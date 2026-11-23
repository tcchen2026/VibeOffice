use crate::Renderer;
use getifaddrs::{getifaddrs, InterfaceFlags};
use reqwest::Url;
use std::{
    collections::HashSet,
    io,
    net::{IpAddr, Ipv4Addr, SocketAddr, SocketAddrV4},
    time::Duration,
};
use tokio::{
    net::UdpSocket,
    sync::mpsc,
    task::JoinSet,
    time::{sleep, sleep_until, Instant},
};

const SEARCH: &[u8]=b"M-SEARCH * HTTP/1.1\r\nHOST: 239.255.255.250:1900\r\nMAN: \"ssdp:discover\"\r\nMX: 1\r\nST: urn:schemas-upnp-org:device:MediaRenderer:1\r\n\r\n";

fn location_header(data: &[u8]) -> Option<Url> {
    let mut lines = std::str::from_utf8(data).ok()?.lines();
    if lines.next()?.split_whitespace().nth(1)? != "200" {
        return None;
    }
    lines
        .filter_map(|line| line.split_once(':'))
        .find(|(name, _)| name.eq_ignore_ascii_case("location"))
        .and_then(|(_, value)| Url::parse(value.trim()).ok())
        .filter(|url| url.scheme() == "http")
}

const SSDP: SocketAddrV4 = SocketAddrV4::new(Ipv4Addr::new(239, 255, 255, 250), 1900);

fn eligible(ip: Ipv4Addr, flags: InterfaceFlags) -> bool {
    flags.contains(InterfaceFlags::UP | InterfaceFlags::MULTICAST)
        && !flags.contains(InterfaceFlags::LOOPBACK)
        && (ip.is_private() || ip.is_link_local())
}

fn interfaces(adapters: io::Result<impl Iterator<Item = getifaddrs::Interface>>) -> Vec<Ipv4Addr> {
    let mut ips = Vec::new();
    for interface in adapters.into_iter().flatten() {
        if let Some(IpAddr::V4(ip)) = interface.address.ip_addr() {
            if eligible(ip, interface.flags) && !ips.contains(&ip) {
                ips.push(ip);
            }
        }
    }
    ips
}

/// Report each distinct renderer as soon as its device description is available.
pub async fn discover_with(wait: Duration, found: impl FnMut(Renderer)) -> io::Result<()> {
    scan_interfaces(interfaces(getifaddrs()), SSDP, wait, found).await
}

/// Search all eligible local IPv4 interfaces and return TVs sorted by name.
pub async fn discover(wait: Duration) -> io::Result<Vec<Renderer>> {
    collect(interfaces(getifaddrs()), wait).await
}

/// Search a particular local IPv4 interface, including one chosen by an embedding app.
pub async fn discover_on(interface: Ipv4Addr, wait: Duration) -> io::Result<Vec<Renderer>> {
    collect(vec![interface], wait).await
}

async fn collect(ips: Vec<Ipv4Addr>, wait: Duration) -> io::Result<Vec<Renderer>> {
    let mut tvs = Vec::new();
    scan_interfaces(ips, SSDP, wait, |r| tvs.push(r)).await?;
    tvs.sort_by(|a, b| a.name.cmp(&b.name).then_with(|| a.id.cmp(&b.id)));
    Ok(tvs)
}

type Reply = (Url, IpAddr, Ipv4Addr);

async fn probe(
    interface: Ipv4Addr,
    target: SocketAddrV4,
    wait: Duration,
    replies: mpsc::Sender<Reply>,
) -> io::Result<()> {
    let socket = socket2::Socket::new(
        socket2::Domain::IPV4,
        socket2::Type::DGRAM,
        Some(socket2::Protocol::UDP),
    )?;
    socket.bind(&SocketAddr::from((interface, 0)).into())?;
    socket.set_multicast_if_v4(&interface)?;
    socket.set_multicast_ttl_v4(2)?;
    socket.set_nonblocking(true)?;
    let socket = UdpSocket::from_std(socket.into())?;
    socket.send_to(SEARCH, target).await?;
    let start = Instant::now();
    let interval = wait.min(Duration::from_secs(1)) / 3;
    let mut sent = 1;
    let mut packet = [0; 8192];
    loop {
        tokio::select! {
            biased;
            _ = sleep_until(start + wait) => return Ok(()),
            _ = sleep_until(start + interval * sent), if sent < 3 => {
                sent += 1;
                // The first send succeeded. Keep listening even if a resend fails.
                let _ = socket.send_to(SEARCH, target).await;
            }
            reply = socket.recv_from(&mut packet) => match reply {
                Ok((len, source)) => {
                    if let Some(url) = location_header(&packet[..len]) {
                        let _ = replies.try_send((url, source.ip(), interface));
                    }
                }
                // Windows reports oversized UDP datagrams and some ICMP replies as receive errors.
                Err(_) => sleep(Duration::from_millis(5)).await,
            }
        }
    }
}

async fn scan_interfaces(
    mut ips: Vec<Ipv4Addr>,
    target: SocketAddrV4,
    wait: Duration,
    mut found: impl FnMut(Renderer),
) -> io::Result<()> {
    // A publicly addressed campus LAN, or unavailable OS interface metadata, can
    // still have a usable system multicast route. A failed private probe is a real error.
    if ips.is_empty() {
        ips.push(Ipv4Addr::UNSPECIFIED);
    }
    let (send, mut replies) = mpsc::channel(64);
    let mut probes = JoinSet::new();
    for ip in ips {
        probes.spawn(probe(ip, target, wait, send.clone()));
    }
    drop(send);
    let mut locations = HashSet::new();
    let mut ids = HashSet::new();
    let mut descriptions = JoinSet::new();
    let mut listening = true;
    while listening || !descriptions.is_empty() {
        tokio::select! {
            reply = replies.recv(), if listening => match reply {
                Some((url, peer, interface)) => {
                    if locations.len() < 64 && locations.insert((url.clone(), peer, interface)) {
                        let local = (!interface.is_unspecified()).then_some(IpAddr::V4(interface));
                        descriptions.spawn(Renderer::read_description(url, Some(peer), local));
                    }
                }
                None => listening = false,
            },
            Some(result) = descriptions.join_next(), if !descriptions.is_empty() => {
                if let Ok(Ok(tv)) = result {
                    if ids.insert(tv.id.clone()) { found(tv); }
                }
            }
        }
    }
    let mut first_error = None;
    let mut searched = false;
    while let Some(result) = probes.join_next().await {
        match result.map_err(io::Error::other).and_then(|r| r) {
            Ok(()) => searched = true,
            Err(e) => {
                first_error.get_or_insert(e);
            }
        }
    }
    match (searched, first_error) {
        (false, Some(e)) => Err(e),
        _ => Ok(()),
    }
}

#[cfg(test)]
mod tests;
