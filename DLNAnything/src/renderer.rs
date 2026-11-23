use crate::{MediaServer, StreamConfig};
use reqwest::{Client, Url};
use serde::Deserialize;
use std::{
    collections::HashSet,
    io,
    net::{IpAddr, Ipv4Addr, SocketAddr},
    time::Duration,
};
use tokio::{
    net::UdpSocket,
    task::JoinSet,
    time::{timeout_at, Instant},
};

const DEVICE: &str = "urn:schemas-upnp-org:device:MediaRenderer:";
const SEARCH: &[u8]=b"M-SEARCH * HTTP/1.1\r\nHOST: 239.255.255.250:1900\r\nMAN: \"ssdp:discover\"\r\nMX: 1\r\nST: urn:schemas-upnp-org:device:MediaRenderer:1\r\n\r\n";
const LIMIT: usize = 256 * 1024;
fn invalid(msg: impl Into<String>) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidData, msg.into())
}
fn client(location: &Url, source: Option<IpAddr>) -> io::Result<Client> {
    let mut builder = Client::builder();
    if let Some(peer) = source {
        let host = location
            .host_str()
            .ok_or_else(|| invalid("missing renderer host"))?;
        if let Ok(ip) = host.trim_matches(&['[', ']'][..]).parse::<IpAddr>() {
            if ip != peer {
                return Err(invalid(
                    "description is on a different host than the SSDP responder",
                ));
            }
        } else {
            // The responder supplied this hostname. Connect only to the address which actually answered
            // SSDP, including subsequent SOAP requests, while preserving its advertised Host header.
            builder = builder.resolve(
                host,
                SocketAddr::new(peer, location.port_or_known_default().unwrap_or(80)),
            );
        }
    }
    builder
        .no_proxy()
        .redirect(reqwest::redirect::Policy::none())
        .timeout(Duration::from_secs(4))
        .build()
        .map_err(io::Error::other)
}
async fn response_text(mut r: reqwest::Response) -> io::Result<(bool, String)> {
    let ok = r.status().is_success();
    let mut out = Vec::new();
    while let Some(chunk) = r.chunk().await.map_err(io::Error::other)? {
        if out.len() + chunk.len() > LIMIT {
            return Err(invalid("UPnP XML exceeds 256 KiB"));
        }
        out.extend_from_slice(&chunk);
    }
    Ok((ok, String::from_utf8(out).map_err(invalid_string)?))
}
fn invalid_string(e: impl std::fmt::Display) -> io::Error {
    invalid(e.to_string())
}
fn escaped(s: &str) -> String {
    quick_xml::escape::escape(s).into_owned()
}

#[derive(Deserialize)]
struct Description {
    #[serde(rename = "URLBase", default)]
    base: String,
    device: Device,
}
#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct Device {
    #[serde(default)]
    device_type: String,
    #[serde(default)]
    friendly_name: String,
    #[serde(rename = "UDN", default)]
    udn: String,
    #[serde(default)]
    service_list: Services,
    #[serde(default)]
    device_list: Devices,
}
#[derive(Default, Deserialize)]
struct Services {
    #[serde(default)]
    service: Vec<Service>,
}
#[derive(Default, Deserialize)]
struct Devices {
    #[serde(default)]
    device: Vec<Device>,
}
#[derive(Deserialize)]
struct Service {
    #[serde(rename = "serviceType")]
    kind: String,
    #[serde(rename = "controlURL")]
    path: String,
}
fn renderer_device(dev: Device) -> Option<Device> {
    if dev.device_type.starts_with(DEVICE) {
        Some(dev)
    } else {
        dev.device_list.device.into_iter().find_map(renderer_device)
    }
}

#[derive(Clone, Debug)]
struct Control {
    kind: String,
    url: Url,
}
impl Control {
    async fn action(&self, client: &Client, action: &str, args: &str) -> io::Result<String> {
        let body=format!("<?xml version=\"1.0\"?><s:Envelope xmlns:s=\"http://schemas.xmlsoap.org/soap/envelope/\" s:encodingStyle=\"http://schemas.xmlsoap.org/soap/encoding/\"><s:Body><u:{action} xmlns:u=\"{}\">{args}</u:{action}></s:Body></s:Envelope>",escaped(&self.kind));
        let r = client
            .post(self.url.clone())
            .header("content-type", "text/xml; charset=\"utf-8\"")
            .header("soapaction", format!("\"{}#{action}\"", self.kind))
            .body(body)
            .send()
            .await
            .map_err(io::Error::other)?;
        let (ok, xml) = response_text(r).await?;
        if ok {
            Ok(xml)
        } else {
            Err(io::Error::other(format!(
                "{action} refused: {}",
                xml.chars().take(1200).collect::<String>()
            )))
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum Transport {
    Hls,
    MpegTs,
}

/// A discovered UPnP AV MediaRenderer. Cached control endpoints are specific to this device.
#[derive(Clone, Debug)]
pub struct Renderer {
    pub id: String,
    pub name: String,
    pub location: Url,
    peer: IpAddr,
    av: Control,
    connection: Option<Control>,
    client: Client,
}

impl Renderer {
    /// Read an explicit device description. Also useful on networks that filter SSDP multicast.
    pub async fn from_location(location: Url) -> io::Result<Self> {
        Self::read_description(location, None).await
    }

    async fn read_description(location: Url, source: Option<IpAddr>) -> io::Result<Self> {
        if location.scheme() != "http"
            || !location.username().is_empty()
            || location.password().is_some()
        {
            return Err(invalid("expected an HTTP device description"));
        }
        let client = client(&location, source)?;
        let r = client
            .get(location.clone())
            .send()
            .await
            .map_err(io::Error::other)?;
        let peer = r
            .remote_addr()
            .ok_or_else(|| invalid("unknown renderer address"))?
            .ip();
        let (ok, xml) = response_text(r).await?;
        if !ok {
            return Err(invalid("device description request failed"));
        }
        let desc: Description = quick_xml::de::from_str(&xml).map_err(invalid_string)?;
        let dev = renderer_device(desc.device)
            .ok_or_else(|| invalid("no MediaRenderer in description"))?;
        let base = if desc.base.trim().is_empty() {
            location.clone()
        } else {
            location.join(desc.base.trim()).map_err(invalid_string)?
        };
        let endpoint = |kind: &str| -> io::Result<Option<Control>> {
            let prefix = format!("urn:schemas-upnp-org:service:{kind}:");
            let Some(s) = dev
                .service_list
                .service
                .iter()
                .find(|s| s.kind.starts_with(&prefix))
            else {
                return Ok(None);
            };
            let url = base.join(s.path.trim()).map_err(invalid_string)?;
            if url.scheme() != "http"
                || url.host() != location.host()
                || !url.username().is_empty()
                || url.password().is_some()
            {
                return Err(invalid("control endpoint must be on the renderer host"));
            }
            Ok(Some(Control {
                kind: s.kind.clone(),
                url,
            }))
        };
        let av = endpoint("AVTransport")?.ok_or_else(|| invalid("renderer lacks AVTransport"))?;
        let connection = endpoint("ConnectionManager")?;
        let id = if dev.udn.is_empty() {
            location.as_str().to_owned()
        } else {
            dev.udn
        };
        Ok(Self {
            id,
            name: dev.friendly_name,
            location,
            peer,
            av,
            connection,
            client,
        })
    }

    pub async fn sink_protocols(&self) -> io::Result<Vec<String>> {
        #[derive(Deserialize)]
        struct Envelope {
            #[serde(rename = "Body")]
            body: SoapBody,
        }
        #[derive(Deserialize)]
        struct SoapBody {
            #[serde(rename = "GetProtocolInfoResponse")]
            info: Info,
        }
        #[derive(Deserialize)]
        struct Info {
            #[serde(rename = "Sink", default)]
            sink: String,
        }
        let Some(c) = &self.connection else {
            return Ok(Vec::new());
        };
        let xml = c.action(&self.client, "GetProtocolInfo", "").await?;
        let e: Envelope = quick_xml::de::from_str(&xml).map_err(invalid_string)?;
        Ok(e.body
            .info
            .sink
            .split(',')
            .map(str::trim)
            .filter(|s| !s.is_empty())
            .map(str::to_owned)
            .collect())
    }

    /// Open the media listener on the local interface used to reach this renderer.
    pub async fn media_server(&self, protocols: &[String]) -> io::Result<MediaServer> {
        self.media_server_with_config(protocols, StreamConfig::default())
            .await
    }

    pub async fn media_server_with_config(
        &self,
        protocols: &[String],
        config: StreamConfig,
    ) -> io::Result<MediaServer> {
        let any = if self.peer.is_ipv4() {
            "0.0.0.0:0"
        } else {
            "[::]:0"
        };
        let socket = UdpSocket::bind(any).await?;
        socket.connect(SocketAddr::new(self.peer, 1900)).await?;
        let (hls, ts) = mime_types(protocols);
        MediaServer::bind_types(
            SocketAddr::new(socket.local_addr()?.ip(), 0),
            hls,
            ts,
            config,
        )
        .await
    }

    /// Select the session URL and start it once. Subsequent canvas updates use only HTTP.
    pub async fn play(
        &self,
        server: &MediaServer,
        transport: Transport,
        title: &str,
    ) -> io::Result<()> {
        let (url, mime) = match transport {
            Transport::Hls => (server.playlist_url(), server.hls_type()),
            Transport::MpegTs => (server.transport_stream_url(), server.ts_type()),
        };
        let didl=format!("<DIDL-Lite xmlns=\"urn:schemas-upnp-org:metadata-1-0/DIDL-Lite/\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:upnp=\"urn:schemas-upnp-org:metadata-1-0/upnp/\"><item id=\"1\" parentID=\"0\" restricted=\"1\"><dc:title>{}</dc:title><upnp:class>object.item.videoItem</upnp:class><res protocolInfo=\"http-get:*:{mime}:*\">{}</res></item></DIDL-Lite>",escaped(title),escaped(&url));
        let args=format!("<InstanceID>0</InstanceID><CurrentURI>{}</CurrentURI><CurrentURIMetaData>{}</CurrentURIMetaData>",escaped(&url),escaped(&didl));
        self.av
            .action(&self.client, "SetAVTransportURI", &args)
            .await?;
        self.av
            .action(
                &self.client,
                "Play",
                "<InstanceID>0</InstanceID><Speed>1</Speed>",
            )
            .await?;
        Ok(())
    }

    pub async fn stop(&self) -> io::Result<()> {
        self.av
            .action(&self.client, "Stop", "<InstanceID>0</InstanceID>")
            .await?;
        Ok(())
    }
}

fn mime_types(protocols: &[String]) -> (&'static str, &'static str) {
    let has = |mime: &str| {
        protocols.iter().any(|p| {
            let mut fields = p.splitn(4, ':');
            fields
                .next()
                .is_some_and(|v| v.eq_ignore_ascii_case("http-get"))
                && fields.nth(1).is_some_and(|v| v.eq_ignore_ascii_case(mime))
        })
    };
    let hls = if !has("application/vnd.apple.mpegurl") && has("application/x-mpegURL") {
        "application/x-mpegURL"
    } else {
        "application/vnd.apple.mpegurl"
    };
    let ts = if !has("video/mp2t") && has("video/mpeg") {
        "video/mpeg"
    } else {
        "video/mp2t"
    };
    (hls, ts)
}

/// Discover IPv4 renderers through the system's multicast route.
pub async fn discover(wait: Duration) -> io::Result<Vec<Renderer>> {
    discover_on(Ipv4Addr::UNSPECIFIED, wait).await
}

pub async fn discover_on(interface: Ipv4Addr, wait: Duration) -> io::Result<Vec<Renderer>> {
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
    socket.send_to(SEARCH, "239.255.255.250:1900").await?;
    let until = Instant::now() + wait;
    let mut buf = [0u8; 8192];
    let mut locations = HashSet::new();
    let mut jobs = JoinSet::new();
    while let Ok(result) = timeout_at(until, socket.recv_from(&mut buf)).await {
        let (n, source) = result?;
        let Some(url) = location_header(&buf[..n]) else {
            continue;
        };
        if locations.len() >= 64 || !locations.insert(url.clone()) {
            continue;
        }
        jobs.spawn(async move {
            Renderer::read_description(url, Some(source.ip()))
                .await
                .ok()
        });
    }
    let mut seen = HashSet::new();
    let mut found = Vec::new();
    while let Some(result) = jobs.join_next().await {
        if let Ok(Some(r)) = result {
            if seen.insert(r.id.clone()) {
                found.push(r);
            }
        }
    }
    found.sort_by(|a, b| a.name.cmp(&b.name).then(a.id.cmp(&b.id)));
    Ok(found)
}

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
