use super::*;
use axum::{routing::get, Router};
use std::collections::HashMap;

#[test]
fn only_up_private_multicast_interfaces_are_searched() {
    let flags = InterfaceFlags::UP | InterfaceFlags::MULTICAST;
    for addr in ["10.0.0.1", "172.31.255.254", "192.168.2.1", "169.254.1.1"] {
        assert!(eligible(addr.parse().unwrap(), flags));
    }
    for addr in ["100.64.1.1", "172.32.0.1", "8.8.8.8", "127.0.0.1"] {
        assert!(!eligible(addr.parse().unwrap(), flags));
    }
    let ip = Ipv4Addr::new(10, 0, 0, 1);
    for invalid in [
        InterfaceFlags::UP,
        InterfaceFlags::MULTICAST,
        flags | InterfaceFlags::LOOPBACK,
    ] {
        assert!(!eligible(ip, invalid));
    }
}

#[tokio::test]
async fn retries_on_each_socket_and_reports_one_tv_before_search_ends() {
    let listener = tokio::net::TcpListener::bind((Ipv4Addr::LOCALHOST, 0))
        .await
        .unwrap();
    let address = listener.local_addr().unwrap();
    let app = Router::new().route("/{name}", get(|| async {
        "<root><device><deviceType>urn:schemas-upnp-org:device:MediaRenderer:1</deviceType><UDN>uuid:tv</UDN><serviceList><service><serviceType>urn:schemas-upnp-org:service:AVTransport:1</serviceType><controlURL>/av</controlURL></service></serviceList></device></root>"
    }));
    let http = tokio::spawn(async move { axum::serve(listener, app).await.unwrap() });
    let udp = UdpSocket::bind((Ipv4Addr::LOCALHOST, 0)).await.unwrap();
    let SocketAddr::V4(target) = udp.local_addr().unwrap() else {
        unreachable!()
    };
    let responder = tokio::spawn(async move {
        let mut counts = HashMap::new();
        let mut buf = [0; 512];
        while counts.values().filter(|n| **n == 3).count() < 2 {
            let (n, source) = udp.recv_from(&mut buf).await.unwrap();
            assert_eq!(&buf[..n], SEARCH);
            let count = counts.entry(source).or_insert(0);
            *count += 1;
            if *count == 2 {
                udp.send_to(&[0; 9000], source).await.unwrap();
                udp.send_to(b"malformed reply", source).await.unwrap();
                let reply = format!(
                    "HTTP/1.1 200 OK\r\nLOCATION: http://{address}/{}\r\n\r\n",
                    source.port()
                );
                udp.send_to(reply.as_bytes(), source).await.unwrap();
            }
        }
    });
    let (found, mut tvs) = mpsc::unbounded_channel();
    // Two distinct local sockets represent interfaces on a machine with just one loopback address.
    let scan = tokio::spawn(scan_interfaces(
        vec![
            Ipv4Addr::new(192, 0, 2, 47),
            Ipv4Addr::LOCALHOST,
            Ipv4Addr::LOCALHOST,
        ],
        target,
        Duration::from_secs(1),
        move |tv| {
            found.send(tv.id).unwrap();
        },
    ));
    assert_eq!(tvs.recv().await.as_deref(), Some("uuid:tv"));
    assert!(!scan.is_finished());
    scan.await.unwrap().unwrap();
    assert!(tvs.recv().await.is_none());
    tokio::time::timeout(Duration::from_secs(1), responder)
        .await
        .unwrap()
        .unwrap();
    http.abort();
}

#[tokio::test]
async fn no_usable_socket_returns_an_error() {
    let target = SocketAddrV4::new(Ipv4Addr::LOCALHOST, 1);
    for ips in [Vec::new(), vec![Ipv4Addr::new(192, 0, 2, 47)]] {
        assert!(
            scan_interfaces(ips, target, Duration::from_millis(30), |_| {})
                .await
                .is_err()
        );
    }
}
