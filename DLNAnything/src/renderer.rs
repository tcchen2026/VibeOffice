use crate::{MediaServer, StreamConfig};
use reqwest::{Client, Url};
use serde::Deserialize;
use std::{
    io,
    net::{IpAddr, SocketAddr},
    time::Duration,
};
use tokio::net::UdpSocket;

const DEVICE: &str = "urn:schemas-upnp-org:device:MediaRenderer:";
const LIMIT: usize = 256 * 1024;
fn invalid(msg: impl Into<String>) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidData, msg.into())
}
fn network_error(e: reqwest::Error) -> io::Error {
    let kind = if e.is_timeout() {
        io::ErrorKind::TimedOut
    } else if e.is_connect() {
        io::ErrorKind::NotConnected
    } else {
        io::ErrorKind::Other
    };
    io::Error::new(kind, e)
}
fn client(location: &Url, source: Option<IpAddr>, local: Option<IpAddr>) -> io::Result<Client> {
    let mut builder = Client::builder().local_address(local);
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
        .map_err(network_error)
}
async fn response_text(mut r: reqwest::Response) -> io::Result<(bool, String)> {
    let ok = r.status().is_success();
    let mut out = Vec::new();
    while let Some(chunk) = r.chunk().await.map_err(network_error)? {
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
            .map_err(network_error)?;
        let (ok, xml) = response_text(r).await?;
        if ok {
            Ok(xml)
        } else {
            Err(io::Error::new(
                io::ErrorKind::Unsupported,
                format!(
                    "{action} refused: {}",
                    xml.chars().take(1200).collect::<String>()
                ),
            ))
        }
    }
}

/// A discovered UPnP AV MediaRenderer. Cached control endpoints are specific to this device.
#[derive(Clone, Debug)]
pub struct Renderer {
    pub id: String,
    pub name: String,
    pub location: Url,
    peer: IpAddr,
    local: Option<IpAddr>,
    av: Control,
    connection: Option<Control>,
    client: Client,
}

impl Renderer {
    /// Read an explicit device description. Also useful on networks that filter SSDP multicast.
    pub async fn from_location(location: Url) -> io::Result<Self> {
        Self::read_description(location, None, None).await
    }

    pub(crate) async fn read_description(
        location: Url,
        source: Option<IpAddr>,
        local: Option<IpAddr>,
    ) -> io::Result<Self> {
        if location.scheme() != "http"
            || !location.username().is_empty()
            || location.password().is_some()
        {
            return Err(invalid("expected an HTTP device description"));
        }
        let client = client(&location, source, local)?;
        let r = client
            .get(location.clone())
            .send()
            .await
            .map_err(network_error)?;
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
            local,
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
        let local = if let Some(ip) = self.local {
            ip
        } else {
            let any = if self.peer.is_ipv4() {
                "0.0.0.0:0"
            } else {
                "[::]:0"
            };
            let socket = UdpSocket::bind(any).await?;
            socket.connect(SocketAddr::new(self.peer, 1900)).await?;
            socket.local_addr()?.ip()
        };
        MediaServer::bind_types(SocketAddr::new(local, 0), hls_type(protocols), config).await
    }

    /// Select the session URL and start it once. Subsequent canvas updates use only HTTP.
    pub async fn play(&self, server: &MediaServer, title: &str) -> io::Result<()> {
        let url = server.playlist_url();
        let mime = server.hls_type();
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

fn hls_type(protocols: &[String]) -> &'static str {
    let has = |mime: &str| {
        protocols.iter().any(|p| {
            let mut fields = p.splitn(4, ':');
            fields
                .next()
                .is_some_and(|v| v.eq_ignore_ascii_case("http-get"))
                && fields.nth(1).is_some_and(|v| v.eq_ignore_ascii_case(mime))
        })
    };
    if !has("application/vnd.apple.mpegurl") && has("application/x-mpegURL") {
        "application/x-mpegURL"
    } else {
        "application/vnd.apple.mpegurl"
    }
}
