use dlnanything::{MediaServer, Renderer, StreamConfig};
use reqwest::{Client, Response, StatusCode};
use std::{
    io::Write,
    process::{Command, Stdio},
    sync::{Arc, Mutex},
    time::Duration,
};
use tokio::time::{sleep_until, timeout, Instant};

// Thirty real WebCodecs Annex B access units from startCanvas at 30 fps, 64×48.
// A white square moves over a green canvas. An AUD marks each access unit in the fixture.
fn pictures() -> Vec<&'static [u8]> {
    let bytes = include_bytes!("fixtures/canvas-30fps.h264");
    let mut offsets: Vec<_> = bytes
        .windows(6)
        .enumerate()
        .filter(|(_, v)| *v == [0, 0, 0, 1, 9, 0xf0])
        .map(|(i, _)| i)
        .collect();
    offsets.push(bytes.len());
    offsets.windows(2).map(|p| &bytes[p[0]..p[1]]).collect()
}

fn client() -> Client {
    Client::builder().no_proxy().build().unwrap()
}

async fn headers(client: &Client, url: &str) -> Response {
    let r = timeout(Duration::from_secs(2), client.get(url).send())
        .await
        .expect("server must flush the HTTP headers before a producer submits a frame")
        .unwrap();
    assert_eq!(r.status(), StatusCode::OK);
    r
}

fn media_url(server: &MediaServer, path: &str) -> String {
    server.playlist_url().replace("live.m3u8", path)
}

#[tokio::test]
async fn http_flushes_headers_and_releases_both_resources_only_when_ready_and_due() {
    let server = MediaServer::bind("127.0.0.1:0".parse().unwrap())
        .await
        .unwrap();
    let client = client();
    let playlist = server.playlist_url();
    let sink = server.frames();
    let key = pictures()[0];
    for url in [&playlist, &media_url(&server, "0.ts")] {
        let probe = client.head(url).send().await.unwrap();
        assert!(probe.status().is_success());
        assert_ne!(
            probe.headers().get("content-length").map(|v| v.as_bytes()),
            Some(&b"0"[..])
        );
    }
    let unnumbered = client
        .head(media_url(&server, "live.ts"))
        .send()
        .await
        .unwrap();
    assert_eq!(unnumbered.status(), StatusCode::NOT_FOUND);
    let zero = media_url(&server, "0.ts");
    let (mut initial, mut segment) =
        tokio::join!(headers(&client, &playlist), headers(&client, &zero));
    let quiet = Duration::from_millis(70);
    let (p, s) = tokio::join!(
        timeout(quiet, initial.chunk()),
        timeout(quiet, segment.chunk())
    );
    assert!(p.is_err() && s.is_err());
    let first = Instant::now();
    assert_eq!(sink.push_annex_b(key).unwrap(), 0);
    let manifest = initial.text().await.unwrap();
    assert!(manifest.contains("#EXTINF:0.500000,\n0.ts"));
    let first_ts = segment.bytes().await.unwrap();
    assert_eq!(presentation_times(&first_ts), [9000]);
    let plain_url = media_url(&server, "1.ts");
    let explicit = format!("{playlist}?_HLS_msn=1");
    // Queue a complete segment early. Its length is known, but its body still has a deadline.
    assert_eq!(sink.push_annex_b(key).unwrap(), 1);
    let (mut reload, mut blocking, mut next) = tokio::join!(
        headers(&client, &playlist),
        headers(&client, &explicit),
        headers(&client, &plain_url)
    );
    assert!(next.content_length().is_some());
    let (p, b, s) = tokio::join!(
        timeout(quiet, reload.chunk()),
        timeout(quiet, blocking.chunk()),
        timeout(quiet, next.chunk())
    );
    assert!(p.is_err() && b.is_err() && s.is_err());
    let (a, b, bytes) = tokio::join!(reload.text(), blocking.text(), next.bytes());
    assert!(first.elapsed() >= Duration::from_millis(495));
    let a = a.unwrap();
    assert!(a.contains("1.ts") && !a.contains("2.ts"));
    assert_eq!(b.unwrap(), a);
    let bytes = bytes.unwrap();
    assert_eq!(presentation_times(&bytes), [54000]);

    let range = client
        .get(&plain_url)
        .header("Range", "bytes=4-19")
        .send()
        .await
        .unwrap();
    assert_eq!(range.status(), StatusCode::PARTIAL_CONTENT);
    assert_eq!(range.bytes().await.unwrap(), bytes.slice(4..20));
    assert_eq!(
        client
            .get(&plain_url)
            .header("Range", "bytes=-0")
            .send()
            .await
            .unwrap()
            .status(),
        416
    );
    let head = client.head(&plain_url).send().await.unwrap();
    assert_eq!(head.headers()["content-length"], bytes.len().to_string());
    assert_eq!(
        client
            .get(media_url(&server, "999.ts"))
            .send()
            .await
            .unwrap()
            .status(),
        400
    );

    // Let the clock pass with no frame. A late producer wakes both HLS readers.
    let next_url = media_url(&server, "2.ts");
    let mut late = headers(&client, &next_url).await;
    let mut late_list = headers(&client, &playlist).await;
    sleep_until(first + Duration::from_millis(1030)).await;
    let (p, s) = tokio::join!(
        timeout(quiet, late_list.chunk()),
        timeout(quiet, late.chunk())
    );
    assert!(p.is_err() && s.is_err());
    sink.push_annex_b(key).unwrap();
    assert!(late_list.text().await.unwrap().contains("2.ts"));
    assert_eq!(presentation_times(&late.bytes().await.unwrap()), [99000]);
    let obsolete = media_url(&server, "3.ts");
    let pending = headers(&client, &obsolete).await;
    drop(server);
    assert!(pending.bytes().await.is_err());
    assert_eq!(
        sink.push_annex_b(key).unwrap_err().kind(),
        std::io::ErrorKind::BrokenPipe
    );
}

// Check the actual transport headers and 90 kHz PTS of every picture, including a
// segment decoded on its own and continuity across HLS segments.
fn presentation_times(ts: &[u8]) -> Vec<u64> {
    assert_eq!(ts.len() % 188, 0);
    let mut result = Vec::new();
    let mut continuity = std::collections::HashMap::new();
    for p in ts.chunks_exact(188) {
        assert_eq!(p[0], 0x47);
        let pid = (u16::from(p[1] & 31) << 8) | u16::from(p[2]);
        if p[3] & 0x10 == 0 {
            continue;
        }
        let cc = p[3] & 15;
        if let Some(previous) = continuity.insert(pid, cc) {
            assert_eq!(cc, (previous + 1) & 15);
        }
        if pid != 0x100 || p[1] & 0x40 == 0 {
            continue;
        }
        let offset = if p[3] & 0x20 != 0 {
            5 + usize::from(p[4])
        } else {
            4
        };
        assert_eq!(&p[offset..offset + 4], &[0, 0, 1, 0xe0]);
        let t = &p[offset + 9..offset + 14];
        result.push(
            (u64::from(t[0] & 14) << 29)
                | (u64::from(t[1]) << 22)
                | (u64::from(t[2] & 0xfe) << 14)
                | (u64::from(t[3]) << 7)
                | u64::from(t[4] >> 1),
        );
    }
    result
}

fn decode_if_available(bytes: &[u8], count: usize) {
    let mut child = match Command::new("ffmpeg")
        .args([
            "-v",
            "error",
            "-xerror",
            "-f",
            "mpegts",
            "-i",
            "pipe:0",
            "-map",
            "0:v:0",
            "-fps_mode",
            "passthrough",
            "-f",
            "framemd5",
            "-",
        ])
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
    {
        Ok(child) => child,
        Err(e) if e.kind() == std::io::ErrorKind::NotFound => return,
        Err(e) => panic!("{e}"),
    };
    child.stdin.take().unwrap().write_all(bytes).unwrap();
    let out = child.wait_with_output().unwrap();
    assert!(
        out.status.success(),
        "{}",
        String::from_utf8_lossy(&out.stderr)
    );
    assert_eq!(
        String::from_utf8(out.stdout)
            .unwrap()
            .lines()
            .filter(|l| !l.starts_with('#'))
            .count(),
        count
    );
}

#[tokio::test]
async fn webcodecs_keyframe_streams_as_one_half_second_segment() {
    // The macOS smoke run passes a picture from the system's WKWebView.
    let key = std::env::var_os("DLNANYTHING_TEST_KEYFRAME")
        .map(|path| std::fs::read(path).expect("read supplied WebCodecs keyframe"))
        .unwrap_or_else(|| pictures()[0].to_owned());
    let server = MediaServer::bind("127.0.0.1:0".parse().unwrap())
        .await
        .unwrap();
    server.frames().push_annex_b(&key).unwrap();
    let client = client();
    let list = headers(&client, &server.playlist_url()).await;
    assert!(list
        .text()
        .await
        .unwrap()
        .contains("#EXTINF:0.500000,\n0.ts"));
    let ts = headers(&client, &media_url(&server, "0.ts"))
        .await
        .bytes()
        .await
        .unwrap();
    assert_eq!(presentation_times(&ts), [9000]);
    decode_if_available(&ts, 1);
}

#[tokio::test]
async fn webcodecs_thirty_fps_produces_two_independently_decodable_segments() {
    let config = StreamConfig::new(30).unwrap();
    let server = MediaServer::bind_with_config("127.0.0.1:0".parse().unwrap(), config)
        .await
        .unwrap();
    let sink = server.frames();
    let frames = pictures();
    assert_eq!(frames.len(), 30);
    assert!(sink.push_annex_b(frames[1]).is_err());
    for (i, frame) in frames.iter().enumerate() {
        assert_eq!(sink.push_annex_b(frame).unwrap(), (i / 15) as u64);
    }
    let client = client();
    let mut all = Vec::new();
    for seq in 0..2 {
        let ts = client
            .get(media_url(&server, &format!("{seq}.ts")))
            .send()
            .await
            .unwrap()
            .bytes()
            .await
            .unwrap();
        assert_eq!(presentation_times(&ts).len(), 15);
        decode_if_available(&ts, 15);
        all.extend_from_slice(&ts);
    }
    assert_eq!(
        presentation_times(&all),
        (0..30).map(|i| 9000 + i * 3000).collect::<Vec<_>>()
    );
    decode_if_available(&all, 30);
}

#[tokio::test]
async fn an_odd_frame_rate_completes_whole_pictures_before_publishing() {
    assert!(StreamConfig::new(1).is_err() && StreamConfig::new(31).is_err());
    let server = MediaServer::bind_with_config(
        "127.0.0.1:0".parse().unwrap(),
        StreamConfig::new(3).unwrap(),
    )
    .await
    .unwrap();
    let frames = pictures();
    let sink = server.frames();
    sink.push_annex_b(frames[0]).unwrap();
    let client = client();
    let mut list = headers(&client, &server.playlist_url()).await;
    assert!(timeout(Duration::from_millis(70), list.chunk())
        .await
        .is_err());
    sink.push_annex_b(frames[1]).unwrap();
    assert!(list
        .text()
        .await
        .unwrap()
        .contains("#EXTINF:0.666667,\n0.ts"));
    assert!(sink.push_annex_b(frames[2]).is_err());
    assert_eq!(sink.push_annex_b(frames[15]).unwrap(), 1);
    assert!(client
        .get(server.playlist_url())
        .send()
        .await
        .unwrap()
        .text()
        .await
        .unwrap()
        .contains("#EXTINF:0.333333,\n1.ts"));
}

#[tokio::test]
async fn renderer_sets_one_uri_and_plays_once_while_the_canvas_changes() {
    use axum::{
        extract::State,
        http::HeaderMap,
        routing::{get, post},
        Router,
    };
    type Actions = Arc<Mutex<Vec<(String, String)>>>;
    const DESCRIPTION: &str = r#"<root><device>
      <deviceType>urn:schemas-upnp-org:device:MediaRenderer:1</deviceType>
      <friendlyName>Test TV</friendlyName><UDN>uuid:tv</UDN><serviceList>
      <service><serviceType>urn:schemas-upnp-org:service:AVTransport:1</serviceType><controlURL>av</controlURL></service>
      <service><serviceType>urn:schemas-upnp-org:service:ConnectionManager:1</serviceType><controlURL>/cm</controlURL></service>
      </serviceList></device></root>"#;
    const PROTOCOLS: &str = r#"<s:Envelope xmlns:s="urn:test"><s:Body><u:GetProtocolInfoResponse xmlns:u="urn:cm">
      <Sink>http-get:*:application/x-mpegURL:*</Sink>
      </u:GetProtocolInfoResponse></s:Body></s:Envelope>"#;

    let actions = Actions::default();
    let record = |State(actions): State<Actions>, headers: HeaderMap, body: String| async move {
        let action = headers["soapaction"].to_str().unwrap().to_owned();
        actions.lock().unwrap().push((action, body));
        "<ok/>"
    };
    let app = Router::new()
        .route("/desc.xml", get(|| async { DESCRIPTION }))
        .route("/cm", post(|| async { PROTOCOLS }))
        .route("/av", post(record))
        .with_state(actions.clone());
    let listener = tokio::net::TcpListener::bind("127.0.0.1:0").await.unwrap();
    let url = format!("http://{}/desc.xml", listener.local_addr().unwrap());
    let tv = tokio::spawn(async move { axum::serve(listener, app).await.unwrap() });
    let r = Renderer::from_location(url.parse().unwrap()).await.unwrap();
    assert_eq!(r.name, "Test TV");
    let media = r
        .media_server(&r.sink_protocols().await.unwrap())
        .await
        .unwrap();
    r.play(&media, "Slides & notes").await.unwrap();
    for _ in 0..3 {
        media.frames().push_annex_b(pictures()[0]).unwrap();
    }
    r.stop().await.unwrap();
    let calls = actions.lock().unwrap();
    assert_eq!(calls.len(), 3);
    assert!(
        calls[0].0.ends_with("#SetAVTransportURI\"") && calls[0].1.contains(&media.playlist_url())
    );
    assert!(
        calls[0].1.contains("application/x-mpegURL")
            && calls[0].1.contains("Slides &amp;amp; notes")
    );
    assert!(calls[1].0.ends_with("#Play\"") && calls[2].0.ends_with("#Stop\""));
    tv.abort();
}
