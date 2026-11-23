use crate::ts::Muxer;
use bytes::Bytes;
use std::{
    collections::VecDeque,
    io,
    sync::{Arc, Mutex},
    time::Duration,
};
use tokio::{
    sync::watch,
    time::{sleep_until, Instant},
};

/// The default canvas produces one frame and one segment per half second.
pub const TARGET_SEGMENT_DURATION: Duration = Duration::from_millis(500);
const CLOCK: u64 = 90_000;
const MAX_SEGMENT_BYTES: usize = 12 * 1024 * 1024;

#[derive(Clone, Copy, Debug)]
pub struct StreamConfig {
    fps: u8,
}
impl Default for StreamConfig {
    fn default() -> Self {
        Self { fps: 2 }
    }
}
impl StreamConfig {
    pub fn new(fps: u8) -> io::Result<Self> {
        if !(2..=30).contains(&fps) {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "fps must be between 2 and 30",
            ));
        }
        Ok(Self { fps })
    }
    pub fn fps(self) -> u8 {
        self.fps
    }
    fn segment(self, frame: u64) -> u64 {
        frame * 2 / u64::from(self.fps)
    }
    fn first_frame(self, seq: u64) -> u64 {
        (seq * u64::from(self.fps)).div_ceil(2)
    }
    fn clock(self, frame: u64) -> u64 {
        frame * CLOCK / u64::from(self.fps)
    }
}
pub const MAX_FRAME_BYTES: usize = 4 * 1024 * 1024;
pub(crate) const WINDOW: usize = 8;
const RETAIN: usize = 20;

#[derive(Clone)]
pub(crate) struct Segment {
    pub seq: u64,
    pub due: Instant,
    pub duration: Duration,
    pub bytes: Bytes,
}

#[derive(Default)]
struct History {
    mux: Muxer,
    frame: u64,
    pending: Vec<u8>,
    segments: VecDeque<Segment>,
    closed: bool,
}

#[derive(Default)]
pub(crate) struct Timeline {
    config: StreamConfig,
    history: Mutex<History>,
    change: watch::Sender<u64>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Missing {
    Expired,
    Future,
    Closed,
    Timeout,
}

/// Give this handle to the host IPC callback. Frames must arrive in capture order.
#[derive(Clone)]
pub struct FrameSink(pub(crate) Arc<Timeline>);

impl FrameSink {
    /// Accept one WebCodecs Annex B picture. The first picture of each segment must be a keyframe. The returned sequence identifies its .ts segment.
    pub fn push_annex_b(&self, frame: &[u8]) -> io::Result<u64> {
        if frame.len() > MAX_FRAME_BYTES {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "H.264 frame exceeds 4 MiB",
            ));
        }
        let mut h = self.0.history.lock().unwrap();
        if h.closed {
            return Err(io::Error::new(io::ErrorKind::BrokenPipe, "cast ended"));
        }
        let now = Instant::now();
        let config = self.0.config;
        let frame_no = h.frame;
        let seq = config.segment(frame_no);
        let start = frame_no == config.first_frame(seq);
        let due = h
            .segments
            .back()
            .map_or(now, |s| (s.due + s.duration).max(now));
        if due > now + TARGET_SEGMENT_DURATION * 3 {
            return Err(io::Error::new(
                io::ErrorKind::WouldBlock,
                "capture is more than three segments ahead",
            ));
        }
        // Bound the segment even when a caller submits several very large access units.
        let max_added = frame.len().div_ceil(176) * 188 + 4096;
        if h.pending.len() + max_added > MAX_SEGMENT_BYTES {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "segment exceeds 12 MiB",
            ));
        }
        let clock = config.clock(frame_no);
        let next_clock = config.clock(frame_no + 1);
        let History { mux, pending, .. } = &mut *h;
        mux.picture(pending, frame, clock, next_clock - clock, start)?;
        h.frame += 1;
        if config.segment(h.frame) == seq {
            return Ok(seq);
        }
        let ticks = next_clock - config.clock(config.first_frame(seq));
        let duration = Duration::from_nanos(ticks * 1_000_000_000 / CLOCK);
        let bytes = Bytes::from(std::mem::take(&mut h.pending));
        h.segments.push_back(Segment {
            seq,
            due,
            duration,
            bytes,
        });
        if h.segments.len() > RETAIN {
            h.segments.pop_front();
        }
        drop(h);
        self.0.wake();
        Ok(seq)
    }
}

impl Timeline {
    pub fn new(config: StreamConfig) -> Self {
        Self {
            config,
            ..Self::default()
        }
    }
    fn wake(&self) {
        self.change.send_modify(|n| *n = n.wrapping_add(1));
    }

    pub fn close(&self) {
        self.history.lock().unwrap().closed = true;
        self.wake();
    }

    pub fn latest(&self) -> Option<u64> {
        let now = Instant::now();
        self.history
            .lock()
            .unwrap()
            .segments
            .iter()
            .rev()
            .find(|s| s.due <= now)
            .map(|s| s.seq)
    }

    pub fn lookup(&self, seq: u64) -> Result<Option<Segment>, Missing> {
        let h = self.history.lock().unwrap();
        if h.closed {
            return Err(Missing::Closed);
        }
        let first = h.segments.front().map_or(0, |s| s.seq);
        if seq < first {
            return Err(Missing::Expired);
        }
        if seq > self.config.segment(h.frame).saturating_add(1) {
            return Err(Missing::Future);
        }
        Ok(h.segments.get((seq - first) as usize).cloned())
    }

    pub async fn wait(&self, seq: u64, deadline: Instant) -> Result<Segment, Missing> {
        // Subscribe before checking the state so a producer cannot wake us between those steps.
        let mut changes = self.change.subscribe();
        loop {
            let segment = self.lookup(seq)?;
            let now = Instant::now();
            if let Some(s) = segment.as_ref() {
                if s.due <= now {
                    return Ok(s.clone());
                }
            }
            if now >= deadline {
                return Err(Missing::Timeout);
            }
            let wake_at = segment.map_or(deadline, |s| s.due.min(deadline));
            tokio::select! {
                _ = sleep_until(wake_at) => (),
                _ = changes.changed() => (),
            }
        }
    }

    pub fn playlist(&self) -> Option<(u64, String)> {
        let now = Instant::now();
        let h = self.history.lock().unwrap();
        let last = h.segments.iter().rev().find(|s| s.due <= now)?.seq;
        let first = last
            .saturating_sub(WINDOW as u64 - 1)
            .max(h.segments.front()?.seq);
        let mut text=format!("#EXTM3U\n#EXT-X-VERSION:3\n#EXT-X-TARGETDURATION:1\n#EXT-X-INDEPENDENT-SEGMENTS\n#EXT-X-SERVER-CONTROL:CAN-BLOCK-RELOAD=YES\n#EXT-X-MEDIA-SEQUENCE:{first}\n");
        for segment in h
            .segments
            .iter()
            .filter(|s| (first..=last).contains(&s.seq))
        {
            use std::fmt::Write;
            writeln!(
                text,
                "#EXTINF:{:.6},\n{}.ts",
                segment.duration.as_secs_f64(),
                segment.seq
            )
            .unwrap();
        }
        Some((last, text))
    }
}
