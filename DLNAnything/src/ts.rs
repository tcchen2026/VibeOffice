//! Single-program, 188-byte MPEG transport stream. One PES per Annex B access unit.
use std::io::{Error, ErrorKind, Result};

const VIDEO: u16 = 0x100;
const PMT: u16 = 0x1000;
const MASK: u64 = (1 << 33) - 1;

#[derive(Default)]
pub(crate) struct Muxer {
    pat_cc: u8,
    pmt_cc: u8,
    video_cc: u8,
}

fn crc32_mpeg(data: &[u8]) -> u32 {
    let mut crc = u32::MAX;
    for &byte in data {
        crc ^= u32::from(byte) << 24;
        for _ in 0..8 {
            crc = if crc & 0x8000_0000 != 0 {
                (crc << 1) ^ 0x04c1_1db7
            } else {
                crc << 1
            };
        }
    }
    crc
}

fn table(out: &mut Vec<u8>, pid: u16, cc: &mut u8, section: &[u8]) {
    let mut p = [0xff; 188];
    p[..5].copy_from_slice(&[0x47, 0x40 | (pid >> 8) as u8, pid as u8, 0x10 | *cc, 0]);
    p[5..5 + section.len()].copy_from_slice(section);
    p[5 + section.len()..9 + section.len()].copy_from_slice(&crc32_mpeg(section).to_be_bytes());
    *cc = (*cc + 1) & 15;
    out.extend_from_slice(&p);
}

fn pcr(base: u64) -> [u8; 6] {
    let t = base & MASK;
    [
        (t >> 25) as u8,
        (t >> 17) as u8,
        (t >> 9) as u8,
        (t >> 1) as u8,
        ((t & 1) as u8) << 7 | 0x7e,
        0,
    ]
}

fn presentation_time(t: u64) -> [u8; 5] {
    [
        0x21 | ((t >> 29) as u8 & 14),
        (t >> 22) as u8,
        ((t >> 14) as u8 & 0xfe) | 1,
        (t >> 7) as u8,
        ((t << 1) as u8) | 1,
    ]
}

// Validate the random-access contract before advancing any transport counters.
fn access_unit(frame: &[u8], independent: bool) -> Result<(bool, bool)> {
    let invalid = || {
        Error::new(
            ErrorKind::InvalidData,
            "expected one Annex B H.264 picture, with SPS/PPS and IDR at segment start",
        )
    };
    let (mut sps, mut pps, mut aud) = (false, false, false);
    let mut picture = None;
    let mut zeros = 0;
    let mut header = false;
    let mut started = false;
    for &byte in frame {
        if header {
            if byte & 0x80 != 0 {
                return Err(invalid());
            }
            match byte & 31 {
                7 if picture.is_none() => sps = true,
                8 if picture.is_none() => pps = true,
                9 if picture.is_none() && !aud => aud = true,
                kind @ (1 | 5) => {
                    let key = kind == 5;
                    if picture.is_some_and(|prior| prior != key)
                        || (independent && !(key && sps && pps))
                    {
                        return Err(invalid());
                    }
                    picture = Some(key);
                }
                2..=4 | 7..=9 => return Err(invalid()),
                _ => (),
            }
            header = false;
            zeros = 0;
        } else if byte == 0 {
            zeros += 1;
        } else if byte == 1 && zeros >= 2 {
            started = true;
            header = true;
            zeros = 0;
        } else {
            if !started {
                return Err(invalid());
            }
            zeros = 0;
        }
    }
    if header {
        return Err(invalid());
    }
    Ok((aud, picture.ok_or_else(invalid)?))
}

impl Muxer {
    pub(crate) fn picture(
        &mut self,
        out: &mut Vec<u8>,
        frame: &[u8],
        clock: u64,
        ticks: u64,
        segment_start: bool,
    ) -> Result<()> {
        let (has_aud, key) = access_unit(frame, segment_start)?;
        if segment_start {
            table(
                out,
                0,
                &mut self.pat_cc,
                &[0, 0xb0, 13, 0, 1, 0xc1, 0, 0, 0, 1, 0xf0, 0],
            );
            table(
                out,
                PMT,
                &mut self.pmt_cc,
                &[
                    2, 0xb0, 18, 0, 1, 0xc1, 0, 0, 0xe1, 0, 0xf0, 0, 0x1b, 0xe1, 0, 0xf0, 0,
                ],
            );
        }
        let mut pes = Vec::with_capacity(frame.len() + 20);
        pes.extend_from_slice(&[0, 0, 1, 0xe0, 0, 0, 0x84, 0x80, 5]);
        pes.extend_from_slice(&presentation_time((clock + 9_000) & MASK));
        if !has_aud {
            pes.extend_from_slice(&[0, 0, 0, 1, 9, 0xf0]);
        }
        pes.extend_from_slice(frame);
        let mut left = pes.as_slice();
        let mut first = true;
        while !left.is_empty() {
            let mut p = [0xff; 188];
            let count = left.len().min(if first { 176 } else { 184 });
            let gap = 184 - count;
            p[..4].copy_from_slice(&[
                0x47,
                (VIDEO >> 8) as u8 | if first { 0x40 } else { 0 },
                VIDEO as u8,
                (if gap > 0 { 0x30 } else { 0x10 }) | self.video_cc,
            ]);
            if gap > 0 {
                p[4] = (gap - 1) as u8;
            }
            if gap > 1 {
                p[5] = if first && key {
                    0x50
                } else if first {
                    0x10
                } else {
                    0
                };
            }
            if first {
                p[6..12].copy_from_slice(&pcr(clock));
            }
            p[4 + gap..].copy_from_slice(&left[..count]);
            out.extend_from_slice(&p);
            self.video_cc = (self.video_cc + 1) & 15;
            left = &left[count..];
            first = false;
        }
        // Maintain the program clock between sparse pictures.
        for offset in (3_600..ticks).step_by(3_600) {
            let mut p = [0xff; 188];
            p[..6].copy_from_slice(&[
                0x47,
                (VIDEO >> 8) as u8,
                VIDEO as u8,
                0x20 | self.video_cc.wrapping_sub(1) & 15,
                183,
                0x10,
            ]);
            p[6..12].copy_from_slice(&pcr(clock + offset));
            out.extend_from_slice(&p);
        }
        Ok(())
    }
}
