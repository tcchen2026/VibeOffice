//! The M5Stack CardputerZero's keyboard (evdev "tca8418c") in the lite edition (docs/desktop.md, Lite edition). The
//! app runs in cage on the device's screen (cardputer/launch.sh), which reads no input itself: this thread reads the
//! keyboard and hands cage what it means as a virtual keyboard and a virtual pointer (zwp_virtual_keyboard_v1,
//! zwlr_virtual_pointer_v1). The keyboard isn't grabbed: M5's launcher, which also reads it, still ends the app when
//! Esc is held.
//!
//! M5's driver resolves its Sym and Fn layers and sticky modifiers itself (hold, one-shot tap, double-tap lock); what
//! a PC keymap can't read is changed here:
//! - Sym symbols arrive as stand-in codes that M5's own software maps to characters; they become the US keyboard's
//!   key for the printed symbol, with Shift where it needs it.
//! - Sym+F/X/Z/C, which M5 leaves unassigned, move the pointer, faster while held; Sym+Enter and Sym+Space are its
//!   left and right buttons. The four pointer keys report only scan records (Linux drops KEY_RESERVED), one per
//!   press and one per release, paired here. A scan without its own key event is such a record even when the driver
//!   syncs it with other keys, as it does for the release that ends a one-shot Sym.
//! - The Sym and Fn keys themselves go nowhere.
//!
//! The keyboard sends no autorepeat: the app repeats held keys itself, at the rate cage gives it.
//!
//! Derived from r86 src/frontend/handheld/cardputer.rs (Apache-2.0).

use std::{
    fs::File,
    io::{Read, Write},
    os::{
        fd::{AsFd, AsRawFd, FromRawFd, OwnedFd},
        unix::net::UnixStream,
    },
    path::{Path, PathBuf},
    sync::Mutex,
    time::Instant,
};

use wayland_client::{
    globals::{registry_queue_init, GlobalListContents},
    protocol::{wl_pointer, wl_registry, wl_seat},
    Connection, Dispatch, EventQueue, QueueHandle,
};
use wayland_protocols_misc::zwp_virtual_keyboard_v1::client::{
    zwp_virtual_keyboard_manager_v1::ZwpVirtualKeyboardManagerV1,
    zwp_virtual_keyboard_v1::ZwpVirtualKeyboardV1,
};
use wayland_protocols_wlr::virtual_pointer::v1::client::{
    zwlr_virtual_pointer_manager_v1::ZwlrVirtualPointerManagerV1,
    zwlr_virtual_pointer_v1::ZwlrVirtualPointerV1,
};

/// The keyboard's evdev name.
const NAME: &str = "tca8418c";

const EV_SYN: u16 = 0;
const EV_KEY: u16 = 1;
const EV_MSC: u16 = 4;
const SYN_REPORT: u16 = 0;
const SYN_DROPPED: u16 = 3;
const MSC_SCAN: u16 = 4;

const ENTER: u16 = 28;
const LEFT_SHIFT: u16 = 42;
const SPACE: u16 = 57;
const BTN_LEFT: u16 = 0x110;
const BTN_RIGHT: u16 = 0x111;
/// The layer keys: Sym (KEY_CD) and Fn (KEY_DVD).
const SYM: u16 = 383;
const FN: u16 = 389;

/// The modifier keys and the modifier each holds down, as masks of the keymap's real modifiers (Shift 1, Lock 2,
/// Control 4, Mod1 8 = Alt, Mod4 64 = Super). A virtual keyboard's keys move no modifier in the compositor
/// (wlroots): its client says which are on (`modifiers`), or Ctrl+Z, Shift+2 and the like arrive as Z and 2.
const MODIFIERS: [(u16, u32); 8] = [
    (LEFT_SHIFT, 1),
    (54, 1), // right Shift
    (29, 4), // Ctrl
    (97, 4),
    (56, 8), // Alt
    (100, 8),
    (125, 64), // Super
    (126, 64),
];
/// Caps Lock: Lock (2) on and off with each press
const CAPS_LOCK: u16 = 58;

/// The modifiers the keys `down` hold.
fn held(down: &[u16]) -> u32 {
    down.iter()
        .filter_map(|code| MODIFIERS.iter().find(|(c, _)| c == code))
        .fold(0, |mods, (_, mask)| mods | mask)
}

/// M5 Sym code -> (US key, with Shift), from M5's character map (m5stack/m5stack-linux-dtoverlays,
/// modules/tca8418-1.0/tca8418_keypad_m5stack_keymap.map). The base layer has no punctuation keys, so these codes
/// only ever mean Sym symbols.
#[rustfmt::skip]
const SYMBOLS: [(u16, u16, bool); 32] = [
    (26, 2, true),    // Sym+1 !
    (27, 3, true),    // Sym+2 @
    (39, 4, true),    // Sym+3 #
    (40, 5, true),    // Sym+4 $
    (41, 6, true),    // Sym+5 %
    (43, 7, true),    // Sym+6 ^
    (51, 8, true),    // Sym+7 &
    (52, 9, true),    // Sym+8 *
    (53, 10, true),   // Sym+9 (
    (94, 11, true),   // Sym+0 )
    (55, 41, true),   // Sym+Q ~
    (69, 41, false),  // Sym+W `
    (70, 12, true),   // Sym+E _
    (71, 12, false),  // Sym+R -
    (72, 13, true),   // Sym+T +
    (73, 13, false),  // Sym+Y =
    (74, 26, false),  // Sym+U [
    (75, 27, false),  // Sym+I ]
    (76, 26, true),   // Sym+O {
    (77, 27, true),   // Sym+P }
    (79, 39, false),  // Sym+A ;
    (80, 39, true),   // Sym+S :
    (81, 40, false),  // Sym+D '
    (82, 40, true),   // Sym+G "
    (83, 51, true),   // Sym+H <
    (85, 52, true),   // Sym+J >
    (86, 43, false),  // Sym+K \
    (89, 43, true),   // Sym+L |
    (90, 51, false),  // Sym+V ,
    (91, 52, false),  // Sym+B .
    (92, 53, false),  // Sym+N /
    (93, 53, true),   // Sym+M ?
];

/// Matrix scans (row * 16 + column) of the Sym-layer pointer keys: F, X, Z, C (up, down, left, right).
const POINTER_SCANS: [i32; 4] = [0x23, 0x33, 0x32, 0x34];
/// Pointer speed in pixels a second, ramping up while held.
const POINTER_START: f32 = 60.0;
const POINTER_MAX: f32 = 360.0;
const POINTER_RAMP: f32 = 0.8;

/// A key or button as cage is told it: an evdev code, pressed (1) or released (0).
#[derive(Debug, PartialEq, Clone, Copy)]
struct Key {
    code: u16,
    value: i32,
}

#[derive(Default)]
struct Keys {
    /// M5's logical Sym state (held, one-shot or locked)
    sym: bool,
    /// the last scan record not yet followed by its key event
    scan: Option<i32>,
    /// translated presses: (M5 code, US code, with Shift), for their release
    held: Vec<(u16, u16, bool)>,
    /// pointer keys down: up, down, left, right
    pointer: [bool; 4],
    moving_since: Option<Instant>,
    last: Option<Instant>,
    rest: (f32, f32),
}

impl Keys {
    /// One input_event from the keyboard; what it means goes to `out`.
    fn event(&mut self, kind: u16, code: u16, value: i32, out: &mut Vec<Key>) {
        match (kind, code) {
            (EV_MSC, MSC_SCAN) => {
                if let Some(scan) = self.scan.replace(value) {
                    self.scan_only(scan);
                }
            }
            (EV_KEY, _) => {
                self.scan = None;
                self.key(code, value, out);
            }
            (EV_SYN, SYN_REPORT) => {
                if let Some(scan) = self.scan.take() {
                    self.scan_only(scan);
                }
            }
            // events were lost: no pairing or layer state can be trusted
            (EV_SYN, SYN_DROPPED) => {
                self.scan = None;
                self.sym = false;
                self.pointer = [false; 4];
                for (_, us, shift) in self.held.drain(..) {
                    out.push(Key { code: us, value: 0 });
                    if shift {
                        out.push(Key {
                            code: LEFT_SHIFT,
                            value: 0,
                        });
                    }
                }
            }
            _ => {}
        }
    }

    /// A press or release of a key the driver reports by scan only.
    fn scan_only(&mut self, scan: i32) {
        if let Some(i) = POINTER_SCANS.iter().position(|&s| s == scan) {
            self.pointer[i] = !self.pointer[i];
        }
    }

    fn key(&mut self, code: u16, value: i32, out: &mut Vec<Key>) {
        if code == SYM {
            self.sym = value != 0;
            return;
        }
        if code == FN {
            return;
        }
        // a release follows its press's translation, even if Sym changed since
        if value == 0 {
            if let Some(i) = self.held.iter().position(|h| h.0 == code) {
                let (_, us, shift) = self.held.remove(i);
                out.push(Key { code: us, value: 0 });
                if shift {
                    out.push(Key {
                        code: LEFT_SHIFT,
                        value: 0,
                    });
                }
                return;
            }
        }
        let translated = match code {
            _ if value != 1 => None,
            ENTER if self.sym => Some((BTN_LEFT, false)),
            SPACE if self.sym => Some((BTN_RIGHT, false)),
            _ => SYMBOLS.iter().find(|s| s.0 == code).map(|s| (s.1, s.2)),
        };
        match translated {
            Some((us, shift)) => {
                self.held.push((code, us, shift));
                if shift {
                    out.push(Key {
                        code: LEFT_SHIFT,
                        value: 1,
                    });
                }
                out.push(Key { code: us, value: 1 });
            }
            None => out.push(Key { code, value }),
        }
    }

    fn moving(&self) -> bool {
        self.pointer.contains(&true)
    }

    /// Pointer motion since the last call, faster while the keys are held.
    fn pointer_motion(&mut self) -> (i32, i32) {
        let now = Instant::now();
        let dt = self
            .last
            .replace(now)
            .map_or(0.0, |last| now.duration_since(last).as_secs_f32().min(0.05));
        if !self.moving() {
            self.moving_since = None;
            self.rest = (0.0, 0.0);
            return (0, 0);
        }
        let since = *self.moving_since.get_or_insert(now);
        let ramp = (now.duration_since(since).as_secs_f32() / POINTER_RAMP).min(1.0);
        let step = (POINTER_START + (POINTER_MAX - POINTER_START) * ramp) * dt;
        let [up, down, left, right] = self.pointer.map(f32::from);
        let x = self.rest.0 + (right - left) * step;
        let y = self.rest.1 + (down - up) * step;
        self.rest = (x.fract(), y.fract());
        (x.trunc() as i32, y.trunc() as i32)
    }
}

/// The keymap cage compiles for the virtual keyboard: a US PC keyboard (the SYMBOLS are its keys).
const KEYMAP: &str = "xkb_keymap {
  xkb_keycodes { include \"evdev+aliases(qwerty)\" };
  xkb_types { include \"complete\" };
  xkb_compat { include \"complete\" };
  xkb_symbols { include \"pc+us+inet(evdev)\" };
};
";

/// /dev/input/eventN of the keyboard, by the name it gives itself in /sys/class/input/eventN/device/name.
fn keyboard() -> Option<PathBuf> {
    std::fs::read_dir("/sys/class/input")
        .ok()?
        .flatten()
        .filter(|e| e.file_name().to_string_lossy().starts_with("event"))
        .find(|e| {
            std::fs::read_to_string(e.path().join("device/name")).is_ok_and(|n| n.trim() == NAME)
        })
        .map(|e| Path::new("/dev/input").join(e.file_name()))
}

/// Starts the keyboard's thread, on the CardputerZero (its keyboard is there) under a Wayland compositor.
pub fn start() {
    let Some(path) = keyboard() else {
        return;
    };
    if std::env::var_os("WAYLAND_DISPLAY").is_none() {
        eprintln!(
            "keyboard: no Wayland compositor to hand {} to",
            path.display()
        );
        return;
    }
    std::thread::spawn(move || {
        if let Err(e) = run(&path) {
            eprintln!("keyboard: {e}");
        }
    });
}

struct State;

impl Dispatch<wl_registry::WlRegistry, GlobalListContents> for State {
    fn event(
        _: &mut Self,
        _: &wl_registry::WlRegistry,
        _: wl_registry::Event,
        _: &GlobalListContents,
        _: &Connection,
        _: &QueueHandle<Self>,
    ) {
    }
}

wayland_client::delegate_noop!(State: ignore wl_seat::WlSeat);
wayland_client::delegate_noop!(State: ignore ZwpVirtualKeyboardManagerV1);
wayland_client::delegate_noop!(State: ignore ZwpVirtualKeyboardV1);
wayland_client::delegate_noop!(State: ignore ZwlrVirtualPointerManagerV1);
wayland_client::delegate_noop!(State: ignore ZwlrVirtualPointerV1);

/// The keymap in a memory file, with its size (the text and its NUL).
fn keymap_file() -> std::io::Result<(OwnedFd, u32)> {
    // SAFETY: memfd_create takes a NUL-terminated name and flags; a new descriptor or -1 comes back.
    let fd = unsafe { libc::memfd_create(c"vibeoffice-keymap".as_ptr(), libc::MFD_CLOEXEC) };
    if fd < 0 {
        return Err(std::io::Error::last_os_error());
    }
    // SAFETY: fd was just made and is owned by nothing else.
    let mut file = unsafe { File::from_raw_fd(fd) };
    file.write_all(KEYMAP.as_bytes())?;
    file.write_all(&[0])?;
    Ok((file.into(), KEYMAP.len() as u32 + 1))
}

/// An input_event as the kernel writes it (64-bit time).
#[repr(C)]
struct InputEvent {
    sec: i64,
    usec: i64,
    kind: u16,
    code: u16,
    value: i32,
}

/// Where the keyboard goes (`follow`): a count of the changes, and the display: none for the app's own (cage's), or
/// the HDMI desktop's for a slide show there (lite.rs).
static FOLLOW: Mutex<(usize, Option<String>)> = Mutex::new((0, None));

/// The keyboard goes to `display` (its Wayland socket's name), or back to the app's own; from its next key on.
pub fn follow(display: Option<String>) {
    let mut f = FOLLOW.lock().unwrap();
    f.0 += 1;
    f.1 = display;
}

/// A virtual keyboard and pointer on one compositor's seat.
struct Seat {
    conn: Connection,
    queue: EventQueue<State>,
    keyboard: ZwpVirtualKeyboardV1,
    pointer: ZwlrVirtualPointerV1,
    /// keys and buttons down here, let go before the keyboard goes elsewhere
    down: Vec<u16>,
    /// Caps Lock's Lock, on or off
    locked: u32,
}

impl Seat {
    /// On the app's own display (WAYLAND_DISPLAY), or on `display` in $XDG_RUNTIME_DIR.
    fn connect(display: Option<&str>) -> Result<Seat, String> {
        let conn = match display {
            None => Connection::connect_to_env().map_err(|e| e.to_string())?,
            Some(name) => {
                let dir = std::env::var_os("XDG_RUNTIME_DIR").ok_or("no XDG_RUNTIME_DIR")?;
                let socket = UnixStream::connect(Path::new(&dir).join(name))
                    .map_err(|e| format!("{name}: {e}"))?;
                Connection::from_socket(socket).map_err(|e| e.to_string())?
            }
        };
        let (globals, queue) = registry_queue_init::<State>(&conn).map_err(|e| e.to_string())?;
        let qh = queue.handle();
        let bind = |what: &str, e: wayland_client::globals::BindError| format!("{what}: {e}");
        let seat: wl_seat::WlSeat = globals
            .bind(&qh, 1..=7, ())
            .map_err(|e| bind("wl_seat", e))?;
        let keyboards: ZwpVirtualKeyboardManagerV1 = globals
            .bind(&qh, 1..=1, ())
            .map_err(|e| bind("zwp_virtual_keyboard_manager_v1", e))?;
        let pointers: ZwlrVirtualPointerManagerV1 = globals
            .bind(&qh, 1..=2, ())
            .map_err(|e| bind("zwlr_virtual_pointer_manager_v1", e))?;
        let keyboard = keyboards.create_virtual_keyboard(&seat, &qh, ());
        let pointer = pointers.create_virtual_pointer(Some(&seat), &qh, ());
        let (keymap, size) = keymap_file().map_err(|e| format!("keymap: {e}"))?;
        keyboard.keymap(1 /* xkb_v1 */, keymap.as_fd(), size);
        conn.flush().map_err(|e| e.to_string())?;
        Ok(Seat {
            conn,
            queue,
            keyboard,
            pointer,
            down: Vec::new(),
            locked: 0,
        })
    }

    fn send(&mut self, time: u32, Key { code, value }: Key) {
        if value > 1 {
            return; // a repeat: the app repeats
        }
        self.down.retain(|&c| c != code);
        if value == 1 {
            self.down.push(code);
        }
        if code == BTN_LEFT || code == BTN_RIGHT {
            let state = if value == 1 {
                wl_pointer::ButtonState::Pressed
            } else {
                wl_pointer::ButtonState::Released
            };
            self.pointer.button(time, code.into(), state);
            self.pointer.frame();
        } else {
            self.keyboard.key(time, code.into(), value as u32);
            if code == CAPS_LOCK && value == 1 {
                self.locked ^= 2;
            }
            if code == CAPS_LOCK || MODIFIERS.iter().any(|(c, _)| *c == code) {
                self.keyboard.modifiers(held(&self.down), 0, self.locked, 0);
            }
        }
    }

    fn motion(&self, time: u32, dx: i32, dy: i32) {
        self.pointer.motion(time, dx.into(), dy.into());
        self.pointer.frame();
    }

    fn flush(&mut self) -> Result<(), String> {
        self.conn.flush().map_err(|e| e.to_string())?;
        self.queue
            .dispatch_pending(&mut State)
            .map(|_| ())
            .map_err(|e| e.to_string())
    }

    /// Lets go of what is down, before the keyboard goes elsewhere.
    fn leave(mut self, time: u32) {
        for code in std::mem::take(&mut self.down) {
            self.send(time, Key { code, value: 0 });
        }
        let _ = self.flush();
    }
}

fn run(path: &Path) -> Result<(), String> {
    let mut device = File::open(path).map_err(|e| format!("{}: {e}", path.display()))?;
    let mut seat = Seat::connect(None)?;
    let mut following = 0;
    let start = Instant::now();
    let time = || start.elapsed().as_millis() as u32;
    let mut keys = Keys::default();
    let mut out = Vec::new();
    const SIZE: usize = std::mem::size_of::<InputEvent>();
    let mut buffer = [0u8; SIZE * 32];
    loop {
        // while the pointer keys are held, a step every 10 ms; else wait for the keyboard
        let mut fds = [libc::pollfd {
            fd: device.as_raw_fd(),
            events: libc::POLLIN,
            revents: 0,
        }];
        let timeout = if keys.moving() { 10 } else { -1 };
        // SAFETY: fds is one valid pollfd, borrowed for the call.
        let ready = unsafe { libc::poll(fds.as_mut_ptr(), 1, timeout) };
        if ready > 0 {
            let len = match device.read(&mut buffer) {
                Ok(0) => return Err(format!("{}: the keyboard is gone", path.display())),
                Ok(len) => len,
                Err(e) if e.kind() == std::io::ErrorKind::Interrupted => continue,
                Err(e) => return Err(format!("{}: {e}", path.display())),
            };
            for chunk in buffer[..len - len % SIZE].chunks_exact(SIZE) {
                // SAFETY: the kernel writes whole input_event records, and any bits are a valid InputEvent.
                let event: InputEvent = unsafe { std::ptr::read_unaligned(chunk.as_ptr().cast()) };
                keys.event(event.kind, event.code, event.value, &mut out);
            }
        }
        // the window went to another display (a slide show on HDMI) or came back
        let (count, display) = FOLLOW.lock().unwrap().clone();
        if count != following {
            following = count;
            match Seat::connect(display.as_deref()) {
                Ok(next) => std::mem::replace(&mut seat, next).leave(time()),
                Err(e) => eprintln!("keyboard: {e}"),
            }
        }
        for key in out.drain(..) {
            seat.send(time(), key);
        }
        let (dx, dy) = keys.pointer_motion();
        if dx != 0 || dy != 0 {
            seat.motion(time(), dx, dy);
        }
        seat.flush()?;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn modifier_keys_hold_their_modifiers() {
        assert_eq!(held(&[]), 0);
        assert_eq!(held(&[29]), 4); // Ctrl
        assert_eq!(held(&[29, 44]), 4); // Ctrl+Z: Z holds nothing
        assert_eq!(held(&[LEFT_SHIFT, 97, 56]), 1 | 4 | 8); // Shift, right Ctrl, Alt
        assert_eq!(held(&[LEFT_SHIFT, 54]), 1); // both Shifts: Shift
        assert_eq!(held(&[125]), 64); // Super
    }

    fn feed(keys: &mut Keys, frames: &[(Option<i32>, Option<(u16, i32)>)]) -> Vec<(u16, i32)> {
        let mut out = Vec::new();
        for &(scan, key) in frames {
            if let Some(scan) = scan {
                keys.event(EV_MSC, MSC_SCAN, scan, &mut out);
            }
            if let Some((code, value)) = key {
                keys.event(EV_KEY, code, value, &mut out);
            }
            keys.event(EV_SYN, SYN_REPORT, 0, &mut out);
        }
        out.iter().map(|k| (k.code, k.value)).collect()
    }

    #[test]
    fn sym_symbols_become_us_keys_with_shift() {
        let mut keys = Keys::default();
        // Sym held, 2 (sent as KEY_RIGHTBRACE) means @: Shift+2; Sym itself goes nowhere
        let out = feed(
            &mut keys,
            &[
                (Some(0x43), Some((SYM, 1))),
                (Some(0x00), Some((27, 1))),
                (Some(0x43), Some((SYM, 0))),
                (Some(0x00), Some((27, 0))),
            ],
        );
        assert_eq!(out, [(42, 1), (3, 1), (3, 0), (42, 0)]);
        // Sym+R (KP7) is -, unshifted; plain letters pass, Fn doesn't
        let out = feed(
            &mut keys,
            &[
                (Some(0x13), Some((71, 1))),
                (Some(0x20), Some((30, 1))),
                (Some(0x42), Some((FN, 1))),
            ],
        );
        assert_eq!(out, [(12, 1), (30, 1)]);
    }

    #[test]
    fn sym_pointer_keys_pair_scan_records_and_click() {
        let mut keys = Keys::default();
        let out = feed(
            &mut keys,
            &[(Some(0x43), Some((SYM, 1))), (Some(0x23), None)],
        );
        assert!(out.is_empty() && keys.pointer == [true, false, false, false]);
        feed(&mut keys, &[(Some(0x23), None)]);
        assert!(!keys.moving());
        let out = feed(
            &mut keys,
            &[
                (Some(0x29), Some((ENTER, 1))),
                (Some(0x29), Some((ENTER, 0))),
            ],
        );
        assert_eq!(out, [(BTN_LEFT, 1), (BTN_LEFT, 0)]);
        // the release that ends a one-shot Sym comes in one frame with Sym's
        let mut out = Vec::new();
        keys.event(EV_MSC, MSC_SCAN, 0x23, &mut out);
        keys.event(EV_SYN, SYN_REPORT, 0, &mut out);
        assert!(keys.moving());
        keys.event(EV_MSC, MSC_SCAN, 0x23, &mut out);
        keys.event(EV_MSC, MSC_SCAN, 0x43, &mut out);
        keys.event(EV_KEY, SYM, 0, &mut out);
        keys.event(EV_SYN, SYN_REPORT, 0, &mut out);
        assert!(!keys.moving() && out.is_empty());
        // without Sym, Enter is Enter
        let out = feed(
            &mut keys,
            &[(Some(0x43), Some((SYM, 0))), (Some(0x29), Some((28, 1)))],
        );
        assert_eq!(out, [(28, 1)]);
    }

    #[test]
    fn the_pointer_speeds_up_while_held() {
        let mut keys = Keys::default();
        feed(&mut keys, &[(Some(0x34), None)]); // Sym+C: right
        keys.pointer_motion();
        std::thread::sleep(std::time::Duration::from_millis(40));
        let (first, y) = keys.pointer_motion();
        std::thread::sleep(std::time::Duration::from_millis(900));
        let _ = keys.pointer_motion(); // a long gap counts as 50 ms
        std::thread::sleep(std::time::Duration::from_millis(40));
        let (later, _) = keys.pointer_motion();
        assert!(
            first >= 1 && y == 0 && later > first,
            "{first} then {later}"
        );
    }
}
