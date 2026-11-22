//! One VibeOffice per user and display on Linux, as tauri-plugin-single-instance gives Windows (main.rs,
//! `second_start`), without the D-Bus the plugin brings here. The first start holds a lock and listens on a Unix
//! socket in a folder only this user can open ($XDG_RUNTIME_DIR, else ~/.vibeoffice); a later start sends it its
//! folder and arguments and exits. They are taken as the command line is: existing files only (doc::files_among).

use std::{
    ffi::OsString,
    fs,
    io::{self, Read, Write},
    os::unix::{
        ffi::{OsStrExt, OsStringExt},
        fs::{DirBuilderExt, PermissionsExt},
        net::{UnixListener, UnixStream},
    },
    path::{Path, PathBuf},
    time::Duration,
};

/// What a start sends at most: its folder and arguments.
const MAX_MESSAGE: u64 = 1 << 20;

pub enum Claim {
    /// This is the first start: the socket where later ones knock (None if it couldn't be made)
    First(Option<UnixListener>),
    /// The first start has this one's files
    Sent,
}

/// `vibeoffice.<display>.<ext>` in `dir` (the lock, the socket): each display (each login on the machine) has
/// its own instance.
fn named(dir: &Path, ext: &str) -> PathBuf {
    let display = std::env::var("WAYLAND_DISPLAY")
        .or_else(|_| std::env::var("DISPLAY"))
        .unwrap_or_default();
    let display: String = display
        .chars()
        .map(|c| if c.is_ascii_alphanumeric() { c } else { '_' })
        .collect();
    dir.join(format!("vibeoffice.{display}.{ext}"))
}

/// Where the socket goes: `VIBEOFFICE_HOME` (tests), else $XDG_RUNTIME_DIR, else ~/.vibeoffice (the lite
/// edition's ~/VibeOffice).
pub fn folder(home: Option<PathBuf>) -> Option<PathBuf> {
    let name = if cfg!(feature = "lite") {
        "VibeOffice"
    } else {
        ".vibeoffice"
    };
    std::env::var_os("VIBEOFFICE_HOME")
        .or_else(|| std::env::var_os("XDG_RUNTIME_DIR"))
        .map(PathBuf::from)
        .or_else(|| home.map(|h| h.join(name)))
}

/// The first start takes the lock and listens; a later one hands over `args` (and `cwd`, which relative paths
/// are from). Anything that fails: this start runs on its own, as before.
pub fn claim(dir: &Path, cwd: &Path, args: &[OsString]) -> Claim {
    let lock = (|| {
        fs::DirBuilder::new()
            .recursive(true)
            .mode(0o700)
            .create(dir)?;
        fs::OpenOptions::new()
            .create(true)
            .truncate(false)
            .write(true)
            .open(named(dir, "lock"))
    })();
    let lock = match lock {
        Ok(lock) => lock,
        Err(e) => {
            eprintln!("one instance: {e}");
            return Claim::First(None);
        }
    };
    let socket = named(dir, "sock");
    match lock.try_lock() {
        Ok(()) => {
            // held until the process ends
            std::mem::forget(lock);
            let _ = fs::remove_file(&socket);
            match UnixListener::bind(&socket) {
                Ok(listener) => {
                    let _ = fs::set_permissions(&socket, fs::Permissions::from_mode(0o600));
                    Claim::First(Some(listener))
                }
                Err(e) => {
                    eprintln!("one instance: {}: {e}", socket.display());
                    Claim::First(None)
                }
            }
        }
        Err(fs::TryLockError::WouldBlock) => {
            // the first start may be binding its socket still
            for _ in 0..50 {
                match send(&socket, cwd, args) {
                    Ok(()) => return Claim::Sent,
                    Err(_) => std::thread::sleep(Duration::from_millis(100)),
                }
            }
            eprintln!("one instance: the running VibeOffice doesn't answer; starting another");
            Claim::First(None)
        }
        Err(fs::TryLockError::Error(e)) => {
            eprintln!("one instance: {e}");
            Claim::First(None)
        }
    }
}

/// The message: each of the folder and the arguments as its length (4 bytes, little-endian) and its bytes.
fn send(socket: &Path, cwd: &Path, args: &[OsString]) -> io::Result<()> {
    let mut message = Vec::new();
    for part in std::iter::once(cwd.as_os_str()).chain(args.iter().map(|a| a.as_os_str())) {
        message.extend_from_slice(&(part.len() as u32).to_le_bytes());
        message.extend_from_slice(part.as_bytes());
    }
    let mut stream = UnixStream::connect(socket)?;
    stream.write_all(&message)?;
    stream.shutdown(std::net::Shutdown::Write)?;
    // the first start answers once it has them
    let mut answer = [0u8; 1];
    stream.set_read_timeout(Some(Duration::from_secs(5)))?;
    stream.read_exact(&mut answer)
}

fn read(stream: &mut UnixStream) -> io::Result<(PathBuf, Vec<OsString>)> {
    stream.set_read_timeout(Some(Duration::from_secs(5)))?;
    let mut bytes = Vec::new();
    stream.take(MAX_MESSAGE + 1).read_to_end(&mut bytes)?;
    if bytes.len() as u64 > MAX_MESSAGE {
        return Err(io::Error::other("too long"));
    }
    let mut parts = Vec::new();
    let mut rest = &bytes[..];
    while !rest.is_empty() {
        let n = rest
            .get(..4)
            .map(|b| u32::from_le_bytes(b.try_into().unwrap()) as usize)
            .ok_or_else(|| io::Error::other("cut short"))?;
        let part = rest
            .get(4..4 + n)
            .ok_or_else(|| io::Error::other("cut short"))?;
        parts.push(OsString::from_vec(part.to_vec()));
        rest = &rest[4 + n..];
    }
    if parts.is_empty() {
        return Err(io::Error::other("no folder"));
    }
    let cwd = PathBuf::from(parts.remove(0));
    if !cwd.is_absolute() {
        return Err(io::Error::other("not a folder"));
    }
    Ok((cwd, parts))
}

/// Takes later starts' messages, each to `arrived(folder, arguments)`, for as long as the app runs.
pub fn listen(listener: UnixListener, arrived: impl Fn(PathBuf, Vec<OsString>) + Send + 'static) {
    std::thread::spawn(move || {
        for stream in listener.incoming() {
            let Ok(mut stream) = stream else { continue };
            match read(&mut stream) {
                Ok((cwd, args)) => {
                    let _ = stream.write_all(b"1");
                    arrived(cwd, args);
                }
                Err(e) => eprintln!("one instance: a message refused: {e}"),
            }
        }
    });
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::mpsc;

    #[test]
    fn a_second_start_hands_its_files_to_the_first() {
        let dir = std::env::temp_dir().join(format!("vo-instance-{}", std::process::id()));
        let _ = fs::remove_dir_all(&dir);
        let Claim::First(Some(listener)) = claim(&dir, Path::new("/"), &[]) else {
            panic!("the first start listens");
        };
        assert_eq!(
            fs::metadata(&dir).unwrap().permissions().mode() & 0o777,
            0o700
        );
        let (tx, rx) = mpsc::channel();
        listen(listener, move |cwd, args| tx.send((cwd, args)).unwrap());
        // (the lock is per open file: this process stands in for a second one)
        let args = [
            OsString::from("Budget.csv"),
            OsString::from_vec(b"caf\xe9.docx".to_vec()),
        ];
        assert!(matches!(
            claim(&dir, Path::new("/home/me"), &args),
            Claim::Sent
        ));
        let (cwd, got) = rx.recv_timeout(Duration::from_secs(5)).unwrap();
        assert_eq!(cwd, PathBuf::from("/home/me"));
        assert_eq!(got, args, "names that aren't UTF-8 too");
        // a message that isn't one is refused, and the listener goes on
        let mut s = UnixStream::connect(named(&dir, "sock")).unwrap();
        s.write_all(&[9, 0, 0, 0, b'x']).unwrap();
        drop(s);
        assert!(matches!(claim(&dir, Path::new("/"), &[]), Claim::Sent));
        assert_eq!(
            rx.recv_timeout(Duration::from_secs(5)).unwrap().1,
            Vec::<OsString>::new()
        );
        let _ = fs::remove_dir_all(&dir);
    }
}
