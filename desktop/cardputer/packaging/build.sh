#!/bin/sh
# The lite edition for the CardputerZero (docs/desktop.md, Lite edition): desktop/ built for aarch64 Linux with
# `--features lite`, linked against the system libraries it runs with there (WebKitGTK 4.1, GTK 3, libsoup 3:
# Raspberry Pi OS is Debian 13 plus its own archive, whose builds of these have Debian's ABI; its WebKitGTK is
# Debian's). The step pack-deb.sh runs first (meta.env); alone it leaves the binary in
# desktop/target/aarch64-unknown-linux-gnu/release/.
#
# No cross gcc and no container: apt, run as this user with its own folders, downloads the arm64 development
# packages of Debian 13 (signed by the Debian archive keyring of this machine) into desktop/target/cardputer/, where
# they are unpacked as a sysroot; clang links against it with lld. Needs: Rust with the aarch64-unknown-linux-gnu
# target, clang, lld, llvm-ar, pkg-config, apt-get and dpkg-deb (a Debian or Ubuntu machine), the network the first
# time.
set -eu
desktop=$(cd "$(dirname "$0")/../.." && pwd)
work=$desktop/target/cardputer
apt=$work/apt
sysroot=$work/sysroot
target=aarch64-unknown-linux-gnu
mkdir -p "$apt/etc/apt.conf.d" "$apt/etc/preferences.d" "$apt/etc/trusted.gpg.d" "$apt/etc/sources.list.d" \
  "$apt/state/lists/partial" "$apt/debs/partial"

keyring=
for k in /usr/share/keyrings/debian-archive-keyring.pgp /usr/share/keyrings/debian-archive-keyring.gpg; do
  [ -f "$k" ] && keyring=$k && break
done
[ -n "$keyring" ] || { echo "no Debian archive keyring (the debian-archive-keyring package)" >&2; exit 1; }
cat > "$apt/etc/sources.list.d/debian.sources" <<EOF
Types: deb
URIs: http://deb.debian.org/debian/
Suites: trixie trixie-updates
Components: main
Signed-By: $keyring

Types: deb
URIs: http://deb.debian.org/debian-security/
Suites: trixie-security
Components: main
Signed-By: $keyring
EOF
cat > "$apt/apt.conf" <<EOF
Dir::Etc "$apt/etc";
Dir::Etc::main "/dev/null";
Dir::Etc::sourcelist "/dev/null";
Dir::Etc::sourceparts "$apt/etc/sources.list.d";
Dir::State "$apt/state";
Dir::State::status "$apt/status";
Dir::Cache "$apt/cache";
Dir::Cache::archives "$apt/debs";
APT::Architecture "arm64";
APT::Architectures "arm64";
APT::Install-Recommends "false";
Debug::NoLocking "true";
EOF
: > "$apt/status"
export APT_CONFIG=$apt/apt.conf
apt-get -q update
apt-get -q -y --download-only install libwebkit2gtk-4.1-dev libgtk-3-dev libsoup-3.0-dev libc6-dev libgcc-14-dev

# the sysroot, made again when the packages change; links that point at / point inside it
ls "$apt/debs" | grep '\.deb$' > "$work/debs.new"
if ! cmp -s "$work/debs.new" "$work/debs.txt" 2>/dev/null || [ ! -d "$sysroot" ]; then
  rm -rf "$sysroot"
  mkdir -p "$sysroot/usr/lib" "$sysroot/usr/bin" "$sysroot/usr/sbin"
  ln -s usr/lib "$sysroot/lib"; ln -s usr/bin "$sysroot/bin"; ln -s usr/sbin "$sysroot/sbin"
  for d in "$apt/debs"/*.deb; do dpkg-deb -x "$d" "$sysroot"; done
  find "$sysroot" -type l -lname '/*' | while read -r l; do
    ln -snf "$(realpath -m --relative-to="$(dirname "$l")" "$sysroot$(readlink "$l")")" "$l"
  done
  mv "$work/debs.new" "$work/debs.txt"
fi
rm -f "$work/debs.new"

cat > "$work/cc" <<EOF
#!/bin/sh
exec clang --target=aarch64-linux-gnu --sysroot=$sysroot -fuse-ld=lld "\$@"
EOF
chmod 755 "$work/cc"

# v2026.11.07-9de300 (tools/current-tree-version.sh): the pages show it; the lite capability instead of the
# Start Center's and the app windows' (capabilities/lite.json)
VERSION=$("$desktop/../tools/current-tree-version.sh")
y=$(echo "$VERSION" | cut -c2-5); m=$(echo "$VERSION" | cut -c7-8); d=$(echo "$VERSION" | cut -c10-11)
m=${m#0}; d=${d#0}
if command -v rustup >/dev/null; then rustup target add "$target" >/dev/null; fi
cd "$desktop"
env VO_VERSION="$VERSION" \
  TAURI_CONFIG="{\"version\":\"$y.$m.$d\",\"app\":{\"security\":{\"capabilities\":[\"lite\"]}}}" \
  CARGO_PROFILE_RELEASE_STRIP=symbols \
  CARGO_TARGET_AARCH64_UNKNOWN_LINUX_GNU_LINKER="$work/cc" CC_aarch64_unknown_linux_gnu="$work/cc" \
  AR_aarch64_unknown_linux_gnu=llvm-ar PKG_CONFIG_ALLOW_CROSS=1 \
  PKG_CONFIG_SYSROOT_DIR_aarch64_unknown_linux_gnu="$sysroot" \
  PKG_CONFIG_LIBDIR_aarch64_unknown_linux_gnu="$sysroot/usr/lib/aarch64-linux-gnu/pkgconfig:$sysroot/usr/share/pkgconfig" \
  PKG_CONFIG_PATH_aarch64_unknown_linux_gnu= \
  cargo build --release --locked --features lite --target "$target"
ls -l "target/$target/release/vibeoffice"
