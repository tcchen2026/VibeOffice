#!/bin/sh
# Installs the app, its licence, helpers.sh and the launcher's entry point (launch.sh) under /usr/share/APPLaunch;
# the icon is the desktop app's (icons/icon.png). Run by pack-deb.sh after build.sh, with STAGE, INSTALL_PREFIX and
# APP_INSTALL_DIR.
set -eu
desktop=$(cd "$(dirname "$0")/../.." && pwd)
app=$STAGE$APP_INSTALL_DIR
install -D -m 0755 "$desktop/target/aarch64-unknown-linux-gnu/release/vibeoffice" "$app/vibeoffice"
install -D -m 0644 "$desktop/../LICENSE" "$app/LICENSE"
install -D -m 0755 "$desktop/cardputer/helpers.sh" "$app/helpers.sh"
sed "s|@APP_DIR@|$APP_INSTALL_DIR|g" "$desktop/cardputer/launch.sh" > "$STAGE$INSTALL_PREFIX/bin/$PKG_NAME"
chmod 0755 "$STAGE$INSTALL_PREFIX/bin/$PKG_NAME"
install -D -m 0644 "$desktop/icons/icon.png" "$STAGE$INSTALL_PREFIX/share/images/$PKG_NAME.png"
