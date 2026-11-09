#!/bin/sh
# VibeOffice Lite on the CardputerZero (docs/desktop.md, Lite edition), as M5's launcher (APPLaunch) starts it: the
# desktop user, no terminal. cage shows the app on the built-in screen (the panel's own DRM device) and reads
# no input itself: the app hands it the keyboard (src/cardputer.rs). When the app ends (Esc on its start page,
# File ▸ Exit, or Esc held: APPLaunch's SIGTERM ends cage), the screen goes back to the launcher. Its messages
# go to $XDG_RUNTIME_DIR, nothing to the disk.
log=${XDG_RUNTIME_DIR:-/tmp}/vibeoffice.log
# the desktop session's idle helpers stop while VibeOffice runs and start again once cage, which keeps this script's
# process (exec), has ended, however it ends (helpers.sh)
@APP_DIR@/helpers.sh stop $$
# the panel's DRM device by udev's name for it, resolved: cage's seat opens devices without following links
panel=$(readlink -f /dev/dri/cardputer-zero-internal)
export WLR_BACKENDS=drm WLR_DRM_DEVICES=$panel WLR_RENDERER=pixman LIBSEAT_BACKEND=noop
exec cage -d -- @APP_DIR@/vibeoffice "$@" >"$log" 2>&1
