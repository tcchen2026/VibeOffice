#!/bin/sh
# VibeOffice Lite on the CardputerZero (docs/desktop.md, Lite edition): the desktop session's idle helpers are stopped
# while VibeOffice runs, for their memory, and started again after it. launch.sh runs `stop PID` before cage (PID:
# its own process, which cage keeps): a watcher (`after PID`), a unit of the user's systemd outside the launcher's
# process tree, starts them again once cage has ended, whatever ended it (Esc, File ▸ Exit, the launcher's held Esc, a
# crash). Only what ran is started again; a stop waits for a resume still going (a quick start again). labwc stays
# (Lectern's slide show uses its HDMI desktop), as do M5's launcher and the session bus.
#   helpers.sh stop [PID] | resume | after PID
set -u
state=${XDG_RUNTIME_DIR:-/tmp}/vibeoffice-helpers
# the sound server and the desktop portals: systemd user units, sockets first so that nothing starts them again
units='pipewire-pulse.socket pipewire.socket pipewire-pulse.service pipewire.service wireplumber.service
filter-chain.service xdg-desktop-portal-wlr.service xdg-desktop-portal-gtk.service xdg-desktop-portal.service'
# the HDMI desktop's panel and file manager (its wallpaper and icons), each under lwrespawn, which starts it again
# while labwc runs
programs='/usr/bin/wf-panel-pi /usr/bin/pcmanfm-pi'

# one stop or resume at a time
lock() { exec 9> "$state.lock" && flock 9; }

stop() {
  lock
  mkdir "$state" 2>/dev/null || return 0   # stopped already, by a run still going
  for u in $units; do
    [ "$(systemctl --user is-active "$u" 2>/dev/null)" = active ] && echo "$u" >> "$state/units"
  done
  [ -f "$state/units" ] && systemctl --user stop $(cat "$state/units")
  n=0
  for p in $programs; do
    for w in $(pgrep -u "$(id -u)" -f "lwrespawn $p\$"); do
      n=$((n + 1))
      echo "$p" > "$state/program.$n"
      cp "/proc/$w/environ" "$state/environ.$n"
      kids=$(pgrep -P "$w")
      kill "$w"   # the wrapper first, or it starts the program again
      for k in $kids; do pkill -P "$k"; kill "$k"; done
    done
  done
}

resume() {
  lock
  [ -d "$state" ] || return 0
  [ -f "$state/units" ] && systemctl --user start $(cat "$state/units")
  for f in "$state"/program.*; do
    [ -f "$f" ] || continue
    # in the session's environment as it was, in a session of its own (setsid -f returns at once)
    xargs -0 -a "$state/environ.${f##*.}" sh -c 'p=$1; shift; exec setsid -f env -i "$@" /bin/sh /usr/bin/lwrespawn "$p"' \
      sh "$(cat "$f")" </dev/null >/dev/null 2>&1
  done
  rm -rf "$state"
}

# the watcher for PID: a transient unit of the user's systemd (M5's launcher adopts what it leaves behind and never
# reaps it), else a session of its own
watch() {
  systemd-run --user --collect --quiet --unit="vibeoffice-helpers-$1" "$0" after "$1" 9>&- 2>/dev/null ||
    { setsid "$0" after "$1" </dev/null >/dev/null 2>&1 9>&- & }
}

case ${1:-} in
stop) stop; [ -z "${2:-}" ] || watch "$2" ;;
resume) resume ;;
after) tail --pid="$2" -f /dev/null; resume ;;
*) echo "usage: $0 stop [PID] | resume | after PID" >&2; exit 2 ;;
esac
