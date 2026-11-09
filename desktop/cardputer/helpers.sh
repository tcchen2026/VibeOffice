#!/bin/sh
# VibeOffice Lite on the CardputerZero (docs/desktop.md, Lite edition): the HDMI desktop's panel and file manager (its
# wallpaper and icons) are stopped while VibeOffice runs, for their memory, and started again after it. launch.sh runs
# `stop PID` before cage (PID: its own process, which cage keeps): a watcher (`after PID`), a unit of the user's
# systemd outside the launcher's process tree, starts them again once cage has ended, whatever ended it (Esc on the
# start page, the launcher's held Esc, a crash). Only what ran is started again; a stop waits for a resume still going
# (a quick start again). labwc stays (Lectern's slide show uses its HDMI desktop), as do M5's launcher, the session
# bus and two kinds of helper: the sound server, which the launcher's own sounds use (stopped under it, its menu froze
# on the way back), and the desktop portals, which VibeOffice's GTK asks for as it starts (the bus starts them again).
#   helpers.sh stop [PID] | resume | after PID
set -u
state=${XDG_RUNTIME_DIR:-/tmp}/vibeoffice-helpers
# each under lwrespawn, which starts it again while labwc runs
programs='/usr/bin/wf-panel-pi /usr/bin/pcmanfm-pi'

# one stop or resume at a time (fd 9; what they start goes without it, or it would hold the lock for good)
lock() { exec 9> "$state.lock" && flock 9; }

stop() {
  lock
  mkdir "$state" 2>/dev/null || return 0   # stopped already, by a run still going
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
  for f in "$state"/program.*; do
    [ -f "$f" ] || continue
    xargs -0 -a "$state/environ.${f##*.}" "$0" spawn "$(cat "$f")" </dev/null >/dev/null 2>&1 9>&-
  done
  rm -rf "$state"
}

# PROGRAM again under lwrespawn, in the session's environment as it was (the arguments after it), in a session of
# its own (setsid -f returns at once) and a scope of its own in the user's systemd: the watcher's unit, which ends
# with its script, would take along all it started
spawn() {
  p=$1
  shift
  systemd-run --user --scope --collect --quiet -- setsid -f env -i "$@" /bin/sh /usr/bin/lwrespawn "$p" ||
    setsid -f env -i "$@" /bin/sh /usr/bin/lwrespawn "$p"
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
spawn) shift; spawn "$@" ;;
*) echo "usage: $0 stop [PID] | resume | after PID" >&2; exit 2 ;;
esac
