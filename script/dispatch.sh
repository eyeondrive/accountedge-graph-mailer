#!/bin/bash
# accountedge_graph_mailer — OfficeLink request router.
#
# AccountEdge uses one OfficeLink applet for several jobs, naming the job in the
# "AppID=" line of "OfficeLink Data" (MYOB.EMAIL, MYOB.EXISTS, Excel/Word export, ...).
# We take over ONLY email. Everything else is handed, unchanged, to the original
# applet, so probes and Excel/Word exports keep working exactly as before.
set -o pipefail

SELF="$(cd "$(dirname "$0")" && pwd)"
CONFIG=""; DRY_RUN=0
ARGS=("$@")
while [ $# -gt 0 ]; do
  case "$1" in
    --config)  CONFIG="$2"; shift ;;
    --dry-run) DRY_RUN=1 ;;
  esac
  shift
done
[ -z "$CONFIG" ] && CONFIG="$SELF/config.yaml"

yget(){ [ -f "$CONFIG" ] && sed -n "s/^$1:[[:space:]]*//p" "$CONFIG" | head -1 | sed 's/[[:space:]][[:space:]]*#.*$//; s/^"//;s/"$//;s/^'\''//;s/'\''$//'; }
expand_tilde(){ case "$1" in "~/"*) printf '%s' "${HOME}/${1#\~/}";; *) printf '%s' "$1";; esac; }

OL_DATA="$(expand_tilde "$(yget officelink_data)")"
ORIG="$(expand_tilde "$(yget original_applet)")"
LOG="$(expand_tilde "$(yget log_path)")"
[ -n "$OL_DATA" ] || OL_DATA="$HOME/Library/Application Support/OfficeLink/OfficeLink Data"
[ -n "$ORIG" ]    || ORIG="$HOME/Library/Application Support/OfficeLink/OfficeLink Processor (original).app"
[ -n "$LOG" ]     || LOG="$HOME/Library/Logs/accountedge_graph_mailer.log"

log(){ printf '%s %s\n' "$(date '+%Y-%m-%d %H:%M:%S')" "$*" >> "$LOG" 2>/dev/null; }

APPID="$(tr -d '\r' < "$OL_DATA" 2>/dev/null | sed -n 's/^AppID=//p' | head -1)"

if [ "$APPID" = "MYOB.EMAIL" ]; then
  [ "$DRY_RUN" = 1 ] && { echo "AppID=$APPID -> route: send.sh (Microsoft Graph)"; }
  log "dispatch AppID=$APPID -> send.sh"
  exec /bin/bash "$SELF/send.sh" "${ARGS[@]}"
fi

# Not email: pass through to the original applet, untouched.
if [ "$DRY_RUN" = 1 ]; then
  echo "AppID=${APPID:-<none>} -> route: original applet ($ORIG)"
  [ -x "$ORIG/Contents/MacOS/applet" ] && echo "original applet present: yes" || echo "original applet present: NO"
  exit 0
fi
log "dispatch AppID=${APPID:-<none>} -> original applet"
if [ -x "$ORIG/Contents/MacOS/applet" ]; then
  "$ORIG/Contents/MacOS/applet"
  rc=$?
  log "original applet exited rc=$rc"
  exit $rc
fi
log "ERROR original applet missing at $ORIG"
osascript -e "display dialog \"AccountEdge OfficeLink: the original helper is missing, so this action (${APPID:-unknown}) could not run.\n\nContact support.\" with title \"AccountEdge Graph Mailer\" buttons {\"OK\"} default button 1 with icon caution" >/dev/null 2>&1
exit 9
