#!/bin/bash
# accountedge_graph_mailer — uninstaller. Restores the original OfficeLink applet
# from the pristine backup. Leaves the tool dir (with your secret/config) in place
# unless you remove it yourself.
set -o pipefail

APPLET="$HOME/Library/Application Support/OfficeLink/OfficeLink Processor.app"
TOOLDIR="$HOME/Library/Application Support/accountedge_graph_mailer"

while [ $# -gt 0 ]; do
  case "$1" in
    --applet)  APPLET="$2"; shift;;
    --tooldir) TOOLDIR="$2"; shift;;
    *) echo "unknown arg: $1" >&2; exit 2;;
  esac; shift
done

PRISTINE="$TOOLDIR/backup/OfficeLink Processor.app.pristine"
ORIGCOPY="$(dirname "$APPLET")/OfficeLink Processor (original).app"
die(){ echo "ERROR: $*" >&2; exit 1; }

[ -e "$PRISTINE" ] || die "no pristine backup at $PRISTINE — refusing to guess; original NOT restored."

echo "Restoring original OfficeLink Processor.app from pristine backup…"
rm -rf "$APPLET"            || die "could not remove current applet"
cp -R "$PRISTINE" "$APPLET" || die "restore failed"
rm -f "$TOOLDIR/.installed"
echo "✓ original applet restored."
if [ -e "$ORIGCOPY" ]; then
  rm -rf "$ORIGCOPY" && echo "✓ removed the pass-through working copy"
fi
echo "• Tool dir left in place (holds your config/secret): $TOOLDIR"
echo "  Remove it yourself if desired:  rm -rf \"$TOOLDIR\""
