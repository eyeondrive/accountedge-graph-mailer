#!/bin/bash
# accountedge_graph_mailer — installer.
# Deploys send.sh + config, and swaps the OfficeLink Processor applet's script to
# send via Graph. Fully reversible: the ORIGINAL applet is backed up once (pristine)
# and never clobbered; uninstall.sh restores it.
set -o pipefail

SELF="$(cd "$(dirname "$0")" && pwd)"
APPLET_SRC="$SELF/script/officelink_applet.applescript"
SEND_SRC="$SELF/script/send.sh"

# Defaults (real locations). Overridable for testing.
APPLET="$HOME/Library/Application Support/OfficeLink/OfficeLink Processor.app"
TOOLDIR="$HOME/Library/Application Support/accountedge_graph_mailer"
CONFIG_SRC=""
DRY=0

while [ $# -gt 0 ]; do
  case "$1" in
    --config)  CONFIG_SRC="$2"; shift;;
    --applet)  APPLET="$2"; shift;;
    --tooldir) TOOLDIR="$2"; shift;;
    --dry-run) DRY=1;;
    *) echo "unknown arg: $1" >&2; exit 2;;
  esac; shift
done

BACKUP="$TOOLDIR/backup"
PRISTINE="$BACKUP/OfficeLink Processor.app.pristine"
MARKER="$TOOLDIR/.installed"
DISPATCH_SRC="$SELF/script/dispatch.sh"
# Working copy of the stock applet, beside ours, for non-email jobs (dispatch.sh calls it).
ORIGCOPY="$(dirname "$APPLET")/OfficeLink Processor (original).app"

die(){ echo "ERROR: $*" >&2; exit 1; }

[ -f "$APPLET_SRC" ]   || die "missing $APPLET_SRC"
[ -f "$SEND_SRC" ]     || die "missing $SEND_SRC"
[ -f "$DISPATCH_SRC" ] || die "missing $DISPATCH_SRC"
[ -d "$APPLET" ]     || die "OfficeLink applet not found at: $APPLET (has AccountEdge emailed at least once on this machine?)"

echo "── install plan ──"
echo "  applet   : $APPLET"
echo "  tool dir : $TOOLDIR"
echo "  pristine : $PRISTINE  (made once, never overwritten)"
echo "  original : $ORIGCOPY  (working copy for non-email jobs)"
echo "  config   : ${CONFIG_SRC:-<none — deploy config.yaml before going live>}"
if [ "$DRY" = 1 ]; then echo "[dry-run] no changes made."; exit 0; fi

mkdir -p "$TOOLDIR" "$BACKUP"

# 1) Pristine backup — ONCE. Never overwrite a good backup with an already-modified applet.
if [ ! -e "$PRISTINE" ]; then
  cp -R "$APPLET" "$PRISTINE" || die "pristine backup failed"
  echo "✓ backed up original applet → $PRISTINE"
else
  echo "✓ pristine backup already exists — kept as-is"
fi

# 1b) Working copy of the stock applet for pass-through (always from the pristine backup)
if [ ! -e "$ORIGCOPY" ]; then
  cp -R "$PRISTINE" "$ORIGCOPY" || die "could not create original-applet working copy"
  echo "✓ placed original applet for non-email jobs → $ORIGCOPY"
else
  echo "✓ original-applet working copy already present"
fi

# 2) Deploy the scripts
cp "$SEND_SRC" "$TOOLDIR/send.sh" && chmod 755 "$TOOLDIR/send.sh" || die "deploy send.sh failed"
cp "$DISPATCH_SRC" "$TOOLDIR/dispatch.sh" && chmod 755 "$TOOLDIR/dispatch.sh" || die "deploy dispatch.sh failed"
cp "$SELF/script/check.sh" "$TOOLDIR/check.sh" && chmod 755 "$TOOLDIR/check.sh" || die "deploy check.sh failed"
echo "✓ deployed send.sh + dispatch.sh + check.sh"

# 3) Deploy config (if given) — lock down the secret
if [ -n "$CONFIG_SRC" ]; then
  [ -f "$CONFIG_SRC" ] || die "config not found: $CONFIG_SRC"
  cp "$CONFIG_SRC" "$TOOLDIR/config.yaml" && chmod 600 "$TOOLDIR/config.yaml" || die "deploy config failed"
  echo "✓ deployed config.yaml (chmod 600)"
else
  echo "• no --config given; place config.yaml in $TOOLDIR before going live"
fi

# 4) Compile our applet script and swap it in
TMPD="$(mktemp -d)"; trap 'rm -rf "$TMPD"' EXIT
osacompile -o "$TMPD/main.scpt" "$APPLET_SRC" || die "osacompile failed"
TARGET="$APPLET/Contents/Resources/Scripts/main.scpt"
# The stock main.scpt is read-only; remove it first (we own the dir) then write ours.
rm -f "$TARGET" || die "could not remove old main.scpt"
cp "$TMPD/main.scpt" "$TARGET" || die "could not write new main.scpt"
chmod 644 "$TARGET"
echo "✓ replaced applet main.scpt"

# 5) Ad-hoc re-sign so the bundle is internally consistent after the swap
if codesign --force --deep --sign - "$APPLET" >/dev/null 2>&1; then
  echo "✓ re-signed applet (ad-hoc)"
else
  echo "• WARN: ad-hoc re-sign failed (a local applet usually still runs)"
fi

# 6) Marker
date '+installed %Y-%m-%d %H:%M:%S' > "$MARKER"
echo "── done ── uninstall with: $SELF/uninstall.sh --applet \"$APPLET\" --tooldir \"$TOOLDIR\""
