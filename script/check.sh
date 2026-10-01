#!/bin/bash
# accountedge_graph_mailer — connection check. Sends NOTHING.
# Signs in to Microsoft with your config and reports which mail permissions the
# app actually holds. Run this after setup, and whenever sending fails.
set -o pipefail

CONFIG="${1:-$HOME/Library/Application Support/accountedge_graph_mailer/config.yaml}"
[ -f "$CONFIG" ] || { echo "config not found: $CONFIG   (usage: check.sh [path/to/config.yaml])"; exit 3; }

yget(){ sed -n "s/^$1:[[:space:]]*//p" "$CONFIG" | head -1 | sed 's/[[:space:]][[:space:]]*#.*$//; s/^"//;s/"$//;s/^'\''//;s/'\''$//'; }
TENANT="$(yget tenant_id)"; CLIENT_ID="$(yget client_id)"
SECRET="$(yget client_secret)"; ITEM="$(yget keychain_item)"; SENDER="$(yget sender)"

if [ -z "$SECRET" ] && [ -n "$ITEM" ]; then
  SECRET="$(security find-generic-password -s "$ITEM" -w 2>/dev/null)" \
    || { echo "✗ Keychain item '$ITEM' not found in your login keychain"; exit 1; }
  echo "✓ secret read from Keychain item '$ITEM' (${#SECRET} characters, not shown)"
fi
[ -n "$TENANT" ] && [ -n "$CLIENT_ID" ] && [ -n "$SECRET" ] || { echo "✗ tenant_id, client_id, or the secret is missing"; exit 1; }

W="$(mktemp -d)"; trap 'rm -rf "$W"' EXIT
curl -sS -X POST "https://login.microsoftonline.com/$TENANT/oauth2/v2.0/token" -d grant_type=client_credentials \
  --data-urlencode "client_id=$CLIENT_ID" --data-urlencode "client_secret=$SECRET" \
  --data-urlencode "scope=https://graph.microsoft.com/.default" > "$W/r.json" 2>/dev/null
unset SECRET
TOKEN="$(plutil -extract access_token raw -o - "$W/r.json" 2>/dev/null)"
if [ -z "$TOKEN" ]; then
  echo "✗ sign-in failed: $(plutil -extract error_description raw -o - "$W/r.json" 2>/dev/null | head -c 300)"
  exit 1
fi
echo "✓ signed in to tenant $TENANT"

P="$(printf '%s' "$TOKEN" | cut -d. -f2 | tr '_-' '/+')"; while [ $(( ${#P} % 4 )) -ne 0 ]; do P="$P="; done
printf '%s' "$P" | base64 -D > "$W/c.json" 2>/dev/null; unset TOKEN
ROLES="$(plutil -extract roles json -o - "$W/c.json" 2>/dev/null)"
echo "  app         : $(plutil -extract app_displayname raw -o - "$W/c.json" 2>/dev/null)"
echo "  permissions : ${ROLES:-none}"
echo "  sends as    : ${SENDER:-<sender not set>}"

ok=1
case "$ROLES" in *'"Mail.Send"'*) ;; *) echo "✗ Mail.Send not granted (Application permission + admin consent)"; ok=0;; esac
case "$ROLES" in *'"Mail.ReadWrite"'*) ;; *) echo "✗ Mail.ReadWrite not granted (Application permission + admin consent)"; ok=0;; esac
[ "$ok" = 1 ] && echo "✓ ready to send" || exit 1
