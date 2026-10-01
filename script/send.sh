#!/bin/bash
# accountedge_graph_mailer — send AccountEdge "OfficeLink" emails via Microsoft Graph.
#
# AccountEdge (native or CrossOver-bottled) writes an "OfficeLink Data" handoff file
# (Subject / Content / Address / CCAddress / BCAddress / Attachment / EmailAction),
# then launches the OfficeLink Processor applet to send it. dispatch.sh routes email
# jobs here; we read the same file and send via Graph — no mail client involved.
#
# ONE route for every email (so every send exercises all of it):
#   create draft -> attach each file (direct if < 3 MB, upload session if larger)
#   -> send draft (Graph files it in Sent Items).
# Needs Graph application permissions Mail.ReadWrite (draft + attachments) and
# Mail.Send (send), restricted to the sending mailbox(es) by an Application Access Policy.
#
# Generic + config-driven: nothing site-specific lives here; identity is in the config.
# Deps: bash, curl, base64, plutil, dd, stat (all stock on macOS). No python/jq.
set -o pipefail

SELF="$(cd "$(dirname "$0")" && pwd)"
DRY_RUN=0
CONFIG=""
GRAPH="https://graph.microsoft.com/v1.0"
DIRECT_MAX=3145728        # < 3 MB: attach in one request; otherwise use an upload session
CHUNK=3276800             # upload-session chunk: 10 x 320 KiB (must be a multiple of 320 KiB, < 4 MB)

usage(){ echo "usage: send.sh [--config PATH] [--dry-run]"; }

while [ $# -gt 0 ]; do
  case "$1" in
    --dry-run) DRY_RUN=1 ;;
    --config)  CONFIG="$2"; shift ;;
    -h|--help) usage; exit 0 ;;
    *) echo "unknown arg: $1" >&2; usage; exit 2 ;;
  esac
  shift
done

# ---- locate config (first hit wins) ----
if [ -z "$CONFIG" ]; then
  for c in \
    "$SELF/../config.yaml" \
    "$SELF/config.yaml" \
    "$HOME/Library/Application Support/accountedge_graph_mailer/config.yaml"
  do
    [ -f "$c" ] && { CONFIG="$c"; break; }
  done
fi
[ -n "$CONFIG" ] && [ -f "$CONFIG" ] || { echo "no config found (looked for config.yaml)" >&2; exit 3; }

# ---- flat-YAML getter (flat "key: value" only) ----
yget(){ sed -n "s/^$1:[[:space:]]*//p" "$CONFIG" | head -1 | sed 's/[[:space:]][[:space:]]*#.*$//; s/^"//;s/"$//;s/^'\''//;s/'\''$//'; }
expand_tilde(){ case "$1" in "~/"*) printf '%s' "${HOME}/${1#\~/}";; *) printf '%s' "$1";; esac; }

TENANT="$(yget tenant_id)"
CLIENT_ID="$(yget client_id)"
CLIENT_SECRET="$(yget client_secret)"
KEYCHAIN_ITEM="$(yget keychain_item)"
SENDER="$(yget sender)"
ALERT_TO="$(yget alert_to)"
OL_DATA="$(expand_tilde "$(yget officelink_data)")"
LOG="$(expand_tilde "$(yget log_path)")"
[ -n "$OL_DATA" ] || OL_DATA="$HOME/Library/Application Support/OfficeLink/OfficeLink Data"
[ -n "$LOG" ] || LOG="$HOME/Library/Logs/accountedge_graph_mailer.log"

log(){ printf '%s %s\n' "$(date '+%Y-%m-%d %H:%M:%S')" "$*" >> "$LOG" 2>/dev/null; }

alert_fail(){
  # Fail LOUD: visible to the operator now, never silent.
  local msg="$1"
  osascript -e "display dialog \"AccountEdge email FAILED:\n${msg}\n\nThe email was NOT sent. Contact support.\" with title \"AccountEdge Graph Mailer\" buttons {\"OK\"} default button 1 with icon caution" >/dev/null 2>&1
}

# ---- resolve secret (config value, else macOS Keychain) ----
if [ -z "$CLIENT_SECRET" ] && [ -n "$KEYCHAIN_ITEM" ]; then
  CLIENT_SECRET="$(security find-generic-password -s "$KEYCHAIN_ITEM" -w 2>/dev/null)"
fi

# ---- parse OfficeLink Data (repeated keys; CRLF-safe) ----
[ -f "$OL_DATA" ] || { echo "OfficeLink Data not found: $OL_DATA" >&2; log "ERROR no OfficeLink Data at $OL_DATA"; exit 4; }

SUBJECT=""; CONTENT=""; ACTION=""; APPID=""
TO=(); CC=(); BCC=(); ATT=()
while IFS= read -r line || [ -n "$line" ]; do
  line="${line%$'\r'}"
  case "$line" in
    AppID=*)       APPID="${line#AppID=}" ;;
    Subject=*)     SUBJECT="${line#Subject=}" ;;
    Content=*)     CONTENT="${line#Content=}" ;;
    Address=*)     TO+=("${line#Address=}") ;;
    CCAddress=*)   CC+=("${line#CCAddress=}") ;;
    BCAddress=*)   BCC+=("${line#BCAddress=}") ;;
    Attachment=*)  ATT+=("${line#Attachment=}") ;;
    EmailAction=*) ACTION="${line#EmailAction=}" ;;
  esac
done < "$OL_DATA"

# Email jobs only — anything else belongs to the original applet (see dispatch.sh).
if [ "$APPID" != "MYOB.EMAIL" ]; then
  log "skip: AppID='$APPID' is not an email job"
  echo "AppID='$APPID' is not an email job — send.sh only handles MYOB.EMAIL." >&2
  exit 0
fi
if [ "$ACTION" != "Send" ]; then
  log "skip: EmailAction='$ACTION' (not Send)"
  echo "EmailAction='$ACTION' (not Send) — nothing to do." >&2
  exit 0
fi
if [ ${#TO[@]} -eq 0 ]; then
  log "ERROR no To recipients"; echo "no To recipients in OfficeLink Data" >&2
  [ "$DRY_RUN" = 1 ] || alert_fail "the email has no recipient."
  exit 4
fi

# ---- attachments: every listed file must exist (never send an invoice email without its invoice) ----
TOTAL=0; MISSING=()
for f in "${ATT[@]}"; do
  if [ -f "$f" ]; then TOTAL=$(( TOTAL + $(stat -f %z "$f") )); else MISSING+=("$f"); fi
done
mb(){ awk -v b="$1" 'BEGIN{printf "%.1f MB", b/1048576}'; }

# ---- JSON builders (plist -> json via plutil = correct escaping, no jq) ----
xml_escape(){ printf '%s' "$1" | sed 's/&/\&amp;/g; s/</\&lt;/g; s/>/\&gt;/g'; }

mime_type(){
  case "$(printf '%s' "$1" | tr 'A-Z' 'a-z')" in
    *.pdf) echo "application/pdf";;
    *.jpg|*.jpeg) echo "image/jpeg";;
    *.png) echo "image/png";;
    *.gif) echo "image/gif";;
    *.tif|*.tiff) echo "image/tiff";;
    *.txt) echo "text/plain";; *.csv) echo "text/csv";;
    *.doc) echo "application/msword";;
    *.docx) echo "application/vnd.openxmlformats-officedocument.wordprocessingml.document";;
    *.xls) echo "application/vnd.ms-excel";;
    *.xlsx) echo "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";;
    *.zip) echo "application/zip";;
    *) echo "application/octet-stream";;
  esac
}

PLIST_HEAD='<?xml version="1.0" encoding="UTF-8"?><!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd"><plist version="1.0">'
to_json(){ plutil -convert json -o "$2" "$1" 2>/dev/null; }   # $1 plist in, $2 json out

recipients_plist(){  # args: addresses -> <array> of {emailAddress:{address}}
  printf '<array>'
  for a in "$@"; do
    printf '<dict><key>emailAddress</key><dict><key>address</key><string>%s</string></dict></dict>' "$(xml_escape "$a")"
  done
  printf '</array>'
}

build_message_json(){  # $1 out.json — message WITHOUT attachments (they're added to the draft)
  {
    printf '%s<dict>' "$PLIST_HEAD"
    printf '<key>subject</key><string>%s</string>' "$(xml_escape "$SUBJECT")"
    printf '<key>body</key><dict><key>contentType</key><string>Text</string><key>content</key><string>%s</string></dict>' "$(xml_escape "$CONTENT")"
    printf '<key>toRecipients</key>%s' "$(recipients_plist "${TO[@]}")"
    [ ${#CC[@]}  -gt 0 ] && printf '<key>ccRecipients</key>%s'  "$(recipients_plist "${CC[@]}")"
    [ ${#BCC[@]} -gt 0 ] && printf '<key>bccRecipients</key>%s' "$(recipients_plist "${BCC[@]}")"
    printf '</dict></plist>'
  } > "$TMP/msg.plist" && to_json "$TMP/msg.plist" "$1"
}

build_attach_json(){   # $1 file, $2 out.json — small attachment, sent inline as base64
  {
    printf '%s<dict>' "$PLIST_HEAD"
    printf '<key>@odata.type</key><string>#microsoft.graph.fileAttachment</string>'
    printf '<key>name</key><string>%s</string>' "$(xml_escape "$(basename "$1")")"
    printf '<key>contentType</key><string>%s</string>' "$(mime_type "$1")"
    printf '<key>contentBytes</key><string>%s</string>' "$(base64 < "$1" | tr -d '\n')"
    printf '</dict></plist>'
  } > "$TMP/att.plist" && to_json "$TMP/att.plist" "$2"
}

build_session_json(){  # $1 file, $2 size, $3 out.json — upload-session request for a large file
  {
    printf '%s<dict><key>AttachmentItem</key><dict>' "$PLIST_HEAD"
    printf '<key>attachmentType</key><string>file</string>'
    printf '<key>name</key><string>%s</string>' "$(xml_escape "$(basename "$1")")"
    printf '<key>size</key><integer>%s</integer>' "$2"
    printf '<key>contentType</key><string>%s</string>' "$(mime_type "$1")"
    printf '</dict></dict></plist>'
  } > "$TMP/ses.plist" && to_json "$TMP/ses.plist" "$3"
}

TMP="$(mktemp -d -t aegm)"
trap 'rm -rf "$TMP"' EXIT

# ---- dry run: show the plan, build the JSON, contact nothing ----
if [ "$DRY_RUN" = 1 ]; then
  echo "── parsed OfficeLink Data ──"
  echo "Subject : $SUBJECT"
  echo "To      : ${TO[*]}"
  [ ${#CC[@]}  -gt 0 ] && echo "Cc      : ${CC[*]}"
  [ ${#BCC[@]} -gt 0 ] && echo "Bcc     : ${BCC[*]}"
  echo "Attach  : ${#ATT[@]} file(s), $(mb $TOTAL) total"
  for f in "${ATT[@]}"; do
    if [ -f "$f" ]; then
      s=$(stat -f %z "$f")
      if [ "$s" -lt "$DIRECT_MAX" ]; then how="direct"; else how="upload session, $(( (s + CHUNK - 1) / CHUNK )) chunk(s)"; fi
      echo "          + $(basename "$f")  $(mb $s)  [$(mime_type "$f")]  -> $how"
    else
      echo "          - $(basename "$f")  MISSING ON DISK -> would refuse to send"
    fi
  done
  build_message_json "$TMP/msg.json" && plutil -convert xml1 -o /dev/null "$TMP/msg.json" \
    && echo "[message JSON: VALID]" || echo "[message JSON: INVALID]"
  echo "Route   : create draft -> attach -> send (Sent Items copy kept by Graph)"
  echo "[dry-run] Graph not contacted."
  exit 0
fi

if [ ${#MISSING[@]} -gt 0 ]; then
  log "ERROR attachment(s) missing: ${MISSING[*]}"
  alert_fail "an attachment could not be found:\n$(basename "${MISSING[0]}")"
  exit 4
fi

# ---- live ----
for pair in "tenant_id:$TENANT" "client_id:$CLIENT_ID" "client_secret:$CLIENT_SECRET" "sender:$SENDER"; do
  [ -n "${pair#*:}" ] || { echo "missing config: ${pair%%:*}" >&2; log "ERROR missing config ${pair%%:*}"; alert_fail "the mailer is not configured (${pair%%:*})."; exit 6; }
done

# OAuth: client credentials
curl -sS -X POST "https://login.microsoftonline.com/$TENANT/oauth2/v2.0/token" \
  -d "grant_type=client_credentials" \
  --data-urlencode "client_id=$CLIENT_ID" \
  --data-urlencode "client_secret=$CLIENT_SECRET" \
  --data-urlencode "scope=https://graph.microsoft.com/.default" > "$TMP/tok.json" 2>/dev/null
TOKEN="$(plutil -extract access_token raw -o - "$TMP/tok.json" 2>/dev/null)"
if [ -z "$TOKEN" ]; then
  log "ERROR token failure: $(head -c 300 "$TMP/tok.json" 2>/dev/null)"
  alert_fail "could not sign in to Microsoft (token request failed)."
  exit 7
fi

# graph METHOD URL [json-file] -> prints HTTP code; response body in $TMP/resp.json
graph(){
  if [ -n "$3" ]; then
    curl -sS -o "$TMP/resp.json" -w '%{http_code}' -X "$1" "$2" \
      -H "Authorization: Bearer $TOKEN" -H "Content-Type: application/json" --data-binary @"$3" 2>/dev/null
  else
    curl -sS -o "$TMP/resp.json" -w '%{http_code}' -X "$1" "$2" \
      -H "Authorization: Bearer $TOKEN" -H "Content-Length: 0" 2>/dev/null
  fi
}
graph_err(){ head -c 400 "$TMP/resp.json" 2>/dev/null; }

DRAFT_ID=""
abort(){  # $1 = user-facing reason; delete the half-built draft so nothing lingers in Drafts
  log "ERROR $1 :: $(graph_err)"
  if [ -n "$DRAFT_ID" ]; then
    curl -sS -o /dev/null -X DELETE "$GRAPH/users/$SENDER/messages/$DRAFT_ID" -H "Authorization: Bearer $TOKEN" 2>/dev/null
  fi
  alert_fail "$1"
  exit 8
}

# 1) draft
build_message_json "$TMP/msg.json" || abort "internal error assembling the message."
HTTP=$(graph POST "$GRAPH/users/$SENDER/messages" "$TMP/msg.json")
[ "$HTTP" = "201" ] || abort "Microsoft would not create the message (HTTP $HTTP)."
DRAFT_ID="$(plutil -extract id raw -o - "$TMP/resp.json" 2>/dev/null)"
[ -n "$DRAFT_ID" ] || abort "Microsoft created the message but returned no ID."

# 2) attachments
for f in "${ATT[@]}"; do
  name="$(basename "$f")"; size=$(stat -f %z "$f")
  if [ "$size" -lt "$DIRECT_MAX" ]; then
    build_attach_json "$f" "$TMP/att.json" || abort "internal error preparing $name."
    HTTP=$(graph POST "$GRAPH/users/$SENDER/messages/$DRAFT_ID/attachments" "$TMP/att.json")
    [ "$HTTP" = "201" ] || abort "could not attach $name (HTTP $HTTP)."
  else
    build_session_json "$f" "$size" "$TMP/ses.json" || abort "internal error preparing $name."
    HTTP=$(graph POST "$GRAPH/users/$SENDER/messages/$DRAFT_ID/attachments/createUploadSession" "$TMP/ses.json")
    [ "$HTTP" = "201" ] || abort "could not start uploading $name ($(mb $size), HTTP $HTTP). Total attachments: $(mb $TOTAL)."
    URL="$(plutil -extract uploadUrl raw -o - "$TMP/resp.json" 2>/dev/null)"
    [ -n "$URL" ] || abort "no upload address returned for $name."
    off=0
    while [ "$off" -lt "$size" ]; do
      dd if="$f" of="$TMP/chunk" bs=327680 skip=$(( off / 327680 )) count=$(( CHUNK / 327680 )) 2>/dev/null
      n=$(stat -f %z "$TMP/chunk"); end=$(( off + n - 1 ))
      HTTP=$(curl -sS -o "$TMP/resp.json" -w '%{http_code}' -X PUT "$URL" \
        -H "Content-Type: application/octet-stream" -H "Content-Range: bytes $off-$end/$size" \
        --data-binary @"$TMP/chunk" 2>/dev/null)
      case "$HTTP" in 200|201) ;; *) abort "upload of $name failed at $(mb $off) of $(mb $size) (HTTP $HTTP)." ;; esac
      off=$(( end + 1 ))
    done
  fi
done

# 3) send (Graph moves it to Sent Items)
HTTP=$(graph POST "$GRAPH/users/$SENDER/messages/$DRAFT_ID/send")
[ "$HTTP" = "202" ] || abort "Microsoft would not send the message (HTTP $HTTP). Total attachments: $(mb $TOTAL)."

log "OK sent subj='$SUBJECT' to=${TO[*]} cc=${CC[*]} bcc=${BCC[*]} att=${#ATT[@]} ($(mb $TOTAL))"
exit 0
