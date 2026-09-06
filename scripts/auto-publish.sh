#!/usr/bin/env bash
# ============================================================================
# auto-publish.sh — publish otomatis multi-platform (windows+macos dari CI,
# linux dari dist/ lokal) ke Print Hub. Untuk dijalankan via cron.
#
# Logika: cari run CI sukses terbaru per workflow (windows & macos) utk tag v*,
# publish bila versi-nya == config.json DAN run id belum tercatat di state file.
# Linux: publish dist/*.AppImage bila versi file/config == versi linux di hub.
# ============================================================================
set -uo pipefail

REPO="yudilee/trayprint"
HUB_URL="${HUB_URL:-https://print-hub.hartonomotor-group.com}"
TOKEN="${HUB_RELEASE_TOKEN:-$(cat /tmp/rel_token.txt 2>/dev/null || true)}"
STATE_FILE="${TRAYPRINT_STATE:-$HOME/.trayprint_publish_state}"
LOCK_FILE="${TRAYPRINT_LOCK:-/tmp/trayprint_auto_publish.lock}"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
LOG_FILE="$HOME/trayprint_auto_publish.log"

exec 9>"$LOCK_FILE"
flock -n 9 || { echo "$(date +%FT%T) lock dipegang proses lain — skip" >> "$LOG_FILE"; exit 0; }

log(){ echo "$(date '+%F %T') $*" >> "$LOG_FILE"; }
[ -n "$TOKEN" ] || { log "TOKEN hub tidak ada — skip"; exit 0; }
PAT=$(printf 'protocol=https\nhost=github.com\n\n' | git credential fill 2>/dev/null | sed -n 's/^password=//p')
[ -n "$PAT" ] || { log "credential github tidak ada — skip"; exit 0; }
touch "$STATE_FILE"

API="https://api.github.com/repos/${REPO}"
AUTH="Authorization: Bearer ${PAT}"
VERSION=$(python3 -c "import json;print(json.load(open('$REPO_DIR/config.json')).get('version','3.0.0'))")
log "=== cek publish v${VERSION} ==="

publish_artifact() { # $1=run_id $2=artifact_name $3=platform $4=ext
  local run_id="$1" art_name="$2" platform="$3" ext="$4" key
  key="${platform}:${run_id}"
  grep -q "^${key}$" "$STATE_FILE" && { log "  ${platform} run ${run_id} sudah dipublish"; return 0; }
  # Anti-duplikat: kalau versi ini sudah jadi latest di hub utk platform tsb, skip
  local cur
  cur=$(curl -s -m 15 "${HUB_URL}/api/v1/agents/version?platform=${platform}" | python3 -c "import sys,json;print((json.load(sys.stdin).get('latest_version') or ''))" 2>/dev/null)
  if [ "$cur" = "$VERSION" ]; then echo "$key" >> "$STATE_FILE"; log "  ${platform} v${VERSION} sudah latest di hub — catat & skip"; return 0; fi
  ART=$(curl -s -H "$AUTH" "${API}/actions/runs/${run_id}/artifacts" \
    | python3 -c "import sys,json;d=json.load(sys.stdin);print(next((a['id'] for a in d['artifacts'] if a['name']=='${art_name}'),''))")
  [ -n "$ART" ] || { log "  ${platform}: artifact ${art_name} tidak ada di run ${run_id}"; return 1; }
  local tmp; tmp=$(mktemp -d)
  curl -sL -H "$AUTH" "${API}/actions/artifacts/${ART}/zip" -o "$tmp/art.zip"
  python3 - "$tmp" "$ext" <<'PY'
import sys, zipfile, glob, os, shutil
d, ext = sys.argv[1], sys.argv[2]
with zipfile.ZipFile(os.path.join(d,'art.zip')) as z: z.extractall(os.path.join(d,'out'))
f=glob.glob(os.path.join(d,'out','**','*'+ext), recursive=True)
if f: shutil.copy(f[0], os.path.join(d,'pkg'+ext)); print(os.path.getsize(os.path.join(d,'pkg'+ext)))
else: print('NONE')
PY
  local size; size=$(python3 -c "import os;print(os.path.getsize('$tmp/pkg$ext'))" 2>/dev/null || echo 0)
  [ "$size" -gt 0 ] || { log "  ${platform}: file tidak ditemukan di artifact"; rm -rf "$tmp"; return 1; }
  local code
  code=$(curl -sS -o "$tmp/resp.json" -w '%{http_code}' -X POST "${HUB_URL}/api/v1/releases" \
    -H "X-Release-Token: ${TOKEN}" \
    -F "version=${VERSION}" -F "platform=${platform}" -F "channel=stable" \
    -F "installer_file=@${tmp}/pkg${ext}")
  rm -rf "$tmp"
  if [ "$code" = "201" ]; then echo "$key" >> "$STATE_FILE"; log "  ${platform} v${VERSION} PUBLISH OK (run ${run_id})"; return 0
  else log "  ${platform} gagal HTTP ${code} (run ${run_id})"; return 1; fi
}

# Windows & macOS dari CI (cari run sukses terbaru per workflow utk tag v*)
RUNS=$(curl -s -H "$AUTH" "${API}/actions/runs?per_page=20")
WIN_RUN=$(echo "$RUNS" | python3 -c "
import sys,json
d=json.load(sys.stdin)
print(next((r['id'] for r in d['workflow_runs']
  if r['status']=='completed' and r['conclusion']=='success'
  and r['name'].startswith('Windows Build')), ''))")
MAC_RUN=$(echo "$RUNS" | python3 -c "
import sys,json
d=json.load(sys.stdin)
print(next((r['id'] for r in d['workflow_runs']
  if r['status']=='completed' and r['conclusion']=='success'
  and r['name'].startswith('macOS Build')), ''))")
[ -n "$WIN_RUN" ] && publish_artifact "$WIN_RUN" trayprint-windows windows .exe
[ -n "$MAC_RUN" ] && publish_artifact "$MAC_RUN" trayprint-macos macos .dmg

# Linux dari dist/ lokal (AppImage)
APPIMAGE=$(ls "$REPO_DIR"/dist/TrayPrint-*.AppImage 2>/dev/null | head -1)
if [ -n "$APPIMAGE" ]; then
  CUR=$(curl -s -m 15 "${HUB_URL}/api/v1/agents/version?platform=linux" | python3 -c "import sys,json;print((json.load(sys.stdin).get('latest_version') or ''))" 2>/dev/null)
  if [ "$CUR" != "$VERSION" ]; then
    code=$(curl -sS -o /tmp/tp_linux_resp.json -w '%{http_code}' -X POST "${HUB_URL}/api/v1/releases" \
      -H "X-Release-Token: ${TOKEN}" -F "version=${VERSION}" -F "platform=linux" \
      -F "channel=stable" -F "installer_file=@${APPIMAGE}")
    [ "$code" = "201" ] && log "  linux v${VERSION} PUBLISH OK (${APPIMAGE})" || log "  linux gagal HTTP ${code}"
  else log "  linux v${VERSION} sudah yang terbaru di hub"; fi
else log "  AppImage linux tidak ada di dist/"; fi

log "=== selesai ==="
exit 0
