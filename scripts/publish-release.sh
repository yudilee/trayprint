#!/usr/bin/env bash
# ============================================================================
# publish-release.sh — Publish TrayPrint Windows EXE ke Print Hub (dari IP lokal)
#
# Kenapa: Cloudflare men-challenge traffic GitHub Actions (403 "Just a moment"),
# jadi publish dilakukan dari mesin lokal/on-prem yang IP-nya tidak diblokir.
#
# Cara pakai:
#   ./scripts/publish-release.sh [RUN_ID]
#     RUN_ID opsional; default = run workflow windows-build terbaru yang SUCCESS.
#   Env: HUB_RELEASE_TOKEN (token /api/v1/releases) — wajib.
#        HUB_URL (default https://print-hub.hartonomotor-group.com)
# ============================================================================
set -euo pipefail

REPO="yudilee/trayprint"
HUB_URL="${HUB_URL:-https://print-hub.hartonomotor-group.com}"
RUN_ID="${1:-}"
TOKEN="${HUB_RELEASE_TOKEN:-$(cat /tmp/rel_token.txt 2>/dev/null || true)}"
[ -n "$TOKEN" ] || { echo "ERROR: HUB_RELEASE_TOKEN tidak ada"; exit 1; }

# PAT GitHub dari credential store (untuk download artifact)
PAT=$(printf 'protocol=https\nhost=github.com\n\n' | git credential fill 2>/dev/null | sed -n 's/^password=//p')
[ -n "$PAT" ] || { echo "ERROR: credential github tidak ditemukan"; exit 1; }

API="https://api.github.com/repos/${REPO}"
AUTH="Authorization: Bearer ${PAT}"

# 1) Cari run terbaru yang sukses (kalau RUN_ID tidak diberikan)
if [ -z "$RUN_ID" ]; then
  RUN_ID=$(curl -s -H "$AUTH" "${API}/actions/runs?per_page=10" \
    | python3 -c "import sys,json;d=json.load(sys.stdin);print(next((r['id'] for r in d['workflow_runs'] if r['conclusion']=='success' and r['status']=='completed'),''))")
  echo "run sukses terbaru: ${RUN_ID:-<tidak ada>}"
fi
[ -n "$RUN_ID" ] || { echo "Tidak ada run sukses."; exit 1; }

# 2) Ambil artifact 'trayprint-windows'
ART=$(curl -s -H "$AUTH" "${API}/actions/runs/${RUN_ID}/artifacts" \
  | python3 -c "import sys,json;d=json.load(sys.stdin);print(next((a['id'] for a in d['artifacts'] if a['name']=='trayprint-windows'),''))")
[ -n "$ART" ] || { echo "ERROR: artifact trayprint-windows tidak ada di run ${RUN_ID}"; exit 1; }
echo "artifact id: ${ART} — mengunduh..."

TMPD=$(mktemp -d)
curl -sL -H "$AUTH" -H "Accept: application/vnd.github+json" \
  "${API}/actions/artifacts/${ART}/zip" -o "${TMPD}/art.zip"
python3 - "$TMPD" <<'PY'
import sys, zipfile, glob, os, shutil
d=sys.argv[1]
with zipfile.ZipFile(os.path.join(d,'art.zip')) as z:
    z.extractall(os.path.join(d,'out'))
exe=glob.glob(os.path.join(d,'out','**','*.exe'), recursive=True)
if exe: shutil.copy(exe[0], os.path.join(d,'trayprint.exe'))
msi=glob.glob(os.path.join(d,'out','**','*.msi'), recursive=True)
if msi: shutil.copy(msi[0], os.path.join(d,'trayprint.msi'))
print('exe:', os.path.getsize(os.path.join(d,'trayprint.exe')) if os.path.exists(os.path.join(d,'trayprint.exe')) else 'TIDAK ADA')
PY
EXE="${TMPD}/trayprint.exe"
[ -f "$EXE" ] || { echo "ERROR: exe tidak ditemukan di artifact"; exit 1; }

# 3) Publish ke hub (dari IP lokal — tidak kena challenge CF)
VERSION=$(python3 -c "import json;print(json.load(open('$(pwd)/config.json')).get('version','3.0.0'))")
echo "Publish EXE v${VERSION} ke ${HUB_URL}"
CODE=$(curl -sS -o "${TMPD}/resp.json" -w '%{http_code}' -X POST "${HUB_URL}/api/v1/releases" \
  -H "X-Release-Token: ${TOKEN}" \
  -F "version=${VERSION}" \
  -F "platform=windows" \
  -F "channel=stable" \
  -F "installer_file=@${EXE}")
echo "HTTP ${CODE}"
head -c 500 "${TMPD}/resp.json"; echo
rm -rf "$TMPD"
[ "$CODE" = "201" ] && echo "PUBLISH OK" || { echo "PUBLISH GAGAL"; exit 1; }
