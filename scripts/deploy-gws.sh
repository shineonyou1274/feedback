#!/usr/bin/env bash
# gws CLI로 시트 만들기 → 시트에 붙은 Apps Script 프로젝트 만들기 → 코드 올리기 → 웹앱 배포
#
# 처음:     ./scripts/deploy-gws.sh init "인간과 심리 1차 과제"
# 코드 수정: ./scripts/deploy-gws.sh push
# 새 버전 배포: ./scripts/deploy-gws.sh deploy "설명"
#
# 필요: gws (Google Workspace CLI, `gws auth login` 완료), python3
#       https://script.google.com/home/usersettings 에서 "Google Apps Script API" 켜기
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
SRC="$ROOT/src"
STATE="$ROOT/.gws-deploy.json"   # scriptId, spreadsheetId, deploymentId 보관 (git에 올리지 않음)

json_get() { python3 -c "import json,sys; print(json.load(sys.stdin)$1)"; }
state_get() { [ -f "$STATE" ] && python3 -c "import json; print(json.load(open('$STATE')).get('$1',''))" || echo ""; }
state_set() {
  python3 - "$STATE" "$1" "$2" <<'PY'
import json, os, sys
p, k, v = sys.argv[1:]
d = json.load(open(p)) if os.path.exists(p) else {}
d[k] = v
json.dump(d, open(p, 'w'), ensure_ascii=False, indent=2)
PY
}

# src/ 폴더 → Apps Script API의 Content 형식(JSON)
build_content() {
  python3 - "$SRC" <<'PY'
import json, os, sys
src = sys.argv[1]
files = []
for name in sorted(os.listdir(src)):
    path = os.path.join(src, name)
    base, ext = os.path.splitext(name)
    kind = {'.gs': 'SERVER_JS', '.html': 'HTML', '.json': 'JSON'}.get(ext)
    if not kind:
        continue
    files.append({'name': base, 'type': kind, 'source': open(path, encoding='utf-8').read()})
print(json.dumps({'files': files}, ensure_ascii=False))
PY
}

cmd_init() {
  local title="${1:-과제 제출·피드백}"
  echo "▶ 스프레드시트 만들기: $title"
  local sheet_json
  sheet_json=$(gws sheets spreadsheets create --json "{\"properties\":{\"title\":\"$title\"}}")
  local ssid
  ssid=$(echo "$sheet_json" | json_get "['spreadsheetId']")
  state_set spreadsheetId "$ssid"

  echo "▶ 시트에 붙은 Apps Script 프로젝트 만들기"
  local proj_json sid
  proj_json=$(gws script projects create --json "{\"title\":\"$title 웹앱\",\"parentId\":\"$ssid\"}")
  sid=$(echo "$proj_json" | json_get "['scriptId']")
  state_set scriptId "$sid"

  cmd_push
  cmd_deploy "첫 배포"

  echo
  echo "✅ 완료"
  echo "  시트:   https://docs.google.com/spreadsheets/d/$ssid/edit"
  echo "  스크립트: https://script.google.com/d/$sid/edit"
  echo
  echo "다음 할 일:"
  echo "  1) 시트를 열고 새로 고침 → '📮 과제 피드백 > ① 처음 설정'을 눌러 권한을 허용하세요."
  echo "  2) 위 웹앱 주소를 설정 시트의 '웹앱 주소' 칸에 붙여 넣으세요."
}

cmd_push() {
  local sid; sid=$(state_get scriptId)
  [ -n "$sid" ] || { echo "scriptId가 없습니다. 먼저 init을 실행하세요." >&2; exit 1; }
  echo "▶ 코드 올리기 ($sid)"
  gws script projects updateContent --params "{\"scriptId\":\"$sid\"}" --json "$(build_content)" >/dev/null
  echo "  올렸습니다."
}

cmd_deploy() {
  local desc="${1:-업데이트 $(date '+%Y-%m-%d %H:%M')}"
  local sid; sid=$(state_get scriptId)
  echo "▶ 새 버전 만들기"
  local ver
  ver=$(gws script projects versions create --params "{\"scriptId\":\"$sid\"}" \
        --json "{\"description\":\"$desc\"}" | json_get "['versionNumber']")
  local dep; dep=$(state_get deploymentId)
  local cfg="{\"deploymentConfig\":{\"versionNumber\":$ver,\"manifestFileName\":\"appsscript\",\"description\":\"$desc\"}}"
  if [ -z "$dep" ]; then
    echo "▶ 웹앱 배포 만들기 (버전 $ver)"
    dep=$(gws script projects deployments create --params "{\"scriptId\":\"$sid\"}" \
          --json "{\"versionNumber\":$ver,\"manifestFileName\":\"appsscript\",\"description\":\"$desc\"}" \
          | json_get "['deploymentId']")
    state_set deploymentId "$dep"
  else
    echo "▶ 같은 주소로 버전 $ver 배포 (학생 주소는 그대로)"
    gws script projects deployments update --params "{\"scriptId\":\"$sid\",\"deploymentId\":\"$dep\"}" \
        --json "$cfg" >/dev/null
  fi
  echo "  웹앱 주소: https://script.google.com/macros/s/$dep/exec"
}

case "${1:-}" in
  init)   shift; cmd_init "$@" ;;
  push)   cmd_push ;;
  deploy) shift; cmd_push; cmd_deploy "$@" ;;
  *) sed -n '2,10p' "$0"; exit 1 ;;
esac
