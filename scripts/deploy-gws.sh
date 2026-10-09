#!/usr/bin/env bash
# gws CLI로 시트 만들기 → 시트에 붙은 Apps Script 만들기 → 코드 올리기 → 웹앱 배포
#
#   ./scripts/deploy-gws.sh check                     준비 상태 점검 (먼저 실행 권장)
#   ./scripts/deploy-gws.sh init "심리 보고서 1차"     처음 한 번 (중간에 실패하면 같은 명령을 다시 실행하면 이어서 함)
#   ./scripts/deploy-gws.sh init --sheet 시트ID        이미 만든 시트에 붙이기
#   ./scripts/deploy-gws.sh push                      코드만 올리기
#   ./scripts/deploy-gws.sh deploy "설명"              코드 올리고 같은 주소로 새 버전 배포
#   ./scripts/deploy-gws.sh status                    저장된 ID와 주소 보기
#
# 필요: gws 0.22 이상 (npm i -g @googleworkspace/cli), python3
set -uo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
SRC="$ROOT/src"
STATE="$ROOT/.gws-deploy.json"
SCOPES_HINT="gws auth login -s script,sheets,drive"

red()   { printf '\033[31m%s\033[0m\n' "$*" >&2; }
green() { printf '\033[32m%s\033[0m\n' "$*"; }
step()  { printf '\n▶ %s\n' "$*"; }

state_get() { [ -f "$STATE" ] && python3 -c "import json; print(json.load(open('$STATE')).get('$1',''))" 2>/dev/null || true; }
state_set() {
  python3 - "$STATE" "$1" "$2" <<'PY'
import json, os, sys
p, k, v = sys.argv[1:]
d = json.load(open(p)) if os.path.exists(p) else {}
d[k] = v
json.dump(d, open(p, 'w'), ensure_ascii=False, indent=2)
PY
}

# 오류 메시지를 보고 무엇을 하면 되는지 한국어로 알려 준다.
explain() {
  local msg="$1"
  red "────────── 실패 ──────────"
  printf '%s\n' "$msg" | head -c 1500 >&2; echo >&2
  red "────────── 이렇게 해 보세요 ──────────"
  case "$msg" in
    *"No credentials"*|*"authError"*|*"invalid_grant"*|*"Token has been expired"*|*'"code":401'*|*'"code": 401'*)
      red "로그인이 안 되어 있거나 만료되었습니다."
      red "  $SCOPES_HINT" ;;
    *"insufficient"*|*"INSUFFICIENT"*|*"ACCESS_TOKEN_SCOPE"*|*"Insufficient Permission"*)
      red "로그인할 때 Apps Script 권한(scope)을 받지 않았습니다. 다시 로그인하면서 script, sheets, drive를 고르세요."
      red "  gws auth logout && $SCOPES_HINT" ;;
    *"User has not enabled the Apps Script API"*)
      red "내 계정에서 Apps Script API가 꺼져 있습니다. 아래에서 켜고 1~2분 뒤 다시 실행하세요."
      red "  https://script.google.com/home/usersettings" ;;
    *"has not been used in project"*|*"SERVICE_DISABLED"*|*"is disabled"*|*"accessNotConfigured"*)
      red "gws가 쓰는 Google Cloud 프로젝트에서 API가 꺼져 있습니다. 메시지 안의 링크를 열어 '사용'을 누르거나,"
      red "  gcloud services enable script.googleapis.com sheets.googleapis.com drive.googleapis.com"
      red "그다음 1~2분 기다렸다가 다시 실행하세요." ;;
    *"access_denied"*|*"not verified"*|*"admin_policy_enforced"*|*"restricted"*)
      red "학교(Workspace) 관리자가 외부 앱 접근을 막았을 수 있습니다. 개인 계정으로 시험하거나 관리자에게 OAuth 앱 허용을 요청하세요." ;;
    *"Duplicate"*|*"same name"*|*"already exists"*)
      red "같은 이름의 파일이 두 개 있습니다. src 폴더 이름을 확인하세요 (.gs와 .html 이름이 겹치면 안 됩니다)." ;;
    *"Requested entity was not found"*)
      red "저장된 ID의 시트나 스크립트를 찾지 못했습니다. 지웠다면 .gws-deploy.json 파일을 지우고 init을 다시 하세요." ;;
    *)
      red "위 메시지를 그대로 복사해 보내 주세요." ;;
  esac
  exit 1
}

# gws를 실행하고, 실패하면 explain. 성공하면 표준 출력(JSON)을 돌려준다.
run_gws() {
  local out err code
  err=$(mktemp)
  out=$(gws "$@" 2>"$err"); code=$?
  if [ $code -ne 0 ] || printf '%s' "$out" | grep -q '"error"[[:space:]]*:'; then
    explain "$(cat "$err"; printf '%s' "$out")"
  fi
  rm -f "$err"
  printf '%s' "$out"
}

json_field() { python3 -c "import json,sys; d=json.load(sys.stdin); print(d$1)"; }

# gws_field "['키']" <gws 인자...> : 성공하면 JSON에서 그 값만 출력, 실패하면 설명을 보여 주고 1을 돌려준다.
gws_field() {
  local field="$1"; shift
  local out
  out=$(run_gws "$@") || return 1
  printf '%s' "$out" | json_field "$field"
}

cmd_check() {
  local ok=1
  step "준비 상태 점검"
  if ! command -v gws >/dev/null; then
    red "✗ gws가 없습니다:  npm install -g @googleworkspace/cli"; exit 1
  fi
  green "✓ $(gws --version 2>/dev/null | head -1)"
  command -v python3 >/dev/null && green "✓ python3" || { red "✗ python3가 필요합니다."; ok=0; }

  local st src
  st=$(gws auth status 2>/dev/null)
  src=$(printf '%s' "$st" | json_field ".get('credential_source','none')" 2>/dev/null || echo none)
  if [ "$src" = "none" ] || [ -z "$src" ]; then
    red "✗ 로그인 정보가 없습니다."
    if printf '%s' "$st" | grep -q '"client_config_exists": false'; then
      red "  먼저 OAuth 클라이언트를 만들어야 합니다 (처음 한 번):"
      red "    gcloud가 있으면:  gws auth setup --login"
      red "    없으면: https://github.com/googleworkspace/cli#manual-oauth-setup-google-cloud-console"
    fi
    red "  그다음:  $SCOPES_HINT"
    exit 1
  fi
  green "✓ 로그인됨 ($src)"

  # 실제로 API를 불러 권한과 API 사용 설정을 확인한다.
  run_gws drive files list --params '{"pageSize":1}' >/dev/null && green "✓ Drive API"
  local probe
  probe=$(gws script projects get --params '{"scriptId":"__check__"}' 2>&1)
  case "$probe" in
    *"not found"*|*"404"*|*"Requested entity"*|*"Invalid"*|*"400"*) green "✓ Apps Script API" ;;
    *) explain "$probe" ;;
  esac
  [ $ok -eq 1 ] && green "준비 완료. 이제 init을 실행하세요."
}

cmd_init() {
  local title="과제 제출·피드백" sheet=""
  while [ $# -gt 0 ]; do
    case "$1" in
      --sheet) sheet="$2"; shift 2 ;;
      *) title="$1"; shift ;;
    esac
  done
  [ -d "$SRC" ] || { red "src 폴더가 없습니다: $SRC"; exit 1; }

  local ssid sid
  ssid="${sheet:-$(state_get spreadsheetId)}"
  if [ -z "$ssid" ]; then
    step "스프레드시트 만들기: $title"
    ssid=$(gws_field "['spreadsheetId']" sheets spreadsheets create --json "$(python3 -c 'import json,sys; print(json.dumps({"properties":{"title":sys.argv[1]}}, ensure_ascii=False))' "$title")") || exit 1
    green "  시트 ID: $ssid"
  else
    step "시트 사용: $ssid"
  fi
  state_set spreadsheetId "$ssid"

  sid="$(state_get scriptId)"
  if [ -z "$sid" ]; then
    step "시트에 붙은 Apps Script 프로젝트 만들기"
    sid=$(gws_field "['scriptId']" script projects create --json "$(python3 -c 'import json,sys; print(json.dumps({"title":sys.argv[1]+" 웹앱","parentId":sys.argv[2]}, ensure_ascii=False))' "$title" "$ssid")") || exit 1
    state_set scriptId "$sid"
    green "  스크립트 ID: $sid"
  else
    step "스크립트 사용: $sid"
  fi

  cmd_push
  cmd_deploy "첫 배포"

  local dep; dep=$(state_get deploymentId)
  echo
  green "✅ 끝났습니다"
  echo "  시트:      https://docs.google.com/spreadsheets/d/$ssid/edit"
  echo "  스크립트:  https://script.google.com/d/$sid/edit"
  echo "  학생 주소: https://script.google.com/macros/s/$dep/exec"
  echo "  관리 화면: https://script.google.com/macros/s/$dep/exec?page=admin"
  echo
  echo "꼭 할 일 (한 번만):"
  echo "  1) 위 '시트'를 열고 새로 고침 → 메뉴 '📮 과제 피드백 > ① 처음 설정' → 권한 허용 → 관리자 비밀번호 정하기"
  echo "     (이걸 하기 전에는 웹앱 주소를 열면 오류가 납니다)"
  echo "  2) 학생 주소를 시크릿 창에서 열어 보세요."
}

cmd_push() {
  local sid; sid=$(state_get scriptId)
  [ -n "$sid" ] || { red "scriptId가 없습니다. 먼저 init을 실행하세요."; exit 1; }
  step "코드 올리기 (src → $sid)"
  (cd "$ROOT" && run_gws script +push --script "$sid" --dir src >/dev/null) || exit 1
  green "  올렸습니다."
}

cmd_deploy() {
  local desc="${1:-업데이트 $(date '+%Y-%m-%d %H:%M')}"
  local sid; sid=$(state_get scriptId)
  step "새 버전 만들기"
  local ver
  ver=$(gws_field "['versionNumber']" script projects versions create --params "{\"scriptId\":\"$sid\"}" \
        --json "$(python3 -c 'import json,sys; print(json.dumps({"description":sys.argv[1]}, ensure_ascii=False))' "$desc")") || exit 1
  green "  버전 $ver"
  local dep; dep=$(state_get deploymentId)
  if [ -z "$dep" ]; then
    step "웹앱 배포 만들기"
    dep=$(gws_field "['deploymentId']" script projects deployments create --params "{\"scriptId\":\"$sid\"}" \
          --json "{\"versionNumber\":$ver,\"manifestFileName\":\"appsscript\",\"description\":\"v$ver\"}") || exit 1
    state_set deploymentId "$dep"
  else
    step "같은 주소로 버전 $ver 배포 (학생 주소는 그대로)"
    run_gws script projects deployments update --params "{\"scriptId\":\"$sid\",\"deploymentId\":\"$dep\"}" \
        --json "{\"deploymentConfig\":{\"versionNumber\":$ver,\"manifestFileName\":\"appsscript\",\"description\":\"v$ver\"}}" >/dev/null || exit 1
  fi
  green "  웹앱 주소: https://script.google.com/macros/s/$dep/exec"
}

cmd_status() {
  [ -f "$STATE" ] || { echo "아직 init을 하지 않았습니다."; exit 0; }
  cat "$STATE"; echo
  local dep; dep=$(state_get deploymentId)
  [ -n "$dep" ] && echo "학생 주소: https://script.google.com/macros/s/$dep/exec"
}

case "${1:-}" in
  check)  cmd_check ;;
  init)   shift; cmd_init "$@" ;;
  push)   cmd_push ;;
  deploy) shift; cmd_push; cmd_deploy "$@" ;;
  status) cmd_status ;;
  *) sed -n '2,11p' "$0"; exit 1 ;;
esac
