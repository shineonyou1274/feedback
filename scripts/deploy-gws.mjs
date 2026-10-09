#!/usr/bin/env node
// gws CLI로 시트 만들기 → 시트에 붙은 Apps Script 만들기 → 코드 올리기 → 웹앱 배포
// 윈도우(cmd, PowerShell)·맥·리눅스 어디서나 같은 명령으로 실행한다.
//
//   node scripts/deploy-gws.mjs check                    준비 상태 점검
//   node scripts/deploy-gws.mjs init "심리 보고서 1차"    처음 한 번 (실패하면 고친 뒤 같은 명령으로 이어서)
//   node scripts/deploy-gws.mjs init --sheet 시트ID       이미 만든 시트에 붙이기
//   node scripts/deploy-gws.mjs push                     코드만 올리기
//   node scripts/deploy-gws.mjs deploy "설명"             코드 올리고 같은 주소로 새 버전 배포
//   node scripts/deploy-gws.mjs status                   저장된 ID와 주소 보기

import { spawnSync } from 'node:child_process';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const SRC = join(ROOT, 'src');
const STATE = join(ROOT, '.gws-deploy.json');
const LOGIN = 'gws auth login -s script,sheets,drive';

const red = s => console.error(`\x1b[31m${s}\x1b[0m`);
const green = s => console.log(`\x1b[32m${s}\x1b[0m`);
const step = s => console.log(`\n▶ ${s}`);

// ───────── 상태 파일 ─────────
const loadState = () => (existsSync(STATE) ? JSON.parse(readFileSync(STATE, 'utf8')) : {});
const saveState = (k, v) => { const d = loadState(); d[k] = v; writeFileSync(STATE, JSON.stringify(d, null, 2)); };

// ───────── gws 찾기 ─────────
// 윈도우의 gws.cmd를 셸로 부르면 JSON 따옴표가 깨지므로, gws 패키지의 run.js를 node로 직접 실행한다.
function findGws() {
  if (process.env.GWS_BIN) return { cmd: process.execPath, pre: [process.env.GWS_BIN] };
  const r = spawnSync('npm', ['root', '-g'], { encoding: 'utf8', shell: process.platform === 'win32' });
  const runJs = r.stdout ? join(r.stdout.trim(), '@googleworkspace', 'cli', 'run.js') : '';
  if (runJs && existsSync(runJs)) return { cmd: process.execPath, pre: [runJs] };
  return null;
}
const GWS = findGws();

function gwsRaw(args) {
  const r = spawnSync(GWS.cmd, [...GWS.pre, ...args], { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
  return { code: r.status, out: r.stdout || '', err: r.stderr || '' };
}

/** gws를 실행하고 JSON을 돌려준다. 실패하면 원인을 설명하고 끝낸다. */
function gws(args) {
  const { code, out, err } = gwsRaw(args);
  let json = null;
  try { json = out.trim() ? JSON.parse(out) : {}; } catch { /* JSON이 아님 */ }
  if (code !== 0 || !json || json.error) explain(`${err}\n${out}`);
  return json;
}

function explain(msg) {
  red('────────── 실패 ──────────');
  console.error(msg.trim().slice(0, 1500));
  red('────────── 이렇게 해 보세요 ──────────');
  const has = (...keys) => keys.some(k => msg.includes(k));
  if (has('No credentials', 'authError', 'invalid_grant', 'Token has been expired', '"code": 401', '"code":401')) {
    red('로그인이 안 되어 있거나 만료되었습니다.');
    red(`  ${LOGIN}`);
  } else if (has('insufficient', 'INSUFFICIENT', 'ACCESS_TOKEN_SCOPE', 'Insufficient Permission')) {
    red('로그인할 때 Apps Script 권한을 받지 않았습니다. 다시 로그인하면서 script, sheets, drive를 고르세요.');
    red(`  gws auth logout`);
    red(`  ${LOGIN}`);
  } else if (has('User has not enabled the Apps Script API')) {
    red('내 계정에서 Apps Script API가 꺼져 있습니다. 아래 주소에서 켜고 1~2분 뒤 다시 실행하세요.');
    red('  https://script.google.com/home/usersettings');
  } else if (has('has not been used in project', 'SERVICE_DISABLED', 'is disabled', 'accessNotConfigured')) {
    red('gws가 쓰는 Google Cloud 프로젝트에서 API가 꺼져 있습니다. 메시지 안의 https://console.developers.google.com/... 링크를 열어 "사용"을 누르세요.');
    red('Apps Script API, Google Sheets API, Google Drive API 세 개가 모두 켜져 있어야 합니다. 그다음 1~2분 뒤 다시 실행하세요.');
  } else if (has('access_denied', 'admin_policy_enforced', 'restricted')) {
    red('학교(Workspace) 관리자가 외부 앱 접근을 막았을 수 있습니다. 개인 구글 계정으로 시험하거나 관리자에게 허용을 요청하세요.');
  } else if (has('Duplicate', 'same name')) {
    red('같은 이름의 파일이 있습니다. src 폴더에서 .gs와 .html 이름이 겹치지 않는지 확인하세요.');
  } else if (has('Requested entity was not found')) {
    red('저장된 ID의 시트나 스크립트를 찾지 못했습니다. 지웠다면 .gws-deploy.json 파일을 지우고 init을 다시 하세요.');
  } else {
    red('위 메시지를 그대로 복사해 보내 주세요.');
  }
  process.exit(1);
}

// ───────── 명령 ─────────
function check() {
  step('준비 상태 점검');
  green(`✓ Node ${process.version}`);
  if (!GWS) {
    red('✗ gws를 찾지 못했습니다. 아래로 설치한 뒤 다시 실행하세요.');
    red('  npm install -g @googleworkspace/cli');
    process.exit(1);
  }
  green(`✓ ${gwsRaw(['--version']).out.split('\n')[0]}`);

  let st = {};
  try { st = JSON.parse(gwsRaw(['auth', 'status']).out); } catch { /* 무시 */ }
  if (!st.credential_source || st.credential_source === 'none') {
    red('✗ gws 로그인 정보가 없습니다.');
    if (st.client_config_exists === false) {
      red('  OAuth 클라이언트가 아직 없습니다 (처음 한 번):');
      red('    gcloud가 있으면:  gws auth setup --login');
      red('    없으면: https://github.com/googleworkspace/cli#manual-oauth-setup-google-cloud-console');
    }
    red(`  그다음:  ${LOGIN}`);
    process.exit(1);
  }
  green(`✓ 로그인됨 (${st.credential_source})`);

  gws(['drive', 'files', 'list', '--params', JSON.stringify({ pageSize: 1 })]);
  green('✓ Drive API');
  const probe = gwsRaw(['script', 'projects', 'get', '--params', JSON.stringify({ scriptId: '__check__' })]);
  const p = probe.err + probe.out;
  if (/not found|Requested entity|Invalid|"code":\s*400|"code":\s*404/i.test(p)) green('✓ Apps Script API');
  else explain(p);
  green('\n준비 완료. 이제 init을 실행하세요.');
}

function init(args) {
  let title = '과제 제출·피드백';
  let sheet = '';
  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--sheet') sheet = args[++i] || '';
    else title = args[i];
  }
  if (!existsSync(SRC)) { red(`src 폴더가 없습니다: ${SRC}`); process.exit(1); }
  if (!GWS) { red('gws를 찾지 못했습니다. 먼저 check를 실행하세요.'); process.exit(1); }

  let st = loadState();
  let ssid = sheet || st.spreadsheetId;
  if (!ssid) {
    step(`스프레드시트 만들기: ${title}`);
    ssid = gws(['sheets', 'spreadsheets', 'create', '--json', JSON.stringify({ properties: { title } })]).spreadsheetId;
    green(`  시트 ID: ${ssid}`);
  } else {
    step(`시트 사용: ${ssid}`);
  }
  saveState('spreadsheetId', ssid);

  let sid = loadState().scriptId;
  if (!sid) {
    step('시트에 붙은 Apps Script 프로젝트 만들기');
    sid = gws(['script', 'projects', 'create', '--json', JSON.stringify({ title: `${title} 웹앱`, parentId: ssid })]).scriptId;
    saveState('scriptId', sid);
    green(`  스크립트 ID: ${sid}`);
  } else {
    step(`스크립트 사용: ${sid}`);
  }

  push();
  deploy('첫 배포');

  const dep = loadState().deploymentId;
  console.log('\n');
  green('✅ 끝났습니다');
  console.log(`  시트:      https://docs.google.com/spreadsheets/d/${ssid}/edit`);
  console.log(`  스크립트:  https://script.google.com/d/${sid}/edit`);
  console.log(`  학생 주소: https://script.google.com/macros/s/${dep}/exec`);
  console.log(`  관리 화면: https://script.google.com/macros/s/${dep}/exec?page=admin`);
  console.log('\n꼭 할 일 (한 번만):');
  console.log("  1) 위 '시트'를 열고 새로 고침 → 메뉴 '📮 과제 피드백 > ① 처음 설정' → 권한 허용 → 관리자 비밀번호 정하기");
  console.log('     (이걸 하기 전에는 웹앱 주소를 열면 오류가 납니다)');
  console.log('  2) 학생 주소를 시크릿 창에서 열어 보세요.');
}

function push() {
  const sid = loadState().scriptId;
  if (!sid) { red('scriptId가 없습니다. 먼저 init을 실행하세요.'); process.exit(1); }
  step(`코드 올리기 (src → ${sid})`);
  gws(['script', '+push', '--script', sid, '--dir', SRC]);
  green('  올렸습니다.');
}

function deploy(desc) {
  desc = desc || `업데이트 ${new Date().toLocaleString('ko-KR')}`;
  const sid = loadState().scriptId;
  step('새 버전 만들기');
  const ver = gws(['script', 'projects', 'versions', 'create', '--params', JSON.stringify({ scriptId: sid }),
    '--json', JSON.stringify({ description: desc })]).versionNumber;
  green(`  버전 ${ver}`);
  const cfg = { versionNumber: ver, manifestFileName: 'appsscript', description: `v${ver}` };
  let dep = loadState().deploymentId;
  if (!dep) {
    step('웹앱 배포 만들기');
    dep = gws(['script', 'projects', 'deployments', 'create', '--params', JSON.stringify({ scriptId: sid }),
      '--json', JSON.stringify(cfg)]).deploymentId;
    saveState('deploymentId', dep);
  } else {
    step(`같은 주소로 버전 ${ver} 배포 (학생 주소는 그대로)`);
    gws(['script', 'projects', 'deployments', 'update', '--params', JSON.stringify({ scriptId: sid, deploymentId: dep }),
      '--json', JSON.stringify({ deploymentConfig: cfg })]);
  }
  green(`  웹앱 주소: https://script.google.com/macros/s/${dep}/exec`);
}

function status() {
  if (!existsSync(STATE)) { console.log('아직 init을 하지 않았습니다.'); return; }
  const st = loadState();
  console.log(JSON.stringify(st, null, 2));
  if (st.deploymentId) console.log(`학생 주소: https://script.google.com/macros/s/${st.deploymentId}/exec`);
}

const [cmd, ...rest] = process.argv.slice(2);
switch (cmd) {
  case 'check': check(); break;
  case 'init': init(rest); break;
  case 'push': push(); break;
  case 'deploy': push(); deploy(rest[0]); break;
  case 'status': status(); break;
  default:
    console.log(readFileSync(fileURLToPath(import.meta.url), 'utf8').split('\n').slice(1, 11).map(l => l.replace(/^\/\/ ?/, '')).join('\n'));
    process.exit(1);
}
