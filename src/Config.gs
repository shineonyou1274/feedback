/**
 * 과정 중심 과제 제출·피드백 웹앱 — 공통 설정과 도우미 함수
 * © shiny-peace
 */

const SHEET = {
  SETTINGS: '설정',
  STAGES: '단계',
  RUBRIC: '루브릭',
  ROSTER: '명단',
  SUBMIT: '제출',
  FEEDBACK: '피드백',
};

/** '설정' 시트의 키 이름 */
const KEY = {
  COURSE: '과목',
  TITLE: '과제명',
  STANDARD: '성취기준',
  GUIDE: '안내문',
  TEACHER: '교사 이름',
  WEBAPP_URL: '웹앱 주소',
  ROSTER_ONLY: '명단에 있는 학생만 제출',
  MAX_MB: '파일 최대 크기(MB)',
  EXTS: '허용 확장자',
  SUBMIT_MAIL: '제출 완료 메일 보내기',
  SUBMIT_MAIL_SUBJECT: '제출 완료 메일 제목',
  ARRIVE_MAIL: '피드백 도착 메일 보내기',
  ARRIVE_MAIL_SUBJECT: '피드백 도착 메일 제목',
  AI_MODEL: 'AI 모델',
  AI_EFFORT: 'AI 생각 깊이',
  SUBMIT_FOLDER: '제출 폴더 ID',
  PDF_FOLDER: 'PDF 폴더 ID',
};

const SUBMIT_HEADERS = [
  '접수번호', '제출시각', '학번', '이름', '이메일', '단계ID', '단계명',
  '이번 단계에서 한 일', '어려웠던 점·질문', '본문', '파일명', '파일링크', '파일ID', '제출차수', '완료메일',
];

/** 피드백 시트: 앞쪽 고정 열 + (루브릭 항목마다 '예상:항목명' 열) + 뒤쪽 고정 열 */
const FB_HEAD = ['접수번호', '학번', '이름', '단계ID', '단계명', '제출시각', '제출파일', '학생 메모'];
const FB_SCORE_PREFIX = '예상:';
const FB_TAIL = [
  '예상 합계', '잘한 점', '보완할 점', '다음 단계까지 할 일', '한 줄 총평',
  '공개', '공개시각', '도착메일', 'PDF', 'PDF ID', 'PDF 해시', 'AI 초안시각',
];
/** 예전 버전에만 있던 열 (정리할 때 지운다) */
const FB_OLD_COLS = ['상태', '공개일', 'PDF 생성시각'];
/** 선생님 의견 칸 */
const FB_CONTENT_COLS = ['잘한 점', '보완할 점', '다음 단계까지 할 일', '한 줄 총평'];

/** 의견이 하나라도 적혀 있으면 '작성됨' */
function hasContent_(fb) {
  return FB_CONTENT_COLS.some(k => String(fb[k] == null ? '' : fb[k]).trim() !== '');
}

/** 화면에 보일 상태: 미작성 / 작성됨 / 공개 */
function fbState_(fb) {
  if (fb['공개'] === true && hasContent_(fb)) return '공개';
  return hasContent_(fb) ? '작성됨' : '미작성';
}

const TIME_BUDGET_MS = 4.5 * 60 * 1000; // 6분 실행 제한 전에 멈춘다.

// ───────────────────────── 시트 도우미 ─────────────────────────

function ss_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('스프레드시트를 찾을 수 없습니다. 시트 메뉴에서 "처음 설정"을 먼저 실행하세요.');
  return SpreadsheetApp.openById(id);
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error(`'${name}' 시트가 없습니다. "처음 설정"을 실행하세요.`);
  return sh;
}

/** 첫 행 머리글 → {머리글: 1부터 시작하는 열 번호} */
function headerMap_(sh) {
  const lastCol = sh.getLastColumn();
  const map = {};
  if (lastCol === 0) return map;
  sh.getRange(1, 1, 1, lastCol).getValues()[0].forEach((h, i) => {
    if (h !== '') map[String(h).trim()] = i + 1;
  });
  return map;
}

/** 시트 전체를 [{머리글: 값, _row: 행번호}] 로 읽는다. */
function readRows_(sh) {
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const head = values[0].map(h => String(h).trim());
  return values.slice(1).map((r, i) => {
    const o = { _row: i + 2 };
    head.forEach((h, c) => { if (h) o[h] = r[c]; });
    return o;
  });
}

/** 한 행의 여러 열을 머리글 이름으로 쓴다. */
function writeCells_(sh, row, map, data) {
  Object.keys(data).forEach(k => {
    if (!map[k]) return;
    sh.getRange(row, map[k]).setValue(data[k]);
  });
}

/** col 열에 값이 있는 마지막 행 번호 (머리글만 있으면 1) */
function lastDataRow_(sh, col) {
  const n = sh.getLastRow();
  if (n < 2 || !col) return Math.max(n, 1);
  const vals = sh.getRange(2, col, n - 1, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) if (vals[i][0] !== '' && vals[i][0] !== null) return i + 2;
  return 1;
}

/**
 * 피드백 시트 정리
 * 1) 접수번호가 없는 줄(예전 버전이 미리 깔아 둔 빈 체크박스 줄)을 지워 실제 행을 위로 올린다.
 * 2) 제출 시트에는 있는데 피드백 시트에 없는 제출을 자동으로 채운다.
 * 결과: {removed, added}
 */
function repairFeedbackSheet_() {
  const sh = sheet_(SHEET.FEEDBACK);
  // 예전 열(상태·공개일 등)을 지우고 새 열을 붙인다.
  FB_OLD_COLS.forEach(h => { const c = headerMap_(sh)[h]; if (c) sh.deleteColumn(c); });
  const have0 = headerMap_(sh);
  FB_TAIL.forEach(h => { if (!have0[h]) sh.getRange(1, sh.getLastColumn() + 1).setValue(h); });
  const map = headerMap_(sh);
  const col = map['접수번호'];
  let removed = 0, added = 0;
  if (col && sh.getLastRow() >= 2) {
    const ids = sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0] == null ? '' : r[0]).trim());
    const hasData = ids.some(Boolean);
    if (!hasData) {
      sh.getRange(2, 1, sh.getMaxRows() - 1, sh.getLastColumn()).clearContent().clearDataValidations();
      removed = ids.length;
    } else {
      // 아래에서부터 연속된 빈 줄 묶음을 지운다.
      let i = ids.length - 1;
      while (i >= 0) {
        if (ids[i]) { i--; continue; }
        let j = i;
        while (j - 1 >= 0 && !ids[j - 1]) j--;
        sh.deleteRows(j + 2, i - j + 1);
        removed += i - j + 1;
        i = j - 1;
      }
    }
  }
  // 빠진 피드백 행 채우기
  const have = new Set(readRows_(sh).map(r => String(r['접수번호'])).filter(Boolean));
  const stages = getStages_();
  readRows_(sheet_(SHEET.SUBMIT)).forEach(r => {
    const rc = String(r['접수번호'] || '');
    if (!rc || have.has(rc)) return;
    const stage = stages.find(st => st.id === String(r['단계ID'])) || { id: String(r['단계ID']), name: String(r['단계명']) };
    addFeedbackRow_({
      receipt: rc, id: normId_(r['학번']), name: normName_(r['이름']), stage, now: r['제출시각'],
      fileName: r['파일명'], fileUrl: r['파일링크'], did: r['이번 단계에서 한 일'], question: r['어려웠던 점·질문'],
      attempt: Number(r['제출차수']) || 1,
    });
    added++;
  });
  const rosterAdded = syncRoster_();
  return { removed, added, rosterAdded };
}

/** 제출한 학생을 명단에 채우고, 처음 설정 때 넣은 예시(1101 홍길동)는 제출이 없으면 지운다. */
function syncRoster_() {
  let sh = ss_().getSheetByName(SHEET.ROSTER);
  if (!sh) {
    sh = ss_().insertSheet(SHEET.ROSTER);
    sh.getRange(1, 1, 1, 3).setValues([['학번', '이름', '이메일']]);
  }
  const subs = readRows_(sheet_(SHEET.SUBMIT)).filter(r => r['학번'] !== '');
  const submitted = new Set(subs.map(r => normId_(r['학번'])));
  // 예시 행 지우기
  const rows = readRows_(sh);
  for (let i = rows.length - 1; i >= 0; i--) {
    const r = rows[i];
    if (normId_(r['학번']) === '1101' && normName_(r['이름']) === '홍길동' && !String(r['이메일'] || '') && !submitted.has('1101')) {
      sh.deleteRow(r._row);
    }
  }
  const have = new Set(readRows_(sh).map(r => normId_(r['학번'])).filter(Boolean));
  let n = 0;
  subs.forEach(r => {
    const id = normId_(r['학번']);
    if (!id || have.has(id)) return;
    have.add(id);
    const row = lastDataRow_(sh, 1) + 1;
    sh.getRange(row, 1).setNumberFormat('@');
    sh.getRange(row, 1, 1, 3).setValues([[id, normName_(r['이름']), String(r['이메일'] || '')]]);
    n++;
  });
  return n;
}

// ───────────────────────── 설정 읽기 ─────────────────────────

function getSettings_() {
  const rows = sheet_(SHEET.SETTINGS).getDataRange().getValues().slice(1);
  const s = {};
  rows.forEach(r => { if (r[0] !== '') s[String(r[0]).trim()] = r[1]; });
  return s;
}

function setSetting_(key, value) {
  const sh = sheet_(SHEET.SETTINGS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === key) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value, '']);
}

function yes_(v) {
  return v === true || /^(y|yes|예|네|사용|true|o)$/i.test(String(v).trim());
}

/** 단계 목록: [{id, name, desc, focus, mode, deadline(Date|null), open, allowLate}] */
function getStages_() {
  return readRows_(sheet_(SHEET.STAGES))
    .filter(r => r['단계ID'] !== '')
    .map(r => ({
      id: String(r['단계ID']).trim(),
      name: String(r['단계명'] || '').trim(),
      desc: String(r['학생 안내'] || ''),
      focus: String(r['피드백 초점'] || ''),
      mode: String(r['제출 방식'] || '파일+글').trim(),
      deadline: r['마감일시'] instanceof Date ? r['마감일시'] : null,
      open: yes_(r['열림']),
      allowLate: yes_(r['마감 후 제출 허용']),
    }));
}

/** 루브릭: [{item, max, high, mid, low}] — 최종 결과물 기준표. 단계마다 '현재 상태로 최종 제출하면 몇 점일지' 예상한다. */
function getRubric_() {
  return readRows_(sheet_(SHEET.RUBRIC))
    .filter(r => r['평가 항목'] !== '')
    .map(r => ({
      item: String(r['평가 항목']).trim(),
      max: Number(r['배점']) || 0,
      high: String(r['상'] || ''),
      mid: String(r['중'] || ''),
      low: String(r['하'] || ''),
    }));
}

function getRoster_() {
  const sh = ss_().getSheetByName(SHEET.ROSTER);
  if (!sh) return [];
  return readRows_(sh)
    .filter(r => r['학번'] !== '')
    .map(r => ({ id: normId_(r['학번']), name: normName_(r['이름']), email: String(r['이메일'] || '').trim() }));
}

// ───────────────────────── 값 정리 ─────────────────────────

function normId_(v) { return String(v || '').replace(/\D/g, ''); }
function normName_(v) { return String(v || '').replace(/\s+/g, ''); }

function fmtTime_(d) {
  if (!(d instanceof Date)) return d ? String(d) : '';
  return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
}

function isEmail_(v) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(v || '').trim()); }

function escHtml_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function webAppUrl_() {
  const s = getSettings_();
  if (s[KEY.WEBAPP_URL]) return String(s[KEY.WEBAPP_URL]);
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

/** 점수 비율 → 상/중/하 */
function levelOf_(score, max) {
  if (score === '' || score == null || !max) return '';
  const r = Number(score) / max;
  if (r >= 0.8) return '상';
  if (r >= 0.5) return '중';
  return '하';
}

/**
 * 시트 메뉴에서만 실행되게 막는다.
 * 이름 끝에 _가 없는 함수는 웹앱 화면에서 google.script.run으로 부를 수 있으므로,
 * 메뉴 함수는 맨 처음에 이것을 불러 웹앱(학생 화면)에서 실행되지 않게 한다.
 */
function uiOnly_() {
  try {
    SpreadsheetApp.getUi();
  } catch (e) {
    throw new Error('시트 메뉴에서만 실행할 수 있습니다.');
  }
}
