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
  '상태', '공개일', '공개', 'PDF', 'PDF ID', 'PDF 생성시각', '도착메일', 'AI 초안시각',
];
/** 고치면 PDF를 다시 만들어야 하는 열 */
const FB_CONTENT_COLS = ['잘한 점', '보완할 점', '다음 단계까지 할 일', '한 줄 총평'];

const STATUS = {
  NEW: '미작성',
  DRAFT: 'AI 초안',
  DONE: '확정',
  PDF: 'PDF 완료',
};

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
