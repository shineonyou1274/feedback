/**
 * 시트 메뉴, 처음 설정, 트리거
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📮 과제 피드백')
    .addItem('① 처음 설정 (시트·폴더 만들기)', 'setupAll')
    .addItem('루브릭 바꾼 뒤: 피드백 시트 점수 열 맞추기', 'syncFeedbackColumns')
    .addSeparator()
    .addItem('② 선택한 행: AI 예상 점수·피드백 초안', 'menuDraftSelected')
    .addItem('② 미작성 전체: AI 초안 (이어서 실행)', 'menuDraftAll')
    .addSeparator()
    .addItem('③ 선택한 행: 확정 표시', 'menuConfirmSelected')
    .addItem('④ 선택한 행: PDF 만들기', 'menuPdfSelected')
    .addItem('④ 확정 전체: PDF 만들기 (이어서 실행)', 'menuPdfAll')
    .addSeparator()
    .addItem('⑤ 선택한 행: 공개', 'menuPublishSelected')
    .addItem('⑤ 선택한 행: 공개 취소', 'menuUnpublishSelected')
    .addItem('피드백 도착 메일 지금 보내기', 'menuSendArrivalMails')
    .addSeparator()
    .addItem('⚙️ 관리자 비밀번호 정하기 (웹 관리 화면)', 'menuSetAdminPassword')
    .addItem('⚙️ AI API 키 등록', 'menuSetApiKey')
    .addItem('⚙️ 매일 아침 도착 메일 자동 발송 켜기', 'installDailyTrigger')
    .addToUi();
}

// ───────────────────────── 처음 설정 ─────────────────────────

function setupAll() {
  uiOnly_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  ensureSheet_(SHEET.SETTINGS, ['키', '값', '설명'], defaultSettings_());
  ensureSheet_(SHEET.STAGES, ['단계ID', '단계명', '학생 안내', '피드백 초점', '제출 방식', '마감일시', '열림', '마감 후 제출 허용'], defaultStages_());
  ensureSheet_(SHEET.RUBRIC, ['평가 항목', '배점', '상', '중', '하'], defaultRubric_());
  ensureSheet_(SHEET.ROSTER, ['학번', '이름', '이메일'], [['1101', '홍길동', '']]);
  ensureSheet_(SHEET.SUBMIT, SUBMIT_HEADERS, []);
  ensureSheet_(SHEET.FEEDBACK, FB_HEAD.concat(FB_TAIL), []);
  syncFeedbackColumns(true);
  ensureFolders_();
  styleSheets_();

  if (!hasAdminPassword_()) menuSetAdminPassword();

  SpreadsheetApp.getUi().alert(
    '처음 설정을 마쳤습니다.\n\n' +
    '1) 확장 프로그램 > Apps Script > 배포 > 새 배포 > 웹 앱으로 배포하세요.\n' +
    '   (실행: 나, 액세스: 모든 사용자)\n' +
    '2) 학생에게는 웹앱 주소를 그대로 알려 주세요. 제출과 피드백 확인을 모두 이 주소에서 합니다.\n' +
    '3) 이제부터는 웹앱 주소 끝에 ?page=admin 을 붙인 관리 화면에서 모두 할 수 있습니다.\n' +
    '   (과제 설정, 명단, 제출 현황, AI 초안, 피드백 작성, PDF, 공개, 메일)');
}

function ensureSheet_(name, headers, rows) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (rows.length) sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    sh.setFrozenRows(1);
    return sh;
  }
  // 이미 있으면 빠진 머리글만 뒤에 붙인다 (기존 자료 보존).
  const map = headerMap_(sh);
  headers.forEach(h => {
    if (!map[h]) {
      const col = sh.getLastColumn() + 1;
      sh.getRange(1, col).setValue(h);
      map[h] = col;
    }
  });
  return sh;
}

function defaultSettings_() {
  return [
    [KEY.COURSE, '인간과 심리', '학생 화면 맨 위에 보입니다.'],
    [KEY.TITLE, '심리 탐구 보고서', ''],
    [KEY.STANDARD, '[12심리01-02]', '피드백 PDF 부제로 들어갑니다. 비워도 됩니다.'],
    [KEY.GUIDE, '보고서는 한 번에 끝내는 것이 아니라 단계마다 제출하고 피드백을 받으며 완성합니다. 각 단계에서 무엇을 했는지, 어디서 막혔는지 솔직하게 적어 주세요.', '제출 화면 안내문'],
    [KEY.TEACHER, '', 'PDF와 메일에 들어갑니다.'],
    [KEY.WEBAPP_URL, '', '웹앱 배포 주소 (메일 속 링크에 씁니다)'],
    [KEY.ROSTER_ONLY, '아니오', '예: 명단 시트에 없는 학번·이름은 제출할 수 없습니다.'],
    [KEY.MAX_MB, 10, '10 이하를 권장합니다.'],
    [KEY.EXTS, 'pdf,docx,hwp,hwpx,pptx,jpg,png', '쉼표로 구분'],
    [KEY.SUBMIT_MAIL, '예', '제출하면 학생에게 접수 확인 메일을 바로 보냅니다.'],
    [KEY.SUBMIT_MAIL_SUBJECT, '[{과제명}] {단계명} 제출이 완료되었습니다', '{이름} {단계명} {과제명} {접수번호} 사용 가능'],
    [KEY.ARRIVE_MAIL, '예', '공개일이 된 피드백을 학생에게 알립니다.'],
    [KEY.ARRIVE_MAIL_SUBJECT, '[{과제명}] {단계명} 피드백이 도착했습니다', ''],
    [KEY.AI_MODEL, 'claude-opus-5-5', 'AI 초안에 쓰는 모델'],
    [KEY.AI_EFFORT, 'medium', 'low / medium / high — 높을수록 꼼꼼하지만 느립니다.'],
    [KEY.SUBMIT_FOLDER, '', '처음 설정 때 자동으로 채워집니다.'],
    [KEY.PDF_FOLDER, '', '처음 설정 때 자동으로 채워집니다.'],
  ];
}

function defaultStages_() {
  const d = n => { const t = new Date(); t.setDate(t.getDate() + n); t.setHours(23, 59, 0, 0); return t; };
  return [
    ['S1', '주제 정하기·탐구 계획', '탐구할 심리 현상, 그 주제를 고른 까닭, 탐구 질문, 찾아볼 이론이나 실험을 적어 계획서로 내세요.',
      '탐구 질문이 구체적이고 검증할 수 있는가, 주제와 이론이 이어지는가', '파일+글', d(7), '예', '예'],
    ['S2', '자료 조사·개요', '찾은 자료(출처 포함)와 보고서 개요(목차)를 내세요.',
      '자료의 신뢰도와 출처 표기, 개요의 논리 흐름, 계획서 피드백 반영 여부', '파일+글', d(14), '아니오', '예'],
    ['S3', '초안', '보고서 초안 전체를 내세요. 미완성이어도 괜찮습니다.',
      '이론을 근거로 현상을 설명하는가, 실험 설계와 결과의 관계, 이전 피드백 반영 여부', '파일+글', d(21), '아니오', '예'],
    ['S4', '최종 보고서', '고쳐 쓴 최종 보고서를 내세요. 무엇을 고쳤는지 "이번 단계에서 한 일"에 적어 주세요.',
      '루브릭 전체 기준, 초안 대비 발전한 점', '파일+글', d(28), '아니오', '아니오'],
  ];
}

function defaultRubric_() {
  return [
    ['탐구 질문과 주제 선정', 20, '현상을 구체적인 질문으로 좁히고 탐구할 까닭을 근거로 밝힌다.', '질문이 있으나 범위가 넓거나 까닭이 막연하다.', '질문이 분명하지 않다.'],
    ['이론·실험의 논리적 이해', 40, '이론과 실험 설계·결과의 관계를 근거를 들어 설명한다.', '이론을 소개하지만 현상과의 연결이 약하다.', '이론 설명이 부정확하거나 빠져 있다.'],
    ['자료 활용과 출처', 20, '믿을 만한 자료를 고르고 출처를 정확히 밝힌다.', '자료는 있으나 출처 표기가 일부 빠진다.', '자료가 부족하거나 출처가 없다.'],
    ['구성과 표현', 20, '서론·본론·결론이 짜임새 있고 문장이 정확하다.', '구성은 갖췄으나 흐름이 끊기는 곳이 있다.', '구성이 흐트러져 내용을 알기 어렵다.'],
  ];
}

function ensureFolders_() {
  const s = getSettings_();
  const ss = ss_();
  const parent = DriveApp.getFileById(ss.getId()).getParents().hasNext()
    ? DriveApp.getFileById(ss.getId()).getParents().next()
    : DriveApp.getRootFolder();
  const title = String(s[KEY.TITLE] || '과제');
  const root = getOrCreateFolder_(parent, `${title} 과제함`);
  if (!s[KEY.SUBMIT_FOLDER]) setSetting_(KEY.SUBMIT_FOLDER, getOrCreateFolder_(root, '제출 파일').getId());
  if (!s[KEY.PDF_FOLDER]) setSetting_(KEY.PDF_FOLDER, getOrCreateFolder_(root, '피드백 PDF').getId());
}

function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/** 루브릭 항목마다 '예상:항목' 열이 피드백 시트에 있도록 맞춘다. 이미 있는 열은 지우지 않는다. */
function syncFeedbackColumns(silent) {
  if (silent !== true) uiOnly_();
  const sh = sheet_(SHEET.FEEDBACK);
  const rubric = getRubric_();
  let map = headerMap_(sh);
  rubric.forEach(r => {
    const h = FB_SCORE_PREFIX + r.item;
    if (map[h]) return;
    const before = map['예상 합계'];
    if (before) {
      sh.insertColumnBefore(before);
      sh.getRange(1, before).setValue(h);
    } else {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(h);
    }
    map = headerMap_(sh);
  });
  styleFeedbackSheet_(sh);
  if (silent !== true) SpreadsheetApp.getActive().toast('피드백 시트의 점수 열을 루브릭에 맞췄습니다.');
}

function styleSheets_() {
  [SHEET.SETTINGS, SHEET.STAGES, SHEET.RUBRIC, SHEET.ROSTER, SHEET.SUBMIT, SHEET.FEEDBACK].forEach(n => {
    const sh = ss_().getSheetByName(n);
    if (!sh) return;
    sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn()))
      .setBackground('#2F5DA8').setFontColor('#ffffff').setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  const st = ss_().getSheetByName(SHEET.STAGES);
  if (st) {
    const m = headerMap_(st);
    st.getRange(2, m['마감일시'], 50, 1).setNumberFormat('yyyy-mm-dd hh:mm');
    const yn = SpreadsheetApp.newDataValidation().requireValueInList(['예', '아니오'], true).build();
    st.getRange(2, m['열림'], 50, 1).setDataValidation(yn);
    st.getRange(2, m['마감 후 제출 허용'], 50, 1).setDataValidation(yn);
    st.getRange(2, m['제출 방식'], 50, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['파일+글', '파일', '글'], true).build());
  }
}

function styleFeedbackSheet_(sh) {
  const m = headerMap_(sh);
  const last = sh.getLastColumn();
  sh.getRange(1, 1, 1, last).setBackground('#2F5DA8').setFontColor('#ffffff').setFontWeight('bold');
  Object.keys(m).forEach(h => {
    if (h.indexOf(FB_SCORE_PREFIX) === 0 || h === '예상 합계') sh.getRange(1, m[h]).setBackground('#FFD93B').setFontColor('#222222');
  });
  const rows = Math.max(sh.getMaxRows() - 1, 1);
  if (m['상태']) {
    sh.getRange(2, m['상태'], rows, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(Object.values(STATUS), true).build());
  }
  if (m['공개']) sh.getRange(2, m['공개'], rows, 1).insertCheckboxes();
  if (m['공개일']) sh.getRange(2, m['공개일'], rows, 1).setNumberFormat('yyyy-mm-dd');
  FB_CONTENT_COLS.forEach(h => { if (m[h]) sh.setColumnWidth(m[h], 260); });
  if (m['학생 메모']) sh.setColumnWidth(m['학생 메모'], 260);
  sh.setFrozenColumns(3);
}

// ───────────────────────── 트리거 ─────────────────────────

function installDailyTrigger() {
  uiOnly_();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'dailyArrivalJob')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyArrivalJob').timeBased().everyDays(1).atHour(8).create();
  SpreadsheetApp.getUi().alert('매일 아침 8시쯤 공개일이 된 피드백의 도착 메일을 보냅니다.');
}

function dailyArrivalJob() {
  sendArrivalMails_();
}

/** (단순 트리거) 피드백 시트 편집: 예상 합계 다시 계산, PDF 만든 뒤 내용을 고치면 상태를 '확정'으로 되돌림 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== SHEET.FEEDBACK || e.range.getRow() < 2) return;
  const m = headerMap_(sh);
  const head = Object.keys(m).find(h => m[h] === e.range.getColumn()) || '';
  const isScore = head.indexOf(FB_SCORE_PREFIX) === 0;
  if (!isScore && FB_CONTENT_COLS.indexOf(head) < 0) return;
  const row = e.range.getRow();
  if (isScore) recalcTotal_(sh, m, row);
  const status = sh.getRange(row, m['상태']).getValue();
  if (status === STATUS.PDF) {
    sh.getRange(row, m['상태']).setValue(STATUS.DONE).setNote('PDF를 만든 뒤 내용이 바뀌었습니다. PDF를 다시 만드세요.');
  }
}

function recalcTotal_(sh, m, row) {
  const rubric = getRubric_();
  let sum = 0, any = false;
  rubric.forEach(r => {
    const c = m[FB_SCORE_PREFIX + r.item];
    if (!c) return;
    const v = sh.getRange(row, c).getValue();
    if (v !== '' && !isNaN(v)) { sum += Number(v); any = true; }
  });
  sh.getRange(row, m['예상 합계']).setValue(any ? sum : '');
}

function menuSetApiKey() {
  uiOnly_();
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Claude API 키 등록', 'console.anthropic.com 에서 발급한 키를 붙여 넣으세요.\n(스크립트 속성에만 저장되고 시트에는 남지 않습니다.)', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const key = res.getResponseText().trim();
  if (!key) return;
  PropertiesService.getScriptProperties().setProperty('ANTHROPIC_API_KEY', key);
  ui.alert('API 키를 저장했습니다.');
}
