/**
 * 교사용 웹 관리 화면 (웹앱 주소 + ?page=admin)
 * 모든 admin* 함수는 맨 먼저 토큰을 확인한다.
 */

const ADMIN_TTL_SEC = 6 * 60 * 60;
const ADMIN_MAX_FAIL = 10;

// ───────────────────────── 비밀번호와 토큰 ─────────────────────────

function hashPw_(pw, salt) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(raw);
}

function hasAdminPassword_() {
  return !!PropertiesService.getScriptProperties().getProperty('ADMIN_PW_HASH');
}

function setAdminPassword_(pw) {
  if (String(pw).length < 6) throw new Error('비밀번호는 6자 이상으로 정하세요.');
  const salt = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperties({
    ADMIN_PW_SALT: salt,
    ADMIN_PW_HASH: hashPw_(String(pw), salt),
  });
}

/** 시트 메뉴: 관리자 비밀번호 정하기 (처음 정하는 일은 시트 주인만 할 수 있다) */
function menuSetAdminPassword() {
  uiOnly_();
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('관리자 비밀번호', '웹 관리 화면(웹앱 주소 끝에 ?page=admin)에 들어갈 비밀번호를 정하세요. 6자 이상.', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  setAdminPassword_(res.getResponseText().trim());
  ui.alert('비밀번호를 정했습니다.');
}

function adminStatus() {
  return { hasPassword: hasAdminPassword_() };
}

function adminLogin(pw) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('adm_fail') || 0);
  if (fails >= ADMIN_MAX_FAIL) throw new Error('비밀번호를 여러 번 틀렸습니다. 10분 뒤에 다시 해 보세요.');
  const props = PropertiesService.getScriptProperties();
  const hash = props.getProperty('ADMIN_PW_HASH');
  if (!hash) throw new Error('관리자 비밀번호가 없습니다. 시트 메뉴 "관리자 비밀번호 정하기"를 먼저 실행하세요.');
  if (hashPw_(String(pw || ''), props.getProperty('ADMIN_PW_SALT')) !== hash) {
    cache.put('adm_fail', String(fails + 1), 600);
    throw new Error('비밀번호가 맞지 않습니다.');
  }
  cache.remove('adm_fail');
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('adm_' + token, '1', ADMIN_TTL_SEC);
  return { token };
}

function adminLogout(token) {
  CacheService.getScriptCache().remove('adm_' + token);
  return true;
}

function requireAdmin_(token) {
  if (!token || !CacheService.getScriptCache().get('adm_' + token)) {
    throw new Error('로그인이 끝났습니다. 다시 로그인하세요.');
  }
}

function adminChangePassword(token, oldPw, newPw) {
  requireAdmin_(token);
  const props = PropertiesService.getScriptProperties();
  if (hashPw_(String(oldPw || ''), props.getProperty('ADMIN_PW_SALT')) !== props.getProperty('ADMIN_PW_HASH')) {
    throw new Error('지금 비밀번호가 맞지 않습니다.');
  }
  setAdminPassword_(String(newPw || ''));
  return true;
}

// ───────────────────────── 제출 현황 ─────────────────────────

function adminDashboard(token) {
  requireAdmin_(token);
  const stages = getStages_();
  const rubric = getRubric_();
  const maxTotal = rubric.reduce((a, r) => a + r.max, 0);
  const subByReceipt = {};
  readRows_(sheet_(SHEET.SUBMIT)).forEach(r => { subByReceipt[r['접수번호']] = r; });
  const now = new Date();

  const rows = readRows_(sheet_(SHEET.FEEDBACK)).filter(r => r['접수번호']).map(fb => {
    const sub = subByReceipt[fb['접수번호']] || {};
    return {
      receipt: fb['접수번호'],
      id: String(fb['학번']),
      name: String(fb['이름']),
      stageId: String(fb['단계ID']),
      stageName: String(fb['단계명']),
      time: fmtTime_(fb['제출시각']),
      ts: fb['제출시각'] instanceof Date ? fb['제출시각'].getTime() : 0,
      attempt: Number(sub['제출차수']) || 1,
      late: isLate_(stages, fb),
      status: fb['상태'] || STATUS.NEW,
      total: fb['예상 합계'],
      published: fb['공개'] === true,
      releaseDate: fb['공개일'] instanceof Date ? Utilities.formatDate(fb['공개일'], 'Asia/Seoul', 'yyyy-MM-dd') : '',
      visible: isReleased_(fb, now),
      hasPdf: !!fb['PDF ID'],
      arrivalMailed: !!fb['도착메일'],
      submitMailed: !!sub['완료메일'],
    };
  }).sort((a, b) => b.ts - a.ts);

  const roster = getRoster_();
  const missing = {};
  stages.forEach(st => {
    const done = new Set(rows.filter(r => r.stageId === st.id).map(r => normId_(r.id)));
    missing[st.id] = roster.filter(p => !done.has(p.id)).map(p => ({ id: p.id, name: p.name }));
  });

  const s = getSettings_();
  return {
    course: String(s[KEY.COURSE] || ''),
    title: String(s[KEY.TITLE] || ''),
    stages: stages.map(st => ({ id: st.id, name: st.name, deadlineText: fmtTime_(st.deadline), open: st.open })),
    rubric: rubric.map(r => ({ item: r.item, max: r.max })),
    maxTotal,
    rows,
    rosterCount: roster.length,
    missing,
    hasApiKey: !!PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY'),
    statuses: Object.values(STATUS),
  };
}

function isLate_(stages, fb) {
  const st = stages.find(s => s.id === String(fb['단계ID']));
  return !!(st && st.deadline && fb['제출시각'] instanceof Date && fb['제출시각'] > st.deadline);
}

// ───────────────────────── 한 학생 피드백 ─────────────────────────

function findFeedback_(receipt) {
  const sh = sheet_(SHEET.FEEDBACK);
  const row = readRows_(sh).find(r => r['접수번호'] === receipt);
  if (!row) throw new Error('해당 제출을 찾지 못했습니다.');
  return { sh, map: headerMap_(sh), row };
}

function adminGetItem(token, receipt) {
  requireAdmin_(token);
  const { row: fb } = findFeedback_(receipt);
  const sub = readRows_(sheet_(SHEET.SUBMIT)).find(r => r['접수번호'] === receipt) || {};
  const rubric = getRubric_();
  const stages = getStages_();
  const stage = stages.find(s => s.id === String(fb['단계ID'])) || {};
  const order = {};
  stages.forEach((s, i) => { order[s.id] = i; });

  const history = readRows_(sheet_(SHEET.FEEDBACK))
    .filter(r => normId_(r['학번']) === normId_(fb['학번']) && r['접수번호'] !== receipt)
    .sort((a, b) => (order[a['단계ID']] - order[b['단계ID']]) || (new Date(a['제출시각']) - new Date(b['제출시각'])))
    .map(r => ({
      receipt: r['접수번호'], stageName: r['단계명'], time: fmtTime_(r['제출시각']),
      status: r['상태'], total: r['예상 합계'], next: r['다음 단계까지 할 일'], improvements: r['보완할 점'],
    }));

  let fileMime = '';
  if (sub['파일ID']) {
    try { fileMime = DriveApp.getFileById(String(sub['파일ID'])).getMimeType(); } catch (e) { fileMime = ''; }
  }

  return {
    receipt,
    id: String(fb['학번']),
    name: String(fb['이름']),
    stageId: String(fb['단계ID']),
    stageName: String(fb['단계명']),
    stageDesc: stage.desc || '',
    stageFocus: stage.focus || '',
    time: fmtTime_(fb['제출시각']),
    late: isLate_(stages, fb),
    attempt: Number(sub['제출차수']) || 1,
    email: String(sub['이메일'] || ''),
    did: String(sub['이번 단계에서 한 일'] || ''),
    question: String(sub['어려웠던 점·질문'] || ''),
    body: String(sub['본문'] || ''),
    fileName: String(sub['파일명'] || ''),
    fileUrl: String(sub['파일링크'] || ''),
    filePreviewable: fileMime === MimeType.PDF,
    scores: rubric.map(r => ({
      item: r.item, max: r.max, high: r.high, mid: r.mid, low: r.low,
      score: fb[FB_SCORE_PREFIX + r.item] === '' ? '' : fb[FB_SCORE_PREFIX + r.item],
    })),
    total: fb['예상 합계'],
    scoreNote: noteOf_(receipt),
    strengths: String(fb['잘한 점'] || ''),
    improvements: String(fb['보완할 점'] || ''),
    nextSteps: String(fb['다음 단계까지 할 일'] || ''),
    summary: String(fb['한 줄 총평'] || ''),
    status: fb['상태'] || STATUS.NEW,
    published: fb['공개'] === true,
    releaseDate: fb['공개일'] instanceof Date ? Utilities.formatDate(fb['공개일'], 'Asia/Seoul', 'yyyy-MM-dd') : '',
    pdfUrl: fb['PDF ID'] ? `https://drive.google.com/file/d/${fb['PDF ID']}/view` : '',
    pdfTime: fmtTime_(fb['PDF 생성시각']),
    aiTime: fmtTime_(fb['AI 초안시각']),
    arrivalMailed: fmtTime_(fb['도착메일']),
    history,
  };
}

function noteOf_(receipt) {
  const { sh, map, row } = findFeedback_(receipt);
  return map['예상 합계'] ? sh.getRange(row._row, map['예상 합계']).getNote() : '';
}

/**
 * data: {scores:{항목: 점수}, strengths, improvements, nextSteps, summary, status, releaseDate('yyyy-MM-dd'|''), published}
 */
function adminSaveFeedback(token, receipt, data) {
  requireAdmin_(token);
  const d = data || {};
  const { sh, map, row } = findFeedback_(receipt);
  const rubric = getRubric_();
  const out = {};
  let total = 0, any = false;
  rubric.forEach(r => {
    const raw = d.scores ? d.scores[r.item] : undefined;
    if (raw === undefined) return;
    if (raw === '' || raw === null) { out[FB_SCORE_PREFIX + r.item] = ''; return; }
    const v = Number(raw);
    if (isNaN(v) || v < 0 || v > r.max) throw new Error(`'${r.item}' 점수는 0부터 ${r.max} 사이로 넣으세요.`);
    out[FB_SCORE_PREFIX + r.item] = v;
  });
  rubric.forEach(r => {
    const k = FB_SCORE_PREFIX + r.item;
    const v = k in out ? out[k] : row[k];
    if (v !== '' && v != null && !isNaN(v)) { total += Number(v); any = true; }
  });
  out['예상 합계'] = any ? total : '';

  const texts = { strengths: '잘한 점', improvements: '보완할 점', nextSteps: '다음 단계까지 할 일', summary: '한 줄 총평' };
  let contentChanged = false;
  Object.keys(texts).forEach(k => {
    if (d[k] === undefined) return;
    const v = String(d[k]).slice(0, 5000);
    if (v !== String(row[texts[k]] || '')) contentChanged = true;
    out[texts[k]] = v;
  });
  rubric.forEach(r => {
    const k = FB_SCORE_PREFIX + r.item;
    if (k in out && String(out[k]) !== String(row[k])) contentChanged = true;
  });

  let status = d.status && Object.values(STATUS).indexOf(d.status) >= 0 ? d.status : row['상태'];
  if (status === STATUS.NEW && contentChanged) status = STATUS.DRAFT;
  // PDF를 만든 뒤 내용을 고치면 다시 만들어야 하므로 확정으로 되돌린다.
  if (row['상태'] === STATUS.PDF && contentChanged && status === STATUS.PDF) status = STATUS.DONE;
  out['상태'] = status;

  if (d.releaseDate !== undefined) {
    out['공개일'] = d.releaseDate ? parseDate_(d.releaseDate) : '';
  }
  if (d.published !== undefined) {
    if (d.published && !row['PDF ID'] && status !== STATUS.PDF) throw new Error('PDF를 먼저 만든 뒤 공개하세요.');
    out['공개'] = !!d.published;
  }
  writeCells_(sh, row._row, map, out);
  if (status === STATUS.DONE && row['상태'] === STATUS.PDF) {
    sh.getRange(row._row, map['상태']).setNote('PDF를 만든 뒤 내용이 바뀌었습니다. PDF를 다시 만드세요.');
  }
  return adminGetItem(token, receipt);
}

function parseDate_(s) {
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error('공개일 형식이 바르지 않습니다.');
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function adminAiDraft(token, receipt) {
  requireAdmin_(token);
  const ctx = aiContext_();
  const row = ctx.fbRows.find(r => r['접수번호'] === receipt);
  if (!row) throw new Error('해당 제출을 찾지 못했습니다.');
  const sh = sheet_(SHEET.FEEDBACK);
  draftRow_(sh, headerMap_(sh), row, ctx);
  return adminGetItem(token, receipt);
}

/** 미작성 제출 전체에 AI 초안. 실행 시간 안에서 할 수 있는 만큼 하고 남은 수를 돌려준다. */
function adminAiDraftPending(token) {
  requireAdmin_(token);
  return makeDrafts_(pendingDraftRows_(), null, false);
}

function receiptsToRows_(receipts) {
  const set = new Set(receipts || []);
  return readRows_(sheet_(SHEET.FEEDBACK)).filter(r => set.has(r['접수번호'])).map(r => r._row);
}

function adminSetStatus(token, receipts, status) {
  requireAdmin_(token);
  if (Object.values(STATUS).indexOf(status) < 0) throw new Error('알 수 없는 상태입니다.');
  const sh = sheet_(SHEET.FEEDBACK);
  const map = headerMap_(sh);
  const rows = receiptsToRows_(receipts);
  rows.forEach(r => sh.getRange(r, map['상태']).setValue(status));
  return { count: rows.length };
}

function adminMakePdfs(token, receipts) {
  requireAdmin_(token);
  const rows = readRows_(sheet_(SHEET.FEEDBACK));
  const r = makePdfs_(receiptsToRows_(receipts), true);
  r.leftReceipts = r.leftRows.map(n => (rows.find(x => x._row === n) || {})['접수번호']).filter(String);
  delete r.leftRows;
  return r;
}

function adminPublish(token, receipts, on, releaseDate) {
  requireAdmin_(token);
  const date = releaseDate === undefined ? undefined : (releaseDate ? parseDate_(releaseDate) : '');
  return setPublish_(receiptsToRows_(receipts), !!on, date);
}

function adminSendArrivalMails(token) {
  requireAdmin_(token);
  return sendArrivalMails_();
}

/** 교사가 제출 파일(PDF)을 관리 화면 안에서 미리 본다. */
function adminSubmissionPdf(token, receipt) {
  requireAdmin_(token);
  const sub = readRows_(sheet_(SHEET.SUBMIT)).find(r => r['접수번호'] === receipt);
  if (!sub || !sub['파일ID']) throw new Error('제출 파일이 없습니다.');
  const f = DriveApp.getFileById(String(sub['파일ID']));
  if (f.getMimeType() !== MimeType.PDF) throw new Error('PDF 파일만 미리 볼 수 있습니다.');
  return { fileName: f.getName(), data: Utilities.base64Encode(f.getBlob().getBytes()) };
}

function adminFeedbackPdf(token, receipt) {
  requireAdmin_(token);
  const { row } = findFeedback_(receipt);
  if (!row['PDF ID']) throw new Error('아직 PDF가 없습니다.');
  const f = DriveApp.getFileById(String(row['PDF ID']));
  return { fileName: f.getName(), data: Utilities.base64Encode(f.getBlob().getBytes()) };
}

// ───────────────────────── 과제 설정 ─────────────────────────

/**
 * 관리 화면에서 고칠 수 있는 설정 키.
 * 파일이 읽히는 순서와 상관없이 쓰도록 함수로 둔다 (Config.gs의 KEY를 맨 위에서 바로 쓰지 않는다).
 */
function editableKeys_() {
  return [
    KEY.COURSE, KEY.TITLE, KEY.STANDARD, KEY.GUIDE, KEY.TEACHER, KEY.WEBAPP_URL,
    KEY.ROSTER_ONLY, KEY.MAX_MB, KEY.EXTS,
    KEY.SUBMIT_MAIL, KEY.SUBMIT_MAIL_SUBJECT, KEY.ARRIVE_MAIL, KEY.ARRIVE_MAIL_SUBJECT,
    KEY.AI_MODEL, KEY.AI_EFFORT,
  ];
}

function adminGetConfig(token) {
  requireAdmin_(token);
  const s = getSettings_();
  const settings = {};
  editableKeys_().forEach(k => { settings[k] = s[k] === undefined ? '' : s[k]; });
  const usedStages = new Set(readRows_(sheet_(SHEET.SUBMIT)).map(r => String(r['단계ID'])));
  return {
    settings,
    keys: editableKeys_(),
    stages: getStages_().map(st => ({
      id: st.id, name: st.name, desc: st.desc, focus: st.focus, mode: st.mode,
      deadline: st.deadline ? Utilities.formatDate(st.deadline, 'Asia/Seoul', "yyyy-MM-dd'T'HH:mm") : '',
      open: st.open, allowLate: st.allowLate, used: usedStages.has(st.id),
    })),
    rubric: getRubric_(),
    roster: getRoster_(),
    hasApiKey: !!PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY'),
    webAppUrl: (() => { try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; } })(),
    dailyTrigger: ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'dailyArrivalJob'),
  };
}

/** cfg: {settings:{}, stages:[...], rubric:[...]} */
function adminSaveConfig(token, cfg) {
  requireAdmin_(token);
  const c = cfg || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (c.settings) {
      editableKeys_().forEach(k => {
        if (c.settings[k] === undefined) return;
        let v = c.settings[k];
        if (k === KEY.MAX_MB) {
          v = Number(v);
          if (!(v > 0 && v <= 25)) throw new Error('파일 최대 크기는 1~25MB 사이로 정하세요.');
        }
        setSetting_(k, typeof v === 'string' ? v.slice(0, 5000) : v);
      });
    }
    if (c.stages) saveStages_(c.stages);
    if (c.rubric) saveRubric_(c.rubric);
  } finally {
    lock.releaseLock();
  }
  return adminGetConfig(token);
}

function saveStages_(list) {
  if (!list.length) throw new Error('단계를 하나 이상 두세요.');
  const sh = sheet_(SHEET.STAGES);
  const old = getStages_();
  const used = new Set(readRows_(sheet_(SHEET.SUBMIT)).map(r => String(r['단계ID'])));
  const ids = new Set();
  let n = old.reduce((m, s) => Math.max(m, Number(String(s.id).replace(/\D/g, '')) || 0), 0);
  const rows = list.map(st => {
    let id = String(st.id || '').trim();
    if (!id) id = 'S' + (++n);
    if (ids.has(id)) throw new Error(`단계ID ${id}가 겹칩니다.`);
    ids.add(id);
    const name = String(st.name || '').trim();
    if (!name) throw new Error('단계 이름을 비워 둘 수 없습니다.');
    let deadline = '';
    if (st.deadline) {
      const m = String(st.deadline).match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
      if (!m) throw new Error(`'${name}' 마감일시 형식이 바르지 않습니다.`);
      deadline = new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]);
    }
    const mode = ['파일+글', '파일', '글'].indexOf(st.mode) >= 0 ? st.mode : '파일+글';
    return [id, name, String(st.desc || ''), String(st.focus || ''), mode, deadline, st.open ? '예' : '아니오', st.allowLate ? '예' : '아니오'];
  });
  old.forEach(s => {
    if (!ids.has(s.id) && used.has(s.id)) throw new Error(`'${s.name}' 단계에는 이미 제출이 있어 지울 수 없습니다. 대신 "열림"을 끄세요.`);
  });
  const head = ['단계ID', '단계명', '학생 안내', '피드백 초점', '제출 방식', '마감일시', '열림', '마감 후 제출 허용'];
  sh.clearContents();
  sh.getRange(1, 1, 1, head.length).setValues([head]);
  sh.getRange(2, 1, rows.length, head.length).setValues(rows);
  sh.getRange(2, 6, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm');
}

function saveRubric_(list) {
  const rows = list
    .filter(r => String(r.item || '').trim())
    .map(r => {
      const max = Number(r.max);
      if (!(max > 0 && max <= 1000)) throw new Error(`'${r.item}' 배점을 바르게 넣으세요.`);
      return [String(r.item).trim(), max, String(r.high || ''), String(r.mid || ''), String(r.low || '')];
    });
  const names = rows.map(r => r[0]);
  if (new Set(names).size !== names.length) throw new Error('평가 항목 이름이 겹칩니다.');
  const sh = sheet_(SHEET.RUBRIC);
  const head = ['평가 항목', '배점', '상', '중', '하'];
  sh.clearContents();
  sh.getRange(1, 1, 1, head.length).setValues([head]);
  if (rows.length) sh.getRange(2, 1, rows.length, head.length).setValues(rows);
  syncFeedbackColumns(true);
}

/** 명단 붙여 넣기: 한 줄에 "학번 이름 이메일" (탭·쉼표·공백 구분) */
function adminSaveRoster(token, text) {
  requireAdmin_(token);
  const rows = [];
  const seen = new Set();
  String(text || '').split(/\r?\n/).forEach((line, i) => {
    const t = line.trim();
    if (!t) return;
    const parts = t.split(/[\t,]+|\s+/).filter(String);
    const id = normId_(parts[0]);
    const name = normName_(parts[1]);
    const email = parts[2] || '';
    if (!id || !name) throw new Error(`${i + 1}번째 줄을 읽지 못했습니다: ${t}`);
    if (email && !isEmail_(email)) throw new Error(`${i + 1}번째 줄 이메일이 바르지 않습니다: ${email}`);
    if (seen.has(id)) throw new Error(`학번 ${id}가 두 번 들어 있습니다.`);
    seen.add(id);
    rows.push([id, name, email]);
  });
  let sh = ss_().getSheetByName(SHEET.ROSTER);
  if (!sh) sh = ss_().insertSheet(SHEET.ROSTER);
  sh.clearContents();
  sh.getRange(1, 1, 1, 3).setValues([['학번', '이름', '이메일']]);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  return { count: rows.length, roster: getRoster_() };
}

function adminSetApiKey(token, key) {
  requireAdmin_(token);
  const k = String(key || '').trim();
  if (!k) throw new Error('키를 넣으세요.');
  PropertiesService.getScriptProperties().setProperty('ANTHROPIC_API_KEY', k);
  return true;
}

function adminToggleDailyTrigger(token, on) {
  requireAdmin_(token);
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'dailyArrivalJob')
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (on) ScriptApp.newTrigger('dailyArrivalJob').timeBased().everyDays(1).atHour(8).create();
  return !!on;
}
