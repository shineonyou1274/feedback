/**
 * 학생용 웹앱: 과제 제출, 제출 확인, 피드백 PDF 보기
 */

function doGet(e) {
  const isAdmin = e && e.parameter && e.parameter.page === 'admin';
  const t = HtmlService.createTemplateFromFile(isAdmin ? 'Admin' : 'Index');
  let url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (err) { /* 배포 전 */ }
  t.baseUrl = url;
  let title = '과제 제출·피드백';
  try { title = `${getSettings_()[KEY.TITLE]} · ${isAdmin ? '선생님 관리' : '제출과 피드백'}`; } catch (err) { /* 설정 전 */ }
  return t.evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/** 학생 화면이 처음 불러오는 정보 */
function getPublicConfig() {
  const s = getSettings_();
  const now = new Date();
  return {
    course: String(s[KEY.COURSE] || ''),
    title: String(s[KEY.TITLE] || ''),
    guide: String(s[KEY.GUIDE] || ''),
    maxMB: Number(s[KEY.MAX_MB]) || 10,
    exts: String(s[KEY.EXTS] || '').split(',').map(x => x.trim().toLowerCase()).filter(String),
    stages: getStages_().map(st => ({
      id: st.id,
      name: st.name,
      desc: st.desc,
      mode: st.mode,
      deadline: st.deadline ? st.deadline.getTime() : null,
      deadlineText: st.deadline ? fmtTime_(st.deadline) : '',
      canSubmit: canSubmit_(st, now),
      open: st.open,
      allowLate: st.allowLate,
    })),
    rubric: getRubric_().map(r => ({ item: r.item, max: r.max, high: r.high })),
  };
}

function canSubmit_(st, now) {
  if (!st.open) return false;
  if (st.deadline && now > st.deadline && !st.allowLate) return false;
  return true;
}

/**
 * 과제 제출
 * payload: {studentId, name, email, stageId, did, question, body, file:{name, mimeType, data(base64)}}
 */
function submitAssignment(payload) {
  const p = payload || {};
  const s = getSettings_();
  const id = normId_(p.studentId);
  const name = normName_(p.name);
  if (!/^\d{4,6}$/.test(id)) throw new Error('학번을 숫자로 바르게 입력하세요.');
  if (!name) throw new Error('이름을 입력하세요.');

  const stage = getStages_().find(st => st.id === String(p.stageId));
  if (!stage) throw new Error('제출할 단계를 고르세요.');
  if (!canSubmit_(stage, new Date())) throw new Error(`'${stage.name}' 단계는 지금 제출할 수 없습니다.`);

  const roster = getRoster_();
  const me = roster.find(r => r.id === id && r.name === name);
  if (yes_(s[KEY.ROSTER_ONLY]) && !me) throw new Error('명단에 없는 학번·이름입니다. 선생님께 확인하세요.');

  const did = String(p.did || '').trim().slice(0, 3000);
  const question = String(p.question || '').trim().slice(0, 2000);
  const body = String(p.body || '').trim().slice(0, 20000);
  if (!did) throw new Error('"이번 단계에서 한 일"을 적어 주세요.');
  const needFile = stage.mode === '파일' || stage.mode === '파일+글';
  const hasFile = p.file && p.file.data;
  if (stage.mode === '파일' && !hasFile) throw new Error('파일을 올려 주세요.');
  if (stage.mode === '글' && !body) throw new Error('본문을 적어 주세요.');
  if (stage.mode === '파일+글' && !hasFile && !body) throw new Error('파일을 올리거나 본문을 적어 주세요.');

  let email = String(p.email || '').trim();
  if (me && me.email) email = me.email; // 명단 이메일을 우선한다.
  if (email && !isEmail_(email)) throw new Error('이메일 주소 형식이 바르지 않습니다.');

  // 파일 검사는 잠금 전에 끝낸다.
  let blob = null;
  if (needFile && hasFile) {
    const maxMB = Number(s[KEY.MAX_MB]) || 10;
    const exts = String(s[KEY.EXTS] || '').split(',').map(x => x.trim().toLowerCase()).filter(String);
    const fname = String(p.file.name || 'file').replace(/[\\/:*?"<>|]/g, '_').slice(0, 120);
    const ext = (fname.split('.').pop() || '').toLowerCase();
    if (exts.length && exts.indexOf(ext) < 0) throw new Error(`올릴 수 있는 파일 형식: ${exts.join(', ')}`);
    const bytes = Utilities.base64Decode(String(p.file.data));
    if (bytes.length > maxMB * 1024 * 1024) throw new Error(`파일은 ${maxMB}MB 이하만 올릴 수 있습니다.`);
    blob = Utilities.newBlob(bytes, p.file.mimeType || 'application/octet-stream', `${id}_${name}_${stage.id}_${fname}`);
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let receipt, now, attempt, fileName = '', fileUrl = '', fileId = '';
  try {
    const sh = sheet_(SHEET.SUBMIT);
    const rows = readRows_(sh);
    attempt = rows.filter(r => normId_(r['학번']) === id && String(r['단계ID']) === stage.id).length + 1;
    now = new Date();
    receipt = `${stage.id}-${id}-${Utilities.formatDate(now, 'Asia/Seoul', 'MMddHHmmss')}`;

    if (blob) {
      const folder = stageFolder_(stage);
      const f = folder.createFile(blob);
      fileName = String(p.file.name);
      fileUrl = f.getUrl();
      fileId = f.getId();
    }

    const map = headerMap_(sh);
    const record = {
      '접수번호': receipt, '제출시각': now, '학번': id, '이름': name, '이메일': email,
      '단계ID': stage.id, '단계명': stage.name, '이번 단계에서 한 일': did, '어려웠던 점·질문': question,
      '본문': body, '파일명': fileName, '파일링크': fileUrl, '파일ID': fileId, '제출차수': attempt, '완료메일': '',
    };
    const line = new Array(sh.getLastColumn()).fill('');
    Object.keys(record).forEach(k => { if (map[k]) line[map[k] - 1] = record[k]; });
    sh.appendRow(line);

    addFeedbackRow_({ receipt, id, name, stage, now, fileName, fileUrl, did, question, attempt });
    if (!me) addToRoster_(id, name, email, roster);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  let mailed = false;
  if (email && yes_(s[KEY.SUBMIT_MAIL])) {
    try {
      sendSubmitMail_({ email, id, name, stage, receipt, now, fileName, did, attempt });
      markSubmitMailed_(receipt);
      mailed = true;
    } catch (err) {
      console.error('제출 완료 메일 실패', err);
    }
  }

  return {
    receipt,
    time: fmtTime_(now),
    stageName: stage.name,
    attempt,
    fileName,
    mailed,
    late: !!(stage.deadline && now > stage.deadline),
  };
}

function stageFolder_(stage) {
  const s = getSettings_();
  const root = DriveApp.getFolderById(String(s[KEY.SUBMIT_FOLDER]));
  return getOrCreateFolder_(root, `${stage.id} ${stage.name}`);
}

function addFeedbackRow_(o) {
  const sh = sheet_(SHEET.FEEDBACK);
  const map = headerMap_(sh);
  const memo = `[한 일] ${o.did}` + (o.question ? `\n[어려운 점·질문] ${o.question}` : '') + (o.attempt > 1 ? `\n(${o.attempt}번째 제출)` : '');
  const record = {
    '접수번호': o.receipt, '학번': o.id, '이름': o.name, '단계ID': o.stage.id, '단계명': o.stage.name,
    '제출시각': o.now, '학생 메모': memo, '상태': STATUS.NEW, '공개': false,
  };
  const line = new Array(sh.getLastColumn()).fill('');
  Object.keys(record).forEach(k => { if (map[k]) line[map[k] - 1] = record[k]; });
  // appendRow는 체크박스만 있는 빈 줄도 '내용 있음'으로 보므로, 접수번호가 있는 마지막 줄 바로 아래에 쓴다.
  const row = lastDataRow_(sh, map['접수번호']) + 1;
  if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 10);
  sh.getRange(row, 1, 1, line.length).setValues([line]);
  if (o.fileUrl && map['제출파일']) {
    sh.getRange(row, map['제출파일']).setFormula(`=HYPERLINK("${o.fileUrl}","${String(o.fileName).replace(/"/g, '')}")`);
  }
  if (map['공개']) sh.getRange(row, map['공개']).insertCheckboxes();
  if (map['상태']) {
    sh.getRange(row, map['상태']).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(Object.values(STATUS), true).build());
  }
}

/** 명단에 없는 학생이 제출하면 명단에 자동으로 넣는다 (미제출 확인과 메일 발송에 쓰인다). */
function addToRoster_(id, name, email, roster) {
  if ((roster || getRoster_()).some(r => r.id === id)) return;
  let sh = ss_().getSheetByName(SHEET.ROSTER);
  if (!sh) {
    sh = ss_().insertSheet(SHEET.ROSTER);
    sh.getRange(1, 1, 1, 3).setValues([['학번', '이름', '이메일']]);
  }
  const row = lastDataRow_(sh, 1) + 1;
  sh.getRange(row, 1).setNumberFormat('@');
  sh.getRange(row, 1, 1, 3).setValues([[id, name, email || '']]);
}

function markSubmitMailed_(receipt) {
  const sh = sheet_(SHEET.SUBMIT);
  const map = headerMap_(sh);
  const r = readRows_(sh).find(x => x['접수번호'] === receipt);
  if (r) sh.getRange(r._row, map['완료메일']).setValue(new Date());
}

// ───────────────────────── 제출 확인 ─────────────────────────

const LOOKUP_MAX_FAIL = 5;
const LOOKUP_BLOCK_SEC = 600;

/** 학번+이름으로 제출 기록, 단계별 진행, 공개된 피드백을 돌려준다. */
function lookup(studentId, name) {
  const id = normId_(studentId);
  const nm = normName_(name);
  if (!id || !nm) throw new Error('학번과 이름을 입력하세요.');

  const cache = CacheService.getScriptCache();
  const failKey = 'fail_' + id;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= LOOKUP_MAX_FAIL) throw new Error('여러 번 틀려서 잠시 조회를 막았습니다. 10분 뒤에 다시 해 보세요.');

  const subs = readRows_(sheet_(SHEET.SUBMIT))
    .filter(r => normId_(r['학번']) === id && normName_(r['이름']) === nm);
  if (!subs.length) {
    cache.put(failKey, String(fails + 1), LOOKUP_BLOCK_SEC);
    return { found: false };
  }
  cache.remove(failKey);

  const fbByReceipt = {};
  readRows_(sheet_(SHEET.FEEDBACK)).forEach(r => { fbByReceipt[r['접수번호']] = r; });
  const rubric = getRubric_();
  const maxTotal = rubric.reduce((a, r) => a + r.max, 0);
  const stages = getStages_();
  const now = new Date();

  const items = subs.map(r => {
    const fb = fbByReceipt[r['접수번호']] || {};
    const released = isReleased_(fb, now);
    const out = {
      receipt: r['접수번호'],
      time: fmtTime_(r['제출시각']),
      ts: r['제출시각'] instanceof Date ? r['제출시각'].getTime() : 0,
      stageId: String(r['단계ID']),
      stageName: r['단계명'],
      attempt: r['제출차수'],
      did: r['이번 단계에서 한 일'],
      question: r['어려웠던 점·질문'],
      fileName: r['파일명'],
      hasBody: !!r['본문'],
      status: released ? 'arrived' : (fb['상태'] && fb['상태'] !== STATUS.NEW ? 'writing' : 'reading'),
    };
    if (released) {
      out.scores = rubric.map(rb => ({ item: rb.item, max: rb.max, score: fb[FB_SCORE_PREFIX + rb.item] }));
      out.total = fb['예상 합계'];
      out.maxTotal = maxTotal;
      out.summary = fb['한 줄 총평'];
      out.next = fb['다음 단계까지 할 일'];
      out.hasPdf = !!fb['PDF ID'];
    }
    return out;
  }).sort((a, b) => b.ts - a.ts);

  const progress = stages.map(st => {
    const mine = items.filter(i => i.stageId === st.id);
    const arrived = mine.find(i => i.status === 'arrived');
    return {
      id: st.id, name: st.name,
      submitted: mine.length > 0,
      feedback: !!arrived,
      total: arrived ? arrived.total : null,
      maxTotal,
    };
  });

  return { found: true, name: nm, items, progress };
}

/** 공개 체크 + 공개일 지남 + PDF 있음 */
function isReleased_(fb, now) {
  if (!fb || fb['공개'] !== true || !fb['PDF ID']) return false;
  const d = fb['공개일'];
  if (d instanceof Date) {
    const day = new Date(d); day.setHours(0, 0, 0, 0);
    if (now < day) return false;
  }
  return true;
}

/** 본인 확인 뒤 공개된 피드백 PDF를 base64로 돌려준다. 파일 공유 설정을 바꾸지 않아도 된다. */
function getFeedbackPdf(studentId, name, receipt) {
  const id = normId_(studentId);
  const nm = normName_(name);
  const fb = readRows_(sheet_(SHEET.FEEDBACK))
    .find(r => r['접수번호'] === receipt && normId_(r['학번']) === id && normName_(r['이름']) === nm);
  if (!fb || !isReleased_(fb, new Date())) throw new Error('아직 볼 수 있는 피드백이 없습니다.');
  const file = DriveApp.getFileById(String(fb['PDF ID']));
  return {
    fileName: file.getName(),
    data: Utilities.base64Encode(file.getBlob().getBytes()),
  };
}
