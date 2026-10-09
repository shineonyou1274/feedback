/**
 * 확정 → PDF 만들기 → 공개
 */

const C = { yellow: '#FFD93B', denim: '#2F5DA8', ink: '#222222', soft: '#FFF6CC' };

// ───────────────────────── PDF (필요할 때 자동으로 만든다) ─────────────────────────

function pdfContext_() {
  const ctx = {
    settings: getSettings_(),
    rubric: getRubric_(),
    stages: getStages_(),
    subByReceipt: {},
    fbRows: readRows_(sheet_(SHEET.FEEDBACK)),
  };
  readRows_(sheet_(SHEET.SUBMIT)).forEach(r => { ctx.subByReceipt[r['접수번호']] = r; });
  return ctx;
}

/** PDF에 들어가는 내용의 지문. 내용이 바뀌면 PDF를 새로 만든다. */
function pdfHash_(row, ctx) {
  const parts = [row['이름'], row['단계명'], row['예상 합계']]
    .concat(ctx.rubric.map(r => r.item + '=' + row[FB_SCORE_PREFIX + r.item]))
    .concat(FB_CONTENT_COLS.map(k => row[k]))
    .concat(growth_(row, ctx).map(h => h.stage + h.total));
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(parts), Utilities.Charset.UTF_8);
  return Utilities.base64Encode(raw);
}

/**
 * 접수번호의 PDF 파일을 돌려준다. 없거나 내용이 바뀌었으면 새로 만든다.
 * 선생님은 PDF를 따로 만들 필요가 없다. 학생이 열거나 선생님이 미리 볼 때 만들어진다.
 */
function ensurePdf_(receipt) {
  const sh = sheet_(SHEET.FEEDBACK);
  const map = headerMap_(sh);
  const ctx = pdfContext_();
  const row = ctx.fbRows.find(r => r['접수번호'] === receipt);
  if (!row) throw new Error('해당 제출을 찾지 못했습니다.');
  if (!hasContent_(row)) throw new Error('아직 피드백 내용이 없습니다.');
  const hash = pdfHash_(row, ctx);
  if (row['PDF ID'] && row['PDF 해시'] === hash) {
    try {
      const f = DriveApp.getFileById(String(row['PDF ID']));
      if (!f.isTrashed()) return f;
    } catch (e) { /* 지워졌으면 새로 만든다 */ }
  }
  const pdf = buildPdf_(row, ctx);
  writeCells_(sh, row._row, map, { 'PDF ID': pdf.getId(), 'PDF 해시': hash });
  if (map['PDF']) sh.getRange(row._row, map['PDF']).setFormula(`=HYPERLINK("${pdf.getUrl()}","PDF 열기")`);
  return pdf;
}

/** 시트 메뉴: 고른 학생의 PDF를 만들어 새 탭에서 연다 (공개 전에 미리 보기). */
function menuPreviewPdf() {
  uiOnly_();
  const rows = selectedFeedbackRows_();
  const all = readRows_(sheet_(SHEET.FEEDBACK));
  const row = all.find(r => r._row === rows[0]);
  if (!row) throw new Error('학생 행을 하나 고르세요.');
  const pdf = ensurePdf_(row['접수번호']);
  const url = pdf.getUrl();
  SpreadsheetApp.getUi().showModelessDialog(HtmlService.createHtmlOutput(
    `<div style="font-family:sans-serif;padding:8px"><p><b>${escHtml_(row['이름'])}</b> 학생 PDF를 엽니다.</p>
     <p><a href="${escHtml_(url)}" target="_blank" style="font-size:16px;font-weight:bold">👉 열리지 않으면 여기를 누르세요</a></p></div>
     <script>window.open(${JSON.stringify(url)}, '_blank');</script>`).setWidth(360).setHeight(130), 'PDF 미리 보기');
}

/** 피드백 문서를 구글 문서로 만들어 PDF로 저장하고, 문서는 지운다. */
function buildPdf_(row, ctx) {
  const s = ctx.settings;
  const sub = ctx.subByReceipt[row['접수번호']] || {};
  const folder = DriveApp.getFolderById(String(s[KEY.PDF_FOLDER]));
  const baseName = `${row['학번']}_${row['이름']}_${row['단계명']}_피드백`;
  const doc = DocumentApp.create(baseName);
  const docFile = DriveApp.getFileById(doc.getId());
  docFile.moveTo(folder);

  try {
    const body = doc.getBody();
    body.setMarginTop(48).setMarginBottom(48).setMarginLeft(56).setMarginRight(56);
    body.clear();

    // Paragraph.setText()는 값을 돌려주지 않으므로 이어 쓰지 않는다.
    const t = body.getParagraphs()[0] || body.appendParagraph('');
    t.setText(`${s[KEY.COURSE]} · ${s[KEY.TITLE]}`);
    t.setHeading(DocumentApp.ParagraphHeading.TITLE);
    t.editAsText().setFontSize(22).setBold(true).setForegroundColor(C.ink);
    const sub1 = body.appendParagraph(`${row['단계명']} 단계 피드백${s[KEY.STANDARD] ? '  |  ' + s[KEY.STANDARD] : ''}`);
    sub1.editAsText().setFontSize(12).setForegroundColor(C.denim).setBold(true);

    // 1. 제출 정보
    heading_(body, '1. 제출 정보');
    infoTable_(body, [
      ['학번', String(row['학번'])],
      ['이름', String(row['이름'])],
      ['단계', `${row['단계명']}${sub['제출차수'] > 1 ? ` (${sub['제출차수']}번째 제출)` : ''}`],
      ['제출 시각', fmtTime_(row['제출시각'])],
      ['제출 파일', String(sub['파일명'] || '(없음)')],
    ]);

    // 2. 나의 과정 기록
    heading_(body, '2. 내가 적은 과정 기록');
    infoTable_(body, [
      ['이번 단계에서 한 일', String(sub['이번 단계에서 한 일'] || '')],
      ['어려웠던 점·질문', String(sub['어려웠던 점·질문'] || '(없음)')],
    ]);

    // 3. 예상 점수
    heading_(body, '3. 예상 점수');
    note_(body, '지금 상태 그대로 최종 보고서를 낸다고 가정한 점수입니다. 다음 단계에서 고치면 얼마든지 올라갑니다.');
    const scoreRows = [['평가 항목', '배점', '예상 점수', '수준']];
    let sum = 0, max = 0;
    ctx.rubric.forEach(r => {
      const v = row[FB_SCORE_PREFIX + r.item];
      const has = v !== '' && v != null && !isNaN(v);
      if (has) sum += Number(v);
      max += r.max;
      scoreRows.push([r.item, String(r.max), has ? String(v) : '-', has ? levelOf_(v, r.max) : '-']);
    });
    scoreRows.push(['합계', String(max), String(row['예상 합계'] !== '' ? row['예상 합계'] : sum), levelOf_(row['예상 합계'] || sum, max)]);
    gridTable_(body, scoreRows, [0.52, 0.14, 0.18, 0.16]);

    // 4. 성장 기록 (이전 단계 예상 점수)
    const history = growth_(row, ctx);
    if (history.length > 1) {
      heading_(body, '4. 단계별 성장 기록');
      gridTable_(body, [['단계', '예상 합계']].concat(history.map(h => [h.stage, `${h.total} / ${max}`])), [0.6, 0.4]);
    }

    // 5. 선생님 의견
    heading_(body, history.length > 1 ? '5. 선생님 의견' : '4. 선생님 의견');
    ['잘한 점', '보완할 점', '다음 단계까지 할 일'].forEach(k => {
      const p = body.appendParagraph(`■ ${k}`);
      p.editAsText().setBold(true).setFontSize(12).setForegroundColor(C.denim);
      const v = body.appendParagraph(String(row[k] || ''));
      v.editAsText().setBold(false).setFontSize(11).setForegroundColor(C.ink);
      v.setLineSpacing(1.5).setSpacingAfter(8);
    });
    if (row['한 줄 총평']) {
      const box = body.appendTable([[`한 줄 총평  |  ${row['한 줄 총평']}`]]);
      box.setBorderColor(C.ink);
      const cell = box.getCell(0, 0);
      cell.setBackgroundColor(C.soft).setPaddingTop(10).setPaddingBottom(10).setPaddingLeft(12);
      cell.editAsText().setBold(true).setFontSize(12).setForegroundColor(C.ink);
    }

    const foot = body.appendParagraph(
      `작성일 ${Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd')}` +
      (s[KEY.TEACHER] ? `  ·  ${s[KEY.TEACHER]} 선생님` : '') + '\n© shiny-peace');
    foot.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    foot.editAsText().setFontSize(9).setForegroundColor('#777777').setBold(false);

    doc.saveAndClose();

    // 예전 PDF가 있으면 휴지통으로 보내고 새로 만든다.
    if (row['PDF ID']) {
      try { DriveApp.getFileById(String(row['PDF ID'])).setTrashed(true); } catch (e) { /* 이미 없음 */ }
    }
    const pdf = folder.createFile(docFile.getAs(MimeType.PDF).setName(baseName + '.pdf'));
    return pdf;
  } finally {
    docFile.setTrashed(true);
  }
}

function heading_(body, text) {
  const h = body.appendParagraph(text);
  h.setHeading(DocumentApp.ParagraphHeading.HEADING2).setSpacingBefore(14).setSpacingAfter(6);
  h.editAsText().setFontSize(14).setBold(true).setForegroundColor(C.ink);
  return h;
}

function note_(body, text) {
  const p = body.appendParagraph(text);
  p.editAsText().setFontSize(9).setItalic(true).setBold(false).setForegroundColor('#666666');
  p.setSpacingAfter(4);
  return p;
}

function infoTable_(body, rows) {
  const tb = body.appendTable(rows);
  tb.setBorderColor(C.ink);
  for (let i = 0; i < rows.length; i++) {
    const k = tb.getCell(i, 0), v = tb.getCell(i, 1);
    k.setBackgroundColor(C.yellow).setWidth(130);
    k.editAsText().setBold(true).setFontSize(10).setForegroundColor(C.ink);
    v.editAsText().setBold(false).setFontSize(10).setForegroundColor(C.ink);
  }
  return tb;
}

function gridTable_(body, rows, ratios) {
  const tb = body.appendTable(rows);
  tb.setBorderColor(C.ink);
  const width = 480;
  const isScore = rows[0][0] === '평가 항목';
  for (let i = 0; i < rows.length; i++) {
    for (let j = 0; j < rows[i].length; j++) {
      const c = tb.getCell(i, j);
      if (ratios && i === 0) c.setWidth(width * ratios[j]);
      c.editAsText().setFontSize(10).setBold(i === 0 || (isScore && i === rows.length - 1)).setForegroundColor(i === 0 ? '#ffffff' : C.ink);
      if (i === 0) c.setBackgroundColor(C.denim);
      else if (isScore && i === rows.length - 1) c.setBackgroundColor(C.soft);
      if (j > 0) c.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    }
  }
  return tb;
}

/** 같은 학생이 이전 단계부터 지금까지 받은 예상 합계 */
function growth_(row, ctx) {
  const order = {};
  ctx.stages.forEach((s, i) => { order[s.id] = i; });
  const cur = order[row['단계ID']];
  const latest = {};
  ctx.fbRows
    .filter(r => normId_(r['학번']) === normId_(row['학번']) && order[r['단계ID']] <= cur && r['예상 합계'] !== '')
    .filter(r => r['접수번호'] === row['접수번호'] || hasContent_(r))
    .forEach(r => {
      const prev = latest[r['단계ID']];
      if (r['단계ID'] === row['단계ID'] && r['접수번호'] !== row['접수번호']) return; // 같은 단계는 지금 행만
      if (!prev || new Date(r['제출시각']) > new Date(prev['제출시각'])) latest[r['단계ID']] = r;
    });
  return Object.keys(latest)
    .sort((a, b) => order[a] - order[b])
    .map(k => ({ stage: latest[k]['단계명'], total: latest[k]['예상 합계'] }));
}

// ───────────────────────── 공개 ─────────────────────────

/**
 * 공개 켜기/끄기. 켜면 공개 시각을 남기고 도착 메일을 보낸다 (한 번만).
 * 결과: {count, empty[], mailed}
 */
function setPublished_(rows, on) {
  const sh = sheet_(SHEET.FEEDBACK);
  const map = headerMap_(sh);
  const all = readRows_(sh);
  let count = 0, mailed = 0;
  const empty = [];
  rows.forEach(r => {
    const row = all.find(x => x._row === r);
    if (!row || !row['접수번호']) return;
    if (on && !hasContent_(row)) {
      empty.push(String(row['이름']));
      sh.getRange(r, map['공개']).setValue(false);
      return;
    }
    sh.getRange(r, map['공개']).setValue(!!on);
    if (on) {
      if (!row['공개시각'] && map['공개시각']) sh.getRange(r, map['공개시각']).setValue(new Date());
      if (!row['도착메일'] && sendArrivalMailFor_(row)) {
        sh.getRange(r, map['도착메일']).setValue(new Date());
        mailed++;
      }
    }
    count++;
  });
  return { count, empty, mailed };
}

/** (설치형 트리거) 시트에서 '공개' 칸을 체크하면 바로 공개 처리와 도착 메일 */
function onFeedbackEditInstalled(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== SHEET.FEEDBACK) return;
  const map = headerMap_(sh);
  if (!map['공개'] || e.range.getColumn() > map['공개'] || e.range.getLastColumn() < map['공개']) return;
  const rows = [];
  for (let r = Math.max(2, e.range.getRow()); r <= e.range.getLastRow(); r++) {
    if (sh.getRange(r, map['공개']).getValue() === true) rows.push(r);
  }
  if (!rows.length) return;
  const res = setPublished_(rows, true);
  if (res.empty.length) {
    SpreadsheetApp.getActive().toast(`의견이 비어 있어 공개하지 않았습니다: ${res.empty.join(', ')}`, '공개', 6);
  } else {
    SpreadsheetApp.getActive().toast(`${res.count}명 공개${res.mailed ? ` · 도착 메일 ${res.mailed}통` : ''}`, '공개', 4);
  }
}

/** 시트에서 공개 체크가 바로 동작하도록 설치형 편집 트리거를 둔다. 예전 매일 트리거는 지운다. */
function ensureTriggers_() {
  const cache = CacheService.getScriptCache();
  if (cache.get('triggers_ok')) return;
  try {
    const ss = ss_();
    const ts = ScriptApp.getProjectTriggers();
    ts.filter(t => t.getHandlerFunction() === 'dailyArrivalJob').forEach(t => ScriptApp.deleteTrigger(t));
    if (!ts.some(t => t.getHandlerFunction() === 'onFeedbackEditInstalled')) {
      ScriptApp.newTrigger('onFeedbackEditInstalled').forSpreadsheet(ss).onEdit().create();
    }
    cache.put('triggers_ok', '1', 6 * 60 * 60);
  } catch (err) {
    console.error('트리거 설치 실패', err);
  }
}
