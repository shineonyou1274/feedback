/**
 * 확정 → PDF 만들기 → 공개
 */

const C = { yellow: '#FFD93B', denim: '#2F5DA8', ink: '#222222', soft: '#FFF6CC' };

function menuConfirmSelected() {
  uiOnly_();
  const sh = sheet_(SHEET.FEEDBACK);
  const map = headerMap_(sh);
  const rows = selectedFeedbackRows_();
  rows.forEach(r => sh.getRange(r, map['상태']).setValue(STATUS.DONE));
  SpreadsheetApp.getActive().toast(`${rows.length}행을 확정했습니다. 이제 PDF를 만드세요.`);
}

function menuPdfSelected() {
  uiOnly_();
  alertPdfResult_(makePdfs_(selectedFeedbackRows_(), true));
}

function menuPdfAll() {
  uiOnly_();
  const rows = readRows_(sheet_(SHEET.FEEDBACK)).filter(r => r['상태'] === STATUS.DONE).map(r => r._row);
  alertPdfResult_(makePdfs_(rows, false));
}

function alertPdfResult_(r) {
  SpreadsheetApp.getUi().alert(
    `PDF ${r.done}건을 만들었습니다.` +
    (r.left ? `\n실행 시간 제한 때문에 ${r.left}건이 남았습니다. 같은 메뉴를 한 번 더 실행하세요.` : '') +
    (r.skipped.length ? `\n건너뜀: ${r.skipped.join(', ')}` : '') +
    (r.errors.length ? `\n\n실패:\n${r.errors.join('\n')}` : '') +
    '\n\n학생에게 보이게 하려면 "공개" 칸을 체크하고 공개일을 정하세요. (공개일이 비어 있으면 바로 보입니다.)');
}

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

/** 피드백 시트의 행 번호 목록으로 PDF를 만든다. 결과: {done, left, skipped[], errors[]} */
function makePdfs_(targetRows, allowAnyStatus) {
  const sh = sheet_(SHEET.FEEDBACK);
  const map = headerMap_(sh);
  const ctx = pdfContext_();
  const all = ctx.fbRows;
  const start = Date.now();
  let done = 0, processed = 0;
  const errors = [], skipped = [];
  for (const r of targetRows) {
    if (Date.now() - start > TIME_BUDGET_MS) break;
    processed++;
    const row = all.find(x => x._row === r);
    if (!row || !row['접수번호']) continue;
    if (!allowAnyStatus && row['상태'] !== STATUS.DONE) continue;
    if (row['상태'] === STATUS.NEW || row['상태'] === '') { skipped.push(`${row['이름']}(미작성)`); continue; }
    try {
      const pdf = buildPdf_(row, ctx);
      writeCells_(sh, r, map, { 'PDF ID': pdf.getId(), 'PDF 생성시각': new Date(), '상태': STATUS.PDF });
      sh.getRange(r, map['PDF']).setFormula(`=HYPERLINK("${pdf.getUrl()}","PDF 열기")`);
      sh.getRange(r, map['상태']).clearNote();
      done++;
    } catch (err) {
      errors.push(`${row['학번']} ${row['이름']}: ${err.message}`);
    }
  }
  return { done, left: targetRows.length - processed, leftRows: targetRows.slice(processed), skipped, errors };
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

    const t = body.getParagraphs()[0];
    t.setText(`${s[KEY.COURSE]} · ${s[KEY.TITLE]}`).setHeading(DocumentApp.ParagraphHeading.TITLE);
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
    .filter(r => r['접수번호'] === row['접수번호'] || r['상태'] === STATUS.DONE || r['상태'] === STATUS.PDF)
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

function menuPublishSelected() { uiOnly_(); alertPublish_(setPublish_(selectedFeedbackRows_(), true), true); }
function menuUnpublishSelected() { uiOnly_(); alertPublish_(setPublish_(selectedFeedbackRows_(), false), false); }

/** 결과: {count, noPdf[]} — releaseDate(Date|''|undefined)를 주면 공개일도 함께 바꾼다. */
function setPublish_(rows, on, releaseDate) {
  const sh = sheet_(SHEET.FEEDBACK);
  const map = headerMap_(sh);
  const all = readRows_(sh);
  let count = 0;
  const noPdf = [];
  rows.forEach(r => {
    const row = all.find(x => x._row === r);
    if (!row) return;
    if (on && !row['PDF ID']) { noPdf.push(row['이름']); return; }
    sh.getRange(r, map['공개']).setValue(on);
    if (on && releaseDate !== undefined) sh.getRange(r, map['공개일']).setValue(releaseDate);
    count++;
  });
  return { count, noPdf };
}

function alertPublish_(r, on) {
  SpreadsheetApp.getUi().alert(
    `${r.count}행을 ${on ? '공개' : '공개 취소'}했습니다.` +
    (on ? '\n공개일이 정해져 있으면 그날부터, 비어 있으면 지금부터 학생 화면에 보입니다.' : '') +
    (r.noPdf.length ? `\n\nPDF가 없어 건너뜀: ${r.noPdf.join(', ')}` : ''));
}
