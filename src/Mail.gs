/**
 * 메일: 제출 완료 알림, 피드백 도착 알림
 */

function fillTemplate_(tpl, vars) {
  return String(tpl || '').replace(/\{([^}]+)\}/g, (m, k) => (k in vars ? vars[k] : m));
}

function mailFrame_(title, inner) {
  return `
  <div style="font-family:'Noto Sans KR',Apple SD Gothic Neo,sans-serif;background:#FFF6CC;padding:24px">
    <div style="max-width:560px;margin:auto;background:#fff;border:3px solid #222;border-radius:20px;box-shadow:6px 6px 0 #2F5DA8;padding:24px">
      <div style="font-size:13px;color:#2F5DA8;font-weight:700">🍌 과제 피드백 도우미</div>
      <h2 style="margin:8px 0 16px;font-size:20px;color:#222">${title}</h2>
      ${inner}
      <p style="margin-top:24px;font-size:12px;color:#777">이 메일은 자동으로 보냈습니다. © shiny-peace</p>
    </div>
  </div>`;
}

function sendSubmitMail_(o) {
  const s = getSettings_();
  const vars = {
    이름: o.name, 단계명: o.stage.name, 과제명: s[KEY.TITLE] || '', 과목: s[KEY.COURSE] || '',
    접수번호: o.receipt, 시각: fmtTime_(o.now),
  };
  const url = webAppUrl_();
  const rows = [
    ['접수 번호', o.receipt],
    ['제출 시각', fmtTime_(o.now)],
    ['단계', `${o.stage.name}${o.attempt > 1 ? ` (${o.attempt}번째 제출)` : ''}`],
    ['파일', o.fileName || '(파일 없음)'],
    ['이번 단계에서 한 일', o.did],
  ].map(([k, v]) => `<tr><td style="padding:6px 10px;background:#FFD93B;font-weight:700;white-space:nowrap;border:1px solid #222">${escHtml_(k)}</td><td style="padding:6px 10px;border:1px solid #222">${escHtml_(v).replace(/\n/g, '<br>')}</td></tr>`).join('');

  const inner = `
    <p style="font-size:15px;line-height:1.7">${escHtml_(o.name)} 학생, <b>${escHtml_(o.stage.name)}</b> 단계 제출이 완료되었습니다.<br>선생님이 읽어 보고 피드백을 준비할게요.</p>
    <table style="border-collapse:collapse;width:100%;font-size:14px">${rows}</table>
    ${url ? `<p style="margin-top:20px"><a href="${escHtml_(url)}" style="display:inline-block;background:#2F5DA8;color:#fff;text-decoration:none;padding:10px 20px;border-radius:999px;font-weight:700">제출 기록 확인하기</a></p>` : ''}`;

  MailApp.sendEmail({
    to: o.email,
    subject: fillTemplate_(s[KEY.SUBMIT_MAIL_SUBJECT] || '[{과제명}] {단계명} 제출 완료', vars),
    htmlBody: mailFrame_('제출이 완료되었습니다 ✅', inner),
    name: s[KEY.TEACHER] ? `${s[KEY.TEACHER]} 선생님` : '과제 피드백 도우미',
  });
}

/** 공개되었고 공개일이 지난 피드백 가운데 아직 알리지 않은 학생에게 메일을 보낸다. */
function sendArrivalMails_() {
  const s = getSettings_();
  if (!yes_(s[KEY.ARRIVE_MAIL])) return { sent: 0, skipped: 0, reason: '설정에서 꺼져 있음' };
  const fbSh = sheet_(SHEET.FEEDBACK);
  const fbMap = headerMap_(fbSh);
  const subs = readRows_(sheet_(SHEET.SUBMIT));
  const emailByReceipt = {};
  subs.forEach(r => { emailByReceipt[r['접수번호']] = String(r['이메일'] || '').trim(); });
  const rosterEmail = {};
  getRoster_().forEach(r => { rosterEmail[r.id + '|' + r.name] = r.email; });

  const url = webAppUrl_();
  const now = new Date();
  let sent = 0, skipped = 0;
  const quota = MailApp.getRemainingDailyQuota();

  readRows_(fbSh).forEach(fb => {
    if (!isReleased_(fb, now) || fb['도착메일']) return;
    const email = rosterEmail[normId_(fb['학번']) + '|' + normName_(fb['이름'])] || emailByReceipt[fb['접수번호']];
    if (!isEmail_(email) || sent >= quota - 1) { skipped++; return; }
    const vars = { 이름: fb['이름'], 단계명: fb['단계명'], 과제명: s[KEY.TITLE] || '', 과목: s[KEY.COURSE] || '', 접수번호: fb['접수번호'] };
    const inner = `
      <p style="font-size:15px;line-height:1.7">${escHtml_(fb['이름'])} 학생, <b>${escHtml_(fb['단계명'])}</b> 단계 피드백이 도착했습니다.</p>
      ${fb['한 줄 총평'] ? `<p style="background:#FFF6CC;border-left:6px solid #FFD93B;padding:10px 14px">${escHtml_(fb['한 줄 총평'])}</p>` : ''}
      <p style="font-size:14px">웹앱에서 학번과 이름을 넣으면 피드백 문서(PDF)를 볼 수 있습니다. 다음 단계까지 할 일을 꼭 확인하세요.</p>
      ${url ? `<p><a href="${escHtml_(url)}" style="display:inline-block;background:#2F5DA8;color:#fff;text-decoration:none;padding:10px 20px;border-radius:999px;font-weight:700">내 피드백 보러 가기</a></p>` : ''}`;
    try {
      MailApp.sendEmail({
        to: email,
        subject: fillTemplate_(s[KEY.ARRIVE_MAIL_SUBJECT] || '[{과제명}] {단계명} 피드백 도착', vars),
        htmlBody: mailFrame_('피드백이 도착했습니다 📬', inner),
        name: s[KEY.TEACHER] ? `${s[KEY.TEACHER]} 선생님` : '과제 피드백 도우미',
      });
      fbSh.getRange(fb._row, fbMap['도착메일']).setValue(new Date());
      sent++;
    } catch (err) {
      console.error('도착 메일 실패', fb['접수번호'], err);
      skipped++;
    }
  });
  return { sent, skipped };
}

function menuSendArrivalMails() {
  uiOnly_();
  const r = sendArrivalMails_();
  SpreadsheetApp.getUi().alert(r.reason ? `보내지 않았습니다: ${r.reason}` : `도착 메일 ${r.sent}통을 보냈습니다.${r.skipped ? ` (이메일이 없거나 실패: ${r.skipped}명)` : ''}`);
}
