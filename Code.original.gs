// ⚠️ 백업본: 웹폼 도입 이전에 운영하던 원본 Apps Script (V4)
// 수정본(Code.gs) 적용 후 문제가 생기면 이 코드로 되돌리면 됩니다.

const FOLDER_ID = '1l4nnQoZvYU661zl24zGKmGh_y_oiLY_q';
const RESPONSE_SHEET_ID = '12783JAW4MrWAuyTH1Slm9TgLIME7vsufJIYvXrq0YQY';
const SHEET_NAME = 'Form_Responses';
const LOGO_URL = 'https://drive.google.com/uc?export=view&id=1-tI22VEIxeuACvM1YoXjTzUc7rFrNe9A';

function onFormSubmit(e) {
  const itemResponses = e.response.getItemResponses();

  const keywords = itemResponses[0].getResponse();
  const orgTypesRaw = itemResponses[1].getResponse();
  const naraSelected = itemResponses[2].getResponse();

  const naraTypesRaw = naraSelected === '예' && itemResponses[3]
    ? itemResponses[3].getResponse()
    : '';

  const infoOffset = naraSelected === '예' ? 4 : 3;
  const name = itemResponses[infoOffset].getResponse();
  const organization = itemResponses[infoOffset + 1].getResponse();
  const phone = itemResponses[infoOffset + 2].getResponse();
  const emails = itemResponses[infoOffset + 3].getResponse();
  const periodRaw = itemResponses.length > infoOffset + 4
    ? itemResponses[infoOffset + 4].getResponse()
    : '1개월';

  const months = periodRaw === '3개월' ? 3 : periodRaw === '2개월' ? 2 : 1;

  const startDate = new Date();
  const expireDate = new Date();
  expireDate.setMonth(expireDate.getMonth() + months);

  const fmt = (d) => `${d.getFullYear()}.${String(d.getMonth()+1).padStart(2,'0')}.${String(d.getDate()).padStart(2,'0')}`;
  const startDateStr = fmt(startDate);
  const expireDateStr = fmt(expireDate);

  const ss = SpreadsheetApp.openById(RESPONSE_SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 11).setValue('활성');
  sheet.getRange(lastRow, 12).setValue(startDateStr);
  sheet.getRange(lastRow, 13).setValue(expireDateStr);

  const emailList = emails.split(',').map(m => m.trim());
  const orgTypes = Array.isArray(orgTypesRaw) ? orgTypesRaw : orgTypesRaw.split(',').map(t => t.trim());
  const naraTypes = naraTypesRaw
    ? (Array.isArray(naraTypesRaw) ? naraTypesRaw : naraTypesRaw.split(',').map(t => t.trim()))
    : [];
  const typeDisplay = orgTypes.map(org => {
    if (org === '나라장터' && naraTypes.length > 0) {
      return `나라장터(${naraTypes.join(', ')})`;
    }
    return org;
  }).join(', ');

  const confirmHtml = `
    <div style="font-family:Arial,sans-serif;max-width:600px;">
      <div style="text-align:center;padding:20px 0;background:#f8f9fa;">
        <img src="${LOGO_URL}" style="height:100px;" alt="Labq">
      </div>
      <h2 style="color:#000000;text-align:center;">[Labq] 공고 맞춤 발송 신청 완료</h2>
      <p style="text-align:center;">안녕하세요, <b>${name}</b>님! 공고 맞춤 발송 신청이 완료되었습니다.</p>
      <table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;width:100%;font-size:13px;text-align:center;">
        <tr style="background:#f0f0f0;"><th>항목</th><th>내용</th></tr>
        <tr><td>키워드</td><td>${keywords || '전체'}</td></tr>
        <tr><td>공고 종류</td><td>${typeDisplay}</td></tr>
        <tr><td>수신 시작일</td><td>${startDateStr}</td></tr>
        <tr><td>수신 마감일</td><td>${expireDateStr}</td></tr>
        <tr><td>수신 이메일</td><td>${emails}</td></tr>
      </table>
      <p style="margin-top:16px;text-align:center;">신청 즉시 오늘 공고를 보내드리고, 이후 매일 오전 8시에 조건에 맞는 공고를 보내드릴게요!</p>
      <p style="color:#aaa;font-size:11px;text-align:center;">문의사항은 labq@labq.kr로 연락해주세요.</p>
    </div>`;

  emailList.forEach(email => {
    if (email) {
      GmailApp.sendEmail(email, '[Labq] 공고 맞춤 발송 신청 완료', '', { htmlBody: confirmHtml });
    }
  });

  // 신청 즉시 오늘 공고 발송
  const keywordList = keywords ? keywords.split(',').map(k => k.trim()) : [];
  const types = [...orgTypes, ...naraTypes];
  sendMail(keywordList, types, typeDisplay, emailList);

  Logger.log(`V4 등록 완료 - ${name} / ${emails} / 만료일: ${expireDateStr}`);
}

function dailySend() {
  const ss = SpreadsheetApp.openById(RESPONSE_SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[10];
    if (status !== '활성') continue;

    const expireDateStr = row[12];
    if (expireDateStr) {
      const parts = String(expireDateStr).split('.');
      const expireDate = new Date(parts[0], parts[1] - 1, parts[2]);
      expireDate.setHours(0, 0, 0, 0);
      if (today > expireDate) {
        sheet.getRange(i + 1, 11).setValue('비활성');
        Logger.log(`만료 처리: ${row[8]}`);
        continue;
      }
    }

    const keywords = row[1] ? row[1].split(',').map(k => k.trim()) : [];
    const orgTypesRaw = row[2];
    const orgTypes = Array.isArray(orgTypesRaw)
      ? orgTypesRaw
      : orgTypesRaw.split(',').map(t => t.trim());
    const naraSelected = row[3];
    const naraTypesRaw = naraSelected === '예' && row[4] ? row[4] : '';
    const naraTypes = naraTypesRaw
      ? (Array.isArray(naraTypesRaw) ? naraTypesRaw : naraTypesRaw.split(',').map(t => t.trim()))
      : [];
    const types = [...orgTypes, ...naraTypes];
    const emails = row[8].split(',').map(m => m.trim());

    const typeDisplay = orgTypes.map(org => {
      if (org === '나라장터' && naraTypes.length > 0) {
        return `나라장터(${naraTypes.join(', ')})`;
      }
      return org;
    }).join(', ');

    sendMail(keywords, types, typeDisplay, emails);
  }
}

function sendMail(keywords, types, typeDisplay, emails) {
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const dd = String(today.getDate()).padStart(2, '0');
  const todayStr = `${yyyy}.${mm}.${dd}`;

  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);
  const yesterdayStr = `${yesterday.getFullYear()}.${String(yesterday.getMonth()+1).padStart(2,'0')}.${String(yesterday.getDate()).padStart(2,'0')}`;

  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();
  let latestFile = null;
  let latestDate = new Date(0);

  while (files.hasNext()) {
    const file = files.next();
    const modified = file.getLastUpdated();
    if (modified > latestDate) {
      latestDate = modified;
      latestFile = file;
    }
  }

  if (!latestFile) {
    Logger.log('발주목록 파일 없음');
    return;
  }

  const ss = SpreadsheetApp.open(latestFile);
  const sheet = ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const richText = sheet.getDataRange().getRichTextValues();
  const totalCount = data.length - 1;

  const filteredRows = [];
  data.slice(1).forEach((row, i) => {
    const 사업명 = row[2] || '';
    const 구분   = row[1] || '';
    const 기관명 = row[4] || '';
    const keywordMatch = keywords.length === 0 || keywords[0] === ''
      || keywords.some(k => 사업명.includes(k));
    const typeMatch = types.some(t => 구분.includes(t) || 기관명.includes(t));
    if (keywordMatch && typeMatch) {
      const dateStr = row[3] ? String(row[3]) : '';
      const deadlineMatch = dateStr.match(/~\s*(\d{4}-\d{2}-\d{2})/);
      const deadline = deadlineMatch ? new Date(deadlineMatch[1]) : new Date('9999-12-31');
      filteredRows.push({ row, richIndex: i + 1, deadline });
    }
  });

  filteredRows.sort((a, b) => a.deadline - b.deadline);

  const subject = `[Labq] ${todayStr} 입찰/조달 공고 맞춤 발송`;

  let html = `
    <div style="font-family:Arial,sans-serif;max-width:900px;">
      <div style="text-align:center;padding:20px 0;background:#f8f9fa;">
        <img src="${LOGO_URL}" style="height:100px;" alt="Labq">
      </div>
      <h3>[공고 수집 기간]</h3>
      <p>${yesterdayStr} 오전 08:00 ~ ${todayStr} 오전 07:59</p>
      <h3>[요청 조건]</h3>
      <p>- 키워드: ${keywords.length === 0 || keywords[0] === '' ? '전체' : keywords.join(', ')}<br>
      - 공고 종류: ${typeDisplay}</p>
      <h3>[맞춤 공고 발송 결과]</h3>
      <p>- 전체 ${totalCount}건 중 ${filteredRows.length}건 매칭<br>`;

  if (filteredRows.length === 0) {
    html += `- <span style="background:#fff3cd; padding:2px 6px;">조건에 맞는 공고가 없습니다.</span></p>`;
  } else {
    html += `- 마감일이 3일 이내인 공고에 하이라이트</p>
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:100%;text-align:center;">
        <tr style="background:#f0f0f0"><th>No</th><th>구분</th><th>사업명</th><th>날짜</th><th>발주기관</th><th>배정예산</th><th>링크</th></tr>`;

    filteredRows.forEach(({ row, richIndex, deadline }) => {
      // ✅ F열(인덱스5) 배정예산, G열(인덱스6) URL로 변경
      const 배정예산 = row[5] || '';
      const linkUrl = richText[richIndex][6].getLinkUrl() || row[6];
      const diffDays = Math.ceil((deadline - today) / (1000 * 60 * 60 * 24));
      const isUrgent = diffDays >= 0 && diffDays <= 3;
      const rowStyle = isUrgent ? 'background:#fff3cd;' : '';
      const urgentTag = isUrgent ? ' <span style="color:red;font-weight:bold;">[마감임박]</span>' : '';

      html += `<tr style="${rowStyle}">
        <td>${row[0]}</td>
        <td>${row[1]}</td>
        <td>${row[2]}${urgentTag}</td>
        <td>${row[3]}</td>
        <td>${row[4]}</td>
        <td>${배정예산}</td>
        <td><a href="${linkUrl}">바로가기</a></td>
      </tr>`;
    });
    html += `</table>`;
  }

  html += `
      <p style="color:#aaa;font-size:11px;margin-top:16px;text-align:center;">
        수신을 원하지 않으시면 labq@labq.kr로 문의해주세요.
      </p>
    </div>`;

  emails.forEach(email => {
    if (email) {
      GmailApp.sendEmail(email, subject, '', { htmlBody: html });
    }
  });

  Logger.log('발송 완료: ' + emails.join(', '));
}
