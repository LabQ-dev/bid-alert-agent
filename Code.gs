const FOLDER_ID = '1l4nnQoZvYU661zl24zGKmGh_y_oiLY_q';
const RESPONSE_SHEET_ID = '12783JAW4MrWAuyTH1Slm9TgLIME7vsufJIYvXrq0YQY';
const SHEET_NAME = 'Form_Responses';
const LOGO_URL = 'https://drive.google.com/uc?export=view&id=1-tI22VEIxeuACvM1YoXjTzUc7rFrNe9A';

// ─────────────────────────────────────────────
// 웹 신청 접수 (웹폼 → POST)
// ─────────────────────────────────────────────
function doPost(e) {
  try {
    const p = e.parameter;

    const keywords = (p.keywords || '').trim();
    const orgTypesRaw = (p.orgTypes || '').trim();          // 콤마 구분 문자열
    const naraSelected = p.naraSelected === '예' ? '예' : '아니오';
    const naraTypesRaw = naraSelected === '예' ? (p.naraTypes || '').trim() : '';
    const budget = p.budget || '모두';
    const name = (p.name || '').trim();
    const organization = (p.organization || '').trim();
    const phone = (p.phone || '').trim();
    const emails = (p.emails || '').trim();
    const periodRaw = p.period || '1개월';

    // 필수값 검증 (구글폼 V4/V5와 동일한 필수 항목)
    if (!orgTypesRaw || !name || !organization || !phone || !emails) {
      return jsonResponse({ ok: false, error: '필수 항목(공고 종류, 이름, 소속, 전화번호, 이메일)이 누락되었습니다.' });
    }

    const months = periodRaw === '3개월' ? 3 : periodRaw === '2개월' ? 2 : 1;
    const startDate = new Date();
    const expireDate = new Date();
    expireDate.setMonth(expireDate.getMonth() + months);

    const fmt = (d) => `${d.getFullYear()}.${String(d.getMonth()+1).padStart(2,'0')}.${String(d.getDate()).padStart(2,'0')}`;
    const startDateStr = fmt(startDate);
    const expireDateStr = fmt(expireDate);

    const emailList = emails.split(',').map(m => m.trim()).filter(Boolean);

    const ss = SpreadsheetApp.openById(RESPONSE_SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    // ✅ 중복 신청 처리: 같은 이메일이 포함된 기존 '활성' 행을 비활성 처리
    deactivatePrevious(sheet, emailList);

    // 새 행 등록 (기존 구글폼 응답과 동일한 열 순서)
    sheet.appendRow([
      new Date(),        // A: 타임스탬프
      keywords,          // B: 키워드
      orgTypesRaw,       // C: 공고 종류
      naraSelected,      // D: 나라장터 선택 여부
      naraTypesRaw,      // E: 나라장터 세부 종류
      name,              // F: 이름
      organization,      // G: 기관명
      phone,             // H: 연락처
      emails,            // I: 수신 이메일
      periodRaw,         // J: 수신 기간
      '활성',            // K: 상태
      startDateStr,      // L: 수신 시작일
      expireDateStr,     // M: 수신 마감일
      budget             // N: 사업예산 규모
    ]);

    const orgTypes = orgTypesRaw.split(',').map(t => t.trim()).filter(Boolean);
    const naraTypes = naraTypesRaw ? naraTypesRaw.split(',').map(t => t.trim()).filter(Boolean) : [];
    const typeDisplay = buildTypeDisplay(orgTypes, naraTypes);

    // 확인 메일 발송
    sendConfirmMail(name, keywords, typeDisplay, startDateStr, expireDateStr, emails, emailList, budget);

    // 신청 즉시 오늘 공고 발송
    const keywordList = keywords ? keywords.split(',').map(k => k.trim()) : [];
    const types = [...orgTypes, ...naraTypes];
    sendMail(keywordList, types, typeDisplay, emailList, budget);

    Logger.log(`웹 등록 완료 - ${name} / ${emails} / 만료일: ${expireDateStr}`);
    return jsonResponse({ ok: true });

  } catch (err) {
    Logger.log('doPost 오류: ' + err);
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 같은 이메일이 하나라도 겹치는 기존 '활성' 행을 '비활성' 처리
// excludeRow: 비활성 처리에서 제외할 행 번호(1-based, 방금 추가된 행)
function deactivatePrevious(sheet, emailList, excludeRow) {
  const data = sheet.getDataRange().getValues();
  const newEmails = emailList.map(m => m.toLowerCase());
  for (let i = 1; i < data.length; i++) {
    const rowNum = i + 1;
    if (excludeRow && rowNum === excludeRow) continue;
    if (data[i][10] !== '활성') continue;
    const rowEmails = String(data[i][8] || '').split(',').map(m => m.trim().toLowerCase());
    if (rowEmails.some(m => newEmails.includes(m))) {
      sheet.getRange(rowNum, 11).setValue('비활성');
      Logger.log(`중복 신청 비활성 처리: ${rowNum}행 (${data[i][8]})`);
    }
  }
}

function buildTypeDisplay(orgTypes, naraTypes) {
  return orgTypes.map(org => {
    if (org === '나라장터' && naraTypes.length > 0) {
      return `나라장터(${naraTypes.join(', ')})`;
    }
    return org;
  }).join(', ');
}

function sendConfirmMail(name, keywords, typeDisplay, startDateStr, expireDateStr, emails, emailList, budget) {
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
        <tr><td>사업예산 규모</td><td>${budget || '모두'}</td></tr>
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
}

// ─────────────────────────────────────────────
// 기존 구글폼 접수 (전환 기간 동안 유지, 중복 처리 추가됨)
// ─────────────────────────────────────────────
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

  const emailList = emails.split(',').map(m => m.trim()).filter(Boolean);

  const ss = SpreadsheetApp.openById(RESPONSE_SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();

  // ✅ 중복 신청 처리: 방금 추가된 행(lastRow)은 제외하고 기존 활성 행 비활성 처리
  deactivatePrevious(sheet, emailList, lastRow);

  sheet.getRange(lastRow, 11).setValue('활성');
  sheet.getRange(lastRow, 12).setValue(startDateStr);
  sheet.getRange(lastRow, 13).setValue(expireDateStr);

  const orgTypes = Array.isArray(orgTypesRaw) ? orgTypesRaw : orgTypesRaw.split(',').map(t => t.trim());
  const naraTypes = naraTypesRaw
    ? (Array.isArray(naraTypesRaw) ? naraTypesRaw : naraTypesRaw.split(',').map(t => t.trim()))
    : [];
  const typeDisplay = buildTypeDisplay(orgTypes, naraTypes);

  sendConfirmMail(name, keywords, typeDisplay, startDateStr, expireDateStr, emails, emailList);

  // 신청 즉시 오늘 공고 발송
  const keywordList = keywords ? keywords.split(',').map(k => k.trim()) : [];
  const types = [...orgTypes, ...naraTypes];
  sendMail(keywordList, types, typeDisplay, emailList);

  Logger.log(`V4 등록 완료 - ${name} / ${emails} / 만료일: ${expireDateStr}`);
}

// ─────────────────────────────────────────────
// 매일 발송 (✅ 구독자별 try-catch 적용: 한 명 실패해도 나머지 계속 발송)
// ─────────────────────────────────────────────
function dailySend() {
  const ss = SpreadsheetApp.openById(RESPONSE_SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const failedList = [];   // ✅ 실패한 구독자 기록

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[10];
    if (status !== '활성') continue;

    // ✅ 구독자 한 명 처리 중 오류가 나도 전체가 중단되지 않도록 try-catch로 감쌈
    try {
      // ✅ 같은 이메일로 재신청한 경우: 가장 최근(아래쪽) 활성 행만 사용
      //    → 최종 검색어 기준으로만 발송되도록 이전 행은 건너뛰고 비활성 처리
      const myEmails = String(row[8] || '').split(',').map(m => m.trim().toLowerCase()).filter(Boolean);
      let superseded = false;
      for (let j = i + 1; j < data.length; j++) {
        if (data[j][10] !== '활성') continue;
        const laterEmails = String(data[j][8] || '').split(',').map(m => m.trim().toLowerCase());
        if (myEmails.some(m => laterEmails.includes(m))) { superseded = true; break; }
      }
      if (superseded) {
        sheet.getRange(i + 1, 11).setValue('비활성');
        Logger.log(`재신청 감지 → 이전 행 비활성 처리: ${i + 1}행 (${row[8]})`);
        continue;
      }

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
      const budget = row[13] || '모두';   // N열: 예전 신청 건은 값이 없으므로 '모두' 처리

      const typeDisplay = buildTypeDisplay(orgTypes, naraTypes);

      sendMail(keywords, types, typeDisplay, emails, budget);

      // ✅ 발송 간격: 연속 호출 부하를 줄여 INTERNAL 오류 빈도 완화
      Utilities.sleep(1000);

    } catch (e) {
      // ✅ 이 구독자만 건너뛰고 다음 구독자 계속 진행
      failedList.push(`${i + 1}행 (${row[8]})`);
      Logger.log(`발송 실패 (건너뜀): ${i + 1}행 (${row[8]}) / 오류: ${e.message}`);
    }
  }

  // ✅ 실패 건 요약 로그 (실행 기록에서 한눈에 확인용)
  if (failedList.length > 0) {
    Logger.log(`⚠️ 발송 실패 총 ${failedList.length}건: ${failedList.join(' / ')}`);
  } else {
    Logger.log('✅ 전체 구독자 발송 정상 완료');
  }
}

// 예산 규모 문자열 → {min, max}(원). '모두'거나 알 수 없으면 null(필터 없음)
// max는 미포함(미만), min은 포함(이상)
function budgetRange(budget) {
  if (budget === '5억 미만') return { min: null, max: 500000000 };
  if (budget === '20억 미만') return { min: null, max: 2000000000 };
  if (budget === '40억 미만') return { min: null, max: 4000000000 };
  if (budget === '40억 이상') return { min: 4000000000, max: null };
  return null;
}

// 배정예산 문자열 → 숫자(원). 파싱 불가 시 null
function parseBudgetAmount(raw) {
  const s = String(raw || '').trim();
  if (!s) return null;
  const eok = s.match(/([\d.,]+)\s*억/);
  if (eok) return parseFloat(eok[1].replace(/,/g, '')) * 100000000;
  const digits = s.replace(/[^\d]/g, '');
  return digits ? Number(digits) : null;
}

// ─────────────────────────────────────────────
// 공고 메일 발송 (변경 없음)
// ─────────────────────────────────────────────
function sendMail(keywords, types, typeDisplay, emails, budget) {
  const range = budgetRange(budget);
  const today = new Date();
  const todayMid = new Date(today);
  todayMid.setHours(0, 0, 0, 0);
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const dd = String(today.getDate()).padStart(2, '0');
  const todayStr = `${yyyy}.${mm}.${dd}`;

  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);
  const yesterdayStr = `${yesterday.getFullYear()}.${String(yesterday.getMonth()+1).padStart(2,'0')}.${String(yesterday.getDate()).padStart(2,'0')}`;

  // ✅ 최신 발주목록 선택: 수정 날짜가 아니라 파일명의 날짜(YYYYMMDD_발주목록) 기준
  //    (옛 파일을 누가 열어서 수정해도 최신 데이터가 잘못 밀리지 않도록)
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();
  let latestFile = null;
  let latestKey = '';

  while (files.hasNext()) {
    const file = files.next();
    const m = file.getName().match(/^(\d{8})_발주목록/);
    if (!m) continue;
    if (m[1] > latestKey) {
      latestKey = m[1];
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

  // 키워드는 대소문자 구분 없이 매칭 (예: 'ai' 로도 'AI' 공고를 찾음)
  const lowerKeywords = keywords.map(k => String(k).toLowerCase());

  const filteredRows = [];
  data.slice(1).forEach((row, i) => {
    const 사업명 = row[2] || '';
    const 구분   = row[1] || '';
    const 기관명 = row[4] || '';
    const 사업명Lower = String(사업명).toLowerCase();
    const keywordMatch = keywords.length === 0 || keywords[0] === ''
      || lowerKeywords.some(k => k !== '' && 사업명Lower.includes(k));
    const typeMatch = types.some(t => 구분.includes(t) || 기관명.includes(t));
    // 예산 필터: 배정예산을 파싱할 수 없는 공고는 제외하지 않고 포함
    const amount = parseBudgetAmount(row[5]);
    const budgetMatch = range === null || amount === null
      || ((range.max === null || amount < range.max)
        && (range.min === null || amount >= range.min));
    if (keywordMatch && typeMatch && budgetMatch) {
      const dateStr = row[3] ? String(row[3]) : '';
      const deadlineMatch = dateStr.match(/~\s*(\d{4}-\d{2}-\d{2})/);
      const deadline = deadlineMatch ? new Date(deadlineMatch[1]) : new Date('9999-12-31');
      // ✅ 발송일 기준 제안마감일이 이미 지난 공고는 제외 (마감일 파싱 불가 시 포함)
      if (deadline < todayMid) return;
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
      - 공고 종류: ${typeDisplay}<br>
      - 사업예산 규모: ${budget || '모두'}</p>
      <h3>[맞춤 공고 발송 결과]</h3>
      <p>- 전체 ${totalCount}건 중 ${filteredRows.length}건 매칭<br>`;

  if (filteredRows.length === 0) {
    html += `- <span style="background:#fff3cd; padding:2px 6px;">조건에 맞는 공고가 없습니다.</span></p>`;
  } else {
    html += `- 마감일이 3일 이내인 공고에 하이라이트</p>
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:100%;text-align:center;">
        <tr style="background:#f0f0f0"><th>No</th><th>구분</th><th>사업명</th><th>날짜</th><th>발주기관</th><th>배정예산</th><th>링크</th></tr>`;

    filteredRows.forEach(({ row, richIndex, deadline }) => {
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
