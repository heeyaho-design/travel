function collectAllFestivals() {
  const baseUrl =
    "https://www.mcst.go.kr/site/s_culture/festival/festivalList.jsp";

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("festa");

  if (!sheet) {
    Logger.log("festa 시트를 찾을 수 없습니다.");
    return;
  }

  // HTML 가져오기
  function fetchHtml(page) {
    const url = baseUrl + "?pCurrentPage=" + page;

    const response = UrlFetchApp.fetch(url, {
      headers: { "User-Agent": "Mozilla/5.0" },
    });

    return response.getContentText();
  }

  // HTML 태그와 특수문자 정리
  function cleanText(text) {
    return text
      .replace(/<[^>]*>/g, "")
      .replace(/&nbsp;/g, " ")
      .replace(/&amp;/g, "&")
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/\s+/g, " ")
      .trim();
  }

  // 기간, 장소, 문의 추출
  function extractField(html, label) {
    const regex = new RegExp(
      "<strong>\\s*" + label + "\\s*</strong>\\s*:\\s*([^<]*)",
    );

    const match = html.match(regex);
    return match ? cleanText(match[1]) : "";
  }

  Logger.log("===== 전국 축제 수집 시작 =====");

  // 마지막 페이지 확인
  const firstHtml = fetchHtml(1);

  const pageNumbers = [...firstHtml.matchAll(/pCurrentPage=(\d+)/g)].map(
    (match) => Number(match[1]),
  );

  const lastPage = pageNumbers.length > 0 ? Math.max(...pageNumbers) : 1;

  Logger.log("확인된 마지막 페이지: " + lastPage);

  const allFestivals = [];

  // 모든 페이지 수집
  for (let page = 1; page <= lastPage; page++) {
    Logger.log("===== " + page + "페이지 수집 =====");

    const html = page === 1 ? firstHtml : fetchHtml(page);

    const cardRegex = /<a href="festivalView\.jsp[^"]*"[\s\S]*?<\/a>/g;
    const cards = html.match(cardRegex) || [];

    let pageCount = 0;

    cards.forEach((card) => {
      const nameMatch = card.match(/class="sub-tit">([\s\S]*?)<\/p>/);

      if (!nameMatch) return;

      const descriptionMatch = card.match(/class="sub-text">([\s\S]*?)<\/div>/);

      const imageMatch = card.match(/<img[^>]+src=["']([^"']+)["']/);

      const name = cleanText(nameMatch[1]);

      const description = descriptionMatch
        ? cleanText(descriptionMatch[1])
        : "";

      const period = extractField(card, "기간");
      const location = extractField(card, "장소");
      const contact = extractField(card, "문의");

      let imageUrl = imageMatch ? imageMatch[1] : "";

      if (imageUrl && imageUrl.startsWith("/")) {
        imageUrl = "https://www.mcst.go.kr" + imageUrl;
      }

      allFestivals.push([
        name,
        period,
        location,
        contact,
        description,
        imageUrl,
      ]);

      pageCount++;
    });

    Logger.log(page + "페이지 수집 결과: " + pageCount + "개");

    Utilities.sleep(200);
  }

  // 제목 행 설정
  sheet
    .getRange(1, 1, 1, 6)
    .setValues([["축제명", "기간", "장소", "문의", "설명", "이미지 URL"]]);

  // 기존 데이터 삭제
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 6).clearContent();
  }

  // 새 데이터 저장
  if (allFestivals.length > 0) {
    sheet.getRange(2, 1, allFestivals.length, 6).setValues(allFestivals);
  }

  Logger.log("------------------------------");
  Logger.log("전체 수집 축제 수: " + allFestivals.length);
  Logger.log("===== 축제 수집 완료 =====");
}

// 기본 화면은 index.html로 제공하고, JSON 요청은 기존 데이터 형식으로 반환
function doGet(e) {
  const callback = e && e.parameter ? e.parameter.callback : "";
  const isJsonp = callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback);
  const isJsonRequest = e && e.parameter && e.parameter.format === "json";

  if (isJsonp || isJsonRequest) {
    const payload = getFestivalData();

    if (isJsonp) {
      return ContentService.createTextOutput(
        callback + "(" + JSON.stringify(payload) + ");",
      ).setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(
      ContentService.MimeType.JSON,
    );
  }

  return HtmlService.createHtmlOutputFromFile("index").setTitle(
    "전국 축제 소식",
  );
}

// index.html에서 호출하는 축제 데이터 조회 함수
function getFestivalData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("festa");

  if (!sheet) {
    return {
      success: false,
      message: "festa 시트를 찾을 수 없습니다.",
      festivals: [],
    };
  }

  const lastRow = sheet.getLastRow();
  const festivals = [];

  if (lastRow > 1) {
    const rows = sheet.getRange(2, 1, lastRow - 1, 6).getDisplayValues();

    rows.forEach((row) => {
      if (!row[0]) return;

      festivals.push({
        name: row[0],
        period: row[1],
        location: row[2],
        contact: row[3],
        description: row[4],
        imageUrl: row[5],
      });
    });
  }

  return {
    success: true,
    festivals: festivals,
    updatedAt: new Date().toISOString(),
  };
}

// users 시트의 이메일 주소마다 festa 시트의 축제 뉴스를 개별 발송
function sendFestivalNewsEmails() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = spreadsheet.getSheetByName("users");
  const festaSheet = spreadsheet.getSheetByName("festa");

  if (!usersSheet) {
    throw new Error("users 시트를 찾을 수 없습니다.");
  }

  if (!festaSheet) {
    throw new Error("festa 시트를 찾을 수 없습니다.");
  }

  const userLastColumn = usersSheet.getLastColumn();
  if (userLastColumn === 0) {
    throw new Error("users 시트에 이메일 항목이 없습니다.");
  }

  const userHeaders = usersSheet
    .getRange(1, 1, 1, userLastColumn)
    .getDisplayValues()[0]
    .map((header) => header.trim());
  const emailColumn = userHeaders.indexOf("이메일");

  if (emailColumn === -1) {
    throw new Error('users 시트에서 "이메일" 열을 찾을 수 없습니다.');
  }

  const userLastRow = usersSheet.getLastRow();
  if (userLastRow < 2) {
    throw new Error("users 시트에 발송 대상 이메일 주소가 없습니다.");
  }

  const emailValues = usersSheet
    .getRange(2, emailColumn + 1, userLastRow - 1, 1)
    .getDisplayValues();
  const recipients = [];
  const seenEmails = {};
  const invalidRows = [];

  emailValues.forEach((row, index) => {
    const email = row[0].trim();
    if (!email) return;

    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      invalidRows.push(index + 2);
      return;
    }

    const normalizedEmail = email.toLowerCase();
    if (!seenEmails[normalizedEmail]) {
      seenEmails[normalizedEmail] = true;
      recipients.push(email);
    }
  });

  if (invalidRows.length > 0) {
    throw new Error(
      "users 시트의 이메일 주소 형식을 확인해 주세요. 행: " +
        invalidRows.join(", "),
    );
  }

  if (recipients.length === 0) {
    throw new Error("발송할 이메일 주소가 없습니다.");
  }

  const festaLastRow = festaSheet.getLastRow();
  if (festaLastRow < 2) {
    throw new Error("festa 시트에 발송할 뉴스 데이터가 없습니다.");
  }

  const festivals = festaSheet
    .getRange(2, 1, festaLastRow - 1, 6)
    .getDisplayValues()
    .filter((row) => row[0].trim());

  if (festivals.length === 0) {
    throw new Error("festa 시트에 발송할 뉴스 데이터가 없습니다.");
  }

  const escapeHtml = (value) =>
    value
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");

  const plainBody = festivals
    .map(
      (row) =>
        row[0] +
        "\n기간: " +
        row[1] +
        "\n장소: " +
        row[2] +
        "\n문의: " +
        row[3] +
        "\n설명: " +
        row[4] +
        "\n이미지 URL: " +
        row[5],
    )
    .join("\n\n--------------------\n\n");

  const htmlCards = festivals
    .map((row) => {
      const details = [
        ["장소", row[2]],
        ["문의", row[3]],
      ]
        .filter((item) => item[1])
        .map(
          (item) =>
            `<tr><td style="padding:5px 12px 5px 0;color:#64748b;font-size:13px;white-space:nowrap;">${item[0]}</td>` +
            `<td style="padding:5px 0;color:#334155;font-size:13px;">${escapeHtml(item[1])}</td></tr>`,
        )
        .join("");
      const imageLink = /^https?:\/\//i.test(row[5])
        ? `<a href="${escapeHtml(row[5])}" style="display:inline-block;margin-top:16px;padding:9px 14px;border:1px solid #cbd5e1;border-radius:6px;color:#334155;font-size:13px;text-decoration:none;">축제 이미지 보기</a>`
        : "";

      return (
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 16px;border:1px solid #e2e8f0;border-radius:10px;background-color:#ffffff;">' +
        '<tr><td style="padding:22px 24px;">' +
        '<div style="margin-bottom:10px;color:#2563eb;font-size:12px;font-weight:bold;letter-spacing:1px;">FESTIVAL NEWS</div>' +
        `<h2 style="margin:0 0 12px;color:#0f172a;font-size:21px;line-height:1.4;">${escapeHtml(row[0])}</h2>` +
        `<div style="display:inline-block;padding:6px 10px;border-radius:20px;background-color:#eff6ff;color:#1d4ed8;font-size:13px;">기간　${escapeHtml(row[1] || "미정")}</div>` +
        (details
          ? `<table role="presentation" cellpadding="0" cellspacing="0" style="margin-top:14px;">${details}</table>`
          : "") +
        (row[4]
          ? `<p style="margin:16px 0 0;color:#475569;font-size:14px;line-height:1.7;white-space:pre-line;">${escapeHtml(row[4])}</p>`
          : "") +
        imageLink +
        "</td></tr></table>"
      );
    })
    .join("");

  const htmlBody =
    '<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">' +
    "<style>" +
    "body{margin:0;padding:0;background-color:#f1f5f9;font-family:Arial,'Malgun Gothic',sans-serif;}" +
    "@media only screen and (max-width:600px){.email-content{padding:20px 14px!important;}.email-title{font-size:23px!important;}}" +
    "</style></head>" +
    "<body style=\"margin:0;padding:0;background-color:#f1f5f9;font-family:Arial,'Malgun Gothic',sans-serif;\">" +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#f1f5f9;">' +
    '<tr><td align="center" style="padding:28px 12px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:680px;">' +
    '<tr><td style="padding:30px 28px;border-radius:12px 12px 0 0;background-color:#1e3a8a;">' +
    '<div style="margin-bottom:8px;color:#bfdbfe;font-size:12px;font-weight:bold;letter-spacing:2px;">LOCAL CULTURE</div>' +
    '<h1 class="email-title" style="margin:0;color:#ffffff;font-size:28px;line-height:1.3;">이번 축제 소식</h1>' +
    `<p style="margin:10px 0 0;color:#dbeafe;font-size:14px;">새로운 축제 ${festivals.length}건을 소개합니다.</p>` +
    '</td></tr><tr><td class="email-content" style="padding:24px 20px 12px;background-color:#f8fafc;">' +
    htmlCards +
    "</td></tr>" +
    '<tr><td style="padding:18px 24px 26px;border-radius:0 0 12px 12px;background-color:#f8fafc;color:#94a3b8;font-size:12px;line-height:1.6;text-align:center;">' +
    "users 시트에 등록된 이메일 주소로 발송된 축제 소식입니다." +
    "</td></tr></table>" +
    "</td></tr></table></body></html>";

  const remainingQuota = MailApp.getRemainingDailyQuota();
  if (remainingQuota < recipients.length) {
    throw new Error(
      "일일 이메일 발송 한도가 부족합니다. 남은 수신자 한도: " +
        remainingQuota +
        ", 발송 대상: " +
        recipients.length,
    );
  }

  recipients.forEach((recipient) => {
    MailApp.sendEmail({
      to: recipient,
      subject: "축제 소식",
      body: plainBody,
      htmlBody: htmlBody,
    });
  });

  Logger.log(
    "축제 뉴스 발송 완료: " +
      recipients.length +
      "명, 뉴스 " +
      festivals.length +
      "건",
  );
}

// 이메일 구독 신청 정보를 users 시트에 저장
function subscribeUser(name, email) {
  name = String(name || "").trim();
  email = String(email || "")
    .trim()
    .toLowerCase();

  // 입력값 확인
  if (!name || !email) {
    return {
      success: false,
      message: "이름과 이메일을 모두 입력해 주세요.",
    };
  }

  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

  if (!emailPattern.test(email)) {
    return {
      success: false,
      message: "올바른 이메일 주소를 입력해 주세요.",
    };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("users");

  if (!sheet) {
    return {
      success: false,
      message: "users 시트를 찾을 수 없습니다.",
    };
  }

  // 시트가 비어 있다면 제목 행 생성
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["이름", "이메일", "구독여부"]);
  }

  const lastRow = sheet.getLastRow();

  // 기존 이메일이 있는지 확인
  if (lastRow >= 2) {
    const users = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

    for (let i = 0; i < users.length; i++) {
      const savedEmail = String(users[i][1] || "")
        .trim()
        .toLowerCase();

      if (savedEmail === email) {
        // 이미 등록된 이메일이면 구독 상태 갱신
        sheet.getRange(i + 2, 1, 1, 3).setValues([[name, email, "구독중"]]);

        return {
          success: true,
          message: "이미 등록된 이메일입니다. 구독 상태를 갱신했습니다.",
        };
      }
    }
  }

  // 신규 구독자 추가
  sheet.appendRow([name, email, "구독중"]);

  return {
    success: true,
    message: "구독 신청이 완료되었습니다!",
  };
}
