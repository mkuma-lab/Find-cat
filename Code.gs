const GEMINI_MODEL = "gemini-2.5-flash";
const GEMINI_API_KEY_PROPERTY = "GEMINI_API_KEY";
const SHEET_HEADERS = ["保存日時", "撮影日時", "緯度", "経度", "項目名", "値"];

function doGet() {
  return HtmlService.createHtmlOutputFromFile("ocr_test")
    .setTitle("画像OCR・スプレッドシート保存");
}

function checkGeminiApiKey() {
  return Boolean(
    PropertiesService.getScriptProperties().getProperty(GEMINI_API_KEY_PROPERTY)
  );
}

function saveGeminiApiKey(apiKey) {
  if (typeof apiKey !== "string" || !apiKey.trim()) {
    throw new Error("Gemini APIキーを入力してください。");
  }

  PropertiesService.getScriptProperties()
    .setProperty(GEMINI_API_KEY_PROPERTY, apiKey.trim());
  return true;
}

function recognizeImage(request) {
  if (!request || typeof request.imageData !== "string" || !request.imageData) {
    throw new Error("画像データがありません。");
  }
  if (typeof request.mimeType !== "string" ||
      !/^image\/(jpeg|png|webp|heic|heif)$/i.test(request.mimeType)) {
    throw new Error("JPEG、PNG、WebP、HEIC/HEIF形式の画像を選択してください。");
  }

  const apiKey = PropertiesService.getScriptProperties()
    .getProperty(GEMINI_API_KEY_PROPERTY);
  if (!apiKey) {
    throw new Error("Gemini APIキーが未設定です。先にAPIキーを保存してください。");
  }

  const prompt = [
    "画像内に記載された情報を読み取り、項目名と値の組にしてください。",
    "見出しとその内容の対応を保ち、画像にない情報を補わないでください。",
    "値が数値ならJSONの数値、文字列ならJSON文字列にしてください。",
    "読み取れる項目がない場合は items を空配列にしてください。",
    '次のJSONだけを返してください: {"items":[{"name":"項目名","value":"値または数値"}]}'
  ].join("\n");

  const response = UrlFetchApp.fetch(
    "https://generativelanguage.googleapis.com/v1beta/models/" +
      GEMINI_MODEL + ":generateContent?key=" + encodeURIComponent(apiKey),
    {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        contents: [{
          parts: [
            { text: prompt },
            { inline_data: { mime_type: request.mimeType, data: request.imageData } }
          ]
        }],
        generationConfig: {
          temperature: 0,
          responseMimeType: "application/json"
        }
      }),
      muteHttpExceptions: true
    }
  );

  const statusCode = response.getResponseCode();
  const responseBody = response.getContentText();
  if (statusCode < 200 || statusCode >= 300) {
    throw new Error("Gemini APIエラー (" + statusCode + "): " + responseBody);
  }

  let apiResult;
  try {
    apiResult = JSON.parse(responseBody);
  } catch (error) {
    throw new Error("Gemini APIの応答をJSONとして読み取れませんでした。");
  }

  const text = apiResult.candidates &&
    apiResult.candidates[0] &&
    apiResult.candidates[0].content &&
    apiResult.candidates[0].content.parts &&
    apiResult.candidates[0].content.parts[0] &&
    apiResult.candidates[0].content.parts[0].text;
  if (typeof text !== "string") {
    throw new Error("Gemini APIから認識結果が返されませんでした。");
  }

  let result;
  try {
    result = JSON.parse(text);
  } catch (error) {
    throw new Error("Geminiの認識結果をJSONとして読み取れませんでした。");
  }
  if (!result || !Array.isArray(result.items)) {
    throw new Error("Geminiの認識結果に items 配列がありません。");
  }

  return result.items.map(function(item, index) {
    if (!item || typeof item.name !== "string" || !item.name.trim() ||
        (typeof item.value !== "string" && typeof item.value !== "number") ||
        (typeof item.value === "number" && !Number.isFinite(item.value))) {
      throw new Error((index + 1) + "件目の項目名または値が正しくありません。");
    }
    return { name: item.name.trim(), value: item.value };
  });
}

function saveResults(request) {
  if (!request || typeof request.spreadsheetName !== "string" ||
      !request.spreadsheetName.trim()) {
    throw new Error("保存ファイル名を入力してください。");
  }
  if (!Array.isArray(request.items)) {
    throw new Error("保存する認識結果がありません。");
  }

  const name = request.spreadsheetName.trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = findOrCreateSpreadsheet(name);
    const sheet = spreadsheet.getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, SHEET_HEADERS.length).setValues([SHEET_HEADERS]);
    }

    const savedAt = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd HH:mm:ss"
    );
    const latitude = validCoordinate(request.latitude, -90, 90);
    const longitude = validCoordinate(request.longitude, -180, 180);
    const captureDate = typeof request.captureDate === "string"
      ? request.captureDate.slice(0, 50)
      : "";
    const items = request.items.map(function(item, index) {
      if (!item || typeof item.name !== "string" || !item.name.trim() ||
          (typeof item.value !== "string" && typeof item.value !== "number") ||
          (typeof item.value === "number" && !Number.isFinite(item.value))) {
        throw new Error((index + 1) + "件目の保存データが正しくありません。");
      }
      return [
        savedAt,
        captureDate,
        latitude,
        longitude,
        safeSheetText(item.name.trim()),
        typeof item.value === "string" ? safeSheetText(item.value) : item.value
      ];
    });

    const rowsToSave = items.length
      ? items
      : [[savedAt, captureDate, latitude, longitude, "", ""]];
    const firstRow = sheet.getLastRow() + 1;
    sheet.getRange(firstRow, 1, rowsToSave.length, SHEET_HEADERS.length)
      .setValues(rowsToSave);
    SpreadsheetApp.flush();

    return {
      spreadsheetName: spreadsheet.getName(),
      spreadsheetUrl: spreadsheet.getUrl(),
      savedCount: items.length
    };
  } finally {
    lock.releaseLock();
  }
}

function findOrCreateSpreadsheet(name) {
  const files = DriveApp.getFilesByName(name);
  const matches = [];
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      matches.push(file);
    }
  }

  if (matches.length > 1) {
    throw new Error("同じ名前のスプレッドシートが複数あります。ファイル名を一意にしてください。");
  }
  if (matches.length === 1) {
    return SpreadsheetApp.openById(matches[0].getId());
  }
  return SpreadsheetApp.create(name);
}

function validCoordinate(value, minimum, maximum) {
  if (value === null || value === undefined || value === "") {
    return "";
  }
  const coordinate = Number(value);
  if (!Number.isFinite(coordinate) || coordinate < minimum || coordinate > maximum) {
    return "";
  }
  return coordinate;
}

function safeSheetText(value) {
  return /^[=+\-@]/.test(value) ? "'" + value : value;
}
