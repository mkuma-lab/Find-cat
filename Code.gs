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
    throw new Error("[GEMINI_SEND_FAILED] 画像データがありません。");
  }
  if (typeof request.mimeType !== "string" ||
      !/^image\/(jpeg|png|webp|heic|heif)$/i.test(request.mimeType)) {
    throw new Error("[GEMINI_SEND_FAILED] JPEG、PNG、WebP、HEIC/HEIF形式の画像を選択してください。");
  }

  const apiKey = PropertiesService.getScriptProperties()
    .getProperty(GEMINI_API_KEY_PROPERTY);
  if (!apiKey) {
    throw new Error("[GEMINI_SEND_FAILED] Gemini APIキーが未設定です。先にAPIキーを保存してください。");
  }

  const prompt = [
    "画像内に記載された情報を読み取り、項目名と値の組にしてください。",
    "見出しとその内容の対応を保ち、画像にない情報を補わないでください。",
    "値が数値ならJSONの数値、文字列ならJSON文字列にしてください。",
    "読み取れる項目がない場合は items を空配列にしてください。",
    '次のJSONだけを返してください: {"items":[{"name":"項目名","value":"値または数値"}]}'
  ].join("\n");

  let response;
  try {
    response = UrlFetchApp.fetch(
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
  } catch (error) {
    throw new Error("[GEMINI_SEND_FAILED] Gemini APIへの送信に失敗しました: " + error.message);
  }

  const statusCode = response.getResponseCode();
  const responseBody = response.getContentText();
  if (statusCode < 200 || statusCode >= 300) {
    throw new Error("[GEMINI_RESULT_FAILED] Gemini APIからエラー応答を受信しました (HTTP " +
      statusCode + "): " + getGeminiErrorMessage(responseBody));
  }

  let apiResult;
  try {
    apiResult = JSON.parse(responseBody);
  } catch (error) {
    throw new Error("[GEMINI_RESULT_FAILED] Gemini APIの応答をJSONとして読み取れませんでした。");
  }

  const text = apiResult.candidates &&
    apiResult.candidates[0] &&
    apiResult.candidates[0].content &&
    apiResult.candidates[0].content.parts &&
    apiResult.candidates[0].content.parts[0] &&
    apiResult.candidates[0].content.parts[0].text;
  if (typeof text !== "string") {
    throw new Error("[GEMINI_RESULT_FAILED] Gemini APIから認識結果が返されませんでした。");
  }

  let result;
  try {
    result = JSON.parse(text);
  } catch (error) {
    throw new Error("[GEMINI_RESULT_FAILED] Geminiの認識結果をJSONとして読み取れませんでした。");
  }
  if (!result || !Array.isArray(result.items)) {
    throw new Error("[GEMINI_RESULT_FAILED] Geminiの認識結果に items 配列がありません。");
  }

  try {
    const items = result.items.map(function(item, index) {
      if (!item || typeof item.name !== "string" || !item.name.trim() ||
          (typeof item.value !== "string" && typeof item.value !== "number") ||
          (typeof item.value === "number" && !Number.isFinite(item.value))) {
        throw new Error((index + 1) + "件目の項目名または値が正しくありません。");
      }
      return { name: item.name.trim(), value: item.value };
    });
    return { items: items };
  } catch (error) {
    throw new Error("[GEMINI_RESULT_FAILED] Geminiの認識結果を取得できませんでした: " + error.message);
  }
}

function getGeminiErrorMessage(responseBody) {
  try {
    const result = JSON.parse(responseBody);
    if (result.error && result.error.message) {
      return result.error.message;
    }
  } catch (error) {
    return "詳細なエラー情報を読み取れませんでした。";
  }
  return "詳細なエラー情報がありません。";
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
  let lockAcquired = false;
  let spreadsheetAction = "";
  try {
    lock.waitLock(30000);
    lockAcquired = true;

    const spreadsheetResult = findOrCreateSpreadsheet(name);
    const spreadsheet = spreadsheetResult.spreadsheet;
    spreadsheetAction = spreadsheetResult.action;
    let sheet;
    try {
      sheet = spreadsheet.getSheets()[0];
      if (!sheet) {
        throw new Error("スプレッドシートにシートがありません。");
      }

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
        spreadsheetAction: spreadsheetAction,
        savedCount: items.length
      };
    } catch (error) {
      throw new Error("[SPREADSHEET_APPEND_FAILED:" + spreadsheetAction + "] " +
        "スプレッドシートは" + (spreadsheetAction === "created" ? "新規作成" : "オープン") +
        "できましたが、データの保存に失敗しました: " + error.message);
    }
  } catch (error) {
    if (/^\[(SPREADSHEET_(SEARCH|OPEN|CREATE|APPEND|OPERATION)_FAILED)/.test(error.message)) {
      throw error;
    }
    throw new Error("[SPREADSHEET_OPERATION_FAILED] " + error.message);
  } finally {
    if (lockAcquired) {
      lock.releaseLock();
    }
  }
}

function findOrCreateSpreadsheet(name) {
  const matches = [];
  try {
    const files = DriveApp.getFilesByName(name);
    while (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
        matches.push(file);
      }
    }
  } catch (error) {
    throw new Error("[SPREADSHEET_SEARCH_FAILED] Google Drive内の検索に失敗しました: " +
      error.message);
  }

  if (matches.length > 1) {
    throw new Error("[SPREADSHEET_SEARCH_FAILED] 同じ名前のスプレッドシートが複数あります。ファイル名を一意にしてください。");
  }
  if (matches.length === 1) {
    try {
      return {
        spreadsheet: SpreadsheetApp.openById(matches[0].getId()),
        action: "opened"
      };
    } catch (error) {
      throw new Error("[SPREADSHEET_OPEN_FAILED] 既存スプレッドシートを開けませんでした: " +
        error.message);
    }
  }
  try {
    return {
      spreadsheet: SpreadsheetApp.create(name),
      action: "created"
    };
  } catch (error) {
    throw new Error("[SPREADSHEET_CREATE_FAILED] 新規スプレッドシートを作成できませんでした: " +
      error.message);
  }
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
