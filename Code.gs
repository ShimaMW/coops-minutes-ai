// =============================================
// COOPs 担当者会議議事録AI — Code.gs (Web API / Vercel対応版)
// =============================================

function getGeminiApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (key && key.trim()) return key.trim();
  throw new Error('スクリプトプロパティに GEMINI_API_KEY が設定されていません。プロジェクト設定から新しいGemini APIキーを設定してください。');
}

function doGet(e) {
  return HtmlService.createHtmlOutput('COOPs 担当者会議議事録AI API Endpoint is Active.');
}

function doPost(e) {
  let responseData = {};
  try {
    const postData = JSON.parse(e.postData.contents || '{}');
    const action = postData.action;
    const payload = postData.payload || {};

    switch(action) {
      case 'generateMinutes':
        responseData = generateMinutes(payload.inputText, payload.audioData);
        break;
      case 'saveLogAndFinish':
        responseData = saveLogAndFinish(
          payload.meetingDate, payload.clientName, payload.inputText,
          payload.summary, payload.item1, payload.item2, payload.item3, payload.item4
        );
        break;
      case 'getHistorySummaryList':
        responseData = getHistorySummaryList();
        break;
      case 'getLogDetail':
        responseData = getLogDetail(payload.id);
        break;
      default:
        responseData = { error: 'Unknown action: ' + action };
    }
  } catch(err) {
    responseData = { error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(responseData))
    .setMimeType(ContentService.MimeType.JSON);
}

function generateMinutes(inputText, audioData) {
  const apiKey = getGeminiApiKey();
  let fileUri = null;
  let fileNameToCleanUp = null;

  if (audioData && audioData.base64) {
    const uploadUrl = 'https://generativelanguage.googleapis.com/upload/v1beta/files?key=' + apiKey;
    const blob = Utilities.newBlob(Utilities.base64Decode(audioData.base64), audioData.mimeType, "meeting_audio");

    const uploadOptions = {
      'method': 'post',
      'contentType': audioData.mimeType,
      'payload': blob,
      'headers': { "X-Goog-Upload-Protocol": "raw" },
      'muteHttpExceptions': true
    };

    try {
      const uploadRes = UrlFetchApp.fetch(uploadUrl, uploadOptions);
      const uploadJson = JSON.parse(uploadRes.getContentText());
      if (uploadJson.error) return { error: "【AIエラー】: " + uploadJson.error.message };
      fileUri = uploadJson.file.uri;
      fileNameToCleanUp = uploadJson.file.name;
    } catch (e) {
      return { error: "【通信エラー】: 音声のアップロードに失敗しました。" + e.message };
    }
  }

  const generateUrl = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;

  const prompt = `
あなたは客観的で正確な「議事録作成のプロフェッショナル」です。
以下の【入力データ】の内容から、サービス担当者会議の議事録を作成してください。

【絶対遵守ルール】
1. 推測や勝手な創作（ハルシネーション）は厳禁ですが、実際にデータ内で語られている「具体的なエピソード」「背景にある理由」「本人や家族の意向」「各専門職の詳細な発言や見解」は絶対に省略せず、【可能な限り詳細に（ボリューミーに）】記載してください。
2. 単なる結果だけでなく、「なぜその結論に至ったか」のプロセスが第三者にも分かるように記録してください。
3. 情報が全くない項目のみ「特記事項なし」としてください。
4. 出力は「である調」で記載してください。
5. テキスト入力があるときは、そちらの情報を優先してください。特に、"item1"と"item2"の方向付けで記載されることがあるので留意してください。

【入力データ】
${inputText ? inputText : '（テキスト入力なし。音声データを優先してください）'}
  `;

  let parts = [{"text": prompt}];
  if (fileUri) {
    parts.push({ "file_data": { "mime_type": audioData.mimeType, "file_uri": fileUri } });
  }

  const payload = { 
    "contents": [{ "parts": parts }],
    "generationConfig": {
      "responseMimeType": "application/json",
      "temperature": 0.2,
      "responseSchema": {
        "type": "OBJECT",
        "properties": {
          "summary": { "type": "STRING", "description": "会議の主な論点、背景、決定事項を300〜500文字程度でしっかり要約" },
          "item1": { "type": "STRING", "description": "開催目的と、検討した議題を①②③とナンバリングして記載し、振り返りがしやすいように各議題100文字以内で要約" },
          "item2": { "type": "STRING", "description": "議題に対応させ、誰がどのような発言をしたかを詳細な文章で記載" },
          "item3": { "type": "STRING", "description": "話し合って合意・決定された内容や、決定に至った具体的な理由を詳細に記載する。このとき、各参加者が本会議後に本会議後に実施することが決まっていれば明記する" },
          "item4": { "type": "STRING", "description": "解決できなかった内容、今後の観察ポイント、次回の開催予定などを記載する。本会議で概ね方針感が決まっていそうであれば「必要時随時検討する」と記載" }
        },
        "required": ["summary", "item1", "item2", "item3", "item4"]
      }
    }
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  let finalResult = null;

  try {
    const response = UrlFetchApp.fetch(generateUrl, options);
    const json = JSON.parse(response.getContentText());
    
    if (json.error) {
      finalResult = { error: "【AI解析エラー】: " + json.error.message };
    } else if (!json.candidates || !json.candidates[0].content) {
      finalResult = { error: "【AIブロックエラー】: セキュリティブロックが発動しました。" };
    } else {
      const text = json.candidates[0].content.parts[0].text;
      finalResult = JSON.parse(text);
    }
  } catch (e) {
    finalResult = { error: "【システムエラー】: " + e.message };
  }

  if (fileNameToCleanUp) {
    try {
      UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/' + fileNameToCleanUp + '?key=' + apiKey, { method: 'delete', muteHttpExceptions: true });
    } catch(e) {}
  }

  return finalResult;
}

function saveLogAndFinish(meetingDate, clientName, inputText, summary, item1, item2, item3, item4) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("スプレッドシートが見つかりません。");
    
    const sheet = ss.getSheets()[0]; 
    sheet.appendRow([
      new Date(),
      meetingDate || "",
      clientName || "名前なし",
      inputText || "",
      summary || "",
      item1 || "",
      item2 || "",
      item3 || "",
      item4 || ""
    ]);
    return { success: true, message: "データベースに保存しました！PCからいつでも呼び出せます。" };
  } catch (e) {
    return { error: e.message };
  }
}

function getHistorySummaryList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("スプレッドシートが見つかりません。");
    
    const sheet = ss.getSheets()[0];
    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return []; 
    
    const data = sheet.getRange(1, 1, lastRow, Math.min(sheet.getLastColumn(), 4)).getValues();
    
    let history = [];
    for (let i = lastRow - 1; i >= 0; i--) {
      const ts = data[i][0];
      if (!ts || ts === "登録日時" || (typeof ts === 'string' && ts.includes("日時"))) continue; 
      
      let dateStr = "";
      if (data[i][1] instanceof Date) {
        dateStr = Utilities.formatDate(data[i][1], "JST", "yyyy/MM/dd");
      } else {
        dateStr = String(data[i][1] || ""); 
      }

      let tsStr = "";
      if (ts instanceof Date) {
        tsStr = Utilities.formatDate(ts, "JST", "yyyy/MM/dd HH:mm");
      } else {
        tsStr = String(ts);
      }

      history.push({
        id: i + 1,
        timestamp: tsStr,
        meetingDate: dateStr,
        clientName: String(data[i][2] || "名前なし")
      });
      
      if (history.length >= 50) break; 
    }
    return history;
  } catch (e) {
    return { error: e.message };
  }
}

function getLogDetail(rowId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("スプレッドシートが見つかりません。");
    
    const sheet = ss.getSheets()[0];
    const row = parseInt(rowId, 10);
    if (isNaN(row) || row < 1 || row > sheet.getLastRow()) {
      return { error: "該当レコードが見つかりません。" };
    }
    
    const data = sheet.getRange(row, 1, 1, Math.max(sheet.getLastColumn(), 9)).getValues()[0];
    let dateStr = (data[1] instanceof Date) ? Utilities.formatDate(data[1], "JST", "yyyy/MM/dd") : String(data[1] || "");
    let tsStr = (data[0] instanceof Date) ? Utilities.formatDate(data[0], "JST", "yyyy/MM/dd HH:mm") : String(data[0] || "");

    return {
      id: row,
      timestamp: tsStr,
      meetingDate: dateStr,
      clientName: String(data[2] || "名前なし"),
      inputText: String(data[3] || ""),
      summary: String(data[4] || ""),
      item1: String(data[5] || ""),
      item2: String(data[6] || ""),
      item3: String(data[7] || ""),
      item4: String(data[8] || "")
    };
  } catch (e) {
    return { error: e.message };
  }
}
