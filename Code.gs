// ============================================================================
// carecore 担当者会議AI議事録（COOPs / コパン訪問看護ステーション対応）
// Code.gs — カイポケ用 & iBow用 両対応・排他制御・後方互換版
// ============================================================================

const SHEET_HEADERS = [
  '登録日時',
  '会議実施日',
  '利用者氏名',
  '補足メモ・条件',
  '全体要約 / iBow全文',
  '①検討した項目',
  '②検討内容',
  '③結論',
  '④残された課題',
  '会議名',
  '開催区分',
  '出力先区分'
];

const FALLBACK_MODELS = ['gemini-3.6-flash', 'gemini-3.5-flash', 'gemini-2.5-flash'];

function getGeminiApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (key && key.trim()) return key.trim();
  throw new Error('スクリプトプロパティに GEMINI_API_KEY が設定されていません。プロジェクト設定からGemini APIキーを設定してください。');
}

function clipText(val, maxLen) {
  const limit = maxLen || 48000;
  const s = String(val == null ? '' : val);
  return s.length > limit ? s.substring(0, limit) + '...(文字数上限のため省略)' : s;
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'ok',
    service: 'carecore 担当者会議AI議事録 API (カイポケ & iBow / コパン訪問看護対応)',
    timestamp: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let responseData = {};
  try {
    const postData = JSON.parse(e.postData.contents || '{}');
    const action = postData.action;
    const payload = postData.payload || {};

    switch (action) {
      case 'ping':
        responseData = {
          success: true,
          message: 'スプレッドシートDBへの接続に成功しました。',
          hasGeminiKey: Boolean(PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'))
        };
        break;
      case 'generateMinutes':
        responseData = generateMinutes(
          payload.inputText,
          payload.audioData,
          payload.meetingName || payload.meetingType,
          payload.modelName,
          payload.targetSystem,
          payload.meetingReason
        );
        break;
      case 'saveLogAndFinish':
        responseData = saveLogAndFinish(
          payload.meetingDate,
          payload.clientName,
          payload.inputText,
          payload.summary,
          payload.item1,
          payload.item2,
          payload.item3,
          payload.item4,
          payload.id,
          payload.meetingName,
          payload.targetSystem,
          payload.meetingReason
        );
        break;
      case 'getHistorySummaryList':
        responseData = getHistorySummaryList();
        break;
      case 'getLogDetail':
        responseData = getLogDetail(payload.id);
        break;
      case 'deleteLog':
        responseData = deleteLog(payload.id);
        break;
      default:
        responseData = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    responseData = { error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(responseData))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrInitSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('スプレッドシートが見つかりません。');
  const sheet = ss.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(SHEET_HEADERS);
    const headerRange = sheet.getRange(1, 1, 1, SHEET_HEADERS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#7a6859');
    headerRange.setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function generateMinutes(inputText, audioData, meetingName, modelName, targetSystem, meetingReason) {
  const apiKey = getGeminiApiKey();
  let fileUri = null;
  let fileNameToCleanUp = null;

  if (audioData && audioData.base64) {
    const uploadUrl = 'https://generativelanguage.googleapis.com/upload/v1beta/files?key=' + apiKey;
    const blob = Utilities.newBlob(Utilities.base64Decode(audioData.base64), audioData.mimeType || 'audio/mp3', 'meeting_audio');

    const uploadOptions = {
      method: 'post',
      contentType: audioData.mimeType || 'audio/mp3',
      payload: blob,
      headers: { 'X-Goog-Upload-Protocol': 'raw' },
      muteHttpExceptions: true
    };

    try {
      const uploadRes = UrlFetchApp.fetch(uploadUrl, uploadOptions);
      const uploadJson = JSON.parse(uploadRes.getContentText());
      if (uploadJson.error) return { error: '【AI音声アップロードエラー】: ' + uploadJson.error.message };
      fileUri = uploadJson.file.uri;
      fileNameToCleanUp = uploadJson.file.name;
    } catch (e) {
      return { error: '【通信エラー】: 音声のアップロードに失敗しました。' + e.message };
    }
  }

  const isIbow = targetSystem === 'ibow';
  const mName = meetingName || 'サービス担当者会議';
  const mReason = meetingReason || '更新・プラン継続';

  const hasPrebuiltPrompt = inputText && (inputText.indexOf('【絶対遵守ルール】') !== -1 || inputText.indexOf('【会議名】') !== -1);

  const prompt = hasPrebuiltPrompt ? inputText : (isIbow ? `
あなたは「コパン訪問看護ステーション」の優秀で客観的な「カンファレンス・会議議事録作成スペシャリスト」です。
以下の【入力データ】の内容から、訪問看護記録ソフト「iBow」の1つの記録ウインドウにそのまま貼り付けられる「${mName}」の議事録を作成してください。

【会議名】${mName}
【実施の理由・開催区分】${mReason}
【出力先】iBow（単一ウインドウ貼付用）

【絶対遵守ルール】
1. 自事業所である「コパン訪問看護ステーション」が【どのようなケア・看護・リハビリを提供するか】【どのような役割を担うか】【カンファレンスや会議で決定した「コパン訪問看護ステーション」の具体的なToDo（観察項目・処置・訪問体制・他職種連携アクション）】に強くフォーカスして作成してください。
2. 開催区分（${mReason}）の背景や検討課題を的確に議事録に反映してください。
3. 推測や勝手な創作（ハルシネーション）は厳禁ですが、実際に語られている本人・家族の意向や主治医・ケアマネ・各専門職の発言・背景理由は省略せず詳細に記載してください。
4. 介護・医療・福祉の専門用語（ケアプラン、バイタル、ADL/IADL、褥瘡、服薬管理、レスパイト、区分変更、ショートステイ等）は、音声のゆらぎや誤変換を文脈から正式な専門用語・正しい漢字表記に自動補正してください。
5. iBowは入力ウインドウが1つだけのため、出力は全て "summary" フィールド1つの中に、以下の構成で見出しを付けて読みやすく整形して出力してください：
   ■ 会議概要・検討内容（本人・家族の意向、各職種の見解や議論の経緯）
   ■ 「コパン訪問看護ステーション」の提供ケア・役割・決定したToDo（箇条書きと詳細）
   ■ 今後の観察ポイント・連携事項・次回予定
6. 出力は簡潔かつ専門的な「である調」で記載してください。

【入力データ】
${inputText ? inputText : '（テキスト入力なし。音声データを優先して詳細に解析してください）'}
  `.trim() : `
あなたは居宅介護支援事業所の優秀で客観的な「議事録作成スペシャリスト」です。
以下の【入力データ】の内容から、介護ソフト「カイポケ」の記録様式に準拠した「${mName}」の議事録を作成してください。

【会議名】${mName}
【実施の理由・開催区分】${mReason}
【出力先】カイポケ用

【絶対遵守ルール】
1. 開催区分（${mReason}）の視点（新規・更新・区分変更・退院カンファレンス等）を意識して作成してください。
2. 推測や勝手な創作（ハルシネーション）は厳禁ですが、実際にデータ内で語られている「具体的なエピソード」「背景にある理由」「本人や家族の意向」「各専門職の詳細な発言や見解」は絶対に省略せず、【可能な限り詳細に（ボリューミーに）】記載してください。
3. 単なる結果だけでなく、「なぜその結論に至ったか」のプロセスが第三者（実地指導・監査等）にも分かるように記録してください。
4. 介護・医療・福祉の専門用語は、音声のゆらぎや誤変換を文脈から正式な専門用語・正しい漢字表記に自動補正してください。
5. 情報が全くない項目のみ「特記事項なし」としてください。
6. 出力は「である調」で記載してください。
7. テキスト入力があるときは、そちらの情報を優先してください。特に、"item1"と"item2"の方向付けで記載されることがあるので留意してください。

【入力データ】
${inputText ? inputText : '（テキスト入力なし。音声データを優先して詳細に解析してください）'}
  `.trim());

  const responseSchema = isIbow ? {
    type: 'OBJECT',
    properties: {
      summary: {
        type: 'STRING',
        description: 'iBowの1つの入力ウインドウにそのまま貼り付ける完成された議事録全文（■会議概要・検討内容、■「コパン訪問看護ステーション」の提供ケア・役割・決定したToDo、■今後の観察ポイント・連携事項 の見出し付きで詳細に記載）'
      }
    },
    required: ['summary']
  } : {
    type: 'OBJECT',
    properties: {
      summary: { type: 'STRING', description: '会議の主な論点、背景、本人・家族の意向、決定事項を300〜500文字程度でしっかり要約' },
      item1: { type: 'STRING', description: '開催目的と、検討した議題を①②③とナンバリングして記載し、振り返りがしやすいように各議題100文字以内で要約' },
      item2: { type: 'STRING', description: '議題に対応させ、どの職種（誰）がどのような発言・専門的見解を述べたかを詳細な文章で記載' },
      item3: { type: 'STRING', description: '話し合って合意・決定された内容や、決定に至った具体的な理由を詳細に記載。各参加者が本会議後に実施する役割分担や頻度・期間が決まっていれば明記する' },
      item4: { type: 'STRING', description: '解決できなかった内容、今後の観察・モニタリングポイント、次回の開催予定などを記載。本会議で概ね方針が決まっていれば今後の観察点とあわせて「必要時随時検討する」と記載' }
    },
    required: ['summary', 'item1', 'item2', 'item3', 'item4']
  };

  const parts = [{ text: prompt }];
  if (fileUri) {
    parts.push({ file_data: { mime_type: audioData.mimeType || 'audio/mp3', file_uri: fileUri } });
  }

  const payload = {
    contents: [{ parts: parts }],
    generationConfig: {
      responseMimeType: 'application/json',
      temperature: 0.2,
      responseSchema: responseSchema
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const models = [modelName || 'gemini-3.6-flash'].concat(FALLBACK_MODELS).filter(function(v, i, a) {
    return v && a.indexOf(v) === i;
  });

  let finalResult = null;
  let lastErrorMsg = '';

  try {
    for (let m = 0; m < models.length; m++) {
      const currentModel = models[m];
      const generateUrl = 'https://generativelanguage.googleapis.com/v1beta/models/' + currentModel + ':generateContent?key=' + apiKey;

      for (let attempt = 0; attempt <= 1; attempt++) {
        try {
          const response = UrlFetchApp.fetch(generateUrl, options);
          const status = response.getResponseCode();
          const json = JSON.parse(response.getContentText() || '{}');

          if (json.error || status >= 400) {
            lastErrorMsg = (json.error && json.error.message) ? json.error.message : ('HTTP ' + status);
            if (status === 404 || lastErrorMsg.toLowerCase().indexOf('not found') !== -1) {
              break; // 次のモデルへフォールバック
            }
            if ((status === 429 || status >= 500) && attempt < 1) {
              Utilities.sleep(2000);
              continue;
            }
            break;
          }

          if (!json.candidates || !json.candidates[0] || !json.candidates[0].content) {
            lastErrorMsg = 'コンテンツ生成がブロックされました。';
            break;
          }

          let text = json.candidates[0].content.parts[0].text || '';
          text = text.trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '').trim();
          finalResult = JSON.parse(text);
          break;
        } catch (innerErr) {
          lastErrorMsg = innerErr.message;
          if (attempt < 1) Utilities.sleep(1500);
        }
      }
      if (finalResult) break;
    }

    if (!finalResult) {
      finalResult = { error: '【AI解析エラー】: ' + (lastErrorMsg || '解析に失敗しました') };
    }
  } catch (e) {
    finalResult = { error: '【システムエラー】: ' + e.message };
  } finally {
    if (fileNameToCleanUp) {
      try {
        UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/' + fileNameToCleanUp + '?key=' + apiKey, {
          method: 'delete',
          muteHttpExceptions: true
        });
      } catch (e) {}
    }
  }

  return finalResult;
}

function saveLogAndFinish(meetingDate, clientName, inputText, summary, item1, item2, item3, item4, existingRowId, meetingName, targetSystem, meetingReason) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const sheet = getOrInitSheet();
    const rowValues = [
      new Date(),
      clipText(meetingDate || '', 200),
      clipText(clientName || '名前なし', 500),
      clipText(inputText || '', 48000),
      clipText(summary || '', 48000),
      clipText(item1 || '', 48000),
      clipText(item2 || '', 48000),
      clipText(item3 || '', 48000),
      clipText(item4 || '', 48000),
      clipText(meetingName || 'サービス担当者会議', 200),
      clipText(meetingReason || '更新・プラン継続', 200),
      targetSystem === 'ibow' ? 'iBow用' : 'カイポケ用'
    ];

    const parsedId = existingRowId ? parseInt(existingRowId, 10) : null;
    if (parsedId && !isNaN(parsedId) && parsedId >= 2 && parsedId <= sheet.getLastRow()) {
      sheet.getRange(parsedId, 1, 1, rowValues.length).setValues([rowValues]);
      return {
        success: true,
        id: parsedId,
        message: '既存の議事録データを上書き更新しました。'
      };
    } else {
      sheet.appendRow(rowValues);
      const newRowId = sheet.getLastRow();
      return {
        success: true,
        id: newRowId,
        message: 'データベースに保存しました！過去ログからいつでも呼び出せます。'
      };
    }
  } catch (e) {
    return { error: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function getHistorySummaryList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return [];

    const sheet = ss.getSheets()[0];
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol < 1) return [];

    const fetchCols = Math.min(Math.max(lastCol, 5), 12);
    const data = sheet.getRange(1, 1, lastRow, fetchCols).getValues();

    const history = [];
    for (let i = lastRow - 1; i >= 1; i--) {
      const ts = data[i][0];
      if (!ts || ts === '登録日時' || (typeof ts === 'string' && (ts.indexOf('日時') !== -1 || ts.indexOf('登録') !== -1))) continue;

      let dateStr = '';
      if (data[i][1] instanceof Date) {
        dateStr = Utilities.formatDate(data[i][1], 'JST', 'yyyy/MM/dd');
      } else {
        dateStr = String(data[i][1] || '');
      }

      let tsStr = '';
      if (ts instanceof Date) {
        tsStr = Utilities.formatDate(ts, 'JST', 'yyyy/MM/dd HH:mm');
      } else {
        tsStr = String(ts);
      }

      const rawInput = String(data[i][3] || '');
      const summaryPreview = data[i][4] ? String(data[i][4]).substring(0, 60) : '';
      let meetingName = data[i][9] ? String(data[i][9]) : '';
      let col10 = data[i][10] ? String(data[i][10]) : '';
      let col11 = data[i][11] ? String(data[i][11]) : '';

      let meetingReason = '';
      let targetSystem = '';

      if (col11) {
        meetingReason = col10;
        targetSystem = col11;
      } else if (col10 === 'iBow用' || col10 === 'カイポケ用' || col10 === 'ibow' || col10 === 'kaipoke') {
        targetSystem = col10;
      } else {
        meetingReason = col10;
      }

      if (!meetingName && rawInput.indexOf('【') === 0) {
        const match = rawInput.match(/^【会議名:([^／\]]+)(?:[／/]\s*区分:([^／\]]+))?(?:[／/]\s*出力先:([^】]+))?】/);
        if (match) {
          meetingName = match[1].trim();
          if (match[2] && !meetingReason) meetingReason = match[2].trim();
          if (match[3] && !targetSystem) targetSystem = match[3].trim();
        }
      }

      history.push({
        id: i + 1,
        timestamp: tsStr,
        meetingDate: dateStr,
        clientName: String(data[i][2] || '名前なし'),
        summaryPreview: summaryPreview,
        meetingName: meetingName || 'サービス担当者会議',
        meetingReason: meetingReason || '更新・プラン継続',
        targetSystem: targetSystem || 'カイポケ用'
      });

      if (history.length >= 100) break;
    }
    return history;
  } catch (e) {
    console.error(e);
    return [];
  }
}

function getLogDetail(rowId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('スプレッドシートが見つかりません。');

    const sheet = ss.getSheets()[0];
    const row = parseInt(rowId, 10);
    if (isNaN(row) || row < 1 || row > sheet.getLastRow()) {
      return { error: '該当レコードが見つかりません。' };
    }

    const data = sheet.getRange(row, 1, 1, Math.max(sheet.getLastColumn(), 12)).getValues()[0];
    const dateStr = (data[1] instanceof Date) ? Utilities.formatDate(data[1], 'JST', 'yyyy/MM/dd') : String(data[1] || '');
    const tsStr = (data[0] instanceof Date) ? Utilities.formatDate(data[0], 'JST', 'yyyy/MM/dd HH:mm') : String(data[0] || '');

    const col10 = String(data[10] || '');
    const col11 = String(data[11] || '');
    let meetingReason = '';
    let targetSystem = '';

    if (col11) {
      meetingReason = col10;
      targetSystem = col11;
    } else if (col10 === 'iBow用' || col10 === 'カイポケ用' || col10 === 'ibow' || col10 === 'kaipoke') {
      targetSystem = col10;
    } else {
      meetingReason = col10;
    }

    return {
      id: row,
      timestamp: tsStr,
      meetingDate: dateStr,
      clientName: String(data[2] || '名前なし'),
      inputText: String(data[3] || ''),
      summary: String(data[4] || ''),
      item1: String(data[5] || ''),
      item2: String(data[6] || ''),
      item3: String(data[7] || ''),
      item4: String(data[8] || ''),
      meetingName: String(data[9] || ''),
      meetingReason: meetingReason,
      targetSystem: targetSystem
    };
  } catch (e) {
    return { error: e.message };
  }
}

function deleteLog(rowId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('スプレッドシートが見つかりません。');

    const sheet = ss.getSheets()[0];
    const row = parseInt(rowId, 10);
    if (isNaN(row) || row < 1 || row > sheet.getLastRow()) {
      return { error: '該当レコードが見つかりません。' };
    }

    sheet.deleteRow(row);
    return { success: true, message: '該当する過去ログを削除しました。' };
  } catch (e) {
    return { error: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
