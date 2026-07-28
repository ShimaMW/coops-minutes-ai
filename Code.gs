// =============================================
// COOPs 議事録AI v4.0 — Code.gs (Web API版 / Vercel連携用)
// =============================================

const MASTER_EMP_SHEET  = 'マスタ_従業員';
const MASTER_TYPE_SHEET = 'マスタ_会議種別';
const LOG_SHEET         = 'アジェンダ議事録ログ';

const GEMINI_MODEL   = 'gemini-1.5-flash';

// ログシートの列インデックス（1始まり）
const COL = {
  record_id:          1,
  timestamp:          2,
  meeting_date:       3,
  dept:               4,
  meeting_type:       5,
  participants:       6,
  client_name:        7,
  user_topics:        8,
  agenda_body:        9,
  status:            10,
  minutes_timestamp: 11,
  input_text:        12,
  summary:           13,
  agenda_items:      14,
  key_discussions:   15,
  action_plans:      16,
  culture_notes:     17,
  next_agenda:       18,
  facilitator_feedback: 19
};

// スクリプトプロパティからGemini APIキーを取得
function getGeminiApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) {
    throw new Error('GASのスクリプトプロパティに GEMINI_API_KEY が設定されていません。プロジェクト設定から追加してください。');
  }
  return key.trim();
}

// =============================================
// HTTP リクエストハンドラ (doPost / doGet)
// =============================================
function doGet(e) {
  return HtmlService.createHtmlOutput('COOPs 議事録AI API Endpoint is Active.');
}

function doPost(e) {
  let responseData = {};
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    const payload = postData.payload || {};

    switch(action) {
      case 'getMasterData':
        responseData = getMasterData();
        break;
      case 'addEmployee':
        responseData = addEmployee(payload.dept, payload.name, payload.role);
        break;
      case 'generateAgenda':
        responseData = generateAgenda(payload);
        break;
      case 'saveAgenda':
        responseData = saveAgenda(payload);
        break;
      case 'getAgendaList':
        responseData = getAgendaList(payload.dept, payload.meetingType);
        break;
      case 'generateMinutes':
        responseData = generateMinutes(payload.inputText, payload.audioData, payload.meetingInfo);
        break;
      case 'saveMinutes':
        responseData = saveMinutes(payload);
        break;
      case 'getHistorySummaryList':
        responseData = getHistorySummaryList(payload.dept, payload.meetingType);
        break;
      case 'getLogDetail':
        responseData = getLogDetail(payload.recordId);
        break;
      case 'deleteLog':
        responseData = deleteLog(payload.recordId);
        break;
      case 'exportToGoogleDoc':
        responseData = exportToGoogleDoc(payload.recordId);
        break;
      case 'searchDriveFiles':
        responseData = searchDriveFiles(payload.keyword);
        break;
      case 'extractDriveText':
        responseData = extractDriveText(payload.fileId, payload.mimeType);
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

// =============================================
// シート初期化
// =============================================
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(MASTER_EMP_SHEET)) {
    const s = ss.insertSheet(MASTER_EMP_SHEET);
    s.appendRow(['部署', '氏名', '役職']);
    styleHeader(s, 3);
  }

  if (!ss.getSheetByName(MASTER_TYPE_SHEET)) {
    const s = ss.insertSheet(MASTER_TYPE_SHEET);
    s.appendRow(['種別名', '説明']);
    styleHeader(s, 2);
    [
      ['月次定例ミーティング', '月に一度の全体会議'],
      ['日次終礼',           '毎日の終礼・引き継ぎ'],
      ['利用者カンファレンス', '利用者のケアプラン検討会'],
      ['研修・勉強会',       'スタッフの学習・技術向上'],
      ['緊急対応会議',       'インシデント・緊急事態への対応'],
      ['採用・人事会議',     '採用活動・人事関連の検討']
    ].forEach(row => s.appendRow(row));
  }

  if (!ss.getSheetByName(LOG_SHEET)) {
    const s = ss.insertSheet(LOG_SHEET);
    const headers = [
      'レコードID','登録日時','会議日','部署','会議種別','参加者','対象利用者',
      'ユーザー入力議題','AIアジェンダ','ステータス',
      '議事録登録日時','入力メモ','全体要約','議題','議論','決定事項','組織文化','次回事項','AI評価'
    ];
    s.appendRow(headers);
    styleHeader(s, headers.length);
    s.setFrozenRows(1);
    s.setColumnWidth(1, 250);
    s.setColumnWidth(9, 400);
    [12,13,14,15,16,17,18,19].forEach(c => s.setColumnWidth(c, 350));
  }
}

function styleHeader(sheet, colCount) {
  sheet.getRange(1, 1, 1, colCount)
    .setFontWeight('bold')
    .setBackground('#354045')
    .setFontColor('#ffffff');
}

// =============================================
// マスタデータ取得
// =============================================
function getMasterData() {
  initSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const empData = ss.getSheetByName(MASTER_EMP_SHEET)
    .getDataRange().getValues().slice(1)
    .filter(r => r[0] && r[1])
    .map(r => ({ dept: String(r[0]), name: String(r[1]), role: String(r[2] || '') }));

  const departments = [...new Set(empData.map(e => e.dept))];

  const meetingTypes = ss.getSheetByName(MASTER_TYPE_SHEET)
    .getDataRange().getValues().slice(1)
    .filter(r => r[0])
    .map(r => ({ name: String(r[0]), desc: String(r[1] || '') }));

  return { employees: empData, departments, meetingTypes };
}

// =============================================
// 従業員追加
// =============================================
function addEmployee(dept, name, role) {
  SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(MASTER_EMP_SHEET)
    .appendRow([dept, name, role || '']);
  return { success: true };
}

// =============================================
// アジェンダ生成
// =============================================
function generateAgenda(p) {
  const participants = (p.participants || []).join('、') || '未定';

  const prompt = `あなたは介護事業所のベテランファシリテーターAIです。
以下の会議情報と議題メモをもとに、実用的な会議アジェンダを作成してください。

【会議情報】
- 会議日: ${p.meetingDate || '未定'}
- 部署: ${p.dept || '未定'}
- 会議種別: ${p.meetingType || '未定'}
- 参加者: ${participants}
- 想定所要時間: ${p.duration || '未定'}

【議題メモ（ユーザー入力）】
${p.topics || '（特記事項なし）'}

【作成方針】
- 各議題には「確認すべきポイント」「注意点」「AIからのアドバイス」を追記してください
- 会議種別に応じた介護現場特有の観点を盛り込んでください
- 【報告】【決定】【議論】等の種別を各議題に付与してください。すべて日本語表記とし、英語（Discussion等）は使用しないでください。

【出力形式】以下のJSONのみで回答（コードブロック不要）:
{
  "title": "会議タイトル（例：2025年4月 月次定例ミーティング）",
  "purpose": "この会議の目的（2〜3文）",
  "outcome": "この会議で達成したい成果・決定事項（2〜3文）",
  "review": "前回からの継続事項・振り返り（なければ空文字）",
  "agenda_items": "各議題の詳細（【報告】【決定】【議論】等の区分、確認ポイント、注意点、AIアドバイスを含む箇条書き。英語表記は使わずすべて日本語で出力）",
  "closing": "クロージング・次回予告に関するテキスト",
  "full_text": "印刷・配布用アジェンダ全文（ヘッダー・各議題・確認ポイントを含む、整形済みテキスト）"
}`;

  const result = callGemini(prompt, null);
  return extractJson(result);
}

// =============================================
// アジェンダ保存
// =============================================
function saveAgenda(d) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    initSheets();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
    const recordId = Utilities.getUuid();
    const now = formatNow();
    
    const clipText = (txt) => (txt && txt.length > 49000) ? txt.substring(0, 49000) + '\n...（文字数上限のため以降省略）' : (txt || '');

    const row = new Array(19).fill('');
    row[COL.record_id - 1]    = recordId;
    row[COL.timestamp - 1]    = now;
    row[COL.meeting_date - 1] = d.meetingDate || '';
    row[COL.dept - 1]         = d.dept || '';
    row[COL.meeting_type - 1] = d.meetingType || '';
    row[COL.participants - 1] = (d.participants || []).join('、');
    row[COL.client_name - 1]  = d.clientName || '';
    row[COL.user_topics - 1]  = clipText(d.userTopics);
    row[COL.agenda_body - 1]  = clipText(d.agendaFullText);
    row[COL.status - 1]       = 'アジェンダのみ';

    sheet.appendRow(row);
    return { success: true, recordId };
  } finally {
    lock.releaseLock();
  }
}

// =============================================
// アジェンダ一覧取得（ドロップダウン用）
// =============================================
function getAgendaList(dept, meetingType) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  let list = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[COL.record_id - 1] || String(r[COL.status - 1]) === '削除済') continue;

    const rowDept = String(r[COL.dept - 1]);
    const rowType = String(r[COL.meeting_type - 1]);

    if (dept && dept !== 'all' && rowDept !== dept) continue;
    if (meetingType && meetingType !== 'all' && rowType !== meetingType) continue;

    let md = r[COL.meeting_date - 1];
    let mdStr = (md instanceof Date) ? Utilities.formatDate(md, Session.getScriptTimeZone(), 'yyyy/MM/dd') : String(md).split('T')[0];

    list.push({
      recordId:    String(r[COL.record_id - 1]),
      meetingDate: mdStr,
      dept:        rowDept,
      meetingType: rowType,
      participants:String(r[COL.participants - 1]),
      clientName:  String(r[COL.client_name - 1]),
      userTopics:  String(r[COL.user_topics - 1]),
      agendaBody:  String(r[COL.agenda_body - 1]),
      status:      String(r[COL.status - 1])
    });
  }

  list.sort((a, b) => b.meetingDate.localeCompare(a.meetingDate));
  return list;
}

// =============================================
// 議事録生成
// =============================================
function generateMinutes(inputText, audioData, info) {
  info = info || {};
  const participants = (info.participants || []).join('、') || '未記載';

  const prompt = `あなたは介護事業所向けの会議ファシリテーターAIです。
以下の情報をもとに、詳細な議事録を作成してください。

【会議情報】
- 会議日: ${info.meetingDate || '未記載'}
- 部署: ${info.dept || '未記載'}
- 会議種別: ${info.meetingType || '未記載'}
- 参加者: ${participants}
- 対象利用者: ${info.clientName || 'なし'}
${info.agendaBody ? `\n【事前アジェンダ】\n${info.agendaBody}` : ''}

【入力メモ・テキスト起こし】
${inputText || '（テキスト入力なし：音声データから生成）'}

【作成方針】
- 提供されたデータ（音声データ・テキストデータ）を詳細までしっかりと読み込んでください
- アジェンダがある場合は各議題に対応した議論・結論を整理してください
- 議論がアジェンダの枠をはみ出た場合、あなたが客観的かつ自由に見て判断し、適切な項目・議題を作成してください
- 発言者が特定できる場合は「田中さん：〜」形式で記録してください
- アクションプランには必ず担当者・期日を含めてください
- 介護施設の運営理念（利用者本位、クローバーイズム等）の観点も含めてください
- 議論の経緯や背景を理解していない人でも、結論やアクションの理由が明確にわかるようにしてください
- 要約で丸められるよりも、会議全体の議論を詳細に記載できるのであれば文字数が多くなってしまっても大丈夫です

【出力形式】以下のJSONのみで回答（コードブロック不要）:
{
  "summary": "会議全体の要約（500〜2000文字）",
  "agenda_items": "議題と振り返り（箇条書き・アジェンダとの対応を明記）",
  "key_discussions": "主な議論・発言（発言者：内容の形式）",
  "action_plans": "決定事項・アクションプラン（担当者・期日を明記）",
  "culture_notes": "組織文化・理念に関する発言や気づき",
  "next_agenda": "次回の検討事項・宿題（具体的に）",
  "facilitator_feedback": "AIファシリテーターとしての評価（良かった点2つ・改善提案2つ）"
}`;

  const result = callGemini(prompt, audioData);
  return extractJson(result);
}

// =============================================
// 議事録保存（setValuesによる一括更新で高速化）
// =============================================
function saveMinutes(d) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    initSheets();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
    const now = formatNow();
    const clipText = (txt) => (txt && txt.length > 49000) ? txt.substring(0, 49000) + '\n...（文字数上限のため以降省略）' : (txt || '');

    if (d.agendaRecordId) {
      // 既存アジェンダ行を一括更新
      const values = sheet.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][COL.record_id - 1]) === String(d.agendaRecordId)) {
          const rowNum = i + 1;
          const rowData = [...values[i]];
          rowData[COL.status - 1]              = '議事録完了';
          rowData[COL.minutes_timestamp - 1]   = now;
          rowData[COL.input_text - 1]          = clipText(d.inputText);
          rowData[COL.summary - 1]             = clipText(d.summary);
          rowData[COL.agenda_items - 1]        = clipText(d.agenda_items);
          rowData[COL.key_discussions - 1]     = clipText(d.key_discussions);
          rowData[COL.action_plans - 1]        = clipText(d.action_plans);
          rowData[COL.culture_notes - 1]       = clipText(d.culture_notes);
          rowData[COL.next_agenda - 1]         = clipText(d.next_agenda);
          rowData[COL.facilitator_feedback - 1]= clipText(d.facilitator_feedback);

          sheet.getRange(rowNum, 1, 1, 19).setValues([rowData]);
          return { success: true, message: 'アジェンダに議事録を紐づけて保存しました' };
        }
      }
      return { error: '対応するアジェンダが見つかりませんでした。' };
    } else {
      // 新規行として保存
      const recordId = Utilities.getUuid();
      const row = new Array(19).fill('');
      row[COL.record_id - 1]           = recordId;
      row[COL.timestamp - 1]           = now;
      row[COL.meeting_date - 1]        = d.meetingDate || '';
      row[COL.dept - 1]                = d.dept || '';
      row[COL.meeting_type - 1]        = d.meetingType || '';
      row[COL.participants - 1]        = (d.participants || []).join('、');
      row[COL.client_name - 1]         = d.clientName || '';
      row[COL.status - 1]              = '議事録完了';
      row[COL.minutes_timestamp - 1]   = now;
      row[COL.input_text - 1]          = clipText(d.inputText);
      row[COL.summary - 1]             = clipText(d.summary);
      row[COL.agenda_items - 1]        = clipText(d.agenda_items);
      row[COL.key_discussions - 1]     = clipText(d.key_discussions);
      row[COL.action_plans - 1]        = clipText(d.action_plans);
      row[COL.culture_notes - 1]       = clipText(d.culture_notes);
      row[COL.next_agenda - 1]         = clipText(d.next_agenda);
      row[COL.facilitator_feedback - 1]= clipText(d.facilitator_feedback);
      sheet.appendRow(row);
      return { success: true, message: '議事録を新規保存しました', recordId };
    }
  } finally {
    lock.releaseLock();
  }
}

// =============================================
// 高速・軽量化：履歴サマリー一覧取得（メタデータのみ）
// =============================================
function getHistorySummaryList(dept, meetingType) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  let list = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const status = String(r[COL.status - 1]);
    if (!r[COL.record_id - 1] || status === '削除済') continue;

    const rowDept = String(r[COL.dept - 1]);
    const rowType = String(r[COL.meeting_type - 1]);

    if (dept && dept !== 'all' && rowDept !== dept) continue;
    if (meetingType && meetingType !== 'all' && rowType !== meetingType) continue;

    let md = r[COL.meeting_date - 1];
    let mdStr = (md instanceof Date) ? Utilities.formatDate(md, Session.getScriptTimeZone(), 'yyyy/MM/dd') : String(md).split('T')[0];

    const hasAgenda  = Boolean(r[COL.agenda_body - 1]);
    const hasMinutes = status === '議事録完了';

    list.push({
      recordId:    String(r[COL.record_id - 1]),
      timestamp:   String(r[COL.timestamp - 1]),
      meetingDate: mdStr,
      dept:        rowDept,
      meetingType: rowType,
      participants:String(r[COL.participants - 1]),
      clientName:  String(r[COL.client_name - 1]),
      status:      status,
      hasAgenda:   hasAgenda,
      hasMinutes:  hasMinutes
    });
  }

  list.sort((a, b) => b.meetingDate.localeCompare(a.meetingDate));
  return list;
}

// =============================================
// 新規：特定レコード1件の詳細データを取得
// =============================================
function getLogDetail(recordId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  if (!sheet) return { error: 'シートが見つかりません' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (String(r[COL.record_id - 1]) === String(recordId)) {
      let md = r[COL.meeting_date - 1];
      let mdStr = (md instanceof Date) ? Utilities.formatDate(md, Session.getScriptTimeZone(), 'yyyy/MM/dd') : String(md).split('T')[0];

      return {
        recordId:            String(r[COL.record_id - 1]),
        timestamp:           String(r[COL.timestamp - 1]),
        meetingDate:         mdStr,
        dept:                String(r[COL.dept - 1]),
        meetingType:         String(r[COL.meeting_type - 1]),
        participants:        String(r[COL.participants - 1]),
        clientName:          String(r[COL.client_name - 1]),
        userTopics:          String(r[COL.user_topics - 1]),
        agendaBody:          String(r[COL.agenda_body - 1]),
        status:              String(r[COL.status - 1]),
        minutesTimestamp:    String(r[COL.minutes_timestamp - 1]),
        inputText:           String(r[COL.input_text - 1]),
        summary:             String(r[COL.summary - 1]),
        agenda_items:        String(r[COL.agenda_items - 1]),
        key_discussions:     String(r[COL.key_discussions - 1]),
        action_plans:        String(r[COL.action_plans - 1]),
        culture_notes:       String(r[COL.culture_notes - 1]),
        next_agenda:         String(r[COL.next_agenda - 1]),
        facilitator_feedback:String(r[COL.facilitator_feedback - 1])
      };
    }
  }
  return { error: '該当レコードが見つかりませんでした' };
}

// =============================================
// ログ削除
// =============================================
function deleteLog(recordId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  if (!sheet) return { error: 'ログシートが見つかりません' };

  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][COL.record_id - 1]) === String(recordId)) {
      const rowNum = i + 1;
      sheet.getRange(rowNum, COL.status).setValue('削除済');
      sheet.hideRows(rowNum);
      return { success: true };
    }
  }
  return { error: '該当レコードが見つかりません' };
}

// =============================================
// Google Docs 出力
// =============================================
function exportToGoogleDoc(recordId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  if (!sheet) return { error: 'ログシートが見つかりません' };

  const values = sheet.getDataRange().getValues();
  let row = null;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][COL.record_id - 1]) === String(recordId)) {
      row = values[i]; break;
    }
  }
  if (!row) return { error: '該当レコードが見つかりません' };

  const dept        = String(row[COL.dept - 1]);
  const meetingDate = String(row[COL.meeting_date - 1]);
  const meetingType = String(row[COL.meeting_type - 1]);
  const participants= String(row[COL.participants - 1]);
  const clientName  = String(row[COL.client_name - 1]);
  const status      = String(row[COL.status - 1]);
  const agendaBody  = String(row[COL.agenda_body - 1]);
  const hasMinutes  = status === '議事録完了';

  const title = `COOPs_${dept}_${meetingDate}_${meetingType}`;
  const doc   = DocumentApp.create(title);
  const body  = doc.getBody();

  const h1 = body.appendParagraph('COOPs 議事録AI');
  h1.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  h1.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph(`${meetingDate}　${dept}　${meetingType}`)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(`参加者：${participants}`)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  if (clientName) body.appendParagraph(`対象利用者：${clientName}`)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendHorizontalRule();

  if (agendaBody) {
    body.appendParagraph('■ アジェンダ').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(agendaBody);
  }

  if (hasMinutes) {
    if (agendaBody) body.appendHorizontalRule();
    body.appendParagraph('■ 議事録').setHeading(DocumentApp.ParagraphHeading.HEADING2);

    const sections = [
      { title: '📌 全体要約',           key: COL.summary - 1 },
      { title: '🔎 議題',               key: COL.agenda_items - 1 },
      { title: '💡 主な議論・発言',      key: COL.key_discussions - 1 },
      { title: '✨ 決定事項・アクション',key: COL.action_plans - 1 },
      { title: '🍀 組織文化・理念',      key: COL.culture_notes - 1 },
      { title: '🎉 次回の検討事項',      key: COL.next_agenda - 1 },
      { title: '🌌 AIファシリテーター評価', key: COL.facilitator_feedback - 1 }
    ];

    sections.forEach(s => {
      const val = String(row[s.key] || '');
      if (!val) return;
      body.appendParagraph(s.title).setHeading(DocumentApp.ParagraphHeading.HEADING3);
      body.appendParagraph(val);
    });
  }

  doc.saveAndClose();
  return { success: true, url: doc.getUrl(), title };
}

// =============================================
// Gemini API 呼び出し (PropertiesService利用)
// =============================================
function callGemini(prompt, audioData) {
  const apiKey = getGeminiApiKey();
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${apiKey}`;

  const parts = [];
  if (audioData && audioData.base64) {
    parts.push({ inline_data: { mime_type: audioData.mimeType, data: audioData.base64 } });
  }
  parts.push({ text: prompt });

  const payload = JSON.stringify({
    contents: [{ role: 'user', parts }],
    generationConfig: { temperature: 0.3, maxOutputTokens: 8192 }
  });

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload,
    muteHttpExceptions: true
  });

  const json = JSON.parse(res.getContentText());
  if (json.error) throw new Error(json.error.message || 'Gemini APIエラー');
  return json.candidates?.[0]?.content?.parts?.[0]?.text || '';
}

// =============================================
// Drive 連携
// =============================================
function searchDriveFiles(keyword) {
  const safeKw = (keyword || '').replace(/'/g, "\\'");
  const query = `title contains '${safeKw}' and trashed = false`;
  const files = DriveApp.searchFiles(query);
  const result = [];
  let count = 0;
  while (files.hasNext() && count < 30) {
    const f = files.next();
    const mime = f.getMimeType();
    if (mime === MimeType.GOOGLE_DOCS || mime === MimeType.PLAIN_TEXT) {
      result.push({
        id: f.getId(),
        name: f.getName(),
        mimeType: mime,
        lastUpdated: f.getLastUpdated().toISOString()
      });
      count++;
    }
  }
  result.sort((a,b) => new Date(b.lastUpdated).getTime() - new Date(a.lastUpdated).getTime());
  return result;
}

function extractDriveText(fileId, mimeType) {
  let text = '';
  if (mimeType === MimeType.GOOGLE_DOCS) {
    const doc = DocumentApp.openById(fileId);
    text = doc.getBody().getText();
  } else if (mimeType === MimeType.PLAIN_TEXT) {
    const file = DriveApp.getFileById(fileId);
    text = file.getBlob().getDataAsString();
  } else {
    return { error: 'この形式（' + mimeType + '）のテキスト抽出には対応していません。' };
  }
  return { success: true, text: text };
}

function extractJson(text) {
  const clean = text.replace(/```json|```/g, '').trim();
  const match = clean.match(/\{[\s\S]*\}/);
  if (!match) throw new Error('AIの出力をJSONとして解析できませんでした');
  return JSON.parse(match[0]);
}

function formatNow() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
}
