/**
 * Google Meetの文字起こしからおすすめ人物の他己紹介文を生成するGAS。
 *
 * 本スクリプトは以下のフローを自動化します。
 * 1. Google Drive に Google Meet の文字起こしファイルが追加されたタイミングでトリガー。
 * 2. ファイルの内容を取得し、生成AIで要約を作成。
 * 3. 所定のヘッダー構成でスプレッドシートを作成し、所定のフォルダに保存。
 * 4. 既存の候補者リストからおすすめ人物を検索。
 * 5. おすすめ人物に合わせた他己紹介文を生成し、Slack に通知。
 *
 * 実行前に CONFIG の各値を環境に合わせて設定してください。
 */

const CONFIG = {
  /** Google Meet の文字起こしを保存しているフォルダの ID */
  TRANSCRIPTION_FOLDER_ID: 'REPLACE_WITH_TRANSCRIPTION_FOLDER_ID',
  /** 要約スプレッドシートを保存するフォルダの ID */
  SUMMARY_FOLDER_ID: 'REPLACE_WITH_SUMMARY_FOLDER_ID',
  /** 要約スプレッドシートに使用するヘッダー */
  SUMMARY_HEADERS: ['氏名', 'ミーティング日時', '要約', 'キーフレーズ', '次のアクション'],
  /** 要約スプレッドシートのシート名 */
  SUMMARY_SHEET_NAME: 'Summary',
  /** 候補者リストを保持しているスプレッドシートの ID */
  PEOPLE_SHEET_ID: 'REPLACE_WITH_PEOPLE_SHEET_ID',
  /** 候補者リストのシート名 */
  PEOPLE_SHEET_NAME: 'People',
  /** Slack Incoming Webhook URL */
  SLACK_WEBHOOK_URL: 'https://hooks.slack.com/services/REPLACE/ME',
  /** Generative Language API (Gemini) のモデル名 */
  MODEL_NAME: 'models/gemini-1.5-pro',
  /** Generative Language API の API キー */
  AI_API_KEY: 'REPLACE_WITH_API_KEY'
};

/**
 * Google Drive のインストール型トリガーから呼び出されるエントリポイント。
 * @param {GoogleAppsScript.Events.DriveOnChange} e
 */
function onTranscriptionUploaded(e) {
  const fileId = e && e.drive && e.drive.fileId;
  if (!fileId) {
    console.log('Drive トリガーに fileId が含まれていません。イベント: ' + JSON.stringify(e));
    return;
  }

  const file = DriveApp.getFileById(fileId);
  if (!isTargetTranscription(file)) {
    console.log('対象外のファイルのため処理をスキップ: ' + file.getName());
    return;
  }

  const transcriptText = extractTranscriptText(file);
  if (!transcriptText) {
    console.error('文字起こしテキストを取得できませんでした: ' + file.getName());
    return;
  }

  const summary = summarizeTranscript(transcriptText, file.getName());
  const summarySpreadsheet = saveSummarySpreadsheet(summary);
  const candidates = loadRecommendationCandidates();
  if (!candidates.length) {
    console.error('候補者リストが空です。スプレッドシートの設定を確認してください。');
    return;
  }

  const recommendation = findRecommendedContact(summary, candidates) || {};
  const introText = generatePeerIntroduction(summary, recommendation);
  sendSlackNotification({ file, summarySpreadsheet, recommendation, introText });
}

/**
 * 対象とする文字起こしファイルかどうかを判定します。
 * @param {GoogleAppsScript.Drive.File} file
 * @return {boolean}
 */
function isTargetTranscription(file) {
  const folderIterator = file.getParents();
  const folderIds = [];
  while (folderIterator.hasNext()) {
    folderIds.push(folderIterator.next().getId());
  }
  const isInTargetFolder = folderIds.indexOf(CONFIG.TRANSCRIPTION_FOLDER_ID) !== -1;
  const mimeType = file.getMimeType();
  const supportedMimeTypes = [
    MimeType.GOOGLE_DOCS,
    MimeType.PLAIN_TEXT,
    'application/vnd.google-apps.document',
    'text/plain'
  ];
  return isInTargetFolder && supportedMimeTypes.indexOf(mimeType) !== -1;
}

/**
 * Google Meet の文字起こしからテキストを抽出します。
 * @param {GoogleAppsScript.Drive.File} file
 * @return {string}
 */
function extractTranscriptText(file) {
  const mimeType = file.getMimeType();
  try {
    if (mimeType === MimeType.GOOGLE_DOCS || mimeType === 'application/vnd.google-apps.document') {
      const doc = DocumentApp.openById(file.getId());
      return doc.getBody().getText();
    }
    const blob = file.getBlob();
    blob.setContentType('text/plain');
    return blob.getDataAsString('UTF-8');
  } catch (error) {
    console.error('文字起こし抽出中にエラーが発生: ' + error);
    return '';
  }
}

/**
 * 文字起こしを要約します。
 * @param {string} transcriptText
 * @param {string} meetingTitle
 * @return {{name: string, meetingDate: string, summary: string, keywords: string[], nextActions: string[]}}
 */
function summarizeTranscript(transcriptText, meetingTitle) {
  const prompt = [
    '以下は Google Meet の文字起こしです。',
    '内容を分析し、JSON で以下のキーを返してください。',
    JSON.stringify(CONFIG.SUMMARY_HEADERS),
    'JSON のキーはそれぞれ `name`, `meetingDate`, `summary`, `keywords`, `nextActions` としてください。',
    'name にはミーティングでフィーチャーされた人物名を入れてください。',
    'meetingDate は YYYY-MM-DD 形式で推測してください。',
    'keywords は 3~5 個の重要トピックを配列で。',
    'nextActions は今後のアクションを 1~3 件、配列で返してください。',
    'ミーティング名: ' + meetingTitle,
    '文字起こし:',
    transcriptText
  ].join('\n\n');

  const response = callGenerativeModel(prompt);
  const text = extractModelText(response);
  try {
    const parsed = JSON.parse(text);
    return {
      name: parsed.name || '不明',
      meetingDate: parsed.meetingDate || '',
      summary: parsed.summary || '',
      keywords: parsed.keywords || [],
      nextActions: parsed.nextActions || []
    };
  } catch (error) {
    console.error('要約の JSON 解析に失敗しました。レスポンス: ' + text);
    return {
      name: '不明',
      meetingDate: '',
      summary: text,
      keywords: [],
      nextActions: []
    };
  }
}

/**
 * 要約データを新規スプレッドシートに保存し、指定フォルダに移動します。
 * @param {{name: string, meetingDate: string, summary: string, keywords: string[], nextActions: string[]}} summary
 * @return {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function saveSummarySpreadsheet(summary) {
  const spreadsheetName = Utilities.formatString('Meet Summary_%s_%s', summary.name, summary.meetingDate);
  const sheet = SpreadsheetApp.create(spreadsheetName).getActiveSheet();
  sheet.setName(CONFIG.SUMMARY_SHEET_NAME);
  sheet.getRange(1, 1, 1, CONFIG.SUMMARY_HEADERS.length).setValues([CONFIG.SUMMARY_HEADERS]);
  const keywordsText = Array.isArray(summary.keywords) ? summary.keywords.join(', ') : summary.keywords;
  const nextActionsText = Array.isArray(summary.nextActions) ? summary.nextActions.join('\n') : summary.nextActions;
  sheet.getRange(2, 1, 1, CONFIG.SUMMARY_HEADERS.length).setValues([[summary.name, summary.meetingDate, summary.summary, keywordsText, nextActionsText]]);

  const spreadsheet = sheet.getParent();
  const file = DriveApp.getFileById(spreadsheet.getId());
  const folder = DriveApp.getFolderById(CONFIG.SUMMARY_FOLDER_ID);
  folder.addFile(file);
  const parents = file.getParents();
  while (parents.hasNext()) {
    const parent = parents.next();
    if (parent.getId() !== CONFIG.SUMMARY_FOLDER_ID) {
      parent.removeFile(file);
    }
  }
  return spreadsheet;
}

/**
 * 推薦候補者リストを読み込みます。
 * @return {Array<{name: string, title: string, expertise: string, slackId: string}>}
 */
function loadRecommendationCandidates() {
  const sheet = SpreadsheetApp.openById(CONFIG.PEOPLE_SHEET_ID).getSheetByName(CONFIG.PEOPLE_SHEET_NAME);
  if (!sheet) {
    throw new Error('候補者シートが見つかりません: ' + CONFIG.PEOPLE_SHEET_NAME);
  }
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  return values
    .filter(function(row) { return row.join('').trim() !== ''; })
    .map(function(row) {
      return headers.reduce(function(acc, header, index) {
        const key = header.toString().trim();
        acc[key] = row[index];
        return acc;
      }, {});
    })
    .map(function(row) {
      return {
        name: row.name || row['氏名'] || '',
        title: row.title || row['肩書'] || '',
        expertise: row.expertise || row['専門'] || '',
        slackId: row.slackId || row['Slack'] || ''
      };
    })
    .filter(function(person) { return person.name; });
}

/**
 * 要約を基におすすめ人物を選出します。
 * @param {{summary: string, keywords: string[]}} summary
 * @param {Array<{name: string, title: string, expertise: string}>} candidates
 * @return {{name: string, title: string, expertise: string, slackId: string}}
 */
function findRecommendedContact(summary, candidates) {
  const prompt = [
    '以下はミーティングの要約情報です。',
    JSON.stringify(summary, null, 2),
    '以下の候補者リストから、最も相性が良い一人を選んでください。',
    '返答は JSON 形式で、`name`, `reason` の 2 つのキーを含めてください。',
    JSON.stringify(candidates, null, 2)
  ].join('\n\n');

  const response = callGenerativeModel(prompt);
  const text = extractModelText(response);
  try {
    const parsed = JSON.parse(text);
    const match = candidates.find(function(person) { return person.name === parsed.name; });
    if (match) {
      return Object.assign({}, match, { reason: parsed.reason || '' });
    }
  } catch (error) {
    console.error('推薦候補の解析に失敗しました。レスポンス: ' + text);
  }
  const fallback = candidates[0];
  if (!fallback) {
    return null;
  }
  return Object.assign({}, fallback, { reason: fallback.reason || '候補者リストの先頭を自動選出しました。' });
}

/**
 * 選出した人物の他己紹介文を生成します。
 * @param {{name: string, summary: string}} summary
 * @param {{name: string, title: string, expertise: string}} recommendation
 * @return {string}
 */
function generatePeerIntroduction(summary, recommendation) {
  const safeRecommendation = recommendation && Object.keys(recommendation).length
    ? recommendation
    : { name: '未設定', title: '', expertise: '' };
  const prompt = [
    '以下のミーティングの要約と紹介対象の情報を基に、100~150 文字程度の他己紹介文を日本語で作成してください。',
    '紹介文は Slack でそのまま送れる文体にしてください。',
    'ミーティング要約:',
    JSON.stringify(summary, null, 2),
    '紹介対象:',
    JSON.stringify(safeRecommendation, null, 2)
  ].join('\n\n');

  const response = callGenerativeModel(prompt);
  return extractModelText(response);
}

/**
 * Slack に結果を通知します。
 * @param {{file: GoogleAppsScript.Drive.File, summarySpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, recommendation: object, introText: string}} payload
 */
function sendSlackNotification(payload) {
  const recommendation = payload.recommendation || { name: '未設定', title: '', slackId: '', reason: '' };
  const spreadsheetUrl = payload.summarySpreadsheet ? payload.summarySpreadsheet.getUrl() : '';
  const displayName = recommendation.slackId ? '<@' + recommendation.slackId + '>' : recommendation.name;
  const slackPayload = {
    text: Utilities.formatString(
      ['新しい文字起こしを処理しました。',
        '*ファイル名*: %s',
        '*要約スプレッドシート*: %s',
        '*おすすめ人物*: %s (%s)',
        '*紹介文*: %s',
        recommendation.reason ? '*推薦理由*: ' + recommendation.reason : ''
      ].filter(Boolean).join('\n'),
      payload.file.getName(),
      spreadsheetUrl,
      displayName,
      recommendation.title,
      payload.introText
    )
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(slackPayload)
  };
  UrlFetchApp.fetch(CONFIG.SLACK_WEBHOOK_URL, options);
}

/**
 * Gemini API を呼び出します。
 * @param {string} prompt
 * @return {Object}
 */
function callGenerativeModel(prompt) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/' + CONFIG.MODEL_NAME + ':generateContent?key=' + encodeURIComponent(CONFIG.AI_API_KEY);
  const payload = {
    contents: [{role: 'user', parts: [{text: prompt}]}]
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify(payload)
  };
  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  if (!data.candidates || !data.candidates.length) {
    throw new Error('AI レスポンスが不正です: ' + response.getContentText());
  }
  return data;
}

/**
 * モデルのレスポンスからテキスト部分を抽出します。
 * @param {Object} response
 * @return {string}
 */
function extractModelText(response) {
  const candidate = response.candidates[0];
  if (candidate && candidate.content && candidate.content.parts && candidate.content.parts.length) {
    return candidate.content.parts.map(function(part) { return part.text || ''; }).join('\n');
  }
  if (candidate && candidate.outputText) {
    return candidate.outputText;
  }
  throw new Error('テキストを抽出できませんでした: ' + JSON.stringify(response));
}
