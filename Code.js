/**
 * 通知表所見作成支援アプリ（GAS）
 * Gemini 2.0 Flash (v1) 対応版
 */

const GEMINI_MODEL = 'gemini-2.0-flash-001';
const PROP_API_KEY = 'GEMINI_API_KEY';
const PROP_STYLE_PROFILE = 'STYLE_PROFILE_V1';
const SHEET_SAMPLES = '文体サンプル';
const SHEET_PROFILE = '文体プロファイル';

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('所見支援')
    .addItem('サイドバーを開く', 'showSidebar')
    .addSeparator()
    .addItem('文体サンプル用シートを作成/表示', 'ensureSampleSheet')
    .addItem('文体分析を実行（サンプルシート）', 'analyzeMyStyle')
    .addItem('文体プロファイルを表示・編集', 'showProfileSheet')
    .addItem('文体プロファイルを再読込', 'loadProfileFromSheet')
    .addSeparator()
    .addItem('APIキー設定', 'showSidebar')
    .addSeparator()
    .addItem('シートを配布用に初期化', 'initializeSheet')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('通知表所見アシスト');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getInitState() {
  const props = PropertiesService.getUserProperties();
  const apiKeySaved = !!(props.getProperty(PROP_API_KEY) || PropertiesService.getScriptProperties().getProperty(PROP_API_KEY));
  const styleProfileJson = props.getProperty(PROP_STYLE_PROFILE);
  const styleProfile = styleProfileJson ? JSON.parse(styleProfileJson) : null;
  const ss = SpreadsheetApp.getActive();
  const sampleSheet = ss.getSheetByName(SHEET_SAMPLES);
  const activeRange = ss.getActiveRange();
  const activeA1 = activeRange ? activeRange.getA1Notation() : '';
  const activeSheetName = ss.getActiveSheet().getName();

  return {
    apiKeySaved: apiKeySaved,
    styleProfileSummary: styleProfile ? summarizeStyleProfile_(styleProfile) : null,
    hasSampleSheet: !!sampleSheet,
    activeA1: activeA1,
    activeSheetName: activeSheetName
  };
}

function getSelectedCellContent() {
  const rangeList = SpreadsheetApp.getActiveRangeList();
  if (rangeList) {
    const allValues = [];
    const ranges = rangeList.getRanges();
    for (let i = 0; i < ranges.length; i++) {
      const values = ranges[i].getValues();
      allValues.push(...values);
    }
    // Flatten the 2D array, filter out empty cells, and join with newlines.
    const content = allValues.flat().filter(cell => cell.toString().trim() !== '').join('\n');
    return content;
  }
  return '';
}

function saveApiKey(key) {
  const trimmed = (key || '').trim();
  if (!trimmed) throw new Error('APIキーが空です。');
  PropertiesService.getUserProperties().setProperty(PROP_API_KEY, trimmed);
  return true;
}

function deleteApiKey() {
  const userProps = PropertiesService.getUserProperties();
  userProps.deleteProperty(PROP_API_KEY);
  
  // 念のためスクリプトプロパティからも削除試行
  const scriptProps = PropertiesService.getScriptProperties();
  scriptProps.deleteProperty(PROP_API_KEY);

  return true; // 成功を示す
}

function ensureSampleSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_SAMPLES);
  if (!sh) {
    sh = ss.insertSheet(SHEET_SAMPLES);
    sh.getRange('A1').setValue('過去に自分で作成した所見文（1セル=1件）');
    sh.setColumnWidths(1, 1, 640);
    sh.getRange('A1').setFontWeight('bold');
    sh.getRange('A2').setNote('例）一学期当初は〜 のように、実名や具体的大会名などは書かないでください。');
    sh.getRange('A2').setWrap(true);
  }
  ss.setActiveSheet(sh);
  return true;
}

function showProfileSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_PROFILE);
  if (!sh) {
    // sh = ss.insertSheet(SHEET_PROFILE);
    // sh.getRange('A1').setValue('ここにプロファイルが表示されます。先に文体分析を実行してください。');
    throw new Error('文体プロファイルシートがまだ作成されていません。先に文体分析を実行してください。');
  }
  ss.setActiveSheet(sh);
  return true;
}

function loadProfileFromSheet() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_PROFILE);
  if (!sh) {
    throw new Error('文体プロファイルシートが見つかりません。');
  }
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  const profile = {};
  const arrayKeys = ['dos', 'donts', 'phrase_bank', 'closing_patterns'];

  data.forEach(row => {
    const key = row[0];
    const value = row[1];
    if (key) {
      if (arrayKeys.includes(key)) {
        profile[key] = value.toString().split('\n').map(s => s.trim()).filter(s => s);
      } else {
        profile[key] = value;
      }
    }
  });

  if (!profile.style_name || !profile.summary) {
    throw new Error('シートからプロファイルを正しく読み込めませんでした。項目名が変更されていないか確認してください。');
  }

  PropertiesService.getUserProperties().setProperty(PROP_STYLE_PROFILE, JSON.stringify(profile));
  return summarizeStyleProfile_(profile);
}

function writeProfileToSheet_(profile) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_PROFILE);
  if (!sh) {
    sh = ss.insertSheet(SHEET_PROFILE);
  }
  sh.clear();
  sh.getRange('A1:B1').setValues([['項目', '内容']]).setFontWeight('bold');
  sh.setFrozenRows(1);

  const rows = [];
  const keys = Object.keys(profile);

  keys.forEach(key => {
    const value = profile[key];
    if (Array.isArray(value)) {
      rows.push([key, value.join('\n')]);
    } else {
      rows.push([key, value]);
    }
  });

  sh.getRange(2, 1, rows.length, 2).setValues(rows).setWrap(true).setVerticalAlignment('top');
  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 600);

  const note = 'このシートの内容を編集してから、メニューやサイドバーの「文体プロファイルを再読込」を実行すると、生成される文章に反映されます。\n' +
               'dos, donts, phrase_bank, closing_patterns の各項目は改行区切りで複数項目を編集できます。';
  sh.getRange('B1').setNote(note);
  ss.setActiveSheet(sh);
}

// 文体分析
function analyzeMyStyle() {
  const samples = readSamples_();
  if (samples.length < 5) {
    throw new Error('サンプル文が不足しています。最低5件以上（推奨10件）を ' + SHEET_SAMPLES + ' シートのA列に貼り付けてください。');
  }
  const profile = analyzeStyleWithGemini_(samples);
  
  writeProfileToSheet_(profile);

  // 新しいシートを作成してヘッダーを設定
  if (profile.parameters && profile.parameters.length > 0) {
    const ss = SpreadsheetApp.getActive();
    const sheetName = '所見作成';
    let sh = ss.getSheetByName(sheetName);
    if (!sh) {
      sh = ss.insertSheet(sheetName);
    }
    sh.clear();
    const headers = profile.parameters.concat(['箇条書きメモ', '生成された所見']);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    SpreadsheetApp.getUi().alert('文体分析が完了し、「' + SHEET_PROFILE + '」シートを更新しました。さらに、所見作成用の「' + sheetName + '」シートを準備しました。');
  } else {
    SpreadsheetApp.getUi().alert('文体分析が完了し、結果を「' + SHEET_PROFILE + '」シートに書き出しました。内容は直接編集可能です。');
  }

  PropertiesService.getUserProperties().setProperty(PROP_STYLE_PROFILE, JSON.stringify(profile));

  return summarizeStyleProfile_(profile);
}

function readSamples_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_SAMPLES);
  if (!sh) return [];
  const values = sh.getRange(2, 1, Math.max(0, sh.getLastRow() - 1), 1).getValues();
  const items = values.map(function(r){ return (r[0] || '').toString().trim(); }).filter(function(s){ return s.length > 0; });
  return items.map(sanitizeForPrivacy_);
}

function summarizeStyleProfile_(profile) {
  return {
    style_name: profile.style_name || '',
    summary: profile.summary || '',
    B_sentence_structure: profile.B_sentence_structure || '',
    D_overall_tone: profile.D_overall_tone || '',
    updatedAt: new Date().toISOString()
  };
}

// 所見生成
function generateRemark(input) {
  input = input || {};
  const memoText = input.memoText;
  const goalCode = input.goalCode;
  const charCount = input.charCount;
  const gradeLevel = input.gradeLevel;
  if (!memoText || !goalCode) throw new Error('メモとゴールを指定してください。');

  const ss = SpreadsheetApp.getActive();
  const range = ss.getActiveRange();
  if (!range) throw new Error('出力先セルが選択されていません。スプレッドシート上で出力したいセルを選んでください。');

  const props = PropertiesService.getUserProperties();
  const styleJson = props.getProperty(PROP_STYLE_PROFILE);
  const styleProfile = styleJson ? JSON.parse(styleJson) : null;

  const memos = memoText.split(/\r?\n/).map(function(s){ return s.trim(); }).filter(function(s){ return !!s; }).map(sanitizeForPrivacy_);
  if (memos.length === 0) throw new Error('箇条書きメモが空です。');

  const resultText = generateRemarkWithGemini_(memos, goalCode, styleProfile, charCount, gradeLevel);
  range.setWrap(true).setValue(resultText);

  return {
    text: resultText,
    writtenTo: ss.getActiveSheet().getName() + '!' + range.getA1Notation()
  };
}

// Gemini呼び出し：文体分析（responseMimeType を削除）
function analyzeStyleWithGemini_(samples) {
  const joined = samples.map(function(s, i){ return '【サンプル' + (i + 1) + '】\n' + s; }).join('\n\n');
  const instruction = `あなたは日本語の文章スタイルを分析する専門家です。
以下は、ある教員が過去に作成した通知表所見文のサンプルです。
特に次の観点を明確に抽出してください:
B：文の構成（一文の長さ、接続詞の使い方、段落の組み立て）
D：全体的なトーン（丁寧さ、温かみ、客観性、励ましの度合い など）
抽出した観点に加えて、所見の文脈で重要となるパラメーター（例：教科、単元、評価観点など）を抽出し、"parameters"としてキーの配列で返してください。
必ず次のJSONスキーマの1オブジェクトのみを返すこと。前置きやコードブロックは不要。
{
  "style_name": "string",
  "summary": "string",
  "B_sentence_structure": "string",
  "D_overall_tone": "string",
  "dos": ["string"],
  "donts": ["string"],
  "phrase_bank": ["string"],
  "closing_patterns": ["string"],
  "parameters": ["string"]
}`;

  const body = {
    contents: [
      { role: 'user', parts: [
        { text: instruction },
        { text: '--- サンプル開始 ---\n' + joined + '\n--- サンプル終了 ---' }
      ] }
    ],
    generationConfig: {
      temperature: 0.2,
      topP: 0.9,
      maxOutputTokens: 2048
    }
  };

  const data = geminiFetch_(body);
  const jsonText = extractTextFromGenerateContent_(data);
  const profile = safeJsonParse_(jsonText);
  if (!profile || !profile.B_sentence_structure || !profile.D_overall_tone) {
    throw new Error('文体分析の結果を正しく取得できませんでした。もう一度お試しください。');
  }
  return profile;
}

// Gemini呼び出し：所見生成
function generateRemarkWithGemini_(memos, goalCode, styleProfile, charCount, gradeLevel) {
  const goalSpec = goalToSpec_(goalCode);
  const styleGuidance = styleProfile ? formatStyleGuidance_(styleProfile) : '丁寧で温かく、簡潔かつ客観性を保った「です・ます調」で書く。';
  const privacyGuard = [
    '固有名詞（生徒名、学校名、具体的な大会名等）は出力に含めない。',
    '学期や回数などの数値は一般化して表現する（例：「複数回」「学期当初」など）。'
  ].join('\n');
  const gradeSpec = gradeToSpec_(gradeLevel);
  const lengthSpec = charCount ? `文字量の目安: ${charCount}字程度（厳密でなくてよい）。` : `文字量の目安: ${goalSpec.recommendedLength}（厳密でなくてよい）。`;

  const prompt = [
    'あなたは日本の学校教員が用いる通知表の所見文を作成する専門アシスタントです。',
    '以下の条件で、日本語の単一段落（必要なら2段落まで）で所見文を作成してください。',
    '',
    '【対象読者】',
    gradeSpec,
    '',
    '【文体指針】',
    styleGuidance,
    '',
    '【目的（ゴール）】',
    goalSpec.instruction,
    '',
    '【守るべき制約】',
    '- ' + privacyGuard,
    '- 過度に断定せず、観察に基づく表現を用いる。',
    '- 読み手（保護者/本人）に配慮し、評価と励ましのバランスを取る。',
    '- 表現の画一化を避け、メモの内容に即して具体と抽象のバランスをとる。',
    '',
    '【入力メモ（個人特定を避けた要約）】',
    memos.map(function(m){ return '- ' + m; }).join('\n'),
    '',
    '【出力フォーマット】',
    '- 日本語の自然な文章。箇条書きは使用しない。',
    `- ${lengthSpec}`,
    '- 最後は前向きな締めで終える。'
  ].join('\n');

  const body = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.6,
      topP: 0.95,
      maxOutputTokens: 2048
    }
  };

  const data = geminiFetch_(body);
  var text = extractTextFromGenerateContent_(data).trim();
  if (!text) throw new Error('所見生成に失敗しました。');
  text = postProcess_(text);
  return text;
}

// 共通: v1 + 2.0
function geminiFetch_(body) {
  const apiKey = PropertiesService.getUserProperties().getProperty(PROP_API_KEY)
    || PropertiesService.getScriptProperties().getProperty(PROP_API_KEY);
  if (!apiKey) throw new Error('Gemini APIキーが未設定です。サイドバーの「設定」で保存してください。');

  const url = 'https://generativelanguage.googleapis.com/v1/models/' +
              encodeURIComponent(GEMINI_MODEL) + ':generateContent?key=' + encodeURIComponent(apiKey);

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code >= 400) {
    throw new Error('Gemini APIエラー (' + code + '): ' + text);
  }
  const data = JSON.parse(text);
  const c = data.candidates && data.candidates[0];
  if (!c || c.finishReason === 'SAFETY') {
    throw new Error('出力がブロックされました。メモの表現を一般化してください。');
  }
  return data;
}

function extractTextFromGenerateContent_(data) {
  try {
    const c = data.candidates && data.candidates[0];
    const parts = (c && c.content && c.content.parts) || [];
    const texts = parts.map(function(p){ return p.text || ''; }).filter(function(s){ return !!s; });
    return texts.join('\n');
  } catch (e) {
    return '';
  }
}

function safeJsonParse_(str) {
  try {
    return JSON.parse(str);
  } catch (e) {
    const m = str && str.match(/\{[\s\S]*\}/);
    if (m) {
      try { return JSON.parse(m[0]); } catch (e2) {}
    }
    return null;
  }
}

// ユーティリティ
function sanitizeForPrivacy_(s) {
  var out = String(s);
  out = out.replace(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi, '[連絡先]');
  out = out.replace(/\b\d{2,4}[-\s]?\d{2,4}[-\s]?\d{3,4}\b/g, '[番号]');
  out = out.replace(/([一-龥]{2,4})(さん|くん|ちゃん)?/g, '[人物]');
  out = out.replace(/([一-龥A-Za-z0-9]+)大会/g, 'ある大会');
  out = out.replace(/([一-龥A-Za-z0-9]+)小学校|([一-龥A-Za-z0-9]+)中学校|([一-龥A-Za-z0-9]+)高等学校/g, 'ある学校');
  out = out.replace(/([1-6])年([1-9])組/g, 'ある学年の学級');
  return out.trim();
}

function gradeToSpec_(gradeLevel) {
  switch (gradeLevel) {
    case 'elementary_1':
      return '読者は小学校1年生とその保護者です。ひらがな中心の、非常に簡単で具体的な言葉を選び、学校生活の基本的な様子が伝わるようにしてください。';
    case 'elementary_2':
      return '読者は小学校2年生とその保護者です。簡単で具体的な言葉を中心にしつつ、少しだけ漢字を交ぜた表現を意識してください。';
    case 'elementary_3':
      return '読者は小学校3年生とその保護者です。抽象的な概念（協力、工夫など）を簡単な言葉で表現し、本人の思考が垣間見えるような記述を心がけてください。';
    case 'elementary_4':
      return '読者は小学校4年生とその保護者です。本人が理解できる範囲で、少し知的な語彙を使い、論理的な思考の芽生えを示すような表現が望ましいです。';
    case 'elementary_5':
      return '読者は小学校5年生とその保護者です。高学年としての自覚や、集団の中での役割を意識した言葉を選んでください。';
    case 'elementary_6':
      return '読者は小学校6年生とその保護者です。中学校進学を視野に入れ、自立心や将来への期待を促すような、少し大人びた表現を含めてください。';
    case 'middle_school':
      return '読者は中学生本人とその保護者です。本人の自覚を促すような、客観的で少し大人びた表現を意識してください。';
    case 'high_school':
      return '読者は高校生本人とその保護者です。自立した個人として尊重し、将来を見据えた客観的な事実と期待を伝えてください。';
    default:
      return '読者は児童・生徒本人とその保護者です。';
  }
}

function goalToSpec_(goalCode) {
  switch ((goalCode || '').toUpperCase()) {
    case 'A':
      return { instruction: '成長の過程を「初期→変化→現在」の流れで、エピソードを一般化して物語的に示す。変化の背景にある努力や姿勢にも触れる。', recommendedLength: '200〜300字' };
    case 'B':
      return { instruction: '保護者に安心感と期待感を与える。事実に基づく肯定的な観察→学校での支援→今後の見通しの順にまとめる。', recommendedLength: '180〜280字' };
    case 'C':
      return { instruction: '次に取れる具体的な行動を2〜3点、温かいトーンで提案する。強制ではなく選択肢として提示する。', recommendedLength: '160〜240字' };
    default:
      return { instruction: '丁寧で温かく、観察に基づくバランスのとれた所見を作成する。', recommendedLength: '180〜260字' };
  }
}

function formatStyleGuidance_(profile) {
  var lines = [];
  if (profile.B_sentence_structure) lines.push('文の構成: ' + profile.B_sentence_structure);
  if (profile.D_overall_tone) lines.push('全体トーン: ' + profile.D_overall_tone);
  if (Array.isArray(profile.dos) && profile.dos.length) lines.push('推奨: ' + profile.dos.slice(0, 5).join(' / '));
  if (Array.isArray(profile.donts) && profile.donts.length) lines.push('避ける: ' + profile.donts.slice(0, 5).join(' / '));
  if (Array.isArray(profile.phrase_bank) && profile.phrase_bank.length) lines.push('言い回し: ' + profile.phrase_bank.slice(0, 5).join(' / '));
  if (Array.isArray(profile.closing_patterns) && profile.closing_patterns.length) lines.push('締め: ' + profile.closing_patterns.slice(0, 3).join(' / '));
  lines.push('文体は「です・ます調」で統一する。');
  return lines.join('\n- ');
}

function postProcess_(text) {
  var t = text.replace(/^[\"'「」]+|[\"'「」]+$/g, '').trim();
  var parts = t.split(/\n{2,}/);
  if (parts.length > 2) t = parts.slice(0, 2).join('\n\n');
  return t;
}

/**
 * 配布用にシートを初期化する
 */
function initializeSheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'シートの初期化',
    '本当にこのスプレッドシートを初期状態に戻しますか？\n' +
    '「文体プロファイル」「所見作成」シートが削除され、' +
    '「文体サンプル」シートの入力データと、保存されたAPIキーが消去されます。\n' +
    'この操作は元に戻せません。',
    ui.ButtonSet.OK_CANCEL
  );

  if (response !== ui.Button.OK) {
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Delete specific sheets
    const sheetsToDelete = [SHEET_PROFILE, '所見作成'];
    sheetsToDelete.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        ss.deleteSheet(sheet);
      }
    });

    // 2. Clear content of the sample sheet (keep header)
    const sampleSheet = ss.getSheetByName(SHEET_SAMPLES);
    if (sampleSheet) {
      const lastRow = sampleSheet.getLastRow();
      if (lastRow > 1) {
        sampleSheet.getRange(2, 1, lastRow - 1, sampleSheet.getLastColumn()).clearContent();
      }
    }

    // 3. Delete user properties
    const userProps = PropertiesService.getUserProperties();
    userProps.deleteProperty(PROP_API_KEY);
    userProps.deleteProperty(PROP_STYLE_PROFILE);

    // 4. Delete script properties as well, just in case
    const scriptProps = PropertiesService.getScriptProperties();
    scriptProps.deleteProperty(PROP_API_KEY);


    ui.alert('初期化が完了しました。このスプレッドシートのコピーを配布できます。');

  } catch (e) {
    Logger.log('Initialization error: ' + e.toString());
    ui.alert('エラーが発生しました。初期化を完了できませんでした。');
  }
}
