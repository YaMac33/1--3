/******************************************************
 * C: スライド生成（雛形）
 * - Form送信 or 親GASからの呼び出しで、1行分を処理
 ******************************************************/

// ========= 設定 =========
const C_CONFIG = {
  SPREADSHEET_ID: '1j5bCNucxL9QVS_iq_RYaeL1_vWTsxHArUSJuRzhQmiw',
  SHEET_NAME: '_QUEUE',

  // 入力列
  COL_THEME: 1, // A: テーマ

  // 出力列
  COL_STATUS: 2,     // B: C_STATUS
  COL_SLIDES_URL: 3, // C: C_SLIDES_URL
  COL_ERROR: 4,      // D: C_ERROR

  // 保存先（任意）
  OUTPUT_FOLDER_ID: '<<保存先フォルダID（任意）>>', // 空ならルート直下
  // テンプレを使うなら（任意）
  TEMPLATE_SLIDES_ID: '', // 例: '1Abc...'; 空なら新規作成

  // 冪等判定：URLが入ってたらスキップ
  SKIP_IF_URL_EXISTS: true,
};

// ========= エントリーポイント（単体運用） =========
function onFormSubmit(e) {
  // インストール型トリガー推奨
  runC_forEvent_(e);
}

/**
 * 親GASから呼ぶ場合の入口（任意）
 * 親側で e をそのまま渡せる設計にすると楽です。
 */
function handleC_fromParent_(e) {
  runC_forEvent_(e);
}

/**
 * 手動実行（最終行を処理）— 動作確認用
 */
function runC_latestRow() {
  const ss = SpreadsheetApp.openById(C_CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(C_CONFIG.SHEET_NAME);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  runC_forRow_(lastRow);
}

// ========= コア処理 =========

function runC_forEvent_(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    // フォーム送信イベントなら range から行番号が取れる
    const row = (e && e.range) ? e.range.getRow() : null;

    // 取れない場合は最終行（親GAS側が独自eを作る時など）
    if (!row) {
      const ss = SpreadsheetApp.openById(C_CONFIG.SPREADSHEET_ID);
      const sh = ss.getSheetByName(C_CONFIG.SHEET_NAME);
      runC_forRow_(sh.getLastRow());
      return;
    }

    // ヘッダー行は無視
    if (row <= 1) return;

    runC_forRow_(row);
  } finally {
    lock.releaseLock();
  }
}

function runC_forRow_(row) {
  const ss = SpreadsheetApp.openById(C_CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(C_CONFIG.SHEET_NAME);

  // 1行分取得（必要な列だけでもOK）
  const theme = String(sh.getRange(row, C_CONFIG.COL_THEME).getValue() || '').trim();
  if (!theme) {
    writeCStatus_(sh, row, 'SKIP', '', 'テーマが空です');
    return;
  }

  const existingUrl = String(sh.getRange(row, C_CONFIG.COL_SLIDES_URL).getValue() || '').trim();
  if (C_CONFIG.SKIP_IF_URL_EXISTS && existingUrl) {
    writeCStatus_(sh, row, 'SKIP', existingUrl, '');
    return;
  }

  writeCStatus_(sh, row, 'RUNNING', '', '');

  try {
    // 1) スライド構成を作る（雛形：固定でもOK。後でAI生成に差し替え可能）
    const deck = buildDeckPlan_(theme);

    // 2) スライドを作成（テンプレコピー or 新規）
    const presId = createPresentation_(deck.title);

    // 3) スライドに流し込み（タイトル/本文/ノート）
    applyDeckPlanToSlides_(presId, deck);

    // 4) 保存先フォルダへ移動（任意）
    const url = `https://docs.google.com/presentation/d/${presId}/edit`;
    moveToFolderIfNeeded_(presId);

    writeCStatus_(sh, row, 'DONE', url, '');
  } catch (err) {
    const msg = (err && err.stack) ? err.stack : String(err);
    writeCStatus_(sh, row, 'ERROR', '', msg);
  }
}

// ========= スライド生成ロジック（雛形） =========

/**
 * スライド構成（雛形）
 * 後でここを「AIでアウトライン生成」に差し替えればOK
 */
function buildDeckPlan_(theme) {
  return {
    title: `${theme}｜スライド`,
    slides: [
      { title: theme, body: '結論 / 要点を1文で', notes: '冒頭で結論を短く言う' },
      { title: '背景', body: 'なぜ今これが重要か\n・理由1\n・理由2', notes: '背景を具体例で補足' },
      { title: 'ポイント', body: '押さえるべき3点\n1.\n2.\n3.', notes: '各ポイントを短く' },
      { title: '次のアクション', body: '今日からやること\n・\n・', notes: '具体的な行動に落とす' },
    ],
  };
}

/**
 * テンプレあり: コピーして使う
 * テンプレなし: 新規作成
 */
function createPresentation_(title) {
  if (C_CONFIG.TEMPLATE_SLIDES_ID) {
    const copied = DriveApp.getFileById(C_CONFIG.TEMPLATE_SLIDES_ID).makeCopy(title);
    return copied.getId();
  }
  const pres = SlidesApp.create(title);
  return pres.getId();
}

/**
 * DeckPlan を Google Slides に反映
 * - 既存スライドを一旦クリアして作り直す（最小実装）
 * - テンプレ運用にする場合は「既存レイアウトに流し込む」に差し替え
 */
function applyDeckPlanToSlides_(presentationId, deck) {
  const pres = SlidesApp.openById(presentationId);

  // 一旦全削除して作り直し（テンプレで固定レイアウトにしたい場合はここを変更）
  const existing = pres.getSlides();
  for (let i = existing.length - 1; i >= 0; i--) {
    existing[i].remove();
  }

  deck.slides.forEach((s, idx) => {
    // タイトル＆本文の簡易レイアウト
    const slide = pres.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);

    // タイトル
    slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE)
      .asShape().getText().setText(String(s.title || `Slide ${idx + 1}`));

    // 本文
    slide.getPlaceholder(SlidesApp.PlaceholderType.BODY)
      .asShape().getText().setText(String(s.body || ''));

    // スピーカーノート
    const notes = slide.getNotesPage().getSpeakerNotesShape().getText();
    notes.setText(String(s.notes || ''));
  });

  pres.saveAndClose();
}

/**
 * 保存先フォルダへ移動（任意）
 */
function moveToFolderIfNeeded_(fileId) {
  const folderId = String(C_CONFIG.OUTPUT_FOLDER_ID || '').trim();
  if (!folderId) return;

  const file = DriveApp.getFileById(fileId);
  const folder = DriveApp.getFolderById(folderId);

  // 既存親フォルダから外す（MyDrive直下など）
  const parents = file.getParents();
  while (parents.hasNext()) {
    const p = parents.next();
    p.removeFile(file);
  }
  folder.addFile(file);
}

// ========= シート書き込み =========
function writeCStatus_(sheet, row, status, url, error) {
  sheet.getRange(row, C_CONFIG.COL_STATUS).setValue(status);
  sheet.getRange(row, C_CONFIG.COL_SLIDES_URL).setValue(url || '');
  sheet.getRange(row, C_CONFIG.COL_ERROR).setValue(error || '');
}

// ========= 初期セットアップ（任意） =========
/**
 * インストール型トリガーを張る（初回だけ手動実行）
 * すでに親GAS側で一括トリガー管理するなら不要
 */
function setupCTrigger() {
  // 既存トリガー掃除（好みで）
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onFormSubmit') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(C_CONFIG.SPREADSHEET_ID)
    .onFormSubmit()
    .create();
}
