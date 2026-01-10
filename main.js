/**
 * =========================================
 * 親：フォーム送信トリガー → キューへジョブ投入
 * =========================================
 *
 * 使い方:
 * 1) CONFIG.SPREADSHEET_ID を回答先スプレッドシートIDにする
 * 2) initQueueAndTrigger() を手動実行（初回のみ）
 * 3) 以後、フォーム送信のたびに onFormSubmit_ が動き、A/B/Cジョブがキューに積まれる
 */

// ===== 設定 =====
const CONFIG = {
  // フォーム回答先スプレッドシートID（必須）
  SPREADSHEET_ID: "1j5bCNucxL9QVS_iq_RYaeL1_vWTsxHArUSJuRzhQmiw",

  // フォーム回答が入るシート名（未指定なら先頭シート）
  RESPONSES_SHEET_NAME: "prmpt",

  // キューシート名
  QUEUE_SHEET_NAME: "_QUEUE",

  // 作るジョブ種別（A/B/C）
  JOB_TYPES: ["A_HTML_GITHUB", "B_BLOG_WP", "CC_SLIDES_GEN"],
};

/**
 * 初回セットアップ：キューシート作成＋フォーム送信トリガー作成
 * ※最初の1回だけ手動実行してください
 */
function initQueueAndTrigger() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // キューシート作成（無ければ）
  let q = ss.getSheetByName(CONFIG.QUEUE_SHEET_NAME);
  if (!q) {
    q = ss.insertSheet(CONFIG.QUEUE_SHEET_NAME);
    q.appendRow([
      "jobId",
      "createdAt",
      "jobType",
      "status",      // PENDING / RUNNING / DONE / ERROR
      "sourceSheet",
      "sourceRow",
      "payloadJson",
      "retryCount",
      "lastError",
      "updatedAt",
    ]);
  }

  // 既存トリガーの重複を避ける（同じ関数のフォーム送信トリガーを削除）
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "onFormSubmit_") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // スプレッドシート（回答先）に対してフォーム送信トリガーを作成
  ScriptApp.newTrigger("onFormSubmit_")
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  Logger.log("✅ initQueueAndTrigger 完了：キューシート & onFormSubmit トリガー作成");
}

/**
 * フォーム送信時に呼ばれる（インストール型トリガー）
 * ここでは “A/B/C のジョブをキューに積むだけ”
 */
function onFormSubmit_(e) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30 * 1000);

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

    const responsesSheet = CONFIG.RESPONSES_SHEET_NAME
      ? ss.getSheetByName(CONFIG.RESPONSES_SHEET_NAME)
      : ss.getSheets()[0];

    const queueSheet = ss.getSheetByName(CONFIG.QUEUE_SHEET_NAME);
    if (!queueSheet) throw new Error("Queue sheet not found. Run initQueueAndTrigger() first.");

    // 送信された行番号（フォーム送信トリガーなら e.range が入る）
    const row = e && e.range ? e.range.getRow() : responsesSheet.getLastRow();

    // 送信内容（必要ならここに入れておく：後でA/B/Cが参照可能）
    // e.namedValues は { "質問": ["回答"] } の形
    const payload = {
      spreadsheetId: ss.getId(),
      sourceSheetName: responsesSheet.getName(),
      sourceRow: row,
      timestamp: (e && e.values) ? e.values[0] : null, // 通常A列がタイムスタンプ
      namedValues: (e && e.namedValues) ? e.namedValues : null,
    };

    // A/B/Cジョブを投入
    const now = new Date();
    const rowsToAppend = CONFIG.JOB_TYPES.map(jobType => ([
      makeJobId_(jobType),
      now,
      jobType,
      "PENDING",
      responsesSheet.getName(),
      row,
      JSON.stringify(payload),
      0,
      "",
      now,
    ]));

    queueSheet.getRange(queueSheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length)
      .setValues(rowsToAppend);

    Logger.log(`✅ Queue enqueued: row=${row} jobs=${CONFIG.JOB_TYPES.join(", ")}`);

  } finally {
    lock.releaseLock();
  }
}

/** ジョブID生成（衝突しにくい形） */
function makeJobId_(jobType) {
  const rand = Utilities.getUuid().slice(0, 8);
  const ts = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMddHHmmssSSS");
  return `${jobType}-${ts}-${rand}`;
}
