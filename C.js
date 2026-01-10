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

/******************************
 * B: 記事生成 → WordPress投稿（方式2：子GAS雛形）
 ******************************/

/** ====== 設定 ====== */
const CONFIG_B = {
  SPREADSHEET_ID: '1j5bCNucxL9QVS_iq_RYaeL1_vWTsxHArUSJuRzhQmiw',
  QUEUE_SHEET_NAME: '_QUEUE',

  // 1回の起動で何件処理するか（親が複数回呼ぶ想定でもOK）
  MAX_ITEMS_PER_RUN: 1,

  // リトライ
  MAX_ATTEMPTS: 5,
  RETRY_BACKOFF_MINUTES: [1, 3, 10, 30, 120], // attempt=1.. の待ち時間

  // WordPress
  WP_BASE_URL: 'https://example.com', // 末尾スラなし
  WP_USERNAME: '<<wpユーザー名>>',
  WP_APP_PASSWORD: '<<wpアプリケーションパスワード>>', // "xxxx xxxx ...." 形式OK
  WP_POST_STATUS: 'publish', // 'draft' などに変更可
  WP_CATEGORY_IDS: [], // 例: [12, 34]
  WP_TAG_IDS: [],      // 例: [56, 78]

  // 生成物の最低品質ガード（空投稿防止）
  MIN_TITLE_LEN: 5,
  MIN_BODY_LEN: 200,
};

/**
 * Queue列名マップ（あなたのキュー列に合わせて変更）
 * ここだけ合わせれば動きます。
 */
const COL = {
  runId: 'runId',
  taskB: 'taskB',
  theme: 'theme',

  statusB: 'statusB',
  attemptB: 'attemptB',
  lockUntilB: 'lockUntilB',

  articleTitle: 'articleTitle',
  articleHtml: 'articleHtml',

  wpPostId: 'wpPostId',
  errorB: 'errorB',
  updatedAtB: 'updatedAtB',
};

/** ====== エントリポイント（親が呼ぶ想定 / 手動実行も可） ====== */
function runWorkerB() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30 * 1000);

  try {
    const ss = SpreadsheetApp.openById(CONFIG_B.SPREADSHEET_ID);
    const sh = ss.getSheetByName(CONFIG_B.QUEUE_SHEET_NAME);
    if (!sh) throw new Error(`Queue sheet not found: ${CONFIG_B.QUEUE_SHEET_NAME}`);

    const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const colIndex = buildColIndex_(header);

    let processed = 0;
    for (let i = 0; i < CONFIG_B.MAX_ITEMS_PER_RUN; i++) {
      const rowNumber = pickNextRowForB_(sh, colIndex);
      if (!rowNumber) break;

      processOneRowB_(sh, colIndex, rowNumber);
      processed++;
    }

    Logger.log(`[B] processed=${processed}`);
  } finally {
    lock.releaseLock();
  }
}

/** ====== 1行処理 ====== */
function processOneRowB_(sh, colIndex, rowNumber) {
  const row = sh.getRange(rowNumber, 1, 1, sh.getLastColumn()).getValues()[0];
  const get = (key) => row[colIndex[key] - 1];

  const runId = String(get('runId') || '').trim();
  const theme = String(get('theme') || '').trim();

  const attempt = Number(get('attemptB') || 0) + 1;
  const now = new Date();

  // RUNNINGへ
  setCells_(sh, colIndex, rowNumber, {
    statusB: 'RUNNING',
    attemptB: attempt,
    errorB: '',
    updatedAtB: now,
  });

  try {
    // 1) 記事生成（雛形：ここをAI生成に差し替え）
    const { title, html } = generateArticleHtml_({ theme, runId });

    if (!title || title.length < CONFIG_B.MIN_TITLE_LEN) {
      throw new Error(`Generated title too short: "${title}"`);
    }
    if (!html || html.length < CONFIG_B.MIN_BODY_LEN) {
      throw new Error(`Generated body too short: len=${(html || '').length}`);
    }

    // 2) WP投稿
    const postId = createWpPost_({
      title,
      html,
      status: CONFIG_B.WP_POST_STATUS,
      categoryIds: CONFIG_B.WP_CATEGORY_IDS,
      tagIds: CONFIG_B.WP_TAG_IDS,
    });

    // 3) DONEへ
    setCells_(sh, colIndex, rowNumber, {
      articleTitle: title,
      articleHtml: html,
      wpPostId: postId,
      statusB: 'DONE',
      lockUntilB: '',
      updatedAtB: new Date(),
    });

    Logger.log(`[B] DONE runId=${runId} row=${rowNumber} wpPostId=${postId}`);
  } catch (e) {
    const msg = `${e && e.stack ? e.stack : e}`;

    // リトライ判定
    if (attempt >= CONFIG_B.MAX_ATTEMPTS) {
      setCells_(sh, colIndex, rowNumber, {
        statusB: 'ERROR',
        errorB: truncate_(msg, 45000),
        lockUntilB: '',
        updatedAtB: new Date(),
      });
      Logger.log(`[B] ERROR(final) runId=${runId} row=${rowNumber}`);
      return;
    }

    // backoff
    const waitMin = CONFIG_B.RETRY_BACKOFF_MINUTES[Math.min(attempt - 1, CONFIG_B.RETRY_BACKOFF_MINUTES.length - 1)];
    const lockUntil = new Date(Date.now() + waitMin * 60 * 1000);

    setCells_(sh, colIndex, rowNumber, {
      statusB: 'PENDING',
      errorB: truncate_(msg, 45000),
      lockUntilB: lockUntil,
      updatedAtB: new Date(),
    });

    Logger.log(`[B] ERROR(retry) runId=${runId} row=${rowNumber} attempt=${attempt} next=${lockUntil.toISOString()}`);
  }
}

/** ====== 次に処理すべき行を選ぶ ====== */
function pickNextRowForB_(sh, colIndex) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;

  const values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  const now = new Date();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowNumber = i + 2;

    const taskB = row[colIndex.taskB - 1];
    const statusB = String(row[colIndex.statusB - 1] || '').trim();
    const lockUntilB = row[colIndex.lockUntilB - 1];

    // B対象判定：TRUE / "TRUE" / "B" など許容
    const isTarget =
      taskB === true ||
      String(taskB).toUpperCase() === 'TRUE' ||
      String(taskB).toUpperCase() === 'B';

    if (!isTarget) continue;

    // status判定（空は未処理扱い）
    const isPending = (statusB === '' || statusB === 'PENDING');
    if (!isPending) continue;

    // lockUntil 判定（未来ならスキップ）
    if (lockUntilB instanceof Date && lockUntilB.getTime() > now.getTime()) continue;

    return rowNumber;
  }
  return null;
}

/** ====== 記事生成（雛形） ======
 * ここを OpenAI/Gemini 生成に差し替えるだけでOK
 */
function generateArticleHtml_({ theme, runId }) {
  const safeTheme = theme || '未指定テーマ';

  const title = `${safeTheme}｜解説記事`;
  const html =
    `<article>` +
    `<h1>${escapeHtml_(title)}</h1>` +
    `<p>この記事はテーマ「${escapeHtml_(safeTheme)}」をもとに自動生成された雛形です。（runId: ${escapeHtml_(runId)}）</p>` +
    `<h2>要点</h2><ul><li>要点1</li><li>要点2</li><li>要点3</li></ul>` +
    `<h2>本文</h2><p>ここに本文が入ります。AI生成に差し替えて運用してください。</p>` +
    `</article>`;

  return { title, html };
}

/** ====== WordPress 投稿 ====== */
function createWpPost_({ title, html, status, categoryIds, tagIds }) {
  const endpoint = `${CONFIG_B.WP_BASE_URL}/wp-json/wp/v2/posts`;

  const payload = {
    title,
    content: html,
    status: status || 'draft',
  };

  if (Array.isArray(categoryIds) && categoryIds.length) payload.categories = categoryIds;
  if (Array.isArray(tagIds) && tagIds.length) payload.tags = tagIds;

  const options = {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    payload: JSON.stringify(payload),
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(`${CONFIG_B.WP_USERNAME}:${CONFIG_B.WP_APP_PASSWORD}`),
    },
    muteHttpExceptions: true,
  };

  const res = UrlFetchApp.fetch(endpoint, options);
  const code = res.getResponseCode();
  const text = res.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error(`WP API Error: status=${code} body=${truncate_(text, 2000)}`);
  }

  const json = JSON.parse(text);
  if (!json || !json.id) throw new Error(`WP API: missing id. body=${truncate_(text, 2000)}`);
  return json.id;
}

/** ====== ユーティリティ ====== */
function buildColIndex_(headerRow) {
  const map = {};
  headerRow.forEach((name, idx) => {
    map[String(name || '').trim()] = idx + 1;
  });

  // 必須キー検証
  Object.keys(COL).forEach((key) => {
    const colName = COL[key];
    if (!map[colName]) throw new Error(`Missing column "${colName}" in header row.`);
    // COLのkey名で参照できるようにする
    map[key] = map[colName];
  });

  return map;
}

function setCells_(sh, colIndex, rowNumber, obj) {
  const updates = Object.keys(obj).map((k) => ({ k, v: obj[k] }));
  updates.forEach(({ k, v }) => {
    const col = colIndex[k];
    if (!col) return;
    sh.getRange(rowNumber, col).setValue(v);
  });
}

function truncate_(s, maxLen) {
  const str = String(s || '');
  return str.length > maxLen ? str.slice(0, maxLen) + '…' : str;
}

function escapeHtml_(s) {
  return String(s || '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

// ===== ワーカー共通設定 =====
const WORKER_CONFIG = {
  SPREADSHEET_ID: "1j5bCNucxL9QVS_iq_RYaeL1_vWTsxHArUSJuRzhQmiw",
  QUEUE_SHEET_NAME: "_QUEUE",
  MAX_PER_RUN: 3,              // 1回の実行で処理する最大件数
  LOCK_SECONDS: 30,
};

// ★A/B/Cごとに変える
const JOB_TYPE = "C_SLIDES_GEN"; // Aは "A_HTML_GITHUB", Bは "B_BLOG_WP"

/**
 * 初回だけ手動実行：時間トリガーを作る（例：1分毎）
 */
function initWorkerTrigger() {
  // 既存の重複トリガーを削除
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "runWorker") ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger("runWorker")
    .timeBased()
    .everyMinutes(1)
    .create();

  Logger.log("✅ Worker trigger created: runWorker every 1 minute");
}

/**
 * 時間主導でキューを消化する
 */
function runWorker() {
  const lock = LockService.getScriptLock();
  lock.waitLock(WORKER_CONFIG.LOCK_SECONDS * 1000);

  try {
    const ss = SpreadsheetApp.openById(WORKER_CONFIG.SPREADSHEET_ID);
    const q = ss.getSheetByName(WORKER_CONFIG.QUEUE_SHEET_NAME);
    if (!q) throw new Error("Queue sheet not found.");

    const values = q.getDataRange().getValues();
    if (values.length <= 1) return;

    const header = values[0];
    const idx = indexMap_(header);

    let processed = 0;

    for (let r = 1; r < values.length; r++) {
      if (processed >= WORKER_CONFIG.MAX_PER_RUN) break;

      const row = values[r];
      const jobType = row[idx.jobType];
      const status  = row[idx.status];

      if (jobType !== JOB_TYPE) continue;
      if (status !== "PENDING") continue;

      const sheetRow = r + 1; // シート上の行番号

      // RUNNINGへ
      updateQueueRow_(q, sheetRow, idx, { status: "RUNNING", updatedAt: new Date(), lastError: "" });

      try {
        const payload = JSON.parse(row[idx.payloadJson] || "{}");

        // ★ここに本処理を差し込む（今はダミー）
        doJob_(payload);

        updateQueueRow_(q, sheetRow, idx, { status: "DONE", updatedAt: new Date() });
      } catch (err) {
        const retry = Number(row[idx.retryCount] || 0) + 1;
        updateQueueRow_(q, sheetRow, idx, {
          status: "ERROR",
          retryCount: retry,
          lastError: String(err && err.stack ? err.stack : err),
          updatedAt: new Date(),
        });
      }

      processed++;
    }

    Logger.log(`✅ runWorker done. jobType=${JOB_TYPE} processed=${processed}`);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 本処理（まずはダミーでOK）
 * payload.namedValues などからテーマを取り出して処理する想定
 */
function doJob_(payload) {
  // 例：フォームの設問「テーマ」を取り出す（存在しない場合もあるので注意）
  const nv = payload.namedValues || {};
  const themeArr = nv["テーマ"] || nv["theme"] || [];
  const theme = Array.isArray(themeArr) ? themeArr[0] : "";

  Logger.log(`[${JOB_TYPE}] doing job... theme=${theme}`);
  // TODO: A/B/C の実処理に置き換える
}

/** ヘッダ名→列index */
function indexMap_(header) {
  const map = {};
  header.forEach((h, i) => map[String(h).trim()] = i);

  const required = ["jobType","status","payloadJson","retryCount","lastError","updatedAt"];
  required.forEach(k => {
    if (!(k in map)) throw new Error(`Queue header missing: ${k}`);
  });

  return map;
}

/** キュー行を部分更新 */
function updateQueueRow_(sheet, rowNum, idx, patch) {
  const now = patch.updatedAt || new Date();

  if (patch.status != null) sheet.getRange(rowNum, idx.status + 1).setValue(patch.status);
  if (patch.retryCount != null) sheet.getRange(rowNum, idx.retryCount + 1).setValue(patch.retryCount);
  if (patch.lastError != null) sheet.getRange(rowNum, idx.lastError + 1).setValue(patch.lastError);
  sheet.getRange(rowNum, idx.updatedAt + 1).setValue(now);
}
