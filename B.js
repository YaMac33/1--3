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
