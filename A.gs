/**
 * =========================================
 * A: HTML生成 → GitHub Pages反映（Aのみ）
 * 親GASの_QUEUE仕様に準拠
 * =========================================
 *
 * Queue schema（親GAS）:
 *  A: jobId
 *  B: createdAt
 *  C: jobType
 *  D: status      // PENDING / RUNNING / DONE / ERROR
 *  E: sourceSheet
 *  F: sourceRow
 *  G: payloadJson
 *  H: retryCount
 *  I: lastError
 *  J: updatedAt
 */

// ===== A側 設定 =====
const A_CONFIG = {
  // 親GASのCONFIGを流用（同一プロジェクトなら参照できます）
  // SPREADSHEET_ID / QUEUE_SHEET_NAME は親CONFIGに合わせる
  JOB_TYPE: "A_HTML_GITHUB",

  // フォーム回答のどの「質問」をテーマとして使うか（namedValuesのキー）
  // 例：フォーム質問が「テーマ」なら "テーマ"
  THEME_QUESTION_KEY: "テーマ",

  // GitHub Pages（コミット先）
  GITHUB: {
    OWNER: "YaMac33",
    REPO: "1--3",
    BRANCH: "main",

    // Script Propertiesに入れるキー名
    TOKEN_PROP_KEY: "GITHUB_TOKEN",

    // ページURL（結果として記録したい場合用。Queueに保存する場合は拡張で可能）
    // ★ここは自分のPages URLに置き換え
    PAGES_URL: "https://PUT_OWNER.github.io/PUT_REPO/",
  },

  // ワーカーの挙動
  WORKER: {
    BATCH_SIZE: 5,           // 1回で処理する件数
    MAX_RETRY: 5,            // retryCountがこれ以上なら諦める
    BACKOFF_MIN: [1, 3, 10, 30, 120], // retryCountに応じた待機（分）
  },
};

/**
 * Aワーカー（時間主導トリガー推奨）
 * - _QUEUEから A_HTML_GITHUB の PENDING/ERROR を拾う
 * - retryCount が上限超えたものはスキップ
 *
 * ※ 今は runWorker() を使ってるはずですが、
 *   こちらも残すなら CONFIG 参照を排除しておく（事故防止）
 */
function runAWorker_() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30 * 1000);

  try {
    // ★CONFIG ではなく WORKER_CONFIG を参照
    const ss = SpreadsheetApp.openById(WORKER_CONFIG.SPREADSHEET_ID);
    const q = ss.getSheetByName(WORKER_CONFIG.QUEUE_SHEET_NAME);
    if (!q) throw new Error("Queue sheet not found. Run initQueueAndTrigger() first.");

    const lastRow = q.getLastRow();
    if (lastRow < 2) return;

    const values = q.getRange(2, 1, lastRow - 1, 10).getValues(); // 2行目以降
    const now = new Date();
    let processed = 0;

    for (let i = 0; i < values.length; i++) {
      if (processed >= A_CONFIG.WORKER.BATCH_SIZE) break;

      const rowIndex = i + 2; // シート行番号

      const jobId = values[i][0];
      const jobType = values[i][2];
      const status = values[i][3];
      const payloadJson = values[i][6];
      const retryCount = Number(values[i][7] || 0);
      const updatedAt = values[i][9] ? new Date(values[i][9]) : null;

      // A以外はスキップ（B/Cは触らない）
      if (jobType !== A_CONFIG.JOB_TYPE) continue;

      // DONEはスキップ
      if (status === "DONE") continue;

      // retry上限超えはスキップ
      if (retryCount >= A_CONFIG.WORKER.MAX_RETRY) continue;

      // ERRORのリトライ待ち（updatedAt + backoff）
      if (status === "ERROR" && updatedAt) {
        const waitMin = getBackoffMin_(retryCount);
        const nextTime = new Date(updatedAt.getTime() + waitMin * 60 * 1000);
        if (nextTime > now) continue;
      }

      // RUNNINGはスキップ（別実行が動いてる想定）
      if (status === "RUNNING") continue;

      // RUNNINGへ更新（先に状態変更して多重実行を減らす）
      updateQueueStatus_(q, rowIndex, "RUNNING", retryCount, "", now);

      try {
        const payload = JSON.parse(payloadJson || "{}");

        // テーマ抽出
        const theme = extractThemeFromPayload_(payload);

        // HTML生成 → GitHub反映
        const resultUrl = executeAHtmlGithub_({ jobId, theme, payload });

        // DONEへ
        updateQueueStatus_(q, rowIndex, "DONE", retryCount, "", new Date());

        Logger.log(`✅ A DONE: jobId=${jobId} url=${resultUrl}`);
        processed++;

      } catch (err) {
        const newRetry = retryCount + 1;
        updateQueueStatus_(q, rowIndex, "ERROR", newRetry, stringifyErr_(err), new Date());
        Logger.log(`❌ A ERROR: jobId=${jobId} retry=${newRetry} err=${stringifyErr_(err)}`);
        processed++;
      }
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * A本体：HTML生成 → GitHubへコミット
 */
function executeAHtmlGithub_({ jobId, theme, payload }) {
  const html = generateHtmlFromTheme_(theme, payload);

  const filePath = buildFilePathFromJobId_(jobId);

  upsertFileToGitHub_(
    A_CONFIG.GITHUB.OWNER,
    A_CONFIG.GITHUB.REPO,
    A_CONFIG.GITHUB.BRANCH,
    filePath,
    html,
    `A_HTML_GITHUB: create ${filePath} (jobId=${jobId})`
  );

  const pageUrl =
    A_CONFIG.GITHUB.PAGES_URL + filePath.replace(/^docs\//, "").replace(/index\.html$/, "");

  // ★docs/index.html を _QUEUE から再生成して更新
  updateDocsIndexFromQueue_();

  return pageUrl;
}

// ===== HTML生成（まずはスタブ。後でAIに差し替え） =====
function generateHtmlFromTheme_(theme, payload) {
  const t = String(theme || "").trim() || "Untitled";

  return `<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>${escapeHtml_(t)}</title>
  <style>
    body{font-family:system-ui,-apple-system,"Segoe UI",Roboto,"Noto Sans JP",sans-serif;margin:40px;line-height:1.7;}
    .card{max-width:900px;padding:28px;border:1px solid #e5e7eb;border-radius:16px;}
    .meta{color:#6b7280;font-size:14px;margin-bottom:10px;}
  </style>
</head>
<body>
  <div class="card">
    <div class="meta">Generated by GAS Queue Worker (A only)</div>
    <h1>${escapeHtml_(t)}</h1>
    <p>このページはフォーム回答から自動生成され、GitHub Pagesに反映されました。</p>
    <p><b>次の拡張候補</b>：テーマ別にファイル名を分ける / 履歴を残す / AIでセクション生成 など。</p>
  </div>
</body>
</html>`;
}

// ===== payload からテーマ抽出 =====
function extractThemeFromPayload_(payload) {
  // payload.namedValues は { "質問": ["回答"] }
  const nv = payload && payload.namedValues ? payload.namedValues : null;
  if (nv && nv[A_CONFIG.THEME_QUESTION_KEY] && nv[A_CONFIG.THEME_QUESTION_KEY][0]) {
    return String(nv[A_CONFIG.THEME_QUESTION_KEY][0]).trim();
  }

  // fallback: 取れない場合はsourceRow等からも取りに行ける（必要なら拡張）
  return "";
}

// ===== Queue更新（親GASの列に合わせる） =====
function updateQueueStatus_(queueSheet, rowIndex, status, retryCount, lastError, updatedAt) {
  // D:status, H:retryCount, I:lastError, J:updatedAt
  queueSheet.getRange(rowIndex, 4).setValue(status);
  queueSheet.getRange(rowIndex, 8).setValue(retryCount);
  queueSheet.getRange(rowIndex, 9).setValue(lastError || "");
  queueSheet.getRange(rowIndex, 10).setValue(updatedAt || new Date());
}

// ===== backoff =====
function getBackoffMin_(retryCount) {
  // retryCount: 1,2,3... を想定
  const idx = Math.max(0, Math.min(retryCount - 1, A_CONFIG.WORKER.BACKOFF_MIN.length - 1));
  return A_CONFIG.WORKER.BACKOFF_MIN[idx];
}

function stringifyErr_(err) {
  return String(err && err.stack ? err.stack : err);
}

// ===== GitHub Contents API: upsert =====
function upsertFileToGitHub_(owner, repo, branch, path, contentText, message) {
  const token = PropertiesService.getScriptProperties().getProperty(A_CONFIG.GITHUB.TOKEN_PROP_KEY);
  if (!token) throw new Error(`Missing GitHub token in Script Properties: ${A_CONFIG.GITHUB.TOKEN_PROP_KEY}`);

  const apiBase = `https://api.github.com/repos/${owner}/${repo}/contents/${encodeURIComponent(path).replace(/%2F/g,'/')}`;
  const headers = {
    Authorization: `token ${token}`,
    "User-Agent": "gas-queue-a-worker",
    Accept: "application/vnd.github+json",
  };

  // 既存sha取得（なければ404）
  let currentSha = null;
  {
    const url = `${apiBase}?ref=${encodeURIComponent(branch)}`;
    const res = UrlFetchApp.fetch(url, { method: "get", headers, muteHttpExceptions: true });
    const code = res.getResponseCode();
    if (code === 200) {
      const json = JSON.parse(res.getContentText());
      currentSha = json.sha;
    } else if (code === 404) {
      currentSha = null;
    } else {
      throw new Error(`GitHub GET contents failed: ${code} ${res.getContentText()}`);
    }
  }

  // PUT（作成/更新）
  const body = {
    message,
    content: Utilities.base64Encode(contentText, Utilities.Charset.UTF_8),
    branch,
  };
  if (currentSha) body.sha = currentSha;

  const put = UrlFetchApp.fetch(apiBase, {
    method: "put",
    headers,
    contentType: "application/json",
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });

  const putCode = put.getResponseCode();
  if (putCode !== 200 && putCode !== 201) {
    throw new Error(`GitHub PUT contents failed: ${putCode} ${put.getContentText()}`);
  }

  const putJson = JSON.parse(put.getContentText());
  return (putJson.commit && putJson.commit.sha) ? putJson.commit.sha : "(no sha)";
}

function escapeHtml_(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

// ===== ワーカー共通設定 =====
const WORKER_CONFIG = {
  SPREADSHEET_ID: "1j5bCNucxL9QVS_iq_RYaeL1_vWTsxHArUSJuRzhQmiw",
  QUEUE_SHEET_NAME: "_QUEUE",
  MAX_PER_RUN: 3,              // 1回の実行で処理する最大件数
  LOCK_SECONDS: 30,
};

// ★A/B/Cごとに変える
const JOB_TYPE = "A_HTML_GITHUB"; // Bは "B_BLOG_WP", Cは "C_SLIDES_GEN"

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

        const jobId = row[idx.jobId]; // jobId 列を使う
        doJob_(payload, jobId);

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

function doJob_(payload, jobId) {
  const theme = getThemeFromPayload_(payload);

  // 既存本体を呼ぶ
  executeAHtmlGithub_({ jobId, theme, payload });
}

/** ヘッダ名→列index */
function indexMap_(header) {
  const map = {};
  header.forEach((h, i) => map[String(h).trim()] = i);

  const required = ["jobId","jobType","status","payloadJson","retryCount","lastError","updatedAt"];
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

function getThemeFromPayload_(payload) {
  const nv = (payload && payload.namedValues) ? payload.namedValues : {};
  const v = nv["テーマ"];

  const theme = Array.isArray(v) ? String(v[0] || "").trim() : String(v || "").trim();

  if (!theme) {
    throw new Error("テーマが空です。フォーム項目名が「テーマ」か、payload.namedValues に入っているか確認してください。");
  }
  return theme;
}

/**
 * jobId → GitHub Pages用のファイルパスを作る
 * Pages = docs 配下 前提
 */
function buildFilePathFromJobId_(jobId) {
  const safe = String(jobId).replace(/[^a-zA-Z0-9-_]/g, "_");
  return `docs/${safe}/index.html`;
}

/**
 * _QUEUE を元に docs/index.html を再生成して GitHub に反映
 * - A_HTML_GITHUB の DONE のみ
 * - 新しい順（createdAt降順）
 */
function updateDocsIndexFromQueue_() {
  // ★CONFIG ではなく WORKER_CONFIG を参照（ここが今回の修正の肝）
  const ss = SpreadsheetApp.openById(WORKER_CONFIG.SPREADSHEET_ID);
  const q = ss.getSheetByName(WORKER_CONFIG.QUEUE_SHEET_NAME);
  if (!q) throw new Error("Queue sheet not found.");

  const values = q.getDataRange().getValues();
  if (values.length <= 1) return;

  const header = values[0];
  const idx = indexMapForQueue_(header);

  // DONE の A だけ集める
  const items = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (row[idx.jobType] !== A_CONFIG.JOB_TYPE) continue;
    if (row[idx.status] !== "DONE") continue;

    const jobId = String(row[idx.jobId] || "");
    const createdAt = row[idx.createdAt];
    const payload = safeJsonParse_(row[idx.payloadJson]);

    const theme = extractThemeFromPayload_(payload) || "(no theme)";
    const filePath = buildFilePathFromJobId_(jobId);
    const pageUrl =
      A_CONFIG.GITHUB.PAGES_URL + filePath.replace(/^docs\//, "").replace(/index\.html$/, "");

    items.push({
      jobId,
      theme,
      createdAt: createdAt ? new Date(createdAt) : null,
      pageUrl,
    });
  }

  // 新しい順
  items.sort((a, b) => (b.createdAt?.getTime?.() || 0) - (a.createdAt?.getTime?.() || 0));

  const indexHtml = renderJobsIndexHtml_(items);

  upsertFileToGitHub_(
    A_CONFIG.GITHUB.OWNER,
    A_CONFIG.GITHUB.REPO,
    A_CONFIG.GITHUB.BRANCH,
    "docs/index.html",
    indexHtml,
    `Update docs/index.html (items=${items.length})`
  );
}

/** _QUEUE ヘッダの index を取得（親GASの列に合わせる） */
function indexMapForQueue_(header) {
  const map = {};
  header.forEach((h, i) => (map[String(h).trim()] = i));

  const required = ["jobId", "createdAt", "jobType", "status", "payloadJson"];
  required.forEach(k => {
    if (!(k in map)) throw new Error(`Queue header missing: ${k}`);
  });

  return {
    jobId: map.jobId,
    createdAt: map.createdAt,
    jobType: map.jobType,
    status: map.status,
    payloadJson: map.payloadJson,
  };
}

function safeJsonParse_(s) {
  try { return JSON.parse(s || "{}"); } catch (e) { return {}; }
}

/** 一覧HTML */
function renderJobsIndexHtml_(items) {
  const rows = items.map(it => {
    const dateStr = it.createdAt
      ? Utilities.formatDate(it.createdAt, "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss")
      : "";
    return `
      <tr>
        <td>${escapeHtml_(dateStr)}</td>
        <td>${escapeHtml_(it.theme)}</td>
        <td><a href="${escapeHtml_(it.pageUrl)}" target="_blank" rel="noopener">Open</a></td>
        <td><code>${escapeHtml_(it.jobId)}</code></td>
      </tr>`;
  }).join("");

  return `<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Jobs Index</title>
  <style>
    body{font-family:system-ui,-apple-system,"Segoe UI",Roboto,"Noto Sans JP",sans-serif;margin:32px;line-height:1.6;}
    table{border-collapse:collapse;width:100%;max-width:1200px;}
    th,td{border:1px solid #e5e7eb;padding:10px;vertical-align:top;}
    th{background:#f8fafc;text-align:left;}
    code{font-size:12px;}
    .meta{color:#6b7280;font-size:13px;margin-bottom:12px;}
  </style>
</head>
<body>
  <h1>Jobs Index</h1>
  <div class="meta">Generated from _QUEUE (A_HTML_GITHUB / DONE). Count: ${items.length}</div>
  <table>
    <thead>
      <tr>
        <th>Created</th>
        <th>Theme</th>
        <th>URL</th>
        <th>JobId</th>
      </tr>
    </thead>
    <tbody>
      ${rows || `<tr><td colspan="4">No DONE jobs yet.</td></tr>`}
    </tbody>
  </table>
</body>
</html>`;
}
