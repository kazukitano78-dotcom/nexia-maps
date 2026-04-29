// ════════════════════════════════════════════════════════════════
// NEXIA 営業マップ - Google Apps Script バックエンド v3
// ════════════════════════════════════════════════════════════════
// 【デプロイ手順】
//   1. スプレッドシート → 拡張機能 → Apps Script
//   2. 全選択(Ctrl+A)して削除 → このコードを貼り付けて保存(Ctrl+S)
//   3. デプロイ → 新しいデプロイ
//      種類: ウェブアプリ / 実行: 自分 / アクセス: 全員
//   4. 新しいURLをindex.htmlのGAS_URLに貼り付け
// ════════════════════════════════════════════════════════════════

const SPREADSHEET_ID = '1gcD6NIxH3U7f4srEIrCtgQavV7-qAct_U7H3wwQXHMw';
const SHEET_NAME = 'シート1';
const SESSION_TTL = 12 * 60 * 60; // 12時間

// ── 許可するGoogleアカウント（小文字で統一）──
const ALLOWED_EMAILS = [
  'kazuki.tano78@gmail.com',
];

function doGet(e) {
  const callback = e.parameter && e.parameter.callback;

  function respond(result) {
    const json = JSON.stringify(result);
    if (callback) {
      const out = ContentService.createTextOutput(callback + '(' + json + ')');
      out.setMimeType(ContentService.MimeType.JAVASCRIPT);
      return out;
    }
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  }

  const action = (e.parameter && e.parameter.action) || '';

  // ログインだけidToken検証（1回だけ・短いセッションIDに交換）
  if (action === 'login') {
    try {
      const email = verifyGoogleToken(e.parameter.idToken);
      const session = createSession(email);
      return respond({ ok: true, session: session, email: email });
    } catch(err) {
      return respond({ error: 'AUTH_ERROR: ' + err.message });
    }
  }

  // 以降は短いセッションIDで認証（URLが短くなりタイムアウトしない）
  try {
    checkSession(e.parameter && e.parameter.s);
  } catch(err) {
    return respond({ error: 'AUTH_ERROR: ' + err.message });
  }

  let result;
  try {
    if (action === 'getMapPoints') {
      result = getMapPoints();
    } else if (action === 'getAllDetails') {
      result = getAllDetails();
    } else if (action === 'getCompanyDetail') {
      result = getCompanyDetail(parseInt(e.parameter.row));
    } else if (action === 'update') {
      result = updateCompany(parseInt(e.parameter.row), JSON.parse(e.parameter.data));
    } else {
      result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }
  return respond(result);
}

// ──────────────────────────────────────────
// POST ハンドラ（update専用 - URL長さ制限回避）
// ──────────────────────────────────────────
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Invalid JSON body' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  function respond(result) {
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const action = body.action || '';
  try {
    checkSession(body.s);
  } catch(err) {
    return respond({ error: 'AUTH_ERROR: ' + err.message });
  }

  let result;
  try {
    if (action === 'update') {
      result = updateCompany(parseInt(body.row), body.data);
    } else {
      result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }
  return respond(result);
}

// ──────────────────────────────────────────
// Google IDトークン検証（JWTペイロードデコード）
// ──────────────────────────────────────────
function verifyGoogleToken(idToken) {
  if (!idToken || idToken === 'null') throw new Error('トークンがありません');
  const parts = idToken.split('.');
  if (parts.length !== 3) throw new Error('不正なトークン形式');
  const b64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
  const padded = b64 + '='.repeat((4 - b64.length % 4) % 4);
  const payload = JSON.parse(
    Utilities.newBlob(Utilities.base64Decode(padded)).getDataAsString('UTF-8')
  );
  if (!payload.email) throw new Error('メール取得失敗');
  if (!['accounts.google.com', 'https://accounts.google.com'].includes(payload.iss))
    throw new Error('不正な発行者');
  if (payload.exp < Date.now() / 1000) throw new Error('トークン期限切れ');
  return payload.email;
}

// ──────────────────────────────────────────
// セッション作成・検証
// ──────────────────────────────────────────
function createSession(email) {
  if (!ALLOWED_EMAILS.includes(email.toLowerCase())) {
    throw new Error('アクセス権限がありません: ' + email);
  }
  const id = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  const exp = Math.floor(Date.now() / 1000) + SESSION_TTL;
  PropertiesService.getScriptProperties().setProperty('s_' + id, email + ':' + exp);
  return id;
}

function checkSession(s) {
  if (!s) throw new Error('セッションがありません');
  const val = PropertiesService.getScriptProperties().getProperty('s_' + s);
  if (!val) throw new Error('無効なセッションです');
  const exp = parseInt(val.split(':')[1]);
  if (exp < Date.now() / 1000) throw new Error('セッション期限切れです。再ログインしてください。');
}

// ──────────────────────────────────────────
// 地図表示用 最小データのみ返す（CacheService付き）
// ──────────────────────────────────────────
function getMapPoints() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('map_points');
  if (cached) return JSON.parse(cached);

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];
  const data    = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headers = data[0].map(h => String(h).trim());
  const idx = {
    la:   headers.indexOf('緯度'),
    lo:   headers.indexOf('経度'),
    n:    headers.indexOf('会社名'),
    f:    headers.indexOf('訪問フラグ'),
    res:  headers.indexOf('訪問結果'),
    l:    headers.indexOf('見込み度'),
    pid:  headers.indexOf('place_id'),
    cust: headers.indexOf('顧客フラグ'),
    bf:   headers.indexOf('ビジネスフォン'),
    mfp:  headers.indexOf('複合機利用'),
    utm:  headers.indexOf('UTM利用'),
    hub:  headers.indexOf('HUB利用'),
    srv:  headers.indexOf('サーバー利用'),
    ml:   headers.indexOf('マイリスト'),
  };
  const points = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const la = parseFloat(data[i][idx.la]);
    const lo = parseFloat(data[i][idx.lo]);
    if (!la || !lo || isNaN(la) || isNaN(lo)) continue;
    points.push({
      r:   i + 1, la, lo,
      n:   idx.n   >= 0 ? data[i][idx.n]   : '',
      f:   idx.f   >= 0 ? data[i][idx.f]   : '',
      res: idx.res >= 0 ? data[i][idx.res] : '',
      l:   idx.l   >= 0 ? data[i][idx.l]   : '',
      pid:  idx.pid  >= 0 ? data[i][idx.pid]  : '',
      cust: idx.cust >= 0 ? data[i][idx.cust] : '',
      bf:   idx.bf   >= 0 ? data[i][idx.bf]   : '',
      mfp:  idx.mfp  >= 0 ? data[i][idx.mfp]  : '',
      utm:  idx.utm  >= 0 ? data[i][idx.utm]  : '',
      hub:  idx.hub  >= 0 ? data[i][idx.hub]  : '',
      srv:  idx.srv  >= 0 ? data[i][idx.srv]  : '',
      ml:   idx.ml   >= 0 ? data[i][idx.ml]   : '',
    });
  }
  try { cache.put('map_points', JSON.stringify(points), 30); } catch(e) {}
  return points;
}

// ──────────────────────────────────────────
// 全社詳細データ一括取得（バックグラウンド用）
// ──────────────────────────────────────────
function getAllDetails() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];
  const data    = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headers = data[0].map(h => String(h).trim());
  const companies = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const obj = { _row: i + 1 };
    headers.forEach((h, j) => { obj[h] = data[i][j] || ''; });
    companies.push(obj);
  }
  return companies;
}

// ──────────────────────────────────────────
// ピンクリック時: 1行分の全データを返す
// ──────────────────────────────────────────
function getCompanyDetail(row) {
  if (!row || row < 2) throw new Error('無効な行番号: ' + row);
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0]
    .map(h => String(h).trim());
  const rowData = sheet.getRange(row, 1, 1, lastCol).getDisplayValues()[0];
  const obj = { _row: row };
  headers.forEach((h, j) => { obj[h] = rowData[j] || ''; });
  return obj;
}

// ──────────────────────────────────────────
// 保存: 該当1行のみ更新
// ──────────────────────────────────────────
function updateCompany(row, data) {
  if (!row || row < 2) throw new Error('行番号が無効: ' + row);
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(h => String(h).trim());

  // 行データを一括取得して更新（setValue個別呼び出しを排除）
  const rowVals = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  Object.keys(data).forEach(key => {
    if (key.startsWith('_')) return;
    const colIdx = headers.indexOf(key);
    if (colIdx >= 0) rowVals[colIdx] = data[key] || '';
  });
  const updateIdx = headers.indexOf('最終更新日');
  if (updateIdx >= 0) {
    rowVals[updateIdx] = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  }
  // 1回のsetValuesで書き込み（高速化）
  sheet.getRange(row, 1, 1, lastCol).setValues([rowVals]);

  // キャッシュを削除（次回リロード時に最新データを取得させる）
  try { CacheService.getScriptCache().remove('map_points'); } catch(e) {}
  return { success: true, row: row };
}
