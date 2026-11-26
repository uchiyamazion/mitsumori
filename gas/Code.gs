// 見積作成 — GAS バックエンド（スプレッドシートにバインドして使用）
const TOKEN = 'CHANGE_ME'; // アクセスキー（必ず変更）
const SEED = { E: [], G: [], T: [] }; // 価格表はリポジトリに含めない（setup後に各シートへ貼り付け）
const TZ = 'Asia/Tokyo';
const S_E = '単価_EHP', S_G = '単価_GHP', S_T = '出張料', S_Q = '見積';

// 初回のみ実行：シート作成と価格表の初期投入
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mk = (name, head, rows) => {
    if (ss.getSheetByName(name)) return;
    const s = ss.insertSheet(name);
    s.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    if (rows.length) s.getRange(2, 1, rows.length, head.length).setValues(rows);
    s.setFrozenRows(1);
  };
  const h = ['コード', '種別', '作業内容', '定価'];
  mk(S_E, h, SEED.E);
  mk(S_G, h, SEED.G);
  mk(S_T, ['エリア', '出張料'], SEED.T);
  mk(S_Q, ['見積No', '更新日時', '宛先', '件名', '税込合計', 'データ'], []);
  ss.getSheetByName(S_Q).getRange('A:B').setNumberFormat('@');
}

const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const sh = n => SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n);
function rows(n, c) { const s = sh(n), l = s.getLastRow(); return l < 2 ? [] : s.getRange(2, 1, l - 1, c).getValues(); }
function findRow(s, no) {
  const l = s.getLastRow(); if (l < 2) return 0;
  const v = s.getRange(2, 1, l - 1, 1).getValues();
  for (let i = 0; i < v.length; i++) if (String(v[i][0]) === no) return i + 2;
  return 0;
}
const price = r => [String(r[0]), String(r[1]), String(r[2]), Number(r[3]) || 0];

function doGet(e) {
  const __g = authGuard_(e); if (__g) return __g;
  try {
    const p = e.parameter;
    if (p.token !== TOKEN) return out({ error: 'unauthorized' });
    if (p.action === 'master') return out({
      E: rows(S_E, 4).filter(r => r[2] !== '').map(price),
      G: rows(S_G, 4).filter(r => r[2] !== '').map(price),
      T: rows(S_T, 2).filter(r => r[0] !== '').map(r => [String(r[0]), Number(r[1]) || 0]),
    });
    if (p.action === 'list') return out(rows(S_Q, 5)
      .map(r => ({ no: String(r[0]), at: String(r[1]), to: r[2], title: r[3], total: r[4] }))
      .sort((a, b) => b.at.localeCompare(a.at)));
    if (p.action === 'get') {
      const s = sh(S_Q), r = findRow(s, p.no);
      return out(r ? JSON.parse(s.getRange(r, 6).getValue()) : { error: 'not found' });
    }
    if (p.action === 'nextno') {
      const d = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd');
      const n = rows(S_Q, 1).filter(r => String(r[0]).startsWith(d)).length + 1;
      return out({ no: d + '-' + ('0' + n).slice(-2) });
    }
    return out({ error: 'bad action' });
  } catch (x) { return out({ error: String(x) }); }
}

// フロントは Content-Type: text/plain で送信（CORSプリフライト回避）
function doPost(e) {
  const __g = authGuard_(e); if (__g) return __g;
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.token !== TOKEN) return out({ error: 'unauthorized' });
    const lock = LockService.getScriptLock(); lock.waitLock(10000);
    try {
      const s = sh(S_Q), r = findRow(s, String(b.no));
      if (b.action === 'save') {
        const row = [String(b.no), Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), b.to || '', b.title || '', Number(b.total) || 0, JSON.stringify(b.data)];
        r ? s.getRange(r, 1, 1, 6).setValues([row]) : s.appendRow(row);
        return out({ ok: 1 });
      }
      if (b.action === 'delete') { if (r) s.deleteRow(r); return out({ ok: 1 }); }
      return out({ error: 'bad action' });
    } finally { lock.releaseLock(); }
  } catch (x) { return out({ error: String(x) }); }
}

// ════════════════════════════════════════════════
// AuthGuard.gs ― Googleログインの本人確認（GAS側）
// 画面から送られたGoogleのIDトークンを確認し、許可リストにあるアカウントだけを通す。
//
// 【設定】Apps Scriptの「プロジェクトの設定 → スクリプト プロパティ」に
//   ALLOWED_EMAILS = a@gmail.com,b@gmail.com（カンマ区切り）を登録する。未設定のときは全員を拒否する。
// 【使い方】doGet(e) と doPost(e) の先頭に、次の2行を入れる。
//   const g = authGuard_(e); if (g) return g;
// ════════════════════════════════════════════════

const GOOGLE_CLIENT_ID = '782190614995-5d4fur401smd1gpnnvs43h3m2q2q2ar3.apps.googleusercontent.com';

function allowedEmails_() {
  const v = PropertiesService.getScriptProperties().getProperty('ALLOWED_EMAILS') || '';
  return v.split(/[\s,;]+/).map(s => s.trim().toLowerCase()).filter(Boolean);
}

// 戻り値：{ ok:true, email } または { ok:false, reason:'unauthorized' | 'forbidden' }
function verifyIdToken_(idToken) {
  if (!idToken) return { ok: false, reason: 'unauthorized' };
  const cache = CacheService.getScriptCache();
  const key = 'idt_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken)).slice(0, 40);
  let info = null;
  const hit = cache.get(key);
  if (hit) { try { info = JSON.parse(hit); } catch (e) { info = null; } }
  if (!info) {
    const res = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
      { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return { ok: false, reason: 'unauthorized' };
    const r = JSON.parse(res.getContentText());
    info = { aud: r.aud, iss: r.iss, email: r.email, email_verified: r.email_verified, exp: r.exp };
    const ttl = Math.min(300, Number(info.exp) - Math.floor(Date.now() / 1000));
    if (ttl > 0) cache.put(key, JSON.stringify(info), ttl);
  }
  const now = Math.floor(Date.now() / 1000);
  const iss = String(info.iss || '');
  if (info.aud !== GOOGLE_CLIENT_ID) return { ok: false, reason: 'unauthorized' };
  if (iss !== 'accounts.google.com' && iss !== 'https://accounts.google.com') return { ok: false, reason: 'unauthorized' };
  if (Number(info.exp) <= now) return { ok: false, reason: 'unauthorized' };
  if (String(info.email_verified) !== 'true') return { ok: false, reason: 'unauthorized' };
  const email = String(info.email || '').toLowerCase();
  if (allowedEmails_().indexOf(email) < 0) return { ok: false, reason: 'forbidden' };
  return { ok: true, email: email };
}

// 許可されていれば null、そうでなければエラー応答（JSON）を返す
function authGuard_(e) {
  let token = (e && e.parameter && e.parameter.idToken) || '';
  if (!token && e && e.postData && e.postData.contents) {
    try { token = (JSON.parse(e.postData.contents) || {}).idToken || ''; } catch (err) { /* JSON以外は無視 */ }
  }
  const a = verifyIdToken_(token);
  if (a.ok) return null;
  return ContentService.createTextOutput(JSON.stringify({
    status: 'error', success: false, ok: false, error: a.reason, data: a.reason, message: a.reason
  })).setMimeType(ContentService.MimeType.JSON);
}
