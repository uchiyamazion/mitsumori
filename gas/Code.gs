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
