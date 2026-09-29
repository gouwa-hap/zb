/**
 * 工作台本地服务器（零依赖，仅用 Node 内置模块）
 * --------------------------------------------------
 * 作用：
 *   1. 托管当前目录的静态文件（index.html / assets / plugins），
 *      让工作台通过 http://localhost:PORT 访问（同源）。
 *   2. 提供 /api/fetch?url=XXX 代理端点：由服务端去抓取目标网页，
 *      再把内容回传。这样浏览器请求的是同源的 localhost，
 *      政府网站的跨域限制从根上被绕开，且数据不经过任何第三方。
 *
 * 启动： node server.js   （或双击「启动工作台.bat」）
 * 默认端口 8080，可用环境变量覆盖： PORT=9000 node server.js
 */
'use strict';
const http = require('http');
const fs = require('fs');
const path = require('path');
const { URL } = require('url');

// 同步到本地 Excel 用的 SheetJS（浏览器版构建，禁用了 readFile，
// 因此始终用 fs.readFileSync + XLSX.read(buffer) 方式读写）
const XLSX = require(path.join(__dirname, 'assets', 'vendor', 'xlsx.full.min.js'));

const ROOT = __dirname;
const PORT = process.env.PORT || 8080;

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.csv': 'text/csv; charset=utf-8',
  '.md': 'text/markdown; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.ico': 'image/x-icon',
};

function serveStatic(req, res) {
  let pathname = decodeURIComponent(new URL(req.url, 'http://localhost').pathname);
  if (pathname === '/') pathname = '/index.html';
  const filePath = path.join(ROOT, pathname);
  // 防止路径穿越
  if (!filePath.startsWith(ROOT)) {
    res.writeHead(403, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('Forbidden');
    return;
  }
  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('404 Not Found: ' + pathname);
      return;
    }
      const ext = path.extname(filePath).toLowerCase();
      res.writeHead(200, {
        'Content-Type': MIME[ext] || 'application/octet-stream',
        'Cache-Control': 'no-store',
      });
      res.end(data);
  });
}

async function proxyFetch(res, target) {
  let parsed;
  try {
    parsed = new URL(target);
  } catch (e) {
    res.writeHead(400, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('非法的目标 URL');
    return;
  }
  if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') {
    res.writeHead(400, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('仅支持 http/https 链接');
    return;
  }
  try {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 25000);
    const resp = await fetch(target, {
      redirect: 'follow',
      signal: controller.signal,
      headers: {
        'User-Agent':
          'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 ' +
          '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        Accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'zh-CN,zh;q=0.9',
      },
    });
    clearTimeout(timer);

    // 读取原始字节，按页面声明的编码解码（兼容 GBK 等中文站点）
    const buf = Buffer.from(await resp.arrayBuffer());
    let charset = 'utf-8';
    const head = buf.slice(0, 2048).toString('latin1');
    const cm = head.match(/charset=([\w-]+)/i);
    if (cm && /gb/i.test(cm[1])) charset = 'gbk';
    const text = new TextDecoder(charset).decode(buf);

    res.writeHead(200, {
      'Content-Type': 'text/html; charset=utf-8',
      'Access-Control-Allow-Origin': '*',
      'Cache-Control': 'no-store',
    });
    res.end(text);
  } catch (e) {
    res.writeHead(502, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('抓取失败：' + (e && e.message ? e.message : String(e)));
  }
}

// ============ 同步到本地 Excel ============

// 目标表头 → 记录字段 的映射（兼容"招标金额"和"招标金额（元）"两种表头）
const SYNC_HEADER_ALIASES = {
  '项目名称': 'projectName',
  '项目单位': 'tenderee',
  '招标代理': 'agency',
  '开标时间': 'openTime',
  '招标方式': 'method',
  '招标金额（元）': 'tenderAmount',
  '招标金额': 'tenderAmount',
  '中标单位': 'winner',
  '中标金额（元）': 'winnerAmount',
  '中标金额': 'winnerAmount',
  '备注': 'remarks',
};
const SYNC_FIELD_ORDER = ['projectName', 'tenderee', 'agency', 'openTime', 'method',
  'tenderAmount', 'winner', 'winnerAmount', 'remarks'];

// 任意格式的"开标时间"值 → 归一化 'YYYY-MM-DD'（Excel 序列数字 / 字符串均可）
function syncDateKey(v) {
  if (v == null || v === '') return '';
  if (v instanceof Date) {
    return v.getUTCFullYear() + '-' + String(v.getUTCMonth() + 1).padStart(2, '0') + '-' + String(v.getUTCDate()).padStart(2, '0');
  }
  const n = Number(v);
  if (!isNaN(n) && /^\d{4,5}(\.\d+)?$/.test(String(v).trim())) {
    const d = new Date((n - 25569) * 86400000);
    return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0');
  }
  const m = String(v).match(/(\d{4})[-\/年.](\d{1,2})[-\/月.](\d{1,2})/);
  if (m) return m[1] + '-' + String(m[2]).padStart(2, '0') + '-' + String(m[3]).padStart(2, '0');
  return String(v).trim();
}

// 同步时写回的开标时间格式：有时刻带时刻，无时刻只留日期
function syncOpenTimeValue(s) {
  if (s == null) return '';
  const str = String(s).trim();
  if (!str) return '';
  const n = Number(str);
  if (!isNaN(n) && /^\d{4,5}(\.\d+)?$/.test(str)) {
    const d = new Date((n - 25569) * 86400000);
    const date = d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0');
    const hh = d.getUTCHours(), mm = d.getUTCMinutes();
    return (hh || mm) ? date + ' ' + String(hh).padStart(2, '0') + ':' + String(mm).padStart(2, '0') : date;
  }
  return str;
}

// 金额：纯数字写成数值（Excel 可直接求和），否则原样字符串
function syncAmountValue(v) {
  if (v == null || v === '') return '';
  const s = String(v).trim().replace(/,/g, '');
  if (/^\d+(\.\d+)?$/.test(s)) return Number(s);
  return String(v).trim();
}

// 写回前备份：同目录下 <名字>.backup-时间戳.xlsx，最多保留 5 份
function backupExcel(filePath) {
  const dir = path.dirname(filePath);
  const base = path.basename(filePath, path.extname(filePath));
  const stamp = new Date();
  const pad = (x) => String(x).padStart(2, '0');
  const ts = stamp.getFullYear() + pad(stamp.getMonth() + 1) + pad(stamp.getDate()) + '-' +
    pad(stamp.getHours()) + pad(stamp.getMinutes()) + pad(stamp.getSeconds());
  const backupPath = path.join(dir, base + '.backup-' + ts + '.xlsx');
  fs.copyFileSync(filePath, backupPath);
  // 只保留最近 5 份备份
  const prefix = base + '.backup-';
  const olds = fs.readdirSync(dir)
    .filter((f) => f.startsWith(prefix) && f.endsWith('.xlsx'))
    .map((f) => path.join(dir, f))
    .filter((f) => f !== backupPath)
    .sort();
  while (olds.length >= 5) fs.unlinkSync(olds.shift());
  return backupPath;
}

async function syncExcel(req, res, body) {
  let payload;
  try {
    payload = JSON.parse(body);
  } catch (e) {
    res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ ok: false, error: '请求体不是合法 JSON' }));
    return;
  }
  const filePath = String(payload.path || '').trim();
  const records = Array.isArray(payload.records) ? payload.records : [];
  if (!filePath || !/\.xlsx$/i.test(filePath)) {
    res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ ok: false, error: '目标文件必须是 .xlsx 路径' }));
    return;
  }
  if (!records.length) {
    res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ ok: false, error: '没有可同步的记录' }));
    return;
  }
  if (!fs.existsSync(filePath)) {
    res.writeHead(404, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ ok: false, error: '目标文件不存在：' + filePath }));
    return;
  }

  let warning = '';
  // Excel 占用锁文件检测（~$ 开头）；可能是残留锁，只提示不阻断
  try {
    const dir = path.dirname(filePath);
    const base = path.basename(filePath);
    const lock = path.join(dir, '~$' + base);
    if (fs.existsSync(lock)) warning = '检测到 Excel 锁文件（~$ 开头），该文件可能正在 Excel 中打开；请先关闭 Excel 再同步，否则改动可能被覆盖。';
  } catch (e) { /* 忽略锁检测失败 */ }

  try {
    // 1. 读取现有表格
    const wb = XLSX.read(fs.readFileSync(filePath), { type: 'buffer' });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });

    // 2. 定位表头行（前 5 行内找"项目名称"）
    let hIdx = -1;
    for (let i = 0; i < Math.min(5, matrix.length); i++) {
      if (Array.isArray(matrix[i]) && matrix[i].some((c) => String(c).includes('项目名称'))) { hIdx = i; break; }
    }
    if (hIdx < 0) throw new Error('表格中找不到表头行（含"项目名称"）');

    // 3. 建立 列号 → 字段 映射
    const colMap = {}; // field -> col index
    matrix[hIdx].forEach((h, i) => {
      const key = SYNC_HEADER_ALIASES[String(h).trim()];
      if (key && colMap[key] == null) colMap[key] = i;
    });
    const seqCol = matrix[hIdx].findIndex((h) => String(h).trim() === '序号');
    if (colMap.projectName == null) throw new Error('表格中找不到"项目名称"列');

    // 4. 现有数据行（表头之后、有项目名称的行）
    const dataRows = [];
    for (let i = hIdx + 1; i < matrix.length; i++) {
      const row = matrix[i];
      const name = String((row || [])[colMap.projectName] || '').trim();
      if (name) dataRows.push({ idx: i, row: row || [] });
    }
    const keyOf = (name, time) => name + '||' + syncDateKey(time);
    const existing = new Map();
    dataRows.forEach((d) => {
      const k = keyOf(String(d.row[colMap.projectName] || '').trim(), d.row[colMap.openTime]);
      if (!existing.has(k)) existing.set(k, d);
    });

    // 5. 合并：新记录追加；已存在的补空（只填 Excel 里为空的格子）
    let added = 0, updated = 0, skipped = 0;
    const newRows = [];
    records.forEach((rec) => {
      const f = (rec && rec.fields) || {};
      const name = String(f.projectName || '').trim();
      if (!name) { skipped++; return; }
      const k = keyOf(name, f.openTime);
      const hit = existing.get(k);
      if (hit) {
        // 补空更新
        let changed = false;
        SYNC_FIELD_ORDER.forEach((key) => {
          const col = colMap[key];
          if (col == null) return;
          const cur = String(hit.row[col] == null ? '' : hit.row[col]).trim();
          const val = key === 'openTime' ? syncOpenTimeValue(f.openTime)
            : key === 'tenderAmount' || key === 'winnerAmount' ? syncAmountValue(f[key])
              : String(f[key] == null ? '' : f[key]).trim();
          if (!cur && val !== '' && val != null) { hit.row[col] = val; changed = true; }
        });
        if (changed) { updated++; existing.set(k, hit); } else skipped++;
      } else {
        const row = new Array(matrix[hIdx].length).fill('');
        SYNC_FIELD_ORDER.forEach((key) => {
          const col = colMap[key];
          if (col == null) return;
          row[col] = key === 'openTime' ? syncOpenTimeValue(f.openTime)
            : key === 'tenderAmount' || key === 'winnerAmount' ? syncAmountValue(f[key])
              : String(f[key] == null ? '' : f[key]).trim();
        });
        newRows.push(row);
        // 同批次内也去重
        existing.set(k, { idx: -1, row });
        added++;
      }
    });

    // 6. 重写数据区：表头之前原样保留 + 表头 + 数据行（重编序号）
    const head = matrix.slice(0, hIdx + 1);
    const body = matrix.slice(hIdx + 1).map((r) => (Array.isArray(r) ? r : []));
    const merged = head.concat(body);
    newRows.forEach((r) => merged.push(r));

    // 重编序号（只给"有项目名称"的行）
    let seq = 0;
    for (let i = hIdx + 1; i < merged.length; i++) {
      const row = merged[i];
      const name = String((row || [])[colMap.projectName] || '').trim();
      if (name) {
        seq++;
        while (row.length <= seqCol) row.push('');
        if (seqCol >= 0) row[seqCol] = seq;
      }
    }

    // 7. 备份 + 写回
    const backupPath = backupExcel(filePath);
    const newWs = XLSX.utils.aoa_to_sheet(merged);
    // 保留原表头行的单元格格式（列宽等不可保留，仅保合并单元格）
    if (ws['!merges']) {
      newWs['!merges'] = ws['!merges'].filter((m) => m.e.r <= hIdx);
    }
    wb.Sheets[sheetName] = newWs;
    const out = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    fs.writeFileSync(filePath, out);

    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({
      ok: true, added, updated, skipped, total: seq,
      backup: backupPath, warning,
    }));
  } catch (e) {
    res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ ok: false, error: '写入失败：' + (e && e.message ? e.message : String(e)) }));
  }
}

// ============ 服务端数据持久化（替代浏览器 localStorage） ============
// 数据保存在 server-data/store.json，按 key 分库（tender / yigong）。
// 前端每次改动都会 PUT 到这里，读取时 GET；彻底摆脱"数据只存在某台浏览器"的局限，
// 因此「更新功能 / 换浏览器 / 换设备」都不会丢数据。
//
// 永久保存三重保险：
//   ① 主库 server-data/store.json（每次改动即写）；
//   ② 每日快照 server-data/backups/store-YYYY-MM-DD.json（主库丢失/损坏时自动恢复）；
//   ③ 启动时「新旧仲裁」：比较主库与最新快照的 savedAt，谁新用谁——
//      防止重新发布时把本地旧备份覆盖掉沙箱里更新的数据（升级不丢数据的关键）。
const DATA_DIR = path.join(__dirname, 'server-data');
const DATA_FILE = path.join(DATA_DIR, 'store.json');
const BACKUP_DIR = path.join(DATA_DIR, 'backups');
const MAX_BACKUPS = 30;
const DATA_KEYS = { tender: true, yigong: true };

/** 读取一个 store 文件，返回 { savedAt, data }；损坏返回 null。
 *  savedAt 取自 _meta（没有 _meta 视为最旧：可能是旧格式文件或部署上传的旧备份） */
function readStoreFile(fp) {
  try {
    const j = JSON.parse(fs.readFileSync(fp, 'utf8'));
    if (!j || typeof j !== 'object') return null;
    const savedAt = (j._meta && j._meta.savedAt) || '';
    return { savedAt: savedAt, data: j };
  } catch (e) { return null; }
}

/** 找到最新一份快照 */
function findNewestBackup() {
  try {
    if (!fs.existsSync(BACKUP_DIR)) return null;
    const files = fs.readdirSync(BACKUP_DIR)
      .filter((f) => /^store-\d{4}-\d{2}-\d{2}\.json$/.test(f))
      .sort();
    for (let i = files.length - 1; i >= 0; i--) {
      const r = readStoreFile(path.join(BACKUP_DIR, files[i]));
      if (r) return r;
    }
  } catch (e) { /* 忽略 */ }
  return null;
}

function loadStore() {
  const main = readStoreFile(DATA_FILE);
  const snap = findNewestBackup();
  // 主库缺失/损坏 → 用快照恢复
  if (!main) {
    if (snap) {
      console.log('[数据] 主库丢失，已从快照恢复（快照时间 ' + snap.savedAt + '）');
      return snap.data;
    }
    return {};
  }
  // 新旧仲裁：主库没有时间戳（旧格式/部署带来的旧备份）或快照更新 → 用快照
  if (snap && snap.savedAt && (!main.savedAt || snap.savedAt > main.savedAt)) {
    console.log('[数据] 检测到快照（' + snap.savedAt + '）比主库' +
      (main.savedAt ? '（' + main.savedAt + '）' : '（无时间戳）') + '新，已采用快照数据');
    return snap.data;
  }
  return main.data || {};
}
let DATA_STORE = loadStore();

function persistStore() {
  try {
    if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
    DATA_STORE._meta = { savedAt: new Date().toISOString() };
    fs.writeFileSync(DATA_FILE, JSON.stringify(DATA_STORE));
    // 每日快照（同一天内覆盖更新，始终是当天最新状态）
    try {
      if (!fs.existsSync(BACKUP_DIR)) fs.mkdirSync(BACKUP_DIR, { recursive: true });
      const d = new Date();
      const name = 'store-' + d.getFullYear() + '-' +
        String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0') + '.json';
      fs.writeFileSync(path.join(BACKUP_DIR, name), JSON.stringify(DATA_STORE));
      // 只保留最近 MAX_BACKUPS 份快照
      const files = fs.readdirSync(BACKUP_DIR)
        .filter((f) => /^store-\d{4}-\d{2}-\d{2}\.json$/.test(f)).sort();
      while (files.length > MAX_BACKUPS) {
        fs.unlinkSync(path.join(BACKUP_DIR, files.shift()));
      }
    } catch (e) { /* 快照失败不影响主库 */ }
  } catch (e) { /* 写入失败不阻断主流程 */ }
}

function handleDataGet(res, key) {
  const arr = Array.isArray(DATA_STORE[key]) ? DATA_STORE[key] : [];
  res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify({ ok: true, records: arr }));
}

function handleDataPut(res, key, body) {
  let payload;
  try {
    payload = JSON.parse(body);
  } catch (e) {
    res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ ok: false, error: '请求体不是合法 JSON' }));
    return;
  }
  const records = Array.isArray(payload.records) ? payload.records : [];
  DATA_STORE[key] = records;
  persistStore();
  res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify({ ok: true, count: records.length }));
}

const server = http.createServer((req, res) => {
  const parsed = new URL(req.url, 'http://localhost');
  if (parsed.pathname === '/api/fetch') {
    const target = parsed.searchParams.get('url');
    if (!target) {
      res.writeHead(400, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('缺少 url 参数');
      return;
    }
    proxyFetch(res, target);
    return;
  }
  if (parsed.pathname === '/api/sync-excel' && req.method === 'POST') {
    const chunks = [];
    req.on('data', (c) => chunks.push(c));
    req.on('end', () => syncExcel(req, res, Buffer.concat(chunks).toString('utf8')));
    return;
  }
  // ============ 服务端数据持久化（替代浏览器 localStorage） ============
  // GET  /api/data/:key  → { ok:true, records:[...] }
  // PUT  /api/data/:key  → body { records:[...] } → { ok:true, count }
  // 数据落在 server-data/store.json，按 key 分库（tender / yigong）。
  const dm = parsed.pathname.match(/^\/api\/data\/([a-z0-9_]+)$/);
  if (dm) {
    const key = dm[1];
    if (!DATA_KEYS[key]) {
      res.writeHead(403, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('未知的数据键');
      return;
    }
    if (req.method === 'GET') { handleDataGet(res, key); return; }
    if (req.method === 'PUT' || req.method === 'POST') {
      const chunks = [];
      req.on('data', (c) => chunks.push(c));
      req.on('end', () => handleDataPut(res, key, Buffer.concat(chunks).toString('utf8')));
      return;
    }
    res.writeHead(405, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('不支持的请求方法');
    return;
  }
  serveStatic(req, res);
});

server.listen(PORT, '0.0.0.0', () => {
  console.log('✅ 工作台已启动： http://localhost:' + PORT);
  console.log('   在浏览器打开上面的地址即可使用（自动抓取已可绕过跨域）。');
  console.log('   按 Ctrl+C 停止服务。');
});
