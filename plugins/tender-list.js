/**
 * 插件：统计招投标项目清单
 * --------------------------------------------------
 * 字段已按文件一《2026年招投标项目交易清单》更新：
 *   开标月份(自动归并) | 序号(自动) | 项目名称 | 项目单位 | 招标代理 | 开标时间
 *   | 招标方式 | 招标金额（元） | 中标单位 | 中标金额（元） | 备注（共 11 列，CSV 含开标月份）
 * 解析器针对「海南省公共资源交易中心」：
 *   - 招标公告页解析前 8 列；
 *   - 自动读取页面隐藏域 gonggaoguid，加载 /ggzyjy/json/<guid>.json 获取同项目全部关联公告，
 *     定位「中标公示/中标结果」文章并解析 中标人→中标单位、中标价格→中标金额（自动填入）。
 */
(function () {
  // ============== ① 表格字段（文件一正式字段） ==============
  const FIELDS = [
    { key: 'openMonth',    label: '开标月份', computed: true },
    { key: 'seq',          label: '序号',          auto: true },
    { key: 'projectName',  label: '项目名称' },
    { key: 'tenderee',     label: '项目单位' },
    { key: 'agency',       label: '招标代理' },
    { key: 'openTime',     label: '开标时间' },
    { key: 'method',       label: '招标方式' },
    { key: 'tenderAmount', label: '招标金额（元）' },
    { key: 'winner',       label: '中标单位' },
    { key: 'winnerAmount', label: '中标金额（元）' },
    { key: 'remarks',      label: '备注' },
  ];

  const STORE_KEY = 'wb_tender_records';
  let records = [];
  // 项目类型分表：'construction' 建设工程（批量抓取/Excel 导入），
  // 'government' 政府采购（在开标日历手动添加）。历史数据无该字段，默认归入建设工程。
  let currentType = 'construction';
  const TYPE_NAMES = { construction: '建设工程', government: '政府采购' };
  function recType(r) { return (r && r.projectType === 'government') ? 'government' : 'construction'; }

  // ============== ② 解析规则（针对海南公共资源交易中心招标公告页） ==============
  function normalizeDate(s) {
    if (!s) return '';
    const m = s.match(/(\d{4})年(\d{2})月(\d{2})日(\d{2})时(\d{2})分/);
    return m ? `${m[1]}-${m[2]}-${m[3]} ${m[4]}:${m[5]}` : s;
  }
  // 解析开标时间为 Date（兼容 "2026-07-08 08:30"、"2026年07月08日08时30分" 等；并兼容 Excel 日期序列数字）
  function parseOpenTime(s) {
    if (s == null) return null;
    const str = String(s).trim();
    if (!str) return null;
    // Excel 日期序列数字（如 46080）→ 视为日期；基准偏移 25569（Excel 1900 系统，25569=1970-01-01）
    const n = Number(str);
    if (!isNaN(n) && n > 10000 && n < 80000 && /^\d{4,5}(\.\d+)?$/.test(str)) {
      return new Date((n - 25569) * 86400000);
    }
    const m = str.match(/(\d{4})[-年](\d{1,2})[-月](\d{1,2})(?:[日\s]*(\d{1,2})?[:时]?(\d{1,2})?)?/);
    if (!m) return null;
    const dt = new Date(+m[1], +m[2] - 1, +m[3], m[4] ? +m[4] : 0, m[5] ? +m[5] : 0);
    return isNaN(dt.getTime()) ? null : dt;
  }
  // 由开标时间得到月份信息：{ key, label }；无有效时间则归为「未注明」
  function monthOf(s) {
    const dt = parseOpenTime(s);
    if (!dt) return { key: '__none__', label: '未注明' };
    return { key: dt.getFullYear() + '-' + String(dt.getMonth() + 1).padStart(2, '0'), label: (dt.getMonth() + 1) + '月份' };
  }

  function cleanAmount(s) {
    return /^[\d.]+$/.test(s) && /\.0$/.test(s) ? s.slice(0, -2) : s;
  }

  // ============== ②½ 统计辅助 ==============
  // 把金额字符串（可能带千分位、小数点）转成数字
  function parseAmount(s) {
    if (s == null) return 0;
    const n = Number(String(s).replace(/[^\d.]/g, ''));
    return isNaN(n) ? 0 : n;
  }
  // 友好显示金额：元 → 万元/亿元
  function formatMoney(n) {
    if (!n) return '¥0';
    if (n >= 1e8) return '¥' + (n / 1e8).toFixed(2) + ' 亿元';
    if (n >= 1e4) return '¥' + (n / 1e4).toFixed(2) + ' 万元';
    return '¥' + Math.round(n).toLocaleString('zh-CN');
  }
  // 按年份统计：项目总数 + 招标金额合计 + 中标金额合计 + 中标下浮率 + 逐月明细
  // 仅统计建设工程项目；政府采购（日历手动添加）不参与按年/按月统计
  function computeStats(year) {
    const byMonth = {}; // 'YYYY-MM' -> { count, tenderAmount, winnerAmount }
    let totalCount = 0;
    let totalTender = 0;   // 招标金额合计
    let totalWinner = 0;   // 中标金额合计
    let baseTender = 0;    // 下浮率计算基数（仅含"既有招标又有中标"的项目，按招标金额加权）
    let baseWinner = 0;
    records.forEach((rec) => {
      if (recType(rec) !== 'construction') return; // 政府采购不参与统计
      const dt = parseOpenTime(rec.fields.openTime || '');
      if (!dt) return; // 未注明开标时间的不计入统计
      if (year && dt.getFullYear() !== year) return;
      totalCount++;
      const t = parseAmount(rec.fields.tenderAmount);
      const w = parseAmount(rec.fields.winnerAmount);
      totalTender += t;
      totalWinner += w;
      if (t > 0 && w > 0) { baseTender += t; baseWinner += w; }
      const key = dt.getFullYear() + '-' + String(dt.getMonth() + 1).padStart(2, '0');
      if (!byMonth[key]) byMonth[key] = { count: 0, tenderAmount: 0, winnerAmount: 0 };
      byMonth[key].count++;
      byMonth[key].tenderAmount += t;
      byMonth[key].winnerAmount += w;
    });
    // 中标下浮率 = (招标金额 - 中标金额) / 招标金额，仅在有中标数据时计算
    const downRate = baseTender > 0 ? ((baseTender - baseWinner) / baseTender) * 100 : null;
    return { totalCount, totalTender, totalWinner, downRate, byMonth };
  }
  // 招标代理字段只保留企业名称，去掉后面跟的地址/联系人/电话等冗余信息
  function cleanAgency(s) {
    if (!s) return '';
    // 1) 遇到"地址/联系人/电话/传真/邮箱/邮编"等后续字段即截断（这些不可能是企业名的一部分）
    let name = s.split(/(地\s*址|联系人|电\s*话|传\s*真|邮\s*箱|邮\s*编|邮政编码|统一社会信用代码|公司地址)/)[0];
    // 2) 若名称后还跟着含联系方式的括号，如"（联系人：X 电话：Y）"，一并去掉
    name = name.replace(/[（(][^）)]*(电话|联系|地址|传真|邮箱)[^）)]*[）)]/g, '');
    // 3) 收尾：去掉上一步截断后残留的孤立左括号与尾随标点/空白
    name = name.replace(/[（(]\s*$/, '').replace(/[，,、:：\s]+$/, '');
    return name.trim();
  }

  // 将网页 HTML 转为纯文本（用于正则解析）
  function toText(raw) {
    if (/<html|<body/i.test(raw)) {
      const doc = new DOMParser().parseFromString(raw, 'text/html');
      return doc.body ? doc.body.textContent : raw;
    }
    return raw;
  }

  // 从文本中抽取中标人 / 中标金额（适用于中标公示 / 中标结果 / 中标候选人公告）。
  // 返回 { winner, winnerAmount }，无则均为空串。
  function extractWinner(text) {
    const get = (re, idx = 1) => {
      const m = text.match(re);
      return m ? (m[idx] || '').trim() : '';
    };
    // 中标人：取「中标人:」到「投标报价 / 中标价格 / 句号」之间的内容；
    // 若公告仅含中标候选人公示，则回退匹配第一中标候选人 / 中标候选人
    let winner = get(/中标人[:：]\s*([\s\S]*?)(?=投标报价|中标价格|。|$)/);
    if (!winner) winner = get(/第一中标候选人[:：]\s*([\s\S]*?)(?=投标报价|中标价格|。|$)/);
    if (!winner) winner = get(/中标候选人[:：]\s*([\s\S]*?)(?=投标报价|中标价格|。|$)/);
    // 去掉末尾可能残留的「投标报价 / 中标价格」等尾巴
    winner = winner.replace(/(投标报价|中标价格|中标金额)[\s\S]*$/, '').replace(/[，,、\s]+$/, '').trim();
    const winnerAmount = cleanAmount(get(/中标(?:价格|金额)[:：]\s*([\d.]+)/));
    return { winner, winnerAmount };
  }

  // 自动关联中标公示：
  //   1) 从招标公告页读取隐藏域 gonggaoguid；
  //   2) 加载 /ggzyjy/json/<guid>.json 获取同项目全部关联公告；
  //   3) 定位「中标公示 / 中标结果」文章（标题含「中标」且非「候选人」，优先最新阶段），
  //      抓取该文章并解析 中标人→中标单位、中标价格→中标金额。
  // 非海南站点（无 gonggaoguid）或解析失败均返回 null，不影响主记录。
  async function fetchWinningInfo(html, url) {
    const m = html.match(/id="gonggaoguid">([^<]+)/);
    if (!m) return null;
    const guid = m[1].trim();
    const jsonUrl = url.replace(/\/jyxx\/.*$/, '/json/' + guid + '.json');
    const r = await fetchPage(jsonUrl);
    if (!r.ok) return null;
    let list;
    try {
      const j = JSON.parse(r.text);
      list = Array.isArray(j) ? j : (j.list || j.data || []);
    } catch (e) { return null; }
    if (!list.length) return null;
    const origin = (() => { try { return new URL(url).origin; } catch (e) { return 'https://ggzy.hainan.gov.cn'; } })();
    // 优先取「中标」且非「候选人」的公告（即中标结果/中标公示），若没有则放宽到含「中标」
    const cands = list.filter((it) => /中标/.test(it.title || '') && !/候选人/.test(it.title || ''));
    const pool = cands.length ? cands : list.filter((it) => /中标/.test(it.title || ''));
    if (!pool.length) return null;
    // 取分类号最大者（阶段最靠后 = 最终结果）
    pool.sort((a, b) => String(a.categorynum || '').localeCompare(String(b.categorynum || '')));
    const item = pool[pool.length - 1];
    let visit = item.visiturl || '';
    if (!visit) return null;
    if (visit.startsWith('/')) visit = origin + visit;
    else if (!/^https?:/i.test(visit)) visit = origin + '/' + visit.replace(/^\//, '');
    const wr = await fetchPage(visit);
    if (!wr.ok) return null;
    const w = extractWinner(toText(wr.text));
    if (!w.winner && !w.winnerAmount) return null;
    return { winner: w.winner, winnerAmount: w.winnerAmount, source: visit };
  }

  function parseTenderPage(raw, url) {
    const text = toText(raw);
    const get = (re, idx = 1) => {
      const m = text.match(re);
      return m ? (m[idx] || '').trim() : '';
    };

    const deadline = get(
      /投标文件递交的截止时间[（(]投标截止时间，下同[)）]为\s*([\d]{4}年[\d]{2}月[\d]{2}日[\d]{2}时[\d]{2}分)/
    );
    // 招标方式：优先识别"机器管招投标"，否则取公开/邀请等
    let method = '机器管';
    if (!/机器管招投标/.test(text)) {
      const m = text.match(/进行(公开招标|邀请招标|竞争性谈判|询价|单一来源采购)/);
      method = m ? m[1] : '';
    }

    const f = {};
    f.projectName  = get(/（机器管招投标）?([\u4e00-\u9fa5A-Za-z0-9·（）()]+?)招标公告/).replace(/^（机器管招投标）/, '');
    f.tenderee     = get(/招标人\(?项目业主\)?为\s*([^，,]+)/);
    f.agency       = cleanAgency(get(/招标代理机构:\s*([^\n]+)/));
    f.openTime     = normalizeDate(deadline);
    f.method       = method;
    f.tenderAmount = cleanAmount(get(/2\.5\s*最高投标限价[（(]或招标控制价[)）]:\s*([\d.]+)/));
    // 若页面本身就是中标公示/中标结果（直接粘贴或抓取），则直接抽取中标人/中标金额
    const w = extractWinner(text);
    f.winner       = w.winner;
    f.winnerAmount = w.winnerAmount;
    f.remarks      = '';
    return f;
  }

  // ============== ②b 跨域抓取：本地代理优先，公共代理兜底 ==============
  // 纯前端直接 fetch 政府网站会被浏览器 CORS 拦截。
  // 解决办法：由同目录的 server.js 提供同源代理端点 /api/fetch，
  // 由服务端去抓页面再回传，从而绕开浏览器跨域限制。
  // 若服务端代理不可用（如 file:// 直接打开），再尝试公共 CORS 代理兜底。
  async function fetchPage(target) {
    const candidates = [];
    if (location.protocol !== 'file:') {
      candidates.push({ type: 'local', url: '/api/fetch?url=' + encodeURIComponent(target) });
    }
    // 公共 CORS 代理（仅 file:// 模式或本地代理不可用时的兜底，可能不稳定/被墙）
    candidates.push({ type: 'proxy', url: 'https://api.allorigins.win/raw?url=' + encodeURIComponent(target) });
    candidates.push({ type: 'proxy', url: 'https://api.codetabs.com/v1/proxy/?quest=' + encodeURIComponent(target) });
    let lastErr = '';
    for (const c of candidates) {
      try {
        const resp = await fetch(c.url);
        if (resp.ok) {
          const text = await resp.text();
          if (text && text.trim().length) return { ok: true, text: text, via: c.type };
          lastErr = '返回内容为空';
        } else {
          lastErr = 'HTTP ' + resp.status;
        }
      } catch (e) {
        lastErr = e && e.message ? e.message : String(e);
      }
    }
    return { ok: false, err: lastErr };
  }

  // 从输入中解析出多条链接（每行一个，或同行以空格/逗号/顿号分隔）
  function splitLinks(text) {
    return String(text || '')
      .split(/[\n\r]+/)
      .flatMap((line) => line.split(/[\s,，、]+/))
      .map((s) => s.trim())
      .filter((s) => /^https?:\/\//i.test(s));
  }

  // ============== ②c 导入 Excel / CSV（浏览器端解析） ==============
  // 表头别名 → 字段 key（兼容文件一「招标金额（元）」等带单位表头）
  const IMPORT_ALIAS = {
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
  // 序号 / 开标月份 在导入时由系统自动处理，不映射到字段
  // Excel 日期序列值 → 本地日期字符串（Excel 1900 日期系统）
  // 基准偏移用 25569（Excel 序列 25569 = 1970-01-01）。注意：必须读取原始序列值(raw)再转换，
  // 不能直接用 SheetJS 的 Date 对象——后者自带 1900 闰年虚增偏移与浮点误差，会按浏览器时区"差一天"。
  function excelSerialToDate(serial) {
    const ms = (serial - 25569) * 86400000;
    const d = new Date(ms);
    const y = d.getUTCFullYear();
    const m = String(d.getUTCMonth() + 1).padStart(2, '0');
    const day = String(d.getUTCDate()).padStart(2, '0');
    const hh = String(d.getUTCHours()).padStart(2, '0');
    const mm = String(d.getUTCMinutes()).padStart(2, '0');
    // 仅当含非零时间才返回时间部分（保留开标时分，便于排序/月份判断一致）
    return (hh === '00' && mm === '00') ? (y + '-' + m + '-' + day) : (y + '-' + m + '-' + day + ' ' + hh + ':' + mm);
  }
  function coerceField(key, v) {
    if (v == null) return '';
    // 开标时间：Excel 日期列可能被读成 Date 对象或序列数字，统一归一化为 YYYY-MM-DD
    if (key === 'openTime') {
      if (v instanceof Date) {
        const y = v.getFullYear();
        const m = String(v.getMonth() + 1).padStart(2, '0');
        const d = String(v.getDate()).padStart(2, '0');
        return y + '-' + m + '-' + d;
      }
      const n = Number(v);
      if (!isNaN(n) && n > 10000 && n < 80000) return excelSerialToDate(n); // 典型 5 位 Excel 日期序列
      return String(v).trim();
    }
    let s = String(v).trim();
    if ((key === 'tenderAmount' || key === 'winnerAmount') && /^[\d.]+$/.test(s) && /\.0$/.test(s)) s = s.slice(0, -2);
    return s;
  }
  // 由矩阵（二维数组）解析：自动定位含「项目名称」的表头行，再按列映射
  function parseMatrix(matrix) {
    const hIdx = matrix.findIndex((r) => Array.isArray(r) && r.some((c) => String(c).includes('项目名称')));
    if (hIdx < 0) return [];
    const header = matrix[hIdx].map((h) => String(h).trim());
    const colMap = {};
    header.forEach((h, i) => { if (IMPORT_ALIAS[h]) colMap[i] = IMPORT_ALIAS[h]; });
    return matrix.slice(hIdx + 1)
      .filter((r) => Array.isArray(r) && r.some((c) => String(c).trim()))
      .map((r) => {
        const f = {};
        Object.keys(colMap).forEach((i) => { f[colMap[i]] = coerceField(colMap[i], r[i]); });
        return f;
      })
      .filter((f) => f.projectName);
  }
  // 极简 CSV 解析（支持引号内的逗号/换行）
  function parseCsvText(text) {
    const rows = [];
    let row = [], field = '', inQ = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (inQ) {
        if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else inQ = false; }
        else field += c;
      } else {
        if (c === '"') inQ = true;
        else if (c === ',') { row.push(field); field = ''; }
        else if (c === '\n' || c === '\r') { if (c === '\r' && text[i + 1] === '\n') i++; row.push(field); rows.push(row); row = []; field = ''; }
        else field += c;
      }
    }
    if (field.length || row.length) { row.push(field); rows.push(row); }
    return parseMatrix(rows);
  }
  function dedupeKey(fields) {
    return ((fields && fields.projectName) || '').trim() + '|' + ((fields && fields.openTime) || '').trim();
  }
  // 合并一批字段对象到 records（按 项目名称+开标时间 去重，避免重复导入）
  function mergeRecords(fieldsList) {
    let added = 0, skipped = 0;
    const seen = new Set(records.map((r) => dedupeKey(r.fields)));
    (fieldsList || []).forEach((f) => {
      if (!f || !f.projectName) return;
      const key = dedupeKey(f);
      if (seen.has(key)) { skipped++; return; }
      seen.add(key);
      addRecord(f, '');
      added++;
    });
    return { added, skipped };
  }
  async function importFile(file, statusEl) {
    statusEl.className = 'status info';
    statusEl.textContent = '解析中…';
    try {
      const buf = await file.arrayBuffer();
      let fieldsList;
      if (/\.csv$/i.test(file.name)) {
        fieldsList = parseCsvText(new TextDecoder('utf-8').decode(buf));
      } else {
        if (!window.XLSX) { statusEl.className = 'status err'; statusEl.textContent = 'Excel 解析库未加载，请刷新页面后重试。'; return; }
        // 读取原始值（raw:true）：日期单元格返回精确的 Excel 序列数字（而非含浮点误差的 Date 对象），
        // 再由 coerceField → excelSerialToDate 统一转换（含 1900 闰年 +1 修正），结果与时区无关、稳定正确。
        const wb = window.XLSX.read(buf, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        fieldsList = parseMatrix(window.XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true }));
      }
      const res = mergeRecords(fieldsList);
      statusEl.className = res.added ? 'status ok' : 'status warn';
      statusEl.textContent = '✓ 导入完成：新增 ' + res.added + ' 条' +
        (res.skipped ? '，跳过重复 ' + res.skipped + ' 条' : '') +
        '；清单共 ' + records.length + ' 条。';
      renderTable(containerRef);
      if (res.added > 0 && autoSyncCb) autoSyncCb(); // 勾选了自动同步则写入本地 Excel
    } catch (err) {
      statusEl.className = 'status err';
      statusEl.textContent = '✗ 解析失败：' + (err && err.message ? err.message : err);
    }
  }
  // 首次打开时，若工作区存在 seed_import.json（由文件一生成），提示一键导入
  function maybeSeed(container) {
    fetch('/seed_import.json')
      .then((r) => (r.ok ? r.json() : null))
      .then((arr) => {
        if (!arr || !arr.length) return;
        if (localStorage.getItem('tender_seed_done')) return;
        const banner = document.createElement('div');
        banner.className = 'banner seed-banner';
        banner.innerHTML =
          '📥 检测到待导入清单 <b>' + arr.length + '</b> 条（来自文件一 Excel）。' +
          '<button class="btn small" id="seed-import">一键导入</button>' +
          '<button class="btn ghost small" id="seed-ignore">忽略</button>';
        container.insertBefore(banner, container.firstChild);
        banner.querySelector('#seed-import').onclick = async () => {
          if (!Workbench.ensureEditable()) return;
          const res = mergeRecords(arr);
          localStorage.setItem('tender_seed_done', '1');
          banner.remove();
          renderTable(container);
          alert('已导入 ' + res.added + ' 条（跳过重复 ' + res.skipped + ' 条），清单共 ' + records.length + ' 条。');
        };
        banner.querySelector('#seed-ignore').onclick = () => {
          localStorage.setItem('tender_seed_done', '1');
          banner.remove();
        };
      })
      .catch(() => {});
  }


  // ============== ③ 存储（服务端持久化，替代浏览器本地） ==============
  function save() {
    // 写入服务端（并自动在本机缓存兜底），数据不再只存于某一台浏览器
    Workbench.saveRecords(STORE_KEY, records);
  }
  async function load() {
    records = await Workbench.loadRecords(STORE_KEY);
  }

  // ============== ④ 导出 / 汇报 ==============
  function download(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename; a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
  function escapeCsv(v) {
    v = String(v == null ? '' : v);
    if (/[",\n\r]/.test(v)) v = '"' + v.replace(/"/g, '""') + '"';
    return v;
  }
  function cellValue(f, r, i) {
    if (f.key === 'openMonth') return monthOf((r.fields && r.fields.openTime) || '').label;
    if (f.key === 'openTime') {
      const v = (r.fields && r.fields.openTime) || '';
      // 显示时归一化：Excel 序列数字 → 本地日期字符串
      const n = Number(v);
      if (!isNaN(n) && n > 10000 && n < 80000 && /^\d{4,5}(\.\d+)?$/.test(String(v).trim())) return excelSerialToDate(n);
      return v;
    }
    if (f.auto) return i + 1; // 序号按当前（开标时间升序）排列后重新编号
    return (r.fields && r.fields[f.key]) || '';
  }
  // 构建表格行：按开标时间升序排序，并为「开标月份」列计算合并信息（同月相邻行合并）
  function buildRows(type) {
    const src = type ? records.filter((r) => recType(r) === type) : records;
    const arr = src.map((r) => ({ rec: r, ot: (r.fields && r.fields.openTime) || '' }));
    arr.sort((a, b) => {
      const da = parseOpenTime(a.ot), db = parseOpenTime(b.ot);
      const ta = da ? da.getTime() : Infinity;
      const tb = db ? db.getTime() : Infinity;
      if (ta !== tb) return ta - tb;
      return (a.rec.addedAt || 0) - (b.rec.addedAt || 0);
    });
    const rows = arr.map((x) => {
      const m = monthOf(x.ot);
      return { rec: x.rec, mkey: m.key, mlabel: m.label };
    });
    rows.forEach((row, i) => {
      if (i > 0 && rows[i - 1].mkey === row.mkey) {
        row.monthStart = false;
      } else {
        let span = 1;
        while (i + span < rows.length && rows[i + span].mkey === row.mkey) span++;
        row.monthStart = true;
        row.monthSpan = span;
      }
    });
    return rows;
  }
  function exportCSV() {
    let csv = '﻿' + FIELDS.map((f) => f.label).join(',') + '\n';
    buildRows(currentType).forEach((row, i) => {
      csv += FIELDS.map((f) => escapeCsv(String(cellValue(f, row.rec, i)).replace(/[\n\r]+/g, ' '))).join(',') + '\n';
    });
    download(new Blob([csv], { type: 'text/csv;charset=utf-8' }), TYPE_NAMES[currentType] + '项目清单.csv');
  }

  function loadDocx() {
    return new Promise((resolve) => {
      if (window.docx) return resolve(window.docx);
      const s = document.createElement('script');
      s.src = 'https://cdn.jsdelivr.net/npm/docx@8.5.1/build/index.umd.min.js';
      s.onload = () => resolve(window.docx);
      s.onerror = () => resolve(null);
      document.head.appendChild(s);
    });
  }
  async function generateWord() {
    const docx = await loadDocx();
    if (!docx) return fallbackDoc();
    const { Document, Paragraph, TextRun, Table, TableRow, TableCell, Packer } = docx;
    const headerRow = new TableRow({
      children: FIELDS.map(
        (f) => new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: f.label, bold: true })] })] })
      ),
    });
    const dataRows = buildRows(currentType).map(
      (row, i) =>
        new TableRow({
          children: FIELDS.map(
            (f) => new TableCell({ children: [new Paragraph(String(cellValue(f, row.rec, i)))] })
          ),
        })
    );
    const doc = new Document({
      sections: [
        {
          children: [
            new Paragraph({ text: TYPE_NAMES[currentType] + '项目清单汇报', heading: 'Heading1' }),
            new Paragraph({ text: '生成时间：' + new Date().toLocaleString('zh-CN') }),
            new Paragraph({ text: '共计项目：' + dataRows.length + ' 条', spacing: { after: 200 } }),
            new Table({ rows: [headerRow, ...dataRows] }),
          ],
        },
      ],
    });
    const blob = await Packer.toBlob(doc);
    download(blob, TYPE_NAMES[currentType] + '项目汇报.docx');
  }
  function fallbackDoc() {
    let html =
      '<html xmlns:o="urn:schemas-microsoft-com:office:office"><head><meta charset="utf-8"></head><body>';
    html += '<h1>' + TYPE_NAMES[currentType] + '项目清单汇报</h1>';
    html += '<p>生成时间：' + new Date().toLocaleString('zh-CN') + '；共计 ' + records.filter((r) => recType(r) === currentType).length + ' 条</p>';
    html += '<table border="1" cellspacing="0"><tr>' +
      FIELDS.map((f) => '<th>' + f.label + '</th>').join('') + '</tr>';
    buildRows(currentType).forEach((row, i) => {
      html += '<tr>' + FIELDS.map((f) => '<td>' + cellValue(f, row.rec, i) + '</td>').join('') + '</tr>';
    });
    html += '</table></body></html>';
    download(new Blob([html], { type: 'application/msword' }), TYPE_NAMES[currentType] + '项目汇报.doc');
  }

  // ============== ⑤b 双击单元格编辑 ==============
  // 双击任意「字段」单元格进入编辑态（input 内联编辑），回车/Esc 提交或取消。
  // 开标月份(计算列)与序号(自动列)不可直接编辑；其它字段均可手动填写。
  // 若编辑的是「开标时间」，提交后会触发 renderTable → buildRows 按开标时间升序重排，
  //   并重新归并「开标月份」合并单元格，即"自动排序到对应的月份"。
  let containerRef = null;
  function editCell(td) {
    if (!Workbench.ensureEditable()) return; // 未进入修改模式禁止编辑
    const field = td.dataset.field;
    const id = td.dataset.id;
    const idx = Number(td.dataset.idx || '0');
    if (!field || !id) return;
    const f = FIELDS.find((x) => x.key === field);
    if (!f || f.computed || f.auto) return; // 开标月份（计算列）和序号（自动列）不可直接编辑
    const rec = records.find((r) => r.id === id);
    if (!rec) return;
    if (td.querySelector('input.cell-edit')) return; // 已在编辑中
    const current = cellValue(f, rec, idx); // 以当前显示值作为初始可编辑内容
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'cell-edit';
    input.value = current;
    td.classList.add('editing');
    td.textContent = '';
    td.appendChild(input);
    input.focus();
    input.select();

    let done = false;
    const finish = (keep) => {
      if (done) return;
      done = true;
      const val = input.value.trim();
      const changed = val !== (rec.fields[field] || '');
      if (keep && changed) {
        rec.fields[field] = val;
        save();
      }
      // 值未变也需移除输入框恢复显示；重渲开销小，统一走 renderTable
      if (containerRef) renderTable(containerRef);
    };
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); finish(true); }
      else if (e.key === 'Escape') { e.preventDefault(); finish(false); }
    });
    input.addEventListener('blur', () => finish(true));
  }

  // ============== ⑤ 渲染 ==============
  // 统计概览卡片（顶部汇总 + 按月明细）
  let statsYear = 2026; // 当前统计年份，由下拉框控制；首次渲染时按数据最新年份校正
  function renderStats(root) {
    const cardsEl = root.querySelector('#stat-cards');
    const gridEl = root.querySelector('#month-grid');
    const yearSel = root.querySelector('#stats-year');
    const totalEl = root.querySelector('#month-total');
    if (!cardsEl || !gridEl) return;

    // 收集数据中出现的年份（仅建设工程，政府采购不参与统计），默认取最新年份
    const years = new Set();
    records.forEach((rec) => {
      if (recType(rec) !== 'construction') return;
      const dt = parseOpenTime(rec.fields.openTime || '');
      if (dt) years.add(dt.getFullYear());
    });
    if (!years.size) years.add(2026);
    const yearArr = [...years].sort((a, b) => a - b);
    if (!yearArr.includes(statsYear)) statsYear = yearArr[yearArr.length - 1];
    if (yearSel) {
      yearSel.innerHTML = yearArr.map((y) => '<option value="' + y + '"' + (y === statsYear ? ' selected' : '') + '>' + y + '年</option>').join('');
    }

    const stats = computeStats(statsYear);
    const downLabel = stats.downRate == null ? '暂无中标数据' : stats.downRate.toFixed(1) + '%';
    cardsEl.innerHTML =
      '<div class="stat-card primary"><div class="stat-num">' + stats.totalCount + '</div>' +
      '<div class="stat-label">' + statsYear + ' 年招标项目总数（个）</div></div>' +
      '<div class="stat-card accent"><div class="stat-num">' + formatMoney(stats.totalTender) + '</div>' +
      '<div class="stat-label">' + statsYear + ' 年招标金额合计</div></div>' +
      '<div class="stat-card teal"><div class="stat-num">' + formatMoney(stats.totalWinner) + '</div>' +
      '<div class="stat-label">' + statsYear + ' 年中标金额合计</div></div>' +
      '<div class="stat-card purple"><div class="stat-num">' + downLabel + '</div>' +
      '<div class="stat-label">' + statsYear + ' 年中标下浮率</div></div>';

    // 固定渲染 1-12 月，网格整齐（无数据的月份显示空态占位）
    const maxTender = Object.keys(stats.byMonth).reduce(
      (m, k) => Math.max(m, (stats.byMonth[k] || {}).tenderAmount || 0), 0) || 1;
    let gridHtml = '';
    for (let mm = 1; mm <= 12; mm++) {
      const key = statsYear + '-' + String(mm).padStart(2, '0');
      const mo = stats.byMonth[key] || { count: 0, tenderAmount: 0, winnerAmount: 0 };
      const pctT = mo.tenderAmount > 0 ? Math.max(4, Math.round((mo.tenderAmount / maxTender) * 100)) : 0;
      const pctW = mo.winnerAmount > 0 ? Math.round((mo.winnerAmount / maxTender) * 100) : 0;
      const moDown = (mo.tenderAmount > 0 && mo.winnerAmount > 0)
        ? (((mo.tenderAmount - mo.winnerAmount) / mo.tenderAmount) * 100).toFixed(1) + '%'
        : null;
      const emptyCls = mo.count ? '' : ' empty';
      gridHtml += '<div class="month-cell-card' + emptyCls + '">' +
        '<div class="mc-head"><span class="mc-month">' + mm + ' 月</span>' +
        '<span class="mc-count">' + mo.count + ' 个项目</span></div>' +
        '<div class="mc-bar"><div class="mc-bar-fill" style="width:' + pctT + '%"></div></div>' +
        '<div class="mc-amount">' + formatMoney(mo.tenderAmount) + ' <span class="mc-tag">招标</span></div>' +
        '<div class="mc-bar win"><div class="mc-bar-fill win" style="width:' + pctW + '%"></div></div>' +
        '<div class="mc-amount win">' + formatMoney(mo.winnerAmount) + ' <span class="mc-tag">中标</span></div>' +
        (moDown ? '<div class="mc-down">下浮率 ' + moDown + '</div>' : '') +
        '</div>';
    }
    gridEl.innerHTML = gridHtml;
    if (totalEl) {
      totalEl.innerHTML = '合计：<b>' + months.length + '</b> 个月 · <b>' + stats.totalCount + '</b> 个项目 · ' +
        '招标 <b>' + formatMoney(stats.totalTender) + '</b> · 中标 <b>' + formatMoney(stats.totalWinner) + '</b>';
    }
  }

  function addRecord(fields, url) {
    const rec = {
      id: Date.now() + '_' + Math.random().toString(36).slice(2, 7),
      fields,
      url: url || '',
      addedAt: Date.now(),
      projectType: 'construction', // 抓取/导入的为建设工程项目
    };
    records.push(rec);
    save();
    return rec;
  }

  // 清单类型切换页签（建设工程 / 政府采购），含计数与标题更新
  function renderTypeTabs(root) {
    const tabs = root.querySelector('#type-tabs');
    if (!tabs) return;
    const nC = records.filter((r) => recType(r) === 'construction').length;
    const nG = records.filter((r) => recType(r) === 'government').length;
    tabs.innerHTML =
      '<button class="type-tab' + (currentType === 'construction' ? ' active' : '') + '" data-type="construction">🏗️ 建设工程项目（' + nC + '）</button>' +
      '<button class="type-tab' + (currentType === 'government' ? ' active' : '') + '" data-type="government">🏛️ 政府采购项目（' + nG + '）</button>';
    tabs.querySelectorAll('.type-tab').forEach((b) => {
      b.onclick = () => {
        if (currentType === b.dataset.type) return;
        currentType = b.dataset.type;
        renderTable(root);
      };
    });
    const title = root.querySelector('#table-title');
    if (title) {
      title.textContent = '② ' + TYPE_NAMES[currentType] + '项目清单（共 ' + FIELDS.length + ' 列，已按开标时间升序）';
    }
  }

  function renderTable(root) {
    containerRef = root;
    const wrap = root.querySelector('#table-wrap');
    if (!wrap) return;
    renderTypeTabs(root); // 页签 + 标题随数据/当前类型刷新
    const list = records.filter((r) => recType(r) === currentType);
    if (!list.length) {
      wrap.innerHTML = currentType === 'government'
        ? '<div class="empty-tip">暂无政府采购项目。可在上方「开标日历」选择日期，点「＋ 添加当天项目」新增（添加后自动归入本表）。</div>'
        : '<div class="empty-tip">暂无数据。抓取招投标链接或导入 Excel 后，将按上方字段自动生成一行。</div>';
      renderStats(root);
      return;
    }
    const rows = buildRows(currentType);
    let html = '<table><thead><tr>';
    FIELDS.forEach((f) => (html += '<th>' + escapeHtml(f.label) + '</th>'));
    html += '<th>操作</th></tr></thead><tbody>';
    rows.forEach((row, i) => {
      html += '<tr>';
      // 开标月份：同月相邻行合并为一个单元格（rowspan）；计算列，不可编辑
      if (row.monthStart) {
        html += '<td class="copyable month-cell" rowspan="' + row.monthSpan + '" title="单击复制" data-field="openMonth" data-id="' + row.rec.id + '" data-idx="' + i + '">' + escapeHtml(row.mlabel) + '</td>';
      }
      // 其余字段（含 序号 自动重编号）；开标月份已在上方单独渲染，此处跳过
      FIELDS.forEach((f) => {
        if (f.key === 'openMonth') return;
        // data-field/id/idx 供双击编辑回写；computed/auto 字段不可编辑，但保留 data-field 用于守卫判断
        html += '<td class="copyable" title="单击复制 · 双击编辑" data-field="' + f.key + '" data-id="' + row.rec.id + '" data-idx="' + i + '">' + escapeHtml(cellValue(f, row.rec, i)) + '</td>';
      });
      html += '<td><button class="del" data-id="' + row.rec.id + '">删除</button></td>';
      html += '</tr>';
    });
    html += '</tbody></table>';
    wrap.innerHTML = html;
    if (window.Workbench && Workbench.makeTableResizable) {
      Workbench.makeTableResizable(wrap.querySelector('table'), 'tender');
    }
    wrap.querySelectorAll('.del').forEach((b) => {
      b.onclick = async () => {
        if (!Workbench.ensureEditable()) return;
        records = records.filter((x) => x.id !== b.dataset.id);
        save();
        renderTable(root);
      };
    });
    wrap.querySelectorAll('td.copyable').forEach((td) => {
      // 单击复制：延迟 320ms 执行。双击编辑的第一击也会触发 click，
      // 若不延迟，会把单元格旧内容写进剪贴板，覆盖用户原本要粘贴的文字
      // （空单元格则写入空内容）→ 表现为"粘贴不了"。延迟后双击的第二击
      // 会 clearTimeout 取消本次复制，从而保护用户剪贴板。
      let copyTimer = null;
      td.onclick = () => {
        if (td.querySelector('input.cell-edit')) return; // 编辑态中不触发复制
        clearTimeout(copyTimer);
        copyTimer = setTimeout(async () => {
          const text = td.textContent || '';
          const ok = await copyText(text);
          if (ok) {
            td.classList.add('copied');
            setTimeout(() => td.classList.remove('copied'), 800);
          } else {
            td.title = '复制失败，请手动选择复制';
          }
        }, 320);
      };
      td.ondblclick = () => {
        clearTimeout(copyTimer); // 取消待执行的单击复制，保护剪贴板
        editCell(td);
      };
    });
    renderStats(root); // 数据变化后同步刷新统计卡片
    // 同步刷新嵌入的开标日历（清单数据变化后，日历上的标记需跟着更新）
    if (window.TenderCalendar && window.TenderCalendar.refresh) TenderCalendar.refresh();
  }
  function escapeHtml(s) {
    return String(s).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
  }
  // 复制文本到剪贴板：优先 Clipboard API（localhost/https），
  // file:// 等非安全上下文用临时 textarea + execCommand 兜底
  async function copyText(text) {
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(text);
        return true;
      }
    } catch (e) {}
    try {
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.style.position = 'fixed';
      ta.style.top = '-9999px';
      ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.focus();
      ta.select();
      const ok = document.execCommand('copy');
      document.body.removeChild(ta);
      return ok;
    } catch (e) {
      return false;
    }
  }

  // ============== ⑥ 插件对象 ==============
  // ============== ⑤½ 同步到本地 Excel ==============
  const DEFAULT_EXCEL_PATH = 'E:\\2026工作\\2026年招投标项目交易清单(1).xlsx';
  const LS_EXCEL_PATH = 'wb_tender_excel_path';
  const LS_AUTO_SYNC = 'wb_tender_auto_sync';
  let autoSyncCb = null; // mount 内设置：勾选「自动同步」时触发

  async function syncToLocalExcel(statusEl) {
    if (!statusEl) return false;
    if (!records.length) {
      statusEl.className = 'status err';
      statusEl.textContent = '清单为空，没有可同步的内容。';
      return false;
    }
    const pathInput = containerRef ? containerRef.querySelector('#excel-path') : null;
    const targetPath = (pathInput ? pathInput.value.trim() : '') || DEFAULT_EXCEL_PATH;
    if (pathInput) pathInput.value = targetPath;
    try { localStorage.setItem(LS_EXCEL_PATH, targetPath); } catch (e) { /* 忽略 */ }

    statusEl.className = 'status info';
    statusEl.textContent = '同步中：正在写入 ' + targetPath + ' …';
    try {
      const resp = await fetch('/api/sync-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ path: targetPath, records: records }),
      });
      let data = null;
      try { data = await resp.json(); } catch (e) { /* 非 JSON 响应 */ }
      if (!resp.ok || !data || !data.ok) {
        throw new Error((data && data.error) || ('HTTP ' + resp.status));
      }
      statusEl.className = data.warning ? 'status warn' : 'status ok';
      let msg = '✓ 已同步到本地 Excel：新增 ' + data.added + ' 条、补空更新 ' + data.updated +
        ' 条、无变化跳过 ' + data.skipped + ' 条，表格现共 ' + data.total + ' 条记录。';
      if (data.backup) msg += '（已自动备份：' + data.backup + '）';
      if (data.warning) msg += ' ⚠ ' + data.warning;
      statusEl.textContent = msg;
      return true;
    } catch (e) {
      statusEl.className = 'status err';
      statusEl.textContent = '同步失败：' + (e && e.message ? e.message : String(e)) +
        '。请确认已通过「启动工作台.bat」启动本地服务（直接双击打开 index.html 无法写本地文件）。';
      return false;
    }
  }

  const plugin = {
    id: 'tender-list',
    name: '统计招投标项目清单',
    icon: '📋',
    description: '根据招投标链接，按文件一字段梳理统计项目清单，支持导出与汇报',
    async mount(container) {
      await load();
      container.innerHTML = `
        <div class="plugin-header">
          <div>
            <h2>${this.name}</h2>
            <div class="desc">${this.description}</div>
          </div>
          <div class="toolbar">
            <button class="btn secondary" id="btn-export-csv">⬇ 导出表格(CSV)</button>
            <button class="btn secondary" id="btn-import">📥 导入Excel/CSV</button>
            <button class="btn" id="btn-report">📄 生成汇报材料(Word)</button>
            <button class="btn ghost" id="btn-clear">清空</button>
          </div>
          <input type="file" id="file-input" accept=".xlsx,.xls,.csv" style="display:none" />
        </div>

        <div class="sync-bar">
          <span class="sync-label">💾 同步目标：</span>
          <input type="text" id="excel-path" class="path-input" spellcheck="false"
            placeholder="本地 Excel 完整路径，如 E:\\2026工作\\2026年招投标项目交易清单(1).xlsx" />
          <label class="auto-sync-label" title="每次解析/导入出新记录后，自动把工作台清单同步写入上面的 Excel 文件">
            <input type="checkbox" id="auto-sync" /> 解析后自动同步
          </label>
          <button class="btn secondary" id="btn-sync-excel">💾 同步到本地Excel</button>
        </div>
        <div class="status" id="sync-status"></div>

        <div class="banner">✅ 字段已按文件一《2026年招投标项目交易清单》更新：<b>10 个正式字段</b> + 新增<b>开标月份</b>列（按开标时间升序排列并合并同月单元格，便于按月统计）；解析招标公告时自动关联同项目「中标公示」填入<b>中标单位 / 中标金额</b>；导出 CSV 含开标月份共 <b>11 列</b>；工作台清单可一键<b>同步写入本地 Excel</b>（自动去重、补空、备份）。</div>
        <div class="status" id="import-status"></div>

        <div class="card stats-card">
          <h3>④ 数据统计概览</h3>
          <div class="stats-toolbar">
            <label class="stats-label">统计年份：</label>
            <select id="stats-year" class="year-select"></select>
            <span class="stats-hint">按「开标时间」自动汇总；未注明开标时间的项目不计入统计。</span>
          </div>
          <div class="stats-top">
            <div class="stats-left">
              <div class="stat-cards" id="stat-cards"></div>
              <div class="month-stats">
                <div class="month-stats-head"><span>📅 按月统计</span><span id="month-total"></span></div>
                <div class="month-grid" id="month-grid"></div>
              </div>
            </div>
            <div class="stats-right">
              <div class="stats-right-title">📅 开标日历</div>
              <div id="calendar-section"></div>
            </div>
          </div>
        </div>

        <div class="card">
          <h3>① 批量输入招投标链接</h3>
          <label class="field-label">每行一个链接（也可一行内用空格/逗号分隔多个）。优先经本地代理抓取；双击打开 HTML 时浏览器会跨域拦截，请用「启动工作台.bat」启动本地代理后访问 localhost。</label>
          <textarea id="url-input" placeholder="https://ggzy.hainan.gov.cn/...&#10;https://ggzy.hainan.gov.cn/...&#10;每行一个链接"></textarea>
          <div style="height:10px"></div>
          <div class="btn-row">
            <button class="btn" id="btn-fetch">🔗 批量抓取并解析</button>
            <button class="btn ghost" id="btn-clear-input">清空输入</button>
          </div>
          <div class="status" id="fetch-status"></div>
        </div>

        <div class="card">
          <h3 id="table-title">② 建设工程项目清单（共 ${FIELDS.length} 列，已按开标时间升序）</h3>
          <div class="type-tabs" id="type-tabs"></div>
          <p class="hint-copy">提示：表格已按「开标时间」升序排列，同月项目合并单元格（左侧「开标月份」列），合并单元格所跨行数即该月项目数；<b>单击</b>任意字段单元格可复制该内容，<b>双击</b>单元格可直接编辑/填写信息（回车保存、Esc 取消）。修改「开标时间」后，表格会自动按月份重新排序并重新归并月份。<b>建设工程</b>为抓取/导入的项目；<b>政府采购</b>为日历中手动添加的项目，两表分开统计、分开导出。</p>
          <div class="table-wrap" id="table-wrap"></div>
        </div>
      `;

      const $ = (id) => container.querySelector(id);
      const urlInput = $('#url-input');
      const fetchStatus = $('#fetch-status');

      // ===== 同步到本地 Excel =====
      const syncStatus = $('#sync-status');
      const excelPathInput = $('#excel-path');
      const autoSyncBox = $('#auto-sync');
      try {
        excelPathInput.value = localStorage.getItem(LS_EXCEL_PATH) || DEFAULT_EXCEL_PATH;
        autoSyncBox.checked = localStorage.getItem(LS_AUTO_SYNC) === '1';
      } catch (e) { /* 忽略 */ }
      autoSyncBox.onchange = () => {
        try { localStorage.setItem(LS_AUTO_SYNC, autoSyncBox.checked ? '1' : '0'); } catch (e) { /* 忽略 */ }
      };
      $('#btn-sync-excel').onclick = () => syncToLocalExcel(syncStatus);
      autoSyncCb = () => {
        if (autoSyncBox.checked && records.length) syncToLocalExcel(syncStatus);
      };

      // ===== 统计年份切换 =====
      const yearSel = $('#stats-year');
      if (yearSel) yearSel.onchange = () => { statsYear = Number(yearSel.value); renderStats(container); };

      $('#btn-fetch').onclick = async () => {
        if (!Workbench.ensureEditable()) return;
        const links = splitLinks(urlInput.value);
        if (!links.length) {
          fetchStatus.className = 'status err';
          fetchStatus.textContent = '请至少填写一个有效的 http(s) 链接（每行一个）。';
          return;
        }
        fetchStatus.className = 'status info';
        let okCount = 0;
        let linkedWin = 0;
        const failList = [];
        for (let i = 0; i < links.length; i++) {
          const url = links[i];
          fetchStatus.textContent = '抓取中（' + (i + 1) + '/' + links.length + '）：' + url;
          const r = await fetchPage(url);
          if (!r.ok) {
            failList.push(url + '（' + (r.err || '失败') + '）');
            continue;
          }
          const f = parseTenderPage(r.text, url);
          const rec = addRecord(f, url);
          // 自动关联同项目的中标公示（若该项目已公示中标单位，则填入中标单位/中标金额）
          fetchStatus.textContent = '抓取中（' + (i + 1) + '/' + links.length + '）：解析完成，正在关联中标公示…';
          try {
            const win = await fetchWinningInfo(r.text, url);
            if (win) {
              rec.fields.winner = win.winner || rec.fields.winner;
              rec.fields.winnerAmount = win.winnerAmount || rec.fields.winnerAmount;
              if (win.winner) linkedWin++;
              save();
            }
          } catch (e) { /* 关联失败不影响主记录 */ }
          okCount++;
          renderTable(container);
        }
        let msg;
        if (failList.length === 0) {
          fetchStatus.className = 'status ok';
          msg = '✓ 全部处理完成：成功 ' + okCount + ' 条，清单共 ' + records.length + ' 条。';
        } else {
          fetchStatus.className = 'status warn';
          msg = '✓ 成功 ' + okCount + ' 条，失败 ' + failList.length + ' 条（抓取超时/被拦截）：' +
            failList.join('；') + '。';
        }
        if (linkedWin > 0) msg += '；已自动关联 ' + linkedWin + ' 条中标公示并填入中标单位/中标金额。';
        fetchStatus.textContent = msg;
        if (okCount > 0 && autoSyncCb) autoSyncCb(); // 勾选了自动同步则写入本地 Excel
      };

      $('#btn-clear-input').onclick = () => {
        urlInput.value = '';
        fetchStatus.textContent = '';
      };

      $('#btn-export-csv').onclick = () => {
        if (!records.length) { alert('清单为空，无法导出。'); return; }
        exportCSV();
      };
      $('#btn-report').onclick = () => {
        if (!records.length) { alert('清单为空，无法生成汇报。'); return; }
        generateWord();
      };
      $('#btn-clear').onclick = async () => {
        const inTab = records.filter((r) => recType(r) === currentType);
        if (!inTab.length) { alert('当前「' + TYPE_NAMES[currentType] + '」清单为空。'); return; }
        if (!Workbench.ensureEditable()) return;
        if (confirm('确定清空「' + TYPE_NAMES[currentType] + '」清单的全部 ' + inTab.length + ' 条记录？（另一张表不受影响）')) {
          records = records.filter((r) => recType(r) !== currentType);
          save(); renderTable(container);
        }
      };

      // 导入 Excel/CSV
      const fileInput = $('#file-input');
      const importStatus = $('#import-status');
      $('#btn-import').onclick = async () => {
        if (!Workbench.ensureEditable()) return;
        fileInput.click();
      };
      fileInput.onchange = (e) => {
        const file = e.target.files && e.target.files[0];
        if (file) importFile(file, importStatus);
        e.target.value = ''; // 允许重复选择同一文件
      };

      // ===== 开标日历（嵌入本页，与清单共享同一份 records） =====
      if (window.TenderCalendar) {
        TenderCalendar.mount(container.querySelector('#calendar-section'), {
          getRecords: () => records,
          onSaved: () => { save(); renderTable(container); },
        });
      }

      renderTable(container);
      maybeSeed(container);
    },
  };

  if (window.Workbench) window.Workbench.register(plugin);
})();
