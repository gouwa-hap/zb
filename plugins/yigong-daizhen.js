/**
 * 插件：以工代赈项目调度
 * --------------------------------------------------
 * 字段按《省级财政以工代赈项目进展情况调度表》原表，共 20 列：
 *   序号(自动) | 市县 | 项目名称 | 资金类型 | 下达资金额度 | 项目投资情况(3) | 项目形象进度(4)
 *   | 资金支付进度(2) | 政策落实情况(4) | 主要建设内容及成果 | 备注
 * 能力：Excel/CSV 导入（自动识别三级表头、跳过总计行）、双击编辑、进度列单击打勾、
 *       新增/删除行、表尾合计、顶部统计卡、导出 CSV、导出 Excel（还原原表三级表头+合并单元格）。
 */
(function () {
  // ============== ① 字段定义 ==============
  // type: money=金额(万元，两位小数) / int=人数(整数) / check=打勾 / 其余为文本
  // sum:  参与表尾合计与顶部统计
  const FIELDS = [
    { key: 'seq',             label: '序号', auto: true, w: 50 },
    { key: 'city',            label: '市县', w: 90 },
    { key: 'projectName',     label: '项目名称', w: 200 },
    { key: 'fundType',        label: '资金类型（省级财政）', w: 170 },
    { key: 'fundQuota',       label: '下达以工代赈资金额度（万元）', type: 'money', sum: true, w: 120 },
    { key: 'investTotal',     label: '项目总投资（万元）', type: 'money', sum: true, w: 110 },
    { key: 'investYigong',    label: '其中：实际用于实施以工代赈项目的资金（万元）', type: 'money', sum: true, w: 130 },
    { key: 'investIntegrate', label: '其中：整合资金（万元）', type: 'money', sum: true, w: 120 },
    { key: 'progPrep',        label: '前期工作', type: 'check', w: 62 },
    { key: 'progBuild',       label: '主体施工', type: 'check', w: 62 },
    { key: 'progDone',        label: '主体完工', type: 'check', w: 62 },
    { key: 'progUsed',        label: '建成投用', type: 'check', w: 62 },
    { key: 'payTotal',        label: '已支付资金总额（万元）', type: 'money', sum: true, w: 120 },
    { key: 'payProvince',     label: '其中：已支付省级财政金额（万元）', type: 'money', sum: true, w: 130 },
    { key: 'wagePaid',        label: '已发放劳务报酬（万元）', type: 'money', sum: true, w: 120 },
    { key: 'workerHired',     label: '已吸纳农村群众务工（人数）', type: 'int', sum: true, w: 110 },
    { key: 'publicPosts',     label: '已设置公益性岗位（个数）', type: 'int', sum: true, w: 110 },
    { key: 'trained',         label: '已组织劳动力技能培训（人数）', type: 'int', sum: true, w: 110 },
    { key: 'content',         label: '主要建设内容及成果', w: 320, long: true },
    { key: 'remarks',         label: '备注', w: 110 },
  ];

  // 两级表头：rowspan:2 的列独占两行；cols 为分组下的子列名（顺序须与 FIELDS 一致）
  const HEAD_TOP = [
    { label: '序号', rowspan: 2 },
    { label: '市县', rowspan: 2 },
    { label: '项目名称', rowspan: 2 },
    { label: '资金类型\n（省级财政）', rowspan: 2 },
    { label: '下达的以工代赈\n资金额度（万元）', rowspan: 2 },
    { label: '项目投资情况', cols: ['项目总投资\n（万元）', '其中：实际用于实施\n以工代赈项目的资金（万元）', '其中：整合资金\n（万元）'] },
    { label: '项目形象进度', cols: ['前期\n工作', '主体\n施工', '主体\n完工', '建成\n投用'] },
    { label: '资金支付进度', cols: ['已支付资金总额\n（万元）', '其中：已支付省级\n财政金额（万元）'] },
    { label: '政策落实情况', cols: ['已发放劳务报酬\n（万元）', '已吸纳农村群众\n务工（人数）', '已设置公益性\n岗位（个数）', '已组织劳动力技能\n培训（人数）'] },
    { label: '主要建设内容及成果', rowspan: 2 },
    { label: '备注', rowspan: 2 },
  ];

  const STORE_KEY = 'wb_yigong_records';
  let records = [];
  const FIELD_MAP = {};
  FIELDS.forEach((f) => { FIELD_MAP[f.key] = f; });

  // ============== ② 数值 / 显示辅助 ==============
  function numOf(v) {
    if (v == null) return 0;
    const s = String(v).trim();
    if (!s) return 0;
    const n = Number(s.replace(/,/g, '').replace(/[^\d.\-]/g, ''));
    return isNaN(n) ? 0 : n;
  }
  // 显示值：金额两位小数、人数取整、打勾显示 √
  function cellText(f, rec, i) {
    if (f.auto) return String(i + 1);
    const v = (rec.fields && rec.fields[f.key]) || '';
    if (f.type === 'check') return v ? '√' : '';
    if (f.type === 'money' && v !== '') { const n = Number(v); return isNaN(n) ? String(v) : n.toFixed(2); }
    if (f.type === 'int' && v !== '') { const n = Number(v); return isNaN(n) ? String(v) : String(Math.round(n)); }
    return String(v);
  }
  // 导出到 Excel 的原始值：金额/人数用数字（便于再计算），打勾用 √
  function rawVal(f, rec, i) {
    if (f.auto) return i + 1;
    const v = (rec.fields && rec.fields[f.key]) || '';
    if (f.type === 'check') return v ? '√' : '';
    if ((f.type === 'money' || f.type === 'int') && v !== '') {
      const n = Number(v);
      if (!isNaN(n)) return f.type === 'int' ? Math.round(n) : n;
    }
    return v;
  }
  function normCheck(v) {
    const s = String(v == null ? '' : v).trim();
    if (!s) return '';
    if (/[√✓✔☑]/.test(s)) return '√';
    if (/^(是|有|已|完成|1|true|y|yes)$/i.test(s)) return '√';
    return '';
  }
  // 导入时按字段类型归一化单元格值
  function normVal(key, v) {
    const f = FIELD_MAP[key];
    if (f && f.type === 'check') return normCheck(v);
    if (v == null) return '';
    const s = String(v).trim();
    if (!s) return '';
    if (f && (f.type === 'money' || f.type === 'int')) {
      if (/^[-—－/\\.\s]*$/.test(s)) return ''; // "-" "—" "/" 等占位符视为空
      const n = Number(s.replace(/,/g, '').replace(/[^\d.\-]/g, ''));
      if (isNaN(n)) return '';
      // 原表金额列多为公式单元格，raw 读出的是浮点全精度值（如 172.480102），
      // 显示的却是两位小数。故金额统一四舍五入到 2 位、人数取整，避免导入后数值失真。
      return f.type === 'money' ? String(Math.round(n * 100) / 100) : String(Math.round(n));
    }
    return s;
  }

  // ============== ③ 导入：三级表头识别 ==============
  // 原表表头跨 3 行且带合并单元格：合并区只有首格有值，故需把 3 行拼成一个完整标题再匹配
  const MATCHERS = [
    ['seq',             (s) => /^序\s*号$/.test(s)],
    ['city',            (s) => /市\s*县|县\s*市|市\s*区/.test(s)],
    ['projectName',     (s) => s.includes('项目名称')],
    ['fundType',        (s) => s.includes('资金类型')],
    ['fundQuota',       (s) => s.includes('下达') && s.includes('资金额度')],
    ['investYigong',    (s) => s.includes('实际用于') || (s.includes('其中') && s.includes('以工代赈项目'))],
    ['investIntegrate', (s) => s.includes('整合资金')],
    ['investTotal',     (s) => s.includes('项目总投资') || s.includes('总投资')],
    ['progPrep',        (s) => s.includes('前期')],
    ['progBuild',       (s) => s.includes('主体施工')],
    ['progDone',        (s) => s.includes('主体完工')],
    ['progUsed',        (s) => s.includes('建成投用')],
    ['payProvince',     (s) => s.includes('已支付') && (s.includes('省级') || s.includes('财政'))],
    ['payTotal',        (s) => s.includes('已支付')],
    ['wagePaid',        (s) => s.includes('劳务报酬')],
    ['workerHired',     (s) => s.includes('吸纳') && (s.includes('务工') || s.includes('群众'))],
    ['publicPosts',     (s) => s.includes('公益性岗位')],
    ['trained',         (s) => s.includes('培训')],
    ['content',         (s) => s.includes('建设内容') || s.includes('建设成果')],
    ['remarks',         (s) => /^备\s*注/.test(s)],
  ];
  function flatHeader(rows, from, colCount) {
    const out = [];
    for (let c = 0; c < colCount; c++) {
      let s = '';
      for (let k = 0; k < 3; k++) {
        const r = rows[from + k] || [];
        s += String(r[c] == null ? '' : r[c]);
      }
      out.push(s.replace(/\s/g, '')); // 去换行/空格后匹配，"项目\n总投资" → "项目总投资"
    }
    return out;
  }
  function parseScheduleMatrix(matrix) {
    const rows = (matrix || []).filter((r) => Array.isArray(r));
    if (!rows.length) return { list: [], warn: '表格为空。' };
    const colCount = rows.reduce((m, r) => Math.max(m, r.length), 0);
    // 定位表头行：前 12 行内找含「项目名称」的行
    let hIdx = -1;
    for (let i = 0; i < Math.min(12, rows.length); i++) {
      if (rows[i].some((c) => String(c == null ? '' : c).includes('项目名称'))) { hIdx = i; break; }
    }
    if (hIdx < 0) return { list: [], warn: '未找到含「项目名称」的表头行，请确认导入的是调度表。' };

    const titles = flatHeader(rows, hIdx, colCount);
    const colKey = titles.map((t) => {
      if (!t) return '';
      for (let i = 0; i < MATCHERS.length; i++) {
        if (MATCHERS[i][1](t)) return MATCHERS[i][0];
      }
      return '';
    });
    if (!colKey.includes('projectName')) {
      return { list: [], warn: '表头未能识别「项目名称」列，导入中止。' };
    }

    const list = [];
    // 数据行从三级表头之后开始
    for (let i = hIdx + 3; i < rows.length; i++) {
      const row = rows[i];
      if (!row.some((c) => String(c == null ? '' : c).trim())) continue; // 跳过空行
      const flat = row.map((c) => String(c == null ? '' : c).replace(/\s/g, ''));
      if (flat.some((s) => /^(总\s*计|合\s*计|小\s*计)$/.test(s))) continue; // 跳过总计/合计行
      const f = {};
      colKey.forEach((k, ci) => { if (k) f[k] = normVal(k, row[ci]); });
      if (!f.projectName) continue; // 无项目名称视为无效行
      list.push(f);
    }
    return { list, warn: '' };
  }
  function dedupeKey(f) {
    return ((f && f.city) || '').trim() + '|' + ((f && f.projectName) || '').trim();
  }
  function mergeRecords(list) {
    let added = 0, skipped = 0;
    const seen = new Set(records.map((r) => dedupeKey(r.fields)));
    (list || []).forEach((f) => {
      if (!f || !f.projectName) return;
      const key = dedupeKey(f);
      if (seen.has(key)) { skipped++; return; }
      seen.add(key);
      addRecord(f);
      added++;
    });
    return { added, skipped };
  }
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
    return rows;
  }
  async function importFile(file, statusEl) {
    statusEl.className = 'status info';
    statusEl.textContent = '解析中…';
    try {
      const buf = await file.arrayBuffer();
      let matrix;
      if (/\.csv$/i.test(file.name)) {
        matrix = parseCsvText(new TextDecoder('utf-8').decode(buf));
      } else {
        if (!window.XLSX) { statusEl.className = 'status err'; statusEl.textContent = 'Excel 解析库未加载，请刷新页面后重试。'; return; }
        const wb = window.XLSX.read(buf, { type: 'array' });
        // raw:true 读原始值，避免日期/数字被二次格式化
        matrix = window.XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '', raw: true });
      }
      const res2 = parseScheduleMatrix(matrix);
      if (res2.warn && !res2.list.length) {
        statusEl.className = 'status err';
        statusEl.textContent = '✗ ' + res2.warn;
        return;
      }
      const res = mergeRecords(res2.list);
      statusEl.className = res.added ? 'status ok' : 'status warn';
      statusEl.textContent = '✓ 导入完成：新增 ' + res.added + ' 条' +
        (res.skipped ? '，跳过重复 ' + res.skipped + ' 条' : '') +
        '；调度表共 ' + records.length + ' 条。' +
        (res2.warn ? ' ⚠ ' + res2.warn : '');
      renderTable(containerRef);
    } catch (err) {
      statusEl.className = 'status err';
      statusEl.textContent = '✗ 解析失败：' + (err && err.message ? err.message : err);
    }
  }

  // ============== ④ 存储（服务端持久化，替代浏览器本地） ==============
  function save() {
    // 写入服务端（并自动在本机缓存兜底），数据不再只存于某一台浏览器
    Workbench.saveRecords(STORE_KEY, records);
  }
  async function load() {
    records = await Workbench.loadRecords(STORE_KEY);
  }
  function addRecord(fields) {
    const rec = {
      id: Date.now() + '_' + Math.random().toString(36).slice(2, 7),
      fields: fields || {},
      addedAt: Date.now(),
    };
    records.push(rec);
    save();
    return rec;
  }

  // ============== ⑤ 统计（表尾合计 + 顶部卡片） ==============
  function computeSums() {
    const s = { count: records.length };
    FIELDS.forEach((f) => {
      if (!f.sum) return;
      let v = 0;
      records.forEach((r) => { v += numOf(r.fields[f.key]); });
      s[f.key] = v;
    });
    return s;
  }
  function money(n) {
    return Number(n || 0).toFixed(2);
  }
  function renderStats(root) {
    const el = root.querySelector('#yg-cards');
    if (!el) return;
    const s = computeSums();
    const payPct = s.fundQuota > 0 ? ((s.payTotal / s.fundQuota) * 100).toFixed(1) + '%' : '—';
    const wagePct = s.payTotal > 0 ? ((s.wagePaid / s.payTotal) * 100).toFixed(1) + '%' : '—';
    const card = (cls, num, label, sub) =>
      '<div class="stat-card ' + cls + '"><div class="stat-num">' + num + '</div>' +
      '<div class="stat-label">' + label + '</div>' +
      (sub ? '<div class="stat-sub">' + sub + '</div>' : '') + '</div>';
    el.innerHTML =
      card('primary', String(s.count), '调度项目总数（个）', '') +
      card('accent', money(s.fundQuota), '下达以工代赈资金额度合计（万元）', '') +
      card('teal', money(s.payTotal), '已支付资金总额合计（万元）', '资金支付进度 ' + payPct + '（占下达额度）') +
      card('purple', money(s.wagePaid), '已发放劳务报酬合计（万元）',
        '占已支付 ' + wagePct + ' · 务工 ' + Math.round(s.workerHired) + ' 人 · 公益岗 ' +
        Math.round(s.publicPosts) + ' 个 · 培训 ' + Math.round(s.trained) + ' 人');
  }

  // ============== ⑥ 渲染表格 ==============
  let containerRef = null;
  function escapeHtml(s) {
    return String(s == null ? '' : s).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
  }
  function buildHeadHtml() {
    let h1 = '', h2 = '';
    HEAD_TOP.forEach((g) => {
      if (g.rowspan === 2) {
        h1 += '<th class="th-group" rowspan="2">' + escapeHtml(g.label).replace(/\n/g, '<br>') + '</th>';
      } else {
        h1 += '<th class="th-group" colspan="' + g.cols.length + '">' + escapeHtml(g.label) + '</th>';
        g.cols.forEach((c) => { h2 += '<th>' + escapeHtml(c).replace(/\n/g, '<br>') + '</th>'; });
      }
    });
    return '<thead><tr>' + h1 + '<th class="th-op" rowspan="2">操作</th></tr><tr>' + h2 + '</tr></thead>';
  }
  function renderTable(root) {
    containerRef = root;
    const wrap = root.querySelector('#yg-table-wrap');
    if (!wrap) return;
    renderStats(root);
    if (!records.length) {
      wrap.innerHTML = '<div class="empty-tip">暂无数据。点击「＋ 新增一行」手工录入，或用「📥 导入Excel/CSV」导入调度表。</div>';
      return;
    }
    let html = '<table class="yg-table">' + buildHeadHtml() + '<tbody>';
    records.forEach((rec, i) => {
      html += '<tr>';
      FIELDS.forEach((f) => {
        const cls = [];
        if (f.auto) cls.push('col-seq');
        if (f.type === 'check') cls.push('check-cell');
        if (f.type === 'money' || f.type === 'int') cls.push('num-cell');
        if (f.long) cls.push('long-cell');
        if (!f.auto) cls.push('copyable');
        const title = f.type === 'check' ? '单击打勾 / 取消' : '单击复制 · 双击编辑';
        html += '<td class="' + cls.join(' ') + '" title="' + title + '"' +
          ' data-field="' + f.key + '" data-id="' + rec.id + '" data-idx="' + i + '">' +
          escapeHtml(cellText(f, rec, i)) + '</td>';
      });
      html += '<td class="op-cell"><button class="del" data-id="' + rec.id + '">删除</button></td>';
      html += '</tr>';
    });
    html += '</tbody>';
    // 表尾合计
    const s = computeSums();
    html += '<tfoot><tr>';
    FIELDS.forEach((f) => {
      if (f.auto) { html += '<td class="col-seq">合计</td>'; return; }
      if (f.sum) { html += '<td class="num-cell">' + (f.type === 'int' ? Math.round(s[f.key]) : money(s[f.key])) + '</td>'; return; }
      html += '<td></td>';
    });
    html += '<td></td></tr></tfoot></table>';
    wrap.innerHTML = html;
    if (window.Workbench && Workbench.makeTableResizable) {
      Workbench.makeTableResizable(wrap.querySelector('table'), 'yigong');
    }

    wrap.querySelectorAll('.del').forEach((b) => {
      b.onclick = async () => {
        if (!Workbench.ensureEditable()) return;
        records = records.filter((x) => x.id !== b.dataset.id);
        save();
        renderTable(root);
      };
    });
    wrap.querySelectorAll('td').forEach((td) => {
      const field = td.dataset.field;
      if (!field) return;
      const f = FIELD_MAP[field];
      if (!f) return;
      if (f.type === 'check') {
        // 进度列：单击切换打勾
        td.onclick = () => {
          if (!Workbench.ensureEditable()) return; // 未进入修改模式禁止打勾
          const rec = records.find((r) => r.id === td.dataset.id);
          if (!rec) return;
          rec.fields[field] = rec.fields[field] ? '' : '√';
          save();
          renderTable(root);
        };
        td.ondblclick = (e) => e.preventDefault();
        return;
      }
      if (f.auto) return;
      // 单击复制：延迟 320ms，避免双击编辑的第一击把旧内容写进剪贴板
      let timer = null;
      td.onclick = () => {
        if (td.querySelector('input.cell-edit, textarea.cell-edit')) return;
        clearTimeout(timer);
        timer = setTimeout(async () => {
          const ok = await copyText(td.textContent || '');
          if (ok) {
            td.classList.add('copied');
            setTimeout(() => td.classList.remove('copied'), 800);
          }
        }, 320);
      };
      td.ondblclick = () => { clearTimeout(timer); editCell(td); };
    });
  }
  function editCell(td) {
    if (!Workbench.ensureEditable()) return; // 未进入修改模式禁止编辑
    const field = td.dataset.field;
    const id = td.dataset.id;
    const idx = Number(td.dataset.idx || '0');
    if (!field || !id) return;
    const f = FIELD_MAP[field];
    if (!f || f.auto) return;
    const rec = records.find((r) => r.id === id);
    if (!rec) return;
    if (td.querySelector('input.cell-edit, textarea.cell-edit')) return;
    const current = (rec.fields && rec.fields[field]) || '';
    td.classList.add('editing');
    td.textContent = '';
    // 长文本用多行输入框，其余用单行
    const el = document.createElement(f.long ? 'textarea' : 'input');
    if (!f.long) el.type = 'text';
    el.className = 'cell-edit';
    el.value = current;
    if (f.long) el.rows = 6;
    td.appendChild(el);
    el.focus();
    if (el.select) el.select();

    let done = false;
    const finish = (keep) => {
      if (done) return;
      done = true;
      // 手工录入同样走归一化：金额保留两位小数、人数取整，与导入口径一致
      const val = f.type === 'check' ? normCheck(el.value) : normVal(f.key, el.value);
      if (keep && val !== (rec.fields[field] || '')) {
        rec.fields[field] = val;
        save();
      }
      if (containerRef) renderTable(containerRef);
    };
    el.addEventListener('keydown', (e) => {
      if (f.long) {
        if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); finish(true); }
        else if (e.key === 'Escape') { e.preventDefault(); finish(false); }
      } else {
        if (e.key === 'Enter') { e.preventDefault(); finish(true); }
        else if (e.key === 'Escape') { e.preventDefault(); finish(false); }
      }
    });
    el.addEventListener('blur', () => finish(true));
  }
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
      ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.select();
      const ok = document.execCommand('copy');
      document.body.removeChild(ta);
      return ok;
    } catch (e) { return false; }
  }

  // ============== ⑦ 导出 ==============
  function download(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename; a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
  function exportCsv() {
    const head1 = [], head2 = [];
    HEAD_TOP.forEach((g) => {
      if (g.rowspan === 2) { head1.push(g.label.replace(/\n/g, '')); head2.push(''); }
      else g.cols.forEach((c) => { head1.push(g.label); head2.push(c.replace(/\n/g, '')); });
    });
    const esc = (v) => {
      let s = String(v == null ? '' : v);
      if (/[",\n\r]/.test(s)) s = '"' + s.replace(/"/g, '""') + '"';
      return s;
    };
    let csv = '\ufeff' + head1.map(esc).join(',') + '\n' + head2.map(esc).join(',') + '\n';
    records.forEach((rec, i) => {
      csv += FIELDS.map((f) => esc(cellText(f, rec, i))).join(',') + '\n';
    });
    download(new Blob([csv], { type: 'text/csv;charset=utf-8;' }), '以工代赈项目调度表.csv');
  }
  // 导出 Excel：还原原表三级表头 + 合并单元格 + 总计行
  function buildAoa() {
    const now = new Date();
    const title = now.getFullYear() + '年省级财政以工代赈项目进展情况调度表（' +
      now.getFullYear() + '年' + (now.getMonth() + 1) + '月更新）';
    const aoa = [];
    aoa.push(['附件2', '', '', '']);
    aoa.push([title]);
    const dateRow = new Array(FIELDS.length).fill('');
    dateRow[4] = '填报日期：';
    dateRow[8] = '填报人：';
    aoa.push(dateRow);
    // 三级表头（与原表一致）
    aoa.push(['序号', '市县', '项目名称', '资金类型\n(省级财政）', '下达的以工代赈资金额度（万元）',
      '项目投资情况', '', '', '项目进展情况', '', '', '', '', '', '政策落实情况', '', '', '', '主要建设内容及成果', '备注']);
    aoa.push(['', '', '', '', '', '项目\n总投资（万元）', '', '',
      '形象进度（在对应选项中打“√”）', '', '', '', '资金支付进度', '',
      '目前，已发放劳务报酬（万元）', '目前，已吸纳农村群众务工情况\n（人数，非人次）',
      '目前，已设置公益性岗位个数（人数，非人次）', '目前，已组织开展劳动力技能培训情况（人数，非人次）', '', '']);
    aoa.push(['', '', '', '', '', '',
      '其中：实际用于实施以工代赈项目的资金（万元）', '其中：整合资金（万元）（含地方投资、社会投资、其他涉农资金等）',
      '前期\n工作', '主体\n施工', '主体\n完工', '建成\n投用',
      '已支付资金总额（万元）', '其中：已支付省级财政金额（万元）', '', '', '', '', '', '']);
    records.forEach((rec, i) => {
      aoa.push(FIELDS.map((f) => rawVal(f, rec, i)));
    });
    // 总计行
    const s = computeSums();
    const tot = new Array(FIELDS.length).fill('');
    tot[0] = '总计';
    FIELDS.forEach((f, i) => {
      if (!f.sum) return;
      tot[i] = f.type === 'int' ? Math.round(s[f.key]) : Number(s[f.key].toFixed(2));
    });
    aoa.push(tot);
    return aoa;
  }
  function exportExcel() {
    if (!window.XLSX) { alert('Excel 解析库未加载，请刷新页面后重试。'); return; }
    const aoa = buildAoa();
    const ws = window.XLSX.utils.aoa_to_sheet(aoa);
    const H = 3; // 表头起始行索引（0 起）：0=附件2，1=标题，2=填报信息
    const m = (r1, c1, r2, c2) => ({ s: { r: r1, c: c1 }, e: { r: r2, c: c2 } });
    ws['!merges'] = [
      m(0, 0, 0, 3), m(1, 0, 1, 19),
      m(2, 0, 2, 3), m(2, 4, 2, 5), m(2, 8, 2, 9), m(2, 18, 2, 19),
      // 表头第一级
      m(H, 5, H, 7), m(H, 8, H, 13), m(H, 14, H, 17),
      // 表头第二级
      m(H + 1, 6, H + 1, 7), m(H + 1, 8, H + 1, 11), m(H + 1, 12, H + 1, 13),
      // 表头纵向合并
      m(H, 0, H + 2, 0), m(H, 1, H + 2, 1), m(H, 2, H + 2, 2), m(H, 3, H + 2, 3), m(H, 4, H + 2, 4),
      m(H, 5, H + 1, 5),
      m(H, 14, H + 1, 14), m(H, 15, H + 1, 15), m(H, 16, H + 1, 16), m(H, 17, H + 1, 17),
      m(H, 18, H + 2, 18), m(H, 19, H + 2, 19),
    ];
    ws['!cols'] = FIELDS.map((f) => ({ wch: Math.max(8, Math.round((f.w || 100) / 8)) }));
    ws['!rows'] = [{ hpt: 18 }, { hpt: 26 }, { hpt: 18 }, { hpt: 34 }, { hpt: 34 }, { hpt: 46 }];
    const wb = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(wb, ws, '汇总');
    window.XLSX.writeFile(wb, '以工代赈项目进展情况调度表.xlsx');
  }

  // ============== ⑧ 插件注册 ==============
  const plugin = {
    id: 'yigong-daizhen',
    name: '以工代赈项目调度',
    icon: '🏗️',
    description: '按《省级财政以工代赈项目进展情况调度表》字段管理项目，支持导入、编辑与导出',
    async mount(container) {
      await load();
      container.innerHTML = `
        <div class="plugin-header">
          <div>
            <h2>${this.name}</h2>
            <div class="desc">${this.description}</div>
          </div>
          <div class="toolbar">
            <button class="btn" id="yg-add">＋ 新增一行</button>
            <button class="btn secondary" id="yg-import">📥 导入Excel/CSV</button>
            <button class="btn secondary" id="yg-export-xlsx">⬇ 导出Excel（原表格式）</button>
            <button class="btn secondary" id="yg-export-csv">⬇ 导出CSV</button>
            <button class="btn ghost" id="yg-clear">清空</button>
          </div>
          <input type="file" id="yg-file" accept=".xlsx,.xls,.csv" style="display:none" />
        </div>

        <div class="status" id="yg-status"></div>

        <div class="banner">📌 字段严格按《省级财政以工代赈项目进展情况调度表》原表 <b>20 列</b>设计：项目投资情况(3 列) · 项目形象进度(4 列) · 资金支付进度(2 列) · 政策落实情况(4 列)。<b>进度列单击即可打勾/取消</b>；其它列<b>单击复制、双击编辑</b>（「主要建设内容及成果」为多行，Ctrl+Enter 保存）。导入会自动识别原表的三级表头并<b>跳过总计行与空行</b>；导出 Excel 会还原原表表头与合并单元格。</div>

        <div class="card stats-card">
          <h3>① 调度情况统计</h3>
          <div class="stat-cards" id="yg-cards"></div>
        </div>

        <div class="card">
          <h3>② 项目调度表</h3>
          <p class="hint-copy">提示：<b>进度列（前期工作/主体施工/主体完工/建成投用）单击打勾</b>；其余列<b>单击复制、双击编辑</b>。表尾自动汇总各金额与人数列；导出 Excel 会还原成上级要求的原表格式（含三级表头、合并单元格、总计行）。</p>
          <div class="table-wrap" id="yg-table-wrap"></div>
        </div>
      `;

      const $ = (id) => container.querySelector(id);
      const statusEl = $('#yg-status');

      $('#yg-add').onclick = async () => {
        if (!Workbench.ensureEditable()) return;
        addRecord({});
        renderTable(container);
        statusEl.className = 'status ok';
        statusEl.textContent = '✓ 已新增一行（共 ' + records.length + ' 条），双击单元格填写内容。';
      };
      $('#yg-clear').onclick = async () => {
        if (!records.length) return;
        if (!Workbench.ensureEditable()) return;
        if (confirm('确定清空全部 ' + records.length + ' 条记录？此操作不可撤销。')) {
          records = []; save(); renderTable(container);
          statusEl.className = 'status info';
          statusEl.textContent = '已清空。';
        }
      };
      $('#yg-export-csv').onclick = () => {
        if (!records.length) { alert('暂无数据，无法导出。'); return; }
        exportCsv();
      };
      $('#yg-export-xlsx').onclick = () => {
        if (!records.length) {
          if (!confirm('当前没有数据，是否导出一份空白模板（含原表三级表头）？')) return;
        }
        exportExcel();
        statusEl.className = 'status ok';
        statusEl.textContent = '✓ 已导出 Excel（原表格式，含三级表头与总计行）。';
      };
      const fileInput = $('#yg-file');
      $('#yg-import').onclick = async () => {
        if (!Workbench.ensureEditable()) return;
        fileInput.click();
      };
      fileInput.onchange = (e) => {
        const file = e.target.files && e.target.files[0];
        if (file) importFile(file, statusEl);
        e.target.value = ''; // 允许重复选择同一文件
      };

      renderTable(container);
    },
  };

  if (window.Workbench) window.Workbench.register(plugin);
})();
