/**
 * 模块：开标日历（嵌入「数据统计概览」卡片右区，非独立插件）
 * --------------------------------------------------
 * 紧凑型月历（参考经典日期选择器样式）：
 *   - 每格只显示日期数字；当天有开标项目的，数字下方显示一个小蓝点；
 *   - 今天 = 蓝色描边；点击选中 = 蓝底白字；
 *   - 日历下方固定一栏「项目名称」，列出选中日期当天的项目（可双击改名）。
 * 数据由宿主插件（tender-list）通过 api 提供，与本清单共享同一份 records。
 *
 * 用法（由 tender-list.js 调用）：
 *   window.TenderCalendar.mount(container, { getRecords, onSaved });
 *   window.TenderCalendar.refresh();
 */
(function () {
  // 当前显示年月 —— 默认始终为「今天所在的年月」
  let curYear = new Date().getFullYear();
  let curMonth = new Date().getMonth() + 1; // 1-12

  let selectedKey = '';     // 当前选中的日期 key（YYYY-MM-DD），默认今天
  let hostContainer = null; // 日历挂载的 DOM 容器
  let api = null;           // 宿主提供的数据/保存回调

  // ============== ① 日期解析工具（兼容清单 openTime 的多种格式） ==============
  function parseOpenTime(s) {
    if (s == null) return null;
    const str = String(s).trim();
    if (!str) return null;
    // Excel 日期序列数字
    const n = Number(str);
    if (!isNaN(n) && n > 10000 && n < 80000 && /^\d{4,5}(\.\d+)?$/.test(str)) {
      return new Date((n - 25569) * 86400000);
    }
    const m = str.match(/(\d{4})[-年](\d{1,2})[-月](\d{1,2})(?:[日\s]*(\d{1,2})?[:时]?(\d{1,2})?)?/);
    if (!m) return null;
    const dt = new Date(+m[1], +m[2] - 1, +m[3], m[4] ? +m[4] : 0, m[5] ? +m[5] : 0);
    return isNaN(dt.getTime()) ? null : dt;
  }

  function pad2(n) { return String(n).padStart(2, '0'); }

  function formatDateKey(dt) {
    if (!dt || isNaN(dt.getTime())) return '';
    return dt.getFullYear() + '-' + pad2(dt.getMonth() + 1) + '-' + pad2(dt.getDate());
  }

  function getRecords() {
    if (api && typeof api.getRecords === 'function') {
      const arr = api.getRecords();
      return Array.isArray(arr) ? arr : [];
    }
    return [];
  }

  // 按日期分组：{ '2026-09-15': [rec, ...] }（同一天按开始时间升序）
  function groupByDate() {
    const map = {};
    getRecords().forEach((rec) => {
      const dt = parseOpenTime((rec.fields && rec.fields.openTime) || '');
      if (!dt) return;
      const key = formatDateKey(dt);
      if (!key) return;
      if (!map[key]) map[key] = [];
      map[key].push(rec);
    });
    Object.keys(map).forEach((k) => {
      map[k].sort((a, b) => {
        const da = parseOpenTime(a.fields.openTime);
        const db = parseOpenTime(b.fields.openTime);
        return (da ? da.getTime() : 0) - (db ? db.getTime() : 0);
      });
    });
    return map;
  }

  // ============== ② 渲染 ==============
  function escapeHtml(s) {
    return String(s).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
  }

  function buildCalendarGrid(year, month) {
    const startWeekDay = new Date(year, month - 1, 1).getDay(); // 0=周日
    const daysInMonth = new Date(year, month, 0).getDate();
    const cells = [];
    const prevMonthDays = new Date(year, month - 1, 0).getDate();
    for (let i = startWeekDay - 1; i >= 0; i--) {
      cells.push({ day: prevMonthDays - i, type: 'prev', month: month - 1, year: month === 1 ? year - 1 : year });
    }
    for (let d = 1; d <= daysInMonth; d++) {
      cells.push({ day: d, type: 'current', month: month, year: year });
    }
    const tail = (7 - (cells.length % 7)) % 7;
    for (let d = 1; d <= tail; d++) {
      cells.push({ day: d, type: 'next', month: month + 1, year: month === 12 ? year + 1 : year });
    }
    return cells;
  }

  function renderDays() {
    const wrap = hostContainer && hostContainer.querySelector('#cal-days');
    if (!wrap) return;
    const byDate = groupByDate();
    const cells = buildCalendarGrid(curYear, curMonth);
    const todayKey = formatDateKey(new Date());
    if (!selectedKey) selectedKey = todayKey;

    let html = '';
    cells.forEach((cell) => {
      const key = formatDateKey(new Date(cell.year, cell.month - 1, cell.day));
      const hasEvents = (byDate[key] || []).length > 0;
      const cls = ['cal-day',
        cell.type === 'current' ? '' : 'adjacent',
        key === todayKey ? 'today' : '',
        key === selectedKey ? 'selected' : '',
        hasEvents ? 'has-events' : ''].filter(Boolean).join(' ');

      html += '<div class="' + cls + '" data-key="' + key + '">' +
        '<span class="cal-day-num">' + cell.day + '</span>' +
        (hasEvents ? '<span class="cal-dot"></span>' : '') +
        '</div>';
    });
    wrap.innerHTML = html;

    // 点击日期 → 选中并刷新下方项目栏
    wrap.querySelectorAll('.cal-day').forEach((el) => {
      el.onclick = () => {
        selectedKey = el.dataset.key;
        // 点击相邻月日期时顺带翻页，保持日历与选中日期同月
        const sel = selectedKey.split('-');
        if (+sel[0] !== curYear || +sel[1] !== curMonth) {
          curYear = +sel[0];
          curMonth = +sel[1];
          syncSelectors();
        }
        renderDays();
      };
    });

    renderEventBar(byDate);
  }

  // ============== ③ 日历下方「项目名称」栏 ==============
  function renderEventBar(byDate) {
    const bar = hostContainer && hostContainer.querySelector('#cal-event-bar');
    if (!bar) return;
    const list = (byDate[selectedKey] || []);
    // 「添加当天项目」按钮：任何日期都可添加（按钮仅在修改模式下显示）
    const addBtn = '<button class="cal-bar-add" id="cal-add-btn" title="在 ' +
      escapeHtml(selectedKey) + ' 新增一条开标项目">＋ 添加当天项目</button>';
    if (!list.length) {
      bar.innerHTML = '<div class="cal-bar-empty">' +
        escapeHtml(selectedKey || '') + '：当天没有开标项目</div>' + addBtn;
    } else {
      const items = list.map((rec) => {
        const name = rec.fields.projectName || '未命名项目';
        return '<span class="cal-bar-item" data-id="' + rec.id + '" title="' +
          escapeHtml(name) + '（修改模式下双击可改名）">' +
          escapeHtml(name) + '</span>';
      }).join('<span class="cal-bar-sep">、</span>');
      bar.innerHTML = '<span class="cal-bar-date">' + escapeHtml(selectedKey) +
        '（' + list.length + ' 个项目）：</span>' + items + addBtn;
    }
    const addEl = bar.querySelector('#cal-add-btn');
    if (addEl) addEl.onclick = addEventForSelected;
  }

  // 在选中日期新增开标项目：点按钮只弹出输入框，输入名称回车才真正创建
  function addEventForSelected() {
    if (!window.Workbench || !Workbench.ensureEditable()) return;
    const bar = hostContainer && hostContainer.querySelector('#cal-event-bar');
    if (!bar || bar.querySelector('.cal-bar-add-input')) return; // 已有输入框，不重复弹

    const btn = bar.querySelector('#cal-add-btn');
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'cal-bar-add-input';
    input.placeholder = '输入项目名称，回车添加（Esc 取消）';
    if (btn) btn.replaceWith(input); else bar.appendChild(input);
    input.focus();

    let done = false;
    const finish = (keep) => {
      if (done) return;
      done = true;
      const name = input.value.trim();
      if (keep && name) {
        const arr = getRecords();
        arr.push({
          id: Date.now() + '_' + Math.random().toString(36).slice(2, 7),
          fields: { projectName: name, openTime: selectedKey + ' 09:00' },
          url: '',
          addedAt: Date.now(),
          projectType: 'government', // 日历手动添加 = 政府采购项目，归入清单第二张表
        });
        // 交给宿主保存并重渲清单（宿主 renderTable 内部会再调 refresh 刷新日历）
        if (api && typeof api.onSaved === 'function') api.onSaved();
        else renderDays();
      } else {
        renderDays(); // 取消/空输入：还原按钮
      }
    };

    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); finish(true); }
      else if (e.key === 'Escape') { e.preventDefault(); finish(false); }
    });
    input.addEventListener('blur', () => finish(true)); // 失焦：有内容则添加，空则取消
  }

  // 项目名编辑（受顶部「修改模式」保护）
  function editEventName(el, rec) {
    if (!window.Workbench || !Workbench.ensureEditable()) return;
    if (el.querySelector('input')) return;
    const original = rec.fields.projectName || '';

    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'ev-edit-input';
    input.value = original;
    el.innerHTML = '';
    el.appendChild(input);
    input.focus();
    input.select();

    let done = false;
    const finish = (keep) => {
      if (done) return;
      done = true;
      if (keep) {
        rec.fields.projectName = input.value.trim();
        // 交给宿主保存并重渲清单（宿主 renderTable 内部会再调 refresh 刷新日历）
        if (api && typeof api.onSaved === 'function') api.onSaved();
        else renderDays();
      } else {
        renderDays();
      }
    };

    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); finish(true); }
      else if (e.key === 'Escape') { e.preventDefault(); finish(false); }
    });
    input.addEventListener('blur', () => finish(true));
  }

  function bindBarEvents() {
    const bar = hostContainer && hostContainer.querySelector('#cal-event-bar');
    if (!bar) return;
    bar.querySelectorAll('.cal-bar-item').forEach((el) => {
      el.ondblclick = () => {
        const rec = getRecords().find((r) => r.id === el.dataset.id);
        if (rec) editEventName(el, rec);
      };
    });
  }

  // 覆盖 renderEventBar 后补绑事件：渲染完直接绑
  const _origRenderEventBar = renderEventBar;
  renderEventBar = function (byDate) {
    _origRenderEventBar(byDate);
    bindBarEvents();
  };

  function syncSelectors() {
    const yearSel = hostContainer && hostContainer.querySelector('#cal-year');
    const monthSel = hostContainer && hostContainer.querySelector('#cal-month');
    if (yearSel) yearSel.value = curYear;
    if (monthSel) monthSel.value = curMonth;
  }

  // ============== ④ 对外接口 ==============
  window.TenderCalendar = {
    /** 挂载日历到指定容器（由 tender-list 页面调用） */
    mount(container, hostApi) {
      hostContainer = container;
      api = hostApi || null;
      if (!hostContainer) return;
      // 打开页面始终回到「当前年月」，默认选中今天
      const now = new Date();
      curYear = now.getFullYear();
      curMonth = now.getMonth() + 1;
      selectedKey = formatDateKey(now);

      hostContainer.innerHTML =
        '<div class="calendar-toolbar">' +
        '<button class="btn ghost small" id="cal-prev" title="上一月">‹</button>' +
        '<span class="cal-title" id="cal-title"></span>' +
        '<button class="btn ghost small" id="cal-next" title="下一月">›</button>' +
        '<button class="btn ghost small" id="cal-today">今天</button>' +
        '</div>' +
        '<div class="calendar-grid">' +
        '<div class="cal-weekday">日</div><div class="cal-weekday">一</div>' +
        '<div class="cal-weekday">二</div><div class="cal-weekday">三</div>' +
        '<div class="cal-weekday">四</div><div class="cal-weekday">五</div>' +
        '<div class="cal-weekday">六</div>' +
        '<div id="cal-days" class="cal-days"></div>' +
        '</div>' +
        '<div class="cal-event-bar" id="cal-event-bar"></div>';

      const prevBtn = hostContainer.querySelector('#cal-prev');
      const nextBtn = hostContainer.querySelector('#cal-next');
      const todayBtn = hostContainer.querySelector('#cal-today');

      prevBtn.onclick = () => {
        curMonth--;
        if (curMonth < 1) { curMonth = 12; curYear--; }
        // 翻页后若选中日期不在本月，改选本月 1 号，避免选中格看不见
        if (!selectedKey.startsWith(curYear + '-' + pad2(curMonth))) selectedKey = curYear + '-' + pad2(curMonth) + '-01';
        syncSelectors();
        renderDays();
      };
      nextBtn.onclick = () => {
        curMonth++;
        if (curMonth > 12) { curMonth = 1; curYear++; }
        if (!selectedKey.startsWith(curYear + '-' + pad2(curMonth))) selectedKey = curYear + '-' + pad2(curMonth) + '-01';
        syncSelectors();
        renderDays();
      };
      todayBtn.onclick = () => {
        const now2 = new Date();
        curYear = now2.getFullYear(); curMonth = now2.getMonth() + 1;
        selectedKey = formatDateKey(now2);
        syncSelectors();
        renderDays();
      };

      renderDays();
    },

    /** 清单数据变化后刷新日历（保持当前查看的年月与选中日期） */
    refresh() {
      if (!hostContainer || !document.body.contains(hostContainer)) return;
      renderDays();
    },
  };

  // 标题（YYYY年MM月）随渲染更新
  const _origRenderDays = renderDays;
  renderDays = function () {
    _origRenderDays();
    const title = hostContainer && hostContainer.querySelector('#cal-title');
    if (title) title.textContent = curYear + '年' + curMonth + '月';
  };
})();
