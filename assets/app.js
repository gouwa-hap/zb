/**
 * 工作台插件注册机制
 * --------------------------------------------------
 * 插件是一个对象：
 *   {
 *     id:      'tender-list',          // 唯一标识
 *     name:    '统计招投标项目清单',     // 功能栏显示名
 *     icon:    '📋',                    // 功能栏图标
 *     description: '...',             // 副标题
 *     mount(container) { ... }        // 点击后渲染到主区
 *   }
 * 新增功能只需写一个这样的对象并调用 Workbench.register(plugin)。
 *
 * 数据保护（修改模式）：
 *   默认处于「只读模式」。所有会改动数据的操作（双击编辑、增加、删除、
 *   导入、清空、批量抓取、粘贴解析等）都必须先点击顶部「🔧 修改」按钮，
 *   通过口令验证进入「修改模式」后才能执行。再次点击该按钮可退出，
 *   退出后所有改动操作被重新锁定。
 */
(function () {
  // 进入修改模式所需的密码
  const EDIT_PASSWORD = '123456';

  const registry = {
    plugins: [],
    activeId: null,
    editing: false, // 修改模式开关（true=可改，false=只读）

    /** 注册一个插件 */
    register(plugin) {
      if (!plugin || !plugin.id) {
        console.warn('[Workbench] 插件缺少 id，已忽略', plugin);
        return;
      }
      if (this.plugins.some((p) => p.id === plugin.id)) {
        console.warn('[Workbench] 插件 id 重复，已忽略：', plugin.id);
        return;
      }
      this.plugins.push(plugin);
      this.renderSidebar();
    },

    /** 渲染左侧功能栏 */
    renderSidebar() {
      const list = document.getElementById('plugin-list');
      if (!list) return;
      list.innerHTML = '';
      this.plugins.forEach((p) => {
        const el = document.createElement('button');
        el.className =
          'plugin-item' + (p.id === this.activeId ? ' active' : '');
        el.innerHTML =
          '<span class="pi-icon">' + (p.icon || '📦') + '</span>' +
          '<span class="pi-name">' + p.name + '</span>';
        el.title = p.description || p.name;
        el.onclick = () => this.activate(p.id);
        list.appendChild(el);
      });
    },

    /** 激活（打开）某个插件 */
    activate(id) {
      this.activeId = id;
      this.renderSidebar();
      const p = this.plugins.find((x) => x.id === id);
      const main = document.getElementById('main');
      if (!main) return;
      main.innerHTML = '';
      if (p && typeof p.mount === 'function') {
        p.mount(main);
      } else {
        main.innerHTML = '<div class="welcome"><p>该功能暂不可用。</p></div>';
      }
    },

    /** 是否处于修改模式 */
    isEditing() {
      return !!this.editing;
    },

    /**
     * 进入修改模式：弹出自定义口令框验证，通过后才置 editing=true。
     * 不使用原生 prompt，且页面/弹窗均不出现「密码」字样或正确口令。
     */
    enterEditMode() {
      return new Promise((resolve) => {
        this.showPasswordModal((ok) => {
          if (ok) {
            this.setEditing(true);
            resolve(true);
          } else {
            resolve(false);
          }
        });
      });
    },

    /**
     * 自定义口令验证框（无「密码」字样，不泄露正确口令）。
     * @param {(ok:boolean)=>void} cb 验证结果回调
     */
    showPasswordModal(cb) {
      const overlay = document.createElement('div');
      overlay.className = 'pwd-overlay';
      overlay.innerHTML =
        '<div class="pwd-box">' +
        '<input type="password" class="pwd-input" placeholder="请输入口令" autocomplete="off" />' +
        '<div class="pwd-actions">' +
        '<button class="pwd-cancel" type="button">取消</button>' +
        '<button class="pwd-ok" type="button">确定</button>' +
        '</div>' +
        '<div class="pwd-err" style="display:none;">口令错误</div>' +
        '</div>';

      const input = overlay.querySelector('.pwd-input');
      const errEl = overlay.querySelector('.pwd-err');
      const close = () => {
        if (overlay && overlay.parentNode) overlay.parentNode.removeChild(overlay);
      };
      const submit = () => {
        const val = (input.value || '').trim();
        if (val === EDIT_PASSWORD) {
          close();
          cb(true);
        } else {
          errEl.style.display = 'block';
          input.value = '';
          input.focus();
        }
      };

      overlay.querySelector('.pwd-ok').onclick = submit;
      overlay.querySelector('.pwd-cancel').onclick = () => {
        close();
        cb(false);
      };
      overlay.onclick = (e) => {
        if (e.target === overlay) {
          close();
          cb(false);
        }
      };
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') { e.preventDefault(); submit(); }
        if (e.key === 'Escape') { close(); cb(false); }
      });

      document.body.appendChild(overlay);
      setTimeout(() => input.focus(), 30);
    },

    /** 切换修改模式（顶部「修改」按钮点击入口） */
    toggleEditMode() {
      if (this.editing) {
        this.setEditing(false);
      } else {
        this.enterEditMode();
      }
    },

    /** 设置修改模式状态并刷新顶部按钮与提示条 */
    setEditing(v) {
      this.editing = !!v;
      this.renderEditUI();
    },

    /** 刷新「修改」按钮与只读/可改提示条 */
    renderEditUI() {
      const btn = document.getElementById('btn-edit-mode');
      const banner = document.getElementById('edit-banner');
      if (btn) {
        if (this.editing) {
          btn.textContent = '✓ 修改中（点此退出）';
          btn.classList.add('active');
        } else {
          btn.textContent = '🔧 修改';
          btn.classList.remove('active');
        }
      }
      if (banner) {
        if (this.editing) {
          banner.textContent =
            '🔓 已进入修改模式：可双击编辑、增加、删除、导入等数据改动。再次点击「🔧 修改」按钮可退出。';
          banner.classList.add('show');
        } else {
          banner.textContent = '';
          banner.classList.remove('show');
        }
      }
      if (document.body) document.body.classList.toggle('edit-mode', this.editing);
      // 修改模式切换后刷新日历（其「添加当天项目」按钮依赖 edit-mode 显隐）
      if (window.TenderCalendar && typeof TenderCalendar.refresh === 'function') {
        TenderCalendar.refresh();
      }
    },

    /**
     * 数据改动操作守卫：在任何会改动数据的操作前调用。
     * 已进入修改模式 → 返回 true，直接放行；
     * 未进入 → 弹出提示并返回 false，调用方应 return 中止本次操作。
     */
    ensureEditable() {
      if (this.editing) return true;
      window.alert(
        '当前为只读模式，无法修改数据。\n请先点击顶部「🔧 修改」按钮进入修改模式。'
      );
      return false;
    },

    /* ============ 服务端数据持久化层（替代 localStorage） ============
     * 数据保存在服务端（server.js 的 server-data/store.json），与浏览器/设备无关，
     * 因此「更新功能 / 换浏览器 / 换设备」都不会丢数据。
     * 同时在本机浏览器 localStorage 留一份缓存，作为离线/服务端不可达时的兜底；
     * 首次上线时若服务端为空但本机有缓存，会自动把本机数据上传到服务端（迁移）。
     * 各插件通过 loadRecords(STORE_KEY) / saveRecords(STORE_KEY, arr) 调用。
     */
    _DATA_KEY_MAP: {
      'wb_tender_records': 'tender',
      'wb_yigong_records': 'yigong',
    },
    _dataSlug(storeKey) {
      return (this._DATA_KEY_MAP && this._DATA_KEY_MAP[storeKey]) || storeKey;
    },
    _dataLocalGet(storeKey) {
      try { const raw = window.localStorage.getItem(storeKey); return raw ? JSON.parse(raw) : []; }
      catch (e) { return []; }
    },
    _dataLocalSet(storeKey, arr) {
      try { window.localStorage.setItem(storeKey, JSON.stringify(arr)); } catch (e) {}
    },
    async loadRecords(storeKey) {
      const slug = this._dataSlug(storeKey);
      let serverArr = null;
      try {
        const r = await fetch('/api/data/' + slug);
        if (r.ok) {
          const j = await r.json();
          if (j && j.ok) serverArr = Array.isArray(j.records) ? j.records : [];
        }
      } catch (e) { serverArr = null; }
      const lsArr = this._dataLocalGet(storeKey);
      if (Array.isArray(serverArr)) {
        // 服务端可达
        if (serverArr.length) { this._dataLocalSet(storeKey, serverArr); return serverArr; }
        // 服务端为空 → 用本机缓存迁移上传（把现有数据上传到数据库）
        if (lsArr.length) { await this.saveRecords(storeKey, lsArr); return lsArr; }
        return [];
      }
      // 服务端不可达 → 离线模式，用本机缓存
      return lsArr;
    },
    async saveRecords(storeKey, arr) {
      // 先写本机缓存兜底
      this._dataLocalSet(storeKey, arr);
      const snapshot = arr;
      const run = async () => {
        try {
          const slug = this._dataSlug(storeKey);
          const r = await fetch('/api/data/' + slug, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ records: snapshot }),
          });
          return r.ok;
        } catch (e) { return false; }
      };
      // 串行化多次保存，避免并发 PUT 互相覆盖
      this._dataChain = (this._dataChain || Promise.resolve()).then(run, run);
      return this._dataChain;
    },

    /**
     * 让表格支持「列宽 + 行高自由拖动」，并强制内容居中。
     * 因各插件的 renderTable 每次都用 innerHTML 重建表格，故需：
     *   1) 每次重建后重新调用本方法（重绑手柄、恢复上次尺寸）；
     *   2) 列宽/行高持久化到 localStorage，避免重渲后丢失。
     * @param {HTMLTableElement} table 目标表格
     * @param {string} storageKey 持久化键前缀（两表不同，如 'tender'/'yigong'）
     */
    makeTableResizable(table, storageKey) {
      if (!table) return;
      storageKey = 'rz:' + storageKey;

      // 叶子表头行（多级表头的最后一行；单列头即唯一一行）
      const thead = table.querySelector('thead');
      let leafRow = null;
      if (thead) {
        const trs = thead.querySelectorAll('tr');
        leafRow = trs[trs.length - 1];
      }
      // 计算底层列数
      let colCount = 0;
      if (leafRow) colCount = leafRow.querySelectorAll('th').length;
      else {
        const fr = table.querySelector('tbody tr');
        if (fr) colCount = fr.children.length;
      }
      if (!colCount) return;

      // 列宽：优先用已存尺寸，否则测量当前（auto 布局）写入存储
      const storedCols = this._rzLoad(storageKey + ':cols');
      const useStored = storedCols && storedCols.length === colCount;
      let widths;
      if (useStored) {
        widths = storedCols;
      } else {
        widths = [];
        if (leafRow) {
          leafRow.querySelectorAll('th').forEach((th) =>
            widths.push(Math.round(th.getBoundingClientRect().width)));
        } else {
          const fr = table.querySelector('tbody tr');
          if (fr) Array.from(fr.children).forEach((c) =>
            widths.push(Math.round(c.getBoundingClientRect().width)));
        }
        this._rzSave(storageKey + ':cols', widths);
      }

      table.classList.add('resizable-table');
      // 建 colgroup 控制列宽（fixed 布局）
      let colgroup = table.querySelector('colgroup');
      if (!colgroup) {
        colgroup = document.createElement('colgroup');
        table.insertBefore(colgroup, table.firstChild);
      }
      colgroup.innerHTML = '';
      widths.forEach((w) => {
        const col = document.createElement('col');
        if (w) col.style.width = w + 'px';
        colgroup.appendChild(col);
      });

      // 列拖动手柄：叶子表头每个 th 右侧
      if (leafRow) {
        leafRow.querySelectorAll('th').forEach((th, i) => {
          if (th.querySelector('.col-resizer')) return;
          th.style.position = 'relative';
          const h = document.createElement('span');
          h.className = 'col-resizer';
          th.appendChild(h);
          h.addEventListener('mousedown', (e) =>
            this._rzColStart(e, colgroup, i, storageKey));
        });
      }

      // 行拖动手柄：tbody 每行底边（整行可拖，不依赖特定单元格）
      const rows = table.querySelectorAll('tbody tr');
      const storedRows = this._rzLoad(storageKey + ':rows');
      table._rzStorageKey = storageKey;
      table.onmousemove = (e) => {
        if (this._rzResizing) return;
        const tr = e.target.closest && e.target.closest('tbody tr');
        if (!tr) { table.style.cursor = ''; return; }
        const r = tr.getBoundingClientRect();
        const nearBottom = e.clientY >= r.bottom - 5 && e.clientY <= r.bottom + 4;
        table.style.cursor = nearBottom ? 'row-resize' : '';
      };
      table.onmousedown = (e) => {
        if (this._rzResizing) return;
        const tr = e.target.closest && e.target.closest('tbody tr');
        if (!tr) return;
        const r = tr.getBoundingClientRect();
        const nearBottom = e.clientY >= r.bottom - 5 && e.clientY <= r.bottom + 4;
        if (!nearBottom) return; // 非底边：放行单元格单击/双击
        e.preventDefault();
        const idx = Array.prototype.indexOf.call(tr.parentNode.children, tr);
        this._rzRowStart(e, tr, idx, storageKey, true);
      };
      rows.forEach((tr, i) => {
        const h0 = storedRows && storedRows[i];
        if (h0) tr.style.height = h0 + 'px';
      });
    },

    /** 列拖动：在文档级监听 mousemove/up，实时改 col 宽度并持久化 */
    _rzColStart(e, colgroup, idx, storageKey) {
      e.preventDefault();
      e.stopPropagation();
      const col = colgroup.children[idx];
      if (!col) return;
      this._rzResizing = true;
      const startX = e.clientX;
      const startW = col.getBoundingClientRect().width;
      const arr = this._rzLoad(storageKey + ':cols') || [];
      const move = (ev) => {
        const w = Math.max(36, Math.round(startW + (ev.clientX - startX)));
        col.style.width = w + 'px';
        arr[idx] = w;
      };
      const up = () => {
        document.removeEventListener('mousemove', move);
        document.removeEventListener('mouseup', up);
        this._rzSave(storageKey + ':cols', arr);
        this._rzResizing = false;
        document.body.style.cursor = '';
      };
      document.addEventListener('mousemove', move);
      document.addEventListener('mouseup', up);
      document.body.style.cursor = 'col-resize';
    },

    /** 行拖动：实时改 tr 高度并持久化；suppressClick 时抑制误触发的单元格 click */
    _rzRowStart(e, tr, idx, storageKey, suppressClick) {
      this._rzResizing = true;
      const startY = e.clientY;
      const startH = tr.getBoundingClientRect().height;
      const arr = this._rzLoad(storageKey + ':rows') || [];
      let clickGuard = null;
      if (suppressClick) {
        clickGuard = (ev) => { ev.stopPropagation(); ev.preventDefault(); };
        document.addEventListener('click', clickGuard, true);
      }
      const move = (ev) => {
        const h = Math.max(26, Math.round(startH + (ev.clientY - startY)));
        tr.style.height = h + 'px';
        arr[idx] = h;
      };
      const up = () => {
        document.removeEventListener('mousemove', move);
        document.removeEventListener('mouseup', up);
        if (clickGuard) document.removeEventListener('click', clickGuard, true);
        this._rzSave(storageKey + ':rows', arr);
        this._rzResizing = false;
        document.body.style.cursor = '';
      };
      document.addEventListener('mousemove', move);
      document.addEventListener('mouseup', up);
      document.body.style.cursor = 'row-resize';
    },

    _rzLoad(key) {
      try {
        const v = window.localStorage.getItem(key);
        return v ? JSON.parse(v) : null;
      } catch (e) { return null; }
    },
    _rzSave(key, val) {
      try { window.localStorage.setItem(key, JSON.stringify(val)); } catch (e) {}
    },
  };

  window.Workbench = registry;

  // 初始化「修改」按钮（本脚本置于 body 末尾，DOM 已就绪）
  const btn = document.getElementById('btn-edit-mode');
  if (btn) {
    btn.onclick = () => registry.toggleEditMode();
    registry.renderEditUI();
  }
})();
