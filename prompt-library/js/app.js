/* =============================================
   AI 超級房仲 Prompt 指令庫 — app.js
   ============================================= */
(() => {
  'use strict';

  const PASSWORD = 'realtor2026';
  const STORAGE_KEY = 'prompt_lib_auth';
  const DEBOUNCE_MS = 300;

  const LEVEL_MAP = {
    '新手': { emoji: '🟢', css: 'tag-green' },
    '中級': { emoji: '🟡', css: 'tag-yellow' },
    '進階': { emoji: '🔴', css: 'tag-red' }
  };

  const TOOL_LABELS = {
    chatgpt: 'ChatGPT',
    gemini: 'Gemini',
    perplexity: 'Perplexity',
    gamma: 'Gamma',
    canva: 'Canva',
    notebooklm: 'NotebookLM'
  };

  let data = null;
  let filteredPrompts = [];
  let activeModule = 'all';
  let searchQuery = '';
  let activeLevels = new Set(['新手', '中級', '進階']);
  let activeTools = new Set(Object.keys(TOOL_LABELS));

  const $ = (sel, ctx = document) => ctx.querySelector(sel);
  const $$ = (sel, ctx = document) => [...ctx.querySelectorAll(sel)];

  document.addEventListener('DOMContentLoaded', init);

  async function init() {
    // 密碼
    if (localStorage.getItem(STORAGE_KEY) === 'true') {
      $('#pw-overlay').classList.add('hidden');
    }
    setupPassword();

    // 載入資料
    try {
      const res = await fetch('data/prompts.json');
      if (!res.ok) throw new Error(res.status);
      data = await res.json();
    } catch (e) {
      console.error('無法載入 prompts.json:', e);
      $('#card-list').innerHTML = '<div class="empty-state"><div class="empty-icon">⚠️</div><h3>資料載入失敗</h3><p>請確認 data/prompts.json 存在</p></div>';
      return;
    }

    $('#total-badge').textContent = data.meta.total + ' 個案例';
    renderSidebar();
    renderMobileTabs();
    renderMobileFilters();
    applyFilters();
  }

  /* --- 密碼 --- */
  function setupPassword() {
    const submit = () => {
      const val = $('#pw-input').value.trim();
      if (val === PASSWORD) {
        localStorage.setItem(STORAGE_KEY, 'true');
        $('#pw-overlay').classList.add('hidden');
      } else {
        $('#pw-error').textContent = '密碼錯誤，請重新輸入';
        $('#pw-input').value = '';
        $('#pw-input').focus();
      }
    };
    $('#pw-submit').addEventListener('click', submit);
    $('#pw-input').addEventListener('keydown', e => {
      if (e.key === 'Enter') submit();
    });
  }

  /* --- Sidebar 渲染 --- */
  function renderSidebar() {
    // 模組 tab
    const tabsEl = $('#module-tabs');
    const allTab = createModuleTab('all', '📋', '全部', data.prompts.length);
    tabsEl.appendChild(allTab);
    data.modules.forEach(m => {
      const count = data.prompts.filter(p => p.module === m.id).length;
      tabsEl.appendChild(createModuleTab(m.id, m.icon, m.name, count));
    });

    // 難度
    const levelEl = $('#level-filters');
    Object.entries(LEVEL_MAP).forEach(([name, info]) => {
      levelEl.appendChild(createCheckbox(name, info.emoji + ' ' + name, 'level', true));
    });

    // 工具
    const toolEl = $('#tool-filters');
    Object.entries(TOOL_LABELS).forEach(([key, label]) => {
      const count = data.prompts.filter(p => p.toolCategories.includes(key)).length;
      if (count > 0) {
        toolEl.appendChild(createCheckbox(key, label + ' (' + count + ')', 'tool', true));
      }
    });

    // 搜尋
    let timer;
    $('#desktop-search').addEventListener('input', e => {
      clearTimeout(timer);
      timer = setTimeout(() => {
        searchQuery = e.target.value.trim().toLowerCase();
        // 同步手機
        if ($('#mobile-search')) $('#mobile-search').value = e.target.value;
        applyFilters();
      }, DEBOUNCE_MS);
    });
  }

  function createModuleTab(id, icon, name, count) {
    const btn = document.createElement('button');
    btn.className = 'module-tab' + (id === activeModule ? ' active' : '');
    btn.innerHTML = `<span class="tab-icon">${esc(icon)}</span>
      <span>${esc(name)}</span>
      <span class="tab-count">${count}</span>`;
    btn.addEventListener('click', () => {
      activeModule = id;
      $$('.module-tab').forEach(t => t.classList.remove('active'));
      btn.classList.add('active');
      // 同步手機 tab
      $$('.mobile-tab').forEach(t => {
        t.classList.toggle('active', t.dataset.id === id);
      });
      applyFilters();
    });
    return btn;
  }

  function createCheckbox(value, label, type, checked) {
    const el = document.createElement('label');
    el.className = 'checkbox-item';
    el.innerHTML = `<input type="checkbox" value="${esc(value)}" ${checked ? 'checked' : ''}><span>${esc(label)}</span>`;
    el.querySelector('input').addEventListener('change', e => {
      const set = type === 'level' ? activeLevels : activeTools;
      if (e.target.checked) set.add(value); else set.delete(value);
      applyFilters();
    });
    return el;
  }

  /* --- 手機版 Tab --- */
  function renderMobileTabs() {
    const container = $('#mobile-tabs');
    if (!container) return;

    const allBtn = document.createElement('button');
    allBtn.className = 'mobile-tab active';
    allBtn.textContent = '全部';
    allBtn.dataset.id = 'all';
    allBtn.addEventListener('click', () => selectMobileTab('all'));
    container.appendChild(allBtn);

    data.modules.forEach(m => {
      const btn = document.createElement('button');
      btn.className = 'mobile-tab';
      btn.textContent = m.icon + ' ' + m.name;
      btn.dataset.id = m.id;
      btn.addEventListener('click', () => selectMobileTab(m.id));
      container.appendChild(btn);
    });

    // 搜尋
    let timer;
    $('#mobile-search').addEventListener('input', e => {
      clearTimeout(timer);
      timer = setTimeout(() => {
        searchQuery = e.target.value.trim().toLowerCase();
        if ($('#desktop-search')) $('#desktop-search').value = e.target.value;
        applyFilters();
      }, DEBOUNCE_MS);
    });

    // 篩選按鈕
    $('#mobile-filter-btn').addEventListener('click', () => {
      const panel = $('#mobile-filter-panel');
      panel.classList.toggle('show');
      $('#mobile-filter-btn').classList.toggle('active');
    });
  }

  function selectMobileTab(id) {
    activeModule = id;
    $$('.mobile-tab').forEach(t => t.classList.toggle('active', t.dataset.id === id));
    $$('.module-tab').forEach(t => t.classList.toggle('active',
      (t.querySelector('.tab-icon')?.nextElementSibling?.textContent === '全部' && id === 'all') ||
      false
    ));
    // 同步桌面
    $$('.module-tab').forEach(t => t.classList.remove('active'));
    const desktopTabs = $$('.module-tab');
    if (id === 'all' && desktopTabs[0]) desktopTabs[0].classList.add('active');
    else {
      const mod = data.modules.find(m => m.id === id);
      if (mod) {
        desktopTabs.forEach(t => {
          if (t.textContent.includes(mod.name)) t.classList.add('active');
        });
      }
    }
    applyFilters();
  }

  /* --- 手機版篩選 chip --- */
  function renderMobileFilters() {
    const levelContainer = $('#mobile-level-chips');
    const toolContainer = $('#mobile-tool-chips');
    if (!levelContainer || !toolContainer) return;

    Object.entries(LEVEL_MAP).forEach(([name, info]) => {
      const chip = document.createElement('span');
      chip.className = 'chip active';
      chip.textContent = info.emoji + ' ' + name;
      chip.dataset.value = name;
      chip.addEventListener('click', () => {
        chip.classList.toggle('active');
        if (chip.classList.contains('active')) activeLevels.add(name);
        else activeLevels.delete(name);
        // 同步桌面
        $$('#level-filters input').forEach(inp => {
          if (inp.value === name) inp.checked = activeLevels.has(name);
        });
        applyFilters();
      });
      levelContainer.appendChild(chip);
    });

    Object.entries(TOOL_LABELS).forEach(([key, label]) => {
      const count = data.prompts.filter(p => p.toolCategories.includes(key)).length;
      if (count === 0) return;
      const chip = document.createElement('span');
      chip.className = 'chip active';
      chip.textContent = label;
      chip.dataset.value = key;
      chip.addEventListener('click', () => {
        chip.classList.toggle('active');
        if (chip.classList.contains('active')) activeTools.add(key);
        else activeTools.delete(key);
        $$('#tool-filters input').forEach(inp => {
          if (inp.value === key) inp.checked = activeTools.has(key);
        });
        applyFilters();
      });
      toolContainer.appendChild(chip);
    });
  }

  /* --- 篩選邏輯 --- */
  function applyFilters() {
    filteredPrompts = data.prompts.filter(p => {
      // 模組
      if (activeModule !== 'all' && p.module !== activeModule) return false;
      // 難度
      if (!activeLevels.has(p.level)) return false;
      // 工具
      if (!p.toolCategories.some(tc => activeTools.has(tc))) return false;
      // 搜尋
      if (searchQuery) {
        const hay = (p.title + p.prompt + p.scenario + p.task).toLowerCase();
        if (!hay.includes(searchQuery)) return false;
      }
      return true;
    });

    $('#filter-badge').textContent = '顯示 ' + filteredPrompts.length + ' 個';
    renderCards();
  }

  /* --- 卡片渲染 --- */
  function renderCards() {
    const list = $('#card-list');
    const empty = $('#empty-state');

    if (filteredPrompts.length === 0) {
      list.innerHTML = '';
      empty.classList.remove('hidden');
      return;
    }
    empty.classList.add('hidden');

    const frag = document.createDocumentFragment();
    filteredPrompts.forEach(p => frag.appendChild(createCard(p)));
    list.innerHTML = '';
    list.appendChild(frag);
  }

  function createCard(p) {
    const card = document.createElement('div');
    card.className = 'prompt-card';

    const levelInfo = LEVEL_MAP[p.level] || { emoji: '⚪', css: 'tag-green' };
    const moduleMeta = data.modules.find(m => m.id === p.module);

    let html = `
      <div class="card-header">
        <div class="card-title-area">
          <div class="card-id">${esc(p.id)} — ${esc(moduleMeta?.name || p.module)}</div>
          <div class="card-title">${esc(p.title)}</div>
        </div>
        <div class="card-tags">
          <span class="tag ${levelInfo.css}">${levelInfo.emoji} ${esc(p.level)}</span>
          <span class="tag-tool">${esc(p.tool)}</span>
        </div>
      </div>`;

    // 情境
    html += `
      <div class="card-scenario">
        <div class="scenario-toggle" onclick="this.nextElementSibling.classList.toggle('show'); this.querySelector('.scenario-arrow').classList.toggle('open')">
          <span class="scenario-label">📋 情境 & 任務</span>
          <span class="scenario-arrow">▼</span>
        </div>
        <div class="scenario-body">
          <div class="scenario-item"><strong>情境：</strong>${esc(p.scenario)}</div>
          <div class="scenario-item"><strong>任務：</strong>${esc(p.task)}</div>
        </div>
      </div>`;

    // 上傳資訊
    if (p.uploadFormat) {
      html += `<div class="upload-info"><strong>上傳：</strong>${esc(p.uploadFormat)}`;
      if (p.uploadFile) html += ` — ${esc(p.uploadFile)}`;
      if (p.uploadNote) html += `<br><em>${esc(p.uploadNote)}</em>`;
      html += '</div>';
    }

    // 系統串接
    if (p.systemNote) {
      html += `<div class="system-note">🔗 系統串接：${esc(p.systemNote)}</div>`;
    }

    // Prompt
    const promptText = p.prompt;
    const isLong = promptText.split('\n').length > 8 || promptText.length > 400;
    html += `
      <div class="card-prompt">
        <div class="prompt-label">Prompt 指令</div>
        <div class="prompt-code${isLong ? ' has-more' : ''}" data-prompt-id="${esc(p.id)}">${esc(promptText)}</div>
        <div class="prompt-actions">
          <button class="copy-btn" data-prompt-id="${esc(p.id)}">📋 複製指令</button>
          ${isLong ? '<button class="expand-btn" data-prompt-id="' + esc(p.id) + '">展開全部</button>' : ''}
        </div>
      </div>`;

    card.innerHTML = html;

    // 複製
    card.querySelector('.copy-btn').addEventListener('click', function() {
      copyToClipboard(promptText, this);
    });

    // 展開
    const expandBtn = card.querySelector('.expand-btn');
    if (expandBtn) {
      expandBtn.addEventListener('click', function() {
        const code = card.querySelector('.prompt-code');
        code.classList.toggle('expanded');
        this.textContent = code.classList.contains('expanded') ? '收合' : '展開全部';
      });
    }

    return card;
  }

  /* --- 複製 --- */
  async function copyToClipboard(text, btn) {
    try {
      await navigator.clipboard.writeText(text);
    } catch {
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.style.cssText = 'position:fixed;left:-9999px';
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      document.body.removeChild(ta);
    }
    const orig = btn.textContent;
    btn.textContent = '✅ 已複製';
    btn.classList.add('copied');
    setTimeout(() => {
      btn.textContent = orig;
      btn.classList.remove('copied');
    }, 2000);
  }

  /* --- XSS 防護 --- */
  function esc(str) {
    if (!str) return '';
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }
})();
