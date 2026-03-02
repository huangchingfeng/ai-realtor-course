/* =============================================
   Prompt 練習場 — playground.js
   ============================================= */
(() => {
  'use strict';

  const PASSWORD = 'realtor2026';
  const STORAGE_KEY = 'prompt_lib_auth';
  const PROFILES_KEY = 'playground_profiles';

  let promptsData = null;
  let varsData = null;
  let currentPromptId = null;
  let userValues = {};

  const $ = (sel, ctx = document) => ctx.querySelector(sel);
  const $$ = (sel, ctx = document) => [...ctx.querySelectorAll(sel)];

  document.addEventListener('DOMContentLoaded', init);

  async function init() {
    // 密碼保護
    if (localStorage.getItem(STORAGE_KEY) === 'true') {
      $('#pw-overlay').classList.add('hidden');
    }
    setupPassword();

    // 載入資料
    try {
      const [pRes, vRes] = await Promise.all([
        fetch('data/prompts.json'),
        fetch('data/variables.json')
      ]);
      if (!pRes.ok || !vRes.ok) throw new Error('載入失敗');
      promptsData = await pRes.json();
      varsData = await vRes.json();
    } catch (e) {
      console.error('載入失敗:', e);
      $('#preview-content').innerHTML = '<div class="pg-empty"><div class="pg-empty-icon">⚠️</div><h3>資料載入失敗</h3><p>請確認 data/prompts.json 和 data/variables.json 存在</p></div>';
      return;
    }

    renderPromptSelect();
    setupEventListeners();
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

  /* --- Prompt 選擇器 --- */
  function renderPromptSelect() {
    const select = $('#prompt-select');
    const modules = promptsData.modules;

    modules.forEach(m => {
      const group = document.createElement('optgroup');
      group.label = m.icon + ' ' + m.name;
      const modulePrompts = promptsData.prompts.filter(p => p.module === m.id);
      modulePrompts.forEach(p => {
        const opt = document.createElement('option');
        opt.value = p.id;
        opt.textContent = p.id + ' — ' + p.title;
        group.appendChild(opt);
      });
      select.appendChild(group);
    });
  }

  /* --- 事件綁定 --- */
  function setupEventListeners() {
    // 選擇 Prompt
    $('#prompt-select').addEventListener('change', e => {
      const id = e.target.value;
      if (!id) {
        resetAll();
        return;
      }
      selectPrompt(id);
    });

    // 複製
    $('#copy-btn').addEventListener('click', () => {
      const text = getReplacedText();
      copyToClipboard(text, $('#copy-btn'));
    });

    // 重置表單
    $('#reset-form-btn')?.addEventListener('click', () => {
      userValues = {};
      $$('#var-form input, #var-form select, #var-form textarea').forEach(el => {
        el.value = '';
        el.classList.remove('has-value');
      });
      renderPreview();
    });

    // 我的資料
    $('#profile-btn').addEventListener('click', () => {
      renderProfileList();
      $('#profile-modal').classList.remove('hidden');
    });
    $('#modal-close').addEventListener('click', () => {
      $('#profile-modal').classList.add('hidden');
    });
    $('#profile-modal').addEventListener('click', e => {
      if (e.target === $('#profile-modal')) $('#profile-modal').classList.add('hidden');
    });

    // 儲存資料組合
    $('#save-profile-btn')?.addEventListener('click', () => {
      $('#save-modal').classList.remove('hidden');
      $('#profile-name-input').value = '';
      $('#profile-name-input').focus();
    });
    $('#save-modal-close').addEventListener('click', () => {
      $('#save-modal').classList.add('hidden');
    });
    $('#save-modal').addEventListener('click', e => {
      if (e.target === $('#save-modal')) $('#save-modal').classList.add('hidden');
    });
    $('#confirm-save-btn').addEventListener('click', () => {
      const name = $('#profile-name-input').value.trim();
      if (!name) return;
      saveProfile(name);
      $('#save-modal').classList.add('hidden');
    });
    $('#profile-name-input').addEventListener('keydown', e => {
      if (e.key === 'Enter') $('#confirm-save-btn').click();
    });
  }

  /* --- 選擇 Prompt --- */
  function selectPrompt(id) {
    currentPromptId = id;
    userValues = {};

    const prompt = promptsData.prompts.find(p => p.id === id);
    if (!prompt) return;

    // 顯示 Prompt 資訊
    const meta = $('#prompt-meta');
    meta.classList.remove('hidden');
    const mod = promptsData.modules.find(m => m.id === prompt.module);
    meta.innerHTML = `
      <div class="pg-meta-title">${esc(prompt.title)}</div>
      <div class="pg-meta-tags">
        <span class="pg-meta-tag">${esc(mod?.icon || '')} ${esc(mod?.name || prompt.module)}</span>
        <span class="pg-meta-tag">${esc(prompt.level)}</span>
        <span class="pg-meta-tag tool">${esc(prompt.tool)}</span>
      </div>`;

    // 渲染表單
    renderForm(id);

    // 渲染預覽
    renderPreview();

    // 顯示動作按鈕
    $('#form-actions').classList.remove('hidden');
    $('#preview-actions').classList.remove('hidden');
  }

  /* --- 渲染變數表單 --- */
  function renderForm(promptId) {
    const formEl = $('#var-form');
    const promptVars = varsData.promptVariables[promptId];

    if (!promptVars || promptVars.length === 0) {
      formEl.innerHTML = '<div class="no-vars-msg">此 Prompt 尚未支援變數替換<br>請直接複製原始 Prompt 並手動修改</div>';
      return;
    }

    // 按 group 分類
    const grouped = {};
    for (const v of promptVars) {
      const fieldDef = findFieldDef(v.key);
      const groupKey = fieldDef ? fieldDef._group : 'other';
      if (!grouped[groupKey]) grouped[groupKey] = [];
      grouped[groupKey].push({ ...v, fieldDef });
    }

    let html = '';
    const groupOrder = ['property', 'finance', 'client', 'time', 'agent', 'content', 'holiday', 'other'];

    for (const gk of groupOrder) {
      if (!grouped[gk]) continue;
      const groupDef = varsData.variableGroups[gk];
      const icon = groupDef?.icon || '📝';
      const label = groupDef?.label || '其他';

      html += `<div class="form-group">
        <div class="form-group-title">${icon} ${esc(label)}</div>`;

      for (const v of grouped[gk]) {
        html += renderField(v);
      }
      html += '</div>';
    }

    formEl.innerHTML = html;

    // 綁定 input 事件
    $$('#var-form input, #var-form select, #var-form textarea').forEach(el => {
      el.addEventListener('input', () => {
        const key = el.dataset.key;
        userValues[key] = el.value.trim();
        el.classList.toggle('has-value', !!el.value.trim());
        renderPreview();
      });
    });
  }

  function renderField(v) {
    const fd = v.fieldDef;
    const label = v.label || fd?.label || v.key;
    const placeholder = fd?.placeholder || '輸入替換值...';
    const suffix = fd?.suffix || '';
    const type = fd?.type || 'text';

    let inputHtml = '';
    if (type === 'select' && fd?.options) {
      inputHtml = `<select data-key="${esc(v.key)}">
        <option value="">選擇...</option>
        ${fd.options.map(o => `<option value="${esc(o)}">${esc(o)}</option>`).join('')}
      </select>`;
    } else if (type === 'textarea') {
      inputHtml = `<textarea data-key="${esc(v.key)}" placeholder="${esc(placeholder)}" rows="2"></textarea>`;
    } else {
      inputHtml = `<input type="${type === 'number' ? 'text' : 'text'}" data-key="${esc(v.key)}" placeholder="${esc(placeholder)}">`;
    }

    const suffixHtml = suffix ? `<span class="field-suffix">${esc(suffix)}</span>` : '';
    const defaultHint = v.defaultInPrompt ? ` <span class="field-suffix">原文：${esc(truncate(v.defaultInPrompt, 20))}</span>` : '';

    return `<div class="form-field">
      <label>${esc(label)}${defaultHint}${suffixHtml}</label>
      ${inputHtml}
    </div>`;
  }

  function findFieldDef(key) {
    for (const [groupKey, group] of Object.entries(varsData.variableGroups)) {
      if (group.fields[key]) {
        return { ...group.fields[key], _group: groupKey };
      }
    }
    // 嘗試從 key 推斷 group
    if (key.startsWith('property_')) return { label: key, _group: 'property' };
    if (key.startsWith('client_')) return { label: key, _group: 'client' };
    if (key.startsWith('asking_') || key.startsWith('budget') || key.startsWith('monthly_') || key.startsWith('management_') || key.startsWith('down_') || key.startsWith('loan_') || key.startsWith('listing_') || key.startsWith('negotiation_')) return { label: key, _group: 'finance' };
    if (key.startsWith('agent_')) return { label: key, _group: 'agent' };
    return { label: key, _group: 'other' };
  }

  /* --- 即時預覽 --- */
  function renderPreview() {
    if (!currentPromptId) return;

    const prompt = promptsData.prompts.find(p => p.id === currentPromptId);
    if (!prompt) return;

    const promptVars = varsData.promptVariables[currentPromptId];
    const previewEl = $('#preview-content');

    if (!promptVars || promptVars.length === 0) {
      previewEl.innerHTML = esc(prompt.prompt).replace(/\n/g, '<br>');
      updateVarCount(0, 0);
      return;
    }

    // 按 defaultInPrompt 長度降序排序
    const sorted = [...promptVars].sort(
      (a, b) => (b.defaultInPrompt || '').length - (a.defaultInPrompt || '').length
    );

    let html = esc(prompt.prompt);

    // 用唯一佔位符避免替換衝突
    const placeholders = [];
    let idx = 0;

    for (const v of sorted) {
      if (!v.defaultInPrompt) continue;
      const search = esc(v.defaultInPrompt);
      const userVal = userValues[v.key];
      const placeholder = `\x00PH${idx}\x00`;

      if (userVal) {
        html = html.replaceAll(search, placeholder);
        placeholders.push({ placeholder, html: `<mark class="var-replaced">${esc(userVal)}</mark>` });
      } else {
        html = html.replaceAll(search, placeholder);
        placeholders.push({ placeholder, html: `<mark class="var-slot">${search}</mark>` });
      }
      idx++;
    }

    // 還原佔位符為 HTML
    for (const ph of placeholders) {
      html = html.replaceAll(ph.placeholder, ph.html);
    }

    html = html.replace(/\n/g, '<br>');
    previewEl.innerHTML = html;

    // 更新計數
    const filled = sorted.filter(v => userValues[v.key]).length;
    updateVarCount(filled, sorted.length);
  }

  function getReplacedText() {
    const prompt = promptsData.prompts.find(p => p.id === currentPromptId);
    if (!prompt) return '';

    const promptVars = varsData.promptVariables[currentPromptId];
    if (!promptVars) return prompt.prompt;

    const sorted = [...promptVars].sort(
      (a, b) => (b.defaultInPrompt || '').length - (a.defaultInPrompt || '').length
    );

    let text = prompt.prompt;
    for (const v of sorted) {
      if (!v.defaultInPrompt) continue;
      const userVal = userValues[v.key];
      if (userVal) {
        text = text.replaceAll(v.defaultInPrompt, userVal);
      }
    }
    return text;
  }

  function updateVarCount(filled, total) {
    const el = $('#var-count');
    if (!el) return;
    if (total === 0) {
      el.textContent = '';
    } else {
      el.textContent = `已填 ${filled}/${total} 個變數`;
    }
  }

  /* --- 我的資料 --- */
  function getProfiles() {
    try {
      const raw = localStorage.getItem(PROFILES_KEY);
      return raw ? JSON.parse(raw) : [];
    } catch { return []; }
  }

  function setProfiles(profiles) {
    localStorage.setItem(PROFILES_KEY, JSON.stringify(profiles));
  }

  function saveProfile(name) {
    const profiles = getProfiles();
    profiles.push({
      id: Date.now().toString(36),
      name,
      createdAt: new Date().toLocaleDateString('zh-TW'),
      data: { ...userValues }
    });
    setProfiles(profiles);
  }

  function loadProfile(id) {
    const profiles = getProfiles();
    const profile = profiles.find(p => p.id === id);
    if (!profile) return;

    // 將 profile 資料填入表單
    for (const [key, val] of Object.entries(profile.data)) {
      userValues[key] = val;
      const el = $(`[data-key="${key}"]`, $('#var-form'));
      if (el) {
        el.value = val;
        el.classList.toggle('has-value', !!val);
      }
    }
    renderPreview();
    $('#profile-modal').classList.add('hidden');
  }

  function deleteProfile(id) {
    const profiles = getProfiles().filter(p => p.id !== id);
    setProfiles(profiles);
    renderProfileList();
  }

  function renderProfileList() {
    const profiles = getProfiles();
    const listEl = $('#profile-list');
    const emptyEl = $('#profile-empty');

    if (profiles.length === 0) {
      listEl.innerHTML = '';
      emptyEl.classList.remove('hidden');
      return;
    }

    emptyEl.classList.add('hidden');
    listEl.innerHTML = profiles.map(p => `
      <div class="profile-item" data-id="${esc(p.id)}">
        <div>
          <div class="profile-item-name">${esc(p.name)}</div>
          <div class="profile-item-date">${esc(p.createdAt)}</div>
        </div>
        <div class="profile-item-actions">
          <button onclick="event.stopPropagation()" class="profile-load" data-id="${esc(p.id)}" title="載入">📥</button>
          <button onclick="event.stopPropagation()" class="profile-delete" data-id="${esc(p.id)}" title="刪除">🗑️</button>
        </div>
      </div>`).join('');

    // 綁定事件
    $$('.profile-load', listEl).forEach(btn => {
      btn.addEventListener('click', () => loadProfile(btn.dataset.id));
    });
    $$('.profile-delete', listEl).forEach(btn => {
      btn.addEventListener('click', () => {
        if (confirm('確定要刪除這組資料？')) deleteProfile(btn.dataset.id);
      });
    });
    $$('.profile-item', listEl).forEach(item => {
      item.addEventListener('click', () => loadProfile(item.dataset.id));
    });
  }

  /* --- 重置 --- */
  function resetAll() {
    currentPromptId = null;
    userValues = {};
    $('#prompt-meta').classList.add('hidden');
    $('#var-form').innerHTML = '';
    $('#form-actions').classList.add('hidden');
    $('#preview-actions').classList.add('hidden');
    $('#preview-content').innerHTML = `
      <div class="pg-empty">
        <div class="pg-empty-icon">🎯</div>
        <h3>選擇一個 Prompt 開始練習</h3>
        <p>從左側選擇模板，填入你的物件與客戶資料<br>右側會即時預覽客製化的 Prompt</p>
      </div>`;
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
    btn.textContent = '✅ 已複製！';
    btn.classList.add('copied');
    setTimeout(() => {
      btn.textContent = orig;
      btn.classList.remove('copied');
    }, 2000);
  }

  /* --- 工具函式 --- */
  function esc(str) {
    if (!str) return '';
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  function truncate(str, len) {
    if (!str) return '';
    return str.length > len ? str.slice(0, len) + '...' : str;
  }
})();
