/* =============================================
   AI 房仲工具中心 — tools-common.js
   共用邏輯：密碼、模板載入、表單、複製
   ============================================= */
;(function () {
  'use strict';

  const PASS = 'realtor2026';
  const AUTH_KEY = 'prompt_lib_auth';

  /* --- 密碼驗證 --- */
  function checkPassword() {
    const overlay = document.getElementById('pw-overlay');
    if (!overlay) return;
    if (localStorage.getItem(AUTH_KEY) === PASS) {
      overlay.classList.add('hidden');
      return;
    }
    overlay.classList.remove('hidden');
    const input = document.getElementById('pw-input');
    const btn = document.getElementById('pw-submit');
    const err = document.getElementById('pw-error');

    function tryLogin() {
      if (input.value === PASS) {
        localStorage.setItem(AUTH_KEY, PASS);
        overlay.classList.add('hidden');
      } else {
        err.textContent = '密碼錯誤，請重新輸入';
        input.value = '';
        input.focus();
      }
    }
    btn.addEventListener('click', tryLogin);
    input.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') tryLogin();
    });
  }

  /* --- 載入模板 --- */
  async function loadTemplates() {
    // 工具子頁在 tools/ 目錄下，需要回上一層
    const paths = ['data/tool-templates.json', '../data/tool-templates.json'];
    for (const p of paths) {
      try {
        const res = await fetch(p);
        if (res.ok) return await res.json();
      } catch (_) { /* try next */ }
    }
    console.error('無法載入 tool-templates.json');
    return null;
  }

  /* --- 動態渲染表單 --- */
  function renderForm(fields, container, onChange) {
    container.innerHTML = '';
    fields.forEach(function (f) {
      const div = document.createElement('div');
      div.className = 'tool-field';

      const label = document.createElement('label');
      label.textContent = f.label;
      if (f.required) {
        const req = document.createElement('span');
        req.className = 'required';
        req.textContent = ' *';
        label.appendChild(req);
      }
      div.appendChild(label);

      let el;
      if (f.type === 'select') {
        el = document.createElement('select');
        (f.options || []).forEach(function (opt) {
          const o = document.createElement('option');
          o.value = opt;
          o.textContent = opt;
          if (opt === f.default) o.selected = true;
          el.appendChild(o);
        });
      } else if (f.type === 'textarea') {
        el = document.createElement('textarea');
        el.placeholder = f.placeholder || '';
        el.rows = 3;
      } else {
        el = document.createElement('input');
        el.type = 'text';
        el.placeholder = f.placeholder || '';
      }
      el.dataset.key = f.key;
      el.addEventListener('input', function () {
        if (el.value.trim()) {
          el.classList.add('has-value');
        } else {
          el.classList.remove('has-value');
        }
        if (typeof onChange === 'function') onChange();
      });
      el.addEventListener('change', function () {
        if (typeof onChange === 'function') onChange();
      });
      div.appendChild(el);
      container.appendChild(div);
    });
  }

  /* --- 收集表單值 --- */
  function collectValues(container) {
    const vals = {};
    container.querySelectorAll('[data-key]').forEach(function (el) {
      vals[el.dataset.key] = el.value.trim();
    });
    return vals;
  }

  /* --- 填入模板 --- */
  function fillTemplate(template, values) {
    let result = template;
    for (const key in values) {
      if (values[key]) {
        const re = new RegExp('\\{\\{' + key + '\\}\\}', 'g');
        result = result.replace(re, values[key]);
      }
    }
    return result;
  }

  /* --- 高亮渲染：已填 vs 未填 --- */
  function highlightTemplate(template, values) {
    let result = escapeHtml(template);
    // 先處理已填的
    for (const key in values) {
      if (values[key]) {
        const re = new RegExp('\\{\\{' + key + '\\}\\}', 'g');
        result = result.replace(re, '<span class="var-highlight">' + escapeHtml(values[key]) + '</span>');
      }
    }
    // 再處理未填的
    result = result.replace(/\{\{(\w+)\}\}/g, '<span class="var-unfilled">{{$1}}</span>');
    return result;
  }

  function escapeHtml(s) {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  /* --- 產出純文字（替換已填，保留未填 placeholder） --- */
  function generatePlainText(template, values) {
    let result = template;
    for (const key in values) {
      if (values[key]) {
        const re = new RegExp('\\{\\{' + key + '\\}\\}', 'g');
        result = result.replace(re, values[key]);
      }
    }
    // 把未填的 {{xxx}} 替換成更友善的 [xxx]
    result = result.replace(/\{\{(\w+)\}\}/g, '[$1]');
    return result;
  }

  /* --- 複製到剪貼簿 --- */
  async function copyToClipboard(text, btn) {
    try {
      await navigator.clipboard.writeText(text);
      const orig = btn.textContent;
      btn.textContent = '✅ 已複製！';
      btn.classList.add('copied');
      setTimeout(function () {
        btn.textContent = orig;
        btn.classList.remove('copied');
      }, 2000);
    } catch (e) {
      // fallback
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.style.position = 'fixed';
      ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      document.body.removeChild(ta);
      const orig = btn.textContent;
      btn.textContent = '✅ 已複製！';
      btn.classList.add('copied');
      setTimeout(function () {
        btn.textContent = orig;
        btn.classList.remove('copied');
      }, 2000);
    }
  }

  /* --- 使用紀錄 --- */
  function saveToHistory(toolId, scenarioId, values) {
    const key = 'tool_history_' + toolId;
    let history = [];
    try { history = JSON.parse(localStorage.getItem(key)) || []; } catch (_) {}
    history.unshift({
      scenario: scenarioId,
      values: values,
      timestamp: Date.now()
    });
    // 只保留最近 20 筆
    if (history.length > 20) history = history.slice(0, 20);
    localStorage.setItem(key, JSON.stringify(history));
  }

  /* --- 匯出 --- */
  window.ToolsCommon = {
    checkPassword: checkPassword,
    loadTemplates: loadTemplates,
    renderForm: renderForm,
    collectValues: collectValues,
    fillTemplate: fillTemplate,
    highlightTemplate: highlightTemplate,
    generatePlainText: generatePlainText,
    copyToClipboard: copyToClipboard,
    saveToHistory: saveToHistory
  };
})();
