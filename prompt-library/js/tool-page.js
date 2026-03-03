/* =============================================
   AI 房仲工具中心 — tool-page.js
   通用工具子頁邏輯
   由各工具 HTML 設定 window.TOOL_ID 後載入
   ============================================= */
;(function () {
  'use strict';

  var TC = window.ToolsCommon;
  var TOOL_ID = window.TOOL_ID;
  if (!TOOL_ID) return console.error('未設定 TOOL_ID');

  TC.checkPassword();

  var toolData = null;
  var currentScenario = null;
  var currentTabId = null; // 帶看準備包用

  // DOM
  var scenarioList = document.getElementById('scenario-list');
  var formContainer = document.getElementById('tool-form');
  var formSection = document.getElementById('form-section');
  var generateBtn = document.getElementById('generate-btn');
  var outputContent = document.getElementById('output-content');
  var outputActions = document.getElementById('output-actions');
  var copyBtn = document.getElementById('copy-btn');
  var copyAllBtn = document.getElementById('copy-all-btn');
  var tabsContainer = document.getElementById('output-tabs');
  var emptyState = document.getElementById('empty-state');

  async function init() {
    var data = await TC.loadTemplates();
    if (!data || !data.tools || !data.tools[TOOL_ID]) {
      console.error('找不到工具：' + TOOL_ID);
      return;
    }
    toolData = data.tools[TOOL_ID];
    renderScenarios();
  }

  function renderScenarios() {
    scenarioList.innerHTML = '';
    var scenarios = toolData.scenarios;
    Object.keys(scenarios).forEach(function (key) {
      var s = scenarios[key];
      var div = document.createElement('div');
      div.className = 'scenario-item';
      div.dataset.id = key;

      var emoji = s.emoji || toolData.icon;
      div.innerHTML =
        '<span class="scenario-emoji">' + emoji + '</span>' +
        '<div class="scenario-text">' +
          '<div class="scenario-name">' + s.label + '</div>' +
          '<div class="scenario-desc">' + (s.description || '') + '</div>' +
        '</div>';

      div.addEventListener('click', function () {
        selectScenario(key);
      });
      scenarioList.appendChild(div);
    });
  }

  function selectScenario(key) {
    currentScenario = key;
    currentTabId = null;

    // 更新選中狀態
    scenarioList.querySelectorAll('.scenario-item').forEach(function (el) {
      el.classList.toggle('active', el.dataset.id === key);
    });

    var scenario = toolData.scenarios[key];

    // 渲染表單
    TC.renderForm(scenario.fields, formContainer, onFormChange);
    formSection.classList.remove('hidden');
    generateBtn.classList.remove('hidden');

    // 清空輸出
    resetOutput();

    // 如果有 tabs（帶看準備包），預設第一個 tab
    if (scenario.tabs && scenario.tabs.length > 0) {
      renderTabs(scenario.tabs);
      currentTabId = scenario.tabs[0].id;
      updateTabActive();
    } else {
      if (tabsContainer) tabsContainer.classList.add('hidden');
    }
  }

  function renderTabs(tabs) {
    if (!tabsContainer) return;
    tabsContainer.innerHTML = '';
    tabsContainer.classList.remove('hidden');
    tabs.forEach(function (tab) {
      var btn = document.createElement('button');
      btn.className = 'tool-tab';
      btn.dataset.tabId = tab.id;
      btn.textContent = tab.icon + ' ' + tab.label;
      btn.addEventListener('click', function () {
        currentTabId = tab.id;
        updateTabActive();
        updatePreview();
      });
      tabsContainer.appendChild(btn);
    });
  }

  function updateTabActive() {
    if (!tabsContainer) return;
    tabsContainer.querySelectorAll('.tool-tab').forEach(function (btn) {
      btn.classList.toggle('active', btn.dataset.tabId === currentTabId);
    });
  }

  function resetOutput() {
    if (emptyState) emptyState.classList.remove('hidden');
    outputContent.innerHTML = '';
    outputContent.classList.add('hidden');
    outputActions.classList.add('hidden');
  }

  function onFormChange() {
    // 即時預覽（如果已經產出過）
    if (outputContent.classList.contains('hidden')) return;
    updatePreview();
  }

  function updatePreview() {
    var scenario = toolData.scenarios[currentScenario];
    var values = TC.collectValues(formContainer);

    if (scenario.tabs && currentTabId) {
      // 找到對應 tab 的 template
      var tab = scenario.tabs.find(function (t) { return t.id === currentTabId; });
      if (tab) {
        outputContent.innerHTML = TC.highlightTemplate(tab.template, values);
      }
    } else {
      outputContent.innerHTML = TC.highlightTemplate(scenario.template, values);
    }
  }

  // 產出按鈕
  if (generateBtn) {
    generateBtn.addEventListener('click', function () {
      if (!currentScenario) return;
      var scenario = toolData.scenarios[currentScenario];
      var values = TC.collectValues(formContainer);

      // 檢查必填
      var missing = (scenario.fields || []).filter(function (f) {
        return f.required && !values[f.key];
      });
      if (missing.length > 0) {
        alert('請填入必填欄位：' + missing.map(function (f) { return f.label; }).join('、'));
        return;
      }

      // 顯示預覽
      if (emptyState) emptyState.classList.add('hidden');
      outputContent.classList.remove('hidden');
      outputActions.classList.remove('hidden');
      updatePreview();

      // 儲存紀錄
      TC.saveToHistory(TOOL_ID, currentScenario, values);
    });
  }

  // 複製按鈕
  if (copyBtn) {
    copyBtn.addEventListener('click', function () {
      var scenario = toolData.scenarios[currentScenario];
      var values = TC.collectValues(formContainer);
      var text;

      if (scenario.tabs && currentTabId) {
        var tab = scenario.tabs.find(function (t) { return t.id === currentTabId; });
        text = tab ? TC.generatePlainText(tab.template, values) : '';
      } else {
        text = TC.generatePlainText(scenario.template, values);
      }
      TC.copyToClipboard(text, copyBtn);
    });
  }

  // 全部複製（帶看準備包用）
  if (copyAllBtn) {
    copyAllBtn.addEventListener('click', function () {
      var scenario = toolData.scenarios[currentScenario];
      if (!scenario.tabs) return;
      var values = TC.collectValues(formContainer);
      var allText = scenario.tabs.map(function (tab, i) {
        return '=== ' + tab.icon + ' ' + tab.label + ' ===\n\n' +
               TC.generatePlainText(tab.template, values);
      }).join('\n\n\n');
      TC.copyToClipboard(allText, copyAllBtn);
    });
  }

  init();
})();
