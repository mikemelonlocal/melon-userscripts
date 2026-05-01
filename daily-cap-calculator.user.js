// ==UserScript==
// @name         Daily Cap Calculator - Melon Local (Enhanced)
// @namespace    https://thepatch.melonlocal.com/
// @version      3.7.0
// @description  Paces budgets evenly through end of month. Auto-fills from page data. Refresh + Freeze. Enhanced with auto-save, export/import, keyboard shortcuts, and improved UX.
// @author       Melon Local
// @match        https://thepatch.melonlocal.com/*
// @grant        GM_getValue
// @grant        GM_setValue
// @run-at       document-idle
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/daily-cap-calculator.user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/daily-cap-calculator.user.js
// ==/UserScript==

(function () {
  'use strict';

  // ============================================================================
  // CONFIGURATION
  // ============================================================================

  const CONFIG = {
    STORAGE_KEY_MINIMIZED: 'dcc-minimized',
    STORAGE_KEY_STATE: 'dcc-state-v3',
    STORAGE_KEY_FROZEN: 'dcc-frozen',
    STORAGE_KEY_DOCKED: 'dcc-docked',
    INIT_FLAG: 'dcc-initialized',
    AUTO_SAVE_DELAY: 1000,
    INIT_TIMEOUT: 5000,
    DEBUG: GM_getValue('dcc_debug', false),
    // Calculator only loads on these page types:
    //   /Agents/Dashboard/{id}                     (with #advertising or #melonmax hash)
    //   /Agents/BudgetDetails?budgetId=...
    //   /MelonMax/MelonMaxBudgetDetails?melonMaxBudgetId=...
    ALLOWED_URL_PATTERNS: [
      /\/Agents\/Dashboard\/\d+/i,
      /\/Agents\/BudgetDetails\b/i,
      /\/MelonMax\/MelonMaxBudgetDetails\b/i
    ]
  };

  // ============================================================================
  // URL VALIDATION
  // ============================================================================

  const DASHBOARD_RE = /\/Agents\/Dashboard\/\d+/i;
  const BUDGET_DETAILS_RE = /\/Agents\/BudgetDetails\b/i;
  const MELONMAX_BUDGET_DETAILS_RE = /\/MelonMax\/MelonMaxBudgetDetails\b/i;

  function isAllowedPage() {
    return CONFIG.ALLOWED_URL_PATTERNS.some(pattern => pattern.test(window.location.href));
  }

  function isDashboardPage() {
    return DASHBOARD_RE.test(window.location.href);
  }

  function isBudgetDetailsPage() {
    const url = window.location.href;
    return BUDGET_DETAILS_RE.test(url) || MELONMAX_BUDGET_DETAILS_RE.test(url);
  }

  function shouldShowCalculator() {
    const url = window.location.href;
    const hash = window.location.hash.toLowerCase();

    // BudgetDetails and MelonMaxBudgetDetails always show
    if (BUDGET_DETAILS_RE.test(url) || MELONMAX_BUDGET_DETAILS_RE.test(url)) {
      return true;
    }

    // Dashboard pages only show with #advertising or #melonmax hash
    if (DASHBOARD_RE.test(url)) {
      return hash.includes('advertising') || hash.includes('melonmax');
    }

    return false;
  }

  // ============================================================================
  // UTILITIES
  // ============================================================================

  const Utils = {
    log(...args) {
      if (CONFIG.DEBUG) console.log('[DCC]', ...args);
    },

    parseMoney(str) {
      return parseFloat((str || '').replace(/[^0-9.]/g, ''));
    },

    escapeHtml(str) {
      return String(str == null ? '' : str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
    },

    padZero(num) {
      return String(num).padStart(2, '0');
    },

    formatDate(date) {
      return `${date.getFullYear()}-${this.padZero(date.getMonth() + 1)}-${this.padZero(date.getDate())}`;
    },

    debounce(func, wait) {
      let timeout;
      return function executedFunction(...args) {
        const later = () => {
          clearTimeout(timeout);
          func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
      };
    },

    showToast(message, type = 'success') {
      const toast = document.createElement('div');
      toast.className = `dcc-toast dcc-toast-${type}`;
      toast.textContent = message;
      toast.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 12px 20px;
        border-radius: 8px;
        font-size: 13px;
        font-weight: 600;
        z-index: 999999;
        animation: dccSlideIn 0.3s ease;
        background: ${type === 'success' ? '#dcfce7' : type === 'error' ? '#fef2f2' : '#fef3c7'};
        border: 1.5px solid ${type === 'success' ? '#86efac' : type === 'error' ? '#fecaca' : '#fde68a'};
        color: ${type === 'success' ? '#166534' : type === 'error' ? '#dc2626' : '#92400e'};
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
      `;

      document.body.appendChild(toast);
      setTimeout(() => {
        toast.style.animation = 'dccSlideOut 0.3s ease';
        setTimeout(() => toast.remove(), 300);
      }, 3000);
    }
  };

  // ============================================================================
  // STYLES
  // ============================================================================

  const CSS = `
    /* Reset & Base */
    #daily-cap-calculator * {
      box-sizing: border-box;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }

    #daily-cap-calculator input[type=number]::-webkit-inner-spin-button,
    #daily-cap-calculator input[type=number]::-webkit-outer-spin-button {
      -webkit-appearance: none;
      margin: 0;
    }

    #daily-cap-calculator input[type=number] {
      -moz-appearance: textfield;
    }

    /* Toast Animations */
    @keyframes dccSlideIn {
      from {
        transform: translateX(400px);
        opacity: 0;
      }
      to {
        transform: translateX(0);
        opacity: 1;
      }
    }

    @keyframes dccSlideOut {
      from {
        transform: translateX(0);
        opacity: 1;
      }
      to {
        transform: translateX(400px);
        opacity: 0;
      }
    }

    /* Product Blocks */
    .dcc-product-block {
      background: #f8faf9;
      border: 1.5px solid #d1fae5;
      border-radius: 9px;
      padding: 13px 14px;
      margin-bottom: 10px;
      transition: border-color 0.2s;
    }

    .dcc-product-block:hover {
      border-color: #86efac;
    }

    .dcc-product-header {
      display: flex;
      align-items: center;
      gap: 8px;
      margin-bottom: 10px;
    }

    .dcc-product-name-input {
      flex: 1;
      font-size: 13px;
      font-weight: 700;
      color: #14532d;
      border: none;
      background: transparent;
      padding: 0;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      cursor: text;
      outline: none;
    }

    .dcc-product-name-input:focus {
      border-bottom: 1.5px solid #4ade80;
    }

    /* Labels */
    .dcc-label {
      font-size: 11px;
      font-weight: 600;
      color: #555;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      display: block;
      margin-bottom: 5px;
    }

    /* Inputs */
    .dcc-input {
      width: 100%;
      padding: 7px 10px;
      border: 1.5px solid #d1d5db;
      border-radius: 7px;
      font-size: 13px;
      color: #222;
      background: #fff;
      transition: border-color 0.2s, background-color 0.2s;
    }

    .dcc-input:focus {
      outline: none;
      border-color: #4ade80;
    }

    .dcc-input.dcc-active-spend {
      border-color: #16a34a;
      background: #f0fdf4;
    }

    .dcc-input.dcc-inactive-spend {
      border-color: #e5e7eb;
      background: #f9fafb;
      color: #bbb;
    }

    .dcc-input.dcc-invalid {
      border-color: #ef4444;
      background: #fef2f2;
    }

    /* Spend Toggle */
    .dcc-spend-toggle-wrap {
      display: flex;
      background: #f3f4f6;
      border-radius: 8px;
      padding: 3px;
      margin-bottom: 10px;
      gap: 3px;
    }

    .dcc-spend-toggle-btn {
      flex: 1;
      padding: 5px 0;
      border: none;
      border-radius: 6px;
      font-size: 11px;
      font-weight: 600;
      cursor: pointer;
      color: #888;
      background: transparent;
      white-space: nowrap;
      transition: all 0.2s;
    }

    .dcc-spend-toggle-btn.active {
      background: #fff;
      color: #14532d;
      box-shadow: 0 1px 4px rgba(0, 0, 0, 0.10);
    }

    /* Toggle Row */
    .dcc-toggle-row {
      display: flex;
      align-items: center;
      gap: 8px;
      margin-bottom: 10px;
      padding: 7px 10px;
      background: #fff;
      border: 1.5px solid #e5e7eb;
      border-radius: 7px;
    }

    .dcc-toggle-label {
      font-size: 12px;
      color: #374151;
      flex: 1;
    }

    .dcc-toggle-label span {
      font-weight: 600;
      color: #14532d;
    }

    .dcc-days-badge {
      display: inline-block;
      background: #dcfce7;
      color: #15803d;
      border-radius: 4px;
      padding: 1px 6px;
      font-size: 10px;
      font-weight: 700;
      margin-left: 4px;
    }

    /* Toggle Switch */
    .dcc-toggle {
      position: relative;
      display: inline-block;
      width: 36px;
      height: 20px;
      flex-shrink: 0;
    }

    .dcc-toggle input {
      opacity: 0;
      width: 0;
      height: 0;
    }

    .dcc-toggle-slider {
      position: absolute;
      cursor: pointer;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background: #d1d5db;
      border-radius: 20px;
      transition: 0.2s;
    }

    .dcc-toggle-slider:before {
      position: absolute;
      content: "";
      height: 14px;
      width: 14px;
      left: 3px;
      bottom: 3px;
      background: white;
      border-radius: 50%;
      transition: 0.2s;
      box-shadow: 0 1px 3px rgba(0, 0, 0, 0.2);
    }

    .dcc-toggle input:checked + .dcc-toggle-slider {
      background: #16a34a;
    }

    .dcc-toggle input:checked + .dcc-toggle-slider:before {
      transform: translateX(16px);
    }

    /* Spend Row */
    .dcc-spend-row {
      display: flex;
      gap: 10px;
      margin-bottom: 10px;
    }

    .dcc-spend-col {
      flex: 1;
    }

    /* Budget Rows */
    .dcc-budget-row {
      display: flex;
      gap: 7px;
      align-items: center;
      margin-bottom: 6px;
    }

    .dcc-budget-row input[type=text] {
      flex: 1.3;
      padding: 6px 8px;
      border: 1.5px solid #d1d5db;
      border-radius: 6px;
      font-size: 12px;
      color: #222;
      background: #fff;
      transition: border-color 0.2s;
    }

    .dcc-budget-row input[type=text]:focus {
      outline: none;
      border-color: #4ade80;
    }

    .dcc-budget-row input[type=number] {
      flex: 1;
      padding: 6px 8px;
      border: 1.5px solid #d1d5db;
      border-radius: 6px;
      font-size: 12px;
      color: #222;
      background: #fff;
      transition: border-color 0.2s;
    }

    .dcc-budget-row input[type=number]:focus {
      outline: none;
      border-color: #4ade80;
    }

    .dcc-budget-row input.dcc-invalid {
      border-color: #ef4444;
      background: #fef2f2;
    }

    .dcc-dollar {
      color: #aaa;
      font-size: 12px;
    }

    .dcc-budget-wk {
      display: inline-flex;
      align-items: center;
      gap: 3px;
      font-size: 10px;
      font-weight: 700;
      color: #6b7280;
      cursor: pointer;
      user-select: none;
      padding: 3px 6px;
      border: 1.5px solid #e5e7eb;
      border-radius: 6px;
      background: #fff;
      transition: all 0.15s;
    }

    .dcc-budget-wk:hover {
      border-color: #d1d5db;
    }

    .dcc-budget-wk input {
      cursor: pointer;
      margin: 0;
      accent-color: #16a34a;
    }

    .dcc-budget-wk.checked {
      background: #fef9c3;
      border-color: #fde047;
      color: #854d0e;
    }

    /* Buttons */
    .dcc-remove-btn {
      background: none;
      border: none;
      color: #f87171;
      font-size: 16px;
      cursor: pointer;
      padding: 2px 3px;
      line-height: 1;
      transition: color 0.2s;
    }

    .dcc-remove-btn:hover {
      color: #dc2626;
    }

    .dcc-add-budget-btn {
      background: none;
      border: 1.5px solid #4ade80;
      color: #16a34a;
      border-radius: 6px;
      padding: 3px 9px;
      font-size: 11px;
      font-weight: 700;
      cursor: pointer;
      transition: all 0.2s;
    }

    .dcc-add-budget-btn:hover {
      background: #f0fdf4;
    }

    .dcc-remove-product-btn {
      background: none;
      border: none;
      color: #f87171;
      font-size: 12px;
      cursor: pointer;
      padding: 2px 5px;
      white-space: nowrap;
      transition: color 0.2s;
    }

    .dcc-remove-product-btn:hover {
      color: #dc2626;
    }

    /* Event Section */
    .dcc-event-section {
      background: #fffbeb;
      border: 1.5px solid #fde68a;
      border-radius: 7px;
      padding: 9px 11px;
      margin-top: 8px;
    }

    .dcc-event-title {
      font-size: 11px;
      font-weight: 700;
      color: #92400e;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }

    .dcc-event-fields {
      display: flex;
      gap: 8px;
    }

    .dcc-event-col {
      flex: 1;
    }

    .dcc-event-label {
      font-size: 10px;
      font-weight: 600;
      color: #78350f;
      text-transform: uppercase;
      letter-spacing: 0.4px;
      display: block;
      margin-bottom: 4px;
    }

    .dcc-event-input {
      width: 100%;
      padding: 6px 9px;
      border: 1.5px solid #fcd34d;
      border-radius: 6px;
      font-size: 12px;
      color: #222;
      background: #fff;
      transition: border-color 0.2s;
    }

    .dcc-event-input:focus {
      outline: none;
      border-color: #f59e0b;
    }

    /* Results Table */
    .dcc-results-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 12px;
      margin-top: 7px;
    }

    .dcc-results-table th {
      text-align: left;
      padding: 3px 5px 5px 0;
      color: #555;
      font-size: 10px;
      font-weight: 700;
      text-transform: uppercase;
      border-bottom: 1.5px solid #bbf7d0;
    }

    .dcc-results-table th:not(:first-child),
    .dcc-results-table td:not(:first-child) {
      text-align: right;
    }

    .dcc-results-table td {
      padding: 4px 5px 4px 0;
      color: #333;
    }

    .dcc-results-table tr:nth-child(even) td {
      background: #f0fdf4;
    }

    .dcc-daily-cap-val {
      font-weight: 700;
      color: #16a34a;
      font-size: 13px;
    }

    .dcc-event-cap-val {
      font-weight: 700;
      color: #d97706;
      font-size: 12px;
    }

    .dcc-result-summary {
      font-size: 11px;
      color: #555;
      margin-bottom: 6px;
    }

    .dcc-result-summary strong {
      color: #14532d;
    }

    /* Header Buttons */
    .dcc-header-btn {
      background: none;
      border: 1.5px solid rgba(255, 255, 255, 0.4);
      color: #fff;
      border-radius: 6px;
      padding: 4px 10px;
      font-size: 11px;
      font-weight: 600;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 4px;
      transition: all 0.2s;
      white-space: nowrap;
    }

    .dcc-header-btn:hover {
      background: rgba(255, 255, 255, 0.15);
    }

    .dcc-header-btn:disabled {
      opacity: 0.5;
      cursor: not-allowed;
    }

    .dcc-freeze-btn.frozen {
      background: rgba(220, 38, 38, 0.35);
      border-color: rgba(248, 113, 113, 0.7);
      color: #fca5a5;
    }

    .dcc-refresh-btn.spinning .dcc-refresh-icon {
      display: inline-block;
      animation: dcc-spin 0.6s linear infinite;
    }

    @keyframes dcc-spin {
      from { transform: rotate(0deg); }
      to { transform: rotate(360deg); }
    }

    /* Empty State */
    .dcc-empty-state {
      text-align: center;
      padding: 40px 20px;
      color: #6b7280;
    }

    .dcc-empty-state-icon {
      font-size: 48px;
      margin-bottom: 12px;
    }

    .dcc-empty-state-title {
      font-size: 14px;
      font-weight: 600;
      margin-bottom: 6px;
      color: #374151;
    }

    .dcc-empty-state-text {
      font-size: 12px;
      margin-bottom: 16px;
    }

    /* Keyboard Shortcuts Hint */
    .dcc-kbd-hint {
      font-size: 10px;
      color: #86efac;
      margin-top: 2px;
      opacity: 0.8;
    }

    .dcc-kbd {
      display: inline-block;
      padding: 1px 4px;
      background: rgba(255, 255, 255, 0.2);
      border-radius: 3px;
      font-family: monospace;
      margin: 0 2px;
    }

    /* ── Sidebar Dock ────────────────────────────────────────────────────── */
    #daily-cap-calculator.dcc-docked {
      bottom: 0 !important;
      right: 0 !important;
      top: 0 !important;
      height: 100vh !important;
      max-height: 100vh !important;
      border-radius: 0 !important;
      box-shadow: -4px 0 24px rgba(0,0,0,0.22) !important;
      transition: none;
    }

    .dcc-dock-btn.docked {
      background: rgba(74, 222, 128, 0.22);
      border-color: rgba(74, 222, 128, 0.65);
      color: #86efac;
    }

    /* ── Per-budget CPE override (shown when event cap toggle is on) ────── */
    .dcc-budget-cpe-wrap {
      display: none;
      align-items: center;
      gap: 3px;
      flex-shrink: 0;
    }

    .dcc-budget-cpe-wrap.visible {
      display: inline-flex;
    }

    .dcc-budget-cpe-label {
      font-size: 10px;
      font-weight: 700;
      color: #b45309;
      white-space: nowrap;
    }

    .dcc-budget-cpe-wrap input {
      width: 54px;
      padding: 5px 6px;
      border: 1.5px solid #fcd34d;
      border-radius: 5px;
      font-size: 11px;
      color: #222;
      background: #fffbeb;
      transition: border-color 0.2s;
    }

    .dcc-budget-cpe-wrap input:focus {
      outline: none;
      border-color: #f59e0b;
    }

    /* ── Live Pacing border colors on spend inputs ───────────────────────── */
    /* These !important rules beat the active-spend / inactive-spend classes  */
    .dcc-input.dcc-pacing-behind  { border-color: #2563eb !important; background: #eff6ff !important; }
    .dcc-input.dcc-pacing-on-pace { border-color: #16a34a !important; background: #f0fdf4 !important; }
    .dcc-input.dcc-pacing-ahead   { border-color: #dc2626 !important; background: #fef2f2 !important; }
  `;

  // ============================================================================
  // STATE MANAGEMENT
  // ============================================================================

  const State = {
    frozen: false,
    frozenState: null,
    docked: false,
    productCounter: 0,
    budgetCounter: 0,
    dom: {},
    autoSaveTimeout: null,

    captureState() {
      const state = {
        date: this.dom.dateInput?.value || '',
        resultsHTML: this.dom.results?.innerHTML || '',
        resultsDisplay: this.dom.results?.style.display || 'none',
        products: []
      };

      document.querySelectorAll('.dcc-product-block').forEach(productBlock => {
        const budgets = [];
        productBlock.querySelectorAll('.dcc-budget-rows-list .dcc-budget-row').forEach(row => {
          const rawSpent = row.dataset.spent;
          budgets.push({
            desc: row.querySelector('.calc-budget-desc')?.value || '',
            cap: row.querySelector('.calc-budget-cap')?.value || '',
            weekdays: row.querySelector('.calc-budget-weekdays')?.checked || false,
            spent: rawSpent === '' || rawSpent === undefined ? null : parseFloat(rawSpent),
            cpe: row.querySelector('.calc-budget-cpe')?.value || ''
          });
        });

        state.products.push({
          name: productBlock.querySelector('.dcc-product-name')?.value || '',
          spendMode: productBlock.dataset.spendMode || 'total',
          totalVal: productBlock.querySelector('.dcc-input-total')?.value || '',
          remainingVal: productBlock.querySelector('.dcc-input-remaining')?.value || '',
          eventOn: productBlock.querySelector('.dcc-event-toggle')?.checked || false,
          eventType: productBlock.querySelector('.dcc-event-type')?.value || '',
          cpe: productBlock.querySelector('.dcc-cpe-input')?.value || '',
          budgets
        });
      });

      return state;
    },

    restoreState(state) {
      if (!state) return;

      Utils.log('Restoring state:', state);

      this.dom.dateInput.value = Utils.formatDate(new Date());
      if (state.resultsHTML) this.dom.results.innerHTML = state.resultsHTML;
      if (state.resultsDisplay) this.dom.results.style.display = state.resultsDisplay;

      if (state.products && state.products.length > 0) {
        this.dom.productsList.innerHTML = '';
        this.productCounter = 0;
        this.budgetCounter = 0;

        state.products.forEach(productData => {
          UI.addProduct(productData);
        });
      }
    }
  };

  // ============================================================================
  // STORAGE
  // ============================================================================

  const Storage = {
    save() {
      try {
        const state = State.captureState();
        localStorage.setItem(CONFIG.STORAGE_KEY_STATE, JSON.stringify(state));
        // Also save frozen state
        localStorage.setItem(CONFIG.STORAGE_KEY_FROZEN, State.frozen ? 'true' : 'false');
        Utils.log('State saved (frozen:', State.frozen, ')');
      } catch (err) {
        console.error('Failed to save state:', err);
      }
    },

    load() {
      try {
        const saved = localStorage.getItem(CONFIG.STORAGE_KEY_STATE);
        if (saved) {
          const state = JSON.parse(saved);
          Utils.log('State loaded');
          return state;
        }
      } catch (err) {
        console.error('Failed to load state:', err);
      }
      return null;
    },

    loadFrozenState() {
      try {
        const frozen = localStorage.getItem(CONFIG.STORAGE_KEY_FROZEN);
        return frozen === 'true';
      } catch (err) {
        console.error('Failed to load frozen state:', err);
      }
      return false;
    },

    autoSave: Utils.debounce(() => {
      if (!State.frozen) {
        Storage.save();
      }
    }, CONFIG.AUTO_SAVE_DELAY),

    exportConfig() {
      const state = State.captureState();
      const json = JSON.stringify(state, null, 2);
      const blob = new Blob([json], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      const today = new Date().toISOString().split('T')[0];
      a.href = url;
      a.download = `daily-cap-config-${today}.json`;
      a.click();
      URL.revokeObjectURL(url);
      Utils.showToast('Configuration exported');
    },

    importConfig() {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = '.json';
      input.onchange = (e) => {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (ev) => {
          try {
            const state = JSON.parse(ev.target.result);
            State.restoreState(state);
            Utils.showToast('Configuration imported');
          } catch (err) {
            console.error('Import failed:', err);
            Utils.showToast('Invalid configuration file', 'error');
          }
        };
        reader.readAsText(file);
      };
      input.click();
    }
  };

  // ============================================================================
  // CALCULATIONS
  // ============================================================================

  const Calculator = {
    countDays(dateStr, weekdaysOnly) {
      if (!dateStr) return null;

      try {
        const [year, month, day] = dateStr.split('-').map(Number);

        // Validate date components
        if (!year || !month || !day || month < 1 || month > 12 || day < 1) {
          return null;
        }

        const daysInMonth = new Date(year, month, 0).getDate();

        if (day > daysInMonth) return null;

        let count = 0;
        for (let i = day; i <= daysInMonth; i++) {
          const dayOfWeek = new Date(year, month - 1, i).getDay();
          if (!weekdaysOnly || (dayOfWeek !== 0 && dayOfWeek !== 6)) {
            count++;
          }
        }

        return count;
      } catch (err) {
        console.error('Date calculation error:', err);
        return null;
      }
    },

    // Days from the 1st of the month through (and including) `dateStr`
    countElapsedDays(dateStr, weekdaysOnly) {
      if (!dateStr) return null;
      try {
        const [year, month, day] = dateStr.split('-').map(Number);
        if (!year || !month || !day) return null;
        let count = 0;
        for (let i = 1; i <= day; i++) {
          const dow = new Date(year, month - 1, i).getDay();
          if (!weekdaysOnly || (dow !== 0 && dow !== 6)) count++;
        }
        return count;
      } catch {
        return null;
      }
    },

    // Total days in the month (full month, not just remaining)
    countTotalDaysInMonth(dateStr, weekdaysOnly) {
      if (!dateStr) return null;
      try {
        const [year, month] = dateStr.split('-').map(Number);
        if (!year || !month) return null;
        const daysInMonth = new Date(year, month, 0).getDate();
        let count = 0;
        for (let i = 1; i <= daysInMonth; i++) {
          const dow = new Date(year, month - 1, i).getDay();
          if (!weekdaysOnly || (dow !== 0 && dow !== 6)) count++;
        }
        return count;
      } catch {
        return null;
      }
    },

    // Time-adjusted pacing: 100 = on pace, >100 = ahead (overspending), <100 = behind
    computePacing(spent, cap, elapsedDays, totalDays) {
      if (!(spent >= 0) || !(cap > 0) || !(elapsedDays > 0) || !(totalDays > 0)) return null;
      return (spent / cap) / (elapsedDays / totalDays) * 100;
    },

    pacingStyle(pct) {
      if (pct == null || !isFinite(pct)) return { color: '#6b7280', label: '—' };
      if (pct < 90) return { color: '#2563eb', label: 'behind' };
      if (pct > 110) return { color: '#dc2626', label: 'ahead' };
      return { color: '#16a34a', label: 'on pace' };
    },

    // Instantly colour a spend input's border to reflect current pacing.
    // Called on every keystroke in the TSA / Remaining fields.
    applyLivePacingToInput(input, productBlock) {
      const PACING_CLASSES = ['dcc-pacing-behind', 'dcc-pacing-on-pace', 'dcc-pacing-ahead'];
      PACING_CLASSES.forEach(c => input.classList.remove(c));

      const dateValue = State.dom.dateInput?.value;
      if (!dateValue) return;

      const tsa = parseFloat(productBlock.querySelector('.dcc-input-total')?.value);
      const rem = parseFloat(productBlock.querySelector('.dcc-input-remaining')?.value);
      // Need both values to compute implied spend
      if (isNaN(tsa) || tsa <= 0 || isNaN(rem)) return;

      const spent   = Math.max(0, tsa - rem);
      const elapsed = this.countElapsedDays(dateValue, false);
      const total   = this.countTotalDaysInMonth(dateValue, false);
      const pct     = this.computePacing(spent, tsa, elapsed, total);

      if (pct == null) return;
      if (pct < 90)      input.classList.add('dcc-pacing-behind');
      else if (pct > 110) input.classList.add('dcc-pacing-ahead');
      else                input.classList.add('dcc-pacing-on-pace');
    },

    validateInput(input) {
      const value = parseFloat(input.value);
      const isValid = !input.value || (!isNaN(value) && value > 0);

      if (input.value && !isValid) {
        input.classList.add('dcc-invalid');
      } else {
        input.classList.remove('dcc-invalid');
      }

      return isValid;
    },

    calculate() {
      const dateValue = State.dom.dateInput.value;
      const results = State.dom.results;
      let html = '';

      if (!dateValue) {
        results.style.display = 'block';
        results.innerHTML = '<div style="color:#dc2626;font-size:13px;padding:8px;background:#fef2f2;border-radius:7px;border:1.5px solid #fecaca;">Please select a valid date.</div>';
        return;
      }

      document.querySelectorAll('.dcc-product-block').forEach(productBlock => {
        const name = Utils.escapeHtml(productBlock.querySelector('.dcc-product-name').value.trim() || 'Product');
        const mode = productBlock.dataset.spendMode || 'total';
        const spendValue = mode === 'total'
          ? parseFloat(productBlock.querySelector('.dcc-input-total').value)
          : parseFloat(productBlock.querySelector('.dcc-input-remaining').value);

        const useEventCap = productBlock.querySelector('.dcc-event-toggle').checked;
        const cpe = parseFloat(productBlock.querySelector('.dcc-cpe-input').value);
        const cpeLabel = Utils.escapeHtml(productBlock.querySelector('.dcc-event-type').value.trim() || 'events');

        // For overall pacing — need both TSA and Remaining regardless of selected mode
        const totalValInput = parseFloat(productBlock.querySelector('.dcc-input-total').value);
        const remainingValInput = parseFloat(productBlock.querySelector('.dcc-input-remaining').value);

        const budgets = [];
        let totalMonthlyCap = 0;

        productBlock.querySelectorAll('.dcc-budget-rows-list .dcc-budget-row').forEach(row => {
          const desc = row.querySelector('.calc-budget-desc').value.trim() || 'Budget';
          const cap = parseFloat(row.querySelector('.calc-budget-cap').value);
          const weekdays = row.querySelector('.calc-budget-weekdays').checked;
          const rawSpent = row.dataset.spent;
          const spent = rawSpent === '' || rawSpent === undefined ? null : parseFloat(rawSpent);

          // Per-budget CPE override (null = use product-level fallback)
          const rawBudgetCpe = parseFloat(row.querySelector('.calc-budget-cpe')?.value);
          const budgetCpe = !isNaN(rawBudgetCpe) && rawBudgetCpe > 0 ? rawBudgetCpe : null;

          if (!isNaN(cap) && cap > 0) {
            budgets.push({ desc, cap, weekdays, spent: isNaN(spent) ? null : spent, budgetCpe });
            totalMonthlyCap += cap;
          }
        });

        // Validation
        if (isNaN(spendValue) || spendValue <= 0) {
          html += `<div style="background:#fef2f2;border:1.5px solid #fecaca;border-radius:8px;padding:10px 12px;margin-bottom:10px;color:#dc2626;font-size:12px;">
            ${name}: Enter a valid ${mode === 'total' ? 'Total Spend Available' : 'Remaining Spend'}.
          </div>`;
          return;
        }

        if (!budgets.length) {
          html += `<div style="background:#fef2f2;border:1.5px solid #fecaca;border-radius:8px;padding:10px 12px;margin-bottom:10px;color:#dc2626;font-size:12px;">
            ${name}: Add at least one budget with Monthly Cap > 0.
          </div>`;
          return;
        }

        // Precompute per-budget day counts + pacing — bail if days are invalid
        let anyInvalidDays = false;
        budgets.forEach(b => {
          b.days = this.countDays(dateValue, b.weekdays);
          if (b.days === null || b.days <= 0) anyInvalidDays = true;
          const elapsed = this.countElapsedDays(dateValue, b.weekdays);
          const total = this.countTotalDaysInMonth(dateValue, b.weekdays);
          b.pacingPct = (b.spent != null && !isNaN(b.spent))
            ? this.computePacing(b.spent, b.cap, elapsed, total)
            : null;
        });

        if (anyInvalidDays) {
          html += `<div style="background:#fef2f2;border:1.5px solid #fecaca;border-radius:8px;padding:10px 12px;margin-bottom:10px;color:#dc2626;font-size:12px;">
            ${name}: Invalid date or no days remaining in month.
          </div>`;
          return;
        }

        // Overall pacing (calendar days): prefer sum of per-budget spend; fall back to TSA − Remaining
        const allSpentKnown = budgets.length > 0 && budgets.every(b => b.spent != null && !isNaN(b.spent));
        let overallSpent = null;
        if (allSpentKnown) {
          overallSpent = budgets.reduce((s, b) => s + b.spent, 0);
        } else if (!isNaN(totalValInput) && !isNaN(remainingValInput) && totalValInput > 0) {
          overallSpent = Math.max(0, totalValInput - remainingValInput);
        }
        const overallCap = !isNaN(totalValInput) && totalValInput > 0 ? totalValInput : totalMonthlyCap;
        const overallElapsed = this.countElapsedDays(dateValue, false);
        const overallTotal = this.countTotalDaysInMonth(dateValue, false);
        const overallPacing = this.computePacing(overallSpent, overallCap, overallElapsed, overallTotal);

        const modeLabel = mode === 'total'
          ? '<span style="background:#dbeafe;color:#1d4ed8;border-radius:4px;padding:1px 6px;font-size:10px;font-weight:700;margin-left:4px;display:inline-block;">Total Spend</span>'
          : '<span style="background:#fef3c7;color:#92400e;border-radius:4px;padding:1px 6px;font-size:10px;font-weight:700;margin-left:4px;display:inline-block;">Remaining Spend</span>';

        // Show alert if spend doesn't match caps
        let mismatchAlert = '';
        if (Math.abs(spendValue - totalMonthlyCap) > 0.01) {
          const diff = spendValue - totalMonthlyCap;
          const diffType = diff > 0 ? 'more' : 'less';
          const diffColor = diff > 0 ? '#2563eb' : '#ea580c';
          const spendLabel = mode === 'total' ? 'Total spend' : 'Remaining spend';
          mismatchAlert = `<div style="background:#fef3c7;border:1.5px solid #fde68a;border-radius:6px;padding:6px 8px;margin-bottom:8px;font-size:11px;color:#92400e;">
            ℹ️ ${spendLabel} is <strong style="color:${diffColor};">$${Math.abs(diff).toFixed(2)} ${diffType}</strong> than total caps ($${totalMonthlyCap.toFixed(2)}). Daily caps distributed proportionally.
          </div>`;
        }

        // Event cap column is shown when the toggle is on AND at least one CPE exists
        // (either the product-level default or a per-budget override).
        const productCpeValid = !isNaN(cpe) && cpe > 0;
        const hasAnyEventCap = useEventCap && (
          productCpeValid || budgets.some(b => b.budgetCpe != null)
        );

        const eventCapHeader = hasAnyEventCap
          ? `<th style="color:#d97706;">${cpeLabel} Cap</th>`
          : '';

        let tableRows = '';
        budgets.forEach(budget => {
          // Share is based on budget cap proportions
          const share = budget.cap / totalMonthlyCap;
          // Each budget gets its proportional slice of spend, divided by its own day count
          const budgetSpend = spendValue * share;
          const dailyCap = budgetSpend / budget.days;
          const dayLabel = budget.weekdays ? 'weekday' : 'day';
          const daysCell = `${budget.days} ${dayLabel}${budget.days !== 1 ? 's' : ''}`;

          // Effective CPE: per-budget override wins, else fall back to product default
          const effectiveCpe = budget.budgetCpe ?? (productCpeValid ? cpe : null);
          const cpeSource = budget.budgetCpe != null ? '(budget rate)' : '(product rate)';
          const eventCapCell = hasAnyEventCap
            ? (effectiveCpe != null
                ? `<td class="dcc-event-cap-val" title="CPE $${effectiveCpe.toFixed(2)} ${cpeSource}">${Math.round(dailyCap / effectiveCpe)}</td>`
                : '<td style="color:#9ca3af;" title="No CPE set for this budget">—</td>')
            : '';

          const pStyle = this.pacingStyle(budget.pacingPct);
          const pacingCell = budget.pacingPct != null
            ? `<td style="color:${pStyle.color};font-weight:700;" title="Spent $${budget.spent.toFixed(2)} of $${budget.cap.toFixed(2)} (${(budget.spent / budget.cap * 100).toFixed(0)}% of cap)">${budget.pacingPct.toFixed(0)}%</td>`
            : '<td style="color:#9ca3af;">—</td>';

          tableRows += `
            <tr>
              <td>${Utils.escapeHtml(budget.desc)}</td>
              <td>$${budget.cap.toFixed(2)}</td>
              <td>${(share * 100).toFixed(1)}%</td>
              <td>${daysCell}</td>
              ${pacingCell}
              <td class="dcc-daily-cap-val">$${dailyCap.toFixed(2)}</td>
              ${eventCapCell}
            </tr>
          `;
        });

        const footerColspan = hasAnyEventCap ? 7 : 6;
        const eventFooter = hasAnyEventCap
          ? (() => {
              const defaultNote = productCpeValid
                ? `default $${cpe.toFixed(2)} per ${cpeLabel.toLowerCase().replace(/s$/, '')}`
                : 'no product default';
              const hasOverrides = budgets.some(b => b.budgetCpe != null);
              const overrideNote = hasOverrides ? ' · per-budget rates override default' : '';
              return `<tr><td colspan="${footerColspan}" style="padding-top:6px;font-size:10px;color:#92400e;background:#fffbeb;">
                &#128200; Event Cap = Daily Cap &divide; CPE (${defaultNote}${overrideNote})
              </td></tr>`;
            })()
          : '';

        // Overall pacing summary line (only when we have enough data to compute it)
        let overallPacingLine = '';
        if (overallPacing != null && overallSpent != null && overallCap > 0) {
          const oStyle = this.pacingStyle(overallPacing);
          overallPacingLine = `<div style="font-size:11px;margin-top:3px;">
            Overall pacing: <strong style="color:${oStyle.color};">${overallPacing.toFixed(0)}% · ${oStyle.label}</strong>
            <span style="color:#6b7280;"> · spent $${overallSpent.toFixed(2)} of $${overallCap.toFixed(2)} (${(overallSpent / overallCap * 100).toFixed(0)}%)</span>
          </div>`;
        }

        html += `
          <div style="background:#f0fdf4;border:1.5px solid #bbf7d0;border-radius:8px;padding:11px 13px;margin-bottom:10px;">
            <div style="display:flex;align-items:center;flex-wrap:wrap;gap:4px;margin-bottom:5px;">
              <span style="font-size:12px;font-weight:700;color:#14532d;text-transform:uppercase;">${name}</span>
              ${modeLabel}
            </div>
            ${mismatchAlert}
            <div class="dcc-result-summary">
              <strong>$${spendValue.toFixed(2)}</strong> across <strong>${budgets.length} budget${budgets.length !== 1 ? 's' : ''}</strong>
            </div>
            ${overallPacingLine}
            <table class="dcc-results-table">
              <thead>
                <tr>
                  <th>Budget</th>
                  <th>Monthly Cap</th>
                  <th>Share</th>
                  <th>Days</th>
                  <th title="Time-adjusted pacing: 100% = on pace, &gt;110% = ahead, &lt;90% = behind">Pacing</th>
                  <th>Daily Cap</th>
                  ${eventCapHeader}
                </tr>
              </thead>
              <tbody>
                ${tableRows}
                ${eventFooter}
              </tbody>
            </table>
          </div>
        `;
      });

      results.style.display = 'block';
      results.innerHTML = html || '<div style="color:#666;font-size:13px;padding:8px;">No products to calculate.</div>';

      // Save state after calculation (even if frozen, to preserve results)
      Storage.save();
    }
  };

  // ============================================================================
  // DATA SCRAPING
  // ============================================================================

  const DataScraper = {
    scrapePageData() {
      const products = [];

      try {
        document.querySelectorAll('.account-section').forEach(section => {
          try {
            const heading = section.querySelector('.mb-0');
            const name = heading?.textContent.trim();

            if (!name || name.length > 30) return;

            const tsaEl = section.querySelector('.spend-pace__top-value');
            const tsa = Utils.parseMoney(tsaEl?.textContent);

            let remaining = NaN;
            const walker = document.createTreeWalker(section, NodeFilter.SHOW_TEXT, null);
            let node;
            while (node = walker.nextNode()) {
              // Case-insensitive partial match — handles "Remaining Spend:", "remaining spend", etc.
              if (/remaining\s*spend/i.test(node.textContent.trim())) {
                remaining = Utils.parseMoney(
                  node.parentElement?.nextElementSibling?.textContent
                );
                break;
              }
            }

            const nameLower = name.toLowerCase();
            const weekdaysDefault = nameLower.includes('call') || nameLower.includes('lead');

            // Per-table: detect the "Spend" column index via header text so pacing works
            // even if column positions shift. Collect rows with their resolved spendIdx.
            const scoredRows = [];
            [...section.querySelectorAll('table')].forEach(table => {
              const headerRow = table.querySelector('thead tr') || table.querySelector('tr');
              const headerCells = headerRow ? [...headerRow.querySelectorAll('th, td')] : [];
              const spendIdx = headerCells.findIndex(c => {
                const t = c.textContent.trim().toLowerCase();
                return /\bspend\b|\bspent\b/.test(t) && !/(per|cost|daily)/.test(t);
              });
              [...table.querySelectorAll('tr')].forEach(row => {
                if (row.querySelectorAll('td').length >= 5) {
                  scoredRows.push({ row, spendIdx });
                }
              });
            });

            const budgets = [];
            scoredRows.forEach(({ row, spendIdx }) => {
              const cells = [...row.querySelectorAll('td')];
              const desc = cells[0]?.textContent.trim();
              const capText = cells[3]?.textContent || '';
              let cap = Utils.parseMoney(capText);

              // "Continuous" budgets have no hard monthly cap — use Total Spend Available instead
              if (/continuous/i.test(capText) && !isNaN(tsa) && tsa > 0) {
                cap = tsa;
              }

              const spentVal = spendIdx >= 0 ? Utils.parseMoney(cells[spendIdx]?.textContent) : NaN;

              if (desc && !isNaN(cap) && cap > 0) {
                budgets.push({
                  desc,
                  cap,
                  weekdays: weekdaysDefault,
                  spent: isNaN(spentVal) ? null : spentVal
                });
              }
            });

            products.push({
              name,
              tsa: isNaN(tsa) ? '' : tsa,
              remaining: isNaN(remaining) ? '' : remaining,
              budgets,
              weekdaysDefault
            });
          } catch (err) {
            Utils.log('Failed to scrape section:', err);
          }
        });

        // --- Advertising page structure (.dashboard-collapse-outer-container) ---
        // Only runs when no .account-section products found (i.e. not on MelonMax)
        if (products.length === 0) {
          document.querySelectorAll('.dashboard-collapse-outer-container').forEach(container => {
            try {
              const h2 = container.querySelector('h2');
              const rawName = h2?.textContent.trim();
              if (!rawName) return;

              // Strip trailing " Budgets" to get a clean product name
              const name = rawName.replace(/\s*Budgets\s*$/i, '').trim();
              if (!name || name.length > 50) return;

              const tables = container.querySelectorAll('table');
              if (tables.length < 2) return; // Need budget table + detail table

              // --- Extract TSA and Remaining Spend from the detail row table ---
              // Detail table cells contain "Spend Available $X" and "Remaining Spend $Y" inline
              let tsa = NaN;
              let remaining = NaN;

              for (let t = 1; t < tables.length; t++) {
                const detailRow = tables[t].querySelector('tr');
                if (!detailRow) continue;
                const cells = detailRow.querySelectorAll('td');

                cells.forEach(td => {
                  const text = td.textContent;
                  // Case-insensitive partial matches for resilience against label wording changes
                  if (/spend\s*available/i.test(text)) {
                    const match = text.match(/\$([0-9,]+\.?[0-9]*)/);
                    if (match) tsa = parseFloat(match[1].replace(/,/g, ''));
                  }
                  if (/remaining\s*spend/i.test(text)) {
                    const match = text.match(/\$([0-9,]+\.?[0-9]*)/);
                    if (match) remaining = parseFloat(match[1].replace(/,/g, ''));
                  }
                });

                if (!isNaN(tsa)) break;
              }

              // --- Extract budget rows from the main budget table ---
              // Headers: [1]=Budget Description, [6]=Monthly Spend ("$750 (100%)" or "$0 (0%)\n<i>$500</i>")
              const budgets = [];
              const rows = tables[0].querySelectorAll('tr');
              const nameLower = name.toLowerCase();
              const productLevelWeekdays = nameLower.includes('call') || nameLower.includes('lead');

              rows.forEach(row => {
                const cells = row.querySelectorAll('td');
                if (cells.length < 7) return; // Skip <th> header rows

                const desc = cells[1]?.textContent.trim().split('\n')[0].trim();
                if (!desc) return;

                // Per-budget weekdays: "Call Only" budget type rows default to weekdays-only
                const isCallOnly = /call\s*only/i.test(row.textContent);
                const weekdays = isCallOnly || productLevelWeekdays;

                // cells[6] format: "$<spent> (<pct>%)" optionally followed by "<i>$<cap></i>" when spend ≠ cap
                const capCell = cells[6];
                const spentMatch = capCell?.textContent.match(/\$([0-9,]+\.?[0-9]*)/);
                const spent = spentMatch ? parseFloat(spentMatch[1].replace(/,/g, '')) : null;

                const iEl = capCell?.querySelector('i');
                let cap = NaN;
                if (iEl) {
                  cap = parseFloat(iEl.textContent.replace(/[^0-9.]/g, ''));
                } else if (spent != null && !isNaN(spent)) {
                  // No <i> tag means spend equals cap
                  cap = spent;
                }

                if (desc && !isNaN(cap) && cap > 0) {
                  budgets.push({
                    desc,
                    cap,
                    weekdays,
                    spent: spent != null && !isNaN(spent) ? spent : null
                  });
                }
              });

              if (budgets.length === 0) return;

              products.push({
                name,
                tsa: isNaN(tsa) ? '' : tsa,
                remaining: isNaN(remaining) ? '' : remaining,
                budgets,
                weekdaysDefault: productLevelWeekdays,
                cpe: '',
                cpeLabel: ''
              });

            } catch (err) {
              Utils.log('Failed to scrape advertising container:', err);
            }
          });
        }

        // Scrape CPE data - but only for MelonMax pages
        const isMelonMaxPage = window.location.href.toLowerCase().includes('melonmax');
        const cpeMap = {};

        if (isMelonMaxPage) {
          // Find a header column by its label text, anywhere on the page.
          // Data grids often render headers outside the <table>, so searching globally is necessary.
          const headerIndexOf = (label) => {
            const normalized = label.toLowerCase();
            for (const el of document.querySelectorAll('th, td, [role="columnheader"]')) {
              if (el.textContent.trim().toLowerCase() === normalized) {
                const parent = el.parentElement;
                if (parent) return [...parent.children].indexOf(el);
              }
            }
            return -1;
          };

          const hMonth = headerIndexOf('Month');
          const hCpc = headerIndexOf('Cost Per Click');
          const hCpCall = headerIndexOf('Cost Per Call');
          const hCpLead = headerIndexOf('Cost Per Lead');

          // Data rows have a leading expand-arrow cell (Month lives at cell index 1).
          // If headers live in a separate structure without that cell, we need to offset.
          const DATA_MONTH_IDX = 1;
          const offset = hMonth >= 0 ? DATA_MONTH_IDX - hMonth : 0;

          const cpcIdx = hCpc >= 0 ? hCpc + offset : 7;
          const cpCallIdx = hCpCall >= 0 ? hCpCall + offset : 10;
          const cpLeadIdx = hCpLead >= 0 ? hCpLead + offset : 13;

          document.querySelectorAll('table').forEach(table => {
            table.querySelectorAll('tr').forEach(row => {
              const cells = [...row.querySelectorAll('td')];
              if (cells.length < 13 || !cells[1]?.textContent.match(/20\d{2}/)) return;

              const cpc = Utils.parseMoney(cells[cpcIdx]?.textContent);
              const cpCall = Utils.parseMoney(cells[cpCallIdx]?.textContent);
              const cpLead = Utils.parseMoney(cells[cpLeadIdx]?.textContent);

              if (!cpeMap.clicks && cpc > 0) cpeMap.clicks = cpc;
              if (!cpeMap.calls && cpCall > 0) cpeMap.calls = cpCall;
              if (!cpeMap.leads && cpLead > 0) cpeMap.leads = cpLead;
            });
          });
        }

        // Assign CPE to products - only for MelonMax
        products.forEach(product => {
          if (isMelonMaxPage) {
            const nameLower = product.name.toLowerCase();
            product.cpe = nameLower.includes('click') ? cpeMap.clicks || ''
              : nameLower.includes('call') ? cpeMap.calls || ''
              : nameLower.includes('lead') ? cpeMap.leads || ''
              : '';
            product.cpeLabel = product.name;
          } else {
            // Advertising pages don't use event caps
            product.cpe = '';
            product.cpeLabel = '';
          }
        });

        Utils.log('Scraped products:', products);
      } catch (err) {
        console.error('Page scraping failed:', err);
      }

      return products.length ? products : this.getDefaultProducts();
    },

    getDefaultProducts() {
      return [
        { name: 'Calls', weekdaysDefault: true, budgets: [], tsa: '', remaining: '', cpe: '', cpeLabel: 'Calls' },
        { name: 'Clicks', weekdaysDefault: false, budgets: [], tsa: '', remaining: '', cpe: '', cpeLabel: 'Clicks' }
      ];
    }
  };

  // ============================================================================
  // UI COMPONENTS
  // ============================================================================

  const UI = {
    addBudgetRow(listEl, desc = '', cap = '', weekdays = false, spent = null, cpe = '', showCpe = false) {
      State.budgetCounter++;
      const row = document.createElement('div');
      row.className = 'dcc-budget-row';
      row.id = `brow-${State.budgetCounter}`;
      if (spent !== null && spent !== undefined && !isNaN(spent)) {
        row.dataset.spent = String(spent);
      }

      const descInput = document.createElement('input');
      descInput.type = 'text';
      descInput.placeholder = 'Description';
      descInput.value = desc;
      descInput.className = 'calc-budget-desc';
      descInput.addEventListener('input', () => Storage.autoSave());

      const dollar = document.createElement('span');
      dollar.className = 'dcc-dollar';
      dollar.textContent = '$';

      const capInput = document.createElement('input');
      capInput.type = 'number';
      capInput.placeholder = 'Monthly Cap';
      capInput.value = cap;
      capInput.min = '0';
      capInput.step = '0.01';
      capInput.className = 'calc-budget-cap';
      capInput.addEventListener('blur', () => Calculator.validateInput(capInput));
      capInput.addEventListener('input', () => Storage.autoSave());

      const wkLabel = document.createElement('label');
      wkLabel.className = 'dcc-budget-wk' + (weekdays ? ' checked' : '');
      wkLabel.title = 'Weekdays only (Mon–Fri)';
      const wkInput = document.createElement('input');
      wkInput.type = 'checkbox';
      wkInput.className = 'calc-budget-weekdays';
      wkInput.checked = !!weekdays;
      const wkText = document.createElement('span');
      wkText.textContent = 'M–F';
      wkLabel.append(wkInput, wkText);
      wkInput.addEventListener('change', () => {
        wkLabel.classList.toggle('checked', wkInput.checked);
        Storage.autoSave();
      });

      // Per-budget CPE override — visible only when the product event toggle is on
      const cpeWrap = document.createElement('div');
      cpeWrap.className = 'dcc-budget-cpe-wrap' + (showCpe ? ' visible' : '');
      cpeWrap.title = 'Cost per event for this budget type (leave blank to use the product default)';
      const cpeLabel = document.createElement('span');
      cpeLabel.className = 'dcc-budget-cpe-label';
      cpeLabel.textContent = '$/evt';
      const cpeInput = document.createElement('input');
      cpeInput.type = 'number';
      cpeInput.placeholder = 'CPE';
      cpeInput.value = cpe || '';
      cpeInput.min = '0';
      cpeInput.step = '0.01';
      cpeInput.className = 'calc-budget-cpe';
      cpeInput.addEventListener('input', () => Storage.autoSave());
      cpeWrap.append(cpeLabel, cpeInput);

      const removeBtn = document.createElement('button');
      removeBtn.className = 'dcc-remove-btn';
      removeBtn.innerHTML = '&times;';
      removeBtn.title = 'Remove budget';

      row.append(descInput, dollar, capInput, wkLabel, cpeWrap, removeBtn);
      listEl.appendChild(row);
    },

    setSpendMode(productBlock, mode) {
      productBlock.dataset.spendMode = mode;

      const btnTotal = productBlock.querySelector('.dcc-btn-total');
      const btnRemaining = productBlock.querySelector('.dcc-btn-remaining');
      const inputTotal = productBlock.querySelector('.dcc-input-total');
      const inputRemaining = productBlock.querySelector('.dcc-input-remaining');
      const labelTotal = productBlock.querySelector('.dcc-label-total');
      const labelRemaining = productBlock.querySelector('.dcc-label-remaining');

      if (mode === 'total') {
        btnTotal.classList.add('active');
        btnRemaining.classList.remove('active');
        inputTotal.classList.add('dcc-active-spend');
        inputTotal.classList.remove('dcc-inactive-spend');
        inputRemaining.classList.add('dcc-inactive-spend');
        inputRemaining.classList.remove('dcc-active-spend');
        labelTotal.style.color = '';
        labelRemaining.style.color = '#bbb';
      } else {
        btnRemaining.classList.add('active');
        btnTotal.classList.remove('active');
        inputRemaining.classList.add('dcc-active-spend');
        inputRemaining.classList.remove('dcc-inactive-spend');
        inputTotal.classList.add('dcc-inactive-spend');
        inputTotal.classList.remove('dcc-active-spend');
        labelRemaining.style.color = '';
        labelTotal.style.color = '#bbb';
      }

      Storage.autoSave();
    },

    addProduct(options = {}) {
      State.productCounter++;
      const productId = State.productCounter;

      const {
        name = '',
        tsa = '',
        remaining = '',
        budgets = [],
        weekdaysDefault = false,
        cpe = '',
        cpeLabel = '',
        spendMode = 'total',
        totalVal = tsa,
        remainingVal = remaining,
        eventOn = false,
        eventType = cpeLabel
      } = options;

      const container = State.dom.productsList;
      const productBlock = document.createElement('div');
      productBlock.className = 'dcc-product-block';
      productBlock.id = `product-${productId}`;

      productBlock.innerHTML = `
        <div class="dcc-product-header">
          <input type="text" placeholder="Product name" value="${Utils.escapeHtml(name)}" class="dcc-product-name-input dcc-product-name">
          <button class="dcc-remove-product-btn" title="Remove product">&#x2715; Remove</button>
        </div>

        <div style="margin-bottom:10px;">
          <label class="dcc-label">Calculate from</label>
          <div class="dcc-spend-toggle-wrap">
            <button class="dcc-spend-toggle-btn active dcc-btn-total">Total Spend Available</button>
            <button class="dcc-spend-toggle-btn dcc-btn-remaining">Remaining Spend</button>
          </div>
        </div>

        <div class="dcc-spend-row">
          <div class="dcc-spend-col">
            <label class="dcc-label dcc-label-total">Total Spend Available ($)</label>
            <input type="number" placeholder="e.g. 2088.16" value="${totalVal}" min="0" step="0.01" class="dcc-input dcc-active-spend dcc-input-total">
          </div>
          <div class="dcc-spend-col">
            <label class="dcc-label dcc-label-remaining" style="color:#bbb;">Remaining Spend ($)</label>
            <input type="number" placeholder="e.g. 1012.12" value="${remainingVal}" min="0" step="0.01" class="dcc-input dcc-inactive-spend dcc-input-remaining">
          </div>
        </div>

        <div style="margin-bottom:10px;">
          <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px;">
            <label class="dcc-label" style="margin:0;">Individual Budgets <span style="font-weight:500;color:#888;text-transform:none;letter-spacing:0;">— toggle M–F per budget</span></label>
            <button class="dcc-add-budget-btn dcc-add-brow-btn">+ Add Budget</button>
          </div>
          <div class="dcc-budget-rows-list"></div>
        </div>

        <div class="dcc-toggle-row" style="border-color:#fde68a;background:#fffbeb;margin-bottom:0;">
          <div class="dcc-toggle-label">
            <span style="color:#92400e;">Convert to Event Cap</span>
            <span style="font-size:11px;color:#b45309;margin-left:6px;">&#247; cost per event</span>
          </div>
          <label class="dcc-toggle">
            <input type="checkbox" class="dcc-event-toggle">
            <span class="dcc-toggle-slider"></span>
          </label>
        </div>

        <div class="dcc-event-section" style="display:none;">
          <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;">
            <span class="dcc-event-title">&#128200; Event Cap Conversion</span>
            <span style="font-size:10px;color:#b45309;">Based on previous month</span>
          </div>
          <div class="dcc-event-fields">
            <div class="dcc-event-col">
              <label class="dcc-event-label">Event Type</label>
              <input type="text" placeholder="e.g. Clicks" value="${Utils.escapeHtml(eventType)}" class="dcc-event-type dcc-event-input">
            </div>
            <div class="dcc-event-col">
              <label class="dcc-event-label">Cost Per Event ($)</label>
              <input type="number" placeholder="e.g. 6.97" value="${cpe}" min="0" step="0.01" class="dcc-cpe-input dcc-event-input">
            </div>
          </div>
        </div>
      `;

      container.appendChild(productBlock);

      // Set initial mode
      this.setSpendMode(productBlock, spendMode);

      // Add event listeners for inputs
      productBlock.querySelectorAll('input').forEach(input => {
        input.addEventListener('input', () => Storage.autoSave());
        if (input.type === 'number') {
          input.addEventListener('blur', () => Calculator.validateInput(input));
        }
      });

      productBlock.querySelector('.dcc-product-name').addEventListener('input', () => Storage.autoSave());

      // Live pacing border colors — fire on every keystroke in either spend field
      const tsaInputEl = productBlock.querySelector('.dcc-input-total');
      const remInputEl = productBlock.querySelector('.dcc-input-remaining');
      [tsaInputEl, remInputEl].forEach(el => {
        if (el) el.addEventListener('input', () => Calculator.applyLivePacingToInput(el, productBlock));
      });

      // Budgets list
      const budgetsList = productBlock.querySelector('.dcc-budget-rows-list');

      // Event toggle — resolve here so we know showCpe for addBudgetRow
      const eventToggle    = productBlock.querySelector('.dcc-event-toggle');
      const eventSection   = productBlock.querySelector('.dcc-event-section');
      const eventToggleRow = productBlock.querySelector('.dcc-toggle-row');

      // Helper: show or hide all per-budget CPE wraps within this product
      const setCpeWrapsVisible = (visible) => {
        productBlock.querySelectorAll('.dcc-budget-cpe-wrap').forEach(w => {
          w.classList.toggle('visible', visible);
        });
      };

      if (budgets.length) {
        budgets.forEach(budget => this.addBudgetRow(
          budgetsList,
          budget.desc,
          budget.cap,
          budget.weekdays !== undefined ? budget.weekdays : weekdaysDefault,
          budget.spent,
          budget.cpe || '',
          eventOn  // show CPE wrap immediately if restoring with event toggle on
        ));
      } else {
        this.addBudgetRow(budgetsList, '', '', weekdaysDefault, null, '', false);
      }

      // If no CPE data, hide the event cap feature entirely
      if (!cpe || cpe === '') {
        if (eventToggleRow) eventToggleRow.style.display = 'none';
        eventSection.style.display = 'none';
      } else {
        // Has CPE data - show toggle and set up handlers
        if (eventOn) {
          eventToggle.checked = true;
          eventSection.style.display = 'block';
          eventToggle.nextElementSibling.style.background = '#f59e0b';
          // CPE wraps already shown via showCpe flag in addBudgetRow above
        }

        eventToggle.addEventListener('change', () => {
          const on = eventToggle.checked;
          eventSection.style.display = on ? 'block' : 'none';
          eventToggle.nextElementSibling.style.background = on ? '#f59e0b' : '';
          setCpeWrapsVisible(on);
          Storage.autoSave();
        });
      }
    },

    showEmptyState() {
      const container = State.dom.productsList;
      container.innerHTML = `
        <div class="dcc-empty-state">
          <div class="dcc-empty-state-icon">📊</div>
          <div class="dcc-empty-state-title">No data found on page</div>
          <div class="dcc-empty-state-text">Add products manually or refresh when on a client page</div>
        </div>
      `;
    },

    populateFromPage() {
      State.dom.productsList.innerHTML = '';
      State.productCounter = 0;
      State.budgetCounter = 0;
      State.dom.results.style.display = 'none';
      State.dom.results.innerHTML = '';

      const scrapedData = DataScraper.scrapePageData();

      if (scrapedData.length > 0) {
        scrapedData.forEach(productData => this.addProduct(productData));
      } else {
        this.showEmptyState();
        this.addProduct({ name: 'Calls', weekdaysDefault: true });
        this.addProduct({ name: 'Clicks', weekdaysDefault: false });
      }

      Storage.autoSave();
    },

    applyDockState() {
      const calc = State.dom.calculator;
      const btn  = State.dom.dockBtn;
      if (State.docked) {
        calc.classList.add('dcc-docked');
        document.body.style.marginRight = '460px';
        btn.classList.add('docked');
        btn.innerHTML = '&#8701; Undock';
      } else {
        calc.classList.remove('dcc-docked');
        document.body.style.marginRight = '';
        btn.classList.remove('docked');
        btn.innerHTML = '&#8700; Dock';
      }
    },

    setMinimized(isMinimized, save = false) {
      const calcBody = State.dom.calcBody;
      const minBtn = State.dom.minBtn;

      if (isMinimized) {
        calcBody.style.display = 'none';
        minBtn.innerHTML = '&#43;';
      } else {
        calcBody.style.display = '';
        minBtn.innerHTML = '&#8722;';
      }

      if (save) {
        localStorage.setItem(CONFIG.STORAGE_KEY_MINIMIZED, isMinimized ? 'true' : 'false');
      }
    }
  };

  // ============================================================================
  // EVENT HANDLERS
  // ============================================================================

  const EventHandlers = {
    setupDelegatedEvents() {
      // Product list event delegation
      State.dom.productsList.addEventListener('click', (e) => {
        const target = e.target;

        // Remove budget row
        if (target.classList.contains('dcc-remove-btn')) {
          target.closest('.dcc-budget-row')?.remove();
          Storage.autoSave();
        }

        // Remove product
        if (target.classList.contains('dcc-remove-product-btn')) {
          target.closest('.dcc-product-block')?.remove();
          Storage.autoSave();
        }

        // Add budget row
        if (target.classList.contains('dcc-add-brow-btn')) {
          const productBlock = target.closest('.dcc-product-block');
          const budgetsList  = productBlock.querySelector('.dcc-budget-rows-list');
          const eventOn      = productBlock.querySelector('.dcc-event-toggle')?.checked || false;
          UI.addBudgetRow(budgetsList, '', '', false, null, '', eventOn);
        }

        // Spend mode toggle
        if (target.classList.contains('dcc-btn-total')) {
          const productBlock = target.closest('.dcc-product-block');
          UI.setSpendMode(productBlock, 'total');
        }

        if (target.classList.contains('dcc-btn-remaining')) {
          const productBlock = target.closest('.dcc-product-block');
          UI.setSpendMode(productBlock, 'remaining');
        }
      });
    },

    setupHeaderButtons() {
      // Minimize button
      State.dom.minBtn.addEventListener('click', () => {
        const isMinimized = State.dom.calcBody.style.display === 'none';
        UI.setMinimized(!isMinimized, true);
      });

      // Refresh button
      State.dom.refreshBtn.addEventListener('click', function() {
        if (State.frozen) {
          State.dom.freezeBtn.style.outline = '3px solid #fbbf24';
          setTimeout(() => {
            State.dom.freezeBtn.style.outline = '';
          }, 600);
          Utils.showToast('Data is frozen', 'warning');
          return;
        }

        const btn = this;
        btn.classList.add('spinning');
        btn.disabled = true;

        setTimeout(() => {
          UI.populateFromPage();
          btn.classList.remove('spinning');
          btn.disabled = false;
          btn.style.background = 'rgba(74,222,128,0.25)';
          setTimeout(() => {
            btn.style.background = '';
          }, 800);
          Utils.showToast('Data refreshed');
        }, 300);
      });

      // Freeze button
      State.dom.freezeBtn.addEventListener('click', () => {
        if (!State.frozen) {
          State.frozen = true;
          State.frozenState = State.captureState();
          State.dom.freezeBtn.classList.add('frozen');
          State.dom.freezeBtn.innerHTML = '&#128274; Frozen';
          localStorage.setItem(CONFIG.STORAGE_KEY_FROZEN, 'true');
          // Save current state including any results
          Storage.save();
          Utils.showToast('Data frozen');
        } else {
          State.frozen = false;
          State.frozenState = null;
          State.dom.freezeBtn.classList.remove('frozen');
          State.dom.freezeBtn.innerHTML = '&#128275; Freeze';
          localStorage.setItem(CONFIG.STORAGE_KEY_FROZEN, 'false');
          Utils.showToast('Data unfrozen');
        }
      });

      // Dock button — toggle sidebar mode
      State.dom.dockBtn.addEventListener('click', () => {
        State.docked = !State.docked;
        UI.applyDockState();
        localStorage.setItem(CONFIG.STORAGE_KEY_DOCKED, State.docked ? 'true' : 'false');
        Utils.showToast(State.docked ? 'Docked to sidebar' : 'Undocked');
      });

      // Export button
      State.dom.exportBtn.addEventListener('click', () => {
        Storage.exportConfig();
      });

      // Import button
      State.dom.importBtn.addEventListener('click', () => {
        Storage.importConfig();
      });

      // Add product button
      State.dom.addProductBtn.addEventListener('click', () => {
        UI.addProduct({});
      });

      // Calculate button
      State.dom.runBtn.addEventListener('click', () => {
        Calculator.calculate();
      });
    },

    setupDateInput() {
      State.dom.dateInput.addEventListener('change', () => {
        Storage.autoSave();
        // Re-apply live pacing colours across all products when the date shifts
        document.querySelectorAll('.dcc-product-block').forEach(block => {
          const mode = block.dataset.spendMode || 'total';
          const el = block.querySelector(mode === 'total' ? '.dcc-input-total' : '.dcc-input-remaining');
          if (el) Calculator.applyLivePacingToInput(el, block);
        });
      });
    },

    setupKeyboardShortcuts() {
      document.addEventListener('keydown', (e) => {
        // Ignore if typing in an input
        if (e.target.matches('input, textarea')) return;

        // Ctrl/Cmd + Enter = Calculate
        if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
          e.preventDefault();
          State.dom.runBtn.click();
        }

        // Ctrl/Cmd + K = Add product
        if ((e.ctrlKey || e.metaKey) && e.key === 'k') {
          e.preventDefault();
          State.dom.addProductBtn.click();
        }

        // Ctrl/Cmd + M = Toggle minimize
        if ((e.ctrlKey || e.metaKey) && e.key === 'm') {
          e.preventDefault();
          State.dom.minBtn.click();
        }

        // Ctrl/Cmd + E = Export
        if ((e.ctrlKey || e.metaKey) && e.key === 'e') {
          e.preventDefault();
          State.dom.exportBtn.click();
        }
      });
    }
  };

  // ============================================================================
  // HASH CHANGE DETECTION
  // ============================================================================

  let hashListenerAdded = false;

  function setupHashChangeListener() {
    if (hashListenerAdded) return; // Only register once across recreations
    hashListenerAdded = true;
    let settleTimeout = null;
    let fallbackTimeout = null;
    let activeObserver = null;

    // Listen for hash changes (switching between #melonmax and #advertising)
    window.addEventListener('hashchange', () => {
      Utils.log('Hash changed to:', window.location.hash);

      // Tear down any in-flight refresh from a prior hash change so rapid
      // switches (#advertising → #melonmax → #advertising) don't race.
      if (settleTimeout) { clearTimeout(settleTimeout); settleTimeout = null; }
      if (fallbackTimeout) { clearTimeout(fallbackTimeout); fallbackTimeout = null; }
      if (activeObserver) { activeObserver.disconnect(); activeObserver = null; }

      // Show/hide calculator based on hash
      if (!shouldShowCalculator()) {
        Utils.log('Hash not allowed - hiding calculator');
        if (State.dom.calculator) State.dom.calculator.style.display = 'none';
        return;
      }
      Utils.log('Hash allowed - showing calculator');
      if (State.dom.calculator) State.dom.calculator.style.display = '';

      // Don't auto-refresh if frozen, or if we're not on a dashboard page
      if (State.frozen || !isDashboardPage()) return;

      Utils.log('Dashboard hash changed - waiting for content to settle');

      // Clear stale results immediately so the user doesn't see old data
      State.dom.results.style.display = 'none';
      State.dom.results.innerHTML = '';

      // Single-fire refresh — guarded so the observer and fallback don't both run
      let refreshed = false;
      const refresh = (reason) => {
        if (refreshed) return;
        refreshed = true;
        if (activeObserver) { activeObserver.disconnect(); activeObserver = null; }
        if (settleTimeout) { clearTimeout(settleTimeout); settleTimeout = null; }
        if (fallbackTimeout) { clearTimeout(fallbackTimeout); fallbackTimeout = null; }
        Utils.log('Refreshing calculator after hash change:', reason);
        UI.populateFromPage();
      };

      // Watch the content area for mutations, debounce 500ms of quiet
      activeObserver = new MutationObserver(() => {
        if (settleTimeout) clearTimeout(settleTimeout);
        settleTimeout = setTimeout(() => refresh('content settled'), 500);
      });

      const mainContent = document.querySelector('.main-content, main, #main, .content') || document.body;
      activeObserver.observe(mainContent, { childList: true, subtree: true });

      // Fallback: always refresh within 1.5s even if no mutations are observed
      fallbackTimeout = setTimeout(() => refresh('fallback timeout'), 1500);
    });
  }

  // ============================================================================
  // SPA NAVIGATION DETECTION
  // ============================================================================

  function setupNavigationListener() {
    let lastPath = window.location.pathname;
    let lastHash = window.location.hash;

    function onUrlChange() {
      const newPath = window.location.pathname;
      const newHash = window.location.hash;

      const pathChanged = newPath !== lastPath;
      const hashChanged = newHash !== lastHash;

      if (!pathChanged && !hashChanged) return; // Truly nothing changed
      lastPath = newPath;
      lastHash = newHash;

      // ── Hash-only change via pushState/replaceState ────────────────────────
      // The native hashchange event only fires for direct location.hash
      // assignments — NOT for pushState. MelonLocal uses pushState to switch
      // tabs, so we handle visibility here when only the fragment changed.
      if (!pathChanged) {
        Utils.log('SPA hash-only change (pushState):', newHash);
        const calc = document.getElementById('daily-cap-calculator');
        if (!calc) return;

        const visible = shouldShowCalculator();
        calc.style.display = visible ? '' : 'none';

        if (visible && !State.frozen && isDashboardPage()) {
          Utils.log('Refreshing data after pushState tab switch');
          State.dom.results.style.display = 'none';
          State.dom.results.innerHTML = '';

          let hnRefreshed = false;
          let hnTimer = null;
          let hnObs = null;

          const doHnRefresh = (reason) => {
            if (hnRefreshed) return;
            hnRefreshed = true;
            if (hnObs) { hnObs.disconnect(); hnObs = null; }
            if (hnTimer) { clearTimeout(hnTimer); hnTimer = null; }
            Utils.log('Populating after pushState hash change:', reason);
            UI.populateFromPage();
          };

          hnObs = new MutationObserver(() => {
            if (hnTimer) clearTimeout(hnTimer);
            hnTimer = setTimeout(() => doHnRefresh('content settled'), 500);
          });
          const hnRoot = document.querySelector('.main-content, main, #main, .content') || document.body;
          hnObs.observe(hnRoot, { childList: true, subtree: true });
          setTimeout(() => doHnRefresh('fallback'), 1500);
        }
        return;
      }

      // ── Path changed (navigating to a different agent / page) ─────────────
      Utils.log('SPA path change detected:', newPath);

      const calc = document.getElementById('daily-cap-calculator');

      if (!isAllowedPage()) {
        // Navigated away from an allowed page — hide the calculator
        if (calc) calc.style.display = 'none';
        return;
      }

      // Navigated to an allowed page
      if (!calc) {
        // Calculator was never created (user started on a non-allowed page) — create it now
        Utils.log('Calculator missing after SPA nav — creating');
        window[CONFIG.INIT_FLAG] = false;
        setTimeout(smartInit, 300);
        return;
      }

      // Calculator already exists — show/hide per hash and refresh data
      const visible = shouldShowCalculator();
      calc.style.display = visible ? '' : 'none';

      if (!visible || State.frozen) return;

      Utils.log('Refreshing calculator data after SPA navigation');
      State.dom.results.style.display = 'none';
      State.dom.results.innerHTML = '';

      // Single-fire refresh with MutationObserver + fallback
      let navRefreshed = false;
      let navSettleTimer = null;
      let navObs = null;

      const doNavRefresh = (reason) => {
        if (navRefreshed) return;
        navRefreshed = true;
        if (navObs) { navObs.disconnect(); navObs = null; }
        if (navSettleTimer) { clearTimeout(navSettleTimer); navSettleTimer = null; }
        Utils.log('Populating after SPA navigation:', reason);
        UI.populateFromPage();
      };

      navObs = new MutationObserver(() => {
        if (navSettleTimer) clearTimeout(navSettleTimer);
        navSettleTimer = setTimeout(() => doNavRefresh('content settled'), 500);
      });

      const root = document.querySelector('.main-content, main, #main, .content') || document.body;
      navObs.observe(root, { childList: true, subtree: true });

      // Fallback: always refresh within 1.5 s even if no mutations are observed
      setTimeout(() => doNavRefresh('fallback'), 1500);
    }

    // Intercept SPA pushState navigation (clicking links, programmatic nav)
    const origPush = history.pushState.bind(history);
    history.pushState = function (...args) {
      origPush(...args);
      setTimeout(onUrlChange, 0);
    };

    // Intercept replaceState (some SPAs use this instead of pushState)
    const origReplace = history.replaceState.bind(history);
    history.replaceState = function (...args) {
      origReplace(...args);
      setTimeout(onUrlChange, 0);
    };

    // Handle browser Back/Forward buttons
    window.addEventListener('popstate', onUrlChange);
  }

  // ============================================================================
  // INITIALIZATION
  // ============================================================================

  function createCalculator() {
    // Remove any existing instance
    ['daily-cap-calculator', 'dcc-style'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.remove();
    });

    // Add styles
    const style = document.createElement('style');
    style.id = 'dcc-style';
    style.textContent = CSS;
    document.head.appendChild(style);

    // Create HTML
    const calculatorHTML = `
      <div id="daily-cap-calculator" style="position:fixed;bottom:24px;right:24px;width:460px;background:#fff;border-radius:12px;box-shadow:0 8px 32px rgba(0,0,0,0.18);z-index:99999;overflow:hidden;max-height:92vh;display:flex;flex-direction:column;">

        <!-- Header -->
        <div style="background:#1a5c38;padding:14px 18px;display:flex;align-items:center;justify-content:space-between;flex-shrink:0;">
          <div>
            <div style="color:#fff;font-size:15px;font-weight:700;">Daily Cap Calculator</div>
            <div style="color:#a8d5b5;font-size:11px;margin-top:2px;">Pace budgets evenly through end of month</div>
            <div class="dcc-kbd-hint">
              <span class="dcc-kbd">⌘/Ctrl</span><span class="dcc-kbd">Enter</span> Calculate
              <span class="dcc-kbd">⌘/Ctrl</span><span class="dcc-kbd">K</span> Add
            </div>
          </div>
          <div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;">
            <button id="calc-export-btn" class="dcc-header-btn" title="Export configuration (⌘/Ctrl+E)">
              <span>⬇</span> Export
            </button>
            <button id="calc-import-btn" class="dcc-header-btn" title="Import configuration">
              <span>⬆</span> Import
            </button>
            <button id="calc-refresh-btn" class="dcc-header-btn" title="Re-pull data from page">
              <span class="dcc-refresh-icon" style="font-size:13px;">&#x21bb;</span> Refresh
            </button>
            <button id="calc-freeze-btn" class="dcc-header-btn dcc-freeze-btn" title="Freeze data so Refresh won't overwrite it">
              &#128275; Freeze
            </button>
            <button id="calc-dock-btn" class="dcc-header-btn dcc-dock-btn" title="Dock calculator to the right sidebar (sets page margin)">
              &#8700; Dock
            </button>
            <button id="calc-minimize-btn" class="dcc-header-btn" title="Minimize (⌘/Ctrl+M)" style="padding:2px 6px;">
              &#8722;
            </button>
          </div>
        </div>

        <!-- Body -->
        <div id="calc-body" style="padding:16px 18px 18px;overflow-y:auto;flex:1;">

          <!-- Date Input -->
          <div style="margin-bottom:14px;">
            <label class="dcc-label">Today's Date</label>
            <input id="calc-date" type="date" class="dcc-input">
          </div>

          <!-- Products List -->
          <div id="calc-products-list"></div>

          <!-- Add Product Button -->
          <button id="calc-add-product-btn" style="width:100%;padding:8px;background:#fff;border:1.5px dashed #4ade80;border-radius:8px;color:#16a34a;font-size:13px;font-weight:600;cursor:pointer;margin-bottom:12px;">
            + Add Product <span style="font-size:11px;color:#86efac;margin-left:4px;">(⌘/Ctrl+K)</span>
          </button>

          <!-- Calculate Button -->
          <button id="calc-run-btn" style="width:100%;padding:10px;background:#4ade80;color:#14532d;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer;transition:background 0.2s;">
            Calculate Daily Caps <span style="font-size:11px;opacity:0.8;">(⌘/Ctrl+Enter)</span>
          </button>

          <!-- Results -->
          <div id="calc-results" style="display:none;margin-top:14px;"></div>

        </div>
      </div>
    `;

    document.body.insertAdjacentHTML('beforeend', calculatorHTML);

    // Cache DOM references
    State.dom = {
      calculator: document.getElementById('daily-cap-calculator'),
      calcBody: document.getElementById('calc-body'),
      dateInput: document.getElementById('calc-date'),
      productsList: document.getElementById('calc-products-list'),
      results: document.getElementById('calc-results'),
      minBtn: document.getElementById('calc-minimize-btn'),
      refreshBtn: document.getElementById('calc-refresh-btn'),
      freezeBtn: document.getElementById('calc-freeze-btn'),
      dockBtn:   document.getElementById('calc-dock-btn'),
      exportBtn: document.getElementById('calc-export-btn'),
      importBtn: document.getElementById('calc-import-btn'),
      addProductBtn: document.getElementById('calc-add-product-btn'),
      runBtn: document.getElementById('calc-run-btn')
    };

    // Set today's date
    State.dom.dateInput.value = Utils.formatDate(new Date());

    // Setup event handlers
    EventHandlers.setupDelegatedEvents();
    EventHandlers.setupHeaderButtons();
    EventHandlers.setupDateInput();
    EventHandlers.setupKeyboardShortcuts();

    // Load minimized state
    const storedMinimized = localStorage.getItem(CONFIG.STORAGE_KEY_MINIMIZED);
    const startMinimized = storedMinimized === null ? true : storedMinimized === 'true';
    UI.setMinimized(startMinimized, storedMinimized === null);

    // Load frozen state FIRST (before any data operations)
    const wasFrozen = Storage.loadFrozenState();
    if (wasFrozen) {
      State.frozen = true;
      State.dom.freezeBtn.classList.add('frozen');
      State.dom.freezeBtn.innerHTML = '&#128274; Frozen';
      Utils.log('Restored frozen state - will not auto-refresh from page');
    }

    // Restore dock state
    if (localStorage.getItem(CONFIG.STORAGE_KEY_DOCKED) === 'true') {
      State.docked = true;
      UI.applyDockState();
    }

    // Determine if we should scrape from page
    const onDashboard = isDashboardPage();
    const onBudgetDetails = isBudgetDetailsPage();
    const currentHash = window.location.hash.toLowerCase();

    Utils.log('Page type - Dashboard:', onDashboard, 'BudgetDetails:', onBudgetDetails, 'Hash:', currentHash);

    // Try to load saved state first
    const savedState = Storage.load();
    if (savedState && savedState.products && savedState.products.length > 0) {
      Utils.log('Loading saved state');
      State.restoreState(savedState);
      if (State.frozen) {
        State.frozenState = State.captureState();
      }

      // If on Dashboard page and NOT frozen, refresh from page to get latest data
      if (onDashboard && !State.frozen) {
        Utils.log('On Dashboard page - refreshing from page to get latest data');
        setTimeout(() => {
          UI.populateFromPage();
        }, 100);
      }
      // If on BudgetDetails page, keep saved state (don't scrape)
      else if (onBudgetDetails) {
        Utils.log('On BudgetDetails page - keeping saved state without scraping');
      }
    } else if (!State.frozen) {
      // Only scrape from page if NOT frozen and no saved state
      Utils.log('No saved state and not frozen, populating from page');
      UI.populateFromPage();
    } else {
      // Frozen but no saved state - show empty state
      Utils.log('Frozen but no saved state, showing empty state');
      UI.showEmptyState();
    }

    // Setup hash change listener for tab switching
    setupHashChangeListener();

    // Check initial hash and hide if not allowed
    if (!shouldShowCalculator()) {
      Utils.log('Initial hash not allowed - hiding calculator');
      State.dom.calculator.style.display = 'none';
    }
  }

  function smartInit() {
    // Check if we're on an allowed page
    if (!isAllowedPage()) {
      Utils.log('Not on allowed page, skipping initialization');
      return;
    }

    // Prevent duplicate initialization
    if (window[CONFIG.INIT_FLAG]) {
      Utils.log('Already initialized');
      return;
    }

    Utils.log('Initializing Daily Cap Calculator');
    window[CONFIG.INIT_FLAG] = true;

    // If we have saved state, just create the calculator immediately
    const hasSavedState = localStorage.getItem(CONFIG.STORAGE_KEY_STATE) !== null;
    if (hasSavedState) {
      Utils.log('Has saved state, initializing immediately');
      createCalculator();
      return;
    }

    // No saved state - check if page has data to scrape
    if (document.querySelector('.account-section')) {
      Utils.log('Found account sections, initializing immediately');
      createCalculator();
    } else {
      Utils.log('No account sections found, watching for DOM changes');

      const observer = new MutationObserver((mutations, obs) => {
        if (document.querySelector('.account-section')) {
          Utils.log('Account sections appeared, initializing');
          obs.disconnect();
          createCalculator();
        }
      });

      observer.observe(document.body, {
        childList: true,
        subtree: true
      });

      // Fallback: initialize anyway after timeout
      setTimeout(() => {
        observer.disconnect();
        if (!document.getElementById('daily-cap-calculator')) {
          Utils.log('Timeout reached, initializing anyway');
          createCalculator();
        }
      }, CONFIG.INIT_TIMEOUT);
    }
  }

  // ============================================================================
  // START
  // ============================================================================

  // Always set up SPA navigation detection so the calculator responds to
  // pushState/popstate regardless of which page the user lands on first.
  setupNavigationListener();

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
      // Check if we have ANY saved state - if so, init immediately
      const hasSavedState = localStorage.getItem(CONFIG.STORAGE_KEY_STATE) !== null;
      const delay = hasSavedState ? 0 : 1500;
      Utils.log('Init delay:', delay, 'ms (hasSavedState:', hasSavedState, ')');
      setTimeout(smartInit, delay);
    });
  } else {
    // Check if we have ANY saved state - if so, init immediately
    const hasSavedState = localStorage.getItem(CONFIG.STORAGE_KEY_STATE) !== null;
    const delay = hasSavedState ? 0 : 1500;
    Utils.log('Init delay:', delay, 'ms (hasSavedState:', hasSavedState, ')');
    setTimeout(smartInit, delay);
  }

})();