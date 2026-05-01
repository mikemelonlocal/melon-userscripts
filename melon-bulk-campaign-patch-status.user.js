// ==UserScript==
// @name         Melon Local – Bulk Campaign Patch Status
// @namespace    https://thepatch.melonlocal.com/
// @version      3.0.2
// @description  Bulk Active/Inactive campaign patch status tool. Sticky collapsible toolbar, pill look-ahead counts, row processing spinner, undo/rollback buffer, Action FAB, budget-status guard, polled verification.
// @author       You
// @match        https://thepatch.melonlocal.com/Agents/BudgetDetails*
// @grant        none
// @run-at       document-idle
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/melon-bulk-campaign-patch-status.user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/melon-bulk-campaign-patch-status.user.js
// ==/UserScript==

(function () {
  'use strict';

  // ── Constants ──────────────────────────────────────────────────────────────
  const TOOLBAR_ID           = 'melon-bulk-toolbar';
  const PANEL_ID             = 'melon-bulk-panel';
  const FAB_ID               = 'melon-fab';
  const HEADER_CB_ID         = 'melon-header-cb';
  const COUNT_ID             = 'melon-bulk-count';
  const STATUS_SEL_ID        = 'melon-bulk-status';
  const ROW_CB_CLASS         = 'melon-campaign-cb';
  const CB_CELL_CLASS        = 'melon-cb-cell';
  const BADGE_CLASS          = 'melon-state-badge';
  const STYLE_ID             = 'melon-bulk-style';
  const AUDIT_BTN_ID         = 'melon-bulk-audit';
  const AUDIT_BODY_CLASS     = 'melon-audit-mode';
  const COLLAPSE_BTN_ID      = 'melon-bulk-collapse';
  const HEADER_COUNT_ID      = 'melon-bulk-header-count';
  const ROW_PROCESSING_CLASS = 'melon-row-processing';
  const COLLAPSE_STORAGE_KEY = 'melonBulkCollapsed';
  const FAB_STORAGE_KEY      = 'melonBulkFAB';

  // Polled verification: up to VERIFY_RETRIES checks at VERIFY_INTERVAL_MS apart.
  const VERIFY_INTERVAL_MS = 300;
  const VERIFY_RETRIES     = 5;

  // Melon Local brand colors.
  const BRAND_ACTIVE  = '#6C2126';
  const BRAND_SUCCESS = '#47B74F';

  // ── Preset filter definitions ──────────────────────────────────────────────
  const PRESET_DEVICES = [
    { key: 'desktop', label: 'Desktop', terms: ['desktop'] },
    { key: 'mobile',  label: 'Mobile',  terms: ['mobile']  },
  ];
  const PRESET_PRODUCTS = [
    { key: 'auto',    label: 'Auto',    terms: ['auto'] },
    { key: 'home',    label: 'Home',    terms: ['home'] },
    { key: 'renters', label: 'Renters', terms: ['renters'] },
    { key: 'condo',   label: 'Condo',   terms: ['condo'] },
    { key: 'branded', label: 'Branded', terms: ['brand'] },
    { key: 'fire',    label: 'Fire',    terms: ['home', 'renters', 'condo'] },
    { key: 'quote',   label: 'Quote',   terms: ['auto', '004'] },
  ];

  // Mutable active-preset state.
  const activeDevices  = new Set();
  const activeProducts = new Set();

  // ── Debounce ───────────────────────────────────────────────────────────────
  function debounce(fn, ms) {
    let t;
    return function () { clearTimeout(t); t = setTimeout(() => fn.apply(this, arguments), ms); };
  }

  // ── Table / row helpers ────────────────────────────────────────────────────
  function findCampaignTable() {
    for (const t of document.querySelectorAll('table')) {
      if (t.querySelector('[role="switch"]')) return t;
    }
    return null;
  }

  function findNameColumnIndex(table) {
    const headerRow = table.querySelector('thead tr') || table.querySelector('tr');
    if (!headerRow) return 0;
    // Exclude our injected checkbox column so the returned index is in original coords.
    const headers = Array.from(headerRow.children).filter(h => !h.classList.contains(CB_CELL_CLASS));
    const idx = headers.findIndex(h => /name|campaign/i.test((h.innerText || '').trim()));
    return idx >= 0 ? idx : 0;
  }

  function getNameColumnIndex(table) {
    if (table.dataset.melonNameIdx == null)
      table.dataset.melonNameIdx = String(findNameColumnIndex(table));
    return parseInt(table.dataset.melonNameIdx, 10);
  }

  function getRowName(table, row) {
    const nameIdx = getNameColumnIndex(table);
    // +1 because we inject our checkbox column at position 0.
    const cell = row.querySelectorAll('td')[nameIdx + 1] || row.querySelectorAll('td')[nameIdx];
    return (cell?.innerText || '').trim() || '(unknown)';
  }

  function getRowSwitch(row) { return row.querySelector('[role="switch"]'); }
  function isSwitchActive(sw) { return sw?.getAttribute('aria-checked') === 'true'; }

  // ── Styles (injected once) ─────────────────────────────────────────────────
  function injectStyles() {
    if (document.getElementById(STYLE_ID)) return;
    const style = document.createElement('style');
    style.id = STYLE_ID;
    style.textContent = `
      /* ── Toolbar ─────────────────────────────────────────────────────────── */
      #${TOOLBAR_ID} {
        background: #fff;
        border: 2px solid ${BRAND_SUCCESS};
        border-radius: 8px;
        padding: 0;
        margin-bottom: 12px;
        font-family: inherit;
        font-size: 13px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.15);
        position: sticky;
        top: 0;
        z-index: 9999;
      }
      #${TOOLBAR_ID} .melon-bulk-header {
        display: flex; align-items: center; gap: 10px;
        padding: 8px 14px; cursor: pointer; user-select: none;
      }
      #${TOOLBAR_ID} .melon-title { font-weight: 700; color: ${BRAND_ACTIVE}; letter-spacing: 0.2px; }
      #${TOOLBAR_ID} .melon-spacer { flex: 1; }
      #${HEADER_COUNT_ID} {
        background: ${BRAND_SUCCESS}; color: #fff;
        font-weight: 600; padding: 2px 8px; border-radius: 10px; font-size: 11px;
      }
      #${HEADER_COUNT_ID}.empty { background: #eee; color: #888; }
      #${COLLAPSE_BTN_ID} {
        background: transparent; border: 1px solid #ccc; border-radius: 4px;
        padding: 2px 8px; font-size: 11px; cursor: pointer; color: #555;
      }
      #${COLLAPSE_BTN_ID}:hover { background: #f5f5f5; }
      #${TOOLBAR_ID} .melon-bulk-body {
        display: flex; align-items: center; gap: 10px; flex-wrap: wrap;
        padding: 12px 14px; border-top: 1px solid #eee;
      }
      #${TOOLBAR_ID}.collapsed .melon-bulk-body { display: none; }
      #${TOOLBAR_ID} label { font-weight: 600; color: #333; }
      #${TOOLBAR_ID} select, #${TOOLBAR_ID} input[type="text"] {
        padding: 5px 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 13px;
      }
      #${TOOLBAR_ID} select { cursor: pointer; }
      #${TOOLBAR_ID} input[type="text"] { width: 160px; }
      #${TOOLBAR_ID} .sep { color: #ccc; }
      #${TOOLBAR_ID} kbd {
        font-size: 10px; padding: 1px 5px; background: #f5f5f5;
        border: 1px solid #ddd; border-bottom-width: 2px; border-radius: 3px; color: #555;
      }

      /* ── Buttons ────────────────────────────────────────────────────────── */
      .melon-bulk-btn {
        padding: 6px 14px; border-radius: 4px; border: none;
        font-size: 13px; font-weight: 600; cursor: pointer; transition: background 0.2s;
      }
      .melon-bulk-btn[disabled] { opacity: 0.5; cursor: not-allowed; }
      #melon-bulk-apply        { background: ${BRAND_SUCCESS}; color: #fff; }
      #melon-bulk-apply:hover  { background: #379b3d; }
      #${AUDIT_BTN_ID}         { background: ${BRAND_ACTIVE}; color: #fff; }
      #${AUDIT_BTN_ID}:hover   { background: #561a1f; }
      #${AUDIT_BTN_ID}.active  { background: #561a1f; box-shadow: inset 0 0 0 2px #fff, 0 0 0 2px ${BRAND_ACTIVE}; }
      #melon-bulk-selall        { background: #1976d2; color: #fff; }
      #melon-bulk-selall:hover  { background: #1256a0; }
      #melon-bulk-selnone       { background: #757575; color: #fff; }
      #melon-bulk-selnone:hover { background: #424242; }
      #melon-bulk-filter-add    { background: #1976d2; color: #fff; }
      #melon-bulk-filter-add:hover { background: #1256a0; }
      #melon-bulk-filter-sub    { background: #757575; color: #fff; }
      #melon-bulk-filter-sub:hover { background: #424242; }
      #melon-fab-toggle         { background: #e65100; color: #fff; }
      #melon-fab-toggle:hover   { background: #bf360c; }
      #melon-fab-toggle.active  { background: #bf360c; box-shadow: inset 0 0 0 2px #fff, 0 0 0 2px #e65100; }
      .melon-confirm-btn        { background: ${BRAND_SUCCESS}; color: #fff; }
      .melon-confirm-btn:hover  { background: #379b3d; }
      .melon-cancel-btn         { background: #757575; color: #fff; }
      .melon-cancel-btn:hover   { background: #424242; }
      .melon-rollback-btn       { background: #e65100; color: #fff; }
      .melon-rollback-btn:hover { background: #bf360c; }
      #${COUNT_ID} { color: #555; font-style: italic; }

      /* ── Preset pills ───────────────────────────────────────────────────── */
      .melon-preset-row {
        display: flex; align-items: center; gap: 6px; flex-wrap: wrap; width: 100%;
      }
      .melon-preset-label { font-weight: 600; color: #555; font-size: 12px; min-width: 52px; }
      .melon-preset-pill {
        padding: 3px 9px; border-radius: 12px; border: 1.5px solid #ccc;
        background: #f9f9f9; font-size: 12px; font-weight: 600; cursor: pointer;
        color: #444; transition: background 0.15s, border-color 0.15s, color 0.15s;
        display: inline-flex; align-items: center; gap: 3px;
      }
      .melon-preset-pill:hover { border-color: #aaa; background: #f0f0f0; }
      .melon-preset-pill.active { background: ${BRAND_ACTIVE}; border-color: ${BRAND_ACTIVE}; color: #fff; }
      .melon-preset-pill.device.active { background: #1976d2; border-color: #1976d2; color: #fff; }
      .pill-count {
        font-weight: 400; font-size: 10px; opacity: 0.75;
        min-width: 12px; text-align: center;
      }
      .melon-preset-pill.active .pill-count { opacity: 0.9; }
      #melon-preset-count { color: #888; font-size: 12px; font-style: italic; }
      #melon-preset-add   { background: ${BRAND_SUCCESS}; color: #fff; }
      #melon-preset-add:hover { background: #379b3d; }
      #melon-preset-sub   { background: #757575; color: #fff; }
      #melon-preset-sub:hover { background: #424242; }

      /* ── Row checkbox / badge ────────────────────────────────────────────── */
      .melon-row-cb { width: 16px; height: 16px; cursor: pointer; accent-color: ${BRAND_SUCCESS}; }
      .${CB_CELL_CLASS} { text-align: center; white-space: nowrap; }
      .${BADGE_CLASS} {
        display: inline-block; margin-left: 6px; padding: 1px 6px;
        font-size: 10px; font-weight: 700; border-radius: 10px; vertical-align: middle;
      }
      .${BADGE_CLASS}.on  { background: #e9f6ea; color: ${BRAND_SUCCESS}; }
      .${BADGE_CLASS}.off { background: #fafafa; color: #757575; border: 1px solid #e0e0e0; }

      /* ── Audit mode ──────────────────────────────────────────────────────── */
      body.${AUDIT_BODY_CLASS} tbody tr:not(:has(.${ROW_CB_CLASS}:checked)) { display: none; }

      /* ── Row processing spinner ──────────────────────────────────────────── */
      @keyframes melon-spin { to { transform: rotate(360deg); } }
      .${ROW_PROCESSING_CLASS} td:not(.${CB_CELL_CLASS}) { opacity: 0.4; }
      .${ROW_PROCESSING_CLASS} [role="switch"] { visibility: hidden; position: relative; }
      .${ROW_PROCESSING_CLASS} [role="switch"]::before {
        content: '';
        visibility: visible;
        position: absolute;
        top: 0; left: 0; right: 0; bottom: 0; margin: auto;
        width: 18px; height: 18px;
        border: 2.5px solid rgba(108,33,38,0.18);
        border-top-color: ${BRAND_ACTIVE};
        border-radius: 50%;
        animation: melon-spin 0.7s linear infinite;
      }

      /* ── Panel ───────────────────────────────────────────────────────────── */
      #${PANEL_ID} {
        background: #fff; border: 2px solid ${BRAND_ACTIVE}; border-radius: 8px;
        padding: 12px 16px; margin-bottom: 12px; font-size: 13px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
      }
      #${PANEL_ID}.active  { border-color: ${BRAND_ACTIVE}; }
      #${PANEL_ID}.active  .panel-title { color: ${BRAND_ACTIVE}; }
      #${PANEL_ID}.success { border-color: ${BRAND_SUCCESS}; }
      #${PANEL_ID}.success .panel-title { color: ${BRAND_SUCCESS}; }
      #${PANEL_ID}.error   { border-color: #e53935; }
      #${PANEL_ID} .panel-title { font-weight: 700; margin-bottom: 8px; }
      #${PANEL_ID} .panel-list {
        max-height: 160px; overflow-y: auto; background: #fafafa;
        border: 1px solid #eee; border-radius: 4px;
        padding: 6px 10px; margin: 8px 0; font-family: monospace; font-size: 12px;
      }
      #${PANEL_ID} .panel-actions { display: flex; gap: 8px; align-items: center; }
      #${PANEL_ID} progress {
        width: 240px; height: 14px; appearance: none; -webkit-appearance: none;
        border: none; border-radius: 7px; overflow: hidden; background: #f1e5e6;
      }
      #${PANEL_ID} progress::-webkit-progress-bar { background: #f1e5e6; border-radius: 7px; }
      #${PANEL_ID} progress::-webkit-progress-value {
        background: linear-gradient(90deg, ${BRAND_ACTIVE} 0%, #8b2a30 100%);
        border-radius: 7px; transition: width 120ms ease-out;
      }
      #${PANEL_ID} progress::-moz-progress-bar {
        background: linear-gradient(90deg, ${BRAND_ACTIVE} 0%, #8b2a30 100%);
        border-radius: 7px;
      }
      #${PANEL_ID}.success progress::-webkit-progress-value {
        background: linear-gradient(90deg, ${BRAND_SUCCESS} 0%, #6cc972 100%);
      }
      #${PANEL_ID}.success progress::-moz-progress-bar {
        background: linear-gradient(90deg, ${BRAND_SUCCESS} 0%, #6cc972 100%);
      }

      /* ── Action FAB ──────────────────────────────────────────────────────── */
      #${FAB_ID} {
        position: fixed; bottom: 24px; right: 24px; z-index: 10001;
        display: flex; flex-direction: column; align-items: flex-end; gap: 8px;
        font-family: inherit;
      }
      #melon-fab-count {
        background: ${BRAND_SUCCESS}; color: #fff;
        border-radius: 12px; padding: 3px 12px;
        font-size: 12px; font-weight: 600;
      }
      #melon-fab-count.empty { background: #aaa; }
      #melon-fab-apply {
        background: ${BRAND_ACTIVE}; color: #fff; border: none; border-radius: 28px;
        padding: 14px 22px; font-size: 14px; font-weight: 700; cursor: pointer;
        box-shadow: 0 4px 16px rgba(108,33,38,0.4);
        display: flex; align-items: center; gap: 10px;
        transition: background 0.2s, box-shadow 0.2s; white-space: nowrap;
        font-family: inherit;
      }
      #melon-fab-apply:hover { background: #561a1f; box-shadow: 0 6px 20px rgba(108,33,38,0.5); }
      #melon-fab-apply[disabled] { opacity: 0.5; cursor: not-allowed; }
      #melon-fab-dismiss {
        background: transparent; border: none; font-size: 11px; color: #888;
        cursor: pointer; text-decoration: underline; padding: 0; font-family: inherit;
      }
    `;
    document.head.appendChild(style);
  }

  // ── Build toolbar (idempotent) ─────────────────────────────────────────────
  function ensureToolbar(table) {
    if (document.getElementById(TOOLBAR_ID)) return;

    const toolbar     = document.createElement('div');
    toolbar.id        = TOOLBAR_ID;
    toolbar.tabIndex  = -1;
    const startCollapsed = localStorage.getItem(COLLAPSE_STORAGE_KEY) !== 'open';
    if (startCollapsed) toolbar.classList.add('collapsed');

    const mac = navigator.platform.includes('Mac');
    toolbar.innerHTML = `
      <div class="melon-bulk-header" title="Click to expand / collapse">
        <span class="melon-title">🍈 Bulk Campaign Patch</span>
        <span id="${HEADER_COUNT_ID}" class="empty">0 selected</span>
        <span class="melon-spacer"></span>
        <button id="${COLLAPSE_BTN_ID}" type="button">${startCollapsed ? 'Expand ▼' : 'Collapse ▲'}</button>
      </div>
      <div class="melon-bulk-body">
        <label>Bulk Patch Status:</label>
        <select id="${STATUS_SEL_ID}">
          <option value="activate">Set ACTIVE (turn ON)</option>
          <option value="deactivate">Set INACTIVE (turn OFF)</option>
        </select>
        <button class="melon-bulk-btn" id="melon-bulk-apply">✓ Apply to Selected</button>
        <span class="sep">|</span>
        <button class="melon-bulk-btn" id="melon-bulk-selall">☑ Select All</button>
        <button class="melon-bulk-btn" id="melon-bulk-selnone">☐ Deselect All</button>
        <button class="melon-bulk-btn" id="${AUDIT_BTN_ID}" title="Isolate selected rows.">🔍 Audit Selection</button>
        <button class="melon-bulk-btn" id="melon-fab-toggle" title="Pin a floating Apply button to the bottom-right corner.">📌 Pin FAB</button>
        <span class="sep">|</span>
        <label for="melon-bulk-filter">Name contains:</label>
        <input type="text" id="melon-bulk-filter" placeholder="e.g. desktop" autocomplete="off" spellcheck="false" />
        <button class="melon-bulk-btn" id="melon-bulk-filter-add" title="Additively check matching rows.">+ Select matching</button>
        <button class="melon-bulk-btn" id="melon-bulk-filter-sub" title="Uncheck matching rows.">− Deselect matching</button>
        <span id="melon-bulk-filter-count" style="color:#888;"></span>
        <span class="sep">|</span>
        <span id="${COUNT_ID}">0 selected</span>
        <span class="sep">|</span>
        <span style="color:#888;font-size:12px;">
          <kbd>${mac ? '⌘' : 'Ctrl'}</kbd>+<kbd>A</kbd> = Select All
        </span>
        <div class="melon-preset-row" style="margin-top:8px;">
          <span class="melon-preset-label">Device:</span>
          ${PRESET_DEVICES.map(d =>
            `<button class="melon-preset-pill device" data-preset-group="device" data-preset-key="${d.key}">${d.label}<span class="pill-count"></span></button>`
          ).join('')}
          <span class="sep" style="margin:0 4px;">|</span>
          <span class="melon-preset-label">Product:</span>
          ${PRESET_PRODUCTS.map(p =>
            `<button class="melon-preset-pill product" data-preset-group="product" data-preset-key="${p.key}">${p.label}<span class="pill-count"></span></button>`
          ).join('')}
        </div>
        <div class="melon-preset-row" style="margin-top:6px;">
          <button class="melon-bulk-btn" id="melon-preset-add" disabled>+ Select preset</button>
          <button class="melon-bulk-btn" id="melon-preset-sub" disabled>− Deselect preset</button>
          <span id="melon-preset-count"></span>
        </div>
      </div>
    `;

    const wrapper = table.closest('div') || table.parentElement;
    wrapper.insertBefore(toolbar, table);

    // All handlers re-resolve the table so stale references don't bite after SPA re-renders.
    const ct = () => findCampaignTable() || table;

    toolbar.querySelector('.melon-bulk-header').addEventListener('click', (e) => {
      if (e.target.closest('button,input,select,a') && e.target.id !== COLLAPSE_BTN_ID) return;
      toggleCollapsed();
    });

    toolbar.querySelector('#melon-bulk-selall').addEventListener('click', selectAll);
    toolbar.querySelector('#melon-bulk-selnone').addEventListener('click', selectNone);
    toolbar.querySelector('#melon-bulk-apply').addEventListener('click', () => onApplyClicked(ct()));
    toolbar.querySelector(`#${AUDIT_BTN_ID}`).addEventListener('click', toggleAuditMode);
    toolbar.querySelector('#melon-fab-toggle').addEventListener('click', () => toggleFAB(ct()));

    toolbar.querySelectorAll('.melon-preset-pill').forEach(pill => {
      pill.addEventListener('click', () => {
        const { presetGroup: group, presetKey: key } = pill.dataset;
        const set = group === 'device' ? activeDevices : activeProducts;
        set.has(key) ? (set.delete(key), pill.classList.remove('active'))
                     : (set.add(key),    pill.classList.add('active'));
        updatePresetCount(ct());
      });
    });

    toolbar.querySelector('#melon-preset-add').addEventListener('click', () => applyPreset(ct(), true));
    toolbar.querySelector('#melon-preset-sub').addEventListener('click', () => applyPreset(ct(), false));

    const filterInput = toolbar.querySelector('#melon-bulk-filter');
    const refreshFC   = () => updateFilterMatchCount(ct(), filterInput.value);
    filterInput.addEventListener('input', refreshFC);
    toolbar.querySelector('#melon-bulk-filter-add').addEventListener('click', () => {
      filterMatchingRows(ct(), filterInput.value, true); refreshFC();
    });
    toolbar.querySelector('#melon-bulk-filter-sub').addEventListener('click', () => {
      filterMatchingRows(ct(), filterInput.value, false); refreshFC();
    });
    filterInput.addEventListener('keydown', (e) => {
      if (e.key !== 'Enter') return;
      e.preventDefault();
      filterMatchingRows(ct(), filterInput.value, !e.shiftKey);
      refreshFC();
    });

    // Restore FAB if it was previously pinned.
    if (localStorage.getItem(FAB_STORAGE_KEY) === 'on') {
      ensureFAB(table);
      const btn = toolbar.querySelector('#melon-fab-toggle');
      if (btn) { btn.textContent = '📌 FAB On ✓'; btn.classList.add('active'); }
    }
  }

  // ── Action FAB ─────────────────────────────────────────────────────────────
  function ensureFAB(table) {
    if (document.getElementById(FAB_ID)) return;
    const fab = document.createElement('div');
    fab.id = FAB_ID;
    fab.innerHTML = `
      <span id="melon-fab-count" class="empty">0 selected</span>
      <button id="melon-fab-apply">✓ Apply to Selected</button>
      <button id="melon-fab-dismiss">Unpin FAB</button>
    `;
    document.body.appendChild(fab);
    fab.querySelector('#melon-fab-apply').addEventListener('click', () => {
      onApplyClicked(findCampaignTable() || table);
    });
    fab.querySelector('#melon-fab-dismiss').addEventListener('click', () => disableFAB());
    updateCount();
  }

  function disableFAB() {
    document.getElementById(FAB_ID)?.remove();
    try { localStorage.setItem(FAB_STORAGE_KEY, 'off'); } catch (_) {}
    const btn = document.getElementById('melon-fab-toggle');
    if (btn) { btn.textContent = '📌 Pin FAB'; btn.classList.remove('active'); }
  }

  function toggleFAB(table) {
    if (document.getElementById(FAB_ID)) {
      disableFAB();
    } else {
      ensureFAB(table);
      try { localStorage.setItem(FAB_STORAGE_KEY, 'on'); } catch (_) {}
      const btn = document.getElementById('melon-fab-toggle');
      if (btn) { btn.textContent = '📌 FAB On ✓'; btn.classList.add('active'); }
    }
  }

  // ── Text filter ────────────────────────────────────────────────────────────
  function getMatchingRows(table, substring) {
    const term = (substring || '').trim().toLowerCase();
    if (!term) return [];
    return Array.from(table.querySelectorAll('tbody tr'))
      .filter(r => getRowName(table, r).toLowerCase().includes(term));
  }

  function filterMatchingRows(table, substring, check) {
    getMatchingRows(table, substring).forEach(row => {
      const cb = row.querySelector(`.${ROW_CB_CLASS}`);
      if (cb && cb.checked !== check) cb.checked = check;
    });
    updateCount();
  }

  function updateFilterMatchCount(table, substring) {
    const el = document.getElementById('melon-bulk-filter-count');
    if (!el) return;
    const term = (substring || '').trim();
    el.textContent = term ? `${getMatchingRows(table, term).length} match` : '';
  }

  // ── Preset filter (device × product intersection) ─────────────────────────
  function rowMatchesPresets(table, row, simDevices, simProducts) {
    const name = getRowName(table, row).toLowerCase();
    const deviceOk  = simDevices.size  === 0 || [...simDevices].some(k => {
      const d = PRESET_DEVICES.find(x => x.key === k);
      return d && d.terms.some(t => name.includes(t));
    });
    const productOk = simProducts.size === 0 || [...simProducts].some(k => {
      const p = PRESET_PRODUCTS.find(x => x.key === k);
      return p && p.terms.some(t => name.includes(t));
    });
    return deviceOk && productOk;
  }

  function getPresetMatchingRows(table) {
    if (activeDevices.size === 0 && activeProducts.size === 0) return [];
    return Array.from(table.querySelectorAll('tbody tr'))
      .filter(r => rowMatchesPresets(table, r, activeDevices, activeProducts));
  }

  // Returns how many rows would match if `key` (in `group`) were added to the
  // current active set — used to populate the look-ahead count badge on each pill.
  function getPillHypotheticalCount(table, group, key) {
    const simDevices  = new Set(activeDevices);
    const simProducts = new Set(activeProducts);
    if (group === 'device') simDevices.add(key);
    else                    simProducts.add(key);
    return Array.from(table.querySelectorAll('tbody tr'))
      .filter(r => rowMatchesPresets(table, r, simDevices, simProducts)).length;
  }

  function updatePresetCount(table) {
    const countEl   = document.getElementById('melon-preset-count');
    const addBtn    = document.getElementById('melon-preset-add');
    const subBtn    = document.getElementById('melon-preset-sub');
    const hasFilter = activeDevices.size > 0 || activeProducts.size > 0;
    const n         = hasFilter ? getPresetMatchingRows(table).length : 0;
    if (countEl) countEl.textContent = hasFilter ? `${n} match` : '';
    if (addBtn)  addBtn.disabled     = !hasFilter;
    if (subBtn)  subBtn.disabled     = !hasFilter;

    // Update look-ahead count badge on every pill.
    document.querySelectorAll('.melon-preset-pill').forEach(pill => {
      const badge = pill.querySelector('.pill-count');
      if (!badge) return;
      const c = getPillHypotheticalCount(table, pill.dataset.presetGroup, pill.dataset.presetKey);
      badge.textContent = c > 0 ? String(c) : '';
    });
  }

  function applyPreset(table, check) {
    getPresetMatchingRows(table).forEach(row => {
      const cb = row.querySelector(`.${ROW_CB_CLASS}`);
      if (cb && cb.checked !== check) cb.checked = check;
    });
    updateCount();
  }

  // ── Row checkboxes / badges (idempotent) ───────────────────────────────────
  function ensureRowCheckboxes(table) {
    const headerRow = table.querySelector('thead tr') || table.querySelector('tr');
    if (headerRow && !headerRow.querySelector(`.${CB_CELL_CLASS}`)) {
      const thCb = document.createElement('th');
      thCb.className = CB_CELL_CLASS;
      thCb.innerHTML = `<input type="checkbox" id="${HEADER_CB_ID}" class="melon-row-cb" title="Select / deselect all">`;
      headerRow.insertBefore(thCb, headerRow.firstElementChild);
    }

    const dataRows = table.querySelector('tbody')
      ? Array.from(table.querySelectorAll('tbody tr'))
      : Array.from(table.querySelectorAll('tr')).slice(1);

    dataRows.forEach((row, i) => {
      if (!row.querySelector(`.${CB_CELL_CLASS}`)) {
        const td = document.createElement('td');
        td.className = CB_CELL_CLASS;
        td.innerHTML = `<input type="checkbox" class="melon-row-cb ${ROW_CB_CLASS}" data-row-index="${i}">`;
        row.insertBefore(td, row.firstElementChild);
      }
      updateRowBadge(row);
    });

    updateCount();
  }

  function updateRowBadge(row) {
    const sw = getRowSwitch(row);
    if (!sw) return;
    const cell = sw.closest('td') || sw.parentElement;
    if (!cell) return;
    let badge = cell.querySelector(`.${BADGE_CLASS}`);
    if (!badge) {
      badge = document.createElement('span');
      badge.className = BADGE_CLASS;
      cell.appendChild(badge);
    }
    const on = isSwitchActive(sw);
    badge.classList.toggle('on', on);
    badge.classList.toggle('off', !on);
    badge.textContent = on ? 'ON' : 'OFF';
  }

  // ── Visual processing state ────────────────────────────────────────────────
  // While isProcessing, replaces the Kendo switch with a BRAND_ACTIVE spinner
  // and dims non-checkbox cells so the user can see which row is in-flight.
  function toggleRowProcessing(row, isProcessing) {
    row.classList.toggle(ROW_PROCESSING_CLASS, isProcessing);
  }

  // ── Selection helpers ──────────────────────────────────────────────────────
  function getRowCheckboxes() {
    return Array.from(document.querySelectorAll(`.${ROW_CB_CLASS}`));
  }

  function getSelectedRows() {
    return getRowCheckboxes().filter(cb => cb.checked).map(cb => cb.closest('tr'));
  }

  function selectAll()  { getRowCheckboxes().forEach(cb => (cb.checked = true));  updateCount(); }
  function selectNone() { getRowCheckboxes().forEach(cb => (cb.checked = false)); updateCount(); }

  function updateCount() {
    const all     = getRowCheckboxes();
    const checked = all.filter(cb => cb.checked).length;

    const countEl = document.getElementById(COUNT_ID);
    if (countEl) countEl.textContent = `${checked} selected`;

    const headerCountEl = document.getElementById(HEADER_COUNT_ID);
    if (headerCountEl) {
      headerCountEl.textContent = `${checked} selected`;
      headerCountEl.classList.toggle('empty', checked === 0);
    }

    const fabCountEl = document.getElementById('melon-fab-count');
    if (fabCountEl) {
      fabCountEl.textContent = `${checked} selected`;
      fabCountEl.classList.toggle('empty', checked === 0);
    }

    const headerCb = document.getElementById(HEADER_CB_ID);
    if (headerCb) {
      headerCb.checked       = checked > 0 && checked === all.length;
      headerCb.indeterminate = checked > 0 && checked < all.length;
    }
  }

  // ── Collapsed state ────────────────────────────────────────────────────────
  function toggleCollapsed() {
    const toolbar = document.getElementById(TOOLBAR_ID);
    if (!toolbar) return;
    const collapsed = toolbar.classList.toggle('collapsed');
    const btn = toolbar.querySelector(`#${COLLAPSE_BTN_ID}`);
    if (btn) btn.textContent = collapsed ? 'Expand ▼' : 'Collapse ▲';
    try { localStorage.setItem(COLLAPSE_STORAGE_KEY, collapsed ? 'closed' : 'open'); } catch (_) {}
  }

  // ── Audit mode ─────────────────────────────────────────────────────────────
  function toggleAuditMode() {
    const btn = document.getElementById(AUDIT_BTN_ID);
    const on  = document.body.classList.toggle(AUDIT_BODY_CLASS);
    if (btn) {
      btn.classList.toggle('active', on);
      btn.textContent = on ? '👁 Show All' : '🔍 Audit Selection';
      btn.title       = on ? 'Restore the full row list.' : 'Isolate selected rows.';
    }
  }

  // ── Panel system ───────────────────────────────────────────────────────────
  function removePanel() {
    document.getElementById(PANEL_ID)?.remove();
  }

  function showPanel(html, klass = '') {
    removePanel();
    const toolbar = document.getElementById(TOOLBAR_ID);
    if (!toolbar) return null;

    if (toolbar.classList.contains('collapsed')) toggleCollapsed();

    const panel = document.createElement('div');
    panel.id = PANEL_ID;
    if (klass) panel.className = klass;
    panel.innerHTML = html;
    toolbar.parentElement.insertBefore(panel, toolbar.nextSibling);

    // Scroll to top so the panel (near the sticky toolbar) is always visible.
    window.scrollTo({ top: 0, behavior: 'smooth' });
    return panel;
  }

  function showConfirmPanel(targets, action) {
    return new Promise(resolve => {
      const list  = targets.map(t => `• ${t.name}  [${t.currentState}]`).join('\n');
      const panel = showPanel(`
        <div class="panel-title">Confirm: set ${targets.length} campaign(s) to ${action}</div>
        <pre class="panel-list">${escapeHtml(list)}</pre>
        <div class="panel-actions">
          <button class="melon-bulk-btn melon-confirm-btn" id="melon-confirm-yes">Confirm</button>
          <button class="melon-bulk-btn melon-cancel-btn"  id="melon-confirm-no">Cancel</button>
        </div>
      `);
      if (!panel) return resolve(false);
      panel.querySelector('#melon-confirm-yes').addEventListener('click', () => { removePanel(); resolve(true); });
      panel.querySelector('#melon-confirm-no').addEventListener('click',  () => { removePanel(); resolve(false); });
    });
  }

  function showProgressPanel(total) {
    const panel = showPanel(`
      <div class="panel-title">Applying changes…</div>
      <progress id="melon-progress" value="0" max="${total}"></progress>
      <span id="melon-progress-text" style="margin-left:10px;">0 / ${total}</span>
    `, 'active');
    return {
      update(done) {
        if (!panel) return;
        const pr = panel.querySelector('#melon-progress');
        const tx = panel.querySelector('#melon-progress-text');
        if (pr) pr.value = done;
        if (tx) tx.textContent = `${done} / ${total}`;
      },
    };
  }

  // stashedState: array of { row, sw, name, wasActive } — null disables the rollback button.
  function showResultPanel(results, stashedState) {
    const failedList   = results.failed.length
      ? `<pre class="panel-list">${escapeHtml(results.failed.map(n => '• ' + n).join('\n'))}</pre>`
      : '';
    const klass        = results.failed.length ? 'error' : 'success';
    const canRollback  = stashedState?.some(s => isSwitchActive(getRowSwitch(s.row)) !== s.wasActive);
    const rollbackHtml = canRollback
      ? `<button class="melon-bulk-btn melon-rollback-btn" id="melon-rollback-btn">↩ Rollback</button>`
      : '';

    const panel = showPanel(`
      <div class="panel-title">${results.failed.length ? 'Done with errors' : 'Done'}</div>
      <div>
        Changed: <b>${results.changed}</b> &nbsp;|&nbsp;
        Already in desired state: <b>${results.alreadyOK}</b> &nbsp;|&nbsp;
        Failed to verify: <b>${results.failed.length}</b>
      </div>
      ${failedList}
      <div class="panel-actions" style="margin-top:8px;">
        <button class="melon-bulk-btn melon-cancel-btn" id="melon-result-dismiss">Dismiss</button>
        ${rollbackHtml}
      </div>
    `, klass);

    if (!panel) return;
    panel.querySelector('#melon-result-dismiss')?.addEventListener('click', removePanel);
    if (canRollback)
      panel.querySelector('#melon-rollback-btn')?.addEventListener('click', () => performRollback(stashedState));
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, c =>
      ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]));
  }

  // ── Rollback ───────────────────────────────────────────────────────────────
  async function performRollback(stashedState) {
    const toRevert = stashedState.filter(s => {
      const sw = getRowSwitch(s.row);
      return sw && isSwitchActive(sw) !== s.wasActive;
    });

    if (toRevert.length === 0) {
      const p = showPanel(`
        <div class="panel-title">Nothing to roll back — rows already match their original state.</div>
        <div class="panel-actions">
          <button class="melon-bulk-btn melon-cancel-btn" id="melon-result-dismiss">Dismiss</button>
        </div>
      `, 'success');
      p?.querySelector('#melon-result-dismiss')?.addEventListener('click', removePanel);
      return;
    }

    const applyBtn = document.getElementById('melon-bulk-apply');
    if (applyBtn) applyBtn.disabled = true;

    const progress = showProgressPanel(toRevert.length);
    const results  = { changed: 0, alreadyOK: 0, failed: [] };

    for (let i = 0; i < toRevert.length; i++) {
      const s  = toRevert[i];
      const sw = getRowSwitch(s.row);
      if (!sw) { results.failed.push(s.name); progress.update(i + 1); continue; }
      toggleRowProcessing(s.row, true);
      sw.click();
      const verified = await verifyStatus(s.row, s.wasActive);
      toggleRowProcessing(s.row, false);
      if (verified) results.changed++;
      else results.failed.push(s.name);
      updateRowBadge(s.row);
      progress.update(i + 1);
    }

    if (applyBtn) applyBtn.disabled = false;
    showResultPanel(results, null); // no recursive rollback
    updateCount();
  }

  // ── Status verification (polled) ───────────────────────────────────────────
  async function verifyStatus(row, targetState, retries = VERIFY_RETRIES) {
    const sw = getRowSwitch(row);
    if (!sw) return false;
    for (let i = 0; i < retries; i++) {
      if (isSwitchActive(sw) === targetState) return true;
      if (i < retries - 1) await sleep(VERIFY_INTERVAL_MS);
    }
    return false;
  }

  // ── Apply ──────────────────────────────────────────────────────────────────
  async function onApplyClicked(table) {
    const rows = getSelectedRows();
    if (rows.length === 0) {
      const p = showPanel(`
        <div class="panel-title">Pick at least one campaign first.</div>
        <div class="panel-actions">
          <button class="melon-bulk-btn melon-cancel-btn" id="melon-result-dismiss">Dismiss</button>
        </div>
      `, 'error');
      p?.querySelector('#melon-result-dismiss')?.addEventListener('click', removePanel);
      return;
    }

    const wantActive = document.getElementById(STATUS_SEL_ID).value === 'activate';
    const action     = wantActive ? 'ACTIVE' : 'INACTIVE';

    // ── Budget-status guard (top-priority safety check) ───────────────────
    if (!wantActive) {
      const budgetStatusText = document.querySelector('.k-dropdownlist .k-input-value-text')?.innerText?.trim();
      if (budgetStatusText?.toLowerCase() === 'active') {
        const allDataRows           = Array.from(table.querySelectorAll('tbody tr'));
        const activeRows            = allDataRows.filter(r => isSwitchActive(getRowSwitch(r)));
        const activeRowsInSelection = activeRows.filter(r => rows.includes(r));

        if (activeRowsInSelection.length === activeRows.length) {
          const p = showPanel(`
            <div class="panel-title">⚠️ Cannot deactivate all campaigns</div>
            <div style="margin-bottom:8px;">
              The budget status is <b>Active</b> — at least one campaign must remain Active.
              Deselect at least one currently-active campaign and try again.
            </div>
            <div style="font-size:12px;color:#555;margin-bottom:8px;">Currently active campaigns in your selection:</div>
            <pre class="panel-list">${escapeHtml(activeRows.map(r => '• ' + getRowName(table, r)).join('\n'))}</pre>
            <div class="panel-actions">
              <button class="melon-bulk-btn melon-cancel-btn" id="melon-result-dismiss">Dismiss</button>
            </div>
          `, 'error');
          p?.querySelector('#melon-result-dismiss')?.addEventListener('click', removePanel);
          return;
        }
      }
    }
    // ── End budget-status guard ───────────────────────────────────────────

    const targets = rows.map(row => {
      const sw = getRowSwitch(row);
      return { row, sw, name: getRowName(table, row), currentState: isSwitchActive(sw) ? 'ON' : 'OFF' };
    }).filter(t => t.sw);

    const ok = await showConfirmPanel(targets, action);
    if (!ok) return;

    // Stash current state before any changes for rollback.
    const stashedState = targets.map(t => ({
      row: t.row, sw: t.sw, name: t.name, wasActive: isSwitchActive(t.sw),
    }));

    const applyBtn = document.getElementById('melon-bulk-apply');
    if (applyBtn) applyBtn.disabled = true;

    const progress = showProgressPanel(targets.length);
    const results  = { changed: 0, alreadyOK: 0, failed: [] };

    for (let i = 0; i < targets.length; i++) {
      const t = targets[i];
      if (isSwitchActive(t.sw) === wantActive) {
        results.alreadyOK++;
      } else {
        toggleRowProcessing(t.row, true);
        t.sw.click();
        const verified = await verifyStatus(t.row, wantActive);
        toggleRowProcessing(t.row, false);
        if (verified) results.changed++;
        else results.failed.push(t.name);
        updateRowBadge(t.row);
      }
      progress.update(i + 1);
    }

    if (applyBtn) applyBtn.disabled = false;
    showResultPanel(results, stashedState);
    updateCount();
  }

  function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

  // ── Main ensure tick ───────────────────────────────────────────────────────
  function ensureUI() {
    const table = findCampaignTable();
    if (!table) return;
    injectStyles();
    ensureToolbar(table);
    ensureRowCheckboxes(table);
    updatePresetCount(table);
  }

  // ── Global event delegation (bound once) ───────────────────────────────────
  document.addEventListener('change', (e) => {
    const tgt = e.target;
    if (!tgt) return;
    if (tgt.id === HEADER_CB_ID) {
      getRowCheckboxes().forEach(cb => (cb.checked = tgt.checked));
      updateCount();
    } else if (tgt.classList?.contains(ROW_CB_CLASS)) {
      updateCount();
    }
  });

  document.addEventListener('keydown', (e) => {
    const toolbar = document.getElementById(TOOLBAR_ID);
    if (!toolbar?.contains(document.activeElement)) return;
    const key = (e.key || '').toLowerCase();
    if ((e.metaKey || e.ctrlKey) && key === 'a') { e.preventDefault(); selectAll(); }
    else if (key === 'escape') removePanel();
  });

  // ── SPA-aware: re-ensure on DOM mutations ──────────────────────────────────
  const debouncedEnsure = debounce(ensureUI, 150);
  new MutationObserver(debouncedEnsure).observe(document.body, { childList: true, subtree: true });

  ensureUI();
  console.log('[Melon Bulk] v3.0.2 ready.');
})();
