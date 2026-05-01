// ==UserScript==
// @name         Melon Local – Bulk Campaign Patch Status
// @namespace    https://thepatch.melonlocal.com/
// @version      2.0.0
// @description  Multiselect toolbar on the Campaigns grid for bulk Active/Inactive patch status changes. SPA-aware, sequential apply with verification.
// @author       You
// @match        https://thepatch.melonlocal.com/Agents/BudgetDetails*
// @grant        none
// @run-at       document-idle
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/melon-bulk-campaign-patch-status.user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/melon-bulk-campaign-patch-status.user.js
// ==/UserScript==

(function () {
  'use strict';

  const TOOLBAR_ID    = 'melon-bulk-toolbar';
  const PANEL_ID      = 'melon-bulk-panel';
  const HEADER_CB_ID  = 'melon-header-cb';
  const COUNT_ID      = 'melon-bulk-count';
  const STATUS_SEL_ID = 'melon-bulk-status';
  const ROW_CB_CLASS  = 'melon-campaign-cb';
  const CB_CELL_CLASS = 'melon-cb-cell';
  const BADGE_CLASS   = 'melon-state-badge';
  const STYLE_ID      = 'melon-bulk-style';
  const CLICK_DELAY_MS = 200;

  // ── Debounce helper ────────────────────────────────────────────────────────
  function debounce(fn, ms) {
    let t;
    return function () {
      clearTimeout(t);
      t = setTimeout(() => fn.apply(this, arguments), ms);
    };
  }

  // ── Find the Campaigns table (heuristic: contains role=switch cells) ───────
  function findCampaignTable() {
    const tables = document.querySelectorAll('table');
    for (const t of tables) {
      if (t.querySelector('[role="switch"]')) return t;
    }
    return null;
  }

  // ── Find the column index of the campaign name by header text ──────────────
  function findNameColumnIndex(table) {
    const headerRow = table.querySelector('thead tr') || table.querySelector('tr');
    if (!headerRow) return 0;
    const headers = Array.from(headerRow.children);
    const idx = headers.findIndex(h => /name|campaign/i.test((h.innerText || '').trim()));
    return idx >= 0 ? idx : 0;
  }

  // Stored on the table element so it survives re-injection without re-querying.
  function getNameColumnIndex(table) {
    if (table.dataset.melonNameIdx == null) {
      table.dataset.melonNameIdx = String(findNameColumnIndex(table));
    }
    return parseInt(table.dataset.melonNameIdx, 10);
  }

  function getRowName(table, row) {
    const nameIdx = getNameColumnIndex(table);
    // +1 because we inject our checkbox column at index 0
    const cell = row.querySelectorAll('td')[nameIdx + 1] || row.querySelectorAll('td')[nameIdx];
    return (cell?.innerText || '').trim() || '(unknown)';
  }

  function getRowSwitch(row) {
    return row.querySelector('[role="switch"]');
  }

  function isSwitchActive(sw) {
    return sw?.getAttribute('aria-checked') === 'true';
  }

  // ── Styles (injected once) ─────────────────────────────────────────────────
  function injectStyles() {
    if (document.getElementById(STYLE_ID)) return;
    const style = document.createElement('style');
    style.id = STYLE_ID;
    style.textContent = `
      #${TOOLBAR_ID} {
        background: #fff;
        border: 2px solid #4caf50;
        border-radius: 8px;
        padding: 12px 16px;
        margin-bottom: 12px;
        display: flex;
        align-items: center;
        gap: 10px;
        flex-wrap: wrap;
        font-family: inherit;
        font-size: 13px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
      }
      #${TOOLBAR_ID} label { font-weight: 600; color: #333; }
      #${TOOLBAR_ID} select {
        padding: 5px 8px;
        border: 1px solid #ccc;
        border-radius: 4px;
        font-size: 13px;
        cursor: pointer;
      }
      .melon-bulk-btn {
        padding: 6px 14px;
        border-radius: 4px;
        border: none;
        font-size: 13px;
        font-weight: 600;
        cursor: pointer;
        transition: background 0.2s;
      }
      .melon-bulk-btn[disabled] { opacity: 0.5; cursor: not-allowed; }
      #melon-bulk-apply        { background: #4caf50; color: #fff; }
      #melon-bulk-apply:hover  { background: #388e3c; }
      #melon-bulk-selall       { background: #1976d2; color: #fff; }
      #melon-bulk-selall:hover { background: #1256a0; }
      #melon-bulk-selnone      { background: #757575; color: #fff; }
      #melon-bulk-selnone:hover{ background: #424242; }
      .melon-confirm-btn       { background: #4caf50; color: #fff; }
      .melon-confirm-btn:hover { background: #388e3c; }
      .melon-cancel-btn        { background: #757575; color: #fff; }
      .melon-cancel-btn:hover  { background: #424242; }
      #${COUNT_ID}             { color: #555; font-style: italic; }
      .melon-row-cb            { width: 16px; height: 16px; cursor: pointer; accent-color: #4caf50; }
      .${CB_CELL_CLASS}        { text-align: center; white-space: nowrap; }
      .${BADGE_CLASS} {
        display: inline-block;
        margin-left: 6px;
        padding: 1px 6px;
        font-size: 10px;
        font-weight: 700;
        border-radius: 10px;
        vertical-align: middle;
      }
      .${BADGE_CLASS}.on  { background: #e8f5e9; color: #2e7d32; }
      .${BADGE_CLASS}.off { background: #fafafa; color: #757575; border: 1px solid #e0e0e0; }
      #${TOOLBAR_ID} .sep { color: #ccc; }
      #${TOOLBAR_ID} kbd {
        font-size: 10px;
        padding: 1px 5px;
        background: #f5f5f5;
        border: 1px solid #ddd;
        border-bottom-width: 2px;
        border-radius: 3px;
        color: #555;
      }
      #${PANEL_ID} {
        background: #fff;
        border: 2px solid #ff9800;
        border-radius: 8px;
        padding: 12px 16px;
        margin-bottom: 12px;
        font-size: 13px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
      }
      #${PANEL_ID}.success { border-color: #4caf50; }
      #${PANEL_ID}.error   { border-color: #e53935; }
      #${PANEL_ID} .panel-title { font-weight: 700; margin-bottom: 8px; }
      #${PANEL_ID} .panel-list {
        max-height: 160px;
        overflow-y: auto;
        background: #fafafa;
        border: 1px solid #eee;
        border-radius: 4px;
        padding: 6px 10px;
        margin: 8px 0;
        font-family: monospace;
        font-size: 12px;
      }
      #${PANEL_ID} .panel-actions { display: flex; gap: 8px; align-items: center; }
      #${PANEL_ID} progress { width: 200px; height: 12px; }
    `;
    document.head.appendChild(style);
  }

  // ── Build the toolbar (idempotent) ─────────────────────────────────────────
  function ensureToolbar(table) {
    if (document.getElementById(TOOLBAR_ID)) return;

    const toolbar = document.createElement('div');
    toolbar.id = TOOLBAR_ID;
    toolbar.tabIndex = -1;
    toolbar.innerHTML = `
      <label>Bulk Patch Status:</label>
      <select id="${STATUS_SEL_ID}">
        <option value="activate">Set ACTIVE (turn ON)</option>
        <option value="deactivate">Set INACTIVE (turn OFF)</option>
      </select>
      <button class="melon-bulk-btn" id="melon-bulk-apply">✓ Apply to Selected</button>
      <span class="sep">|</span>
      <button class="melon-bulk-btn" id="melon-bulk-selall">☑ Select All</button>
      <button class="melon-bulk-btn" id="melon-bulk-selnone">☐ Deselect All</button>
      <span id="${COUNT_ID}">0 selected</span>
      <span class="sep">|</span>
      <span style="color:#888;">Tip: <kbd>${navigator.platform.includes('Mac') ? '⌘' : 'Ctrl'}</kbd>+<kbd>A</kbd> in toolbar = Select All</span>
    `;

    const wrapper = table.closest('div') || table.parentElement;
    wrapper.insertBefore(toolbar, table);

    toolbar.querySelector('#melon-bulk-selall').addEventListener('click', selectAll);
    toolbar.querySelector('#melon-bulk-selnone').addEventListener('click', selectNone);
    toolbar.querySelector('#melon-bulk-apply').addEventListener('click', () => onApplyClicked(table));
  }

  // ── Header + per-row checkboxes (idempotent) ───────────────────────────────
  function ensureRowCheckboxes(table) {
    const headerRow = table.querySelector('thead tr') || table.querySelector('tr');
    if (headerRow && !headerRow.querySelector(`.${CB_CELL_CLASS}`)) {
      const thCb = document.createElement('th');
      thCb.className = CB_CELL_CLASS;
      thCb.innerHTML = `<input type="checkbox" id="${HEADER_CB_ID}" class="melon-row-cb" title="Select / deselect all">`;
      headerRow.insertBefore(thCb, headerRow.firstElementChild);
    }

    const allRows = Array.from(table.querySelectorAll('tr'));
    const dataRows = table.querySelector('tbody')
      ? Array.from(table.querySelectorAll('tbody tr'))
      : allRows.slice(1);

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

  // ── Per-row state badge ────────────────────────────────────────────────────
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

  // ── Selection helpers ──────────────────────────────────────────────────────
  function getRowCheckboxes() {
    return Array.from(document.querySelectorAll(`.${ROW_CB_CLASS}`));
  }

  function getSelectedRows() {
    return getRowCheckboxes().filter(cb => cb.checked).map(cb => cb.closest('tr'));
  }

  function selectAll() {
    getRowCheckboxes().forEach(cb => (cb.checked = true));
    updateCount();
  }

  function selectNone() {
    getRowCheckboxes().forEach(cb => (cb.checked = false));
    updateCount();
  }

  function updateCount() {
    const all = getRowCheckboxes();
    const checked = all.filter(cb => cb.checked).length;
    const countEl = document.getElementById(COUNT_ID);
    if (countEl) countEl.textContent = `${checked} selected`;
    const headerCb = document.getElementById(HEADER_CB_ID);
    if (headerCb) {
      headerCb.checked = checked > 0 && checked === all.length;
      headerCb.indeterminate = checked > 0 && checked < all.length;
    }
  }

  // ── Inline confirm / status panel ──────────────────────────────────────────
  function removePanel() {
    const p = document.getElementById(PANEL_ID);
    if (p) p.remove();
  }

  function showPanel(html, klass = '') {
    removePanel();
    const toolbar = document.getElementById(TOOLBAR_ID);
    if (!toolbar) return null;
    const panel = document.createElement('div');
    panel.id = PANEL_ID;
    if (klass) panel.className = klass;
    panel.innerHTML = html;
    toolbar.parentElement.insertBefore(panel, toolbar.nextSibling);
    return panel;
  }

  function showConfirmPanel(targets, action) {
    return new Promise(resolve => {
      const list = targets.map(t => `• ${t.name}  [${t.currentState}]`).join('\n');
      const panel = showPanel(`
        <div class="panel-title">Confirm: set ${targets.length} campaign(s) to ${action}</div>
        <pre class="panel-list">${escapeHtml(list)}</pre>
        <div class="panel-actions">
          <button class="melon-bulk-btn melon-confirm-btn" id="melon-confirm-yes">Confirm</button>
          <button class="melon-bulk-btn melon-cancel-btn" id="melon-confirm-no">Cancel</button>
        </div>
      `);
      if (!panel) return resolve(false);
      panel.querySelector('#melon-confirm-yes').addEventListener('click', () => { removePanel(); resolve(true); });
      panel.querySelector('#melon-confirm-no').addEventListener('click', () => { removePanel(); resolve(false); });
    });
  }

  function showProgressPanel(total) {
    const panel = showPanel(`
      <div class="panel-title">Applying changes…</div>
      <progress id="melon-progress" value="0" max="${total}"></progress>
      <span id="melon-progress-text" style="margin-left:10px;">0 / ${total}</span>
    `);
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

  function showResultPanel(results) {
    const failedList = results.failed.length
      ? `<pre class="panel-list">${escapeHtml(results.failed.map(n => '• ' + n).join('\n'))}</pre>`
      : '';
    const klass = results.failed.length ? 'error' : 'success';
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
      </div>
    `, klass);
    if (panel) panel.querySelector('#melon-result-dismiss').addEventListener('click', removePanel);
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]));
  }

  // ── Apply ──────────────────────────────────────────────────────────────────
  async function onApplyClicked(table) {
    const rows = getSelectedRows();
    if (rows.length === 0) {
      showPanel(`<div class="panel-title">Pick at least one campaign first.</div>
        <div class="panel-actions"><button class="melon-bulk-btn melon-cancel-btn" id="melon-result-dismiss">Dismiss</button></div>`, 'error');
      const p = document.getElementById(PANEL_ID);
      p?.querySelector('#melon-result-dismiss')?.addEventListener('click', removePanel);
      return;
    }

    const wantActive = document.getElementById(STATUS_SEL_ID).value === 'activate';
    const action = wantActive ? 'ACTIVE' : 'INACTIVE';

    const targets = rows.map(row => {
      const sw = getRowSwitch(row);
      return {
        row,
        sw,
        name: getRowName(table, row),
        currentState: isSwitchActive(sw) ? 'ON' : 'OFF',
      };
    }).filter(t => t.sw);

    const ok = await showConfirmPanel(targets, action);
    if (!ok) return;

    const applyBtn = document.getElementById('melon-bulk-apply');
    if (applyBtn) applyBtn.disabled = true;

    const progress = showProgressPanel(targets.length);
    const results = { changed: 0, alreadyOK: 0, failed: [] };

    for (let i = 0; i < targets.length; i++) {
      const t = targets[i];
      const isActive = isSwitchActive(t.sw);
      if (isActive === wantActive) {
        results.alreadyOK++;
      } else {
        t.sw.click();
        await sleep(CLICK_DELAY_MS);
        if (isSwitchActive(t.sw) === wantActive) {
          results.changed++;
        } else {
          results.failed.push(t.name);
        }
        updateRowBadge(t.row);
      }
      progress.update(i + 1);
    }

    if (applyBtn) applyBtn.disabled = false;
    showResultPanel(results);
    updateCount();
  }

  function sleep(ms) {
    return new Promise(r => setTimeout(r, ms));
  }

  // ── Main ensure tick ───────────────────────────────────────────────────────
  function ensureUI() {
    const table = findCampaignTable();
    if (!table) return;
    injectStyles();
    ensureToolbar(table);
    ensureRowCheckboxes(table);
  }

  // ── Global event delegation (bound once) ───────────────────────────────────
  document.addEventListener('change', (e) => {
    const tgt = e.target;
    if (!tgt) return;
    if (tgt.id === HEADER_CB_ID) {
      getRowCheckboxes().forEach(cb => (cb.checked = tgt.checked));
      updateCount();
    } else if (tgt.classList && tgt.classList.contains(ROW_CB_CLASS)) {
      updateCount();
    }
  });

  document.addEventListener('keydown', (e) => {
    const toolbar = document.getElementById(TOOLBAR_ID);
    if (!toolbar || !toolbar.contains(document.activeElement)) return;
    const key = (e.key || '').toLowerCase();
    if ((e.metaKey || e.ctrlKey) && key === 'a') {
      e.preventDefault();
      selectAll();
    } else if (key === 'escape') {
      removePanel();
    }
  });

  // ── SPA-aware: re-ensure on DOM changes ────────────────────────────────────
  const debouncedEnsure = debounce(ensureUI, 150);
  const observer = new MutationObserver(debouncedEnsure);
  observer.observe(document.body, { childList: true, subtree: true });

  ensureUI();
  console.log('[Melon Bulk] v2.0.0 ready.');
})();
