// ==UserScript==
// @name         CIB Report Quick Filters
// @namespace    https://thepatch.melonlocal.com/
// @version      1.4.0
// @description  Adds a quick-filter panel to the CIB Report table, allowing fast multi-column filtering
// @author       You
// @match        https://thepatch.melonlocal.com/Reports/CIB*
// @grant        none
// @run-at       document-idle
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/cib-report-quick-filters.user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/cib-report-quick-filters.user.js
// ==/UserScript==

(function () {
  'use strict';

  // ─── Config ──────────────────────────────────────────────────────────────────
  const STORAGE_KEY       = 'cib-filters-v1';
  const PRESETS_KEY       = 'cib-filters-presets-v1';
  const COLLAPSE_KEY      = 'cib-filters-collapsed-v1';
  const GROUP_STATE_PREFIX= 'cib-group-collapsed-';
  const DEBOUNCE_MS       = 400;

  // ─── Groups ──────────────────────────────────────────────────────────────────
  const GROUPS = [
    { id: 'agent',      title: 'Agent Information' },
    { id: 'agency',     title: 'Agency Details'    },
    { id: 'financials', title: 'Financials'        },
    { id: 'location',   title: 'Location'          },
  ];

  // ─── Column definitions ──────────────────────────────────────────────────────
  // type: 'string' | 'number' | 'date' | 'select' | 'multiselect'
  // multiSplit: true → treat the field as a comma-separated list and match with 'contains'
  const COLUMNS = [
    // Agent Information
    { group: 'agent',      field: 'Agent.FullName',         label: 'Full Name',            type: 'string' },
    { group: 'agent',      field: 'Agent.CSMName',          label: 'CSM',                  type: 'multiselect' },
    { group: 'agent',      field: 'Agent.PartnerAgent',     label: 'Partner Agent',        type: 'string' },
    { group: 'agent',      field: 'AgentStatus',            label: 'Agent Status',         type: 'select',
      options: [{ value: 'Active', label: 'Active' }, { value: 'Inactive', label: 'Inactive' }] },
    { group: 'agent',      field: 'Agent.StartDate',        label: 'Start Date (from)',    type: 'date', operator: 'gte' },
    { group: 'agent',      field: 'Agent.StartDate',        label: 'Start Date (to)',      type: 'date', operator: 'lte', fieldAlias: 'Agent.StartDate_to' },

    // Agency Details
    { group: 'agency',     field: 'Agent.AgentCompany',     label: 'Company Name',         type: 'string' },
    { group: 'agency',     field: 'Agency.AgencyTypeName',  label: 'Agency Type',          type: 'select',
      options: [{ value: 'Legacy', label: 'Legacy' }, { value: 'MOA', label: 'MOA' }] },
    { group: 'agency',     field: 'Agency.EntityId',        label: 'Entity ID',            type: 'string' },
    { group: 'agency',     field: 'AgencyStatus',           label: 'Agency Status',        type: 'multiselect',
      options: [{ value: 'Active', label: 'Active' }, { value: 'Inactive', label: 'Inactive' }] },
    { group: 'agency',     field: 'CompanyName',            label: 'Agency Website Name',  type: 'string' },
    { group: 'agency',     field: 'Agency.OfficePhone',     label: 'Office Phone',         type: 'string' },
    { group: 'agency',     field: 'PackageList',            label: 'Packages',             type: 'multiselect', multiSplit: true },
    { group: 'agency',     field: 'LicensedStates',         label: 'Licensed States',      type: 'string' },

    // Financials
    { group: 'financials', field: 'GoogleAdSpend',          label: 'Google Ad Spend ≥',    type: 'number', operator: 'gte' },
    { group: 'financials', field: 'CallOnlySpend',          label: 'Call Only Spend ≥',    type: 'number', operator: 'gte' },
    { group: 'financials', field: 'MicrosoftAdSpend',       label: 'Bing Ad Spend ≥',      type: 'number', operator: 'gte' },
    { group: 'financials', field: 'FacebookAdSpend',        label: 'FB Ad Spend ≥',        type: 'number', operator: 'gte' },
    { group: 'financials', field: 'CustomAdSpend',          label: 'Custom Campaign ≥',    type: 'number', operator: 'gte' },

    // Location
    { group: 'location',   field: 'Agency.State',           label: 'State',                type: 'multiselect' },
    { group: 'location',   field: 'Agency.City',            label: 'City',                 type: 'multiselect' },
    { group: 'location',   field: 'Agency.Zip',             label: 'Zip',                  type: 'string' },
  ];

  // ─── Styles ──────────────────────────────────────────────────────────────────
  const STYLES = `
    #cib-filter-panel {
      background: #fff;
      border: 1px solid #d0d7de;
      border-radius: 6px;
      margin: 0 16px 12px;
      font-family: inherit;
      font-size: 13px;
      box-shadow: 0 1px 4px rgba(0,0,0,.08);
    }
    #cib-filter-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      background: #1b4332;
      color: #fff;
      padding: 8px 14px;
      cursor: pointer;
      user-select: none;
      gap: 8px;
      border-radius: 6px 6px 0 0;
    }
    #cib-filter-header span { font-weight: 600; font-size: 13px; flex: 1; }
    #cib-filter-header .cib-badge {
      background: #52b788;
      border-radius: 10px;
      padding: 1px 8px;
      font-size: 11px;
      font-weight: 700;
      display: none;
    }
    #cib-filter-header .cib-toggle { font-size: 11px; opacity: .7; }

    #cib-filter-body { padding: 12px 14px 10px; }

    /* Search */
    #cib-filter-search-wrap { position: relative; margin-bottom: 10px; }
    #cib-filter-search-icon {
      position: absolute; left: 9px; top: 50%; transform: translateY(-50%);
      color: #999; pointer-events: none;
    }
    #cib-filter-search {
      width: 100%; border: 1px solid #ccc; border-radius: 4px;
      padding: 6px 30px; font-size: 13px; box-sizing: border-box; outline: none;
      transition: border-color .15s;
    }
    #cib-filter-search:focus { border-color: #1b4332; }
    #cib-filter-search-clear {
      position: absolute; right: 6px; top: 50%; transform: translateY(-50%);
      background: transparent; border: none; color: #999; cursor: pointer;
      font-size: 16px; line-height: 1; padding: 2px 6px; border-radius: 50%;
      display: none;
    }
    #cib-filter-search-clear:hover { color: #333; background: rgba(0,0,0,.06); }
    #cib-filter-search-wrap.has-value #cib-filter-search-clear { display: block; }

    /* Presets */
    #cib-filter-presets {
      display: flex; gap: 8px; align-items: center; flex-wrap: wrap;
      margin-bottom: 14px; padding-bottom: 10px; border-bottom: 1px solid #eee;
    }
    #cib-filter-presets .cib-presets-label {
      font-size: 11px; font-weight: 700; color: #555;
      text-transform: uppercase; letter-spacing: .4px;
    }
    #cib-preset-select {
      border: 1px solid #ccc; border-radius: 4px; padding: 4px 7px;
      font-size: 13px; background: #fff; min-width: 200px; outline: none;
    }
    #cib-preset-select:focus { border-color: #1b4332; }
    .cib-preset-btn {
      background: transparent; color: #1b4332; border: 1px solid #ccc;
      border-radius: 4px; padding: 4px 12px; font-size: 12px;
      font-weight: 600; cursor: pointer; transition: all .15s;
    }
    .cib-preset-btn:hover { background: #d8f3dc; border-color: #95d5b2; }
    .cib-preset-btn:disabled { opacity: .4; cursor: not-allowed; }
    .cib-preset-btn.danger { color: #b00; }
    .cib-preset-btn.danger:hover { background: #fde2e2; border-color: #f5b6b6; }

    /* Groups */
    .cib-filter-group {
      margin-bottom: 10px;
      border: 1px solid #ececec;
      border-radius: 4px;
    }
    .cib-filter-group:last-of-type { margin-bottom: 4px; }
    .cib-filter-group-title {
      font-size: 11px; font-weight: 700; color: #1b4332;
      text-transform: uppercase; letter-spacing: .6px;
      margin: 0; padding: 8px 10px; background: #f7f9f8;
      border-radius: 4px 4px 0 0;
      display: flex; justify-content: space-between; align-items: center;
      cursor: pointer; user-select: none;
      transition: background .15s;
    }
    .cib-filter-group.is-collapsed .cib-filter-group-title { border-radius: 4px; }
    .cib-filter-group-title:hover { background: #eef3f0; }
    .cib-filter-group-title::after {
      content: '▾'; font-size: 12px; color: #1b4332;
      transition: transform .25s ease;
    }
    .cib-filter-group.is-collapsed .cib-filter-group-title::after { transform: rotate(-90deg); }

    .cib-group-content-wrapper {
      overflow: visible;
      transition: max-height .3s ease, opacity .25s ease, padding .25s ease;
      max-height: 4000px;
      opacity: 1;
      padding: 10px;
    }
    .cib-filter-group.is-collapsed .cib-group-content-wrapper {
      max-height: 0; opacity: 0;
      padding-top: 0; padding-bottom: 0;
      overflow: hidden;
    }

    .cib-filter-group-rows {
      display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
      gap: 10px 12px;
    }

    /* Rows */
    .cib-filter-row { display: flex; flex-direction: column; gap: 3px; min-width: 0; }
    .cib-filter-row label.cib-row-label {
      font-size: 11px; font-weight: 600; color: #555;
      text-transform: uppercase; letter-spacing: .4px;
    }
    .cib-input-wrap { position: relative; display: block; }
    .cib-filter-row input,
    .cib-filter-row select,
    .cib-ms-button {
      border: 1px solid #ccc; border-radius: 4px;
      padding: 4px 26px 4px 7px; font-size: 13px;
      outline: none; width: 100%; box-sizing: border-box;
      transition: border-color .15s;
      background: #fff;
    }
    .cib-filter-row select { padding-right: 7px; }
    .cib-input-wrap.is-date input { padding-right: 32px; }
    .cib-filter-row input:focus,
    .cib-filter-row select:focus,
    .cib-ms-button:focus { border-color: #1b4332; }
    .cib-filter-row input.active,
    .cib-filter-row select.active,
    .cib-ms-button.active { border-color: #2d6a4f; background: #f0fff4; }

    .cib-input-clear {
      display: none; position: absolute; right: 6px; top: 50%;
      transform: translateY(-50%);
      background: transparent; border: none; color: #999;
      cursor: pointer; font-size: 16px; line-height: 1;
      padding: 0 4px; border-radius: 50%;
    }
    .cib-input-clear:hover { color: #333; background: rgba(0,0,0,.06); }
    .cib-input-wrap.has-value .cib-input-clear { display: block; }
    .cib-input-wrap.is-date .cib-input-clear { right: 28px; }

    .cib-filter-row input[type="number"]::-webkit-outer-spin-button,
    .cib-filter-row input[type="number"]::-webkit-inner-spin-button {
      -webkit-appearance: none; margin: 0;
    }
    .cib-filter-row input[type="number"] { -moz-appearance: textfield; }

    /* Multi-select */
    .cib-ms-button {
      text-align: left; cursor: pointer; color: #888; font-family: inherit;
    }
    .cib-ms-button.active { color: #1b4332; font-weight: 600; }
    .cib-ms-button::after {
      content: '▾';
      position: absolute; right: 8px; top: 50%; transform: translateY(-50%);
      color: #999; font-size: 10px; pointer-events: none;
    }
    .cib-input-wrap.is-ms.has-value .cib-ms-button::after { display: none; }

    .cib-ms-popover {
      position: absolute; top: calc(100% + 4px); left: 0;
      z-index: 1000;
      width: 320px; max-width: 90vw;
      background: #fff; border: 1px solid #ccc; border-radius: 6px;
      box-shadow: 0 4px 14px rgba(0,0,0,.15);
      padding: 8px;
      box-sizing: border-box;
    }
    .cib-ms-popover.flip-right { left: auto; right: 0; }
    .cib-ms-popover[hidden] { display: none; }
    .cib-ms-search {
      width: 100%; border: 1px solid #ccc; border-radius: 4px;
      padding: 4px 7px; font-size: 12px; margin-bottom: 6px;
      box-sizing: border-box; outline: none;
    }
    .cib-ms-search:focus { border-color: #1b4332; }
    .cib-ms-actions {
      display: flex; align-items: center; gap: 8px;
      margin-bottom: 4px; padding: 0 2px; font-size: 11px;
    }
    .cib-ms-action {
      background: transparent; border: none; color: #1b4332;
      cursor: pointer; padding: 2px 0; font-size: 11px;
      font-weight: 600; text-decoration: underline;
    }
    .cib-ms-action:hover { color: #2d6a4f; }
    .cib-ms-count { margin-left: auto; color: #888; font-weight: 600; }
    .cib-ms-list {
      max-height: 240px; overflow-y: auto;
      border: 1px solid #eee; border-radius: 4px;
    }
    .cib-ms-item {
      display: flex; align-items: center; gap: 6px;
      padding: 4px 8px; cursor: pointer; font-size: 12px;
      user-select: none;
      line-height: 1.4;
      overflow-wrap: anywhere;
    }
    .cib-ms-item:hover { background: #f0fff4; }
    .cib-ms-item input { margin: 0; flex-shrink: 0; }
    .cib-ms-empty {
      padding: 14px; color: #888; text-align: center;
      font-size: 12px; font-style: italic;
    }

    /* No-match notice */
    #cib-filter-no-match {
      display: none; text-align: center;
      color: #888; font-size: 12px; padding: 18px 8px;
    }

    /* Actions (sticky to viewport bottom while panel scrolls past) */
    #cib-filter-actions {
      display: flex; gap: 8px; align-items: center;
      margin-top: 10px; padding: 10px 0; border-top: 1px solid #eee;
      position: sticky; bottom: 0; background: #fff;
      z-index: 20;
    }
    #cib-filter-apply {
      background: #1b4332; color: #fff; border: none;
      border-radius: 4px; padding: 6px 18px;
      font-size: 13px; font-weight: 600; cursor: pointer;
      transition: background .2s; min-width: 130px;
    }
    #cib-filter-apply:hover { background: #2d6a4f; }
    #cib-filter-apply.applied { background: #52b788; }
    #cib-filter-clear {
      background: transparent; color: #666; border: 1px solid #ccc;
      border-radius: 4px; padding: 5px 14px;
      font-size: 13px; cursor: pointer; transition: all .15s;
    }
    #cib-filter-clear:hover { border-color: #999; color: #333; }
    #cib-filter-status { font-size: 12px; color: #666; margin-left: auto; }
    #cib-filter-status strong { color: #1b4332; }

    /* Chips */
    #cib-filter-chips {
      display: flex; flex-wrap: wrap; gap: 6px;
      margin-top: 10px; padding-top: 10px; border-top: 1px solid #eee;
    }
    #cib-filter-chips:empty { display: none; }
    .cib-chip {
      display: inline-flex; align-items: center; gap: 4px;
      background: #d8f3dc; color: #1b4332; border: 1px solid #95d5b2;
      border-radius: 12px; padding: 2px 4px 2px 10px;
      font-size: 12px; line-height: 1.4;
    }
    .cib-chip strong { font-weight: 700; margin-right: 2px; }
    .cib-chip-x {
      background: transparent; border: 0; color: #1b4332;
      cursor: pointer; font-size: 14px; font-weight: 700;
      line-height: 1; padding: 0 6px; border-radius: 50%;
    }
    .cib-chip-x:hover { background: rgba(0,0,0,.08); }
  `;

  // ─── Helpers ─────────────────────────────────────────────────────────────────
  function getGrid() {
    const el = document.querySelector('[data-role="grid"]');
    if (!el || !window.$) return null;
    return window.$(el).data('kendoGrid');
  }

  function waitForGrid(cb, retries = 40) {
    const g = getGrid();
    if (g) { cb(g); return; }
    if (retries <= 0) { console.warn('[CIB Filters] Kendo grid not found'); return; }
    setTimeout(() => waitForGrid(cb, retries - 1), 250);
  }

  function debounce(fn, ms) {
    let t;
    return function (...args) {
      clearTimeout(t);
      t = setTimeout(() => fn.apply(this, args), ms);
    };
  }

  function readJSON(key, fallback) {
    try { return JSON.parse(localStorage.getItem(key) || JSON.stringify(fallback)); }
    catch { return fallback; }
  }
  function writeJSON(key, value) {
    try { localStorage.setItem(key, JSON.stringify(value)); }
    catch { /* quota / disabled — ignore */ }
  }

  function getNestedValue(obj, path) {
    if (!obj) return undefined;
    const parts = path.split('.');
    let cur = obj;
    for (const p of parts) {
      if (cur == null) return undefined;
      cur = (typeof cur.get === 'function') ? cur.get(p) : cur[p];
    }
    return cur;
  }

  // Returns sorted distinct values for a column, sourced from the grid or from col.options.
  // Reads from _pristineData so active filters don't shrink the option list.
  function getOptionsFor(col) {
    if (col.options) return col.options.map(o => ({ value: o.value, label: o.label || o.value }));
    const grid = getGrid();
    if (!grid) return [];
    const ds = grid.dataSource;
    const data = ds._pristineData && ds._pristineData.length ? ds._pristineData : ds.data();
    const set = new Set();
    data.forEach(item => {
      const v = getNestedValue(item, col.field);
      if (v == null || v === '') return;
      if (col.multiSplit) {
        String(v).split(',').forEach(s => {
          const t = s.trim();
          if (t) set.add(t);
        });
      } else {
        set.add(String(v));
      }
    });
    return [...set]
      .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }))
      .map(v => ({ value: v, label: v }));
  }

  // ─── Build UI ─────────────────────────────────────────────────────────────────
  function buildPanel() {
    const style = document.createElement('style');
    style.textContent = STYLES;
    document.head.appendChild(style);

    const saved = readJSON(STORAGE_KEY, {});

    const panel = document.createElement('div');
    panel.id = 'cib-filter-panel';

    // Header
    const header = document.createElement('div');
    header.id = 'cib-filter-header';
    header.innerHTML = `
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
        <polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/>
      </svg>
      <span>Quick Filters</span>
      <span class="cib-badge" id="cib-active-count"></span>
      <span class="cib-toggle" id="cib-toggle">▼</span>
    `;
    panel.appendChild(header);

    // Body
    const body = document.createElement('div');
    body.id = 'cib-filter-body';

    // Search
    const searchWrap = document.createElement('div');
    searchWrap.id = 'cib-filter-search-wrap';
    searchWrap.innerHTML = `
      <svg id="cib-filter-search-icon" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
      </svg>
      <input id="cib-filter-search" type="text" placeholder="Find a filter…" autocomplete="off" />
      <button id="cib-filter-search-clear" type="button" title="Clear search">×</button>
    `;
    body.appendChild(searchWrap);

    const searchInput = searchWrap.querySelector('#cib-filter-search');
    const searchClear = searchWrap.querySelector('#cib-filter-search-clear');

    // Presets
    const presetsWrap = document.createElement('div');
    presetsWrap.id = 'cib-filter-presets';
    presetsWrap.innerHTML = `
      <span class="cib-presets-label">Preset</span>
      <select id="cib-preset-select"><option value="">— Load preset… —</option></select>
      <button id="cib-preset-save" class="cib-preset-btn" type="button">Save Current</button>
      <button id="cib-preset-delete" class="cib-preset-btn danger" type="button" disabled>Delete</button>
    `;
    body.appendChild(presetsWrap);

    const presetSelect   = presetsWrap.querySelector('#cib-preset-select');
    const presetSaveBtn  = presetsWrap.querySelector('#cib-preset-save');
    const presetDeleteBtn= presetsWrap.querySelector('#cib-preset-delete');

    // Groups container
    const groupsContainer = document.createElement('div');
    groupsContainer.id = 'cib-filter-groups';
    body.appendChild(groupsContainer);

    const inputs = {}; // key → input entry

    GROUPS.forEach(group => {
      const cols = COLUMNS.filter(c => c.group === group.id);
      if (cols.length === 0) return;

      const groupEl = document.createElement('section');
      groupEl.className = 'cib-filter-group';
      groupEl.dataset.groupId = group.id;

      const title = document.createElement('h4');
      title.className = 'cib-filter-group-title';
      title.textContent = group.title;
      groupEl.appendChild(title);

      const contentWrapper = document.createElement('div');
      contentWrapper.className = 'cib-group-content-wrapper';

      const grid = document.createElement('div');
      grid.className = 'cib-filter-group-rows';

      cols.forEach(col => {
        const key = col.fieldAlias || col.field;
        const row = document.createElement('div');
        row.className = 'cib-filter-row';
        row.dataset.label = col.label.toLowerCase();

        const lbl = document.createElement('label');
        lbl.className = 'cib-row-label';
        lbl.textContent = col.label;
        row.appendChild(lbl);

        const entry = buildInput(col, key, saved[key]);
        row.appendChild(entry.wrapEl);
        grid.appendChild(row);

        entry.rowEl = row;
        inputs[key] = entry;
      });

      contentWrapper.appendChild(grid);
      groupEl.appendChild(contentWrapper);
      groupsContainer.appendChild(groupEl);

      // Restore per-group collapse state (no animation on first paint)
      const groupKey = GROUP_STATE_PREFIX + group.id;
      if (localStorage.getItem(groupKey) === 'collapsed') {
        groupEl.classList.add('is-collapsed');
      }

      // Re-enable overflow:visible after expand transition so popovers aren't clipped
      contentWrapper.addEventListener('transitionend', (e) => {
        if (e.propertyName !== 'max-height' || e.target !== contentWrapper) return;
        if (!groupEl.classList.contains('is-collapsed')) {
          contentWrapper.style.maxHeight = 'none';
        }
      });

      title.addEventListener('click', () => {
        const willCollapse = !groupEl.classList.contains('is-collapsed');
        if (willCollapse) {
          // Lock current height so the transition has a value to animate from
          contentWrapper.style.maxHeight = contentWrapper.scrollHeight + 'px';
          // Force reflow
          void contentWrapper.offsetHeight;
          groupEl.classList.add('is-collapsed');
          contentWrapper.style.maxHeight = '';
        } else {
          groupEl.classList.remove('is-collapsed');
          contentWrapper.style.maxHeight = contentWrapper.scrollHeight + 'px';
        }
        try { localStorage.setItem(groupKey, willCollapse ? 'collapsed' : 'expanded'); } catch {}
      });
    });

    // No-match message
    const noMatch = document.createElement('div');
    noMatch.id = 'cib-filter-no-match';
    noMatch.textContent = 'No filters match your search.';
    body.appendChild(noMatch);

    // Action bar
    const actions = document.createElement('div');
    actions.id = 'cib-filter-actions';

    const applyBtn = document.createElement('button');
    applyBtn.id = 'cib-filter-apply';
    applyBtn.textContent = 'Apply Filters';
    applyBtn.addEventListener('click', () => { applyFilters(); flashApplied(); });

    const clearAllBtn = document.createElement('button');
    clearAllBtn.id = 'cib-filter-clear';
    clearAllBtn.textContent = 'Clear All';
    clearAllBtn.addEventListener('click', clearFilters);

    const status = document.createElement('span');
    status.id = 'cib-filter-status';

    actions.appendChild(applyBtn);
    actions.appendChild(clearAllBtn);
    actions.appendChild(status);
    body.appendChild(actions);

    // Chips
    const chips = document.createElement('div');
    chips.id = 'cib-filter-chips';
    body.appendChild(chips);

    panel.appendChild(body);

    // ── Collapse toggle ────────────────────────────────────────────────────────
    let collapsed = localStorage.getItem(COLLAPSE_KEY) === '1';
    const toggleEl = header.querySelector('#cib-toggle');
    function applyCollapsed() {
      body.style.display = collapsed ? 'none' : '';
      toggleEl.textContent = collapsed ? '▶' : '▼';
    }
    applyCollapsed();
    header.addEventListener('click', () => {
      collapsed = !collapsed;
      try { localStorage.setItem(COLLAPSE_KEY, collapsed ? '1' : '0'); } catch {}
      applyCollapsed();
    });

    // ── Search ─────────────────────────────────────────────────────────────────
    function applySearch(query) {
      const q = (query || '').trim().toLowerCase();
      searchWrap.classList.toggle('has-value', q !== '');
      let anyVisible = false;
      groupsContainer.querySelectorAll('.cib-filter-group').forEach(groupEl => {
        let groupVisible = false;
        groupEl.querySelectorAll('.cib-filter-row').forEach(rowEl => {
          const match = !q || rowEl.dataset.label.includes(q);
          rowEl.style.display = match ? '' : 'none';
          if (match) groupVisible = true;
        });
        groupEl.style.display = groupVisible ? '' : 'none';
        if (groupVisible) anyVisible = true;
      });
      noMatch.style.display = (q && !anyVisible) ? 'block' : 'none';
    }
    searchInput.addEventListener('input', () => applySearch(searchInput.value));
    searchClear.addEventListener('click', () => {
      searchInput.value = '';
      applySearch('');
      searchInput.focus();
    });

    // ── Presets ────────────────────────────────────────────────────────────────
    function refreshPresetDropdown() {
      const presets = readJSON(PRESETS_KEY, {});
      while (presetSelect.options.length > 1) presetSelect.remove(1);
      Object.keys(presets).sort((a, b) => a.localeCompare(b)).forEach(name => {
        const opt = document.createElement('option');
        opt.value = name;
        opt.textContent = name;
        presetSelect.appendChild(opt);
      });
    }
    function applyPresetByName(name) {
      const presets = readJSON(PRESETS_KEY, {});
      const values = presets[name];
      if (!values) return;
      Object.entries(inputs).forEach(([key, entry]) => {
        entry.setValue(values[key]);
      });
      applyFilters();
      flashApplied();
    }
    function saveCurrentAsPreset() {
      const name = (prompt('Save filters as preset. Name?') || '').trim();
      if (!name) return;
      const presets = readJSON(PRESETS_KEY, {});
      if (presets[name] && !confirm(`Overwrite preset "${name}"?`)) return;
      const values = {};
      Object.entries(inputs).forEach(([key, entry]) => {
        if (entry.isEmpty()) return;
        values[key] = entry.getValue();
      });
      presets[name] = values;
      writeJSON(PRESETS_KEY, presets);
      refreshPresetDropdown();
      presetSelect.value = name;
      presetDeleteBtn.disabled = false;
    }
    function deleteSelectedPreset() {
      const name = presetSelect.value;
      if (!name) return;
      if (!confirm(`Delete preset "${name}"?`)) return;
      const presets = readJSON(PRESETS_KEY, {});
      delete presets[name];
      writeJSON(PRESETS_KEY, presets);
      refreshPresetDropdown();
      presetSelect.value = '';
      presetDeleteBtn.disabled = true;
    }

    refreshPresetDropdown();
    presetSelect.addEventListener('change', () => {
      const name = presetSelect.value;
      presetDeleteBtn.disabled = !name;
      if (name) applyPresetByName(name);
    });
    presetSaveBtn.addEventListener('click', saveCurrentAsPreset);
    presetDeleteBtn.addEventListener('click', deleteSelectedPreset);

    // ── Status binding ─────────────────────────────────────────────────────────
    let dataBoundBound = false;
    function bindDataBound(grid) {
      if (dataBoundBound) return;
      dataBoundBound = true;
      grid.dataSource.bind('change', () => {
        const filter = grid.dataSource.filter();
        const hasFilter = !!(filter && filter.filters && filter.filters.length);
        if (!hasFilter) { status.innerHTML = ''; return; }
        const n = grid.dataSource.total();
        status.innerHTML = `<strong>${n.toLocaleString()}</strong> result${n !== 1 ? 's' : ''}`;
      });
    }

    // ── Chip rendering ─────────────────────────────────────────────────────────
    function renderChips(activeFilters) {
      chips.innerHTML = '';
      activeFilters.forEach(({ col, entry, displayValue }) => {
        const chip = document.createElement('span');
        chip.className = 'cib-chip';
        chip.innerHTML = `<strong></strong><span></span>`;
        chip.querySelector('strong').textContent = col.label + ':';
        chip.querySelector('span').textContent = ' ' + displayValue;
        const x = document.createElement('button');
        x.className = 'cib-chip-x';
        x.type = 'button';
        x.title = 'Remove filter';
        x.textContent = '×';
        x.addEventListener('click', (e) => {
          e.stopPropagation();
          entry.clear();
          applyFilters();
        });
        chip.appendChild(x);
        chips.appendChild(chip);
      });
    }

    // ── Apply logic ────────────────────────────────────────────────────────────
    function applyFilters() {
      const grid = getGrid();
      if (!grid) {
        status.innerHTML = '<span style="color:#b00">Grid not ready</span>';
        return;
      }
      bindDataBound(grid);

      const filterParts = [];
      const activeForChips = [];
      const toSave = {};

      Object.entries(inputs).forEach(([key, entry]) => {
        if (entry.isEmpty()) return;
        const col = entry.col;
        const value = entry.getValue();
        toSave[key] = value;

        if (col.type === 'multiselect') {
          const op = col.multiSplit ? 'contains' : 'eq';
          const subFilters = value.map(v => ({ field: col.field, operator: op, value: v }));
          if (subFilters.length === 0) return;
          filterParts.push(subFilters.length === 1 ? subFilters[0] : { logic: 'or', filters: subFilters });
          const display = value.length <= 3 ? value.join(', ') : `${value.length} selected`;
          activeForChips.push({ col, entry, displayValue: display });
        } else if (col.type === 'number') {
          const parsed = parseFloat(value);
          if (isNaN(parsed)) return;
          filterParts.push({ field: col.field, operator: col.operator || 'gte', value: parsed });
          activeForChips.push({ col, entry, displayValue: parsed.toLocaleString() });
        } else if (col.type === 'date') {
          const parsed = new Date(value);
          if (isNaN(parsed.getTime())) return;
          if (col.operator === 'lte') parsed.setHours(23, 59, 59, 999);
          filterParts.push({ field: col.field, operator: col.operator, value: parsed });
          activeForChips.push({ col, entry, displayValue: value });
        } else {
          const op = col.operator || (col.type === 'string' ? 'contains' : 'eq');
          filterParts.push({ field: col.field, operator: op, value });
          activeForChips.push({ col, entry, displayValue: value });
        }
      });

      writeJSON(STORAGE_KEY, toSave);

      const badge = document.getElementById('cib-active-count');
      if (filterParts.length > 0) {
        badge.textContent = filterParts.length;
        badge.style.display = '';
      } else {
        badge.style.display = 'none';
      }

      renderChips(activeForChips);

      grid.dataSource.filter(
        filterParts.length > 0
          ? { logic: 'and', filters: filterParts }
          : {}
      );
    }

    const scheduleApply = debounce(applyFilters, DEBOUNCE_MS);

    // ── Applied! feedback ──────────────────────────────────────────────────────
    let appliedTimer;
    function flashApplied() {
      applyBtn.textContent = 'Applied ✓';
      applyBtn.classList.add('applied');
      clearTimeout(appliedTimer);
      appliedTimer = setTimeout(() => {
        applyBtn.textContent = 'Apply Filters';
        applyBtn.classList.remove('applied');
      }, 900);
    }

    // ── Clear All ──────────────────────────────────────────────────────────────
    function clearFilters() {
      Object.values(inputs).forEach(entry => entry.clear());
      const badge = document.getElementById('cib-active-count');
      badge.style.display = 'none';
      status.innerHTML = '';
      chips.innerHTML = '';
      writeJSON(STORAGE_KEY, {});
      presetSelect.value = '';
      presetDeleteBtn.disabled = true;

      const grid = getGrid();
      if (grid) {
        bindDataBound(grid);
        grid.dataSource.filter({});
      }
    }

    // ── Input factory ──────────────────────────────────────────────────────────
    function buildInput(col, key, savedValue) {
      if (col.type === 'multiselect') return buildMultiselectInput(col, key, savedValue);

      const wrap = document.createElement('div');
      wrap.className = 'cib-input-wrap';
      if (col.type === 'date') wrap.classList.add('is-date');

      let inp;
      if (col.type === 'select') {
        inp = document.createElement('select');
        const blank = document.createElement('option');
        blank.value = ''; blank.textContent = '— All —';
        inp.appendChild(blank);
        (col.options || []).forEach(o => {
          const opt = document.createElement('option');
          opt.value = o.value; opt.textContent = o.label;
          inp.appendChild(opt);
        });
      } else if (col.type === 'date') {
        inp = document.createElement('input');
        inp.type = 'date';
      } else if (col.type === 'number') {
        inp = document.createElement('input');
        inp.type = 'number';
        inp.placeholder = '0'; inp.min = '0';
      } else {
        inp = document.createElement('input');
        inp.type = 'text';
        inp.placeholder = 'contains…';
      }

      inp.dataset.key = key;

      if (savedValue != null && savedValue !== '') {
        inp.value = Array.isArray(savedValue) ? savedValue.join(', ') : savedValue;
      }

      let clearBtn = null;
      if (col.type !== 'select') {
        clearBtn = document.createElement('button');
        clearBtn.type = 'button';
        clearBtn.className = 'cib-input-clear';
        clearBtn.textContent = '×';
        clearBtn.tabIndex = -1;
        clearBtn.title = 'Clear';
        clearBtn.addEventListener('click', (e) => {
          e.stopPropagation();
          inp.value = '';
          refresh();
          scheduleApply();
          inp.focus();
        });
      }

      wrap.appendChild(inp);
      if (clearBtn) wrap.appendChild(clearBtn);

      function refresh() {
        const has = inp.value !== '';
        wrap.classList.toggle('has-value', has);
        inp.classList.toggle('active', has);
      }

      const onChange = () => { refresh(); scheduleApply(); };
      inp.addEventListener('input', onChange);
      inp.addEventListener('change', onChange);
      inp.addEventListener('keydown', e => {
        if (e.key === 'Enter') { e.preventDefault(); applyFilters(); flashApplied(); }
      });

      refresh();

      return {
        col, wrapEl: wrap, focusEl: inp,
        getValue: () => inp.value.trim(),
        setValue: (v) => {
          if (Array.isArray(v)) v = v.join(', ');
          inp.value = v == null ? '' : String(v);
          refresh();
        },
        clear: () => { inp.value = ''; refresh(); },
        isEmpty: () => inp.value.trim() === '',
        refresh,
      };
    }

    function buildMultiselectInput(col, key, savedValue) {
      const selected = new Set();
      // Only accept array-shaped saved values. Strings from v1.2.x (when these
      // were `contains` text fields) would re-apply as strict `eq` and zero out
      // the result set, so we drop them and let the user re-select.
      if (Array.isArray(savedValue)) savedValue.forEach(v => selected.add(String(v)));

      const wrap = document.createElement('div');
      wrap.className = 'cib-input-wrap is-ms';

      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'cib-ms-button';

      const popover = document.createElement('div');
      popover.className = 'cib-ms-popover';
      popover.hidden = true;
      popover.addEventListener('click', e => e.stopPropagation());
      popover.addEventListener('mousedown', e => e.stopPropagation());

      const clearBtn = document.createElement('button');
      clearBtn.type = 'button';
      clearBtn.className = 'cib-input-clear';
      clearBtn.textContent = '×';
      clearBtn.tabIndex = -1;
      clearBtn.title = 'Clear all';
      clearBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        selected.clear();
        refresh();
        scheduleApply();
      });

      wrap.appendChild(button);
      wrap.appendChild(clearBtn);
      wrap.appendChild(popover);

      function refresh() {
        const n = selected.size;
        button.textContent = n === 0
          ? '— Any —'
          : (n === 1 ? `1 selected: ${[...selected][0]}` : `${n} selected`);
        button.classList.toggle('active', n > 0);
        wrap.classList.toggle('has-value', n > 0);
      }

      function openPopover() {
        if (!popover.hidden) return;
        renderPopover();
        popover.classList.remove('flip-right');
        popover.hidden = false;
        // If anchored-left popover would overflow the viewport, flip to right-anchor.
        const rect = popover.getBoundingClientRect();
        if (rect.right > window.innerWidth - 8) {
          popover.classList.add('flip-right');
        }
        setTimeout(() => {
          const sb = popover.querySelector('.cib-ms-search');
          if (sb) sb.focus();
        }, 0);
      }
      function closePopover() { popover.hidden = true; }

      function renderPopover() {
        popover.innerHTML = '';
        const options = getOptionsFor(col);
        // Preserve any selected values not currently in options
        selected.forEach(v => {
          if (!options.some(o => o.value === v)) options.push({ value: v, label: v });
        });
        options.sort((a, b) => String(a.label).localeCompare(String(b.label), undefined, { numeric: true, sensitivity: 'base' }));
        if (!window.__cibMsLogged) {
          window.__cibMsLogged = true;
          console.log('[CIB Filters] first multi-select options for', col.field, options.slice(0, 5));
        }

        const search = document.createElement('input');
        search.type = 'text';
        search.className = 'cib-ms-search';
        search.placeholder = options.length > 8 ? 'Filter list…' : '';
        if (options.length <= 8) search.style.display = 'none';
        popover.appendChild(search);

        const actions = document.createElement('div');
        actions.className = 'cib-ms-actions';
        actions.innerHTML = `
          <button type="button" class="cib-ms-action" data-action="all">Select all visible</button>
          <button type="button" class="cib-ms-action" data-action="none">Clear</button>
          <span class="cib-ms-count"></span>
        `;
        popover.appendChild(actions);

        const list = document.createElement('div');
        list.className = 'cib-ms-list';
        popover.appendChild(list);

        let visible = options.slice();

        function renderList(q) {
          const query = (q || '').toLowerCase();
          list.innerHTML = '';
          visible = options.filter(o => !query || o.label.toLowerCase().includes(query));
          if (visible.length === 0) {
            const empty = document.createElement('div');
            empty.className = 'cib-ms-empty';
            empty.textContent = options.length === 0 ? 'No values in data yet' : 'No matches';
            list.appendChild(empty);
            return;
          }
          visible.forEach(o => {
            const item = document.createElement('label');
            item.className = 'cib-ms-item';
            item.title = String(o.label);
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.checked = selected.has(o.value);
            cb.addEventListener('change', () => {
              if (cb.checked) selected.add(o.value);
              else selected.delete(o.value);
              refresh();
              updateCount();
              scheduleApply();
            });
            item.appendChild(cb);
            item.appendChild(document.createTextNode(' ' + String(o.label)));
            list.appendChild(item);
          });
        }
        function updateCount() {
          actions.querySelector('.cib-ms-count').textContent = `${selected.size}/${options.length}`;
        }
        search.addEventListener('input', () => renderList(search.value));
        actions.querySelector('[data-action="all"]').addEventListener('click', () => {
          visible.forEach(o => selected.add(o.value));
          renderList(search.value);
          refresh();
          updateCount();
          scheduleApply();
        });
        actions.querySelector('[data-action="none"]').addEventListener('click', () => {
          selected.clear();
          renderList(search.value);
          refresh();
          updateCount();
          scheduleApply();
        });
        renderList();
        updateCount();
      }

      button.addEventListener('click', (e) => {
        e.stopPropagation();
        if (popover.hidden) openPopover();
        else closePopover();
      });

      refresh();

      return {
        col, wrapEl: wrap, focusEl: button,
        getValue: () => [...selected],
        setValue: (v) => {
          selected.clear();
          if (Array.isArray(v)) v.forEach(x => selected.add(String(x)));
          else if (v != null && v !== '') selected.add(String(v));
          refresh();
        },
        clear: () => { selected.clear(); refresh(); },
        isEmpty: () => selected.size === 0,
        refresh,
        _closePopover: closePopover,
      };
    }

    // Close any open popovers on outside click
    document.addEventListener('click', () => {
      Object.values(inputs).forEach(entry => entry._closePopover && entry._closePopover());
    });

    // Restore saved values on load
    if (Object.keys(saved).some(k => saved[k] != null && saved[k] !== '' && !(Array.isArray(saved[k]) && saved[k].length === 0))) {
      setTimeout(applyFilters, 0);
    }

    return panel;
  }

  // ─── Inject panel into page ───────────────────────────────────────────────────
  function inject() {
    const toolbar = document.querySelector('[role="toolbar"][aria-label="grid toolbar"]');
    if (!toolbar) { setTimeout(inject, 500); return; }

    const container = toolbar.closest('.k-grid') || toolbar.parentElement;
    if (!container || !container.parentNode) { setTimeout(inject, 500); return; }

    try {
      const panel = buildPanel();
      container.before(panel);
    } catch (err) {
      console.error('[CIB Filters] inject failed:', err);
    }
  }

  waitForGrid(() => inject());

})();
