// ==UserScript==
// @name         MelonPatch - Searchable & Sortable Budget Tables
// @namespace    https://thepatch.melonlocal.com/
// @version      1.1
// @description  Adds DataTables search, sort, and per-column checkbox filters (SearchPanes) to Legacy Office Budget tables on the Agent Dashboard.
// @author       You
// @match        https://thepatch.melonlocal.com/Agents/Dashboard/*
// @grant        GM_addStyle
// @grant        GM_addElement
// @require      https://code.jquery.com/jquery-3.7.1.min.js
// @require      https://cdn.datatables.net/1.13.7/js/jquery.dataTables.min.js
// @require      https://cdn.datatables.net/select/1.7.0/js/dataTables.select.min.js
// @require      https://cdn.datatables.net/searchpanes/2.2.0/js/dataTables.searchPanes.min.js
// @run-at       document-idle
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/patch-budget-table-search-sort.user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/patch-budget-table-search-sort.user.js
// ==/UserScript==

(function () {
  'use strict';

  // Load DataTables + extension CSS as real <link>s (avoids render-blocking @import / CSP issues)
  [
    'https://cdn.datatables.net/1.13.7/css/jquery.dataTables.min.css',
    'https://cdn.datatables.net/select/1.7.0/css/select.dataTables.min.css',
    'https://cdn.datatables.net/searchpanes/2.2.0/css/searchPanes.dataTables.min.css',
  ].forEach(href => GM_addElement(document.head, 'link', { rel: 'stylesheet', href }));

  GM_addStyle(`
    .melon-table__container .dataTables_wrapper { font-size: 13px; }
    .melon-table__container .dataTables_filter input {
      margin-left: 6px; border: 1px solid #ccc; border-radius: 4px; padding: 3px 6px;
    }
    .melon-table__container .dataTables_length select {
      border: 1px solid #ccc; border-radius: 4px; padding: 2px 4px;
    }
    .melon-table__container .dataTables_info,
    .melon-table__container .dataTables_paginate { margin-top: 6px; }
    /* Keep the checkbox-filter panes compact so they don't dominate the page */
    .melon-table__container .dtsp-searchPanes .dtsp-searchPane { max-width: 240px; }
    .melon-table__container .dtsp-panesContainer { margin-bottom: 8px; }
  `);

  function initTable(table) {
    if (!table || $.fn.DataTable.isDataTable(table)) return;

    // These tables put their header <th>s inside <tbody>; move the first
    // such row into a proper <thead> so DataTables can use it.
    // NOTE: assumes a single header row — revisit if a 2-row header appears.
    const tbody = table.querySelector('tbody');
    if (!tbody) return;

    const headerRow = Array.from(tbody.querySelectorAll('tr'))
      .find(row => row.querySelector('th'));
    if (!headerRow) return;

    let thead = table.querySelector('thead');
    if (!thead) {
      thead = document.createElement('thead');
      table.insertBefore(thead, tbody);
    }
    thead.appendChild(headerRow);

    $(table).DataTable({
      // 'P' renders the SearchPanes checkbox filters above the table.
      dom: 'Plfrtip',
      pageLength: 25,
      order: [],                                       // preserve original row order
      columnDefs: [
        { orderable: false, targets: -1 },             // action-button column: no sort
        { searchPanes: { show: false }, targets: -1 }, // ...and no checkbox pane
      ],
      searchPanes: {
        // Auto-detect categorical columns: only build a pane when the share of
        // unique values is at/below this threshold. Repetitive columns (status,
        // product, type) get checkboxes; high-cardinality ones ($ amounts, names)
        // are skipped automatically. Raise toward 1 to show more panes.
        threshold: 0.6,
        cascadePanes: true,   // filtering one pane narrows the others' options
        viewTotal: true,      // show match counts next to each value
        initCollapsed: true,  // panes start collapsed to stay compact
        layout: 'columns-3',
        dtOpts: { select: { style: 'multi' }, paging: false },
      },
      language: {
        search: 'Search:',
        lengthMenu: 'Show _MENU_ entries',
      }
    });
  }

  function initAllTables() {
    document.querySelectorAll('table.melon-table.budget-details-table')
      .forEach(initTable);
  }

  // The page builds the tables with its own JS. Watch for them to appear,
  // but debounce so we don't initialize a half-built table mid-render.
  let debounce;
  const observer = new MutationObserver(() => {
    const pending = document.querySelectorAll(
      'table.melon-table.budget-details-table:not(.dataTable)'
    );
    if (!pending.length) return;
    clearTimeout(debounce);
    debounce = setTimeout(initAllTables, 150);
  });

  observer.observe(document.body, { childList: true, subtree: true });

  // Handle the case where tables are already present.
  if (document.readyState === 'complete') {
    initAllTables();
  } else {
    window.addEventListener('load', initAllTables);
  }
})();
