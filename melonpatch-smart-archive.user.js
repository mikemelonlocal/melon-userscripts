// ==UserScript==
// @name         MelonPatch Smart Archive
// @namespace    https://thepatch.melonlocal.com/
// @version      1.1.0
// @description  Adds a "Smart Archive" button next to the Archive button on the Tasks page. Lets you choose to archive Done, Unsuccessful, or both statuses in one click. v1.1.0: pagination-aware, hold-to-confirm, optimistic removal, inline error state.
// @author       MelonLocal
// @match        https://thepatch.melonlocal.com/Agents/Dashboard/*
// @grant        none
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/melonpatch-smart-archive.user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/melonpatch-smart-archive.user.js
// ==/UserScript==

(function () {
  'use strict';

  // ─── Config ────────────────────────────────────────────────────────────────
  const BUTTON_ID         = 'smart-archive-btn';
  const WRAPPER_ID        = 'smart-archive-wrapper';
  const MODAL_ID          = 'smart-archive-modal';
  const HOLD_DURATION_MS  = 1500;
  const BRAND_GREEN       = '#2d6a4f';
  const BRAND_GREEN_DARK  = '#1f4f39';

  // ─── Inject CSS ────────────────────────────────────────────────────────────
  const style = document.createElement('style');
  style.textContent = `
    #${WRAPPER_ID} { order: -4; }

    #${BUTTON_ID} {
      background-color: ${BRAND_GREEN} !important;
      color: #fff !important;
      border-color:  ${BRAND_GREEN} !important;
      font-family: Poppins, sans-serif;
      cursor: pointer;
    }
    #${BUTTON_ID}:hover {
      background-color: ${BRAND_GREEN_DARK} !important;
      border-color: ${BRAND_GREEN_DARK} !important;
    }

    /* ── Modal overlay ── */
    #${MODAL_ID}-overlay {
      display: none;
      position: fixed;
      inset: 0;
      background: rgba(0,0,0,.45);
      z-index: 99998;
      align-items: center;
      justify-content: center;
    }
    #${MODAL_ID}-overlay.sa-open { display: flex; }

    /* ── Modal box ── */
    #${MODAL_ID} {
      background: #fff;
      border-radius: 12px;
      box-shadow: 0 8px 32px rgba(0,0,0,.22);
      padding: 28px 32px 24px;
      width: 420px;
      font-family: Poppins, sans-serif;
      position: relative;
      z-index: 99999;
    }
    #${MODAL_ID} h2 {
      margin: 0 0 6px;
      font-size: 17px;
      font-weight: 600;
      color: #1a1a1a;
      display: flex;
      align-items: center;
      gap: 8px;
    }
    #${MODAL_ID} .sa-subtitle {
      font-size: 12.5px;
      color: #666;
      margin-bottom: 20px;
    }

    /* ── Status checkboxes ── */
    #${MODAL_ID} .sa-options {
      display: flex;
      flex-direction: column;
      gap: 10px;
      margin-bottom: 14px;
    }
    #${MODAL_ID} .sa-option {
      display: flex;
      align-items: center;
      gap: 10px;
      padding: 10px 14px;
      border: 1.5px solid #e0e0e0;
      border-radius: 8px;
      cursor: pointer;
      transition: border-color .15s, background .15s;
      user-select: none;
    }
    #${MODAL_ID} .sa-option:hover {
      border-color: ${BRAND_GREEN};
      background: #f0faf5;
    }
    #${MODAL_ID} .sa-option input[type=checkbox],
    #${MODAL_ID} .sa-scope input[type=checkbox] {
      width: 16px;
      height: 16px;
      accent-color: ${BRAND_GREEN};
      cursor: pointer;
      flex-shrink: 0;
    }
    #${MODAL_ID} .sa-option-label { flex: 1; }
    #${MODAL_ID} .sa-option-label strong {
      display: block;
      font-size: 14px;
      color: #1a1a1a;
    }
    #${MODAL_ID} .sa-option-label span {
      font-size: 11.5px;
      color: #888;
    }
    #${MODAL_ID} .sa-badge {
      font-size: 11px;
      font-weight: 600;
      padding: 2px 8px;
      border-radius: 20px;
      flex-shrink: 0;
    }
    #${MODAL_ID} .sa-badge.done         { background: #d1fae5; color: #065f46; }
    #${MODAL_ID} .sa-badge.unsuccessful { background: #fee2e2; color: #991b1b; }

    /* ── Pagination scope toggle ── */
    #${MODAL_ID} .sa-scope {
      display: none;
      align-items: center;
      gap: 10px;
      padding: 10px 14px;
      border: 1.5px dashed #cdd5d0;
      border-radius: 8px;
      margin-bottom: 16px;
      background: #fafafa;
      user-select: none;
      cursor: pointer;
    }
    #${MODAL_ID} .sa-scope.sa-shown { display: flex; }
    #${MODAL_ID} .sa-scope-label { flex: 1; }
    #${MODAL_ID} .sa-scope-label strong {
      display: block;
      font-size: 13px;
      color: #1a1a1a;
    }
    #${MODAL_ID} .sa-scope-label span {
      font-size: 11.5px;
      color: #888;
    }

    /* ── Preview / status box ── */
    #${MODAL_ID} .sa-preview {
      font-size: 12.5px;
      color: #555;
      background: #f5f5f2;
      border: 1.5px solid transparent;
      border-radius: 7px;
      padding: 10px 12px;
      margin-bottom: 20px;
      min-height: 34px;
      display: flex;
      align-items: center;
      gap: 6px;
      transition: background .15s, color .15s, border-color .15s;
    }
    #${MODAL_ID} .sa-preview.sa-warn {
      color: #b45309;
      background: #fffbeb;
    }
    #${MODAL_ID} .sa-preview.sa-error {
      color: #991b1b;
      background: #fef2f2;
      border-color: #dc2626;
    }

    /* ── Action buttons ── */
    #${MODAL_ID} .sa-actions {
      display: flex;
      gap: 10px;
      justify-content: flex-end;
    }
    #${MODAL_ID} .sa-cancel {
      padding: 7px 18px;
      border: 1.5px solid #d0d0d0;
      border-radius: 8px;
      background: #fff;
      color: #444;
      font-family: Poppins, sans-serif;
      font-size: 13.5px;
      cursor: pointer;
    }
    #${MODAL_ID} .sa-cancel:hover { background: #f5f5f5; }

    /* ── Hold-to-confirm button (progress fill) ── */
    #${MODAL_ID} .sa-confirm {
      position: relative;
      overflow: hidden;
      isolation: isolate;
      padding: 7px 22px;
      min-width: 140px;
      border: none;
      border-radius: 8px;
      background: ${BRAND_GREEN};
      color: #fff;
      font-family: Poppins, sans-serif;
      font-size: 13.5px;
      font-weight: 600;
      cursor: pointer;
      user-select: none;
      -webkit-user-select: none;
      -webkit-touch-callout: none;
    }
    #${MODAL_ID} .sa-confirm .sa-confirm-label {
      position: relative;
      z-index: 1;
      pointer-events: none;
    }
    #${MODAL_ID} .sa-confirm::before {
      content: '';
      position: absolute;
      inset: 0;
      background: ${BRAND_GREEN_DARK};
      transform: translateX(-100%);
      transition: transform 180ms ease-out;
      z-index: 0;
    }
    #${MODAL_ID} .sa-confirm.sa-holding::before {
      transform: translateX(0);
      transition: transform ${HOLD_DURATION_MS}ms linear;
    }
    #${MODAL_ID} .sa-confirm:disabled {
      background: #a0a0a0;
      cursor: not-allowed;
    }
    #${MODAL_ID} .sa-confirm:disabled::before { display: none; }
  `;
  document.head.appendChild(style);

  // ─── Helpers ───────────────────────────────────────────────────────────────
  function taskWord(n) { return n === 1 ? '1 task' : n + ' tasks'; }

  function getGrid() {
    const el = document.getElementById('OpenTasksData');
    return el ? $(el).data('kendoGrid') : null;
  }

  /** Tasks on the currently-loaded page only. */
  function getMatchingTasks(statuses) {
    const grid = getGrid();
    if (!grid) return [];
    return grid.dataSource.data().filter(t => statuses.includes(t.Status));
  }

  /** True when the dataSource has more rows than fit in one page. */
  function hasMultiplePages() {
    const grid = getGrid();
    if (!grid) return false;
    const ds = grid.dataSource;
    const total    = typeof ds.total    === 'function' ? ds.total()    : 0;
    const pageSize = typeof ds.pageSize === 'function' ? ds.pageSize() : 0;
    return pageSize > 0 && total > pageSize;
  }

  /**
   * Fetch every task across all pages, return those whose Status matches.
   * Temporarily widens pageSize to total(); the original pageSize is captured
   * in `originalPaging` and restored after the archive POST resolves.
   */
  let originalPaging = null;
  function fetchAllMatchingTasks(statuses) {
    const grid = getGrid();
    if (!grid) return Promise.resolve([]);
    const ds = grid.dataSource;

    // All data already client-side — no fetch needed.
    if (ds.data().length >= ds.total()) {
      return Promise.resolve(ds.data().filter(t => statuses.includes(t.Status)));
    }

    return new Promise((resolve, reject) => {
      originalPaging = {
        pageSize: ds.pageSize(),
        page: ds.page()
      };

      const onChange = () => {
        ds.unbind('change', onChange);
        ds.unbind('error', onError);
        resolve(ds.data().filter(t => statuses.includes(t.Status)));
      };
      const onError = (e) => {
        ds.unbind('change', onChange);
        ds.unbind('error', onError);
        originalPaging = null;
        reject(e || new Error('dataSource error'));
      };
      ds.bind('change', onChange);
      ds.bind('error', onError);
      ds.pageSize(ds.total());
      ds.page(1);
    });
  }

  /**
   * After an archive completes, return the grid to its pre-fetch paging.
   * Calling pageSize()/page() triggers a single re-read; if there's nothing
   * to restore, fall back to a plain read() so the grid still refreshes.
   */
  function refreshGrid() {
    const grid = getGrid();
    if (!grid) return;
    const ds = grid.dataSource;
    if (originalPaging) {
      const { pageSize, page } = originalPaging;
      originalPaging = null;
      ds.pageSize(pageSize);
      // pageSize() resets to page 1 — restore the user's page if needed
      if (page && page !== ds.page()) ds.page(page);
    } else {
      ds.read();
    }
  }

  function getRVT() {
    const el = document.querySelector('input[name="__RequestVerificationToken"]');
    return el ? el.value : null;
  }

  // ─── Modal state ───────────────────────────────────────────────────────────
  const modalState = {
    holdTimer: null,
    inFlight: false
  };

  // ─── Build modal (once) ────────────────────────────────────────────────────
  function buildModal() {
    if (document.getElementById(MODAL_ID + '-overlay')) return;

    const overlay = document.createElement('div');
    overlay.id = MODAL_ID + '-overlay';
    overlay.innerHTML = `
      <div id="${MODAL_ID}" role="dialog" aria-modal="true" aria-labelledby="sa-title">
        <h2 id="sa-title">🗂 Smart Archive</h2>
        <p class="sa-subtitle">Select which task statuses to archive, then hold Confirm.</p>

        <div class="sa-options">
          <label class="sa-option" id="sa-opt-done">
            <input type="checkbox" id="sa-check-done" value="Done" checked />
            <div class="sa-option-label">
              <strong>Done</strong>
              <span>Tasks marked as completed</span>
            </div>
            <span class="sa-badge done" id="sa-count-done">— tasks</span>
          </label>

          <label class="sa-option" id="sa-opt-unsuccessful">
            <input type="checkbox" id="sa-check-unsuccessful" value="Unsuccessful" />
            <div class="sa-option-label">
              <strong>Unsuccessful</strong>
              <span>Tasks that could not be completed</span>
            </div>
            <span class="sa-badge unsuccessful" id="sa-count-unsuccessful">— tasks</span>
          </label>
        </div>

        <label class="sa-scope" id="sa-scope">
          <input type="checkbox" id="sa-check-allpages" />
          <div class="sa-scope-label">
            <strong>Archive across all pages</strong>
            <span id="sa-scope-detail">Includes matching tasks not currently visible</span>
          </div>
        </label>

        <div class="sa-preview" id="sa-preview">
          Select at least one status above.
        </div>

        <div class="sa-actions">
          <button class="sa-cancel" id="sa-cancel" type="button">Cancel</button>
          <button class="sa-confirm" id="sa-confirm" type="button" disabled>
            <span class="sa-confirm-label">Hold to Archive</span>
          </button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) closeModal();
    });

    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') closeModal();
    });

    document.getElementById('sa-cancel').addEventListener('click', closeModal);

    ['sa-check-done', 'sa-check-unsuccessful', 'sa-check-allpages'].forEach(id => {
      document.getElementById(id).addEventListener('change', updatePreview);
    });

    setupHoldToConfirm();
  }

  // ─── Hold-to-confirm wiring ────────────────────────────────────────────────
  function setupHoldToConfirm() {
    const btn = document.getElementById('sa-confirm');

    const start = (e) => {
      if (btn.disabled || modalState.inFlight) return;
      e.preventDefault();
      btn.classList.add('sa-holding');
      modalState.holdTimer = setTimeout(() => {
        modalState.holdTimer = null;
        btn.classList.remove('sa-holding');
        runArchive();
      }, HOLD_DURATION_MS);
    };

    const cancel = () => {
      if (modalState.holdTimer) {
        clearTimeout(modalState.holdTimer);
        modalState.holdTimer = null;
      }
      btn.classList.remove('sa-holding');
    };

    btn.addEventListener('mousedown', start);
    btn.addEventListener('touchstart', start, { passive: false });
    ['mouseup', 'mouseleave', 'touchend', 'touchcancel', 'blur'].forEach(ev =>
      btn.addEventListener(ev, cancel)
    );
  }

  // ─── Open / close ──────────────────────────────────────────────────────────
  function openModal() {
    buildModal();
    document.getElementById('sa-check-allpages').checked = false;
    setPreviewState('default');
    resetConfirmButton();
    updatePreview();
    document.getElementById(MODAL_ID + '-overlay').classList.add('sa-open');
  }

  function closeModal() {
    const overlay = document.getElementById(MODAL_ID + '-overlay');
    if (overlay) overlay.classList.remove('sa-open');
    if (modalState.holdTimer) {
      clearTimeout(modalState.holdTimer);
      modalState.holdTimer = null;
    }
    const btn = document.getElementById('sa-confirm');
    if (btn) btn.classList.remove('sa-holding');
  }

  function resetConfirmButton() {
    const btn = document.getElementById('sa-confirm');
    if (!btn) return;
    btn.classList.remove('sa-holding');
    const label = btn.querySelector('.sa-confirm-label');
    if (label) label.textContent = 'Hold to Archive';
  }

  // ─── Preview / status box ──────────────────────────────────────────────────
  function setPreviewState(state, message) {
    const preview = document.getElementById('sa-preview');
    if (!preview) return;
    preview.classList.remove('sa-warn', 'sa-error');
    if (state === 'warn')  preview.classList.add('sa-warn');
    if (state === 'error') preview.classList.add('sa-error');
    if (typeof message === 'string') preview.textContent = message;
  }

  function updatePreview() {
    const grid = getGrid();
    const scope = document.getElementById('sa-scope');
    const allPagesCheckbox = document.getElementById('sa-check-allpages');
    const showScope = hasMultiplePages();

    scope.classList.toggle('sa-shown', showScope);
    if (!showScope) allPagesCheckbox.checked = false;
    const allPages = allPagesCheckbox.checked;

    const doneChecked  = document.getElementById('sa-check-done').checked;
    const unsucChecked = document.getElementById('sa-check-unsuccessful').checked;

    const doneTasks  = getMatchingTasks(['Done']);
    const unsucTasks = getMatchingTasks(['Unsuccessful']);

    document.getElementById('sa-count-done').textContent         = taskWord(doneTasks.length);
    document.getElementById('sa-count-unsuccessful').textContent = taskWord(unsucTasks.length);

    if (grid && showScope) {
      const total    = grid.dataSource.total();
      const pageSize = grid.dataSource.pageSize();
      const pages    = Math.ceil(total / pageSize);
      document.getElementById('sa-scope-detail').textContent =
        `${total} total tasks across ${pages} pages — only the current page is checked otherwise`;
    }

    const confirmBtn = document.getElementById('sa-confirm');

    if (!doneChecked && !unsucChecked) {
      setPreviewState('warn', 'Select at least one status above.');
      confirmBtn.disabled = true;
      return;
    }

    if (allPages) {
      setPreviewState(
        'default',
        '✔ All matching tasks across every page will be archived.'
      );
      confirmBtn.disabled = false;
      return;
    }

    const selectedCount =
      (doneChecked  ? doneTasks.length  : 0) +
      (unsucChecked ? unsucTasks.length : 0);

    if (selectedCount === 0) {
      const msg = showScope
        ? 'No matching tasks on this page. Toggle "across all pages" to widen the scope.'
        : 'No matching tasks found in the current view.';
      setPreviewState('warn', msg);
      confirmBtn.disabled = true;
    } else {
      setPreviewState('default', `✔ ${taskWord(selectedCount)} on this page will be archived.`);
      confirmBtn.disabled = false;
    }
  }

  // ─── Archive flow ──────────────────────────────────────────────────────────
  async function runArchive() {
    if (modalState.inFlight) return;

    const statuses = [];
    if (document.getElementById('sa-check-done').checked)         statuses.push('Done');
    if (document.getElementById('sa-check-unsuccessful').checked) statuses.push('Unsuccessful');
    if (!statuses.length) return;

    const allPages   = document.getElementById('sa-check-allpages').checked;
    const confirmBtn = document.getElementById('sa-confirm');
    const labelEl    = confirmBtn.querySelector('.sa-confirm-label');

    modalState.inFlight = true;
    confirmBtn.disabled = true;
    labelEl.textContent = allPages ? 'Loading all pages…' : 'Archiving…';

    let tasks;
    try {
      tasks = allPages
        ? await fetchAllMatchingTasks(statuses)
        : getMatchingTasks(statuses);
    } catch (err) {
      modalState.inFlight = false;
      confirmBtn.disabled = false;
      labelEl.textContent = 'Hold to Archive';
      setPreviewState('error', 'Could not load tasks across pages — please try again.');
      return;
    }

    if (!tasks.length) {
      modalState.inFlight = false;
      confirmBtn.disabled = false;
      labelEl.textContent = 'Hold to Archive';
      setPreviewState('warn', 'No matching tasks found.');
      return;
    }

    const rvt = getRVT();
    if (!rvt) {
      modalState.inFlight = false;
      confirmBtn.disabled = false;
      labelEl.textContent = 'Hold to Archive';
      setPreviewState('error', 'Security token not found — please refresh the page and try again.');
      return;
    }

    const ids = tasks.map(t => t.TaskId);
    labelEl.textContent = `Archiving ${taskWord(ids.length)}…`;

    const xhr = new XMLHttpRequest();
    xhr.open('POST', '/Tasks/BulkArchive', true);
    xhr.setRequestHeader('Content-Type', 'application/json');
    xhr.setRequestHeader('Accept', '*/*');
    xhr.setRequestHeader('RequestVerificationToken', rvt);
    xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');

    xhr.onload = function () {
      modalState.inFlight = false;
      labelEl.textContent = 'Hold to Archive';

      if (xhr.status === 200) {
        // Optimistic local cleanup so rows disappear immediately
        const grid = getGrid();
        if (grid) {
          tasks.forEach(t => {
            try { grid.dataSource.remove(t); } catch (_) { /* item already gone */ }
          });
          refreshGrid();
        }
        confirmBtn.disabled = false;
        closeModal();
      } else {
        confirmBtn.disabled = false;
        setPreviewState(
          'error',
          `Archive failed (HTTP ${xhr.status}) — ${taskWord(ids.length)} could not be archived.`
        );
      }
    };

    xhr.onerror = function () {
      modalState.inFlight = false;
      confirmBtn.disabled = false;
      labelEl.textContent = 'Hold to Archive';
      setPreviewState(
        'error',
        `Network error — ${taskWord(ids.length)} could not be archived. Please try again.`
      );
    };

    xhr.send(JSON.stringify(ids));
  }

  // ─── Inject toolbar button ─────────────────────────────────────────────────
  function injectButton() {
    if (document.getElementById(WRAPPER_ID)) return;

    const toolbar = document.querySelector('.k-grid-toolbar');
    if (!toolbar) return;

    const archiveBtn = toolbar.querySelector('.tasks_bulk_archive');
    if (!archiveBtn) return;

    const spacer = toolbar.querySelector('.k-spacer');
    if (!spacer) return;

    const wrapper = document.createElement('div');
    wrapper.id = WRAPPER_ID;
    wrapper.className = 'k-toolbar-item';
    wrapper.setAttribute('data-overflow', 'never');

    const btn = document.createElement('button');
    btn.id = BUTTON_ID;
    btn.className = 'k-button k-button-md k-rounded-md k-button-solid k-button-solid-base k-toolbar-button';
    btn.type = 'button';
    btn.setAttribute('role', 'button');
    btn.setAttribute('aria-disabled', 'false');
    btn.innerHTML = '<span class="k-button-text">🗂 Smart Archive</span>';

    btn.addEventListener('click', openModal);

    wrapper.appendChild(btn);
    toolbar.insertBefore(wrapper, spacer);
    wrapper.style.order = '-4';
  }

  // ─── Watch for the toolbar to appear (loads after the Tasks tab click) ─────
  function waitForToolbar() {
    if (document.querySelector('.k-grid-toolbar .tasks_bulk_archive')) {
      injectButton();
      return;
    }

    const observer = new MutationObserver(() => {
      if (document.querySelector('.k-grid-toolbar .tasks_bulk_archive')) {
        injectButton();
      }
    });

    observer.observe(document.body, { childList: true, subtree: true });
  }

  waitForToolbar();

})();
