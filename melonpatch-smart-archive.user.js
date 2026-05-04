// ==UserScript==
// @name         MelonPatch Smart Archive
// @namespace    https://thepatch.melonlocal.com/
// @version      1.0
// @description  Adds a "Smart Archive" button next to the Archive button on the Tasks page. Lets you choose to archive Done, Unsuccessful, or both statuses in one click.
// @author       MelonLocal
// @match        https://thepatch.melonlocal.com/Agents/Dashboard/*
// @grant        none
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/melonpatch-smart-archive.user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/melonpatch-smart-archive.user.js
// ==/UserScript==

(function () {
  'use strict';

  // ─── Config ────────────────────────────────────────────────────────────────
  const BUTTON_ID  = 'smart-archive-btn';
  const WRAPPER_ID = 'smart-archive-wrapper';
  const MODAL_ID   = 'smart-archive-modal';

  // ─── Inject CSS ────────────────────────────────────────────────────────────
  const style = document.createElement('style');
  style.textContent = `
    #${WRAPPER_ID} {
      order: -4;          /* Sit alongside the other toolbar buttons */
    }

    #${BUTTON_ID} {
      background-color: #2d6a4f !important;
      color: #fff !important;
      border-color:  #2d6a4f !important;
      font-family: Poppins, sans-serif;
      cursor: pointer;
    }
    #${BUTTON_ID}:hover {
      background-color: #1f4f39 !important;
      border-color: #1f4f39 !important;
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
    #${MODAL_ID}-overlay.sa-open {
      display: flex;
    }

    /* ── Modal box ── */
    #${MODAL_ID} {
      background: #fff;
      border-radius: 12px;
      box-shadow: 0 8px 32px rgba(0,0,0,.22);
      padding: 28px 32px 24px;
      width: 400px;
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

    /* ── Checkboxes ── */
    #${MODAL_ID} .sa-options {
      display: flex;
      flex-direction: column;
      gap: 10px;
      margin-bottom: 22px;
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
      border-color: #2d6a4f;
      background: #f0faf5;
    }
    #${MODAL_ID} .sa-option input[type=checkbox] {
      width: 16px;
      height: 16px;
      accent-color: #2d6a4f;
      cursor: pointer;
      flex-shrink: 0;
    }
    #${MODAL_ID} .sa-option-label {
      flex: 1;
    }
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

    /* ── Preview count ── */
    #${MODAL_ID} .sa-preview {
      font-size: 12.5px;
      color: #555;
      background: #f5f5f2;
      border-radius: 7px;
      padding: 8px 12px;
      margin-bottom: 20px;
      min-height: 34px;
      display: flex;
      align-items: center;
      gap: 6px;
    }
    #${MODAL_ID} .sa-preview.sa-warn { color: #b45309; background: #fffbeb; }

    /* ── Buttons ── */
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
    #${MODAL_ID} .sa-confirm {
      padding: 7px 20px;
      border: none;
      border-radius: 8px;
      background: #2d6a4f;
      color: #fff;
      font-family: Poppins, sans-serif;
      font-size: 13.5px;
      font-weight: 600;
      cursor: pointer;
    }
    #${MODAL_ID} .sa-confirm:hover   { background: #1f4f39; }
    #${MODAL_ID} .sa-confirm:disabled {
      background: #a0a0a0;
      cursor: not-allowed;
    }
  `;
  document.head.appendChild(style);

  // ─── Helpers ───────────────────────────────────────────────────────────────

  /** Return a count label: "4 tasks" or "1 task" */
  function taskWord(n) { return n === 1 ? '1 task' : n + ' tasks'; }

  /** Get the Kendo grid instance (null if not ready) */
  function getGrid() {
    const el = document.getElementById('OpenTasksData');
    return el ? $(el).data('kendoGrid') : null;
  }

  /** Get tasks from the grid matching the chosen statuses */
  function getMatchingTasks(statuses) {
    const grid = getGrid();
    if (!grid) return [];
    return grid.dataSource.data().filter(t => statuses.includes(t.Status));
  }

  /** Get the request-verification token */
  function getRVT() {
    const el = document.querySelector('input[name="__RequestVerificationToken"]');
    return el ? el.value : null;
  }

  // ─── Build modal (once) ────────────────────────────────────────────────────
  function buildModal() {
    if (document.getElementById(MODAL_ID + '-overlay')) return;

    const overlay = document.createElement('div');
    overlay.id = MODAL_ID + '-overlay';
    overlay.innerHTML = `
      <div id="${MODAL_ID}" role="dialog" aria-modal="true" aria-labelledby="sa-title">
        <h2 id="sa-title">🗂 Smart Archive</h2>
        <p class="sa-subtitle">Select which task statuses to archive, then click Confirm.</p>

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

        <div class="sa-preview" id="sa-preview">
          Select at least one status above.
        </div>

        <div class="sa-actions">
          <button class="sa-cancel" id="sa-cancel">Cancel</button>
          <button class="sa-confirm" id="sa-confirm" disabled>Archive</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);

    // Close on overlay click
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) closeModal();
    });

    // Close on Escape
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') closeModal();
    });

    document.getElementById('sa-cancel').addEventListener('click', closeModal);
    document.getElementById('sa-confirm').addEventListener('click', runArchive);

    // Live-update counts whenever a checkbox changes
    ['sa-check-done', 'sa-check-unsuccessful'].forEach(id => {
      document.getElementById(id).addEventListener('change', updatePreview);
    });
  }

  function openModal() {
    buildModal();
    updatePreview();
    document.getElementById(MODAL_ID + '-overlay').classList.add('sa-open');
  }

  function closeModal() {
    const overlay = document.getElementById(MODAL_ID + '-overlay');
    if (overlay) overlay.classList.remove('sa-open');
  }

  /** Recalculate the per-status counts and update the preview line */
  function updatePreview() {
    const doneChecked   = document.getElementById('sa-check-done').checked;
    const unsucChecked  = document.getElementById('sa-check-unsuccessful').checked;

    const doneTasks   = getMatchingTasks(['Done']);
    const unsucTasks  = getMatchingTasks(['Unsuccessful']);

    document.getElementById('sa-count-done').textContent         = taskWord(doneTasks.length);
    document.getElementById('sa-count-unsuccessful').textContent = taskWord(unsucTasks.length);

    const selected = [];
    if (doneChecked)  selected.push(...doneTasks);
    if (unsucChecked) selected.push(...unsucTasks);

    const preview   = document.getElementById('sa-preview');
    const confirmBtn = document.getElementById('sa-confirm');

    if (selected.length === 0) {
      const reason = (!doneChecked && !unsucChecked)
        ? 'Select at least one status above.'
        : 'No matching tasks found in the current view.';
      preview.textContent = reason;
      preview.classList.add('sa-warn');
      confirmBtn.disabled = true;
    } else {
      preview.textContent = `✔ ${taskWord(selected.length)} will be archived.`;
      preview.classList.remove('sa-warn');
      confirmBtn.disabled = false;
    }
  }

  /** POST to BulkArchive then refresh the grid */
  function runArchive() {
    const statuses = [];
    if (document.getElementById('sa-check-done').checked)         statuses.push('Done');
    if (document.getElementById('sa-check-unsuccessful').checked) statuses.push('Unsuccessful');

    const tasks = getMatchingTasks(statuses);
    if (!tasks.length) return;

    const rvt = getRVT();
    if (!rvt) {
      alert('Security token not found — please refresh the page and try again.');
      return;
    }

    const ids = tasks.map(t => t.TaskId);
    const confirmBtn = document.getElementById('sa-confirm');
    confirmBtn.disabled = true;
    confirmBtn.textContent = 'Archiving…';

    const xhr = new XMLHttpRequest();
    xhr.open('POST', '/Tasks/BulkArchive', true);
    xhr.setRequestHeader('Content-Type', 'application/json');
    xhr.setRequestHeader('Accept', '*/*');
    xhr.setRequestHeader('RequestVerificationToken', rvt);
    xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');

    xhr.onload = function () {
      closeModal();
      if (xhr.status === 200) {
        const grid = getGrid();
        if (grid) grid.dataSource.read();
      } else {
        alert('Archive failed (HTTP ' + xhr.status + '). Please try again.');
      }
      confirmBtn.disabled = false;
      confirmBtn.textContent = 'Archive';
    };

    xhr.onerror = function () {
      closeModal();
      alert('Network error while archiving. Please try again.');
      confirmBtn.disabled = false;
      confirmBtn.textContent = 'Archive';
    };

    xhr.send(JSON.stringify(ids));
  }

  // ─── Inject toolbar button ─────────────────────────────────────────────────
  function injectButton() {
    // Only inject once
    if (document.getElementById(WRAPPER_ID)) return;

    const toolbar = document.querySelector('.k-grid-toolbar');
    if (!toolbar) return;

    const archiveBtn = toolbar.querySelector('.tasks_bulk_archive');
    if (!archiveBtn) return;

    const spacer = toolbar.querySelector('.k-spacer');
    if (!spacer) return;

    // Wrapper mirrors the structure of other k-toolbar items
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

    // Insert before the spacer so it stays in the left button group,
    // then use CSS order: -4 to sit alongside the other toolbar buttons
    toolbar.insertBefore(wrapper, spacer);
    wrapper.style.order = '-4';
  }

  // ─── Watch for the toolbar to appear (it loads after the Tasks tab click) ──
  function waitForToolbar() {
    // If the toolbar is already present, inject immediately
    if (document.querySelector('.k-grid-toolbar .tasks_bulk_archive')) {
      injectButton();
      return;
    }

    // Otherwise observe the DOM for it
    const observer = new MutationObserver(() => {
      if (document.querySelector('.k-grid-toolbar .tasks_bulk_archive')) {
        injectButton();
        // Keep observing in case the grid is re-rendered (e.g. tab switch)
      }
    });

    observer.observe(document.body, { childList: true, subtree: true });
  }

  waitForToolbar();

})();