// ==UserScript==
// @name         MelonPatch - Note Tools (Enhanced)
// @namespace    melonlocal
// @version      2.13
// @description  Adds Copy (Rich), Edit Note, Copy to Tasks, and Delete buttons on task notes. Fixes stale NumComments bubble counts including replies. Works in both grid and tree views.
// @match        https://thepatch.melonlocal.com/*
// @grant        none
// @run-at       document-start
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/MelonPatch%20-%20Note%20Tools%20(Enhanced).user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/MelonPatch%20-%20Note%20Tools%20(Enhanced).user.js
// ==/UserScript==

(function () {

  'use strict';

  // ─── Native XHR Intercept (synchronous, prototype-level) ─────────────────
  //
  // Covers all *TasksData endpoints. At readyState=4: parse JSON, fetch true
  // comment counts via sync XHR, patch NumComments, override responseText,
  // then call Kendo's handler. No flash.
  //
  // Selector: '.commentIdentifier' counts both top-level notes
  // (.commentContainer.commentIdentifier) AND replies (.reply.commentIdentifier).
  //
  (function installXHRIntercept() {

    const TASKS_DATA_RE = /\/Tasks\/\w+TasksData/;
    const COMMENT_SELECTOR = '.commentIdentifier';
    const OrigAEL = XMLHttpRequest.prototype.addEventListener;

    XMLHttpRequest.prototype.addEventListener = function(type, fn, options) {
      if (type === 'readystatechange' && typeof fn === 'function') {
        const xhr = this;
        const kendoHandler = fn;

        const syncWrapper = function() {
          if (
            xhr.readyState === 4 &&
            xhr.status === 200 &&
            xhr.responseURL &&
            TASKS_DATA_RE.test(xhr.responseURL)
          ) {
            let parsed;
            try { parsed = JSON.parse(xhr.responseText); }
            catch(e) { kendoHandler.call(xhr); return; }

            const items = (parsed && parsed.Data) || parsed;

            if (Array.isArray(items)) {
              const toSync = items.filter(t => t.ConversationId && t.NumComments > 0);
              let corrected = 0;

              toSync.forEach(task => {
                try {
                  const syncXhr = new XMLHttpRequest();
                  syncXhr.open('GET',
                    '/Comments/GetCommentsPartial?conversationId=' + task.ConversationId + '&page=0',
                    false // synchronous
                  );
                  syncXhr.send();
                  if (syncXhr.status === 200) {
                    const doc = new DOMParser().parseFromString(syncXhr.responseText, 'text/html');
                    const count = doc.querySelectorAll(COMMENT_SELECTOR).length;
                    if (task.NumComments !== count) {
                      console.log('[ML-Tools] task ' + task.TaskId + ': ' + task.NumComments + ' → ' + count);
                      task.NumComments = count;
                      corrected++;
                    }
                  }
                } catch(e) {
                  console.error('[ML-Tools] sync fetch failed for task ' + task.TaskId, e);
                }
              });

              if (corrected > 0) {
                const correctedText = JSON.stringify(parsed);
                try {
                  Object.defineProperty(xhr, 'responseText', {
                    get: () => correctedText,
                    configurable: true
                  });
                  Object.defineProperty(xhr, 'response', {
                    get: () => correctedText,
                    configurable: true
                  });
                } catch(e) {
                  console.warn('[ML-Tools] Could not override responseText:', e);
                }
              }
            }
          }

          kendoHandler.call(xhr);
        };

        return OrigAEL.call(this, type, syncWrapper, options);
      }

      return OrigAEL.call(this, type, fn, options);
    };

    console.log('[ML-Tools] Sync XHR intercept installed');

  })();

  // ─── DOM work waits until <body> exists ───────────────────────────────────
  function whenDomReady(fn) {
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', fn, { once: true });
    } else {
      fn();
    }
  }

  // ─── Styles ───────────────────────────────────────────────────────────────
  const styleCss = `
    .ml-note-actions {
      display: inline-flex;
      gap: 6px;
      margin-left: 10px;
      vertical-align: middle;
    }
    .ml-note-btn {
      background: #ffffff;
      border: 1px solid #d1d5db;
      border-radius: 6px;
      padding: 3px 10px;
      font-size: 11px;
      font-weight: 500;
      cursor: pointer;
      color: #374151;
      transition: all 0.15s ease;
      box-shadow: 0 1px 2px rgba(0,0,0,0.05);
      line-height: 1.5;
    }
    .ml-note-btn:hover {
      background: #f9fafb;
      border-color: #9ca3af;
      transform: translateY(-1px);
    }
    .ml-note-btn:active { transform: translateY(0); }
    .ml-note-btn.copied { color: #059669; border-color: #059669; background: #ecfdf5; }
    .ml-note-btn.delete { color: #dc2626; border-color: #fca5a5; }
    .ml-note-btn.delete:hover { background: #fef2f2; border-color: #dc2626; }
    .ml-note-btn:disabled { opacity: 0.5; cursor: not-allowed; transform: none; }
    .ml-delete-confirm {
      display: flex;
      align-items: center;
      gap: 8px;
      margin-top: 8px;
      padding: 8px 12px;
      background: #fef2f2;
      border: 1px solid #fca5a5;
      border-radius: 6px;
      font-size: 12px;
      color: #991b1b;
      font-family: system-ui, -apple-system, sans-serif;
      animation: ml-slide-down 0.2s ease-out;
    }
    @keyframes ml-slide-down {
      from { opacity: 0; transform: translateY(-5px); }
      to   { opacity: 1; transform: translateY(0); }
    }
    .ml-delete-confirm span { flex: 1; }
    .ml-delete-yes {
      padding: 4px 12px; border: none; border-radius: 4px;
      background: #dc2626; color: #fff; font-size: 11px; font-weight: 600; cursor: pointer;
    }
    .ml-delete-yes:hover { background: #b91c1c; }
    .ml-delete-yes:disabled { background: #9ca3af; cursor: not-allowed; }
    .ml-delete-no {
      padding: 4px 12px; border: 1px solid #d1d5db; border-radius: 4px;
      background: #fff; font-size: 11px; font-weight: 500; cursor: pointer; color: #374151;
    }
    .ml-delete-no:hover { background: #f9fafb; }
    .ml-delete-no:disabled { opacity: 0.5; cursor: not-allowed; }
    #ml-copy-modal-overlay {
      position: fixed; inset: 0;
      background: rgba(0,0,0,0.5);
      backdrop-filter: blur(2px);
      z-index: 99999;
      display: flex; align-items: center; justify-content: center;
    }
    #ml-copy-modal {
      background: #fff; border-radius: 12px; padding: 24px;
      width: min(600px, 92vw); max-height: 85vh; display: flex; flex-direction: column;
      box-shadow: 0 20px 25px -5px rgba(0,0,0,0.1), 0 10px 10px -5px rgba(0,0,0,0.04);
      font-family: system-ui, -apple-system, sans-serif;
    }
    #ml-copy-modal h3 { margin: 0 0 12px 0; font-size: 16px; font-weight: 700; color: #111827; }
    #ml-copy-modal .ml-note-preview {
      background: #f3f4f6; border-radius: 6px; padding: 8px 12px; font-size: 12px;
      color: #4b5563; margin-bottom: 16px; border-left: 4px solid #d1d5db;
      max-height: 50px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
    }
    #ml-task-search {
      width: 100%; padding: 8px 12px; margin-bottom: 12px; border: 1px solid #d1d5db;
      border-radius: 6px; font-size: 14px; outline: none; box-sizing: border-box;
    }
    #ml-task-search:focus { border-color: #2563eb; box-shadow: 0 0 0 3px rgba(191,219,254,0.6); }
    #ml-copy-modal .ml-task-list {
      overflow-y: auto; flex: 1; border: 1px solid #e5e7eb; border-radius: 6px; margin-bottom: 16px;
    }
    #ml-copy-modal .ml-task-list::-webkit-scrollbar { width: 8px; }
    #ml-copy-modal .ml-task-list::-webkit-scrollbar-track { background: #f9fafb; }
    #ml-copy-modal .ml-task-list::-webkit-scrollbar-thumb { background: #d1d5db; border-radius: 10px; border: 2px solid #f9fafb; }
    #ml-copy-modal .ml-task-list::-webkit-scrollbar-thumb:hover { background: #9ca3af; }
    .ml-task-table { width: 100%; border-collapse: collapse; font-size: 13px; }
    .ml-task-table thead { position: sticky; top: 0; background: #f9fafb; z-index: 1; }
    .ml-task-table th {
      padding: 8px 12px; text-align: left; font-size: 11px; font-weight: 600; color: #6b7280;
      text-transform: uppercase; letter-spacing: 0.05em; border-bottom: 1px solid #e5e7eb;
    }
    .ml-task-table th.col-check { width: 36px; text-align: center; }
    .ml-task-table th.col-status { width: 100px; }
    .ml-task-table tbody tr { cursor: pointer; transition: background 0.1s; border-bottom: 1px solid #f3f4f6; }
    .ml-task-table tbody tr:last-child { border-bottom: none; }
    .ml-task-table tbody tr:hover { background: #f9fafb; }
    .ml-task-table tbody tr.ml-row-checked { background: #eff6ff; }
    .ml-task-table tbody tr.ml-row-checked:hover { background: #dbeafe; }
    .ml-task-table td { padding: 8px 12px; vertical-align: middle; }
    .ml-task-table td.col-check { text-align: center; }
    .ml-no-results { padding: 20px; text-align: center; color: #9ca3af; font-style: italic; display: none; }
    .ml-task-table input[type=checkbox] { width: 15px; height: 15px; cursor: pointer; accent-color: #2563eb; }
    #ml-header-checkbox { width: 15px; height: 15px; cursor: pointer; accent-color: #2563eb; }
    .ml-task-title-cell { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 300px; }
    .ml-task-status {
      font-size: 10px; padding: 2px 6px; border-radius: 9999px;
      background: #e5e7eb; color: #4b5563; display: inline-block; white-space: nowrap;
    }
    #ml-copy-modal .ml-modal-footer {
      display: flex; justify-content: space-between; align-items: center; gap: 12px; flex-wrap: wrap;
    }
    .ml-footer-left { display: flex; gap: 12px; align-items: center; }
    #ml-selection-count { font-size: 12px; color: #6b7280; }
    .ml-modal-actions { display: flex; gap: 8px; }
    .ml-btn-cancel { padding: 8px 16px; border: 1px solid #d1d5db; border-radius: 6px; background: #fff; cursor: pointer; font-size: 13px; font-weight: 500; }
    .ml-btn-copy { padding: 8px 16px; border: none; border-radius: 6px; background: #166534; color: #fff; cursor: pointer; font-size: 13px; font-weight: 600; }
    .ml-btn-copy:disabled { background: #9ca3af; cursor: not-allowed; }
    #ml-status-msg { font-size: 12px; margin-top: 10px; min-height: 1.2em; text-align: center; }
    .ml-toast {
      position: fixed; bottom: 24px; left: 50%; transform: translateX(-50%);
      background: #111827; color: #fff; padding: 10px 16px; border-radius: 8px;
      font-size: 13px; font-family: system-ui, -apple-system, sans-serif;
      box-shadow: 0 10px 25px -5px rgba(0,0,0,0.3);
      z-index: 100000; animation: ml-toast-in 0.2s ease-out;
    }
    @keyframes ml-toast-in {
      from { opacity: 0; transform: translate(-50%, 10px); }
      to   { opacity: 1; transform: translate(-50%, 0); }
    }
  `;

  // ─── Helpers ──────────────────────────────────────────────────────────────
  const COMMENT_SELECTOR = '.commentIdentifier';

  const getCsrfToken = () =>
    document.querySelector('input[name="__RequestVerificationToken"]')?.value || '';

  const getCurrentTaskId = () =>
    document.querySelector('input[name="TaskId"]')?.value || null;

  const getNoteText = (container) =>
    container.querySelector('.editor-contents.patchNote')?.innerHTML?.trim() || '';

  const getNoteTextPlain = (container) =>
    container.querySelector('.editor-contents.patchNote')?.innerText?.trim() || '';

  function showToast(message, ms = 2500) {
    const t = document.createElement('div');
    t.className = 'ml-toast';
    t.textContent = message;
    document.body.appendChild(t);
    setTimeout(() => t.remove(), ms);
  }

  function getKendoWidget() {
    if (typeof jQuery === 'undefined') return null;
    return jQuery('[data-role="treelist"]').data('kendoTreeList')
        || jQuery('[data-role="grid"]').data('kendoGrid');
  }

  async function runPool(items, limit, worker) {
    let idx = 0;
    const runners = Array.from({ length: Math.min(limit, items.length) }, async () => {
      while (idx < items.length) { const i = idx++; await worker(items[i], i); }
    });
    await Promise.all(runners);
  }

  // ─── In-session bubble sync (post-action: delete, copy-to-tasks) ──────────
  async function fetchTrueCount(conversationId) {
    const res = await fetch('/Comments/GetCommentsPartial?conversationId=' + conversationId + '&page=0');
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const html = await res.text();
    const doc = new DOMParser().parseFromString(html, 'text/html');
    return doc.querySelectorAll(COMMENT_SELECTOR).length;
  }

  async function syncTaskBubble(taskId, conversationId) {
    if (!conversationId || !taskId) return;
    try {
      const trueCount = await fetchTrueCount(conversationId);
      if (typeof jQuery === 'undefined') return;

      // Try both widgets — task may live in either grid or treelist (or both)
      const widgets = [
        jQuery('[data-role="treelist"]').data('kendoTreeList'),
        jQuery('[data-role="grid"]').data('kendoGrid')
      ].filter(Boolean);

      for (const widget of widgets) {
        const item = Array.from(widget.dataSource.data())
          .find(d => String(d.TaskId) === String(taskId));
        if (!item) continue;
        const before = item.NumComments;
        if (before !== trueCount) {
          if (typeof item.set === 'function') item.set('NumComments', trueCount);
          else item.NumComments = trueCount;
          console.log('[ML-Tools] syncTaskBubble: task ' + taskId + ' ' + before + ' → ' + trueCount);
        }
      }
    } catch(e) {
      console.error('[ML-Tools] syncTaskBubble failed for task', taskId, e);
    }
  }

  // ─── API wrappers ─────────────────────────────────────────────────────────
  async function getTaskVm(taskId) {
    const r = await fetch('/Tasks/GetTaskVm?taskId=' + taskId);
    return r.json();
  }

  async function postNote(conversationId, taskId, htmlText) {
    const token = getCsrfToken();
    return fetch('/Comments/AddNewNote', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'RequestVerificationToken': token,
        'X-Requested-With': 'XMLHttpRequest'
      },
      body: JSON.stringify({
        ConversationId: parseInt(conversationId),
        Text: htmlText,
        ParentType: 'Task',
        ParentId: String(taskId)
      })
    });
  }

  async function deleteNote(commentId) {
    const token = getCsrfToken();
    return fetch('/Comments/DeleteComment/' + commentId, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        'RequestVerificationToken': token,
        'X-Requested-With': 'XMLHttpRequest'
      }
    });
  }

  function getAllTasks() {
    if (typeof jQuery === 'undefined') return [];
    const grid     = jQuery('[data-role="grid"]').data('kendoGrid');
    const treelist = jQuery('[data-role="treelist"]').data('kendoTreeList');

    // Both widgets may be mounted simultaneously (one per view); read from any
    // that has data and dedupe by TaskId. Previously we hard-preferred treelist,
    // which returned [] in grid view because the hidden treelist was empty.
    const rows = [];
    if (treelist) rows.push(...Array.from(treelist.dataSource.data()));
    if (grid)     rows.push(...Array.from(grid.dataSource.data()));
    if (!rows.length) return [];

    const seen = new Set();
    return rows
      .map(d => (d.toJSON ? d.toJSON() : d))
      .filter(item => {
        const tid = item?.TaskId;
        if (tid == null || seen.has(tid)) return false;
        seen.add(tid);
        return true;
      })
      .map(item => ({
        TaskId: item.TaskId,
        Title: (item.Title || '').trim(),
        ParentTaskId: item.ParentTaskId,
        Status: item.Status || item.TaskStatus
      }));
  }

  function waitForNoteEditor(timeoutMs = 3000) {
    return new Promise((resolve, reject) => {
      const check = () => {
        const frame = document.querySelector('#commentEditor iframe');
        const doc = frame?.contentDocument || frame?.contentWindow?.document;
        if (doc?.body && doc.body.getAttribute('contenteditable')) return doc;
        return null;
      };
      const found = check();
      if (found) return resolve(found);
      const mo = new MutationObserver(() => {
        const d = check();
        if (d) { mo.disconnect(); clearTimeout(timer); resolve(d); }
      });
      mo.observe(document.body, { childList: true, subtree: true });
      const timer = setTimeout(() => { mo.disconnect(); reject(new Error('Editor not found')); }, timeoutMs);
    });
  }

  // ─── Copy to Tasks Modal ──────────────────────────────────────────────────
  function openCopyModal(noteHtml, notePreviewText, sourceTaskId, triggerEl) {
    document.getElementById('ml-copy-modal-overlay')?.remove();

    const tasks = getAllTasks();
    if (!tasks.length) {
      showToast('No tasks found. Make sure the Task list or Tree view is loaded.');
      return;
    }

    const overlay = document.createElement('div');
    overlay.id = 'ml-copy-modal-overlay';
    overlay.setAttribute('role', 'dialog');
    overlay.setAttribute('aria-modal', 'true');
    overlay.setAttribute('aria-labelledby', 'ml-copy-modal-title');
    overlay.innerHTML = `
      <div id="ml-copy-modal">
        <h3 id="ml-copy-modal-title">📋 Copy Note to Tasks</h3>
        <div class="ml-note-preview" id="ml-note-preview"></div>
        <input type="text" id="ml-task-search" placeholder="Search by task title..." autocomplete="off" aria-label="Search tasks">
        <div class="ml-task-list" id="ml-task-list" role="list">
          <table class="ml-task-table">
            <thead>
              <tr>
                <th class="col-check"><input type="checkbox" id="ml-header-checkbox" aria-label="Select all visible"></th>
                <th class="col-title">Task</th>
                <th class="col-status">Status</th>
              </tr>
            </thead>
            <tbody id="ml-task-tbody"></tbody>
          </table>
          <div id="ml-no-results" class="ml-no-results">No tasks match your search</div>
        </div>
        <div class="ml-modal-footer">
          <div class="ml-footer-left"><span id="ml-selection-count">0 selected</span></div>
          <div class="ml-modal-actions">
            <button class="ml-btn-cancel" id="ml-modal-cancel" type="button">Cancel</button>
            <button class="ml-btn-copy" id="ml-modal-confirm" type="button" disabled>Copy to Selected</button>
          </div>
        </div>
        <div id="ml-status-msg" aria-live="polite"></div>
      </div>
    `;
    document.body.appendChild(overlay);
    document.getElementById('ml-note-preview').textContent = notePreviewText;

    const tbody          = document.getElementById('ml-task-tbody');
    const searchInput    = document.getElementById('ml-task-search');
    const confirmBtn     = document.getElementById('ml-modal-confirm');
    const headerCheckbox = document.getElementById('ml-header-checkbox');
    const selectionCount = document.getElementById('ml-selection-count');
    const noResults      = document.getElementById('ml-no-results');

    tasks.filter(t => String(t.TaskId) !== String(sourceTaskId)).forEach(task => {
      const tr = document.createElement('tr');
      tr.dataset.taskid = task.TaskId;

      const tdCheck = document.createElement('td'); tdCheck.className = 'col-check';
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.className = 'ml-task-checkbox';
      checkbox.dataset.taskid = task.TaskId;
      tdCheck.appendChild(checkbox);

      const tdTitle = document.createElement('td');
      tdTitle.className = 'ml-task-title-cell';
      tdTitle.textContent = task.Title;

      const tdStatus = document.createElement('td');
      if (task.Status) {
        const b = document.createElement('span');
        b.className = 'ml-task-status';
        b.textContent = task.Status;
        tdStatus.appendChild(b);
      }

      tr.append(tdCheck, tdTitle, tdStatus);
      tbody.appendChild(tr);

      tr.addEventListener('click', (e) => {
        if (e.target === checkbox) return;
        checkbox.checked = !checkbox.checked;
        tr.classList.toggle('ml-row-checked', checkbox.checked);
        updateSelectionState();
      });

      checkbox.addEventListener('change', () => {
        tr.classList.toggle('ml-row-checked', checkbox.checked);
        updateSelectionState();
      });
    });

    function visibleRows() {
      return Array.from(tbody.querySelectorAll('tr')).filter(tr => tr.style.display !== 'none');
    }

    function updateSelectionState() {
      const n = overlay.querySelectorAll('.ml-task-checkbox:checked').length;
      confirmBtn.disabled = n === 0;
      confirmBtn.textContent = n === 0 ? 'Copy to Selected' : 'Copy to ' + n + ' task' + (n === 1 ? '' : 's');
      selectionCount.textContent = n + ' selected';
      const visible = visibleRows();
      const vc = visible.filter(tr => tr.querySelector('.ml-task-checkbox')?.checked);
      headerCheckbox.checked = vc.length > 0 && vc.length === visible.length;
      headerCheckbox.indeterminate = vc.length > 0 && vc.length < visible.length;
    }

    headerCheckbox.addEventListener('change', () => {
      visibleRows().forEach(tr => {
        const cb = tr.querySelector('.ml-task-checkbox');
        if (cb) { cb.checked = headerCheckbox.checked; tr.classList.toggle('ml-row-checked', cb.checked); }
      });
      updateSelectionState();
    });

    searchInput.addEventListener('input', (e) => {
      const term = e.target.value.toLowerCase();
      let count = 0;
      Array.from(tbody.querySelectorAll('tr')).forEach(tr => {
        const match = (tr.querySelector('.ml-task-title-cell')?.textContent.toLowerCase() || '').includes(term);
        tr.style.display = match ? '' : 'none';
        if (match) count++;
      });
      noResults.style.display = count === 0 ? 'block' : 'none';
      updateSelectionState();
    });

    setTimeout(() => searchInput.focus(), 50);

    const close = () => {
      document.removeEventListener('keydown', onKey);
      overlay.remove();
      triggerEl?.focus();
    };

    const onKey = (e) => { if (e.key === 'Escape') close(); };
    document.addEventListener('keydown', onKey);
    document.getElementById('ml-modal-cancel').onclick = close;
    overlay.onclick = (e) => { if (e.target === overlay) close(); };

    confirmBtn.onclick = async () => {
      const selected = Array.from(overlay.querySelectorAll('.ml-task-checkbox:checked'))
        .map(cb => parseInt(cb.dataset.taskid));
      if (!selected.length) return;

      const status = document.getElementById('ml-status-msg');
      confirmBtn.disabled = true;
      headerCheckbox.disabled = true;
      searchInput.disabled = true;
      Array.from(tbody.querySelectorAll('tr')).forEach(tr => tr.style.pointerEvents = 'none');

      const total = selected.length;
      let done = 0, ok = 0, fail = 0;

      status.style.color = '';
      status.textContent = 'Sending 0 / ' + total + '...';

      const syncTargets = [];

      await runPool(selected, 4, async (id) => {
        try {
          const vm = await getTaskVm(id);
          if (!vm.ConversationId) throw new Error('No conversation ID');
          const res = await postNote(vm.ConversationId, id, noteHtml);
          if (!res.ok) throw new Error('HTTP ' + res.status);
          syncTargets.push({ taskId: id, conversationId: vm.ConversationId });
          ok++;
        } catch(err) {
          console.error('Copy failed for task', id, err);
          fail++;
        } finally {
          done++;
          status.textContent = 'Sending ' + done + ' / ' + total + '...';
        }
      });

      await runPool(syncTargets, 4, ({ taskId, conversationId }) =>
        syncTaskBubble(taskId, conversationId)
      );

      status.style.color = fail ? '#b45309' : '#166534';
      status.textContent = fail
        ? '✅ Copied to ' + ok + ' task' + (ok === 1 ? '' : 's') + ', ⚠️ ' + fail + ' failed'
        : '✅ Successfully copied to ' + ok + ' task' + (ok === 1 ? '' : 's') + '.';

      setTimeout(close, 1800);
    };
  }

  // ─── Inject buttons into each note ────────────────────────────────────────
  function injectNoteButtons(container) {
    if (container.querySelector('.ml-note-actions')) return;
    const actions = container.querySelector('.comment-actions');
    if (!actions) return;

    const noteHtml  = getNoteText(container);
    const notePlain = getNoteTextPlain(container);
    if (!noteHtml) return;

    const commentId = container.dataset.commentid || container.id?.split('_').pop();

    const btnGroup = document.createElement('span');
    btnGroup.className = 'ml-note-actions';

    // ── Copy ──
    const copyBtn = document.createElement('button');
    copyBtn.className = 'ml-note-btn';
    copyBtn.textContent = '📋 Copy';
    copyBtn.onclick = async () => {
      try {
        await navigator.clipboard.write([new ClipboardItem({
          'text/html':  new Blob([noteHtml],  { type: 'text/html'  }),
          'text/plain': new Blob([notePlain], { type: 'text/plain' })
        })]);
        copyBtn.textContent = '✓ Copied';
        copyBtn.classList.add('copied');
        setTimeout(() => { copyBtn.textContent = '📋 Copy'; copyBtn.classList.remove('copied'); }, 2000);
      } catch {
        try {
          await navigator.clipboard.writeText(notePlain);
          showToast('Rich copy unavailable — copied as plain text');
        } catch {
          showToast('Clipboard copy failed');
        }
      }
    };

    // ── Edit Note ──
    const editNoteBtn = document.createElement('button');
    editNoteBtn.className = 'ml-note-btn';
    editNoteBtn.textContent = '✏️ Edit Note';
    editNoteBtn.onclick = async () => {
      const orig = editNoteBtn.textContent;
      editNoteBtn.textContent = '⌛ Opening...';
      const addBtn = document.querySelector('.addNoteButton');
      if (addBtn) addBtn.click();
      try {
        const doc = await waitForNoteEditor();
        doc.body.innerHTML = noteHtml;
      } catch {
        showToast('Could not open the note editor');
      } finally {
        editNoteBtn.textContent = orig;
      }
    };

    // ── Copy to Tasks ──
    const copyToTasksBtn = document.createElement('button');
    copyToTasksBtn.className = 'ml-note-btn';
    copyToTasksBtn.textContent = '📤 Copy to Tasks';
    copyToTasksBtn.onclick = () => {
      const preview = notePlain.length > 70 ? notePlain.substring(0, 70) + '…' : notePlain;
      openCopyModal(noteHtml, preview, getCurrentTaskId(), copyToTasksBtn);
    };

    // ── Delete ──
    const deleteBtn = document.createElement('button');
    deleteBtn.className = 'ml-note-btn delete';
    deleteBtn.textContent = '🗑️ Delete';
    let confirmBanner = null;
    const removeConfirmBanner = () => { confirmBanner?.remove(); confirmBanner = null; };

    deleteBtn.onclick = () => {
      if (confirmBanner) { removeConfirmBanner(); return; }
      if (!commentId) { showToast('Could not find comment ID — delete unavailable.'); return; }

      confirmBanner = document.createElement('div');
      confirmBanner.className = 'ml-delete-confirm';
      const msg    = document.createElement('span'); msg.textContent = 'Delete this note? This cannot be undone.';
      const yesBtn = document.createElement('button'); yesBtn.className = 'ml-delete-yes'; yesBtn.textContent = 'Delete';
      const noBtn  = document.createElement('button'); noBtn.className = 'ml-delete-no';  noBtn.textContent = 'Cancel';
      confirmBanner.append(msg, yesBtn, noBtn);
      actions.insertAdjacentElement('afterend', confirmBanner);

      noBtn.onclick = removeConfirmBanner;

      yesBtn.onclick = async () => {
        yesBtn.disabled = true;
        noBtn.disabled  = true;
        yesBtn.textContent = '⌛ Deleting…';
        try {
          const res = await deleteNote(commentId);
          if (!res.ok) throw new Error('HTTP ' + res.status);
          container.remove();
          const taskId = getCurrentTaskId();
          if (taskId) {
            try {
              const vm = await getTaskVm(taskId);
              if (vm?.ConversationId) await syncTaskBubble(taskId, vm.ConversationId);
            } catch(e) {
              console.warn('[ML-Tools] Could not sync bubble after delete:', e);
            }
          }
          showToast('Note deleted.');
        } catch(err) {
          console.error('Delete failed', err);
          showToast('Delete failed — please try again.');
          yesBtn.disabled = false;
          noBtn.disabled  = false;
          yesBtn.textContent = 'Delete';
        }
      };
    };

    btnGroup.append(copyBtn, editNoteBtn, copyToTasksBtn, deleteBtn);
    actions.appendChild(btnGroup);
  }

  // ─── Init ─────────────────────────────────────────────────────────────────
  whenDomReady(() => {
    const styleEl = document.createElement('style');
    styleEl.textContent = styleCss;
    document.head.appendChild(styleEl);

    let pending = false;
    const scan = () => {
      document.querySelectorAll('.commentContainer.commentIdentifier').forEach(injectNoteButtons);
    };

    const observer = new MutationObserver(() => {
      if (pending) return;
      pending = true;
      requestAnimationFrame(() => { pending = false; scan(); });
    });

    observer.observe(document.body, { childList: true, subtree: true });
    scan();
  });

})();