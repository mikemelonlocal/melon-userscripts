// ==UserScript==
// @name         Patch Gong Workflow Helper
// @namespace    http://tampermonkey.net/
// @version      1.10.0
// @description  Controls Patch calls using data from Google Sheets. Resumes across agent navigations via localStorage. Log panel replaces alert spam, Abort stops mid-run, config stored per-browser.
// @match        https://thepatch.melonlocal.com/*
// @run-at       document-end
// @grant        none
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/patch-gong-workflow-helper.user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/patch-gong-workflow-helper.user.js
// ==/UserScript==

(function () {
  const VERSION = "workflow-html-v1.10.0";

  // -----------------------------
  // CONFIG (defaults; per-browser overrides via Config button -> localStorage)
  // -----------------------------
  const DEFAULTS = {
    SHEET_ID: "1WSVMb1yV18kj2YzzbJ3gx_TF9fSrk1ExNnVLjHEX9KY",
    QUEUE_SHEET_GID: "0",
    CALL_TYPE: "Monthly",
    DONE_WEBHOOK_URL: "",
  };

  function cfg(key) {
    try {
      const v = localStorage.getItem(`patchWF.cfg.${key}`);
      if (v !== null && v !== "") return v;
    } catch (e) {}
    return DEFAULTS[key];
  }
  function setCfg(key, value) {
    try {
      if (value === null || value === "") {
        localStorage.removeItem(`patchWF.cfg.${key}`);
      } else {
        localStorage.setItem(`patchWF.cfg.${key}`, String(value));
      }
    } catch (e) {}
  }

  const COL_AGENT_ID = 0;
  const COL_AGENT_NAME = 1;
  const COL_MEETING_TITLE = 2;
  const COL_NOTES = 3;
  const COL_CALL_DATE = 4;
  const COL_DONE = 5;

  const TARGET_STATUS = "Completed";
  const PRIMARY_GRID_STATUS = "Scheduled";
  const FALLBACK_GRID_STATUS = "Completed";

  const STATE_STORAGE_KEY = "patchWF.state.v1";

  // -----------------------------
  // Step names + order
  // -----------------------------
  const STEP_LOAD_ROW = "LOAD_ROW";
  const STEP_NAVIGATE_AGENT = "NAVIGATE_AGENT";
  const STEP_FIND_OR_CREATE = "FIND_OR_CREATE";
  const STEP_OPEN_DETAILS = "OPEN_DETAILS";
  const STEP_APPLY_NOTES = "APPLY_NOTES";
  const STEP_SET_STATUS = "SET_STATUS";
  const STEP_MARK_DONE = "MARK_DONE";
  const STEP_FINISHED = "FINISHED";

  const stepOrder = [
    STEP_LOAD_ROW,
    STEP_NAVIGATE_AGENT,
    STEP_FIND_OR_CREATE,
    STEP_OPEN_DETAILS,
    STEP_APPLY_NOTES,
    STEP_SET_STATUS,
    STEP_MARK_DONE,
  ];

  function freshWfState() {
    return {
      step: STEP_LOAD_ROW,
      rowIndex: null,
      agentId: null,
      meetingTitle: null,
      callDateRaw: null,
      notes: null,
      scenario: null,
      callId: null,
      lastRowTitle: null,
    };
  }

  // -----------------------------
  // Module state
  // -----------------------------
  let queueRows = [];
  let wfState = freshWfState();
  let wfRowEl = null; // DOM ref only; never persisted

  let isRunning = false;
  let abortRequested = false;
  let currentMode = null; // "full" | "step" | "dry" | null

  // -----------------------------
  // Persistence (survives page reloads for Full Auto resume)
  // -----------------------------
  function saveState() {
    try {
      const payload = {
        queueRows,
        wfState,
        running: isRunning && currentMode === "full",
        mode: currentMode,
      };
      localStorage.setItem(STATE_STORAGE_KEY, JSON.stringify(payload));
    } catch (e) {}
  }
  function loadSavedState() {
    try {
      const raw = localStorage.getItem(STATE_STORAGE_KEY);
      if (!raw) return null;
      return JSON.parse(raw);
    } catch (e) {
      return null;
    }
  }
  function clearSavedState() {
    try { localStorage.removeItem(STATE_STORAGE_KEY); } catch (e) {}
  }

  // -----------------------------
  // Utilities
  // -----------------------------
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  async function waitFor(predicate, { timeoutMs = 15000, intervalMs = 150 } = {}) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      if (abortRequested) return null;
      try {
        const v = predicate();
        if (v) return v;
      } catch (e) {}
      await sleep(intervalMs);
    }
    return null;
  }

  function norm(txt) {
    return (txt || "").replace(/\s+/g, " ").trim();
  }

  function isDashboardCallsPage() {
    const u = location.href;
    return /\/Agents\/Dashboard\/\d+/.test(u) && /#calls\b/.test(u);
  }

  function getAgentIdFromDashboardUrl() {
    const m = location.pathname.match(/\/Agents\/Dashboard\/(\d+)/i);
    return m ? Number(m[1]) : null;
  }

  function buildDashboardCallsUrl(agentId) {
    return `${location.origin}/Agents/Dashboard/${agentId}#calls`;
  }

  function getNextPendingRowIndex() {
    for (let i = 0; i < queueRows.length; i++) {
      if (!queueRows[i].done) return i;
    }
    return null;
  }
  function pendingCount() {
    return queueRows.filter((r) => !r.done).length;
  }
  function summarizeRow(row) {
    return `Row #${row.rowNumber} | AgentId ${row.agentId} | "${row.meetingTitle}" | ${row.callDateRaw}`;
  }

  // -----------------------------
  // Log panel (replaces alert spam)
  // -----------------------------
  function log(msg, level = "info") {
    const line = `[${new Date().toTimeString().slice(0, 8)}] ${msg}`;
    const logs = document.getElementById("patchWFLogs");
    if (logs) {
      const el = document.createElement("div");
      const colors = { info: "#222", warn: "#b26b00", error: "#b00020", ok: "#0a7d2a" };
      el.style.cssText = `color:${colors[level] || "#222"};padding:2px 0;border-bottom:1px solid #f0f0f0;white-space:pre-wrap;`;
      el.textContent = line;
      logs.appendChild(el);
      while (logs.children.length > 200) logs.removeChild(logs.firstChild);
      logs.scrollTop = logs.scrollHeight;
    }
    console.log(`[WF][${level}]`, msg);
  }

  // -----------------------------
  // Google Sheets loader
  // -----------------------------
  async function loadQueueFromSheet() {
    const sheetId = cfg("SHEET_ID");
    const gid = cfg("QUEUE_SHEET_GID");
    const url = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?gid=${gid}&tqx=out:json`;
    const res = await fetch(url);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const text = await res.text();
    const json = JSON.parse(text.substring(47).slice(0, -2));
    const rows = json.table.rows || [];
    const loaded = [];

    for (let i = 1; i < rows.length; i++) {
      const rowObj = rows[i];
      if (!rowObj || !rowObj.c) continue;
      const r = rowObj.c;

      const agentIdCell = r[COL_AGENT_ID];
      const agentNameCell = r[COL_AGENT_NAME];
      const meetingTitleCell = r[COL_MEETING_TITLE];
      const notesCell = r[COL_NOTES];
      const callDateCell = r[COL_CALL_DATE];
      const doneCell = r[COL_DONE];

      if (!agentIdCell || !meetingTitleCell) continue;

      const idValue = agentIdCell.v;
      const numericId = Number(idValue);
      if (!numericId || Number.isNaN(numericId)) {
        log(`Skipping row ${i + 1}: invalid AgentId ${idValue}`, "warn");
        continue;
      }

      const done = doneCell && String(doneCell.v || "").trim().toLowerCase() === "yes";

      loaded.push({
        rowNumber: i + 1,
        agentId: numericId,
        agentName: agentNameCell ? agentNameCell.v : "",
        meetingTitle: meetingTitleCell.v,
        notes: notesCell ? String(notesCell.v || "") : "",
        callDateRaw: callDateCell ? callDateCell.v : "",
        done,
      });
    }

    queueRows = loaded;
    wfState = freshWfState();
    saveState();
    log(`Loaded ${queueRows.length} rows. Pending: ${pendingCount()}.`, "ok");
  }

  // -----------------------------
  // Optional Done webhook (Apps Script)
  // text/plain avoids CORS preflight; Apps Script parses JSON body itself.
  // -----------------------------
  async function postDoneWebhook(row) {
    const url = cfg("DONE_WEBHOOK_URL");
    if (!url) return;
    try {
      await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({
          rowNumber: row.rowNumber,
          agentId: row.agentId,
          meetingTitle: row.meetingTitle,
          done: true,
        }),
      });
      log(`Webhook: marked row #${row.rowNumber} Done in sheet`, "ok");
    } catch (e) {
      log(`Webhook failed for row #${row.rowNumber}: ${e.message}`, "warn");
    }
  }

  // -----------------------------
  // Dashboard helpers
  // -----------------------------
  function getDashboardRows() {
    return Array.from(
      document.querySelectorAll("tr.k-table-row.k-master-row, tr.k-master-row")
    );
  }

  function getStatusFromGridRow(row) {
    const badge = row.querySelector(".gridFilledCell");
    return badge ? norm(badge.textContent) : "";
  }

  function findCallRowOnDashboard({ title, callType }) {
    const rows = getDashboardRows();
    const wantTitle = norm(title);
    const wantType = callType ? norm(callType) : null;
    let scheduledMatch = null;
    let completedMatch = null;
    for (const row of rows) {
      const titleEl = row.querySelector(".task-title");
      if (!titleEl) continue;
      const titleText = norm(titleEl.textContent);
      if (titleText !== wantTitle) continue;
      const typeCell = row.querySelector("td.k-table-td:nth-child(3)");
      const typeText = typeCell ? norm(typeCell.textContent) : "";
      if (wantType && typeText !== wantType) continue;
      const statusText = getStatusFromGridRow(row);
      if (statusText === PRIMARY_GRID_STATUS) scheduledMatch = row;
      else if (statusText === FALLBACK_GRID_STATUS && !completedMatch) completedMatch = row;
    }
    return scheduledMatch || completedMatch || null;
  }

  function getCallsGridRowInfos() {
    return getDashboardRows().map((row) => {
      const checkbox = row.querySelector('input[type="checkbox"][aria-label="Select row"]');
      return { row, checkbox };
    });
  }

  function findRowInfoByExactTitle(title) {
    const infos = getCallsGridRowInfos();
    const wantTitle = norm(title);
    let scheduledInfo = null;
    let completedInfo = null;
    for (const info of infos) {
      const tEl = info.row.querySelector(".task-title");
      if (!tEl) continue;
      const titleText = norm(tEl.textContent);
      if (titleText !== wantTitle) continue;
      const statusText = getStatusFromGridRow(info.row);
      if (statusText === PRIMARY_GRID_STATUS) scheduledInfo = info;
      else if (statusText === FALLBACK_GRID_STATUS && !completedInfo) completedInfo = info;
    }
    return scheduledInfo || completedInfo || null;
  }

  function clickCallRow(info) {
    if (!info) return;
    if (info.checkbox && !info.checkbox.checked) info.checkbox.click();
    const clickable =
      info.row.querySelector("td .myLink.cTitle") ||
      info.row.querySelector("td .task-title") ||
      info.row.querySelector(".task-title") ||
      info.row.querySelector("td") ||
      info.row;
    clickable.click();
  }

  async function waitForSidebarAndNotesUi(timeoutMs = 12000) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      if (abortRequested) return null;
      const sidebar = document.querySelector("#sidebarRight");
      const addNoteBtn = findAddNoteButton();
      const latestNote = getLatestNoteDiv();
      if (sidebar && (addNoteBtn || latestNote)) {
        return { sidebar, addNoteBtn, latestNote };
      }
      await sleep(150);
    }
    return null;
  }

  // -----------------------------
  // New Call
  // -----------------------------
  async function clickNewCall() {
    const btn =
      document.querySelector("button#addNewCall.btn.melon-green") ||
      [...document.querySelectorAll("button")].find((b) => norm(b.textContent) === "New Call");
    if (!btn) throw new Error("Could not find 'New Call' button on Dashboard.");
    btn.click();
    const form = await waitFor(
      () => document.querySelector("input#Title") && document.querySelector("input#ScheduledTime"),
      { timeoutMs: 15000, intervalMs: 200 }
    );
    if (!form) throw new Error("New Call form did not appear.");
    await sleep(400);
  }

  function getScheduledPicker() {
    if (!window.jQuery) return null;
    const $input = window.jQuery("input#ScheduledTime");
    if (!$input.length) return null;
    return $input.data("kendoDateTimePicker") || $input.data("kendoDatePicker") || null;
  }

  async function setScheduledDateTime(raw) {
    const dtInput = document.querySelector("input#ScheduledTime");
    if (!dtInput) throw new Error("New Call ScheduledTime input not found.");
    if (!raw) return;
    const picker = getScheduledPicker();
    let valueToUse = raw;
    if (!(raw instanceof Date)) {
      const d = new Date(raw);
      if (!Number.isNaN(d.getTime())) valueToUse = d;
    }
    if (picker) {
      picker.value(valueToUse);
      picker.trigger("change");
    } else {
      dtInput.focus();
      dtInput.select();
      dtInput.value = String(raw);
      dtInput.dispatchEvent(new Event("input", { bubbles: true }));
      dtInput.dispatchEvent(new Event("change", { bubbles: true }));
      await sleep(200);
      dtInput.blur();
    }
    await sleep(300);
  }

  async function fillAndSaveNewCall({ title, callDateRaw }) {
    const titleInput = document.querySelector("input#Title");
    if (!titleInput) throw new Error("New Call Title input not found.");
    titleInput.value = title || "";
    titleInput.dispatchEvent(new Event("input", { bubbles: true }));
    titleInput.dispatchEvent(new Event("change", { bubbles: true }));

    const modalFooter = document.querySelector(".modal-footer");
    if (!modalFooter) throw new Error("New Call modal footer not found.");

    if (callDateRaw) {
      await setScheduledDateTime(callDateRaw);
      modalFooter.click();
      await sleep(300);
    }

    const getSaveButton = () =>
      document.querySelector("button#newCallSave.btn.melon-green") ||
      [...document.querySelectorAll("button")].find((b) => (b.textContent || "").trim() === "Save");

    const isNewCallModalOpen = () => {
      const el = document.querySelector("input#Title");
      return !!(el && el.offsetParent !== null);
    };

    const clickSaveOnce = async () => {
      const saveBtn = getSaveButton();
      if (!saveBtn) throw new Error("New Call Save button not found.");
      saveBtn.click();
      await sleep(1800);
    };

    await clickSaveOnce();
    if (isNewCallModalOpen()) await clickSaveOnce();

    const errorBar =
      document.querySelector("div[role='alert']") ||
      document.querySelector(".alert-danger") ||
      document.querySelector(".validation-summary-errors");
    if (isNewCallModalOpen() && errorBar && errorBar.offsetParent !== null) {
      throw new Error("Patch did not accept Scheduled Date and Time; validation error still visible.");
    }
  }

  async function ensureCallExists({ title, callType, callDateRaw } = {}) {
    const ct = callType || cfg("CALL_TYPE");
    let existing = findCallRowOnDashboard({ title, callType: ct });
    if (existing) return { created: false, rowEl: existing };
    await clickNewCall();
    await fillAndSaveNewCall({ title, callDateRaw });
    const createdRow = await waitFor(
      () => findCallRowOnDashboard({ title, callType: ct }),
      { timeoutMs: 20000, intervalMs: 600 }
    );
    if (!createdRow) throw new Error(`New call "${title}" did not appear in grid after save.`);
    return { created: true, rowEl: createdRow };
  }

  // -----------------------------
  // Status
  // -----------------------------
  function getCurrentStatusText() {
    const btn = document.querySelector("#callStatusButton");
    return btn ? norm(btn.textContent) : "";
  }

  function openStatusMenu(dryRun) {
    const btn = document.querySelector("#callStatusButton");
    if (!btn) return { ok: false, error: "Could not find #callStatusButton." };
    if (dryRun) return { ok: true, info: "Dry run: would click status button." };
    btn.click();
    return { ok: true, info: "Clicked status button." };
  }

  async function setCallStatusByLabel(labelText, { dryRun }) {
    const plan = { ok: true, actions: [] };
    const opened = openStatusMenu(dryRun);
    plan.actions.push(opened);
    if (!opened.ok) return { ok: false, error: opened.error, plan };
    if (!dryRun) await sleep(300);
    const match = await waitFor(
      () => {
        const labels = [...document.querySelectorAll("label[onclick*='ChangeCallStatus(']")];
        return labels.find((l) => norm(l.textContent) === labelText) || null;
      },
      { timeoutMs: 5000 }
    );
    if (!match) return { ok: false, error: `Status option not found by text: "${labelText}"`, plan };
    if (dryRun) {
      plan.actions.push({ ok: true, info: `Dry run: would click "${labelText}".` });
      return plan;
    }
    match.click();
    plan.actions.push({ ok: true, info: `Clicked "${labelText}".` });
    await waitFor(() => getCurrentStatusText() === labelText, { timeoutMs: 8000, intervalMs: 200 });
    return plan;
  }

  // -----------------------------
  // Notes builder / parser
  // -----------------------------
  function buildCallNoteHtml({
    keyPoints = [],
    nextSteps = [],
    keyPointsTitle = "Key points",
    nextStepsTitle = "Next steps",
  } = {}) {
    const esc = (s) =>
      String(s ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#39;");
    const renderList = (items) => {
      const clean = (items || []).map((x) => {
        if (typeof x === "string") return { title: "", body: x };
        return { title: x?.title ?? "", body: x?.body ?? "" };
      });
      if (!clean.length) return "<p><em>None.</em></p>";
      return `
        <ol>
          ${clean
            .map(({ title, body }) => {
              const t = esc(title).trim();
              const b = esc(body).trim();
              if (t && b) return `<li><strong>${t}:</strong> ${b}</li>`;
              if (t && !b) return `<li><strong>${t}</strong></li>`;
              return `<li>${b}</li>`;
            })
            .join("")}
        </ol>
      `.trim();
    };
    return `
      <p><strong>${esc(keyPointsTitle)}</strong></p>
      ${renderList(keyPoints)}
      <p><strong>${esc(nextStepsTitle)}</strong></p>
      ${renderList(nextSteps)}
    `.trim();
  }

  function parseRawNotesToSections(raw) {
    const text = String(raw || "").replace(/\s+/g, " ").trim();
    const keyIdx = text.toLowerCase().indexOf("key points");
    const nextIdx = text.toLowerCase().indexOf("next steps");
    let keyBlock = "";
    let nextBlock = "";
    if (keyIdx !== -1 && nextIdx !== -1) {
      keyBlock = text.slice(keyIdx + "key points".length, nextIdx).trim();
      nextBlock = text.slice(nextIdx + "next steps".length).trim();
    } else {
      keyBlock = text;
    }
    const splitNumbered = (block) => {
      const parts = block.split(/\s(?=\d+\.\s)/g);
      return parts
        .map((p) => p.trim())
        .filter((p) => /^\d+\.\s/.test(p))
        .map((p) => p.replace(/^\d+\.\s*/, "").trim());
    };
    const keyItemsRaw = splitNumbered(keyBlock);
    const nextItemsRaw = splitNumbered(nextBlock);
    const keyPoints = keyItemsRaw.map((item) => {
      const m = item.match(/^([^:]+):\s*(.*)$/);
      if (m) return { title: m[1].trim(), body: m[2].trim() };
      return { title: "", body: item };
    });
    const nextSteps = nextItemsRaw.map((item) => item.trim()).filter(Boolean);
    return { keyPoints, nextSteps };
  }

  function getLatestNoteDiv() {
    const notes = [...document.querySelectorAll("div.editor-contents.patchNote.myLink")];
    return notes.length ? notes[notes.length - 1] : null;
  }

  function findAddNoteButton() {
    return (
      document.querySelector("button.addNoteButton") ||
      [...document.querySelectorAll("button")].find((b) => norm(b.textContent) === "Add Note") ||
      null
    );
  }
  function findNewNoteSaveButton() {
    return (
      document.querySelector("button.saveNote") ||
      [...document.querySelectorAll("button[type='submit']")].find((b) => b.classList.contains("saveNote")) ||
      null
    );
  }
  function findEditNoteSaveButton() {
    return (
      document.querySelector("button#editNoteSave") ||
      document.querySelector("button[data-commentid][id*='Save']") ||
      null
    );
  }

  function getKendoEditorInstance() {
    const ta =
      document.querySelector("textarea.k-raw-content[data-role='editor']") ||
      document.querySelector("textarea.k-raw-content") ||
      document.querySelector("#NoteEditor") ||
      null;
    if (ta && window.jQuery) {
      const inst = window.jQuery(ta).data("kendoEditor");
      if (inst) return inst;
    }
    return null;
  }
  function getEditorTextarea() {
    return (
      document.querySelector("textarea.k-raw-content[data-role='editor']") ||
      document.querySelector("textarea.k-raw-content") ||
      document.querySelector("#NoteEditor") ||
      document.querySelector("textarea[name='NoteEditor']") ||
      null
    );
  }
  function getEditorIframe() {
    return (
      document.querySelector("iframe.k-iframe[aria-label='editor']") ||
      document.querySelector("iframe.k-iframe") ||
      null
    );
  }

  async function setEditorHtml(noteHtml) {
    const inst = getKendoEditorInstance();
    if (inst) {
      inst.value(noteHtml);
      try { inst.trigger("change"); } catch (e) {}
      return { ok: true, info: "Set editor via kendoEditor.value()." };
    }
    const ta = getEditorTextarea();
    if (ta) {
      ta.value = noteHtml;
      ta.dispatchEvent(new Event("input", { bubbles: true }));
      ta.dispatchEvent(new Event("change", { bubbles: true }));
      return { ok: true, info: "Set editor via textarea.value." };
    }
    const iframe = getEditorIframe();
    if (iframe && iframe.contentDocument && iframe.contentDocument.body) {
      iframe.contentDocument.body.innerHTML = noteHtml;
      return { ok: true, info: "Set editor via iframe.body.innerHTML." };
    }
    return { ok: false, error: "No editor found (Kendo/textarea/iframe)." };
  }

  async function openEditorForNewNote({ dryRun }) {
    const btn = findAddNoteButton();
    if (!btn) return { ok: false, error: "Could not find Add Note button." };
    if (dryRun) return { ok: true, info: "Dry run: would click Add Note." };
    btn.click();
    await waitFor(
      () => findNewNoteSaveButton() || getEditorTextarea() || getKendoEditorInstance() || getEditorIframe(),
      { timeoutMs: 12000 }
    );
    return { ok: true, info: "Opened new note editor." };
  }

  async function openEditorForExistingNote(latestNoteDiv, { dryRun }) {
    if (!latestNoteDiv) return { ok: false, error: "No existing note div." };
    if (dryRun) return { ok: true, info: "Dry run: would click latest note." };
    latestNoteDiv.click();
    await waitFor(
      () => findEditNoteSaveButton() || getEditorTextarea() || getKendoEditorInstance() || getEditorIframe(),
      { timeoutMs: 12000 }
    );
    return { ok: true, info: "Opened existing note editor." };
  }

  async function clickSaveAndWait({ mode, dryRun }) {
    const btn = mode === "edit" ? findEditNoteSaveButton() : findNewNoteSaveButton();
    if (!btn) return { ok: false, error: `Could not find ${mode} Save button.` };
    if (dryRun) return { ok: true, info: `Dry run: would click ${mode} Save.` };
    btn.click();
    await sleep(800);
    return { ok: true, info: `Clicked ${mode} Save.` };
  }

  // -----------------------------
  // Step implementations
  // -----------------------------
  function ensureRowLoaded() {
    const idx = getNextPendingRowIndex();
    if (idx === null) return null;
    const row = queueRows[idx];
    wfState.rowIndex = idx;
    wfState.agentId = row.agentId;
    wfState.meetingTitle = row.meetingTitle;
    wfState.callDateRaw = row.callDateRaw;
    wfState.notes = row.notes;
    wfState.callId = null;
    wfState.lastRowTitle = row.meetingTitle;
    wfRowEl = null;
    return row;
  }

  async function stepLoadRow() {
    const row = ensureRowLoaded();
    if (!row) return { ok: false, info: "No pending rows left." };
    wfState.step = STEP_NAVIGATE_AGENT;
    return { ok: true, info: `Loaded row: ${summarizeRow(row)}` };
  }

  async function stepNavigateAgent({ dryRun }) {
    const row = queueRows[wfState.rowIndex];
    const currentAgentId = getAgentIdFromDashboardUrl();
    const url = buildDashboardCallsUrl(row.agentId);
    if (currentAgentId === row.agentId && isDashboardCallsPage()) {
      if (!dryRun) await sleep(1000);
      wfState.step = STEP_FIND_OR_CREATE;
      return { ok: true, info: "Already on correct agent dashboard." };
    }
    if (dryRun) return { ok: true, info: `Would navigate to agent dashboard: ${url}` };

    // Persist state BEFORE navigating so the reloaded script can resume.
    wfState.step = STEP_FIND_OR_CREATE;
    saveState();
    log(`Navigating to agent ${row.agentId}...`, "info");
    window.location.href = url;
    return { ok: true, navigating: true, info: "Navigating to agent dashboard..." };
  }

  async function stepFindOrCreate({ dryRun }) {
    const row = queueRows[wfState.rowIndex];
    if (!isDashboardCallsPage()) {
      return { ok: false, info: "Not on dashboard; cannot find/create meeting." };
    }
    wfRowEl = null;
    wfState.meetingTitle = row.meetingTitle;
    wfState.callDateRaw = row.callDateRaw;
    wfState.callId = null;
    wfState.lastRowTitle = row.meetingTitle;

    const callType = cfg("CALL_TYPE");

    if (dryRun) {
      const existing = !!findCallRowOnDashboard({ title: row.meetingTitle, callType });
      wfState.scenario = existing ? "B" : "C";
      wfState.step = STEP_OPEN_DETAILS;
      return {
        ok: true,
        info: existing
          ? `Meeting "${row.meetingTitle}" exists; would open it.`
          : `Meeting "${row.meetingTitle}" does NOT exist; would create via New Call.`,
      };
    }

    try {
      const { created, rowEl } = await ensureCallExists({
        title: row.meetingTitle,
        callType,
        callDateRaw: row.callDateRaw,
      });
      let finalRowEl = rowEl || null;
      if (!finalRowEl) {
        const info = findRowInfoByExactTitle(row.meetingTitle);
        if (!info) {
          return { ok: false, info: `Meeting "${row.meetingTitle}" was created/found, but its grid row cannot be located.` };
        }
        finalRowEl = info.row;
      }
      wfState.scenario = created ? "C" : "B";
      wfRowEl = finalRowEl;
      wfState.lastRowTitle = row.meetingTitle;
      wfState.step = STEP_OPEN_DETAILS;
      return {
        ok: true,
        info: created
          ? `Created meeting "${row.meetingTitle}".`
          : `Meeting "${row.meetingTitle}" already exists.`,
      };
    } catch (e) {
      return { ok: false, info: `Failed to create/find meeting: ${e.message}` };
    }
  }

  async function stepOpenDetails({ dryRun }) {
    const row = queueRows[wfState.rowIndex];
    if (!row) return { ok: false, info: "OPEN_DETAILS: No current queue row." };
    const title = row.meetingTitle;

    let info = null;
    if (wfRowEl && document.contains(wfRowEl)) {
      const checkbox = wfRowEl.querySelector('input[type="checkbox"][aria-label="Select row"]') || null;
      info = { row: wfRowEl, checkbox };
    } else {
      info = findRowInfoByExactTitle(title);
      if (!info) return { ok: false, info: `OPEN_DETAILS: Row for "${title}" not found in grid.` };
      wfRowEl = info.row;
      wfState.lastRowTitle = title;
    }

    if (dryRun) {
      wfState.step = STEP_APPLY_NOTES;
      return { ok: true, info: `OPEN_DETAILS (DRY): Would click row for "${title}".` };
    }

    clickCallRow(info);
    const ready = await waitForSidebarAndNotesUi(12000);
    if (!ready) return { ok: false, info: "OPEN_DETAILS: Call Details did not load Notes UI in time." };
    wfState.step = STEP_APPLY_NOTES;
    return { ok: true, info: "OPEN_DETAILS: Clicked grid row, Notes UI ready." };
  }

  async function stepApplyNotes({ dryRun }) {
    const hasNotesUi = getLatestNoteDiv() || findAddNoteButton() || document.querySelector("button.addNoteButton");
    if (!hasNotesUi) return { ok: false, info: "Notes UI not ready yet." };
    const raw = wfState.notes || "";
    const { keyPoints, nextSteps } = parseRawNotesToSections(raw);
    const noteHtml = buildCallNoteHtml({ keyPoints, nextSteps });
    const latestNote = getLatestNoteDiv();

    if (!latestNote) {
      if (dryRun) {
        wfState.step = STEP_SET_STATUS;
        return { ok: true, info: "No note exists. Would add new note." };
      }
      const opened = await openEditorForNewNote({ dryRun: false });
      if (!opened.ok) return opened;
      const setRes = await setEditorHtml(noteHtml);
      if (!setRes.ok) return setRes;
      const saved = await clickSaveAndWait({ mode: "new", dryRun: false });
      if (!saved.ok) return saved;
      wfState.step = STEP_SET_STATUS;
      return { ok: true, info: "Added new formatted note and saved." };
    }

    if (dryRun) {
      wfState.step = STEP_SET_STATUS;
      return { ok: true, info: "Note exists. Would overwrite latest note." };
    }

    const opened = await openEditorForExistingNote(latestNote, { dryRun: false });
    if (!opened.ok) return opened;
    const setRes = await setEditorHtml(noteHtml);
    if (!setRes.ok) return setRes;
    const saved = await clickSaveAndWait({ mode: "edit", dryRun: false });
    if (!saved.ok) return saved;
    wfState.step = STEP_SET_STATUS;
    return { ok: true, info: "Overwrote existing note." };
  }

  async function stepSetStatus({ dryRun }) {
    const statusText = getCurrentStatusText();
    if (!statusText) return { ok: false, info: "Cannot find status button; is Call Details open?" };
    if (statusText === TARGET_STATUS) {
      wfState.step = STEP_MARK_DONE;
      return { ok: true, info: `Status already ${TARGET_STATUS}.` };
    }
    const plan = await setCallStatusByLabel(TARGET_STATUS, { dryRun });
    if (!plan.ok) return { ok: false, info: plan.error || "Failed to set status." };
    wfState.step = STEP_MARK_DONE;
    return { ok: true, info: `Status set to ${TARGET_STATUS}.` };
  }

  async function stepMarkDone({ dryRun }) {
    const idx = wfState.rowIndex;
    if (idx == null || !queueRows[idx]) {
      wfState.step = STEP_FINISHED;
      return { ok: false, info: "No row index in state." };
    }
    const row = queueRows[idx];
    if (dryRun) {
      wfState.step = STEP_FINISHED;
      return { ok: true, info: `Would mark row #${row.rowNumber} as Done.` };
    }
    row.done = true;
    wfState.step = STEP_FINISHED;
    // Fire-and-forget webhook (errors already logged)
    postDoneWebhook(row);
    saveState();
    return { ok: true, info: `Marked row #${row.rowNumber} done.` };
  }

  // -----------------------------
  // Orchestration
  // -----------------------------
  async function runOneStep({ dryRun }) {
    if (abortRequested) return { ok: false, info: "Aborted." };
    let step = wfState.step;
    if (step === STEP_FINISHED) {
      wfState.step = STEP_LOAD_ROW;
      step = STEP_LOAD_ROW;
    }
    let result;
    if (step === STEP_LOAD_ROW) result = await stepLoadRow({ dryRun });
    else if (step === STEP_NAVIGATE_AGENT) result = await stepNavigateAgent({ dryRun });
    else if (step === STEP_FIND_OR_CREATE) result = await stepFindOrCreate({ dryRun });
    else if (step === STEP_OPEN_DETAILS) result = await stepOpenDetails({ dryRun });
    else if (step === STEP_APPLY_NOTES) result = await stepApplyNotes({ dryRun });
    else if (step === STEP_SET_STATUS) result = await stepSetStatus({ dryRun });
    else if (step === STEP_MARK_DONE) result = await stepMarkDone({ dryRun });
    else result = { ok: false, info: `Unknown step: ${step}` };
    log(`${step} ${dryRun ? "(DRY)" : "(LIVE)"}: ${result.info || ""}`, result.ok ? "info" : "error");
    return result;
  }

  function guardRunning(label) {
    if (isRunning) {
      log(`Already running (${currentMode}); ignoring ${label}.`, "warn");
      return false;
    }
    return true;
  }

  async function runDryWorkflow() {
    if (!guardRunning("Dry Run")) return;
    if (!queueRows.length) { log("No queue loaded yet. Click Load from Sheet first.", "warn"); return; }
    isRunning = true; abortRequested = false; currentMode = "dry"; updateRunningUi();
    const originalState = JSON.parse(JSON.stringify(wfState));
    try {
      wfState.step = STEP_LOAD_ROW;
      for (const s of stepOrder) {
        if (abortRequested) { log("Dry run aborted.", "warn"); break; }
        wfState.step = s;
        const res = await runOneStep({ dryRun: true });
        if (!res.ok) break;
        if (s === STEP_MARK_DONE) break;
      }
      log("Dry run complete.", "ok");
    } finally {
      wfState = originalState;
      isRunning = false; abortRequested = false; currentMode = null; updateRunningUi();
    }
  }

  async function runFullWorkflow({ resume = false } = {}) {
    if (!resume && !guardRunning("Full Auto")) return;
    if (!queueRows.length) { log("No queue loaded yet.", "warn"); return; }

    isRunning = true; abortRequested = false; currentMode = "full"; updateRunningUi(); saveState();

    let navigatingAway = false;
    try {
      let resuming = resume;
      while (true) {
        if (abortRequested) { log("Aborted.", "warn"); return; }

        let startStep;
        if (resuming) {
          startStep = wfState.step || STEP_LOAD_ROW;
          resuming = false;
          log(`Resuming at ${startStep} for row #${queueRows[wfState.rowIndex]?.rowNumber}`, "info");
        } else {
          const nextIdx = getNextPendingRowIndex();
          if (nextIdx === null) {
            log("Full Auto finished: no more pending rows.", "ok");
            return;
          }
          wfState = freshWfState();
          startStep = STEP_LOAD_ROW;
        }

        const startIdx = stepOrder.indexOf(startStep);
        const stepsToRun = startIdx >= 0 ? stepOrder.slice(startIdx) : stepOrder;

        for (const s of stepsToRun) {
          if (abortRequested) return;
          wfState.step = s;
          saveState();
          const res = await runOneStep({ dryRun: false });
          if (!res.ok) {
            log(`Full Auto stopped at ${s} for row #${queueRows[wfState.rowIndex]?.rowNumber}: ${res.info || "Error"}`, "error");
            return;
          }
          if (res.navigating) {
            // Page is about to unload. Do NOT clear running state — next load resumes.
            navigatingAway = true;
            return;
          }
          if (s === STEP_MARK_DONE) break;
        }
      }
    } finally {
      if (!navigatingAway) {
        isRunning = false; abortRequested = false; currentMode = null;
        saveState();
        updateRunningUi();
      }
    }
  }

  async function runStepWorkflow() {
    if (!guardRunning("Step")) return;
    if (!queueRows.length) { log("No queue loaded yet.", "warn"); return; }
    isRunning = true; abortRequested = false; currentMode = "step"; updateRunningUi();
    try {
      const prevStep = wfState.step;
      const res = await runOneStep({ dryRun: false });
      log(`Step ${prevStep} done: ${res.info || ""}`, res.ok ? "ok" : "error");
      if (wfState.step === STEP_FINISHED) log("Row complete.", "ok");
    } finally {
      isRunning = false; abortRequested = false; currentMode = null; saveState(); updateRunningUi();
    }
  }

  function requestAbort() {
    if (!isRunning) { log("Nothing running to abort.", "warn"); return; }
    abortRequested = true;
    log("Abort requested; will stop at next checkpoint.", "warn");
  }

  // -----------------------------
  // UI
  // -----------------------------
  let uiRefs = {};

  function updateRunningUi() {
    if (!uiRefs.abortBtn) return;
    uiRefs.abortBtn.disabled = !isRunning;
    uiRefs.abortBtn.style.opacity = isRunning ? "1" : "0.5";
    const runButtons = [uiRefs.dryBtn, uiRefs.fullBtn, uiRefs.stepBtn, uiRefs.loadBtn, uiRefs.hardResetBtn];
    for (const b of runButtons) {
      if (!b) continue;
      b.disabled = isRunning;
      b.style.opacity = isRunning ? "0.5" : "1";
    }
    if (uiRefs.statusEl) {
      uiRefs.statusEl.textContent = isRunning
        ? `Running (${currentMode})…`
        : `Idle · ${queueRows.length} rows, ${pendingCount()} pending`;
    }
  }

  function openConfigDialog() {
    const keys = ["SHEET_ID", "QUEUE_SHEET_GID", "CALL_TYPE", "DONE_WEBHOOK_URL"];
    for (const k of keys) {
      const current = cfg(k);
      const next = prompt(`${k} (leave blank to use default "${DEFAULTS[k]}")`, current ?? "");
      if (next === null) return; // cancelled
      setCfg(k, next);
    }
    log("Config updated.", "ok");
  }

  function addControlButtons() {
    const old = document.getElementById("patchWFControls");
    if (old) old.remove();

    const box = document.createElement("div");
    box.id = "patchWFControls";
    box.style.cssText =
      "position:fixed;bottom:20px;right:20px;z-index:99999;display:flex;flex-direction:column;gap:4px;background:#fff;border:1px solid #ccc;padding:6px 8px;border-radius:4px;font-family:sans-serif;font-size:12px;min-width:220px;box-shadow:0 2px 8px rgba(0,0,0,0.1);";

    const header = document.createElement("div");
    header.style.cssText =
      "display:flex;align-items:center;justify-content:space-between;cursor:move;margin-bottom:4px;";

    const title = document.createElement("span");
    title.textContent = `Patch WF ${VERSION}`;
    title.style.fontWeight = "600";

    const toggle = document.createElement("button");
    toggle.textContent = "+";
    toggle.style.cssText =
      "width:20px;height:20px;border:none;border-radius:3px;background:#999;color:#fff;cursor:pointer;font-size:14px;line-height:18px;padding:0;";

    header.appendChild(title);
    header.appendChild(toggle);
    box.appendChild(header);

    const inner = document.createElement("div");
    inner.id = "patchWFControlsInner";
    inner.style.display = "none";
    inner.style.flexDirection = "column";
    inner.style.gap = "4px";
    box.appendChild(inner);

    const statusEl = document.createElement("div");
    statusEl.style.cssText = "font-size:11px;color:#555;padding:2px 0;";
    statusEl.textContent = "Idle";
    inner.appendChild(statusEl);
    uiRefs.statusEl = statusEl;

    function makeButton(label, bg, handler) {
      const btn = document.createElement("button");
      btn.textContent = label;
      btn.style.cssText = `padding:6px 10px;margin:0;border:none;border-radius:4px;cursor:pointer;font-size:12px;color:#fff;background:${bg};`;
      btn.onclick = handler;
      return btn;
    }

    uiRefs.loadBtn = makeButton("Load from Sheet", "#2196F3", () => {
      loadQueueFromSheet().catch((e) => { log(`Error loading sheet: ${e.message}`, "error"); });
    });
    uiRefs.hardResetBtn = makeButton("Hard Reset (reload sheet)", "#E91E63", () => {
      queueRows = [];
      wfState = freshWfState();
      wfRowEl = null;
      clearSavedState();
      loadQueueFromSheet().catch((e) => { log(`Error loading sheet: ${e.message}`, "error"); });
    });
    uiRefs.dryBtn = makeButton("Dry Run (next row)", "#607D8B", () => { runDryWorkflow(); });
    uiRefs.fullBtn = makeButton("Full Auto (next rows)", "#4CAF50", () => { runFullWorkflow(); });
    uiRefs.stepBtn = makeButton("Step (advance 1 step)", "#FF9800", () => { runStepWorkflow(); });
    uiRefs.abortBtn = makeButton("Abort", "#d32f2f", () => { requestAbort(); });
    uiRefs.abortBtn.disabled = true;
    uiRefs.abortBtn.style.opacity = "0.5";

    const resetRetryBtn = makeButton("Reset (retry same row)", "#9E9E9E", () => {
      wfState = freshWfState();
      wfRowEl = null;
      saveState();
      log("State reset; next run reloads the same pending row.", "info");
      updateRunningUi();
    });
    const resetSkipBtn = makeButton("Reset (skip this row)", "#795548", () => {
      if (wfState.rowIndex != null && queueRows[wfState.rowIndex]) {
        queueRows[wfState.rowIndex].done = true;
      }
      wfState = freshWfState();
      wfRowEl = null;
      saveState();
      log("Current row marked done (in memory); next run moves to the following row.", "info");
      updateRunningUi();
    });
    const configBtn = makeButton("Config…", "#455A64", () => { openConfigDialog(); });

    inner.appendChild(uiRefs.loadBtn);
    inner.appendChild(uiRefs.hardResetBtn);
    inner.appendChild(uiRefs.dryBtn);
    inner.appendChild(uiRefs.fullBtn);
    inner.appendChild(uiRefs.stepBtn);
    inner.appendChild(uiRefs.abortBtn);
    inner.appendChild(resetRetryBtn);
    inner.appendChild(resetSkipBtn);
    inner.appendChild(configBtn);

    // Log panel
    const logsWrap = document.createElement("div");
    logsWrap.style.cssText = "margin-top:6px;border-top:1px solid #eee;padding-top:4px;";
    const logsLabel = document.createElement("div");
    logsLabel.textContent = "Log";
    logsLabel.style.cssText = "font-size:10px;color:#888;margin-bottom:2px;";
    const logs = document.createElement("div");
    logs.id = "patchWFLogs";
    logs.style.cssText = "max-height:160px;overflow-y:auto;font-family:ui-monospace,Menlo,monospace;font-size:11px;background:#fafafa;border:1px solid #eee;padding:4px;border-radius:3px;";
    const clearLogsBtn = makeButton("Clear log", "#9E9E9E", () => { logs.innerHTML = ""; });
    clearLogsBtn.style.fontSize = "10px";
    clearLogsBtn.style.padding = "3px 6px";
    clearLogsBtn.style.marginTop = "4px";
    logsWrap.appendChild(logsLabel);
    logsWrap.appendChild(logs);
    logsWrap.appendChild(clearLogsBtn);
    inner.appendChild(logsWrap);

    document.body.appendChild(box);

    let minimized = true;
    toggle.addEventListener("click", (e) => {
      e.stopPropagation();
      minimized = !minimized;
      inner.style.display = minimized ? "none" : "flex";
      toggle.textContent = minimized ? "+" : "–";
    });

    // Drag (clamped to viewport)
    let isDragging = false;
    let startX = 0, startY = 0, startLeft = 0, startTop = 0;
    function ensureAbsolutePosition() {
      const rect = box.getBoundingClientRect();
      box.style.top = rect.top + "px";
      box.style.left = rect.left + "px";
      box.style.bottom = "auto";
      box.style.right = "auto";
    }
    header.addEventListener("mousedown", (e) => {
      isDragging = true;
      ensureAbsolutePosition();
      startX = e.clientX; startY = e.clientY;
      const rect = box.getBoundingClientRect();
      startLeft = rect.left; startTop = rect.top;
      document.addEventListener("mousemove", onDrag);
      document.addEventListener("mouseup", onStopDrag);
      e.preventDefault();
    });
    function onDrag(e) {
      if (!isDragging) return;
      const rect = box.getBoundingClientRect();
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;
      const newLeft = Math.max(0, Math.min(window.innerWidth - rect.width, startLeft + dx));
      const newTop = Math.max(0, Math.min(window.innerHeight - rect.height, startTop + dy));
      box.style.left = newLeft + "px";
      box.style.top = newTop + "px";
    }
    function onStopDrag() {
      isDragging = false;
      document.removeEventListener("mousemove", onDrag);
      document.removeEventListener("mouseup", onStopDrag);
    }

    updateRunningUi();
  }

  // -----------------------------
  // Init + auto-resume
  // -----------------------------
  window.PatchWF = {
    get version() { return VERSION; },
    get queueRows() { return queueRows; },
    get wfState() { return wfState; },
    get isRunning() { return isRunning; },
    abort: requestAbort,
    clearSavedState,
  };
  addControlButtons();
  log(`Loaded ${VERSION}`, "info");

  const saved = loadSavedState();
  if (saved) {
    if (Array.isArray(saved.queueRows) && saved.queueRows.length) {
      queueRows = saved.queueRows;
    }
    if (saved.wfState && typeof saved.wfState === "object") {
      wfState = { ...freshWfState(), ...saved.wfState };
    }
    updateRunningUi();
    log(`Restored ${queueRows.length} rows from local storage, ${pendingCount()} pending.`, "info");
    if (saved.running && saved.mode === "full") {
      log("Resuming Full Auto after page reload...", "info");
      (async () => {
        // Wait for the dashboard calls page to settle before resuming,
        // so stepFindOrCreate can see the grid.
        await sleep(1200);
        await waitFor(() => isDashboardCallsPage() && getDashboardRows().length > 0, {
          timeoutMs: 20000,
          intervalMs: 300,
        });
        runFullWorkflow({ resume: true });
      })();
    }
  }
})();
