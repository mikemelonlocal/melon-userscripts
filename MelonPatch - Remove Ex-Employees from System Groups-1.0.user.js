// ==UserScript==
// @name         MelonPatch - Remove Ex-Employees from System Groups
// @namespace    http://tampermonkey.net/
// @version      2.2.0
// @description  Highlights ex-employees on System Group pages, plus inline panels to bulk-remove ex-employees and bulk-add active Melons. Themed and structured to match Patch Targeting Helper.
// @match        https://thepatch.melonlocal.com/SystemGroups/Edit*
// @run-at       document-end
// @grant        none
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/MelonPatch%20-%20Remove%20Ex-Employees%20from%20System%20Groups-1.0.user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/MelonPatch%20-%20Remove%20Ex-Employees%20from%20System%20Groups-1.0.user.js
// ==/UserScript==

(function () {
  "use strict";

  // ============================================================
  // CONSTANTS
  // ============================================================

  // Keep in sync with @version in the userscript header above.
  const VERSION = "patch-ex-employees-v2.2.0";
  const DEBUG = false;

  const _debug = { DEBUG_MODE: DEBUG };

  const TIMING = {
    DEFAULT_DELAY_MS: 120,
    MAX_WAIT_ITERATIONS: 200,
    WAIT_ITERATION_DELAY_MS: 150,
    FOCUS_DELAY_MS: 0,
    RETRY_BASE_BACKOFF_MS: 200,
    RETRY_MAX_ATTEMPTS: 2, // total attempts after the first call = 2 retries
    // Safety-net polling for re-injection after the page rebuilds the chip
    // list (e.g., after a manual add/remove). Idempotent, so cheap to run.
    REINJECT_INTERVAL_MS: 1500,
  };

  const COLORS = {
    alpine: "#FEF8E9",
    cactus: "#47B74F",
    lemonSun: "#F1CB20",
    sand: "#EDDFDB",
    clover: "#40A74C",
    mustardSeed: "#CC8F15",
    whitneyPink: "#FF9B94",
    watermelonSugar: "#E9736E",
    mojave: "#CFBA97",
    pine: "#114E38",
    coconut: "#644414",
    cranberry: "#6C2126",
  };

  const PANEL_IDS = {
    inlinePanelClass: "patch-inline-panel",
    inlinePanelStatusClass: "patch-inline-status",
    panelId: "patch-ex-employee-panel",
    triggerBtnId: "patchExEmployeeBtn",
    chipMarkClass: "patch-ex-employee-chip",
    addPanelId: "patch-add-members-panel",
    addTriggerBtnId: "patchAddMembersBtn",
  };

  const ENDPOINTS = {
    REMOVE_MEMBER: "/SystemGroups/Edit?handler=RemoveTargetSystemGroupMember",
    ADD_MEMBER: "/SystemGroups/Edit",
    GET_MELONS: "/Lists/GetMelonVms",
  };

  const MELON_NAME_FIELDS = ["FullName", "DisplayName", "Name", "FirstName", "LastName", "Email"];

  function melonDisplayName(m) {
    if (!m) return "";
    if (m.FullName) return String(m.FullName);
    if (m.DisplayName) return String(m.DisplayName);
    const fn = m.FirstName || m.firstName || "";
    const ln = m.LastName || m.lastName || "";
    const combined = `${fn} ${ln}`.trim();
    if (combined) return combined;
    if (m.Name) return String(m.Name);
    if (m.Email) return String(m.Email);
    return `MelonId ${m.MelonId}`;
  }

  function melonMatchesQuery(m, qLower) {
    if (!qLower) return true;
    if (m.MelonId != null && String(m.MelonId).includes(qLower)) return true;
    for (const k of MELON_NAME_FIELDS) {
      const v = m[k];
      if (v && String(v).toLowerCase().includes(qLower)) return true;
    }
    if (m.FirstName && m.LastName) {
      if (`${m.FirstName} ${m.LastName}`.toLowerCase().includes(qLower)) return true;
    }
    return false;
  }

  // ============================================================
  // UTILITY FUNCTIONS
  // ============================================================

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
      clearTimeout(timeout);
      timeout = setTimeout(() => func(...args), wait);
    };
  }

  function normalizeText(s) {
    return String(s || "").replace(/ /g, " ").replace(/\s+/g, " ").trim();
  }

  function insertAfter(newNode, referenceNode) {
    const parent = referenceNode?.parentNode;
    if (!parent) return;
    parent.insertBefore(newNode, referenceNode.nextSibling);
  }

  function log(message, ...args) {
    if (DEBUG || _debug.DEBUG_MODE) {
      console.log(`[PatchExEmployees] ${message}`, ...args);
    }
  }

  function logError(message, ...args) {
    console.error(`[PatchExEmployees] ${message}`, ...args);
  }

  // ============================================================
  // PAGE DETECTION
  // ============================================================

  const PageDetector = {
    get hrefLower() { return String(location.href || "").toLowerCase(); },
    get hostLower() { return String(location.hostname || "").toLowerCase(); },

    get isPatch() {
      return this.hostLower === "thepatch.melonlocal.com";
    },
    get isSystemGroupEdit() {
      return this.isPatch && this.hrefLower.includes("/systemgroups/edit");
    },
  };

  // ============================================================
  // GLOBAL STYLES (inline panel theming)
  // ============================================================

  function injectGlobalStyles() {
    if (document.getElementById("patch-ex-employees-styles")) return;
    const styleTag = document.createElement("style");
    styleTag.id = "patch-ex-employees-styles";
    styleTag.textContent = `
      .${PANEL_IDS.inlinePanelClass} {
        margin-top: 12px;
        padding: 12px 14px;
        background: ${COLORS.alpine};
        border: 1px solid ${COLORS.mojave};
        border-radius: 8px;
        font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
        color: ${COLORS.coconut};
        box-shadow: 0 1px 2px rgba(100, 68, 20, 0.08);
      }
      .${PANEL_IDS.inlinePanelClass}[hidden] { display: none !important; }
      .${PANEL_IDS.inlinePanelClass}-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 8px;
      }
      .${PANEL_IDS.inlinePanelClass}-title {
        font-size: 14px;
        font-weight: 600;
        color: ${COLORS.coconut};
      }
      .${PANEL_IDS.inlinePanelClass}-close {
        background: none; border: none; cursor: pointer;
        font-size: 20px; line-height: 1; color: ${COLORS.coconut}; padding: 0 6px;
      }
      .${PANEL_IDS.inlinePanelClass}-close:hover { color: ${COLORS.cranberry}; }
      .${PANEL_IDS.inlinePanelStatusClass} {
        margin-top: 6px;
        font-size: 12px;
        min-height: 16px;
        color: ${COLORS.coconut};
      }
      .${PANEL_IDS.inlinePanelStatusClass}.is-error { color: ${COLORS.cranberry}; }
      .${PANEL_IDS.inlinePanelStatusClass}.is-info  { color: ${COLORS.coconut}; }
      .${PANEL_IDS.inlinePanelStatusClass}.is-warn  { color: ${COLORS.mustardSeed}; }
      .${PANEL_IDS.inlinePanelStatusClass} .badge {
        display: inline-block;
        padding: 1px 6px;
        border-radius: 10px;
        background: ${COLORS.sand};
        color: ${COLORS.coconut};
        margin-left: 4px;
        font-weight: 600;
      }
      .${PANEL_IDS.inlinePanelClass}-button-row {
        display: flex; gap: 6px; justify-content: flex-end; margin-top: 10px;
        flex-wrap: wrap;
      }
      .patch-inline-btn {
        border: none;
        padding: 7px 14px;
        border-radius: 4px;
        cursor: pointer;
        font-size: 13px;
        font-weight: 500;
        transition: background-color 120ms ease;
      }
      .patch-inline-btn[disabled] { cursor: progress; opacity: 0.85; }
      .patch-inline-btn--primary   { background: ${COLORS.cactus};         color: #fff; }
      .patch-inline-btn--secondary { background: ${COLORS.sand};           color: ${COLORS.coconut}; }
      .patch-inline-btn--danger    { background: ${COLORS.watermelonSugar}; color: #fff; }
      .patch-inline-btn--busy      { background: ${COLORS.mustardSeed} !important; color: #fff !important; }
      .patch-ex-employee-list {
        margin-top: 4px;
        padding: 8px;
        background: #fff;
        border: 1px solid ${COLORS.mojave};
        border-radius: 4px;
        max-height: 180px;
        overflow-y: auto;
        font-size: 12px;
        color: ${COLORS.coconut};
      }
      .patch-ex-employee-list-empty {
        color: ${COLORS.mustardSeed};
        font-style: italic;
      }
      .patch-ex-employee-list-item {
        padding: 2px 0;
        font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
      }
      .${PANEL_IDS.chipMarkClass} {
        background-color: ${COLORS.whitneyPink} !important;
        border-color: ${COLORS.watermelonSugar} !important;
        color: ${COLORS.cranberry} !important;
      }
      .patch-add-members-search {
        width: 100%;
        box-sizing: border-box;
        padding: 7px 10px;
        font-size: 13px;
        border: 1px solid ${COLORS.mojave};
        border-radius: 4px;
        margin-bottom: 6px;
        background: #fff;
        color: ${COLORS.coconut};
      }
      .patch-add-members-search:focus {
        outline: none;
        border-color: ${COLORS.cactus};
        box-shadow: 0 0 0 2px rgba(71, 183, 79, 0.15);
      }
      .patch-add-members-list {
        margin-top: 4px;
        padding: 4px;
        background: #fff;
        border: 1px solid ${COLORS.mojave};
        border-radius: 4px;
        max-height: 280px;
        overflow-y: auto;
        font-size: 13px;
        color: ${COLORS.coconut};
      }
      .patch-add-members-list-item {
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 4px 6px;
        border-radius: 3px;
        cursor: pointer;
        user-select: none;
      }
      .patch-add-members-list-item:hover {
        background: ${COLORS.alpine};
      }
      .patch-add-members-list-item input[type="checkbox"] {
        margin: 0;
        cursor: pointer;
        accent-color: ${COLORS.cactus};
      }
      .patch-add-members-list-item-name {
        flex: 1;
      }
      .patch-add-members-list-item-id {
        color: ${COLORS.mustardSeed};
        font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
        font-size: 11px;
      }
      .patch-add-members-list-empty {
        padding: 8px;
        color: ${COLORS.mustardSeed};
        font-style: italic;
      }
    `;
    document.head.appendChild(styleTag);
  }

  // ============================================================
  // API CLIENT — direct calls to the page's backend endpoints
  // ============================================================
  // Endpoint contract:
  //   GET /Lists/GetMelonVms                 -> [{ MelonId: number, ... }]
  //   POST /SystemGroups/Edit?handler=RemoveTargetSystemGroupMember
  //     Content-Type: application/json; charset=UTF-8
  //     RequestVerificationToken: <from hidden input>
  //     Body: {"SystemGroupId": <num>, "MelonId": <num>}

  const ApiClient = {
    getAntiforgeryToken() {
      const input = document.querySelector('input[name="__RequestVerificationToken"]');
      return input?.value || "";
    },

    getSystemGroupId() {
      const hidden = document.querySelector('input[name="systemGroupId"]');
      if (hidden?.value) return String(hidden.value);
      const fromUrl = new URLSearchParams(window.location.search).get("systemGroupId");
      return fromUrl ? String(fromUrl) : null;
    },

    async getActiveMelons() {
      const res = await fetch(ENDPOINTS.GET_MELONS, { credentials: "same-origin" });
      if (!res.ok) {
        throw new Error(`HTTP ${res.status} ${res.statusText} fetching active melons`);
      }
      return await res.json();
    },

    async getActiveMelonIds() {
      const melons = await this.getActiveMelons();
      return new Set(melons.map((m) => m.MelonId));
    },

    /**
     * POST add-member with retry + exponential backoff. Mirrors the page's
     * own AddNewSystemGroupMember function: bare /SystemGroups/Edit URL,
     * JSON body {SystemGroupId, MelonId}, anti-forgery header.
     */
    async addMember({ systemGroupId, melonId, token }, opts = {}) {
      const maxRetries = opts.maxRetries ?? TIMING.RETRY_MAX_ATTEMPTS;
      const baseBackoff = opts.baseBackoffMs ?? TIMING.RETRY_BASE_BACKOFF_MS;
      const body = JSON.stringify({
        SystemGroupId: Number(systemGroupId),
        MelonId: Number(melonId),
      });

      let lastError;
      for (let attempt = 0; attempt <= maxRetries; attempt++) {
        try {
          const response = await fetch(ENDPOINTS.ADD_MEMBER, {
            method: "POST",
            credentials: "same-origin",
            headers: {
              "Content-Type": "application/json; charset=utf-8",
              "X-Requested-With": "XMLHttpRequest",
              RequestVerificationToken: token,
            },
            body,
          });
          if (!response.ok) {
            if (response.status >= 500 || response.status === 429) {
              throw new Error(`HTTP ${response.status} ${response.statusText}`);
            }
            throw Object.assign(
              new Error(`HTTP ${response.status} ${response.statusText}`),
              { terminal: true }
            );
          }
          return;
        } catch (error) {
          lastError = error;
          if (error?.terminal || attempt === maxRetries) break;
          const backoff = baseBackoff * Math.pow(2, attempt);
          log(`addMember retry ${attempt + 1}/${maxRetries} for MelonId ${melonId} after ${backoff}ms`, error?.message);
          await sleep(backoff);
        }
      }
      throw lastError ?? new Error("Unknown error in addMember");
    },

    /**
     * POST RemoveTargetSystemGroupMember with retry + exponential backoff.
     * Throws if all attempts fail.
     */
    async removeMember({ systemGroupId, melonId, token }, opts = {}) {
      const maxRetries = opts.maxRetries ?? TIMING.RETRY_MAX_ATTEMPTS;
      const baseBackoff = opts.baseBackoffMs ?? TIMING.RETRY_BASE_BACKOFF_MS;
      const body = JSON.stringify({
        SystemGroupId: Number(systemGroupId),
        MelonId: Number(melonId),
      });

      let lastError;
      for (let attempt = 0; attempt <= maxRetries; attempt++) {
        try {
          const response = await fetch(ENDPOINTS.REMOVE_MEMBER, {
            method: "POST",
            credentials: "same-origin",
            headers: {
              "Content-Type": "application/json; charset=UTF-8",
              "X-Requested-With": "XMLHttpRequest",
              RequestVerificationToken: token,
            },
            body,
          });
          if (!response.ok) {
            // Treat 5xx and 429 as retryable; 4xx (other) as terminal.
            if (response.status >= 500 || response.status === 429) {
              throw new Error(`HTTP ${response.status} ${response.statusText}`);
            }
            throw Object.assign(
              new Error(`HTTP ${response.status} ${response.statusText}`),
              { terminal: true }
            );
          }
          return; // success
        } catch (error) {
          lastError = error;
          if (error?.terminal || attempt === maxRetries) break;
          const backoff = baseBackoff * Math.pow(2, attempt);
          log(`removeMember retry ${attempt + 1}/${maxRetries} for MelonId ${melonId} after ${backoff}ms`, error?.message);
          await sleep(backoff);
        }
      }
      throw lastError ?? new Error("Unknown error in removeMember");
    },
  };

  // ============================================================
  // EX-EMPLOYEE OPERATIONS
  // ============================================================

  const ExEmployeeOperations = {
    /**
     * Scan the Kendo chip list, mark ex-employee chips with our CSS class,
     * and return the list of ex-employees.
     */
    async scan() {
      // Try Kendo's class first; fall back to anything carrying data-chip-id
      // inside the container, in case the page swaps to a non-Kendo render.
      const container = document.getElementById("systemGroupMembersChipList");
      let chips = container
        ? Array.from(container.querySelectorAll(".k-chip[data-chip-id], [data-chip-id]"))
        : [];
      if (!chips.length) {
        return { chips: [], exEmployees: [], reason: "no-chips" };
      }
      let activeMelonIds;
      try {
        activeMelonIds = await ApiClient.getActiveMelonIds();
      } catch (error) {
        logError("Failed to load active melons:", error);
        return { chips, exEmployees: [], reason: "fetch-failed", error };
      }
      const exEmployees = [];
      for (const chip of chips) {
        const melonId = Number(chip.getAttribute("data-chip-id"));
        if (!Number.isFinite(melonId)) continue;
        if (activeMelonIds.has(melonId)) {
          chip.classList.remove(PANEL_IDS.chipMarkClass);
          if (chip.title === "Ex-employee (not in active Melons list)") {
            chip.removeAttribute("title");
          }
          continue;
        }
        chip.classList.add(PANEL_IDS.chipMarkClass);
        chip.title = "Ex-employee (not in active Melons list)";
        const name = normalizeText(chip.textContent) || `MelonId ${melonId}`;
        exEmployees.push({ chip, melonId, name });
      }
      log("Scan complete", { total: chips.length, exEmployees: exEmployees.length });
      return { chips, exEmployees, reason: "ok" };
    },

    /**
     * Remove ex-employees one at a time via the API. opts.onProgress({completed,total,phase})
     * fires after each item.
     */
    async run(exEmployees, delayMs = TIMING.DEFAULT_DELAY_MS, opts = {}) {
      if (!exEmployees.length) {
        return { ok: false, reason: "empty", message: "Nothing to remove." };
      }
      const token = ApiClient.getAntiforgeryToken();
      const systemGroupId = ApiClient.getSystemGroupId();
      if (!token || !systemGroupId) {
        logError("API remove unavailable", { hasToken: !!token, hasSystemGroupId: !!systemGroupId });
        return {
          ok: false,
          reason: "no-api",
          message: "Cannot remove via API: missing anti-forgery token or systemGroupId.",
        };
      }
      let removed = 0;
      const failed = [];
      for (let i = 0; i < exEmployees.length; i++) {
        const entry = exEmployees[i];
        try {
          await ApiClient.removeMember({ systemGroupId, melonId: entry.melonId, token });
          try { entry.chip.remove(); } catch (_) {}
          removed++;
          await sleep(delayMs);
        } catch (error) {
          logError(`API remove failed for MelonId ${entry.melonId}:`, error?.message || error);
          failed.push(entry);
        }
        if (opts.onProgress) opts.onProgress({ completed: i + 1, total: exEmployees.length, phase: "Removing" });
      }
      return {
        ok: true,
        completed: removed,
        total: exEmployees.length,
        failed,
        reload: removed > 0,
        message: `Removed ${removed} of ${exEmployees.length} ex-employee(s)` +
          (failed.length ? `  •  Failed: ${failed.length}` : ""),
      };
    },
  };

  // ============================================================
  // ADD-MEMBER OPERATIONS
  // ============================================================

  const AddMemberOperations = {
    /**
     * Read MelonIds for everyone currently in the chip list, so the picker
     * can hide them.
     */
    getCurrentMemberIds() {
      const ids = new Set();
      const container = document.getElementById("systemGroupMembersChipList");
      if (!container) return ids;
      const chips = container.querySelectorAll(".k-chip[data-chip-id], [data-chip-id]");
      for (const chip of chips) {
        const id = Number(chip.getAttribute("data-chip-id"));
        if (Number.isFinite(id)) ids.add(id);
      }
      return ids;
    },

    /**
     * Pull all active melons and remove anyone already a member.
     */
    async getCandidates() {
      const all = await ApiClient.getActiveMelons();
      const current = this.getCurrentMemberIds();
      const candidates = all
        .filter((m) => Number.isFinite(Number(m?.MelonId)))
        .filter((m) => !current.has(Number(m.MelonId)));
      candidates.sort((a, b) => melonDisplayName(a).localeCompare(melonDisplayName(b)));
      log("Add-member candidates", { total: all.length, alreadyMembers: current.size, candidates: candidates.length });
      return candidates;
    },

    /**
     * Add each selected melon. opts.onProgress({completed,total,phase}) fires
     * after each item.
     */
    async run(melons, delayMs = TIMING.DEFAULT_DELAY_MS, opts = {}) {
      if (!melons.length) {
        return { ok: false, reason: "empty", message: "Nothing to add." };
      }
      const token = ApiClient.getAntiforgeryToken();
      const systemGroupId = ApiClient.getSystemGroupId();
      if (!token || !systemGroupId) {
        logError("API add unavailable", { hasToken: !!token, hasSystemGroupId: !!systemGroupId });
        return {
          ok: false,
          reason: "no-api",
          message: "Cannot add via API: missing anti-forgery token or systemGroupId.",
        };
      }
      let added = 0;
      const failed = [];
      for (let i = 0; i < melons.length; i++) {
        const m = melons[i];
        try {
          await ApiClient.addMember({ systemGroupId, melonId: m.MelonId, token });
          added++;
          await sleep(delayMs);
        } catch (error) {
          logError(`API add failed for MelonId ${m.MelonId}:`, error?.message || error);
          failed.push(m);
        }
        if (opts.onProgress) opts.onProgress({ completed: i + 1, total: melons.length, phase: "Adding" });
      }
      return {
        ok: true,
        completed: added,
        total: melons.length,
        failed,
        reload: added > 0,
        message: `Added ${added} of ${melons.length} member(s)` +
          (failed.length ? `  •  Failed: ${failed.length}` : ""),
      };
    },
  };

  // ============================================================
  // INLINE DASHBOARD HELPERS
  // ============================================================

  const InlineDashboard = {
    makeButton(text, variant = "primary") {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = text;
      btn.className = `patch-inline-btn patch-inline-btn--${variant}`;
      btn.dataset.originalText = text;
      btn.dataset.originalVariant = variant;
      return btn;
    },

    setBusy(btn, busy, label) {
      if (busy) {
        btn.disabled = true;
        btn.classList.add("patch-inline-btn--busy");
        if (label != null) btn.textContent = label;
      } else {
        btn.disabled = false;
        btn.classList.remove("patch-inline-btn--busy");
        btn.textContent = btn.dataset.originalText || btn.textContent;
      }
    },
  };

  // ============================================================
  // PANEL BUILDER
  // ============================================================

  function buildExEmployeePanel() {
    const panel = document.createElement("div");
    panel.id = PANEL_IDS.panelId;
    panel.className = PANEL_IDS.inlinePanelClass;
    panel.hidden = true;
    panel.dataset.bulkPanel = "ex-employees";

    // Header
    const header = document.createElement("div");
    header.className = `${PANEL_IDS.inlinePanelClass}-header`;
    const titleEl = document.createElement("div");
    titleEl.className = `${PANEL_IDS.inlinePanelClass}-title`;
    titleEl.textContent = "Remove Ex-Employees";
    const closeBtn = document.createElement("button");
    closeBtn.type = "button";
    closeBtn.className = `${PANEL_IDS.inlinePanelClass}-close`;
    closeBtn.setAttribute("aria-label", "Close");
    closeBtn.textContent = "×";
    header.appendChild(titleEl);
    header.appendChild(closeBtn);

    // List of detected ex-employees
    const list = document.createElement("div");
    list.className = "patch-ex-employee-list";

    // Status line
    const status = document.createElement("div");
    status.className = `${PANEL_IDS.inlinePanelStatusClass} is-info`;

    // Buttons
    const btnRow = document.createElement("div");
    btnRow.className = `${PANEL_IDS.inlinePanelClass}-button-row`;
    const cancelBtn = InlineDashboard.makeButton("Cancel", "secondary");
    const rescanBtn = InlineDashboard.makeButton("Rescan", "secondary");
    const removeBtn = InlineDashboard.makeButton("Remove Ex-Employees", "danger");
    btnRow.appendChild(cancelBtn);
    btnRow.appendChild(rescanBtn);
    btnRow.appendChild(removeBtn);

    panel.appendChild(header);
    panel.appendChild(list);
    panel.appendChild(status);
    panel.appendChild(btnRow);

    let lastScan = { exEmployees: [], reason: "pending" };

    const renderList = () => {
      list.innerHTML = "";
      if (!lastScan.exEmployees.length) {
        const empty = document.createElement("div");
        empty.className = "patch-ex-employee-list-empty";
        empty.textContent = lastScan.reason === "pending"
          ? "Click Rescan to detect ex-employees."
          : "No ex-employees detected.";
        list.appendChild(empty);
        return;
      }
      for (const entry of lastScan.exEmployees) {
        const row = document.createElement("div");
        row.className = "patch-ex-employee-list-item";
        row.textContent = `${entry.name} (MelonId ${entry.melonId})`;
        list.appendChild(row);
      }
    };

    const renderStatus = () => {
      status.classList.remove("is-error", "is-warn", "is-info");
      const onPage = document.querySelectorAll(
        "#systemGroupMembersChipList .k-chip[data-chip-id], #systemGroupMembersChipList [data-chip-id]"
      ).length;
      const count = lastScan.exEmployees.length;
      const onPageBadge = `<span class="badge">${onPage} on page</span>`;
      if (lastScan.reason === "fetch-failed") {
        status.classList.add("is-error");
        status.innerHTML = `Could not fetch active melons. ${onPageBadge}`;
        return;
      }
      if (lastScan.reason === "no-chips") {
        status.classList.add("is-warn");
        status.textContent = "No member chips visible yet.";
        return;
      }
      if (lastScan.reason === "pending") {
        status.classList.add("is-info");
        status.innerHTML = `Ready. ${onPageBadge}`;
        return;
      }
      status.classList.add("is-info");
      status.innerHTML = count
        ? `Found <strong>${count}</strong> ex-employee(s). ${onPageBadge}`
        : `No ex-employees found. ${onPageBadge}`;
    };

    const updateRemoveBtnState = () => {
      removeBtn.disabled = lastScan.exEmployees.length === 0;
    };

    const rescan = async () => {
      InlineDashboard.setBusy(rescanBtn, true, "Scanning...");
      try {
        lastScan = await ExEmployeeOperations.scan();
      } catch (error) {
        logError("Rescan failed:", error);
        lastScan = { exEmployees: [], reason: "fetch-failed", error };
      }
      InlineDashboard.setBusy(rescanBtn, false);
      renderList();
      renderStatus();
      updateRemoveBtnState();
    };

    const close = () => { panel.hidden = true; };
    cancelBtn.addEventListener("click", close);
    closeBtn.addEventListener("click", close);
    rescanBtn.addEventListener("click", () => { rescan(); });

    removeBtn.addEventListener("click", async () => {
      if (removeBtn.disabled) return;
      const count = lastScan.exEmployees.length;
      if (!count) return;
      if (!confirm(`Remove ${count} ex-employee(s) from this system group? This cannot be undone.`)) return;

      cancelBtn.disabled = true;
      rescanBtn.disabled = true;
      InlineDashboard.setBusy(removeBtn, true, `(0/${count}) Removing...`);
      const onProgress = ({ completed, total }) => {
        removeBtn.textContent = `(${completed}/${total}) Removing...`;
      };

      const result = await ExEmployeeOperations.run(
        lastScan.exEmployees,
        TIMING.DEFAULT_DELAY_MS,
        { onProgress }
      );
      InlineDashboard.setBusy(removeBtn, false);
      cancelBtn.disabled = false;
      rescanBtn.disabled = false;

      status.classList.remove("is-info", "is-warn", "is-error");
      if (result.ok) {
        status.classList.add("is-info");
        status.textContent = result.message;
        if (result.reload) {
          status.textContent += "  •  Reloading page...";
          setTimeout(() => location.reload(), 800);
        } else {
          rescan();
        }
      } else {
        status.classList.add("is-warn");
        status.textContent = result.message || "No-op.";
      }
    });

    // Initial empty render
    renderList();
    renderStatus();
    updateRemoveBtnState();

    return {
      panel,
      rescan,
      show: async () => {
        panel.hidden = false;
        await rescan();
      },
      hide: close,
      toggle: async () => {
        if (panel.hidden) {
          panel.hidden = false;
          await rescan();
        } else {
          close();
        }
      },
      isVisible: () => !panel.hidden,
    };
  }

  // ============================================================
  // ADD-MEMBERS PANEL BUILDER
  // ============================================================

  function buildAddMembersPanel() {
    const panel = document.createElement("div");
    panel.id = PANEL_IDS.addPanelId;
    panel.className = PANEL_IDS.inlinePanelClass;
    panel.hidden = true;
    panel.dataset.bulkPanel = "add-members";

    const header = document.createElement("div");
    header.className = `${PANEL_IDS.inlinePanelClass}-header`;
    const titleEl = document.createElement("div");
    titleEl.className = `${PANEL_IDS.inlinePanelClass}-title`;
    titleEl.textContent = "Add Members";
    const closeBtn = document.createElement("button");
    closeBtn.type = "button";
    closeBtn.className = `${PANEL_IDS.inlinePanelClass}-close`;
    closeBtn.setAttribute("aria-label", "Close");
    closeBtn.textContent = "×";
    header.appendChild(titleEl);
    header.appendChild(closeBtn);

    const search = document.createElement("input");
    search.type = "text";
    search.className = "patch-add-members-search";
    search.placeholder = "Search by name, email, or MelonId...";

    const list = document.createElement("div");
    list.className = "patch-add-members-list";

    const status = document.createElement("div");
    status.className = `${PANEL_IDS.inlinePanelStatusClass} is-info`;

    const btnRow = document.createElement("div");
    btnRow.className = `${PANEL_IDS.inlinePanelClass}-button-row`;
    const cancelBtn = InlineDashboard.makeButton("Cancel", "secondary");
    const refreshBtn = InlineDashboard.makeButton("Refresh", "secondary");
    const selectVisibleBtn = InlineDashboard.makeButton("Select visible", "secondary");
    const clearBtn = InlineDashboard.makeButton("Clear", "secondary");
    const addBtn = InlineDashboard.makeButton("Add Selected", "primary");
    btnRow.appendChild(cancelBtn);
    btnRow.appendChild(refreshBtn);
    btnRow.appendChild(clearBtn);
    btnRow.appendChild(selectVisibleBtn);
    btnRow.appendChild(addBtn);

    panel.appendChild(header);
    panel.appendChild(search);
    panel.appendChild(list);
    panel.appendChild(status);
    panel.appendChild(btnRow);

    let allCandidates = [];
    let filtered = [];
    let selected = new Map(); // MelonId -> melon object
    let loadState = "pending"; // pending | ok | error

    const renderStatus = () => {
      status.classList.remove("is-error", "is-warn", "is-info");
      if (loadState === "error") {
        status.classList.add("is-error");
        status.textContent = "Could not load active Melons.";
        return;
      }
      status.classList.add("is-info");
      const sel = selected.size;
      const shown = filtered.length;
      const total = allCandidates.length;
      status.innerHTML =
        `<strong>${sel}</strong> selected ` +
        `<span class="badge">${shown} shown</span> ` +
        `<span class="badge">${total} addable</span>`;
    };

    const updateAddBtnState = () => {
      addBtn.disabled = selected.size === 0;
    };

    const renderList = () => {
      list.innerHTML = "";
      if (loadState === "pending") {
        const empty = document.createElement("div");
        empty.className = "patch-add-members-list-empty";
        empty.textContent = "Loading active Melons...";
        list.appendChild(empty);
        return;
      }
      if (loadState === "error") {
        const empty = document.createElement("div");
        empty.className = "patch-add-members-list-empty";
        empty.textContent = "Failed to load. Try Refresh.";
        list.appendChild(empty);
        return;
      }
      if (!filtered.length) {
        const empty = document.createElement("div");
        empty.className = "patch-add-members-list-empty";
        empty.textContent = allCandidates.length
          ? "No matches for that search."
          : "No melons available to add (everyone is already a member).";
        list.appendChild(empty);
        return;
      }
      const frag = document.createDocumentFragment();
      for (const m of filtered) {
        const row = document.createElement("label");
        row.className = "patch-add-members-list-item";

        const cb = document.createElement("input");
        cb.type = "checkbox";
        cb.checked = selected.has(m.MelonId);
        cb.addEventListener("change", () => {
          if (cb.checked) selected.set(m.MelonId, m);
          else selected.delete(m.MelonId);
          renderStatus();
          updateAddBtnState();
        });

        const nameEl = document.createElement("span");
        nameEl.className = "patch-add-members-list-item-name";
        nameEl.textContent = melonDisplayName(m);

        const idEl = document.createElement("span");
        idEl.className = "patch-add-members-list-item-id";
        idEl.textContent = `#${m.MelonId}`;

        row.appendChild(cb);
        row.appendChild(nameEl);
        row.appendChild(idEl);
        frag.appendChild(row);
      }
      list.appendChild(frag);
    };

    const applyFilter = () => {
      const q = (search.value || "").toLowerCase().trim();
      filtered = q
        ? allCandidates.filter((m) => melonMatchesQuery(m, q))
        : allCandidates.slice();
      renderList();
      renderStatus();
    };

    const refresh = async () => {
      loadState = "pending";
      InlineDashboard.setBusy(refreshBtn, true, "Loading...");
      renderList();
      renderStatus();
      try {
        allCandidates = await AddMemberOperations.getCandidates();
        // Drop any selections that are no longer candidates (e.g. someone
        // got added in another tab).
        const stillValid = new Set(allCandidates.map((m) => m.MelonId));
        for (const id of [...selected.keys()]) {
          if (!stillValid.has(id)) selected.delete(id);
        }
        loadState = "ok";
      } catch (error) {
        logError("Failed to load candidates:", error);
        allCandidates = [];
        loadState = "error";
      }
      InlineDashboard.setBusy(refreshBtn, false);
      applyFilter();
      updateAddBtnState();
    };

    const close = () => { panel.hidden = true; };
    cancelBtn.addEventListener("click", close);
    closeBtn.addEventListener("click", close);
    refreshBtn.addEventListener("click", () => { refresh(); });

    clearBtn.addEventListener("click", () => {
      selected.clear();
      renderList();
      renderStatus();
      updateAddBtnState();
    });

    selectVisibleBtn.addEventListener("click", () => {
      for (const m of filtered) selected.set(m.MelonId, m);
      renderList();
      renderStatus();
      updateAddBtnState();
    });

    search.addEventListener("input", debounce(applyFilter, 100));

    addBtn.addEventListener("click", async () => {
      if (addBtn.disabled) return;
      const picks = [...selected.values()];
      if (!picks.length) return;
      if (!confirm(`Add ${picks.length} member(s) to this system group?`)) return;

      cancelBtn.disabled = true;
      refreshBtn.disabled = true;
      clearBtn.disabled = true;
      selectVisibleBtn.disabled = true;
      search.disabled = true;
      InlineDashboard.setBusy(addBtn, true, `(0/${picks.length}) Adding...`);
      const onProgress = ({ completed, total }) => {
        addBtn.textContent = `(${completed}/${total}) Adding...`;
      };

      const result = await AddMemberOperations.run(picks, TIMING.DEFAULT_DELAY_MS, { onProgress });

      InlineDashboard.setBusy(addBtn, false);
      cancelBtn.disabled = false;
      refreshBtn.disabled = false;
      clearBtn.disabled = false;
      selectVisibleBtn.disabled = false;
      search.disabled = false;

      status.classList.remove("is-info", "is-warn", "is-error");
      if (result.ok) {
        status.classList.add("is-info");
        status.textContent = result.message;
        if (result.reload) {
          status.textContent += "  •  Reloading page...";
          setTimeout(() => location.reload(), 800);
        }
      } else {
        status.classList.add("is-warn");
        status.textContent = result.message || "No-op.";
      }
    });

    renderList();
    renderStatus();
    updateAddBtnState();

    return {
      panel,
      refresh,
      show: async () => {
        panel.hidden = false;
        await refresh();
        setTimeout(() => { try { search.focus(); } catch (_) {} }, TIMING.FOCUS_DELAY_MS);
      },
      hide: close,
      toggle: async () => {
        if (panel.hidden) {
          panel.hidden = false;
          await refresh();
          setTimeout(() => { try { search.focus(); } catch (_) {} }, TIMING.FOCUS_DELAY_MS);
        } else {
          close();
        }
      },
      isVisible: () => !panel.hidden,
    };
  }

  // ============================================================
  // BUTTON INJECTION
  // ============================================================

  const ButtonInjector = {
    _dashboard: null,
    _addDashboard: null,

    inject() {
      if (!PageDetector.isSystemGroupEdit) return;
      const haveExBtn = !!document.getElementById(PANEL_IDS.triggerBtnId);
      const haveAddBtn = !!document.getElementById(PANEL_IDS.addTriggerBtnId);
      if (haveExBtn && haveAddBtn) return;

      const chipList = document.getElementById("systemGroupMembersChipList");
      if (!chipList) return;

      // Build (or reuse) dashboards.
      const exDashboard = haveExBtn ? this._dashboard : buildExEmployeePanel();
      const addDashboard = haveAddBtn ? this._addDashboard : buildAddMembersPanel();
      this._dashboard = exDashboard;
      this._addDashboard = addDashboard;

      const anchorParent = chipList.parentElement;

      const buildTriggerBtn = (id, text, onClick) => {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.id = id;
        btn.className = "btn btn--small melon-green";
        btn.style.marginBottom = "10px";
        btn.style.marginRight = "10px";
        btn.textContent = text;
        btn.addEventListener("click", onClick);
        return btn;
      };

      // Add Members trigger
      if (!haveAddBtn) {
        const addTrigger = buildTriggerBtn(PANEL_IDS.addTriggerBtnId, "Add Members", () => {
          exDashboard.hide();
          addDashboard.toggle();
        });
        if (anchorParent) anchorParent.insertBefore(addTrigger, chipList);
        else chipList.before(addTrigger);
        insertAfter(addDashboard.panel, addTrigger);
      }

      // Remove Ex-Employees trigger (mounted after Add Members button so the
      // visual order is "Add | Remove" reading left-to-right).
      if (!haveExBtn) {
        const exTrigger = buildTriggerBtn(PANEL_IDS.triggerBtnId, "Remove Ex-Employees", () => {
          addDashboard.hide();
          exDashboard.toggle();
        });
        const addTrigger = document.getElementById(PANEL_IDS.addTriggerBtnId);
        if (addTrigger) {
          insertAfter(exTrigger, addTrigger);
        } else if (anchorParent) {
          anchorParent.insertBefore(exTrigger, chipList);
        } else {
          chipList.before(exTrigger);
        }
        insertAfter(exDashboard.panel, exTrigger);

        // Auto-scan on injection so ex-employee chips get marked even before
        // the user opens the panel.
        exDashboard.rescan().catch((e) => logError("Initial rescan failed:", e));
      }
    },

    rescan() {
      if (this._dashboard) {
        this._dashboard.rescan().catch((e) => logError("Manual rescan failed:", e));
      }
    },

    injectAllInlineUI() {
      this.inject();
    },
  };

  // ============================================================
  // INITIALIZATION
  // ============================================================

  const AppController = {
    mutationObserver: null,
    beforeUnloadHandler: null,
    reinjectIntervalId: null,
    _waiting: false,

    async waitForChipList() {
      if (this._waiting) return;
      this._waiting = true;
      try {
        for (let i = 0; i < TIMING.MAX_WAIT_ITERATIONS; i++) {
          const list = document.getElementById("systemGroupMembersChipList");
          if (list) {
            ButtonInjector.inject();
            return;
          }
          await sleep(TIMING.WAIT_ITERATION_DELAY_MS);
        }
      } finally {
        this._waiting = false;
      }
    },

    tick() {
      this.waitForChipList().catch((e) => logError("waitForChipList:", e));
    },

    handleMutations: debounce(function () {
      ButtonInjector.injectAllInlineUI();
    }, 300),

    rescan() {
      ButtonInjector.rescan();
    },

    init() {
      if (!PageDetector.isSystemGroupEdit) return;
      log("Initializing", VERSION);
      injectGlobalStyles();
      this.tick();

      this.mutationObserver = new MutationObserver(() => this.handleMutations());
      this.mutationObserver.observe(document.documentElement || document.body, {
        childList: true,
        subtree: true,
      });

      // Safety-net poll. The page replaces the chip-list subtree after manual
      // adds/removes, which can outpace the debounced MutationObserver. inject()
      // is idempotent (early-returns when the trigger button is already there),
      // so polling is cheap and self-healing.
      this.reinjectIntervalId = setInterval(() => {
        if (PageDetector.isSystemGroupEdit) {
          ButtonInjector.injectAllInlineUI();
        }
      }, TIMING.REINJECT_INTERVAL_MS);

      Object.assign(_debug, {
        ApiClient,
        ExEmployeeOperations,
        AddMemberOperations,
        ButtonInjector,
        PageDetector,
        InlineDashboard,
      });
      window.PatchExEmployees = { version: VERSION, _debug };
      window.PatchExEmployeesDebug = {
        enableDebug: () => { _debug.DEBUG_MODE = true; },
        disableDebug: () => { _debug.DEBUG_MODE = false; },
        injectButtons: () => {
          ButtonInjector.injectAllInlineUI();
          console.log("[PatchExEmployees] Manual UI injection triggered");
        },
        rescan: () => AppController.rescan(),
      };

      log("Loaded successfully");
    },

    cleanup() {
      if (this.mutationObserver) {
        this.mutationObserver.disconnect();
        this.mutationObserver = null;
      }
      if (this.reinjectIntervalId) {
        clearInterval(this.reinjectIntervalId);
        this.reinjectIntervalId = null;
      }
      if (this.beforeUnloadHandler) {
        window.removeEventListener("beforeunload", this.beforeUnloadHandler);
        this.beforeUnloadHandler = null;
      }
      log("Cleaned up");
    },
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => AppController.init());
  } else {
    AppController.init();
  }

  AppController.beforeUnloadHandler = () => AppController.cleanup();
  window.addEventListener("beforeunload", AppController.beforeUnloadHandler);
})();
