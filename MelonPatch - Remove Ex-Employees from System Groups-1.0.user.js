// ==UserScript==
// @name         MelonPatch - Remove Ex-Employees from System Groups
// @namespace    http://tampermonkey.net/
// @version      2.0.0
// @description  Highlights ex-employees on System Group pages and provides an inline panel to bulk-remove them. Themed and structured to match Patch Targeting Helper.
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
  const VERSION = "patch-ex-employees-v2.0.0";
  const DEBUG = false;

  const _debug = { DEBUG_MODE: DEBUG };

  const TIMING = {
    DEFAULT_DELAY_MS: 120,
    MAX_WAIT_ITERATIONS: 200,
    WAIT_ITERATION_DELAY_MS: 150,
    FOCUS_DELAY_MS: 0,
    RETRY_BASE_BACKOFF_MS: 200,
    RETRY_MAX_ATTEMPTS: 2, // total attempts after the first call = 2 retries
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
  };

  const ENDPOINTS = {
    REMOVE_MEMBER: "/SystemGroups/Edit?handler=RemoveTargetSystemGroupMember",
    GET_MELONS: "/Lists/GetMelonVms",
  };

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

    async getActiveMelonIds() {
      const res = await fetch(ENDPOINTS.GET_MELONS, { credentials: "same-origin" });
      if (!res.ok) {
        throw new Error(`HTTP ${res.status} ${res.statusText} fetching active melons`);
      }
      const melons = await res.json();
      return new Set(melons.map((m) => m.MelonId));
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
      const chips = Array.from(document.querySelectorAll("#systemGroupMembersChipList .k-chip"));
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
      const onPage = document.querySelectorAll("#systemGroupMembersChipList .k-chip").length;
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
  // BUTTON INJECTION
  // ============================================================

  const ButtonInjector = {
    _dashboard: null,

    inject() {
      if (!PageDetector.isSystemGroupEdit) return;
      if (document.getElementById(PANEL_IDS.triggerBtnId)) return;
      const chipList = document.getElementById("systemGroupMembersChipList");
      if (!chipList) return;

      const dashboard = buildExEmployeePanel();
      this._dashboard = dashboard;

      const triggerBtn = document.createElement("button");
      triggerBtn.type = "button";
      triggerBtn.id = PANEL_IDS.triggerBtnId;
      triggerBtn.className = "btn btn--small melon-green";
      triggerBtn.style.marginBottom = "10px";
      triggerBtn.style.marginRight = "10px";
      triggerBtn.textContent = "Remove Ex-Employees";
      triggerBtn.addEventListener("click", () => dashboard.toggle());

      const anchorParent = chipList.parentElement;
      if (anchorParent) {
        anchorParent.insertBefore(triggerBtn, chipList);
      } else {
        chipList.before(triggerBtn);
      }
      insertAfter(dashboard.panel, triggerBtn);

      // Auto-scan on injection so chips get marked even before the user
      // opens the panel.
      dashboard.rescan().catch((e) => logError("Initial rescan failed:", e));
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
    _waiting: false,

    async waitForChipList() {
      if (this._waiting) return;
      this._waiting = true;
      try {
        for (let i = 0; i < TIMING.MAX_WAIT_ITERATIONS; i++) {
          const list = document.getElementById("systemGroupMembersChipList");
          if (list && list.querySelector(".k-chip")) {
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

      Object.assign(_debug, {
        ApiClient,
        ExEmployeeOperations,
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
