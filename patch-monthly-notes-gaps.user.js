// ==UserScript==
// @name         Patch Monthly Notes Gaps
// @namespace    http://tampermonkey.net/
// @version      1.0.6
// @description  On the Patch Calls grid, identify months missing a final-status call (Completed, Canceled, Agent No Show) for a given year by title, and bulk-archive calls for that year by title. Uses the Melon color palette and only appears on the Calls tab.
// @match        https://thepatch.melonlocal.com/Agents/Dashboard/*
// @match        https://thepatch.melonlocal.com/agents/dashboard/*
// @run-at       document-end
// @grant        none
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/patch-monthly-notes-gaps.user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/patch-monthly-notes-gaps.user.js
// ==/UserScript==

(function () {
  "use strict";

  // Keep in sync with @version in the metadata block above.
  const VERSION = "monthly-notes-gaps-v1.0.6";

  // -----------------------------
  // CONFIG
  // -----------------------------
  const DEFAULT_START_MINIMIZED = true;
  const DRAG_BOUNDARY_PADDING = 6;
  const MIN_YEAR = 2000;
  const MAX_YEAR = 2099;
  const ARCHIVE_BUTTON_WAIT_MS = 2000;
  const ARCHIVE_BUTTON_POLL_MS = 50;
  const TOAST_TIMEOUT_MS = 8000;

  const UI_POS_KEY = "patchMonthlyNotesGaps_uiPos_v1";
  const UI_MIN_KEY = "patchMonthlyNotesGaps_uiMin_v1";

  const MONTH_NAMES = [
    "january", "february", "march", "april", "may", "june",
    "july", "august", "september", "october", "november", "december"
  ];
  const MONTH_LABELS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  const MONTH_REGEX = new RegExp(`\\b(${MONTH_NAMES.join("|")})\\b`, "i");
  const YEAR_REGEX = /\b(\d{4})\b/;

  const FINAL_STATUSES = new Set(["Completed", "Canceled", "Agent No Show"]);

  const MELON_COLORS = {
    alpine: "#FEF8E9",
    sand: "#EDDFDB",
    mojave: "#CFBA97",
    cactus: "#47B74F",
    clover: "#40A74C",
    cloverHover: "#368E40",
    pine: "#114E38",
    pineHover: "#0D3D2B",
    coconut: "#644414",
    cranberry: "#6C2126",
    cranberryHover: "#551A1E"
  };

  // -----------------------------
  // Utilities
  // -----------------------------
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  function norm(txt) {
    return (txt || "").replace(/\s+/g, " ").trim();
  }

  function getAgentIdFromDashboardUrl() {
    const m = location.pathname.match(/\/Agents\/Dashboard\/(\d+)/i);
    return m ? Number(m[1]) : null;
  }

  function getDashboardRows() {
    try {
      return Array.from(
        document.querySelectorAll("tr.k-table-row.k-master-row, tr.k-master-row")
      );
    } catch (e) {
      console.error("[PatchMonthlyNotesGaps] Error getting rows:", e);
      return [];
    }
  }

  // -----------------------------
  // UI persistence helpers
  // -----------------------------
  function saveUiPos(pos) {
    try {
      localStorage.setItem(UI_POS_KEY, JSON.stringify(pos));
    } catch (e) {
      console.warn("[PatchMonthlyNotesGaps] Failed to save UI pos:", e);
    }
  }

  function loadUiPos() {
    try {
      const raw = localStorage.getItem(UI_POS_KEY);
      if (!raw) return null;
      const p = JSON.parse(raw);
      if (typeof p.left !== "number" || typeof p.top !== "number") return null;
      return p;
    } catch (e) {
      console.warn("[PatchMonthlyNotesGaps] Failed to load UI pos:", e);
      return null;
    }
  }

  function saveUiMinimized(min) {
    try {
      localStorage.setItem(UI_MIN_KEY, min ? "1" : "0");
    } catch (e) {
      console.warn("[PatchMonthlyNotesGaps] Failed to save UI min:", e);
    }
  }

  function loadUiMinimized() {
    try {
      const raw = localStorage.getItem(UI_MIN_KEY);
      if (raw === null) return null;
      return raw === "1";
    } catch (e) {
      console.warn("[PatchMonthlyNotesGaps] Failed to load UI min:", e);
      return null;
    }
  }

  function clamp(n, min, max) {
    return Math.max(min, Math.min(max, n));
  }

  function makeDraggable(box, handle) {
    let dragging = false;
    let startX = 0;
    let startY = 0;
    let startLeft = 0;
    let startTop = 0;

    const onMouseDown = (e) => {
      if (e.button !== 0) return;
      // Don't start a drag when clicking interactive controls inside the handle.
      if (e.target.closest("button, input, select, textarea, a")) return;

      dragging = true;
      const rect = box.getBoundingClientRect();
      startX = e.clientX;
      startY = e.clientY;
      startLeft = rect.left;
      startTop = rect.top;

      box.style.right = "auto";
      box.style.bottom = "auto";
      box.style.left = `${startLeft}px`;
      box.style.top = `${startTop}px`;

      document.addEventListener("mousemove", onMouseMove, true);
      document.addEventListener("mouseup", onMouseUp, true);
      e.preventDefault();
      e.stopPropagation();
    };

    const onMouseMove = (e) => {
      if (!dragging) return;
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;

      const rect = box.getBoundingClientRect();
      const w = rect.width;
      const h = rect.height;

      const maxLeft = window.innerWidth - w - DRAG_BOUNDARY_PADDING;
      const maxTop = window.innerHeight - h - DRAG_BOUNDARY_PADDING;

      const newLeft = clamp(
        startLeft + dx,
        DRAG_BOUNDARY_PADDING,
        Math.max(DRAG_BOUNDARY_PADDING, maxLeft)
      );
      const newTop = clamp(
        startTop + dy,
        DRAG_BOUNDARY_PADDING,
        Math.max(DRAG_BOUNDARY_PADDING, maxTop)
      );

      box.style.left = `${newLeft}px`;
      box.style.top = `${newTop}px`;

      saveUiPos({ left: newLeft, top: newTop });
    };

    const onMouseUp = () => {
      if (!dragging) return;
      dragging = false;
      document.removeEventListener("mousemove", onMouseMove, true);
      document.removeEventListener("mouseup", onMouseUp, true);
    };

    handle.addEventListener("mousedown", onMouseDown, true);
  }

  // -----------------------------
  // Title month/year parsing
  // -----------------------------
  function getMonthFromTitle(title) {
    const t = String(title || "");

    const monthMatch = t.match(MONTH_REGEX);
    if (!monthMatch) return null;
    const month = MONTH_NAMES.indexOf(monthMatch[1].toLowerCase()) + 1;

    const yearMatch = t.match(YEAR_REGEX);
    if (!yearMatch) return null;
    const year = parseInt(yearMatch[1], 10);
    if (year < MIN_YEAR || year > MAX_YEAR) return null;

    return { year, month };
  }

  // -----------------------------
  // Grid scanners
  // -----------------------------
  function getGridFinalMonthsByYear() {
    const rows = getDashboardRows();
    const years = {};

    for (const row of rows) {
      try {
        const titleEl = row.querySelector(".task-title");
        if (!titleEl) continue;

        const info = getMonthFromTitle(norm(titleEl.textContent));
        if (!info) continue;

        const statusBadge = row.querySelector(".gridFilledCell");
        const statusText = statusBadge ? norm(statusBadge.textContent) : "";
        if (!FINAL_STATUSES.has(statusText)) continue;

        const { year, month } = info;
        if (!years[year]) years[year] = new Set();
        years[year].add(month);
      } catch (e) {
        console.warn("[PatchMonthlyNotesGaps] Error processing row:", e);
      }
    }

    return years;
  }

  function selectCallsByTitleYear(year) {
    const rows = getDashboardRows();
    let matchCount = 0;

    for (const row of rows) {
      try {
        const titleEl = row.querySelector(".task-title");
        if (!titleEl) continue;

        const info = getMonthFromTitle(norm(titleEl.textContent));
        if (!info || info.year !== year) continue;

        const checkbox = row.querySelector(
          'input[type="checkbox"][aria-label="Select row"]'
        );
        if (checkbox && !checkbox.checked) {
          checkbox.click();
          // Some grid frameworks listen for `change` rather than `click`.
          checkbox.dispatchEvent(new Event("change", { bubbles: true }));
          matchCount++;
        }
      } catch (e) {
        console.warn("[PatchMonthlyNotesGaps] Error selecting row:", e);
      }
    }

    return matchCount;
  }

  function findArchiveButton() {
    return (
      document.querySelector("button.bulk_archive_calls.k-grid-archive") ||
      [...document.querySelectorAll("button")].find(
        (b) => norm(b.textContent) === "Archive"
      ) ||
      null
    );
  }

  async function waitForEnabledArchiveButton(timeoutMs = ARCHIVE_BUTTON_WAIT_MS) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      const btn = findArchiveButton();
      if (btn && !btn.disabled && btn.getAttribute("aria-disabled") !== "true") {
        return btn;
      }
      await sleep(ARCHIVE_BUTTON_POLL_MS);
    }
    return null;
  }

  // -----------------------------
  // UI (Melon palette + white)
  // -----------------------------
  function makeBtn({ label, bg, hoverBg, handler, disabled = false }) {
    const btn = document.createElement("button");
    btn.textContent = label;
    btn.disabled = disabled;
    btn.style.padding = "7px 10px";
    btn.style.margin = "0";
    btn.style.border = "none";
    btn.style.borderRadius = "5px";
    btn.style.cursor = disabled ? "not-allowed" : "pointer";
    btn.style.fontSize = "12px";
    btn.style.color = "#FFFFFF";
    btn.style.background = disabled ? MELON_COLORS.sand : bg;
    btn.style.textAlign = "left";
    btn.style.opacity = disabled ? "0.6" : "1";
    btn.style.transition = "background 0.15s ease";

    if (!disabled) {
      btn.onmouseenter = () => { btn.style.background = hoverBg; };
      btn.onmouseleave = () => { btn.style.background = bg; };
    }

    btn.onclick = async () => {
      if (disabled) return;

      const originalText = btn.textContent;
      const originalBg = btn.style.background;
      btn.disabled = true;
      btn.style.cursor = "wait";
      btn.style.opacity = "0.6";
      btn.textContent = originalText + " ...";

      try {
        await handler();
      } catch (e) {
        console.error("[PatchMonthlyNotesGaps] Button handler error:", e);
        if (typeof window.__patchMonthlyNotesGapsToast === "function") {
          window.__patchMonthlyNotesGapsToast(e.message || "Error.", "error");
        }
      } finally {
        btn.disabled = false;
        btn.style.cursor = "pointer";
        btn.style.opacity = "1";
        btn.textContent = originalText;
        btn.style.background = originalBg;
      }
    };
    return btn;
  }

  // Document-level handler stored so removeUi() can detach it; prevents
  // listener buildup across SPA route changes.
  let escKeydownHandler = null;

  function addUi() {
    const old = document.getElementById("patchMonthlyNotesGapsBox");
    if (old) return;

    const box = document.createElement("div");
    box.id = "patchMonthlyNotesGapsBox";
    box.style.position = "fixed";
    box.style.zIndex = "99999";
    box.style.display = "flex";
    box.style.flexDirection = "column";
    box.style.gap = "6px";
    box.style.background = "#FFFFFF";
    box.style.border = `1px solid ${MELON_COLORS.sand}`;
    box.style.padding = "0";
    box.style.borderRadius = "8px";
    box.style.fontFamily =
      "system-ui, -apple-system, BlinkMacSystemFont, sans-serif";
    box.style.fontSize = "12px";
    box.style.minWidth = "260px";
    box.style.boxShadow = "0 4px 10px rgba(0,0,0,0.15)";

    const savedPos = loadUiPos();
    if (savedPos) {
      box.style.left = `${savedPos.left}px`;
      box.style.top = `${savedPos.top}px`;
    } else {
      box.style.right = "20px";
      box.style.bottom = "20px";
    }

    // Header
    const headerRow = document.createElement("div");
    headerRow.style.display = "flex";
    headerRow.style.alignItems = "center";
    headerRow.style.justifyContent = "space-between";
    headerRow.style.gap = "10px";
    headerRow.style.cursor = "move";
    headerRow.style.userSelect = "none";
    headerRow.style.padding = "8px 10px";
    headerRow.style.background = MELON_COLORS.cactus;
    headerRow.style.borderRadius = "8px 8px 0 0";
    headerRow.title = "Drag to move";

    const header = document.createElement("div");
    header.textContent = "Patch Monthly Notes Gaps";
    header.style.fontWeight = "700";
    header.style.letterSpacing = "0.02em";
    header.style.color = "#FFFFFF";

    const minBtn = document.createElement("button");
    minBtn.style.padding = "4px 10px";
    minBtn.style.margin = "0";
    minBtn.style.border = "none";
    minBtn.style.borderRadius = "999px";
    minBtn.style.cursor = "pointer";
    minBtn.style.fontSize = "11px";
    minBtn.style.fontWeight = "600";
    minBtn.style.color = "#FFFFFF";
    minBtn.style.background = MELON_COLORS.pine;
    minBtn.style.transition = "background 0.15s ease";
    minBtn.textContent = "Minimize";
    minBtn.title = "Toggle panel (ESC)";
    minBtn.onmouseenter = () => { minBtn.style.background = MELON_COLORS.pineHover; };
    minBtn.onmouseleave = () => { minBtn.style.background = MELON_COLORS.pine; };

    const headerBtns = document.createElement("div");
    headerBtns.style.display = "flex";
    headerBtns.style.gap = "6px";
    headerBtns.style.alignItems = "center";
    headerBtns.appendChild(minBtn);

    headerRow.appendChild(header);
    headerRow.appendChild(headerBtns);
    box.appendChild(headerRow);

    // Body
    const bodyWrap = document.createElement("div");
    bodyWrap.style.display = "flex";
    bodyWrap.style.flexDirection = "column";
    bodyWrap.style.gap = "6px";
    bodyWrap.style.padding = "8px 10px 10px 10px";
    bodyWrap.style.background = MELON_COLORS.alpine;
    bodyWrap.style.color = MELON_COLORS.pine;

    const sub = document.createElement("div");
    sub.textContent =
      "Check final-stage coverage and archive by year (titles like \"January 2025\").";
    sub.style.opacity = "0.98";
    sub.style.fontSize = "11px";
    sub.style.color = MELON_COLORS.pine;
    bodyWrap.appendChild(sub);

    // Year controls
    const yearRow = document.createElement("div");
    yearRow.style.display = "flex";
    yearRow.style.alignItems = "center";
    yearRow.style.gap = "4px";
    yearRow.style.marginTop = "2px";

    const yearLabel = document.createElement("span");
    yearLabel.textContent = "Year:";
    yearLabel.style.minWidth = "32px";
    yearLabel.style.color = MELON_COLORS.pine;

    const yearInput = document.createElement("input");
    yearInput.type = "number";
    yearInput.value = new Date().getFullYear();
    yearInput.min = String(MIN_YEAR);
    yearInput.max = String(MAX_YEAR);
    yearInput.style.width = "80px";
    yearInput.style.fontSize = "11px";
    yearInput.style.borderRadius = "4px";
    yearInput.style.border = `1px solid ${MELON_COLORS.sand}`;
    yearInput.style.padding = "2px 4px";
    yearInput.style.background = "#FFFFFF";
    yearInput.id = "patchMonthlyNotesGapsYearInput";

    yearRow.appendChild(yearLabel);
    yearRow.appendChild(yearInput);
    bodyWrap.appendChild(yearRow);

    // Inline, non-blocking status / toast area
    const statusEl = document.createElement("div");
    statusEl.style.display = "none";
    statusEl.style.marginTop = "4px";
    statusEl.style.padding = "6px 8px";
    statusEl.style.borderRadius = "4px";
    statusEl.style.fontSize = "11px";
    statusEl.style.whiteSpace = "pre-wrap";
    statusEl.style.lineHeight = "1.35";
    statusEl.style.border = `1px solid ${MELON_COLORS.sand}`;
    statusEl.style.background = "#FFFFFF";
    statusEl.style.color = MELON_COLORS.pine;
    bodyWrap.appendChild(statusEl);

    let statusTimer = null;
    function showToast(message, variant = "info") {
      if (statusTimer) {
        clearTimeout(statusTimer);
        statusTimer = null;
      }
      statusEl.textContent = message;
      statusEl.style.display = "block";
      if (variant === "error") {
        statusEl.style.borderColor = MELON_COLORS.cranberry;
        statusEl.style.color = MELON_COLORS.cranberry;
      } else if (variant === "success") {
        statusEl.style.borderColor = MELON_COLORS.clover;
        statusEl.style.color = MELON_COLORS.pine;
      } else {
        statusEl.style.borderColor = MELON_COLORS.sand;
        statusEl.style.color = MELON_COLORS.pine;
      }
      // Auto-expand the panel so the user can actually read the message.
      if (minimized) {
        minimized = false;
        applyMinState();
      }
      statusTimer = setTimeout(() => {
        statusEl.style.display = "none";
        statusEl.textContent = "";
        statusTimer = null;
      }, TOAST_TIMEOUT_MS);
    }
    window.__patchMonthlyNotesGapsToast = showToast;

    // Check missing months button
    bodyWrap.appendChild(
      makeBtn({
        label: "Check Missing Months (by title)",
        bg: MELON_COLORS.clover,
        hoverBg: MELON_COLORS.cloverHover,
        handler: async () => {
          const agentId = getAgentIdFromDashboardUrl();
          if (!agentId) {
            showToast("Could not determine agent id from URL.", "error");
            return;
          }

          const yearVal = parseInt(yearInput.value, 10);
          if (!yearVal || Number.isNaN(yearVal) || yearVal < MIN_YEAR || yearVal > MAX_YEAR) {
            showToast(`Enter a valid year (${MIN_YEAR}-${MAX_YEAR}).`, "error");
            return;
          }

          const yearsMap = getGridFinalMonthsByYear();
          const monthSet = yearsMap[yearVal];

          if (!monthSet) {
            showToast(
              `No Completed/Canceled/Agent No Show calls found on this grid for ${yearVal} (by title).\nAll 12 months are missing.`,
              "info"
            );
            return;
          }

          const missing = [];
          for (let m = 1; m <= 12; m++) {
            if (!monthSet.has(m)) missing.push(m);
          }

          if (!missing.length) {
            showToast(
              `Agent ${agentId}: every month in ${yearVal} has a final-stage call (by title).`,
              "success"
            );
          } else {
            const monthNames = missing.map((m) => MONTH_LABELS[m - 1]).join(", ");
            showToast(
              `Agent ${agentId} missing final-stage calls in ${yearVal}:\n${monthNames}`,
              "info"
            );
          }
        }
      })
    );

    // Archive button
    bodyWrap.appendChild(
      makeBtn({
        label: "Archive Calls for Year (by title)",
        bg: MELON_COLORS.cranberry,
        hoverBg: MELON_COLORS.cranberryHover,
        handler: async () => {
          const agentId = getAgentIdFromDashboardUrl();
          if (!agentId) {
            showToast("Could not determine agent id from URL.", "error");
            return;
          }

          const yearVal = parseInt(yearInput.value, 10);
          if (!yearVal || Number.isNaN(yearVal) || yearVal < MIN_YEAR || yearVal > MAX_YEAR) {
            showToast(`Enter a valid year (${MIN_YEAR}-${MAX_YEAR}).`, "error");
            return;
          }

          const count = selectCallsByTitleYear(yearVal);
          if (!count) {
            showToast(
              `No calls found for ${yearVal} by title on this page for agent ${agentId}.`,
              "info"
            );
            return;
          }

          // Destructive action: keep confirm() as an explicit user-intent gate.
          const confirmArchive = confirm(
            `Archive ${count} call(s) for ${yearVal} for agent ${agentId}?`
          );
          if (!confirmArchive) {
            showToast("Archive canceled.", "info");
            return;
          }

          const archiveBtn = await waitForEnabledArchiveButton();
          if (!archiveBtn) {
            showToast(
              "Archive button did not become enabled in time. Selected rows may not have registered with the grid.",
              "error"
            );
            return;
          }
          try {
            archiveBtn.click();
            showToast(`Archive requested for ${count} call(s) in ${yearVal}.`, "success");
          } catch (e) {
            console.error("[PatchMonthlyNotesGaps] Error clicking archive:", e);
            showToast(`Archive click failed: ${e.message}`, "error");
          }
        }
      })
    );

    box.appendChild(bodyWrap);
    document.body.appendChild(box);

    makeDraggable(box, headerRow);

    // Minimize behavior (persisted). Start minimized on fresh installs.
    let minimized = (function () {
      const stored = loadUiMinimized();
      if (localStorage.getItem(UI_MIN_KEY) === null) return DEFAULT_START_MINIMIZED;
      return stored;
    })();

    const applyMinState = () => {
      if (minimized) {
        bodyWrap.style.display = "none";
        minBtn.textContent = "Expand";
        box.style.minWidth = "260px";
      } else {
        bodyWrap.style.display = "flex";
        minBtn.textContent = "Minimize";
        box.style.minWidth = "340px";
      }
      saveUiMinimized(minimized);
    };

    minBtn.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      minimized = !minimized;
      applyMinState();
    };

    // Keyboard shortcut: ESC to toggle minimize.
    escKeydownHandler = (e) => {
      if (e.key !== "Escape" || e.ctrlKey || e.altKey || e.metaKey) return;
      const ae = document.activeElement;
      if (ae && (ae.tagName === "INPUT" || ae.tagName === "TEXTAREA" || ae.isContentEditable)) {
        return;
      }
      minimized = !minimized;
      applyMinState();
      e.preventDefault();
    };
    document.addEventListener("keydown", escKeydownHandler);

    applyMinState();
  }

  function removeUi() {
    const box = document.getElementById("patchMonthlyNotesGapsBox");
    if (box) box.remove();
    if (escKeydownHandler) {
      document.removeEventListener("keydown", escKeydownHandler);
      escKeydownHandler = null;
    }
    if (window.__patchMonthlyNotesGapsToast) {
      delete window.__patchMonthlyNotesGapsToast;
    }
  }

  // -----------------------------
  // Calls view detection + routing
  // -----------------------------
  function isCallsHash() {
    return (location.hash || "").toLowerCase() === "#calls";
  }

  function isOnCallsView() {
    return /\/agents\/dashboard\/[0-9]+/i.test(location.pathname) && isCallsHash();
  }

  function runCallsLogic() {
    addUi();
    console.log("[PatchMonthlyNotesGaps] Calls logic initialized.");
  }

  function handleRouteChange() {
    if (isOnCallsView()) {
      runCallsLogic();
    } else {
      removeUi();
    }
  }

  window.PatchMonthlyNotesGaps = { version: VERSION };

  // Event-driven SPA navigation detection: patch pushState/replaceState so we
  // get notified on programmatic route changes, plus popstate and hashchange
  // for back/forward and hash edits. Replaces the previous setInterval poll.
  (function patchHistoryForLocationChange() {
    const EVENT_NAME = "patch-mng-locationchange";
    const emit = () => window.dispatchEvent(new Event(EVENT_NAME));

    const origPush = history.pushState;
    const origReplace = history.replaceState;
    history.pushState = function (...args) {
      const ret = origPush.apply(this, args);
      emit();
      return ret;
    };
    history.replaceState = function (...args) {
      const ret = origReplace.apply(this, args);
      emit();
      return ret;
    };
    window.addEventListener("popstate", emit);
    window.addEventListener(EVENT_NAME, handleRouteChange);
  })();
  window.addEventListener("hashchange", handleRouteChange);

  // Initial check (direct load)
  handleRouteChange();

  console.log("[PatchMonthlyNotesGaps] Loaded", VERSION);
})();
