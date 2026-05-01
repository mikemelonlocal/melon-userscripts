// ==UserScript==
// @name         Patch Monthly Notes Gaps
// @namespace    http://tampermonkey.net/
// @version      1.0.9
// @description  On the Patch Calls grid, identify months missing a final-status call (Completed, Canceled, Agent No Show) for a given year by title, and bulk-archive calls for that year by title. Uses the Melon color palette and only appears on the Calls tab.
// @author       Melon Local
// @match        https://thepatch.melonlocal.com/Agents/Dashboard/*
// @match        https://thepatch.melonlocal.com/agents/dashboard/*
// @run-at       document-end
// @grant        none
// @noframes
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/patch-monthly-notes-gaps.user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/patch-monthly-notes-gaps.user.js
// @supportURL   https://github.com/mikemelonlocal/melon-userscripts/issues
// @homepageURL  https://github.com/mikemelonlocal/melon-userscripts
// ==/UserScript==

(function () {
  "use strict";

  // Keep CURRENT_VERSION in sync with @version in the metadata block above.
  // VERSION is the human-readable tag used in console logs and window export.
  const CURRENT_VERSION = "1.0.9";
  const VERSION = `monthly-notes-gaps-v${CURRENT_VERSION}`;

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
  const DOCK_WIDTH_PX = 340;
  // CSS class applied to the panel when docked. Drag handler bails on this.
  const DOCKED_CLASS = "pmng-docked";

  // Auto-update: lightweight in-panel notice that complements Tampermonkey's
  // own update mechanism (driven by @updateURL / @downloadURL above).
  const REMOTE_SCRIPT_URL =
    "https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/patch-monthly-notes-gaps.user.js";
  const UPDATE_CHECK_INTERVAL_MS = 6 * 60 * 60 * 1000; // 6 hours

  const UI_POS_KEY = "patchMonthlyNotesGaps_uiPos_v1";
  const UI_MIN_KEY = "patchMonthlyNotesGaps_uiMin_v1";
  const UI_DOCK_KEY = "patchMonthlyNotesGaps_uiDocked_v1";
  const UPDATE_CHECK_KEY = "patchMonthlyNotesGaps_lastUpdateCheck_v1";
  const UPDATE_LATEST_KEY = "patchMonthlyNotesGaps_latestVersion_v1";

  const MONTH_NAMES = [
    "january", "february", "march", "april", "may", "june",
    "july", "august", "september", "october", "november", "december"
  ];
  const MONTH_LABELS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  const MONTH_INITIALS = ["J", "F", "M", "A", "M", "J", "J", "A", "S", "O", "N", "D"];
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

  /**
   * Resilient row scraper. Filters Kendo grid for data rows only:
   * - Must have role="row" (Kendo's standard for data + header rows)
   * - Excludes grouping rows
   * - Excludes anything inside a <thead> (column headers, filter rows)
   */
  function getDashboardRows() {
    try {
      const rows = Array.from(
        document.querySelectorAll('tr[role="row"]:not(.k-grouping-row)')
      );
      return rows.filter((row) => !row.closest("thead"));
    } catch (e) {
      console.error("[PatchMonthlyNotesGaps] Error getting rows:", e);
      return [];
    }
  }

  // -----------------------------
  // Auto-update helpers
  // -----------------------------
  /**
   * Numeric semver compare. Returns 1 if a > b, -1 if a < b, 0 if equal.
   * Non-numeric segments are coerced to 0, which is fine for our X.Y.Z scheme.
   */
  function compareVersions(a, b) {
    const pa = String(a).split(".").map((n) => parseInt(n, 10) || 0);
    const pb = String(b).split(".").map((n) => parseInt(n, 10) || 0);
    const len = Math.max(pa.length, pb.length);
    for (let i = 0; i < len; i++) {
      const x = pa[i] || 0;
      const y = pb[i] || 0;
      if (x > y) return 1;
      if (x < y) return -1;
    }
    return 0;
  }

  /**
   * Best-effort fetch of the remote script's @version. Returns the remote
   * version string if it's strictly newer than CURRENT_VERSION; otherwise null.
   * Rate-limited via localStorage so we don't hammer GitHub on every UI mount.
   */
  async function checkForUpdate() {
    try {
      const lastRaw = localStorage.getItem(UPDATE_CHECK_KEY);
      const last = lastRaw ? parseInt(lastRaw, 10) : 0;
      const cached = localStorage.getItem(UPDATE_LATEST_KEY);

      // If we checked recently, return whatever we cached.
      if (last && Date.now() - last < UPDATE_CHECK_INTERVAL_MS) {
        if (cached && compareVersions(cached, CURRENT_VERSION) > 0) return cached;
        return null;
      }

      const res = await fetch(REMOTE_SCRIPT_URL, { cache: "no-cache" });
      if (!res.ok) return null;
      const text = await res.text();

      // Only look in the metadata block to avoid matching the VERSION constant.
      const meta = text.split("==/UserScript==")[0] || "";
      const m = meta.match(/@version\s+(\S+)/);
      if (!m) return null;
      const remote = m[1];

      localStorage.setItem(UPDATE_CHECK_KEY, String(Date.now()));
      localStorage.setItem(UPDATE_LATEST_KEY, remote);

      return compareVersions(remote, CURRENT_VERSION) > 0 ? remote : null;
    } catch (e) {
      console.warn("[PatchMonthlyNotesGaps] Update check failed:", e);
      return null;
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

  function saveUiDocked(docked) {
    try {
      localStorage.setItem(UI_DOCK_KEY, docked ? "1" : "0");
    } catch (e) {
      console.warn("[PatchMonthlyNotesGaps] Failed to save UI dock:", e);
    }
  }

  function loadUiDocked() {
    try {
      return localStorage.getItem(UI_DOCK_KEY) === "1";
    } catch (e) {
      console.warn("[PatchMonthlyNotesGaps] Failed to load UI dock:", e);
      return false;
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
      // Don't drag when docked — the panel is pinned to the viewport edge.
      if (box.classList.contains(DOCKED_CLASS)) return;
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

  /**
   * Selects all rows whose title month/year matches `year`.
   *
   * Kendo-aware: clicks the row checkbox, dispatches `change`, AND adds the
   * `.k-selected` class plus `aria-selected="true"` on the parent <tr> so the
   * grid's bulk-action toolbar recognizes the selection.
   *
   * @returns {{count: number, months: Set<number>}}
   */
  function selectCallsByTitleYear(year) {
    const rows = getDashboardRows();
    const months = new Set();
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
          checkbox.dispatchEvent(new Event("change", { bubbles: true }));
        }

        // Sync Kendo's row-level selection state regardless of checkbox path,
        // so bulk-action buttons treat these rows as selected.
        row.classList.add("k-selected");
        row.setAttribute("aria-selected", "true");

        matchCount++;
        months.add(info.month);
      } catch (e) {
        console.warn("[PatchMonthlyNotesGaps] Error selecting row:", e);
      }
    }

    return { count: matchCount, months };
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

    // Dock button — pins the panel as a right-side rail and shifts page content
    // left via document.body.style.marginRight.
    const dockBtn = document.createElement("button");
    dockBtn.style.padding = "4px 10px";
    dockBtn.style.margin = "0";
    dockBtn.style.border = "none";
    dockBtn.style.borderRadius = "999px";
    dockBtn.style.cursor = "pointer";
    dockBtn.style.fontSize = "11px";
    dockBtn.style.fontWeight = "600";
    dockBtn.style.color = "#FFFFFF";
    dockBtn.style.background = MELON_COLORS.pine;
    dockBtn.style.transition = "background 0.15s ease";
    dockBtn.textContent = "Dock";
    dockBtn.title = "Dock panel to the right sidebar";
    dockBtn.onmouseenter = () => { dockBtn.style.background = MELON_COLORS.pineHover; };
    dockBtn.onmouseleave = () => { dockBtn.style.background = MELON_COLORS.pine; };

    const headerBtns = document.createElement("div");
    headerBtns.style.display = "flex";
    headerBtns.style.gap = "6px";
    headerBtns.style.alignItems = "center";
    headerBtns.appendChild(dockBtn);
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

    // Update-available banner (hidden until checkForUpdate() resolves with a
    // newer version). Clicking the link opens the .user.js file directly so
    // Tampermonkey's userscript-detector takes over and offers re-install.
    const updateBanner = document.createElement("div");
    updateBanner.style.display = "none";
    updateBanner.style.marginTop = "2px";
    updateBanner.style.padding = "6px 8px";
    updateBanner.style.borderRadius = "4px";
    updateBanner.style.fontSize = "11px";
    updateBanner.style.lineHeight = "1.35";
    updateBanner.style.background = "#FFFFFF";
    updateBanner.style.border = `1px solid ${MELON_COLORS.cactus}`;
    updateBanner.style.color = MELON_COLORS.pine;
    bodyWrap.appendChild(updateBanner);

    function showUpdateBanner(remoteVersion) {
      updateBanner.textContent = "";
      const label = document.createElement("span");
      label.textContent = `Update available: v${remoteVersion} (you have v${CURRENT_VERSION}). `;
      const link = document.createElement("a");
      link.textContent = "Install";
      link.href = REMOTE_SCRIPT_URL;
      link.target = "_blank";
      link.rel = "noopener noreferrer";
      link.style.color = MELON_COLORS.clover;
      link.style.fontWeight = "700";
      link.style.textDecoration = "underline";
      updateBanner.appendChild(label);
      updateBanner.appendChild(link);
      updateBanner.style.display = "block";
    }

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

    // -----------------------------
    // Heatmap: 4x3 grid of month squares
    // -----------------------------
    const heatmapWrap = document.createElement("div");
    heatmapWrap.style.display = "grid";
    heatmapWrap.style.gridTemplateColumns = "repeat(4, 1fr)";
    heatmapWrap.style.gap = "4px";
    heatmapWrap.style.marginTop = "4px";
    heatmapWrap.title = "Month coverage. Click \"Check Missing Months\" to populate.";

    const heatmapSquares = [];
    for (let i = 0; i < 12; i++) {
      const sq = document.createElement("div");
      sq.textContent = MONTH_INITIALS[i];
      sq.title = `${MONTH_LABELS[i]} — unknown`;
      sq.dataset.month = String(i + 1);
      sq.style.background = MELON_COLORS.sand;
      sq.style.color = "#FFFFFF";
      sq.style.fontWeight = "700";
      sq.style.fontSize = "11px";
      sq.style.textAlign = "center";
      sq.style.padding = "6px 0";
      sq.style.borderRadius = "3px";
      sq.style.userSelect = "none";
      sq.style.transition = "background 0.15s ease";
      heatmapWrap.appendChild(sq);
      heatmapSquares.push(sq);
    }
    bodyWrap.appendChild(heatmapWrap);

    /**
     * Paint heatmap state.
     * @param {Set<number>|null} foundMonths
     *   - Set of months 1-12 that have a final-stage call. Missing months are
     *     filled with cranberry. Null/undefined means "no data scanned" so all
     *     12 squares are treated as missing.
     */
    function paintHeatmap(foundMonths) {
      for (let i = 0; i < 12; i++) {
        const month = i + 1;
        const sq = heatmapSquares[i];
        if (foundMonths && foundMonths.has(month)) {
          sq.style.background = MELON_COLORS.clover;
          sq.title = `${MONTH_LABELS[i]} — found`;
        } else {
          sq.style.background = MELON_COLORS.cranberry;
          sq.title = `${MONTH_LABELS[i]} — missing`;
        }
      }
    }

    function resetHeatmap() {
      for (let i = 0; i < 12; i++) {
        const sq = heatmapSquares[i];
        sq.style.background = MELON_COLORS.sand;
        sq.title = `${MONTH_LABELS[i]} — unknown`;
      }
    }

    // Reset to "unknown" whenever the user changes the year, since stale colors
    // would lie about the new year's coverage.
    yearInput.addEventListener("input", resetHeatmap);

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
          const monthSet = yearsMap[yearVal] || null;

          // Paint the heatmap regardless of outcome.
          paintHeatmap(monthSet);

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

          const { count, months } = selectCallsByTitleYear(yearVal);
          if (!count) {
            showToast(
              `No calls found for ${yearVal} by title on this page for agent ${agentId}.`,
              "info"
            );
            return;
          }

          // Audit summary breakdown.
          const monthList = [...months]
            .sort((a, b) => a - b)
            .map((m) => MONTH_LABELS[m - 1].slice(0, 3))
            .join(", ");
          const rowWord = count === 1 ? "row" : "rows";
          const auditMsg =
            `I found ${count} ${rowWord} for ${yearVal}.\n` +
            `Months covered: ${monthList || "—"}.\n\n` +
            `Archive these ${count} ${rowWord}?`;

          const confirmArchive = confirm(auditMsg);
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
            showToast(`Archive requested for ${count} ${rowWord} in ${yearVal}.`, "success");
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

    // Dock behavior (persisted). When docked, the panel becomes a full-height
    // right-side rail and the page content shifts left to make room.
    let docked = loadUiDocked();
    // Remember the floating position so undocking can restore it.
    let savedFloatPos = null;

    const applyDockState = () => {
      if (docked) {
        // Snapshot current floating position before pinning to the edge.
        if (!box.classList.contains(DOCKED_CLASS)) {
          const rect = box.getBoundingClientRect();
          savedFloatPos = { left: rect.left, top: rect.top };
        }
        box.classList.add(DOCKED_CLASS);
        box.style.left = "auto";
        box.style.top = "0";
        box.style.right = "0";
        box.style.bottom = "0";
        box.style.width = `${DOCK_WIDTH_PX}px`;
        box.style.height = "100vh";
        box.style.maxHeight = "100vh";
        box.style.borderRadius = "0";
        box.style.boxShadow = "-4px 0 24px rgba(0,0,0,0.18)";
        // Shift the host page content left so the dock doesn't cover it.
        document.body.style.marginRight = `${DOCK_WIDTH_PX}px`;
        dockBtn.textContent = "Undock";
        dockBtn.title = "Return to floating panel";
      } else {
        box.classList.remove(DOCKED_CLASS);
        box.style.right = "auto";
        box.style.bottom = "auto";
        box.style.width = "";
        box.style.height = "";
        box.style.maxHeight = "";
        box.style.borderRadius = "8px";
        box.style.boxShadow = "0 4px 10px rgba(0,0,0,0.15)";
        document.body.style.marginRight = "";
        // Restore the floating position the user had before docking, falling
        // back to the persisted position or the default bottom-right.
        const restore = savedFloatPos || loadUiPos();
        if (restore && typeof restore.left === "number" && typeof restore.top === "number") {
          box.style.left = `${restore.left}px`;
          box.style.top = `${restore.top}px`;
        } else {
          box.style.left = "auto";
          box.style.top = "auto";
          box.style.right = "20px";
          box.style.bottom = "20px";
        }
        dockBtn.textContent = "Dock";
        dockBtn.title = "Dock panel to the right sidebar";
      }
      saveUiDocked(docked);
    };

    dockBtn.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      docked = !docked;
      applyDockState();
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
    applyDockState();

    // Kick off a non-blocking update check. If a newer version is published
    // at REMOTE_SCRIPT_URL, surface an in-panel banner. Tampermonkey itself
    // also handles updates via @updateURL/@downloadURL on its own schedule;
    // this just gives the user a faster signal inside the panel.
    checkForUpdate().then((remote) => {
      if (remote && document.body.contains(box)) {
        showUpdateBanner(remote);
      }
    });
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
    // Always clear the body margin we may have set while docked, so leaving
    // the Calls view doesn't leave the host page squashed.
    if (document.body && document.body.style.marginRight) {
      document.body.style.marginRight = "";
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
  // for back/forward and hash edits.
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
