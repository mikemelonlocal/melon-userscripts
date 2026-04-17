// ==UserScript==
// @name         Patch Monthly Auto-Creator
// @namespace    http://tampermonkey.net/
// @version      1.3.0
// @description  Create next month's Monthly call based on the latest one on the current agent's Calls tab
// @match        https://thepatch.melonlocal.com/Agents/Dashboard/*
// @match        https://thepatch.melonlocal.com/agents/dashboard/*
// @run-at       document-end
// @grant        none
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/patch-monthly-auto-creator.user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/patch-monthly-auto-creator.user.js
// ==/UserScript==

(function () {
  "use strict";

  // ---------- Configuration constants ----------

  const CONFIG = {
    TIMEOUTS: {
      WAIT_FOR_DEFAULT: 15000,
      WAIT_FOR_INTERVAL: 150,
      FORM_APPEAR: 15000,
      FORM_APPEAR_INTERVAL: 200,
      AFTER_CLICK: 400,
      TITLE_INPUT_FOCUS: 150,
      TITLE_INPUT_CLEAR: 100,
      TITLE_INPUT_SET: 300,
      DATETIME_FOCUS: 100,
      DATETIME_SET: 200,
      DATETIME_BLUR: 300,
      SAVE_CLICK: 2000,
      SAVE_RESOLVE: 10000,
      SAVE_RESOLVE_INTERVAL: 150,
      TOAST_INFO_MS: 5000,
      TOAST_ERROR_MS: 8000
    },
    BUTTON_RETRY_DELAYS: [50, 300, 900],
    SELECTORS: {
      // Form inputs (IDs may change - document which app version these work with)
      TITLE_INPUT: "input#Title",
      TITLE_INPUT_FALLBACK: "input[name='Title']",
      SCHEDULED_INPUT: "input#ScheduledTime",
      SAVE_BUTTON: "button#newCallSave",

      // Grid and rows
      GRID_ROWS: "tr.k-master-row", // Simplified to avoid duplicates
      TABLE_CELLS: "td.k-table-td, td",
      TASK_TITLE: ".task-title",

      // Buttons
      NEW_CALL_PRIMARY: "button.addNewCall.btn.melon-green",

      // Error indicators
      ALERT_ROLE: "div[role='alert']",
      ALERT_DANGER: ".alert-danger",
      VALIDATION_ERRORS: ".validation-summary-errors"
    },
    DATE_VALIDATION: {
      MIN_YEAR: 2020,
      MAX_YEARS_AHEAD: 5
    }
  };

  // ---------- Brand tokens ----------

  const BRAND = {
    Alpine: "#FEF8E9",
    Cactus: "#47B74F",
    LemonSun: "#F1CB20",
    Sand: "#EDDFDB",
    Clover: "#40A74C",
    MustardSeed: "#CC8F15",
    WhitneyPink: "#FF9B94",
    WatermelonSugar: "#E9736E",
    Mojave: "#CFBA97",
    Pine: "#114E38",
    Coconut: "#644414",
    Cranberry: "#6C2126"
  };

  // ---------- Logger ----------

  const logger = {
    prefix: "[PatchMonthly]",
    info: (...args) => console.log(logger.prefix, ...args),
    warn: (...args) => console.warn(logger.prefix, ...args),
    error: (...args) => console.error(logger.prefix, ...args),
    debug: (...args) => console.debug(logger.prefix, ...args)
  };

  // ---------- Notification system ----------

  function ensureToastContainer() {
    let container = document.getElementById("patchMonthlyToastContainer");
    if (container) return container;
    container = document.createElement("div");
    container.id = "patchMonthlyToastContainer";
    Object.assign(container.style, {
      position: "fixed",
      top: "16px",
      right: "16px",
      zIndex: "2147483647",
      display: "flex",
      flexDirection: "column",
      gap: "8px",
      pointerEvents: "none",
      maxWidth: "380px"
    });
    document.body.appendChild(container);
    return container;
  }

  function showNotification(message, type = "info") {
    if (type === "error") logger.error(message);
    else if (type === "warn") logger.warn(message);
    else logger.info(message);

    const container = ensureToastContainer();
    const toast = document.createElement("div");

    const bg =
      type === "error" ? BRAND.WatermelonSugar :
      type === "warn" ? BRAND.LemonSun :
      BRAND.Cactus;
    const fg = type === "warn" ? BRAND.Coconut : BRAND.Alpine;

    Object.assign(toast.style, {
      background: bg,
      color: fg,
      padding: "12px 16px",
      borderRadius: "10px",
      boxShadow: "0 6px 20px rgba(0,0,0,0.18)",
      fontFamily: "system-ui, -apple-system, 'Segoe UI', sans-serif",
      fontSize: "14px",
      lineHeight: "1.4",
      pointerEvents: "auto",
      cursor: "pointer",
      opacity: "0",
      transform: "translateY(-8px)",
      transition: "opacity 160ms ease, transform 160ms ease",
      whiteSpace: "pre-wrap"
    });
    toast.textContent = message;

    let removed = false;
    const remove = () => {
      if (removed) return;
      removed = true;
      toast.style.opacity = "0";
      toast.style.transform = "translateY(-8px)";
      setTimeout(() => toast.remove(), 200);
    };
    toast.addEventListener("click", remove);

    container.appendChild(toast);
    requestAnimationFrame(() => {
      toast.style.opacity = "1";
      toast.style.transform = "translateY(0)";
    });

    const ttl = type === "error"
      ? CONFIG.TIMEOUTS.TOAST_ERROR_MS
      : CONFIG.TIMEOUTS.TOAST_INFO_MS;
    setTimeout(remove, ttl);
  }

  // ---------- Basic helpers ----------

  function norm(txt) {
    return (txt || "").replace(/\s+/g, " ").trim();
  }

  function isCallsPage() {
    const href = location.href.toLowerCase();
    const isDashboard = /agents\/dashboard\/\d+/.test(href);

    const hash = (location.hash || "").toLowerCase();
    const isCallsHash = hash === "#calls";

    const hasCallsGridText = document.body && document.body.textContent
      ? document.body.textContent.includes("Active Calls")
      : false;

    return isDashboard && (isCallsHash || hasCallsGridText);
  }

  function getGridRows() {
    return Array.from(document.querySelectorAll(CONFIG.SELECTORS.GRID_ROWS));
  }

  // Legacy positional fallback used only if headers can't be read.
  const POSITIONAL_FALLBACK = { title: 1, "call type": 2, status: 3, "scheduled time": 4 };

  function getColumnMap(table) {
    if (!table) return null;
    const headers = table.querySelectorAll("thead th");
    if (!headers.length) return null;
    const map = {};
    headers.forEach((th, i) => {
      const text = norm(th.textContent).toLowerCase();
      if (text) map[text] = i;
    });
    // Must contain at least title + scheduled time to be trusted
    if (map["title"] == null || map["scheduled time"] == null) return null;
    return map;
  }

  function getRowInfo(row, colMap) {
    const tds = row.querySelectorAll(CONFIG.SELECTORS.TABLE_CELLS);
    if (!tds.length) return null;

    const map = colMap || POSITIONAL_FALLBACK;
    const cell = (name) => {
      const idx = map[name];
      return idx != null && tds[idx] ? norm(tds[idx].textContent) : "";
    };

    const titleEl = row.querySelector(CONFIG.SELECTORS.TASK_TITLE);
    const title = titleEl ? norm(titleEl.textContent) : cell("title");
    const type = cell("call type");
    const status = cell("status");
    const scheduled = cell("scheduled time");

    return { row, title, type, status, scheduled };
  }

  /**
   * Parse scheduled time string
   * Expected format: "14:00 12/24/2025" (HH:MM MM/DD/YYYY)
   * Note: Uses browser's local timezone
   */
  function parseScheduled(s) {
    const m = s.match(/^(\d{1,2}):(\d{2})\s+(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (!m) return null;

    const hour = Number(m[1]);
    const minute = Number(m[2]);
    const month = Number(m[3]) - 1; // JS months are 0-indexed
    const day = Number(m[4]);
    const year = Number(m[5]);

    // Validate ranges
    if (hour < 0 || hour > 23) return null;
    if (minute < 0 || minute > 59) return null;
    if (month < 0 || month > 11) return null;
    if (day < 1 || day > 31) return null;

    const d = new Date(year, month, day, hour, minute);

    // Additional validation: check if date is reasonable
    if (isNaN(d.getTime())) return null;
    if (year < CONFIG.DATE_VALIDATION.MIN_YEAR) return null;
    if (year > new Date().getFullYear() + CONFIG.DATE_VALIDATION.MAX_YEARS_AHEAD) return null;

    return d;
  }

  function getLatestMonthlyCall() {
    const rows = getGridRows();
    if (!rows.length) return null;

    const table = rows[0].closest("table");
    const colMap = getColumnMap(table);
    if (!colMap) {
      logger.warn("Could not read grid headers — falling back to positional column indexes");
    }

    const infos = rows
      .map((row) => getRowInfo(row, colMap))
      .filter(
        (info) =>
          info &&
          info.type === "Monthly" &&
          info.scheduled &&
          parseScheduled(info.scheduled)
      );

    if (!infos.length) return null;

    infos.sort((a, b) => {
      const da = parseScheduled(a.scheduled);
      const db = parseScheduled(b.scheduled);
      return db - da; // newest first
    });

    return infos[0];
  }

  function getNthWeekdayInfo(date) {
    const year = date.getFullYear();
    const month = date.getMonth();
    const targetDow = date.getDay();
    const day = date.getDate();

    let count = 0;
    for (let d = 1; d <= day; d++) {
      const tmp = new Date(year, month, d);
      if (tmp.getDay() === targetDow) count++;
    }
    return { nth: count, weekday: targetDow };
  }

  /**
   * Get the date for the nth occurrence of a weekday in a month
   * If the nth occurrence doesn't exist (e.g., 5th Tuesday in a month with only 4),
   * falls back to the last occurrence of that weekday
   */
  function getDateForNthWeekday(year, monthIndex, nth, weekday, hour, minute) {
    let count = 0;
    let result = null;

    // Try to find the exact nth occurrence
    for (let d = 1; d <= 31; d++) {
      const dt = new Date(year, monthIndex, d, hour, minute);
      if (dt.getMonth() !== monthIndex) break; // Went into next month
      if (dt.getDay() === weekday) {
        count++;
        if (count === nth) {
          result = dt;
          break;
        }
      }
    }

    // Fallback: if nth occurrence doesn't exist, use the last occurrence
    if (!result) {
      logger.debug(`No ${nth}th occurrence of weekday ${weekday} in ${year}-${monthIndex+1}, using last occurrence`);
      let last = null;
      for (let d = 1; d <= 31; d++) {
        const dt = new Date(year, monthIndex, d, hour, minute);
        if (dt.getMonth() !== monthIndex) break;
        if (dt.getDay() === weekday) last = dt;
      }
      result = last;
    }

    return result;
  }

  /**
   * Format date as "Month YYYY" (e.g., "January 2025")
   * Note: Uses English month names - may need localization if app supports other languages
   */
  function monthTitleFromDate(date) {
    const monthNames = [
      "January","February","March","April","May","June",
      "July","August","September","October","November","December"
    ];
    return monthNames[date.getMonth()] + " " + date.getFullYear();
  }

  /**
   * Format date for display in success message
   * Returns format: "MM/DD/YYYY at HH:MM"
   */
  function formatDateForDisplay(date) {
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const year = date.getFullYear();
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    return `${month}/${day}/${year} at ${hours}:${minutes}`;
  }

  function findCallByExactTitle(title) {
    const rows = getGridRows();
    if (!rows.length) return null;
    const colMap = getColumnMap(rows[0].closest("table"));
    for (const row of rows) {
      const info = getRowInfo(row, colMap);
      if (!info) continue;
      if (info.title === title && info.type === "Monthly") return info;
    }
    return null;
  }

  // ---------- Wait helper ----------

  async function waitFor(predicate, timeoutMs = CONFIG.TIMEOUTS.WAIT_FOR_DEFAULT, intervalMs = CONFIG.TIMEOUTS.WAIT_FOR_INTERVAL) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      try {
        const v = predicate();
        if (v) return v;
      } catch (e) {
        // Only log unexpected errors
        if (e.message && !e.message.includes("Cannot read")) {
          logger.debug("waitFor predicate error:", e.message);
        }
      }
      await new Promise((r) => setTimeout(r, intervalMs));
    }
    return null;
  }

  /**
   * Wait for jQuery to be available
   * Some features depend on jQuery for Kendo widget access
   */
  async function waitForJQuery(timeoutMs = 5000) {
    return await waitFor(() => window.jQuery, timeoutMs, 100);
  }

  // ---------- New Call + Kendo helpers ----------

  async function clickNewCall() {
    const btn =
      document.querySelector(CONFIG.SELECTORS.NEW_CALL_PRIMARY) ||
      Array.from(document.querySelectorAll("button")).find(
        (b) => norm(b.textContent) === "New Call"
      );
    if (!btn) throw new Error("Could not find New Call button on Dashboard.");

    logger.debug("Clicking New Call button");
    btn.click();

    const form = await waitFor(
      () =>
        document.querySelector(CONFIG.SELECTORS.TITLE_INPUT) &&
        document.querySelector(CONFIG.SELECTORS.SCHEDULED_INPUT),
      CONFIG.TIMEOUTS.FORM_APPEAR,
      CONFIG.TIMEOUTS.FORM_APPEAR_INTERVAL
    );
    if (!form) throw new Error("New Call form did not appear.");

    await new Promise((r) => setTimeout(r, CONFIG.TIMEOUTS.AFTER_CLICK));
  }

  async function getScheduledPicker() {
    // Wait for jQuery if it's not ready yet
    const jQuery = await waitForJQuery();
    if (!jQuery) {
      logger.debug("jQuery not available, will use fallback date input method");
      return null;
    }

    const input = jQuery(CONFIG.SELECTORS.SCHEDULED_INPUT);
    if (!input.length) return null;

    return (
      input.data("kendoDateTimePicker") ||
      input.data("kendoDatePicker") ||
      null
    );
  }

  async function setScheduledDateTime(raw) {
    const dtInput = document.querySelector(CONFIG.SELECTORS.SCHEDULED_INPUT);
    if (!dtInput) throw new Error("New Call ScheduledTime input not found.");
    if (!raw) return;

    const picker = await getScheduledPicker();
    let valueToUse = raw;
    if (!(raw instanceof Date)) {
      const d = new Date(raw);
      if (!Number.isNaN(d.getTime())) valueToUse = d;
    }

    if (picker) {
      logger.debug("Using Kendo picker to set date");
      picker.value(valueToUse);
      picker.trigger("change");
    } else {
      logger.debug("Using fallback method to set date");
      dtInput.focus();
      await new Promise((r) => setTimeout(r, CONFIG.TIMEOUTS.DATETIME_FOCUS));
      dtInput.select();
      dtInput.value = String(valueToUse);
      dtInput.dispatchEvent(new Event("input", { bubbles: true }));
      dtInput.dispatchEvent(new Event("change", { bubbles: true }));
      await new Promise((r) => setTimeout(r, CONFIG.TIMEOUTS.DATETIME_SET));
      dtInput.blur();
    }
    await new Promise((r) => setTimeout(r, CONFIG.TIMEOUTS.DATETIME_BLUR));
  }

  // ---------- New call fill + save ----------

  async function fillAndSaveNewCall(title, callType, callDateRaw) {
    const titleInput =
      document.querySelector(CONFIG.SELECTORS.TITLE_INPUT) ||
      document.querySelector(CONFIG.SELECTORS.TITLE_INPUT_FALLBACK);
    if (!titleInput) throw new Error("New Call Title input not found.");

    logger.debug("Setting title:", title);
    titleInput.focus();
    await new Promise((r) => setTimeout(r, CONFIG.TIMEOUTS.TITLE_INPUT_FOCUS));

    // Clear existing value
    titleInput.value = "";
    titleInput.dispatchEvent(new Event("input", { bubbles: true }));
    await new Promise((r) => setTimeout(r, CONFIG.TIMEOUTS.TITLE_INPUT_CLEAR));

    // Set new value
    titleInput.value = title || "";
    titleInput.dispatchEvent(new Event("input", { bubbles: true }));
    titleInput.dispatchEvent(new Event("change", { bubbles: true }));

    await new Promise((r) => setTimeout(r, CONFIG.TIMEOUTS.TITLE_INPUT_SET));

    // Verify title stuck
    const currentTitle = (titleInput.value || "").trim();
    if (!currentTitle) {
      throw new Error(
        "Title did not stick on New Call modal even after delay; aborting save."
      );
    }

    // Set date/time if provided
    if (callDateRaw) {
      logger.debug("Setting date:", callDateRaw);
      await setScheduledDateTime(callDateRaw);
    }

    function isNewCallModalOpen() {
      const el =
        document.querySelector(CONFIG.SELECTORS.TITLE_INPUT) ||
        document.querySelector(CONFIG.SELECTORS.TITLE_INPUT_FALLBACK);
      return !!el && el.offsetParent !== null;
    }

    function findVisibleErrorBar() {
      const candidates = [
        CONFIG.SELECTORS.ALERT_ROLE,
        CONFIG.SELECTORS.ALERT_DANGER,
        CONFIG.SELECTORS.VALIDATION_ERRORS
      ];
      for (const sel of candidates) {
        for (const el of document.querySelectorAll(sel)) {
          if (el.offsetParent !== null && norm(el.textContent)) return el;
        }
      }
      return null;
    }

    const saveBtn = document.querySelector(CONFIG.SELECTORS.SAVE_BUTTON);
    if (!saveBtn) throw new Error("Save button not found.");

    logger.debug("Clicking Save button");
    saveBtn.click();

    // Wait for one of: modal closes (success) or an error bar appears (failure).
    // Replaces the old "click twice and hope" pattern, which risked duplicate saves.
    const outcome = await waitFor(
      () => {
        const err = findVisibleErrorBar();
        if (err) return { error: err };
        if (!isNewCallModalOpen()) return { success: true };
        return null;
      },
      CONFIG.TIMEOUTS.SAVE_RESOLVE,
      CONFIG.TIMEOUTS.SAVE_RESOLVE_INTERVAL
    );

    if (!outcome) {
      throw new Error("Save timed out — the form neither closed nor reported an error.");
    }
    if (outcome.error) {
      const errorText = norm(outcome.error.textContent) || "Unknown validation error";
      throw new Error("Patch did not accept the form: " + errorText);
    }
  }

  // ---------- Main logic ----------

  async function createNextMonthlyCallFromLatest() {
    if (!isCallsPage()) {
      showNotification("Not on an Agent Calls page.", "warn");
      return;
    }

    logger.info("Starting Monthly call creation process");

    const latest = getLatestMonthlyCall();
    if (!latest) {
      showNotification("No Monthly calls with a Scheduled Time found.", "warn");
      return;
    }

    logger.debug("Latest Monthly call:", latest.title, latest.scheduled);

    const latestDate = parseScheduled(latest.scheduled);
    if (!latestDate) {
      showNotification("Could not parse Scheduled Time of latest Monthly call.", "error");
      return;
    }

    const { nth, weekday } = getNthWeekdayInfo(latestDate);
    const hour = latestDate.getHours();
    const minute = latestDate.getMinutes();

    logger.debug(`Pattern: ${nth}th occurrence of weekday ${weekday} at ${hour}:${minute}`);

    // Calculate next month (handles year rollover: December -> January)
    const nextMonthIndex = latestDate.getMonth() + 1;
    const nextYear =
      nextMonthIndex > 11 ? latestDate.getFullYear() + 1 : latestDate.getFullYear();
    const monthIndexNormalized = nextMonthIndex % 12;

    const targetDate = getDateForNthWeekday(
      nextYear,
      monthIndexNormalized,
      nth,
      weekday,
      hour,
      minute
    );
    if (!targetDate) {
      showNotification("Could not compute next month's matching weekday.", "error");
      return;
    }

    const targetTitle = monthTitleFromDate(targetDate);
    logger.debug("Target call:", targetTitle, targetDate);

    // Check if already exists
    const existing = findCallByExactTitle(targetTitle);
    if (existing) {
      showNotification('Monthly call "' + targetTitle + '" already exists.', "info");
      return;
    }

    const confirmed = window.confirm(
      "Create this Monthly call?\n\n" +
      "  Title:  " + targetTitle + "\n" +
      "  When:   " + formatDateForDisplay(targetDate)
    );
    if (!confirmed) {
      logger.info("User cancelled Monthly call creation");
      showNotification("Cancelled — no call created.", "info");
      return;
    }

    try {
      await clickNewCall();
      await fillAndSaveNewCall(targetTitle, "Monthly", targetDate);
      showNotification(
        'Successfully created Monthly call "' +
          targetTitle +
          '" scheduled for ' +
          formatDateForDisplay(targetDate) +
          '.',
        "info"
      );
    } catch (e) {
      logger.error("Error creating Monthly call:", e);
      showNotification(
        "Error creating Monthly call: " +
          (e && e.message ? e.message : String(e)),
        "error"
      );
    }
  }

  // ---------- UI helpers ----------

  function injectMonthlyButtonStylesFrom(newCallBtn, btn) {
    const cs = window.getComputedStyle(newCallBtn);

    // Typography from New Call
    btn.style.fontFamily = cs.fontFamily;
    btn.style.fontSize = cs.fontSize;
    btn.style.fontWeight = cs.fontWeight;
    btn.style.lineHeight = cs.lineHeight;
    btn.style.boxSizing = cs.boxSizing;

    // Same flex behavior as New Call
    btn.style.display = cs.display || "inline-flex";
    btn.style.alignItems = cs.alignItems || "center";
    btn.style.justifyContent = cs.justifyContent || "center";

    // Use the exact same height as New Call so they match
    btn.style.height = cs.height;
    btn.style.minHeight = cs.minHeight;
    btn.style.maxHeight = cs.maxHeight;

    // Horizontal padding tuned for longer label
    const verticalPadding = cs.paddingTop || "8px";
    btn.style.paddingTop = verticalPadding;
    btn.style.paddingBottom = verticalPadding;
    btn.style.paddingLeft = "18px";
    btn.style.paddingRight = "18px";

    // Match radius
    btn.style.borderRadius = cs.borderRadius || "999px";

    // Keep text on one line
    btn.style.whiteSpace = "nowrap";

    // Placement
    btn.style.marginRight = "8px";

    // Brand colors
    btn.style.border = `1px solid ${BRAND.Cactus}`;
    btn.style.background = BRAND.Cactus;
    btn.style.color = BRAND.Alpine;
    btn.style.cursor = "pointer";
    btn.style.letterSpacing = "0.02em";

    // Hover state
    btn.addEventListener("mouseenter", () => {
      btn.style.background = BRAND.Pine;
      btn.style.borderColor = BRAND.Pine;
    });
    btn.addEventListener("mouseleave", () => {
      btn.style.background = BRAND.Cactus;
      btn.style.borderColor = BRAND.Cactus;
    });

    // Focus state
    btn.addEventListener("focus", () => {
      btn.style.outline = "none";
      btn.style.boxShadow = `0 0 0 2px ${BRAND.Alpine}, 0 0 0 4px ${BRAND.Pine}`;
    });
    btn.addEventListener("blur", () => {
      btn.style.boxShadow = "none";
    });
  }

  function preventNewCallWrap() {
    const newCallBtn = Array.from(document.querySelectorAll("button")).find(
      (b) => norm(b.textContent) === "New Call"
    );
    if (!newCallBtn) return;

    newCallBtn.style.whiteSpace = "nowrap";
  }

  // Cache to avoid repeated searches
  let cachedButtonContainer = null;

  function addMonthlyButton() {
    if (!isCallsPage()) return;
    if (document.getElementById("autoMonthlyCreateBtn")) return;

    // Use cached container if available
    let newCallBtn;
    if (cachedButtonContainer && cachedButtonContainer.parentElement) {
      newCallBtn = cachedButtonContainer;
    } else {
      newCallBtn = Array.from(document.querySelectorAll("button")).find(
        (b) => norm(b.textContent) === "New Call"
      );
      if (newCallBtn) {
        cachedButtonContainer = newCallBtn;
      }
    }

    if (!newCallBtn || !newCallBtn.parentElement) return;

    const btn = document.createElement("button");
    btn.id = "autoMonthlyCreateBtn";
    btn.type = "button";
    btn.textContent = "Create Next Monthly Call";

    injectMonthlyButtonStylesFrom(newCallBtn, btn);

    btn.addEventListener("click", () => {
      createNextMonthlyCallFromLatest();
    });

    newCallBtn.parentElement.insertBefore(btn, newCallBtn);

    // Keep "New Call" on a single line
    preventNewCallWrap();

    logger.debug("Monthly button added to page");
  }

  // ---------- Robust startup ----------

  function startObservers() {
    logger.info("Script initialized");
    addMonthlyButton();

    // Use more targeted observation - watch for the button container area
    // This reduces the performance impact of observing the entire document
    const targetNode = document.querySelector("main") || document.body;

    const obs = new MutationObserver(() => {
      addMonthlyButton();
    });

    obs.observe(targetNode, {
      childList: true,
      subtree: true
    });

    // Handle hash changes (e.g., navigating to #calls tab)
    window.addEventListener("hashchange", () => {
      logger.debug("Hash changed, attempting to add button");
      // Try multiple times with increasing delays to handle async rendering
      CONFIG.BUTTON_RETRY_DELAYS.forEach(delay => {
        setTimeout(() => addMonthlyButton(), delay);
      });
    });

    // Cleanup on page unload
    window.addEventListener("beforeunload", () => {
      try {
        obs.disconnect();
        logger.debug("Observer disconnected");
      } catch (e) {
        // Ignore errors during cleanup
      }
    });
  }

  startObservers();
})();