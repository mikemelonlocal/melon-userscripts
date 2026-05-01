// ==UserScript==
// @name         Patch Monthly Auto-Creator
// @namespace    http://tampermonkey.net/
// @version      1.4.1
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
      SAVE_RESOLVE: 10000,
      SAVE_RESOLVE_INTERVAL: 150,
      TOAST_INFO_MS: 5000,
      TOAST_WARN_MS: 7000,
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
      GRID_ROWS: "tr.k-master-row",
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
    },
    DOM_IDS: {
      AUTO_BTN: "autoMonthlyCreateBtn",
      STYLES: "melonAutoCreateStyles",
      OVERLAY: "melonAutoCreateOverlay",
      TOAST_CONTAINER: "patchMonthlyToastContainer"
    },
    CSS_CLASSES: {
      AUTO_BTN: "melon-auto-create-btn"
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

  const MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];

  const WEEKDAY_NAMES = [
    "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"
  ];

  // ---------- Logger ----------

  const logger = {
    prefix: "[PatchMonthly]",
    info: (...args) => console.log(logger.prefix, ...args),
    warn: (...args) => console.warn(logger.prefix, ...args),
    error: (...args) => console.error(logger.prefix, ...args),
    debug: (...args) => console.debug(logger.prefix, ...args)
  };

  // ---------- Style injection ----------

  function injectStyles() {
    if (document.getElementById(CONFIG.DOM_IDS.STYLES)) return;

    const style = document.createElement("style");
    style.id = CONFIG.DOM_IDS.STYLES;
    // !important throughout so we beat any host stylesheet without copying
    // computed styles off another button at runtime.
    style.textContent = `
      .${CONFIG.CSS_CLASSES.AUTO_BTN} {
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        box-sizing: border-box !important;
        padding: 8px 18px !important;
        margin-right: 8px !important;
        border: 1px solid ${BRAND.Cactus} !important;
        background: ${BRAND.Cactus} !important;
        color: ${BRAND.Alpine} !important;
        border-radius: 999px !important;
        font-family: inherit !important;
        font-size: inherit !important;
        font-weight: 600 !important;
        line-height: 1.4 !important;
        letter-spacing: 0.02em !important;
        white-space: nowrap !important;
        cursor: pointer !important;
        transition: background-color 120ms ease, border-color 120ms ease, box-shadow 120ms ease !important;
      }
      .${CONFIG.CSS_CLASSES.AUTO_BTN}:hover {
        background: ${BRAND.Pine} !important;
        border-color: ${BRAND.Pine} !important;
      }
      .${CONFIG.CSS_CLASSES.AUTO_BTN}:focus {
        outline: none !important;
        background: ${BRAND.Pine} !important;
        border-color: ${BRAND.Pine} !important;
        box-shadow: 0 0 0 2px ${BRAND.Alpine}, 0 0 0 4px ${BRAND.Pine} !important;
      }
      .${CONFIG.CSS_CLASSES.AUTO_BTN}:active {
        transform: translateY(1px);
      }
      /* Keep the original "New Call" button on a single line */
      ${CONFIG.SELECTORS.NEW_CALL_PRIMARY} {
        white-space: nowrap !important;
      }
      /* Transparent input-blocker shown while a save is in flight. */
      #${CONFIG.DOM_IDS.OVERLAY} {
        position: fixed !important;
        inset: 0 !important;
        z-index: 2147483646 !important;
        background: transparent !important;
        cursor: progress !important;
        pointer-events: all !important;
      }
    `;
    document.head.appendChild(style);
    logger.debug("Stylesheet injected");
  }

  // ---------- Notification system ----------

  function ensureToastContainer() {
    let container = document.getElementById(CONFIG.DOM_IDS.TOAST_CONTAINER);
    if (container) return container;
    container = document.createElement("div");
    container.id = CONFIG.DOM_IDS.TOAST_CONTAINER;
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

    const ttl =
      type === "error" ? CONFIG.TIMEOUTS.TOAST_ERROR_MS :
      type === "warn" ? CONFIG.TIMEOUTS.TOAST_WARN_MS :
      CONFIG.TIMEOUTS.TOAST_INFO_MS;
    setTimeout(remove, ttl);
  }

  // ---------- Input blocker overlay ----------

  function showInputBlocker() {
    let overlay = document.getElementById(CONFIG.DOM_IDS.OVERLAY);
    if (overlay) return overlay;
    overlay = document.createElement("div");
    overlay.id = CONFIG.DOM_IDS.OVERLAY;
    overlay.setAttribute("aria-hidden", "true");
    // Swallow stray clicks/keys; programmatic clicks (saveBtn.click()) are
    // unaffected by pointer-events, so our automation still works.
    const swallow = (e) => {
      e.stopPropagation();
      e.preventDefault();
    };
    overlay.addEventListener("click", swallow, true);
    overlay.addEventListener("mousedown", swallow, true);
    overlay.addEventListener("keydown", swallow, true);
    document.body.appendChild(overlay);
    logger.debug("Input blocker shown");
    return overlay;
  }

  function hideInputBlocker() {
    const overlay = document.getElementById(CONFIG.DOM_IDS.OVERLAY);
    if (overlay) {
      overlay.remove();
      logger.debug("Input blocker removed");
    }
  }

  // ---------- Basic helpers ----------

  function norm(txt) {
    return (txt || "").replace(/\s+/g, " ").trim();
  }

  function ordinalSuffix(n) {
    const v = n % 100;
    if (v >= 11 && v <= 13) return "th";
    switch (n % 10) {
      case 1: return "st";
      case 2: return "nd";
      case 3: return "rd";
      default: return "th";
    }
  }

  function escapeRegex(s) {
    return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
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

    if (hour < 0 || hour > 23) return null;
    if (minute < 0 || minute > 59) return null;
    if (month < 0 || month > 11) return null;
    if (day < 1 || day > 31) return null;

    const d = new Date(year, month, day, hour, minute);

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
    return {
      nth: Math.ceil(date.getDate() / 7),
      weekday: date.getDay()
    };
  }

  /**
   * Get the date for the nth occurrence of a weekday in a month.
   * Returns { date, usedFallback }. usedFallback is true when the requested
   * nth occurrence doesn't exist in that month and we substitute the last
   * occurrence of that weekday.
   */
  function getDateForNthWeekday(year, monthIndex, nth, weekday, hour, minute) {
    let count = 0;
    let result = null;

    for (let d = 1; d <= 31; d++) {
      const dt = new Date(year, monthIndex, d, hour, minute);
      if (dt.getMonth() !== monthIndex) break;
      if (dt.getDay() === weekday) {
        count++;
        if (count === nth) {
          result = dt;
          break;
        }
      }
    }

    if (result) return { date: result, usedFallback: false };

    // Fallback: use the last occurrence of the weekday in this month.
    let last = null;
    for (let d = 1; d <= 31; d++) {
      const dt = new Date(year, monthIndex, d, hour, minute);
      if (dt.getMonth() !== monthIndex) break;
      if (dt.getDay() === weekday) last = dt;
    }
    return { date: last, usedFallback: true };
  }

  function monthLabel(date) {
    return MONTH_NAMES[date.getMonth()] + " " + date.getFullYear();
  }

  /**
   * Build the next call's title.
   *
   * Title intelligence: if the previous title contains the previous month's
   * "Month YYYY" or bare "Month" name, replace just that substring with the
   * new equivalent so custom naming conventions ("Q4 2024 December Sync")
   * are preserved. Falls back to plain "Month YYYY".
   */
  function buildNextMonthlyTitle(targetDate, previousTitle, previousDate) {
    const newMonth = MONTH_NAMES[targetDate.getMonth()];
    const newMonthYear = monthLabel(targetDate);

    if (previousTitle && previousDate) {
      const prevMonth = MONTH_NAMES[previousDate.getMonth()];
      const prevYear = previousDate.getFullYear();

      // 1) Try the most specific match: "Month YYYY" together.
      const monthYearRe = new RegExp(
        "\\b" + escapeRegex(prevMonth) + "\\s+" + escapeRegex(String(prevYear)) + "\\b",
        "i"
      );
      if (monthYearRe.test(previousTitle)) {
        const next = previousTitle.replace(monthYearRe, newMonthYear);
        logger.debug(`Title preserved by Month-Year swap: "${previousTitle}" -> "${next}"`);
        return next;
      }

      // 2) Bare month name.
      const monthRe = new RegExp("\\b" + escapeRegex(prevMonth) + "\\b", "i");
      if (monthRe.test(previousTitle)) {
        const next = previousTitle.replace(monthRe, newMonth);
        logger.debug(`Title preserved by Month-only swap: "${previousTitle}" -> "${next}"`);
        return next;
      }
    }

    return newMonthYear;
  }

  /**
   * Format date for display in success message
   * Returns format: "MM/DD/YYYY at HH:MM"
   */
  function formatDateForDisplay(date) {
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    const year = date.getFullYear();
    const hours = String(date.getHours()).padStart(2, "0");
    const minutes = String(date.getMinutes()).padStart(2, "0");
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
        if (e.message && !e.message.includes("Cannot read")) {
          logger.debug("waitFor predicate error:", e.message);
        }
      }
      await new Promise((r) => setTimeout(r, intervalMs));
    }
    return null;
  }

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
    // Block user input for the duration of fill+save. Programmatic clicks
    // bypass pointer-events so our own automation is unaffected.
    showInputBlocker();
    try {
      const titleInput =
        document.querySelector(CONFIG.SELECTORS.TITLE_INPUT) ||
        document.querySelector(CONFIG.SELECTORS.TITLE_INPUT_FALLBACK);
      if (!titleInput) throw new Error("New Call Title input not found.");

      logger.debug("Setting title:", title);
      titleInput.focus();
      await new Promise((r) => setTimeout(r, CONFIG.TIMEOUTS.TITLE_INPUT_FOCUS));

      titleInput.value = "";
      titleInput.dispatchEvent(new Event("input", { bubbles: true }));
      await new Promise((r) => setTimeout(r, CONFIG.TIMEOUTS.TITLE_INPUT_CLEAR));

      titleInput.value = title || "";
      titleInput.dispatchEvent(new Event("input", { bubbles: true }));
      titleInput.dispatchEvent(new Event("change", { bubbles: true }));

      await new Promise((r) => setTimeout(r, CONFIG.TIMEOUTS.TITLE_INPUT_SET));

      const currentTitle = (titleInput.value || "").trim();
      if (!currentTitle) {
        throw new Error(
          "Title did not stick on New Call modal even after delay; aborting save."
        );
      }

      if (callDateRaw) {
        logger.debug("Setting date:", callDateRaw);
        await setScheduledDateTime(callDateRaw);
      }

      const isNewCallModalOpen = () => {
        const el =
          document.querySelector(CONFIG.SELECTORS.TITLE_INPUT) ||
          document.querySelector(CONFIG.SELECTORS.TITLE_INPUT_FALLBACK);
        return !!el && el.offsetParent !== null;
      };

      const findVisibleErrorBar = () => {
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
      };

      const saveBtn = document.querySelector(CONFIG.SELECTORS.SAVE_BUTTON);
      if (!saveBtn) throw new Error("Save button not found.");

      logger.debug("Clicking Save button");
      saveBtn.click();

      // Wait for one of: modal closes (success) or an error bar appears (failure).
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
    } finally {
      hideInputBlocker();
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

    const { date: targetDate, usedFallback } = getDateForNthWeekday(
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

    if (usedFallback) {
      const wkName = WEEKDAY_NAMES[weekday];
      showNotification(
        `Heads up: ${monthLabel(targetDate)} has no ${nth}${ordinalSuffix(nth)} ${wkName}. ` +
        `Falling back to the LAST ${wkName} of the month (${formatDateForDisplay(targetDate)}). ` +
        `Confirm this matches your intended cadence before saving.`,
        "warn"
      );
    }

    const targetTitle = buildNextMonthlyTitle(targetDate, latest.title, latestDate);
    logger.debug("Target call:", targetTitle, targetDate);

    const existing = findCallByExactTitle(targetTitle);
    if (existing) {
      showNotification('Monthly call "' + targetTitle + '" already exists.', "info");
      return;
    }

    const confirmed = window.confirm(
      "Create this Monthly call?\n\n" +
      "  Title:  " + targetTitle + "\n" +
      "  When:   " + formatDateForDisplay(targetDate) +
      (usedFallback ? "\n\n(Note: fell back to last " + WEEKDAY_NAMES[weekday] + " — see warning toast.)" : "")
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
        'Successfully created Monthly call "' + targetTitle +
          '" scheduled for ' + formatDateForDisplay(targetDate) + ".",
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

  let cachedButtonContainer = null;

  function addMonthlyButton() {
    if (!isCallsPage()) return;
    if (document.getElementById(CONFIG.DOM_IDS.AUTO_BTN)) return;

    let newCallBtn;
    if (cachedButtonContainer && cachedButtonContainer.parentElement) {
      newCallBtn = cachedButtonContainer;
    } else {
      newCallBtn = Array.from(document.querySelectorAll("button")).find(
        (b) => norm(b.textContent) === "New Call"
      );
      if (newCallBtn) cachedButtonContainer = newCallBtn;
    }

    if (!newCallBtn || !newCallBtn.parentElement) return;

    const btn = document.createElement("button");
    btn.id = CONFIG.DOM_IDS.AUTO_BTN;
    btn.type = "button";
    btn.className = CONFIG.CSS_CLASSES.AUTO_BTN;
    btn.textContent = "Create Next Monthly Call";

    btn.addEventListener("click", () => {
      createNextMonthlyCallFromLatest();
    });

    newCallBtn.parentElement.insertBefore(btn, newCallBtn);
    logger.debug("Monthly button added to page");
  }

  // ---------- Robust startup ----------

  function startObservers() {
    logger.info("Script initialized");
    injectStyles();
    addMonthlyButton();

    const targetNode = document.querySelector("main") || document.body;

    const obs = new MutationObserver(() => {
      addMonthlyButton();
    });

    obs.observe(targetNode, {
      childList: true,
      subtree: true
    });

    window.addEventListener("hashchange", () => {
      logger.debug("Hash changed, attempting to add button");
      CONFIG.BUTTON_RETRY_DELAYS.forEach((delay) => {
        setTimeout(() => addMonthlyButton(), delay);
      });
    });

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
