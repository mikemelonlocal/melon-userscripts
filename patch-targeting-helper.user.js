// ==UserScript==
// @name         Patch Targeting Helper – Bulk Add + Bulk Remove (Targets) + Bulk Move (BudgetDetails ListBoxes)
// @namespace    http://tampermonkey.net/
// @version      3.3.0
// @description  Inline bulk Add/Remove for Edit Advertising Targets (County/City/Zip) + Bulk Move for Kendo ListBoxes on BudgetDetails screens. Validation counts, live progress, retried API calls. Optional zip coverage analysis vs. City/County/DMA/State.
// @match        https://thepatch.melonlocal.com/*
// @run-at       document-end
// @grant        none
// @connect      raw.githubusercontent.com
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/patch-targeting-helper.user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/patch-targeting-helper.user.js
// ==/UserScript==

(function () {
  "use strict";

  // ============================================================
  // CONSTANTS
  // ============================================================

  // Keep in sync with @version in the userscript header above.
  const VERSION = "patch-targeting-helper-bulk-v3.3.0";
  const DEBUG = false;

  // Coverage analysis: bump EXPECTED when the JSON schema or data
  // changes meaningfully, so cached copies are invalidated.
  const ZIP_GEO = {
    URL: "https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/zip-geo-mapping.json",
    EXPECTED_VERSION: "1",
    CACHE_KEY: "patch-helper:zip-geo-mapping:v1",
    TOGGLE_KEY: "patch-helper:coverage-toggle",
    THRESHOLD: 0.5, // groups with >= 50% selected are shown
    MISSING_PREVIEW_MAX: 30,
  };

  const _debug = { DEBUG_MODE: DEBUG };

  window.PatchTargetingHelperDebug = {
    enableDebug: () => { _debug.DEBUG_MODE = true; },
    disableDebug: () => { _debug.DEBUG_MODE = false; },
    injectButtons: () => {
      ButtonInjector.injectAllInlineUI();
      console.log("[PatchTargetingHelper] Manual UI injection triggered");
    },
  };

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

  const SURFACE = {
    background: "#ffffff",
    border: "#edede8",
  };

  const PANEL_IDS = {
    inlinePanelClass: "patch-inline-panel",
    inlinePanelStatusClass: "patch-inline-status",
  };

  const TARGET_TYPES = {
    COUNTY: { key: "county", inputId: "newTargetCounty", handlerName: "NewTargetCounty", pretty: "Counties", singular: "County" },
    CITY:   { key: "city",   inputId: "newTargetCity",   handlerName: "NewTargetCity",   pretty: "Cities",   singular: "City" },
    ZIP:    { key: "zip",    inputId: "newTargetZip",    handlerName: "NewTargetZip",    pretty: "Zip Codes", singular: "Zip Code" },
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

  function parseLinesOrCsv(text) {
    const raw = String(text || "")
      .split(/[\n,]+/g)
      .map((s) => normalizeText(s))
      .filter(Boolean);
    const seen = new Set();
    const out = [];
    for (const v of raw) {
      const k = v.toLowerCase();
      if (!seen.has(k)) { seen.add(k); out.push(v); }
    }
    return out;
  }

  function parseZips(text) {
    const normalized = String(text || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    const raw = normalized.split(/[\s,]+/).map((s) => s.trim()).filter(Boolean);
    const cleaned = raw
      .map((z) => z.split("-")[0].replace(/[^\d]/g, ""))
      .filter((z) => /^\d{5}$/.test(z));
    const seen = new Set();
    const out = [];
    for (const z of cleaned) {
      if (!seen.has(z)) { seen.add(z); out.push(z); }
    }
    return out;
  }

  function setNativeValue(el, value) {
    try {
      const proto = Object.getPrototypeOf(el);
      const desc = proto ? Object.getOwnPropertyDescriptor(proto, "value") : null;
      if (desc?.set) desc.set.call(el, value);
      else el.value = value;
      el.dispatchEvent(new Event("input", { bubbles: true }));
      el.dispatchEvent(new Event("change", { bubbles: true }));
    } catch (error) {
      logError("Error setting native value:", error);
      el.value = value;
    }
  }

  function insertAfter(newNode, referenceNode) {
    const parent = referenceNode?.parentNode;
    if (!parent) return;
    parent.insertBefore(newNode, referenceNode.nextSibling);
  }

  function log(message, ...args) {
    if (DEBUG || _debug.DEBUG_MODE) {
      console.log(`[PatchTargetingHelper] ${message}`, ...args);
    }
  }

  function logError(message, ...args) {
    console.error(`[PatchTargetingHelper] ${message}`, ...args);
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
    get isMelonMaxBudgetDetails() {
      const h = this.hrefLower;
      return this.isPatch && (
        h.includes("/melonmax/melonmaxbudgetdetails") ||
        h.includes("/melonmaxbudgetaddtargetingpartial") ||
        h.includes("/agents/melonmaxbudgetaddtargetingpartial")
      );
    },
    get isAgentsBudgetDetails() {
      return this.isPatch && this.hrefLower.includes("/agents/budgetdetails");
    },
  };

  // ============================================================
  // GLOBAL STYLES (inline panel theming)
  // ============================================================

  function injectGlobalStyles() {
    if (document.getElementById("patch-helper-styles")) return;
    const styleTag = document.createElement("style");
    styleTag.id = "patch-helper-styles";
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
      .${PANEL_IDS.inlinePanelClass} textarea {
        width: 100%; box-sizing: border-box;
        font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
        font-size: 13px; padding: 8px;
        border: 1px solid ${COLORS.mojave};
        border-radius: 4px;
        resize: vertical; min-height: 90px;
        background: #fff;
      }
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

      .patch-coverage-toggle {
        display: flex; align-items: center; gap: 8px;
        margin-top: 10px;
        font-size: 13px; color: ${COLORS.coconut};
        cursor: pointer; user-select: none;
      }
      .patch-coverage-toggle input { cursor: pointer; }
      .patch-coverage-results {
        margin-top: 8px;
        background: #fff;
        border: 1px solid ${COLORS.mojave};
        border-radius: 6px;
        padding: 8px 10px;
        font-size: 12px;
        color: ${COLORS.coconut};
      }
      .patch-coverage-results[hidden] { display: none !important; }
      .patch-coverage-header {
        font-size: 12px; margin-bottom: 6px; color: ${COLORS.coconut};
      }
      .patch-coverage-subtle { color: ${COLORS.mojave}; font-weight: normal; }
      .patch-coverage-empty, .patch-coverage-loading, .patch-coverage-error {
        font-style: italic; color: ${COLORS.mojave}; padding: 4px 0;
      }
      .patch-coverage-error { color: ${COLORS.cranberry}; font-style: normal; }
      .patch-coverage-list { list-style: none; padding: 0; margin: 0; }
      .patch-coverage-row {
        border-top: 1px solid ${COLORS.alpine};
        padding: 6px 0;
      }
      .patch-coverage-row:first-child { border-top: none; }
      .patch-coverage-row-head {
        display: grid;
        grid-template-columns: 64px 1fr auto auto;
        gap: 8px;
        align-items: baseline;
      }
      .patch-coverage-type {
        font-size: 10px; text-transform: uppercase; letter-spacing: 0.5px;
        font-weight: 700; color: #fff;
        background: ${COLORS.pine};
        padding: 2px 6px; border-radius: 3px;
        text-align: center;
      }
      .patch-coverage-type[data-type="City"]   { background: ${COLORS.cactus}; }
      .patch-coverage-type[data-type="County"] { background: ${COLORS.mustardSeed}; }
      .patch-coverage-type[data-type="DMA"]    { background: ${COLORS.pine}; }
      .patch-coverage-type[data-type="State"]  { background: ${COLORS.coconut}; }
      .patch-coverage-label { font-weight: 600; }
      .patch-coverage-pct {
        font-weight: 700; color: ${COLORS.pine};
        font-variant-numeric: tabular-nums;
      }
      .patch-coverage-count {
        font-size: 11px; color: ${COLORS.mojave};
        font-variant-numeric: tabular-nums;
      }
      .patch-coverage-missing {
        margin-top: 4px; padding-left: 72px;
      }
      .patch-coverage-missing summary {
        cursor: pointer; font-size: 11px; color: ${COLORS.coconut};
      }
      .patch-coverage-missing-list {
        font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
        font-size: 11px;
        padding: 4px 0;
        word-break: break-word;
        color: ${COLORS.coconut};
      }
      .patch-coverage-copy {
        font-size: 11px; padding: 3px 8px;
        border: 1px solid ${COLORS.mojave}; background: ${COLORS.sand};
        color: ${COLORS.coconut};
        border-radius: 3px; cursor: pointer;
      }
      .patch-coverage-copy:hover { background: ${COLORS.alpine}; }
      .patch-coverage-complete {
        padding-left: 72px; font-size: 11px;
        color: ${COLORS.cactus}; font-weight: 600;
      }
      .patch-coverage-warn-invalid {
        margin-top: 4px; font-size: 11px;
        color: ${COLORS.mustardSeed};
      }
    `;
    document.head.appendChild(styleTag);
  }

  // ============================================================
  // API CLIENT — direct calls to the page's backend endpoints
  // ============================================================
  // Endpoint contract captured from a real chip-close network request:
  //   POST /Agents/RemoveTarget
  //   Content-Type: application/json; charset=UTF-8
  //   RequestVerificationToken: <from hidden input>
  //   Body: {"AgencyId":"7374","TargetName":"89052","TargetType":"Zip","TargetId":0}

  const ApiClient = {
    getAntiforgeryToken() {
      const input = document.querySelector('input[name="__RequestVerificationToken"]');
      return input?.value || "";
    },

    getAgencyId() {
      // 1. Hidden input (most common in ASP.NET MVC forms)
      const hidden = document.querySelector(
        'input[name="AgencyId"], input[name="agencyId"], input[name="agency-id"]'
      );
      if (hidden?.value) return String(hidden.value);

      // 2. Globals
      if (window.AgencyId != null && window.AgencyId !== "") return String(window.AgencyId);
      if (window.agencyId != null && window.agencyId !== "") return String(window.agencyId);

      // 3. data-* attributes
      const dataEl = document.querySelector("[data-agency-id]");
      if (dataEl?.dataset?.agencyId) return String(dataEl.dataset.agencyId);

      // 4. Inline onclicks that pass it
      const anyOnclick = document.querySelector('[onclick*="AgencyId"], [onclick*="agencyId"]');
      if (anyOnclick) {
        const m = anyOnclick.getAttribute("onclick").match(/(?:AgencyId|agencyId)\s*[:=]\s*['"]?(\d+)['"]?/);
        if (m) return m[1];
      }

      return null;
    },

    /**
     * POST /Agents/RemoveTarget with retry + exponential backoff. Throws if
     * all attempts fail.
     * @param {{agencyId:string, name:string, type:string, token:string}} args
     * @param {{maxRetries?:number, baseBackoffMs?:number}} [opts]
     */
    async removeTarget({ agencyId, name, type, token }, opts = {}) {
      const maxRetries = opts.maxRetries ?? TIMING.RETRY_MAX_ATTEMPTS;
      const baseBackoff = opts.baseBackoffMs ?? TIMING.RETRY_BASE_BACKOFF_MS;
      const pretty = type.charAt(0).toUpperCase() + type.slice(1).toLowerCase();
      const body = JSON.stringify({
        AgencyId: String(agencyId),
        TargetName: String(name),
        TargetType: pretty,
        TargetId: 0,
      });

      let lastError;
      for (let attempt = 0; attempt <= maxRetries; attempt++) {
        try {
          const response = await fetch("/Agents/RemoveTarget", {
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
          log(`removeTarget retry ${attempt + 1}/${maxRetries} for "${name}" after ${backoff}ms`, error?.message);
          await sleep(backoff);
        }
      }
      throw lastError ?? new Error("Unknown error in removeTarget");
    },
  };

  // ============================================================
  // BULK ADD OPERATIONS
  // ============================================================

  const BulkAddOperations = {
    getTargetTypeConfig(typeKey) {
      return ({ county: TARGET_TYPES.COUNTY, city: TARGET_TYPES.CITY, zip: TARGET_TYPES.ZIP })[typeKey] || null;
    },

    /**
     * Parse + dedup planning. Pure function — no side effects on DOM.
     * @param {string} typeKey
     * @param {string} rawText
     * @returns {{values:string[], toAdd:string[], skipped:string[], existingSize:number}}
     */
    plan(typeKey, rawText) {
      const values = typeKey === "zip" ? parseZips(rawText) : parseLinesOrCsv(rawText);
      const { existing } = BulkRemoveOperations.getExistingTargetNames(typeKey);
      const seenInBatch = new Set();
      const toAdd = [];
      const skipped = [];
      for (const val of values) {
        const key = normalizeText(val).toLowerCase();
        if (existing.has(key) || seenInBatch.has(key)) {
          skipped.push(val);
        } else {
          seenInBatch.add(key);
          toAdd.push(val);
        }
      }
      return { values, toAdd, skipped, existingSize: existing.size };
    },

    /**
     * Execute the bulk add. opts.onProgress({completed,total,phase}) is
     * invoked after each item.
     */
    async run(typeKey, rawText, delayMs = TIMING.DEFAULT_DELAY_MS, opts = {}) {
      const targetType = this.getTargetTypeConfig(typeKey);
      if (!targetType) throw new Error(`Unknown target type: ${typeKey}`);
      const { inputId, handlerName, pretty } = targetType;

      const input = document.getElementById(inputId);
      const handler = window[handlerName];
      if (!input || typeof handler !== "function") {
        throw new Error(`Cannot find input or handler for ${pretty}`);
      }

      const { toAdd, skipped, values } = this.plan(typeKey, rawText);
      log(`Bulk add plan`, { typeKey, total: values.length, toAdd: toAdd.length, skipped: skipped.length });

      if (!values.length) {
        return { ok: false, reason: "empty", message: `No valid ${pretty} found in the input.` };
      }
      if (!toAdd.length) {
        return { ok: true, completed: 0, total: 0, skipped: skipped.length, message: `All ${values.length} ${pretty} are already targeted. Nothing to add.` };
      }

      let completed = 0;
      for (let i = 0; i < toAdd.length; i++) {
        const val = toAdd[i];
        try {
          input.focus();
          setNativeValue(input, val);
          await sleep(30);
          handler();
          await sleep(delayMs);
          completed++;
        } catch (error) {
          logError(`Error adding ${val}:`, error);
        }
        if (opts.onProgress) opts.onProgress({ completed, total: toAdd.length, phase: "Adding" });
      }

      return {
        ok: true,
        completed,
        total: toAdd.length,
        skipped: skipped.length,
        message: `Added ${completed} of ${toAdd.length} ${pretty}` +
          (skipped.length ? `  •  Skipped ${skipped.length} duplicate(s)` : ""),
      };
    },
  };

  // ============================================================
  // BULK REMOVE OPERATIONS
  // ============================================================

  const BulkRemoveOperations = {
    /**
     * Existing target names for dedup. Tries the legacy table; falls back
     * to the chip UI scanner.
     */
    getExistingTargetNames(typeKey) {
      const existing = new Set();

      const rows = Array.from(document.querySelectorAll("#exampleTable tbody tr"));
      for (const row of rows) {
        const cols = row.querySelectorAll("td");
        if (cols.length < 2) continue;
        existing.add(normalizeText(cols[1].textContent).toLowerCase());
      }

      if (existing.size === 0) {
        for (const name of this.collectChipTargetNames(typeKey)) {
          existing.add(name.toLowerCase());
        }
      }

      return { existing, typeDetectionWorks: false };
    },

    /**
     * Walk up from the type's Add input to find the section container, then
     * scan for chip-like elements. Three strategies, most specific first.
     */
    collectChipTargetNames(typeKey) {
      const cap = typeKey.charAt(0).toUpperCase() + typeKey.slice(1);
      const addBtn = document.querySelector(`button[onclick="NewTarget${cap}()"]`);
      const input = document.getElementById(`newTarget${cap}`);
      const anchor = addBtn || input;
      if (!anchor) return [];

      const CLOSE_ICONS = /[×✕✖✗✘⨯⊝⊜×⊗✕✖✗✘]/g;
      const seen = new Set();
      const names = [];

      const cleanText = (raw) =>
        normalizeText(String(raw || "").replace(CLOSE_ICONS, ""));

      const HEADING_NOISE = /^(target counties|target cities|target zips|target states|licensed states|add|bulk add|bulk remove|cancel|remove all|target county|target city|target zip)$/i;

      const add = (raw) => {
        const n = cleanText(raw);
        if (!n || n.length > 100) return;
        if (HEADING_NOISE.test(n)) return;
        if (typeKey === "zip" && !/^\d{5}$/.test(n)) return;
        const k = n.toLowerCase();
        if (seen.has(k)) return;
        seen.add(k);
        names.push(n);
      };

      let container = anchor.parentElement;
      for (let depth = 0; depth < 8 && container; depth++, container = container.parentElement) {
        // Strategy 1: explicit chip classes
        const classChips = container.querySelectorAll(
          "[class*='chip'], [class*='pill'], [class*='tag'], [class*='badge'], [class*='token']"
        );
        for (const c of classChips) {
          if (c.contains(anchor)) continue;
          add(c.textContent);
        }
        if (names.length) return names;

        // Strategy 2: short-text wrappers containing a close icon/button.
        const candidates = container.querySelectorAll("span, div, li, button, a");
        const matches = [];
        for (const c of candidates) {
          if (c.contains(anchor) || c === container) continue;
          const closeEl = c.querySelector(
            "button[aria-label*='remove' i], button[aria-label*='close' i], " +
              "[class*='close'], [class*='remove'], [class*='delete'], svg"
          );
          const hasIcon = CLOSE_ICONS.test(c.textContent || "");
          CLOSE_ICONS.lastIndex = 0;
          if (!closeEl && !hasIcon) continue;
          const text = cleanText(c.textContent);
          if (!text) continue;
          if (matches.some((m) => m.el.contains(c))) continue;
          matches.push({ el: c, text });
        }
        if (matches.length >= 2) {
          for (const m of matches) add(m.text);
          if (names.length) return names;
        }
      }

      // Strategy 3: zip-only doc-wide 5-digit leaf scan
      if (typeKey === "zip") {
        const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_ELEMENT);
        while (walker.nextNode()) {
          const el = walker.currentNode;
          if (el.children.length !== 0) continue;
          const t = (el.textContent || "").trim();
          if (/^\d{5}$/.test(t)) add(t);
        }
      }

      return names;
    },

    /**
     * Bulk Remove (paste-based). Same v3.0.0 logic for the legacy table
     * path; if it finds no rows, falls back to the AJAX endpoint with
     * progress callbacks.
     */
    async run(typeKey, rawText, delayMs = TIMING.DEFAULT_DELAY_MS, opts = {}) {
      const targetType = BulkAddOperations.getTargetTypeConfig(typeKey);
      if (!targetType) throw new Error(`Unknown target type: ${typeKey}`);
      const { pretty } = targetType;

      const values = typeKey === "zip" ? parseZips(rawText) : parseLinesOrCsv(rawText);
      if (!values.length) {
        return { ok: false, reason: "empty", message: `No valid ${pretty} found in the input.` };
      }

      const normalizedValues = values.map((v) => normalizeText(v).toLowerCase());
      const rows = Array.from(document.querySelectorAll("#exampleTable tbody tr"));
      let removed = 0;

      // Path A: legacy DOM table
      let domAttempted = 0;
      for (const row of rows) {
        try {
          const cols = row.querySelectorAll("td");
          if (cols.length < 2) continue;
          const cellText = normalizeText(cols[1].textContent).toLowerCase();
          if (normalizedValues.includes(cellText)) {
            const link = row.querySelector('a[onclick*="DeleteAdvertisingTarget"]');
            if (link) {
              link.click();
              await sleep(delayMs);
              removed++;
              domAttempted++;
              if (opts.onProgress) opts.onProgress({ completed: removed, total: values.length, phase: "Removing" });
            }
          }
        } catch (error) {
          logError("Error removing row:", error);
        }
      }

      if (removed > 0) {
        return { ok: true, completed: removed, total: values.length, message: `Removed ${removed} ${pretty}` };
      }

      // Path B: AJAX
      const apiResult = await this._removeViaApi(typeKey, values, delayMs, opts);
      if (!apiResult) {
        return { ok: false, reason: "no-api", message: `Cannot remove via API: missing anti-forgery token or AgencyId.` };
      }
      return {
        ok: true,
        completed: apiResult.removed,
        total: values.length,
        failed: apiResult.failed,
        reload: apiResult.removed > 0,
        message: `Removed ${apiResult.removed} of ${values.length} ${pretty}` +
          (apiResult.failed.length ? `  •  Failed: ${apiResult.failed.length}` : ""),
      };
    },

    /**
     * Remove every target of the given type. DOM-first, API-fallback.
     */
    async removeAll(typeKey, delayMs = TIMING.DEFAULT_DELAY_MS, opts = {}) {
      const targetType = BulkAddOperations.getTargetTypeConfig(typeKey);
      if (!targetType) throw new Error(`Unknown target type: ${typeKey}`);
      const { pretty } = targetType;

      // Path A: legacy DOM table
      const domRows = Array.from(document.querySelectorAll("#exampleTable tbody tr"))
        .filter((r) => r.querySelector('a[onclick*="DeleteAdvertisingTarget"]'));
      if (domRows.length) {
        let removed = 0;
        for (const row of domRows) {
          try {
            const link = row.querySelector('a[onclick*="DeleteAdvertisingTarget"]');
            link.click();
            await sleep(delayMs);
            removed++;
            if (opts.onProgress) opts.onProgress({ completed: removed, total: domRows.length, phase: "Removing" });
          } catch (error) {
            logError("Error in removeAll (DOM path):", error);
          }
        }
        return { ok: true, completed: removed, total: domRows.length, message: `Removed ${removed} of ${domRows.length} row(s)` };
      }

      // Path B: chip + AJAX
      const names = this.collectChipTargetNames(typeKey);
      if (!names.length) {
        return { ok: false, reason: "empty", message: `No ${pretty} found to remove.` };
      }
      const apiResult = await this._removeViaApi(typeKey, names, delayMs, opts);
      if (!apiResult) {
        return { ok: false, reason: "no-api", message: `Cannot remove via API: missing anti-forgery token or AgencyId.` };
      }
      return {
        ok: true,
        completed: apiResult.removed,
        total: names.length,
        failed: apiResult.failed,
        reload: apiResult.removed > 0,
        message: `Removed ${apiResult.removed} of ${names.length} ${pretty}` +
          (apiResult.failed.length ? `  •  Failed: ${apiResult.failed.length}` : ""),
      };
    },

    async _removeViaApi(typeKey, names, delayMs, opts = {}) {
      const token = ApiClient.getAntiforgeryToken();
      const agencyId = ApiClient.getAgencyId();
      if (!token || !agencyId) {
        logError("API remove unavailable", { hasToken: !!token, hasAgencyId: !!agencyId });
        return null;
      }
      let removed = 0;
      const failed = [];
      for (let i = 0; i < names.length; i++) {
        const name = names[i];
        try {
          await ApiClient.removeTarget({ agencyId, name, type: typeKey, token });
          removed++;
          await sleep(delayMs);
        } catch (error) {
          logError(`API remove failed for ${name}:`, error?.message || error);
          failed.push(name);
        }
        if (opts.onProgress) opts.onProgress({ completed: i + 1, total: names.length, phase: "Removing" });
      }
      return { removed, failed };
    },
  };

  // ============================================================
  // BULK MOVE OPERATIONS (Kendo ListBox)
  // ============================================================

  const BulkMoveOperations = {
    getListBoxPairBySelectIds(sourceSelectId, destSelectId) {
      try {
        const sel1 = document.getElementById(sourceSelectId);
        const sel2 = document.getElementById(destSelectId);
        if (!sel1 || !sel2) return null;
        const available = window.jQuery?.(sel1).data("kendoListBox");
        const useThese = window.jQuery?.(sel2).data("kendoListBox");
        if (!available || !useThese) return null;
        return { available, useThese };
      } catch (error) {
        logError("Error getting ListBox pair:", error);
        return null;
      }
    },

    getItemText(item, lb) {
      if (!item) return "";
      const f = lb?.options?.dataTextField;
      if (f && item[f] != null) return String(item[f]);
      for (const k of ["County", "Zip", "State", "City", "DMA", "text", "Text", "TargetName", "value", "Value"]) {
        if (item[k] != null) return String(item[k]);
      }
      return "";
    },

    getItemValue(item, lb) {
      if (!item) return "";
      const f = lb?.options?.dataValueField;
      if (f && item[f] != null) return String(item[f]);
      for (const k of ["value", "Value", "Zip", "County", "State", "City", "DMA"]) {
        if (item[k] != null) return String(item[k]);
      }
      return "";
    },

    getItemsArray(lb) {
      try {
        const ds = lb.dataSource;
        if (!ds?.data) return [];
        return ds.data().map((item) => ({
          text: normalizeText(this.getItemText(item, lb)),
          value: this.getItemValue(item, lb),
          raw: item,
        }));
      } catch (error) {
        logError("Error getting items array:", error);
        return [];
      }
    },

    moveItemsByText(sourceLb, destLb, textArray, opts = {}) {
      try {
        const wanted = textArray.map((t) => normalizeText(t).toLowerCase());
        const sourceItems = this.getItemsArray(sourceLb);
        const toMove = sourceItems.filter((item) => wanted.includes(item.text.toLowerCase()));
        if (!toMove.length) return 0;
        const sourceDs = sourceLb.dataSource;
        const destDs = destLb.dataSource;
        for (let i = 0; i < toMove.length; i++) {
          const item = toMove[i];
          try {
            destDs.add(item.raw);
            const found = sourceDs.data().find((d) => this.getItemValue(d, sourceLb) === item.value);
            if (found) sourceDs.remove(found);
          } catch (error) {
            logError("Error moving item:", item.text, error);
          }
          if (opts.onProgress) opts.onProgress({ completed: i + 1, total: toMove.length, phase: "Moving" });
        }
        return toMove.length;
      } catch (error) {
        logError("Error in moveItemsByText:", error);
        return 0;
      }
    },

    moveAll(sourceLb, destLb, opts = {}) {
      try {
        const sourceItems = this.getItemsArray(sourceLb);
        if (!sourceItems.length) return 0;
        const sourceDs = sourceLb.dataSource;
        const destDs = destLb.dataSource;
        for (let i = 0; i < sourceItems.length; i++) {
          try {
            destDs.add(sourceItems[i].raw);
          } catch (error) {
            logError("Error adding item to destination:", sourceItems[i].text, error);
          }
          if (opts.onProgress) opts.onProgress({ completed: i + 1, total: sourceItems.length, phase: "Moving" });
        }
        try {
          const allData = sourceDs.data().slice();
          for (const d of allData) sourceDs.remove(d);
        } catch (error) {
          logError("Error removing items from source:", error);
        }
        return sourceItems.length;
      } catch (error) {
        logError("Error in moveAll:", error);
        return 0;
      }
    },

    findZipDragDropHeader_MelonMax() {
      const modal = document.querySelector('.modal.show, #dashboardModal');
      if (!modal) return null;
      const h6 = Array.from(modal.querySelectorAll("h6"));
      for (const el of h6) {
        const text = normalizeText(el.textContent);
        if (text.includes("Choose") && text.includes("Zip") && text.length < 50) return el;
      }
      const others = Array.from(modal.querySelectorAll("h5, h4, h3, label"));
      for (const el of others) {
        if (el.children.length > 3) continue;
        const text = normalizeText(el.textContent);
        if (text.includes("Choose") && text.includes("Zip") && text.length < 50) return el;
      }
      return null;
    },
  };

  // ============================================================
  // ZIP GEO DATA — lazy-loaded, cached zip → {state, city, county, DMA} map
  // ============================================================
  // Backing data is generated from ZIP_to_DMA_Mapping.xlsx by
  // scripts/build-zip-geo-mapping.py and committed to the repo as
  // zip-geo-mapping.json. The userscript fetches it on first use and caches
  // the parsed result in localStorage keyed by EXPECTED_VERSION so cached
  // copies invalidate when the schema changes.

  const ZipGeoData = {
    _data: null,
    _loading: null,
    _inverse: null,

    isLoaded() { return !!this._data; },

    async load() {
      if (this._data) return this._data;
      if (this._loading) return this._loading;
      this._loading = (async () => {
        const cached = this._readCache();
        if (cached) {
          log("ZipGeoData: loaded from cache", cached.stats);
          this._data = cached;
          return cached;
        }
        log("ZipGeoData: fetching", ZIP_GEO.URL);
        const res = await fetch(ZIP_GEO.URL, { credentials: "omit", cache: "force-cache" });
        if (!res.ok) throw new Error(`HTTP ${res.status} ${res.statusText}`);
        const text = await res.text();
        const data = JSON.parse(text);
        if (!data?.zips || data?.v !== ZIP_GEO.EXPECTED_VERSION) {
          throw new Error(`Unexpected mapping shape or version: v=${data?.v}`);
        }
        try { localStorage.setItem(ZIP_GEO.CACHE_KEY, text); }
        catch (e) { logError("ZipGeoData: cache write failed (quota?)", e?.message || e); }
        this._data = data;
        return data;
      })();
      try { return await this._loading; }
      finally { this._loading = null; }
    },

    _readCache() {
      try {
        const raw = localStorage.getItem(ZIP_GEO.CACHE_KEY);
        if (!raw) return null;
        const parsed = JSON.parse(raw);
        if (parsed?.v !== ZIP_GEO.EXPECTED_VERSION) return null;
        return parsed;
      } catch (e) {
        logError("ZipGeoData: cache read failed", e?.message || e);
        return null;
      }
    },

    clearCache() {
      try { localStorage.removeItem(ZIP_GEO.CACHE_KEY); } catch {}
      this._data = null;
      this._inverse = null;
    },

    zipInfo(zip) {
      if (!this._data) return null;
      const t = this._data.zips[zip];
      if (!t) return null;
      const [state, cityIdx, countyIdx, dmaCode] = t;
      return {
        state,
        city: this._data.cities[cityIdx],
        county: this._data.counties[countyIdx],
        dmaCode,
        dmaName: this._data.dmas[dmaCode] || dmaCode,
      };
    },

    /**
     * Build reverse indexes (group → Set of all zips in group) lazily. The
     * forward map has ~42k entries, so this is a single ~10ms walk and
     * memoized for the page lifetime.
     */
    inverse() {
      if (this._inverse) return this._inverse;
      if (!this._data) throw new Error("ZipGeoData not loaded");
      const cityToZips = new Map();
      const countyToZips = new Map();
      const stateToZips = new Map();
      const dmaToZips = new Map();
      const cities = this._data.cities;
      const counties = this._data.counties;
      for (const zip in this._data.zips) {
        const [state, cityIdx, countyIdx, dmaCode] = this._data.zips[zip];
        const cityKey = `${cities[cityIdx]}, ${state}`;
        const countyKey = `${counties[countyIdx]}, ${state}`;
        let set;
        set = cityToZips.get(cityKey);   if (!set) cityToZips.set(cityKey, set = new Set());   set.add(zip);
        set = countyToZips.get(countyKey); if (!set) countyToZips.set(countyKey, set = new Set()); set.add(zip);
        set = stateToZips.get(state);    if (!set) stateToZips.set(state, set = new Set());    set.add(zip);
        set = dmaToZips.get(dmaCode);    if (!set) dmaToZips.set(dmaCode, set = new Set());    set.add(zip);
      }
      this._inverse = { cityToZips, countyToZips, stateToZips, dmaToZips };
      return this._inverse;
    },
  };

  // ============================================================
  // COVERAGE ANALYZER
  // ============================================================
  // Given a list of selected zips, return groups (City/County/DMA/State)
  // where the selection covers at least THRESHOLD of the group's zips,
  // sorted by coverage desc. Used by the optional zip coverage UI.

  const CoverageAnalyzer = {
    /**
     * @param {string[]} selectedZips
     * @returns {{
     *   total:number, recognized:number, unrecognized:string[],
     *   results: Array<{type:string,label:string,key:string,percent:number,selectedCount:number,totalCount:number,missing:string[]}>
     * }}
     */
    analyze(selectedZips) {
      const data = ZipGeoData._data;
      if (!data) return { total: 0, recognized: 0, unrecognized: [], results: [] };

      const selectedSet = new Set();
      const unrecognized = [];
      for (const z of selectedZips) {
        if (data.zips[z]) selectedSet.add(z);
        else unrecognized.push(z);
      }
      if (!selectedSet.size) {
        return { total: selectedZips.length, recognized: 0, unrecognized, results: [] };
      }

      const inv = ZipGeoData.inverse();
      const cityHits = new Map();
      const countyHits = new Map();
      const stateHits = new Map();
      const dmaHits = new Map();

      for (const zip of selectedSet) {
        const info = ZipGeoData.zipInfo(zip);
        if (!info) continue;
        const cityKey = `${info.city}, ${info.state}`;
        const countyKey = `${info.county}, ${info.state}`;
        cityHits.set(cityKey, (cityHits.get(cityKey) || 0) + 1);
        countyHits.set(countyKey, (countyHits.get(countyKey) || 0) + 1);
        stateHits.set(info.state, (stateHits.get(info.state) || 0) + 1);
        dmaHits.set(info.dmaCode, (dmaHits.get(info.dmaCode) || 0) + 1);
      }

      const results = [];
      const collect = (type, hits, indexMap, labelFn) => {
        for (const [key, selectedCount] of hits) {
          const fullSet = indexMap.get(key);
          if (!fullSet) continue;
          const total = fullSet.size;
          const percent = selectedCount / total;
          if (percent < ZIP_GEO.THRESHOLD) continue;
          const missing = [];
          for (const z of fullSet) if (!selectedSet.has(z)) missing.push(z);
          missing.sort();
          results.push({ type, key, label: labelFn(key), percent, selectedCount, totalCount: total, missing });
        }
      };

      collect("City",   cityHits,   inv.cityToZips,   (k) => k);
      collect("County", countyHits, inv.countyToZips, (k) => k);
      collect("DMA",    dmaHits,    inv.dmaToZips,    (code) => `${data.dmas[code] || code} (${code})`);
      collect("State",  stateHits,  inv.stateToZips,  (k) => k);

      results.sort((a, b) =>
        b.percent - a.percent ||
        b.totalCount - a.totalCount ||
        a.label.localeCompare(b.label)
      );

      return {
        total: selectedZips.length,
        recognized: selectedSet.size,
        unrecognized,
        results,
      };
    },
  };

  // ============================================================
  // INLINE DASHBOARD — replaces ModalBuilder + FocusManager
  // ============================================================

  const InlineDashboard = {
    /**
     * Build a small button styled by variant.
     */
    makeButton(text, variant = "primary") {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = text;
      btn.className = `patch-inline-btn patch-inline-btn--${variant}`;
      btn.dataset.originalText = text;
      btn.dataset.originalVariant = variant;
      return btn;
    },

    /**
     * Toggle a button into/out of "busy" state. Stores originals on
     * data-* attrs so it can be restored exactly.
     */
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

    /**
     * Build an inline panel.
     *
     * @param {object} cfg
     * @param {string} cfg.id           - Unique DOM id.
     * @param {string} cfg.title        - Panel header text.
     * @param {"add"|"remove"|"move"} cfg.mode
     * @param {object} cfg.context      - Mode-specific data.
     * @param {Function} cfg.onAction   - Async (intent, ui) => result. UI has {panel, status, runBtn, removeAllBtn?, textarea}.
     * @param {Function} [cfg.onValidate] - (rawText, ui) => string  — sets the status label.
     * @param {Array<{label:string, intent:string, variant?:string}>} cfg.actions
     */
    create(cfg) {
      const panel = document.createElement("div");
      panel.id = cfg.id;
      panel.className = PANEL_IDS.inlinePanelClass;
      panel.hidden = true;
      panel.dataset.bulkPanel = cfg.mode;

      // Header
      const header = document.createElement("div");
      header.className = `${PANEL_IDS.inlinePanelClass}-header`;
      const titleEl = document.createElement("div");
      titleEl.className = `${PANEL_IDS.inlinePanelClass}-title`;
      titleEl.textContent = cfg.title;
      const closeBtn = document.createElement("button");
      closeBtn.type = "button";
      closeBtn.className = `${PANEL_IDS.inlinePanelClass}-close`;
      closeBtn.setAttribute("aria-label", "Close");
      closeBtn.textContent = "×";
      header.appendChild(titleEl);
      header.appendChild(closeBtn);

      // Textarea (omitted for "move" mode? we still want it for pasted move)
      const textarea = document.createElement("textarea");
      textarea.placeholder = cfg.placeholder || "Paste items, one per line or comma-separated...";

      // Status span
      const status = document.createElement("div");
      status.className = `${PANEL_IDS.inlinePanelStatusClass} is-info`;

      // Buttons row
      const btnRow = document.createElement("div");
      btnRow.className = `${PANEL_IDS.inlinePanelClass}-button-row`;

      const cancelBtn = this.makeButton("Cancel", "secondary");

      const actionButtons = [];
      for (const action of cfg.actions || []) {
        const btn = this.makeButton(action.label, action.variant || "primary");
        btn.dataset.intent = action.intent;
        actionButtons.push(btn);
      }

      btnRow.appendChild(cancelBtn);
      for (const b of actionButtons) btnRow.appendChild(b);

      panel.appendChild(header);
      panel.appendChild(textarea);
      panel.appendChild(status);
      panel.appendChild(btnRow);

      const ui = { panel, textarea, status, actionButtons, cancelBtn, closeBtn };

      // Validation (live status)
      const runValidation = () => {
        try {
          const labelText = cfg.onValidate ? cfg.onValidate(textarea.value, ui) : "";
          status.innerHTML = labelText || "";
          status.classList.remove("is-error", "is-warn");
          status.classList.add("is-info");
        } catch (error) {
          logError("Validation error:", error);
        }
      };
      const debouncedValidate = debounce(runValidation, 150);
      textarea.addEventListener("input", debouncedValidate);

      // Close handlers
      const close = () => {
        panel.hidden = true;
      };
      cancelBtn.addEventListener("click", close);
      closeBtn.addEventListener("click", close);

      // Action click handlers
      for (const btn of actionButtons) {
        btn.addEventListener("click", async () => {
          if (btn.disabled) return;
          // Disable other actions during the run
          actionButtons.forEach((b) => { if (b !== btn) b.disabled = true; });
          cancelBtn.disabled = true;
          try {
            await cfg.onAction(btn.dataset.intent, {
              panel, textarea, status, runBtn: btn, ui, runValidation,
            });
          } catch (error) {
            logError("Action error:", error);
            status.classList.remove("is-info", "is-warn");
            status.classList.add("is-error");
            status.textContent = `Error: ${error?.message || error}`;
            this.setBusy(btn, false);
          } finally {
            actionButtons.forEach((b) => { if (b !== btn) b.disabled = false; });
            cancelBtn.disabled = false;
          }
        });
      }

      return {
        panel,
        textarea,
        status,
        actionButtons,
        cancelBtn,
        show() {
          panel.hidden = false;
          runValidation();
          setTimeout(() => { try { textarea.focus(); } catch {} }, TIMING.FOCUS_DELAY_MS);
        },
        hide: close,
        toggle() { panel.hidden ? this.show() : this.hide(); },
        isVisible() { return !panel.hidden; },
        runValidation,
      };
    },
  };

  // ============================================================
  // INLINE ACTION RUNNERS — wire operations to inline panel UI
  // ============================================================

  function escapeHtml(s) {
    return String(s ?? "").replace(/[&<>"']/g, (c) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
    })[c]);
  }

  /**
   * Attach the "Show coverage analysis" toggle + result panel to a Bulk Add
   * Zip dashboard. Wires the textarea so analysis updates live (debounced),
   * and remembers the toggle state in localStorage.
   */
  function attachCoverageAnalysis(dashboard) {
    const panel = dashboard.panel;
    const textarea = dashboard.textarea;

    const toggleLabel = document.createElement("label");
    toggleLabel.className = "patch-coverage-toggle";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = localStorage.getItem(ZIP_GEO.TOGGLE_KEY) === "1";
    const labelSpan = document.createElement("span");
    labelSpan.textContent = "Show coverage analysis (City / County / DMA / State)";
    toggleLabel.appendChild(checkbox);
    toggleLabel.appendChild(labelSpan);

    const results = document.createElement("div");
    results.className = "patch-coverage-results";
    results.hidden = !checkbox.checked;

    const btnRow = panel.querySelector(`.${PANEL_IDS.inlinePanelClass}-button-row`);
    panel.insertBefore(toggleLabel, btnRow);
    panel.insertBefore(results, btnRow);

    const render = async () => {
      if (!checkbox.checked) { results.hidden = true; return; }
      results.hidden = false;

      if (!ZipGeoData.isLoaded()) {
        results.innerHTML = `<div class="patch-coverage-loading">Loading zip → geo data…</div>`;
        try { await ZipGeoData.load(); }
        catch (e) {
          results.innerHTML = `<div class="patch-coverage-error">Failed to load geo data: ${escapeHtml(e?.message || e)}</div>`;
          return;
        }
      }

      const existing = BulkRemoveOperations.collectChipTargetNames("zip");
      const pasted = parseZips(textarea.value || "");
      const merged = Array.from(new Set([...existing, ...pasted]));

      if (!merged.length) {
        results.innerHTML = `<div class="patch-coverage-empty">Paste zip codes (or have some already targeted) to see coverage.</div>`;
        return;
      }

      const analysis = CoverageAnalyzer.analyze(merged);
      renderCoverageResults(results, analysis, {
        existingCount: existing.length,
        pastedCount: pasted.length,
        mergedCount: merged.length,
      });
    };

    const debouncedRender = debounce(render, 250);
    textarea.addEventListener("input", debouncedRender);
    checkbox.addEventListener("change", () => {
      try { localStorage.setItem(ZIP_GEO.TOGGLE_KEY, checkbox.checked ? "1" : "0"); } catch {}
      render();
    });

    // Re-render whenever the panel is shown so existing chips drive an
    // initial view even before the user pastes anything.
    const origShow = dashboard.show.bind(dashboard);
    dashboard.show = function () {
      origShow();
      render();
    };

    // After an Add action finishes the page has new chips, so coverage is
    // stale. The action button toggles to disabled while running and back
    // when done — wait for that transition, then re-render.
    for (const actionBtn of dashboard.actionButtons) {
      actionBtn.addEventListener("click", () => {
        if (!checkbox.checked) return;
        let sawDisabled = false;
        const tick = () => {
          if (!checkbox.checked || panel.hidden) return;
          if (actionBtn.disabled) { sawDisabled = true; setTimeout(tick, 200); return; }
          if (sawDisabled) { render(); return; }
          setTimeout(tick, 75);
        };
        setTimeout(tick, 50);
      });
    }
  }

  function renderCoverageResults(box, analysis, counts) {
    const { results, recognized, unrecognized } = analysis;
    const header = `
      <div class="patch-coverage-header">
        Coverage of <strong>${counts.mergedCount}</strong> unique zip${counts.mergedCount === 1 ? "" : "s"}
        <span class="patch-coverage-subtle">
          (${counts.existingCount} on page${counts.pastedCount ? ` + ${counts.pastedCount} pasted` : ""}${recognized !== counts.mergedCount ? `, ${recognized} recognized` : ""})
        </span>
      </div>`;

    const warn = unrecognized.length
      ? `<div class="patch-coverage-warn-invalid">${unrecognized.length} zip${unrecognized.length === 1 ? " is" : "s are"} not in the mapping and were excluded from coverage.</div>`
      : "";

    if (!results.length) {
      box.innerHTML = header + warn + `<div class="patch-coverage-empty">No group is at least ${Math.round(ZIP_GEO.THRESHOLD * 100)}% covered yet.</div>`;
      return;
    }

    const rows = results.map((r) => {
      const pct = Math.round(r.percent * 100);
      const previewZips = r.missing.slice(0, ZIP_GEO.MISSING_PREVIEW_MAX).join(", ");
      const more = r.missing.length > ZIP_GEO.MISSING_PREVIEW_MAX
        ? `  …+${r.missing.length - ZIP_GEO.MISSING_PREVIEW_MAX} more`
        : "";
      const allZips = r.missing.join(",");
      const missingBlock = r.missing.length
        ? `<details class="patch-coverage-missing">
             <summary>${r.missing.length} missing zip${r.missing.length === 1 ? "" : "s"}</summary>
             <div class="patch-coverage-missing-list">${escapeHtml(previewZips)}${escapeHtml(more)}</div>
             <button type="button" class="patch-coverage-copy" data-zips="${escapeHtml(allZips)}">Copy missing zips</button>
           </details>`
        : `<div class="patch-coverage-complete">✓ Complete — every zip in this group is selected</div>`;
      return `
        <li class="patch-coverage-row">
          <div class="patch-coverage-row-head">
            <span class="patch-coverage-type" data-type="${escapeHtml(r.type)}">${escapeHtml(r.type)}</span>
            <span class="patch-coverage-label">${escapeHtml(r.label)}</span>
            <span class="patch-coverage-pct">${pct}%</span>
            <span class="patch-coverage-count">${r.selectedCount}/${r.totalCount}</span>
          </div>
          ${missingBlock}
        </li>`;
    }).join("");

    box.innerHTML = header + warn + `<ul class="patch-coverage-list">${rows}</ul>`;

    box.querySelectorAll(".patch-coverage-copy").forEach((btn) => {
      btn.addEventListener("click", async (e) => {
        e.preventDefault();
        const zips = btn.dataset.zips || "";
        try {
          await navigator.clipboard.writeText(zips.replace(/,/g, "\n"));
          const orig = btn.textContent;
          btn.textContent = "✓ Copied";
          setTimeout(() => { btn.textContent = orig; }, 1500);
        } catch (err) {
          logError("Clipboard write failed", err);
          btn.textContent = "Copy failed";
        }
      });
    });
  }

  function buildAddPanel(typeKey, pretty) {
    const dashboard = InlineDashboard.create({
      id: `patch-bulk-add-${typeKey}`,
      title: `Bulk Add ${pretty}`,
      mode: "add",
      placeholder: typeKey === "zip"
        ? "Paste zip codes (one per line or comma-separated)..."
        : `Paste ${pretty.toLowerCase()} (one per line or comma-separated)...`,
      actions: [{ label: `Add ${pretty}`, intent: "add", variant: "primary" }],
      onValidate: (raw) => {
        if (!raw.trim()) return `Paste ${pretty.toLowerCase()} to see how many will be added.`;
        const { values, toAdd, skipped } = BulkAddOperations.plan(typeKey, raw);
        const addCount = `<strong>${toAdd.length}</strong>`;
        const totalCount = `<span class="badge">${values.length} pasted</span>`;
        const dupNote = skipped.length
          ? ` <span class="badge" style="background:${COLORS.lemonSun};color:${COLORS.coconut}">${skipped.length} duplicate(s)</span>`
          : "";
        return `Will add ${addCount} ${pretty.toLowerCase()} ${totalCount}${dupNote}`;
      },
      onAction: async (intent, { textarea, status, runBtn, runValidation }) => {
        if (intent !== "add") return;
        const rawText = textarea.value || "";
        const { toAdd } = BulkAddOperations.plan(typeKey, rawText);
        if (!toAdd.length) {
          status.classList.remove("is-info");
          status.classList.add("is-warn");
          status.textContent = "Nothing to add (all values are empty or already targeted).";
          return;
        }

        InlineDashboard.setBusy(runBtn, true, `(0/${toAdd.length}) Adding...`);
        const onProgress = ({ completed, total }) => {
          runBtn.textContent = `(${completed}/${total}) Adding...`;
        };

        const result = await BulkAddOperations.run(typeKey, rawText, TIMING.DEFAULT_DELAY_MS, { onProgress });
        InlineDashboard.setBusy(runBtn, false);

        status.classList.remove("is-info", "is-warn", "is-error");
        if (result.ok) {
          status.classList.add("is-info");
          status.textContent = result.message;
          // Re-run validation against the (now updated) chip list so dup
          // counts update if the user keeps the panel open.
          runValidation();
        } else {
          status.classList.add("is-warn");
          status.textContent = result.message || "No-op.";
        }
      },
    });
    if (typeKey === "zip") attachCoverageAnalysis(dashboard);
    return dashboard;
  }

  function buildRemovePanel(typeKey, pretty) {
    const dashboard = InlineDashboard.create({
      id: `patch-bulk-remove-${typeKey}`,
      title: `Bulk Remove ${pretty}`,
      mode: "remove",
      placeholder: typeKey === "zip"
        ? "Paste zip codes to remove (one per line or comma-separated)..."
        : `Paste ${pretty.toLowerCase()} to remove (one per line or comma-separated)...`,
      actions: [
        { label: `Remove ${pretty}`, intent: "remove", variant: "danger" },
        { label: `Remove ALL ${pretty}`, intent: "removeAll", variant: "danger" },
      ],
      onValidate: (raw) => {
        const existingCount = BulkRemoveOperations.collectChipTargetNames(typeKey).length
          || Array.from(document.querySelectorAll("#exampleTable tbody tr")).length;
        const onPage = `<span class="badge">${existingCount} on page</span>`;
        if (!raw.trim()) {
          return `Paste ${pretty.toLowerCase()} to see how many will be removed, or click <strong>Remove ALL ${pretty}</strong>. ${onPage}`;
        }
        const values = typeKey === "zip" ? parseZips(raw) : parseLinesOrCsv(raw);
        return `Will remove <strong>${values.length}</strong> ${pretty.toLowerCase()} ${onPage}`;
      },
      onAction: async (intent, { textarea, status, runBtn, runValidation, ui }) => {
        const isRemoveAll = intent === "removeAll";

        // Confirm guard for destructive Remove ALL
        if (isRemoveAll) {
          const onPage = BulkRemoveOperations.collectChipTargetNames(typeKey).length
            || Array.from(document.querySelectorAll("#exampleTable tbody tr")).length;
          if (!onPage) {
            status.classList.remove("is-info");
            status.classList.add("is-warn");
            status.textContent = `No ${pretty} to remove.`;
            return;
          }
          if (!confirm(`Remove ALL ${onPage} ${pretty}? This cannot be undone.`)) return;
        }

        const rawText = textarea.value || "";
        const expected = isRemoveAll
          ? (BulkRemoveOperations.collectChipTargetNames(typeKey).length
              || Array.from(document.querySelectorAll("#exampleTable tbody tr")).length)
          : (typeKey === "zip" ? parseZips(rawText).length : parseLinesOrCsv(rawText).length);

        if (expected === 0) {
          status.classList.remove("is-info");
          status.classList.add("is-warn");
          status.textContent = "Nothing to remove.";
          return;
        }

        InlineDashboard.setBusy(runBtn, true, `(0/${expected}) Removing...`);
        const onProgress = ({ completed, total }) => {
          runBtn.textContent = `(${completed}/${total}) Removing...`;
        };

        const result = isRemoveAll
          ? await BulkRemoveOperations.removeAll(typeKey, TIMING.DEFAULT_DELAY_MS, { onProgress })
          : await BulkRemoveOperations.run(typeKey, rawText, TIMING.DEFAULT_DELAY_MS, { onProgress });

        InlineDashboard.setBusy(runBtn, false);

        status.classList.remove("is-info", "is-warn", "is-error");
        if (result.ok) {
          status.classList.add("is-info");
          status.textContent = result.message;
          if (result.reload) {
            status.textContent += "  •  Reloading page...";
            setTimeout(() => location.reload(), 500);
          } else {
            runValidation();
          }
        } else {
          status.classList.add("is-warn");
          status.textContent = result.message || "No-op.";
        }
      },
    });
    return dashboard;
  }

  function buildMovePanel({ id, titleProvider, pairGetter }) {
    const titleText = typeof titleProvider === "function"
      ? (titleProvider(window.jQuery) || "Bulk Move Targeting")
      : (titleProvider || "Bulk Move Targeting");

    const dashboard = InlineDashboard.create({
      id,
      title: titleText,
      mode: "move",
      placeholder: "Paste items to move (one per line or comma-separated)...",
      actions: [
        { label: "Pasted: Available → Use These", intent: "pasted-AtoU", variant: "primary" },
        { label: "Pasted: Use These → Available", intent: "pasted-UtoA", variant: "primary" },
        { label: "ALL: Available → Use These",    intent: "all-AtoU",    variant: "primary" },
        { label: "ALL: Use These → Available",    intent: "all-UtoA",    variant: "primary" },
      ],
      onValidate: (raw) => {
        const lbs = pairGetter();
        if (!lbs) return `<span style="color:${COLORS.cranberry}">ListBoxes not found yet — make sure the targeting area is visible.</span>`;
        const aCount = BulkMoveOperations.getItemsArray(lbs.available).length;
        const uCount = BulkMoveOperations.getItemsArray(lbs.useThese).length;
        const lhs = `<span class="badge">${aCount} Available</span> <span class="badge">${uCount} Use These</span>`;
        if (!raw.trim()) return `Paste items, or use the ALL buttons. ${lhs}`;
        const values = parseLinesOrCsv(raw);
        return `Will move <strong>${values.length}</strong> pasted item(s). ${lhs}`;
      },
      onAction: async (intent, { textarea, status, runBtn, runValidation }) => {
        const lbs = pairGetter();
        if (!lbs) {
          status.classList.remove("is-info");
          status.classList.add("is-error");
          status.textContent = "Could not find the two listboxes. Make sure the targeting area is visible.";
          return;
        }
        const isPasted = intent.startsWith("pasted-");
        const isAtoU = intent.endsWith("-AtoU");
        const [from, to] = isAtoU ? [lbs.available, lbs.useThese] : [lbs.useThese, lbs.available];
        const summary = isAtoU ? "Available → Use These" : "Use These → Available";

        let total;
        if (isPasted) {
          const values = parseLinesOrCsv(textarea.value || "");
          if (!values.length) {
            status.classList.add("is-warn");
            status.textContent = "Paste items first.";
            return;
          }
          // Estimate total (values that exist in source)
          const sourceTexts = new Set(BulkMoveOperations.getItemsArray(from).map((x) => x.text.toLowerCase()));
          total = values.filter((v) => sourceTexts.has(normalizeText(v).toLowerCase())).length;
        } else {
          total = BulkMoveOperations.getItemsArray(from).length;
        }

        if (total === 0) {
          status.classList.add("is-warn");
          status.textContent = "Nothing to move.";
          return;
        }

        InlineDashboard.setBusy(runBtn, true, `(0/${total}) Moving...`);
        const onProgress = ({ completed, total }) => {
          runBtn.textContent = `(${completed}/${total}) Moving...`;
        };

        let moved;
        if (isPasted) {
          const values = parseLinesOrCsv(textarea.value || "");
          moved = BulkMoveOperations.moveItemsByText(from, to, values, { onProgress });
        } else {
          moved = BulkMoveOperations.moveAll(from, to, { onProgress });
        }

        InlineDashboard.setBusy(runBtn, false);
        status.classList.remove("is-info", "is-warn", "is-error");
        status.classList.add("is-info");
        status.textContent = `Moved ${moved} item(s) ${summary}.`;
        runValidation();
      },
    });
    return dashboard;
  }

  // ============================================================
  // BUTTON INJECTION
  // ============================================================

  const ButtonInjector = {
    // Track injected dashboards so handlers don't double-fire and we can
    // hide siblings when one opens.
    _injected: new WeakMap(), // addBtn -> { addPanel, removePanel }

    createMatchingButton(refBtn, text, onClick) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = text;
      btn.className = refBtn?.className || "btn btn--small melon-green";
      btn.removeAttribute("style");
      btn.dataset.bulkInjected = "1";
      if (typeof onClick === "function") btn.addEventListener("click", onClick);
      return btn;
    },

    injectForTargetType(handlerName, typeKey, pretty) {
      const addBtn = document.querySelector(`button[onclick="${handlerName}()"]`);
      if (!addBtn) return;
      if (addBtn.dataset.bulkPanelInjected) return;

      // Normalize style of any pre-existing Bulk Add/Remove buttons
      const siblings = Array.from(addBtn.parentElement?.children || []);
      for (const el of siblings) {
        if (el?.tagName === "BUTTON") {
          const text = normalizeText(el.textContent);
          if (text === "Bulk Add" || text === "Bulk Remove") {
            el.className = addBtn.className;
            el.removeAttribute("style");
          }
        }
      }

      const addPanel = buildAddPanel(typeKey, pretty);
      const removePanel = buildRemovePanel(typeKey, pretty);

      // Mount the panels after the entire input/button row
      const rowParent = addBtn.parentElement;
      const mountAnchor = rowParent || addBtn;
      insertAfter(addPanel.panel, mountAnchor);
      insertAfter(removePanel.panel, addPanel.panel);

      // Inject Bulk Remove + Bulk Add buttons (Bulk Add inserted last so it
      // ends up adjacent to the Add button).
      if (!addBtn.dataset.bulkRemoveInjected) {
        const bulkRemoveBtn = this.createMatchingButton(addBtn, "Bulk Remove", () => {
          addPanel.hide();
          removePanel.toggle();
        });
        insertAfter(bulkRemoveBtn, addBtn);
        addBtn.dataset.bulkRemoveInjected = "1";
      }
      if (!addBtn.dataset.bulkAddInjected) {
        const bulkAddBtn = this.createMatchingButton(addBtn, "Bulk Add", () => {
          removePanel.hide();
          addPanel.toggle();
        });
        insertAfter(bulkAddBtn, addBtn);
        addBtn.dataset.bulkAddInjected = "1";
      }

      addBtn.dataset.bulkPanelInjected = "1";
      this._injected.set(addBtn, { addPanel, removePanel });
    },

    injectTargetButtons() {
      this.injectForTargetType("NewTargetCounty", "county", "Counties");
      this.injectForTargetType("NewTargetCity", "city", "Cities");
      this.injectForTargetType("NewTargetZip", "zip", "Zip Codes");
    },

    /**
     * Inject Bulk Move inline panel for MelonMax page.
     */
    injectBulkMoveButton_MelonMax() {
      if (!PageDetector.isMelonMaxBudgetDetails) return;
      if (document.getElementById("patch-bulk-move-melonmax")) return;
      const header = BulkMoveOperations.findZipDragDropHeader_MelonMax();
      if (!header) return;

      const dashboard = buildMovePanel({
        id: "patch-bulk-move-melonmax",
        titleProvider: "Bulk Move Targeting",
        pairGetter: () => BulkMoveOperations.getListBoxPairBySelectIds("UpdateBudgetTargetId", "UpdateListbox2"),
      });

      const triggerBtn = document.createElement("button");
      triggerBtn.type = "button";
      triggerBtn.id = "patchBulkMoveZipsBtn";
      triggerBtn.className = "btn btn--small melon-green";
      triggerBtn.style.marginLeft = "10px";
      triggerBtn.style.marginTop = "10px";
      triggerBtn.textContent = "Bulk Move";
      triggerBtn.addEventListener("click", () => dashboard.toggle());

      try {
        header.insertAdjacentElement("afterend", triggerBtn);
        insertAfter(dashboard.panel, triggerBtn);
      } catch (error) {
        logError("MelonMax inline injection failed:", error);
        try {
          header.parentElement?.appendChild(triggerBtn);
          header.parentElement?.appendChild(dashboard.panel);
        } catch (e2) {
          logError("MelonMax fallback injection failed:", e2);
        }
      }
    },

    /**
     * Inject Bulk Move inline panels for Agents BudgetDetails (one or two
     * areas).
     */
    injectBulkMoveButtons_AgentsBudgetDetails() {
      if (!PageDetector.isAgentsBudgetDetails) return;

      const wireOne = ({ headerId, exampleId, btnId, panelId, sourceSelectId, destSelectId, dropdownId, suffix }) => {
        const header = document.getElementById(headerId);
        const example = document.getElementById(exampleId);
        if (!header || !example) return;
        if (document.getElementById(btnId)) return;

        const titleProvider = ($) => {
          try {
            const ddl = $?.(`#${dropdownId}`)?.data?.("kendoDropDownList");
            const raw = ddl ? String(ddl.text() || "").trim() : "Targeting";
            const label = raw.charAt(0).toUpperCase() + raw.slice(1).toLowerCase();
            return `Bulk Move ${label}${suffix ? " " + suffix : ""}`;
          } catch {
            return "Bulk Move Targeting";
          }
        };

        const dashboard = buildMovePanel({
          id: panelId,
          titleProvider,
          pairGetter: () => BulkMoveOperations.getListBoxPairBySelectIds(sourceSelectId, destSelectId),
        });

        const btn = document.createElement("button");
        btn.type = "button";
        btn.id = btnId;
        btn.className = "btn btn--small melon-green";
        btn.style.marginLeft = "10px";
        btn.textContent = `Bulk Move${suffix ? " " + suffix : ""}`;
        btn.addEventListener("click", () => dashboard.toggle());

        header.insertAdjacentElement("afterend", btn);
        insertAfter(dashboard.panel, btn);
      };

      wireOne({
        headerId: "UpdateTargetTypeHeader",
        exampleId: "UpdateExample",
        btnId: "patchBulkMoveZipsBtn_update1",
        panelId: "patch-bulk-move-update1",
        sourceSelectId: "UpdateBudgetTargetId",
        destSelectId: "UpdateListbox2",
        dropdownId: "UpdateBudgetTargetTypeId",
        suffix: "",
      });

      const container2 = document.getElementById("UpdateTargetTypeContainer2");
      const visible2 = container2 && container2.style.display !== "none" && container2.offsetParent !== null;
      if (visible2) {
        wireOne({
          headerId: "UpdateTargetTypeHeader2",
          exampleId: "UpdateExample2",
          btnId: "patchBulkMoveZipsBtn_update2",
          panelId: "patch-bulk-move-update2",
          sourceSelectId: "UpdateBudgetTargetId2",
          destSelectId: "UpdateListbox22",
          dropdownId: "UpdateBudgetTargetTypeId2",
          suffix: "(2nd Target)",
        });
      }
    },

    injectAllInlineUI() {
      this.injectTargetButtons();
      this.injectBulkMoveButton_MelonMax();
      this.injectBulkMoveButtons_AgentsBudgetDetails();
    },
  };

  // ============================================================
  // INITIALIZATION
  // ============================================================

  const AppController = {
    mutationObserver: null,
    beforeUnloadHandler: null,
    _waiting: false,

    async waitForEditTargetsButtons() {
      if (this._waiting) return;
      this._waiting = true;
      try {
        for (let i = 0; i < TIMING.MAX_WAIT_ITERATIONS; i++) {
          const any = ["NewTargetCounty", "NewTargetCity", "NewTargetZip"]
            .some((h) => document.querySelector(`button[onclick="${h}()"]`));
          if (any) {
            ButtonInjector.injectTargetButtons();
            return;
          }
          await sleep(TIMING.WAIT_ITERATION_DELAY_MS);
        }
      } finally {
        this._waiting = false;
      }
    },

    tick() {
      this.waitForEditTargetsButtons().catch((e) => logError("waitForEditTargetsButtons:", e));
      ButtonInjector.injectBulkMoveButton_MelonMax();
      ButtonInjector.injectBulkMoveButtons_AgentsBudgetDetails();
    },

    handleMutations: debounce(function () {
      ButtonInjector.injectAllInlineUI();
    }, 300),

    init() {
      log("Initializing", VERSION);
      injectGlobalStyles();
      this.tick();

      this.mutationObserver = new MutationObserver(() => this.handleMutations());
      this.mutationObserver.observe(document.documentElement || document.body, {
        childList: true,
        subtree: true,
      });

      Object.assign(_debug, {
        BulkAddOperations,
        BulkRemoveOperations,
        BulkMoveOperations,
        ButtonInjector,
        PageDetector,
        ApiClient,
        InlineDashboard,
        ZipGeoData,
        CoverageAnalyzer,
      });
      window.PatchTargetingHelper = { version: VERSION, _debug };

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
