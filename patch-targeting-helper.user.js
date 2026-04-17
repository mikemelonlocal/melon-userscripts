// ==UserScript==
// @name         Patch Targeting Helper – Bulk Add + Bulk Remove (Targets) + Bulk Move (BudgetDetails ListBoxes)
// @namespace    http://tampermonkey.net/
// @version      3.0.0
// @description  Bulk Add + Bulk Remove for Edit Advertising Targets (County/City/Zip) + Bulk Move for Kendo ListBoxes on BudgetDetails screens (County, Zip, State, City, DMA).
// @match        https://thepatch.melonlocal.com/*
// @run-at       document-end
// @grant        none
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/patch-targeting-helper.user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/patch-targeting-helper.user.js
// ==/UserScript==

(function () {
  "use strict";

  // ============================================================
  // CONSTANTS
  // ============================================================

  // Keep in sync with @version in the userscript header above.
  const VERSION = "patch-targeting-helper-bulk-v3.0.0";
  const DEBUG = false; // Set to true to enable detailed console logging

  // Shared mutable state so window.PatchTargetingHelperDebug works before init().
  const _debug = {
    DEBUG_MODE: DEBUG,
  };

  // Expose a way to toggle debug mode and manually trigger button injection
  window.PatchTargetingHelperDebug = {
    enableDebug: () => { _debug.DEBUG_MODE = true; },
    disableDebug: () => { _debug.DEBUG_MODE = false; },
    injectButtons: () => {
      ButtonInjector.injectBulkMoveButton_MelonMax();
      ButtonInjector.injectBulkMoveButtons_AgentsBudgetDetails();
      ButtonInjector.injectTargetButtons();
      console.log("[PatchTargetingHelper] Manual button injection triggered");
    }
  };

  const TIMING = {
    DEFAULT_DELAY_MS: 120,
    MAX_WAIT_ITERATIONS: 200,
    WAIT_ITERATION_DELAY_MS: 150,
    FOCUS_DELAY_MS: 0,
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

  const ELEMENT_IDS = {
    BULK_TEXTAREA: "patchBulkTextarea",
    BULK_REMOVE_TEXTAREA: "patchBulkRemoveTextarea",
    BULK_MOVE_TEXTAREA: "patchBulkMoveZipsTextarea",
    BULK_MOVE_BTN: "patchBulkMoveZipsBtn",
    BULK_MOVE_BTN_UPDATE1: "patchBulkMoveZipsBtn_update1",
    BULK_MOVE_BTN_UPDATE2: "patchBulkMoveZipsBtn_update2",
  };

  const TARGET_TYPES = {
    COUNTY: {
      key: "county",
      inputId: "newTargetCounty",
      handlerName: "NewTargetCounty",
      pretty: "Counties",
    },
    CITY: {
      key: "city",
      inputId: "newTargetCity",
      handlerName: "NewTargetCity",
      pretty: "Cities",
    },
    ZIP: {
      key: "zip",
      inputId: "newTargetZip",
      handlerName: "NewTargetZip",
      pretty: "Zip Codes",
    },
  };

  // ============================================================
  // UTILITY FUNCTIONS
  // ============================================================

  /**
   * Sleep utility for async operations
   * @param {number} ms - Milliseconds to sleep
   * @returns {Promise<void>}
   */
  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  /**
   * Debounce function calls
   * @param {Function} func - Function to debounce
   * @param {number} wait - Wait time in ms
   * @returns {Function}
   */
  function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
      const later = () => {
        clearTimeout(timeout);
        func(...args);
      };
      clearTimeout(timeout);
      timeout = setTimeout(later, wait);
    };
  }

  /**
   * Normalize whitespace in text
   * @param {string} s - Text to normalize
   * @returns {string}
   */
  function normalizeText(s) {
    return String(s || "").replace(/\s+/g, " ").trim();
  }

  /**
   * Parse lines or CSV, removing duplicates (case-insensitive)
   * @param {string} text - Text to parse
   * @returns {string[]}
   */
  function parseLinesOrCsv(text) {
    const raw = String(text || "")
      .split(/[\n,]+/g)
      .map((s) => normalizeText(s))
      .filter(Boolean);

    const seen = new Set();
    const out = [];
    for (const v of raw) {
      const k = v.toLowerCase();
      if (!seen.has(k)) {
        seen.add(k);
        out.push(v);
      }
    }
    return out;
  }

  /**
   * Parse and validate zip codes
   * @param {string} text - Text containing zip codes
   * @returns {string[]}
   */
  function parseZips(text) {
    // Convert to string and normalize line endings
    const normalized = String(text || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");

    // Split on whitespace (including newlines) and commas
    const raw = normalized
      .split(/[\s,]+/)
      .map((s) => s.trim())
      .filter(Boolean);

    // Accept ZIP+4 like "12345-6789" by keeping only the leading 5 digits.
    const cleaned = raw
      .map((z) => z.split("-")[0].replace(/[^\d]/g, ""))
      .filter((z) => /^\d{5}$/.test(z));

    // Deduplicate
    const seen = new Set();
    const out = [];
    for (const z of cleaned) {
      if (!seen.has(z)) {
        seen.add(z);
        out.push(z);
      }
    }

    // Only log errors when parsing fails
    if (out.length === 0 && text && text.trim()) {
      logError("parseZips: No valid zips found", {
        inputLength: text.length,
        inputSample: text.substring(0, 50),
        rawTokens: raw.length,
        cleanedTokens: cleaned.length
      });
    }

    return out;
  }

  /**
   * Set native input value and trigger events
   * @param {HTMLElement} el - Input element
   * @param {string} value - Value to set
   */
  function setNativeValue(el, value) {
    try {
      const proto = Object.getPrototypeOf(el);
      const desc = proto ? Object.getOwnPropertyDescriptor(proto, "value") : null;
      if (desc?.set) {
        desc.set.call(el, value);
      } else {
        el.value = value;
      }

      el.dispatchEvent(new Event("input", { bubbles: true }));
      el.dispatchEvent(new Event("change", { bubbles: true }));
    } catch (error) {
      console.error("[PatchTargetingHelper] Error setting native value:", error);
      el.value = value;
    }
  }

  /**
   * Insert element after reference node
   * @param {Node} newNode - Node to insert
   * @param {Node} referenceNode - Reference node
   */
  function insertAfter(newNode, referenceNode) {
    const parent = referenceNode?.parentNode;
    if (!parent) return;
    parent.insertBefore(newNode, referenceNode.nextSibling);
  }

  /**
   * Log with namespace prefix (only if DEBUG is enabled)
   * @param {string} message - Message to log
   * @param {...any} args - Additional arguments
   */
  function log(message, ...args) {
    if (DEBUG || _debug.DEBUG_MODE) {
      console.log(`[PatchTargetingHelper] ${message}`, ...args);
    }
  }

  /**
   * Log error with namespace prefix (always logs)
   * @param {string} message - Error message
   * @param {...any} args - Additional arguments
   */
  function logError(message, ...args) {
    console.error(`[PatchTargetingHelper] ${message}`, ...args);
  }

  // ============================================================
  // PAGE DETECTION
  // ============================================================

  const PageDetector = {
    hrefLower: String(location.href || "").toLowerCase(),
    hostLower: String(location.hostname || "").toLowerCase(),

    get isPatch() {
      return this.hostLower === "thepatch.melonlocal.com";
    },

    get isMelonMaxBudgetDetails() {
      return (
        this.isPatch &&
        (this.hrefLower.includes("/melonmax/melonmaxbudgetdetails") ||
          this.hrefLower.includes("/melonmaxbudgetaddtargetingpartial") ||
          this.hrefLower.includes("/agents/melonmaxbudgetaddtargetingpartial"))
      );
    },

    get isAgentsBudgetDetails() {
      return this.isPatch && this.hrefLower.includes("/agents/budgetdetails");
    },
  };

  // ============================================================
  // GLOBAL STYLES
  // ============================================================

  function injectGlobalStyles() {
    const styleTag = document.createElement("style");
    styleTag.textContent = `
      #${ELEMENT_IDS.BULK_TEXTAREA},
      #${ELEMENT_IDS.BULK_REMOVE_TEXTAREA},
      #${ELEMENT_IDS.BULK_MOVE_TEXTAREA} {
        pointer-events: auto !important;
        z-index: 2147483647 !important;
        position: relative;
      }
    `;
    document.head.appendChild(styleTag);
  }

  // ============================================================
  // FOCUS TRAP MITIGATION
  // ============================================================

  const FocusManager = {
    bulkFocusinBlocker: null,
    dashFocusTrapRestore: null,
    bulkModalRemovalObserver: null,

    /**
     * Get Bootstrap modal instance from element
     * @param {HTMLElement} modalEl - Modal element
     * @returns {object|null}
     */
    getBootstrapModalInstance(modalEl) {
      if (!modalEl) return null;

      const b = window.bootstrap;
      if (b?.Modal?.getInstance) {
        try {
          return b.Modal.getInstance(modalEl);
        } catch (error) {
          logError("Error getting Bootstrap modal instance:", error);
        }
      }

      try {
        const props = Object.getOwnPropertyNames(modalEl);
        for (const k of props) {
          const v = modalEl[k];
          if (
            v &&
            typeof v === "object" &&
            v.constructor &&
            String(v.constructor.name).toLowerCase().includes("modal")
          ) {
            return v;
          }
        }
      } catch (error) {
        logError("Error finding modal instance via property inspection:", error);
      }

      return modalEl.__bs_modal || modalEl._bsModal || null;
    },

    /**
     * Try to deactivate dashboard focus trap
     * @returns {Function|null} - Restore function or null
     */
    tryDeactivateDashboardFocusTrap() {
      const dashEl = document.getElementById("dashboardModal");
      if (!dashEl) return null;

      const inst = this.getBootstrapModalInstance(dashEl);
      if (!inst) return null;

      const ft = inst._focustrap;
      if (!ft || typeof ft.deactivate !== "function" || typeof ft.activate !== "function") {
        return null;
      }

      try {
        ft.deactivate();
        return () => {
          try {
            ft.activate();
          } catch (error) {
            logError("Error reactivating focus trap:", error);
          }
        };
      } catch (error) {
        logError("Error deactivating focus trap:", error);
        return null;
      }
    },

    /**
     * Install focusin bypass for bulk modal
     * @param {HTMLElement} bulkModalEl - Bulk modal element
     */
    installFocusinBypassForBulkModal(bulkModalEl) {
      if (!bulkModalEl || this.bulkFocusinBlocker) return;

      this.bulkFocusinBlocker = function (e) {
        try {
          if (bulkModalEl.contains(e.target)) {
            e.stopImmediatePropagation();
          }
        } catch (error) {
          logError("Error in focusin blocker:", error);
        }
      };

      document.addEventListener("focusin", this.bulkFocusinBlocker, true);
    },

    /**
     * Uninstall focusin bypass
     */
    uninstallFocusinBypass() {
      if (!this.bulkFocusinBlocker) return;
      try {
        document.removeEventListener("focusin", this.bulkFocusinBlocker, true);
      } catch (error) {
        logError("Error removing focusin listener:", error);
      }
      this.bulkFocusinBlocker = null;
    },

    /**
     * Suspend dashboard focus management
     * @param {HTMLElement} bulkModalEl - Bulk modal element
     */
    suspendDashboardFocusManagement(bulkModalEl) {
      // Fail-safe: avoid stacking focus overrides
      this.restoreDashboardFocusManagement();

      this.dashFocusTrapRestore = this.tryDeactivateDashboardFocusTrap();
      this.installFocusinBypassForBulkModal(bulkModalEl);

      // Auto-restore if modal is removed without our close handlers firing
      try {
        if (this.bulkModalRemovalObserver) {
          this.bulkModalRemovalObserver.disconnect();
        }
      } catch (error) {
        logError("Error disconnecting observer:", error);
      }

      this.bulkModalRemovalObserver = new MutationObserver(() => {
        try {
          if (!bulkModalEl || !document.body || !document.body.contains(bulkModalEl)) {
            this.restoreDashboardFocusManagement();
          }
        } catch (error) {
          logError("Error in mutation observer:", error);
        }
      });

      // Observe only the modal's parent — watching document.body with subtree
      // fires on every descendant mutation in the page.
      const observeTarget = bulkModalEl?.parentNode || document.body;
      try {
        this.bulkModalRemovalObserver.observe(observeTarget, {
          childList: true,
        });
      } catch (error) {
        logError("Error starting mutation observer:", error);
      }
    },

    /**
     * Restore dashboard focus management
     */
    restoreDashboardFocusManagement() {
      try {
        if (this.bulkModalRemovalObserver) {
          this.bulkModalRemovalObserver.disconnect();
        }
      } catch (error) {
        logError("Error disconnecting observer during restore:", error);
      }
      this.bulkModalRemovalObserver = null;

      this.uninstallFocusinBypass();
      if (this.dashFocusTrapRestore) {
        try {
          this.dashFocusTrapRestore();
        } catch (error) {
          logError("Error restoring focus trap:", error);
        }
        this.dashFocusTrapRestore = null;
      }
    },
  };

  // ============================================================
  // BULK ADD OPERATIONS
  // ============================================================

  const BulkAddOperations = {
    /**
     * Run bulk add operation
     * @param {string} typeKey - Type of target (county, city, zip)
     * @param {string} rawText - Raw text input
     * @param {number} delayMs - Delay between operations
     * @returns {Promise<void>}
     */
    async run(typeKey, rawText, delayMs = TIMING.DEFAULT_DELAY_MS) {
      const targetType = this.getTargetTypeConfig(typeKey);
      if (!targetType) {
        throw new Error(`Unknown target type: ${typeKey}`);
      }

      const { inputId, handlerName, pretty } = targetType;

      const input = document.getElementById(inputId);
      const handler = window[handlerName];

      if (!input || typeof handler !== "function") {
        throw new Error(`Cannot find input or handler for ${pretty}`);
      }

      log(`Bulk add started for ${typeKey}`, {
        rawTextLength: rawText?.length,
        rawTextSample: rawText?.substring(0, 100),
        typeKey
      });

      const values = typeKey === "zip" ? parseZips(rawText) : parseLinesOrCsv(rawText);

      log(`Parsed values:`, {
        count: values.length,
        sample: values.slice(0, 5)
      });

      if (!values.length) {
        const msg = `No valid ${pretty} found in the input.`;
        alert(DEBUG
          ? `${msg}\n\nDebug: Input had ${rawText?.length || 0} chars.\nSample: "${rawText?.substring(0, 50)}"\n\nCheck console for details.`
          : msg
        );
        logError(`No valid items parsed for ${typeKey}`, {
          rawTextLength: rawText?.length,
          rawTextType: typeof rawText,
          isEmpty: !rawText || !rawText.trim()
        });
        return;
      }

      let completed = 0;

      for (let i = 0; i < values.length; i++) {
        const val = values[i];
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
      }

      const msg = `Bulk Add Complete.\nProcessed: ${completed} of ${values.length} ${pretty}`;
      alert(msg);
      log(msg);
    },

    /**
     * Get target type configuration
     * @param {string} typeKey - Type key
     * @returns {object|null}
     */
    getTargetTypeConfig(typeKey) {
      const typeMap = {
        county: TARGET_TYPES.COUNTY,
        city: TARGET_TYPES.CITY,
        zip: TARGET_TYPES.ZIP,
      };
      return typeMap[typeKey] || null;
    },
  };

  // ============================================================
  // BULK REMOVE OPERATIONS
  // ============================================================

  const BulkRemoveOperations = {
    /**
     * Run bulk remove operation
     * @param {string} typeKey - Type of target
     * @param {string} rawText - Raw text input
     * @param {number} delayMs - Delay between operations
     * @returns {Promise<void>}
     */
    async run(typeKey, rawText, delayMs = TIMING.DEFAULT_DELAY_MS) {
      const targetType = BulkAddOperations.getTargetTypeConfig(typeKey);
      if (!targetType) {
        throw new Error(`Unknown target type: ${typeKey}`);
      }

      const { pretty } = targetType;

      log(`Bulk remove started for ${typeKey}`, {
        rawTextLength: rawText?.length,
        rawTextSample: rawText?.substring(0, 100)
      });

      const values = typeKey === "zip" ? parseZips(rawText) : parseLinesOrCsv(rawText);

      log(`Parsed values for removal:`, {
        count: values.length,
        sample: values.slice(0, 5)
      });

      if (!values.length) {
        alert(`No valid ${pretty} found in the input.`);
        logError(`No valid items parsed for removal: ${typeKey}`, {
          rawTextLength: rawText?.length
        });
        return;
      }

      const normalizedValues = values.map((v) => normalizeText(v).toLowerCase());

      const rows = Array.from(document.querySelectorAll("#exampleTable tbody tr"));
      let removed = 0;

      for (const row of rows) {
        try {
          const cols = row.querySelectorAll("td");
          if (cols.length < 2) continue;

          const cellText = normalizeText(cols[1].textContent).toLowerCase();
          if (normalizedValues.includes(cellText)) {
            const deleteLink = row.querySelector('a[onclick*="DeleteAdvertisingTarget"]');
            if (deleteLink) {
              deleteLink.click();
              await sleep(delayMs);
              removed++;
            }
          }
        } catch (error) {
          logError("Error removing row:", error);
        }
      }

      const msg = `Bulk Remove Complete.\nRemoved: ${removed} ${pretty}`;
      alert(msg);
      log(msg);
    },
  };

  // ============================================================
  // BULK MOVE OPERATIONS (Kendo ListBox)
  // ============================================================

  const BulkMoveOperations = {
    /**
     * Find Kendo ListBox pair by select IDs
     * @param {string} sourceSelectId - Source select ID
     * @param {string} destSelectId - Destination select ID
     * @returns {object|null}
     */
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

    /**
     * Get text from ListBox item respecting dataTextField configuration
     * @param {object} item - ListBox item
     * @param {object} listBox - Kendo ListBox instance
     * @returns {string}
     */
    getItemText(item, listBox) {
      if (!item) return "";

      // Try to get the configured dataTextField
      let dataTextField = null;
      try {
        if (listBox?.options?.dataTextField) {
          dataTextField = listBox.options.dataTextField;
        }
      } catch (error) {
        logError("Error getting dataTextField:", error);
      }

      // If dataTextField is configured and exists on item, use it
      if (dataTextField && item[dataTextField] != null) {
        return String(item[dataTextField]);
      }

      // Fallback: check common property names
      if (item.County != null) return String(item.County);
      if (item.Zip != null) return String(item.Zip);
      if (item.State != null) return String(item.State);
      if (item.City != null) return String(item.City);
      if (item.DMA != null) return String(item.DMA);
      if (item.text != null) return String(item.text);
      if (item.Text != null) return String(item.Text);
      if (item.TargetName != null) return String(item.TargetName);
      if (item.value != null) return String(item.value);
      if (item.Value != null) return String(item.Value);

      return "";
    },

    /**
     * Get value from ListBox item respecting dataValueField configuration
     * @param {object} item - ListBox item
     * @param {object} listBox - Kendo ListBox instance
     * @returns {string}
     */
    getItemValue(item, listBox) {
      if (!item) return "";

      // Try to get the configured dataValueField
      let dataValueField = null;
      try {
        if (listBox?.options?.dataValueField) {
          dataValueField = listBox.options.dataValueField;
        }
      } catch (error) {
        logError("Error getting dataValueField:", error);
      }

      // If dataValueField is configured and exists on item, use it
      if (dataValueField && item[dataValueField] != null) {
        return String(item[dataValueField]);
      }

      // Fallback: check common property names
      if (item.value != null) return String(item.value);
      if (item.Value != null) return String(item.Value);
      if (item.Zip != null) return String(item.Zip);
      if (item.County != null) return String(item.County);
      if (item.State != null) return String(item.State);
      if (item.City != null) return String(item.City);
      if (item.DMA != null) return String(item.DMA);

      return "";
    },

    /**
     * Get items as array with text and value
     * @param {object} lb - Kendo ListBox
     * @returns {Array}
     */
    getItemsArray(lb) {
      try {
        const ds = lb.dataSource;
        if (!ds?.data) return [];
        const data = ds.data();
        return data.map((item) => ({
          text: normalizeText(this.getItemText(item, lb)),
          value: this.getItemValue(item, lb),
          raw: item, // Keep the original item
        }));
      } catch (error) {
        logError("Error getting items array:", error);
        return [];
      }
    },

    /**
     * Move items by text from source to destination
     * @param {object} sourceLb - Source ListBox
     * @param {object} destLb - Destination ListBox
     * @param {string[]} textArray - Array of text values to move
     * @returns {number} Number of items moved
     */
    moveItemsByText(sourceLb, destLb, textArray) {
      try {
        const normalized = textArray.map((t) => normalizeText(t).toLowerCase());
        const sourceItems = this.getItemsArray(sourceLb);
        const toMove = sourceItems.filter((item) =>
          normalized.includes(item.text.toLowerCase())
        );

        if (!toMove.length) return 0;

        const sourceDs = sourceLb.dataSource;
        const destDs = destLb.dataSource;

        for (const item of toMove) {
          try {
            // Add the raw item (preserves all properties)
            destDs.add(item.raw);
            // Find and remove from source
            const found = sourceDs.data().find((d) =>
              this.getItemValue(d, sourceLb) === item.value
            );
            if (found) sourceDs.remove(found);
          } catch (error) {
            logError("Error moving item:", item.text, error);
          }
        }

        return toMove.length;
      } catch (error) {
        logError("Error in moveItemsByText:", error);
        return 0;
      }
    },

    /**
     * Move all items from source to destination
     * @param {object} sourceLb - Source ListBox
     * @param {object} destLb - Destination ListBox
     * @returns {number} Number of items moved
     */
    moveAll(sourceLb, destLb) {
      try {
        const sourceItems = this.getItemsArray(sourceLb);
        if (!sourceItems.length) return 0;

        const sourceDs = sourceLb.dataSource;
        const destDs = destLb.dataSource;

        for (const item of sourceItems) {
          try {
            // Add the raw item (preserves all properties)
            destDs.add(item.raw);
          } catch (error) {
            logError("Error adding item to destination:", item.text, error);
          }
        }

        try {
          const allData = sourceDs.data().slice();
          for (const d of allData) {
            sourceDs.remove(d);
          }
        } catch (error) {
          logError("Error removing items from source:", error);
        }

        return sourceItems.length;
      } catch (error) {
        logError("Error in moveAll:", error);
        return 0;
      }
    },

    /**
     * Find zip drag-drop header for MelonMax
     * @returns {HTMLElement|null}
     */
    findZipDragDropHeader_MelonMax() {
      // First, try to find it inside the visible modal
      const modal = document.querySelector('.modal.show, #dashboardModal');

      if (modal) {
        // Look for h6 specifically first (that's what we want)
        const h6Candidates = Array.from(modal.querySelectorAll("h6"));

        for (const el of h6Candidates) {
          const text = normalizeText(el.textContent);
          // Match "Choose Zip (Drag and Drop)" specifically
          if (text.includes("Choose") && text.includes("Zip") && text.length < 50) {
            log("Found h6 header inside modal:", {
              tag: el.tagName,
              text: text,
              classes: el.className
            });
            return el;
          }
        }

        // Fallback to other headers
        const otherCandidates = Array.from(
          modal.querySelectorAll("h5, h4, h3, label")
        );

        for (const el of otherCandidates) {
          // Skip if too many children (likely a container)
          if (el.children.length > 3) continue;

          const text = normalizeText(el.textContent);
          if (
            text.includes("Choose") &&
            text.includes("Zip") &&
            text.length < 50
          ) {
            log("Found header inside modal:", {
              tag: el.tagName,
              text: text,
              classes: el.className
            });
            return el;
          }
        }
      }

      log("Could not find drag-drop header in modal");
      return null;
    },
  };

  // ============================================================
  // MODAL BUILDER
  // ============================================================

  const ModalBuilder = {
    /**
     * Create a Bootstrap-style modal
     * @param {object} config - Modal configuration
     * @returns {object} Modal elements
     */
    createModal(config) {
      const {
        id,
        title,
        textareaId,
        textareaPlaceholder,
        onClose,
        ariaLabel = title,
      } = config;

      const modal = document.createElement("div");
      modal.className = "modal fade show";
      modal.id = id;
      modal.style.display = "block";
      modal.style.backgroundColor = "rgba(0,0,0,0.5)";
      modal.setAttribute("role", "dialog");
      modal.setAttribute("aria-modal", "true");
      modal.setAttribute("aria-labelledby", `${id}-title`);

      const dialog = document.createElement("div");
      dialog.className = "modal-dialog";
      dialog.style.marginTop = "50px";
      dialog.setAttribute("role", "document");

      const content = document.createElement("div");
      content.className = "modal-content";
      content.style.border = `1px solid ${SURFACE.border}`;

      // Header
      const header = document.createElement("div");
      header.className = "modal-header";
      header.style.backgroundColor = COLORS.alpine;
      header.style.borderBottom = `1px solid ${SURFACE.border}`;

      const h5 = document.createElement("h5");
      h5.className = "modal-title";
      h5.id = `${id}-title`;
      h5.textContent = title;

      const closeBtn = document.createElement("button");
      closeBtn.type = "button";
      closeBtn.className = "close";
      closeBtn.innerHTML = "&times;";
      closeBtn.setAttribute("aria-label", "Close");
      closeBtn.style.background = "none";
      closeBtn.style.border = "none";
      closeBtn.style.fontSize = "1.5rem";
      closeBtn.style.cursor = "pointer";

      header.appendChild(h5);
      header.appendChild(closeBtn);

      // Body
      const body = document.createElement("div");
      body.className = "modal-body";

      const textarea = document.createElement("textarea");
      textarea.id = textareaId;
      textarea.className = "form-control";
      textarea.rows = 10;
      textarea.placeholder = textareaPlaceholder;
      textarea.style.width = "100%";
      textarea.style.fontFamily = "monospace";
      textarea.setAttribute("aria-label", ariaLabel);

      body.appendChild(textarea);

      // Footer
      const footer = document.createElement("div");
      footer.className = "modal-footer";
      footer.style.borderTop = `1px solid ${SURFACE.border}`;

      content.appendChild(header);
      content.appendChild(body);
      content.appendChild(footer);
      dialog.appendChild(content);
      modal.appendChild(dialog);

      // ESC key handler declared first so closeHandler can unbind it.
      let escHandler;

      const closeHandler = () => {
        try {
          if (escHandler) {
            document.removeEventListener("keydown", escHandler);
            escHandler = null;
          }
          modal.remove();
          FocusManager.restoreDashboardFocusManagement();
          if (onClose) onClose();
        } catch (error) {
          logError("Error closing modal:", error);
        }
      };

      closeBtn.onclick = closeHandler;

      escHandler = (e) => {
        if (e.key === "Escape") closeHandler();
      };
      document.addEventListener("keydown", escHandler);

      return {
        modal,
        dialog,
        content,
        header,
        body,
        footer,
        textarea,
        closeBtn,
        close: closeHandler,
      };
    },

    /**
     * Create an action button
     * @param {string} text - Button text
     * @param {Function} onClick - Click handler
     * @param {string} variant - Button variant (primary, secondary, danger)
     * @returns {HTMLButtonElement}
     */
    createActionButton(text, onClick, variant = "primary") {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = text;
      btn.style.marginLeft = "5px";

      const variantStyles = {
        primary: {
          backgroundColor: COLORS.cactus,
          color: "#fff",
        },
        secondary: {
          backgroundColor: COLORS.sand,
          color: COLORS.coconut,
        },
        danger: {
          backgroundColor: COLORS.watermelonSugar,
          color: "#fff",
        },
      };

      const style = variantStyles[variant] || variantStyles.primary;
      Object.assign(btn.style, {
        ...style,
        border: "none",
        padding: "8px 16px",
        borderRadius: "4px",
        cursor: "pointer",
        fontSize: "14px",
      });

      btn.addEventListener("click", onClick);
      btn.addEventListener("mouseenter", () => {
        btn.style.opacity = "0.9";
      });
      btn.addEventListener("mouseleave", () => {
        btn.style.opacity = "1";
      });

      return btn;
    },
  };

  // ============================================================
  // BULK ADD DIALOG
  // ============================================================

  function openBulkAddDialog(typeKey) {
    const targetType = BulkAddOperations.getTargetTypeConfig(typeKey);
    if (!targetType) {
      logError("Unknown target type:", typeKey);
      return;
    }

    const { pretty } = targetType;

    const modalConfig = {
      id: "patchBulkAddModal",
      title: `Bulk Add ${pretty}`,
      textareaId: ELEMENT_IDS.BULK_TEXTAREA,
      textareaPlaceholder:
        typeKey === "zip"
          ? "Paste zip codes (one per line or comma-separated)..."
          : `Paste ${pretty.toLowerCase()} (one per line or comma-separated)...`,
      ariaLabel: `Bulk add ${pretty.toLowerCase()}`,
    };

    const { modal, footer, textarea, close } = ModalBuilder.createModal(modalConfig);

    const cancelBtn = ModalBuilder.createActionButton("Cancel", close, "secondary");
    const runBtn = ModalBuilder.createActionButton(
      `Add ${pretty}`,
      async () => {
        const rawText = textarea.value || "";
        close();
        try {
          await BulkAddOperations.run(typeKey, rawText);
        } catch (error) {
          logError("Bulk add error:", error);
          alert(`Error during bulk add: ${error.message}`);
        }
      },
      "primary"
    );

    footer.appendChild(cancelBtn);
    footer.appendChild(runBtn);

    document.body.appendChild(modal);
    FocusManager.suspendDashboardFocusManagement(modal);

    setTimeout(() => {
      try {
        textarea.focus();
      } catch (error) {
        logError("Error focusing textarea:", error);
      }
    }, TIMING.FOCUS_DELAY_MS);
  }

  // ============================================================
  // BULK REMOVE DIALOG
  // ============================================================

  function openBulkRemoveDialog(typeKey) {
    const targetType = BulkAddOperations.getTargetTypeConfig(typeKey);
    if (!targetType) {
      logError("Unknown target type:", typeKey);
      return;
    }

    const { pretty } = targetType;

    const modalConfig = {
      id: "patchBulkRemoveModal",
      title: `Bulk Remove ${pretty}`,
      textareaId: ELEMENT_IDS.BULK_REMOVE_TEXTAREA,
      textareaPlaceholder:
        typeKey === "zip"
          ? "Paste zip codes to remove (one per line or comma-separated)..."
          : `Paste ${pretty.toLowerCase()} to remove (one per line or comma-separated)...`,
      ariaLabel: `Bulk remove ${pretty.toLowerCase()}`,
    };

    const { modal, footer, textarea, close } = ModalBuilder.createModal(modalConfig);

    const cancelBtn = ModalBuilder.createActionButton("Cancel", close, "secondary");
    const runBtn = ModalBuilder.createActionButton(
      `Remove ${pretty}`,
      async () => {
        const rawText = textarea.value || "";
        close();
        try {
          await BulkRemoveOperations.run(typeKey, rawText);
        } catch (error) {
          logError("Bulk remove error:", error);
          alert(`Error during bulk remove: ${error.message}`);
        }
      },
      "danger"
    );

    footer.appendChild(cancelBtn);
    footer.appendChild(runBtn);

    document.body.appendChild(modal);
    FocusManager.suspendDashboardFocusManagement(modal);

    setTimeout(() => {
      try {
        textarea.focus();
      } catch (error) {
        logError("Error focusing textarea:", error);
      }
    }, TIMING.FOCUS_DELAY_MS);
  }

  // ============================================================
  // BULK MOVE DIALOG
  // ============================================================

  function openBulkMoveModal(pairGetter, titleGetter) {
    const getTitle = () => {
      try {
        if (typeof titleGetter === "function") {
          const result = titleGetter(window.jQuery);
          return typeof result === "string" ? result : "Bulk Move Targeting";
        }
        return titleGetter || "Bulk Move Targeting";
      } catch (error) {
        logError("Error getting modal title:", error);
        return "Bulk Move Targeting";
      }
    };

    const modalConfig = {
      id: "patchBulkMoveModal",
      title: getTitle(),
      textareaId: ELEMENT_IDS.BULK_MOVE_TEXTAREA,
      textareaPlaceholder: "Paste items to move (one per line or comma-separated)...",
      ariaLabel: "Bulk move targeting items",
    };

    const { modal, footer, textarea, close } = ModalBuilder.createModal(modalConfig);

    const cancelBtn = ModalBuilder.createActionButton("Cancel", close, "secondary");

    const NO_LISTBOXES_MSG = "Could not find the two listboxes. Make sure the targeting area is visible.";

    const makeMoveBtn = (label, mode, direction) =>
      ModalBuilder.createActionButton(
        label,
        () => {
          const lbs = pairGetter();
          if (!lbs) {
            alert(NO_LISTBOXES_MSG);
            close();
            return;
          }
          const [from, to] = direction === "AtoU"
            ? [lbs.available, lbs.useThese]
            : [lbs.useThese, lbs.available];
          const summary = direction === "AtoU"
            ? "Available → Use These"
            : "Use These → Available";

          let moved;
          if (mode === "pasted") {
            const values = parseLinesOrCsv(textarea.value || "");
            if (!values.length) {
              close();
              return;
            }
            moved = BulkMoveOperations.moveItemsByText(from, to, values);
          } else {
            moved = BulkMoveOperations.moveAll(from, to);
          }
          alert(`Moved ${moved} item(s) from ${summary}.`);
          close();
        },
        "primary"
      );

    footer.appendChild(cancelBtn);
    footer.appendChild(makeMoveBtn("Move (Pasted) Available → Use These", "pasted", "AtoU"));
    footer.appendChild(makeMoveBtn("Move (Pasted) Use These → Available", "pasted", "UtoA"));
    footer.appendChild(makeMoveBtn("Move ALL Available → Use These", "all", "AtoU"));
    footer.appendChild(makeMoveBtn("Move ALL Use These → Available", "all", "UtoA"));

    document.body.appendChild(modal);
    FocusManager.suspendDashboardFocusManagement(modal);

    setTimeout(() => {
      try {
        textarea.focus();
      } catch (error) {
        logError("Error focusing textarea:", error);
      }
    }, TIMING.FOCUS_DELAY_MS);
  }

  // ============================================================
  // BUTTON INJECTION
  // ============================================================

  const ButtonInjector = {
    /**
     * Create a button matching style of reference button
     * @param {HTMLElement} refBtn - Reference button
     * @param {string} text - Button text
     * @param {Function} onClick - Click handler
     * @returns {HTMLButtonElement}
     */
    createMatchingButton(refBtn, text, onClick) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = text;
      btn.className = refBtn?.className || "btn btn--small melon-green";
      btn.removeAttribute("style");
      btn.dataset.bulkInjected = "1";
      if (typeof onClick === "function") {
        btn.addEventListener("click", onClick);
      }
      return btn;
    },

    /**
     * Inject bulk buttons for a target type
     * @param {string} handlerName - Handler function name
     * @param {string} typeKey - Target type key
     */
    injectForTargetType(handlerName, typeKey) {
      const addBtn = document.querySelector(`button[onclick="${handlerName}()"]`);
      if (!addBtn) return;

      // Normalize existing bulk button styles
      const siblings = Array.from(addBtn.parentElement?.children || []);
      for (const el of siblings) {
        if (el && el.tagName === "BUTTON") {
          const text = normalizeText(el.textContent);
          if (text === "Bulk Add" || text === "Bulk Remove") {
            el.className = addBtn.className;
            el.removeAttribute("style");
          }
        }
      }

      // Inject Bulk Remove button
      if (!addBtn.dataset.bulkRemoveInjected) {
        const removeBtn = this.createMatchingButton(addBtn, "Bulk Remove", () =>
          openBulkRemoveDialog(typeKey)
        );
        insertAfter(removeBtn, addBtn);
        addBtn.dataset.bulkRemoveInjected = "1";
      }

      // Inject Bulk Add button
      if (!addBtn.dataset.bulkAddInjected) {
        const bulkAddBtn = this.createMatchingButton(addBtn, "Bulk Add", () =>
          openBulkAddDialog(typeKey)
        );
        insertAfter(bulkAddBtn, addBtn);
        addBtn.dataset.bulkAddInjected = "1";
      }
    },

    /**
     * Inject all target buttons
     */
    injectTargetButtons() {
      this.injectForTargetType("NewTargetCounty", "county");
      this.injectForTargetType("NewTargetCity", "city");
      this.injectForTargetType("NewTargetZip", "zip");
    },

    /**
     * Inject bulk move button for MelonMax
     */
    injectBulkMoveButton_MelonMax() {
      if (!PageDetector.isMelonMaxBudgetDetails) return;

      // Don't inject if button already exists
      if (document.getElementById(ELEMENT_IDS.BULK_MOVE_BTN)) return;

      const header = BulkMoveOperations.findZipDragDropHeader_MelonMax();
      if (!header) {
        log("Could not find drag-drop header for MelonMax bulk move button");
        return;
      }

      log("Found drag-drop header, injecting bulk move button", {
        headerTag: header.tagName,
        headerText: header.textContent.substring(0, 50)
      });

      const btn = document.createElement("button");
      btn.id = ELEMENT_IDS.BULK_MOVE_BTN;
      btn.type = "button";
      btn.className = "btn btn--small melon-green";
      btn.style.marginLeft = "10px";
      btn.style.marginTop = "10px";
      btn.textContent = "Bulk Move";

      btn.addEventListener("click", () =>
        openBulkMoveModal(
          () => BulkMoveOperations.getListBoxPairBySelectIds("UpdateBudgetTargetId", "UpdateListbox2"),
          () => "Bulk Move Targeting"
        )
      );

      // Try different insertion strategies
      try {
        // First try: insert after the header
        header.insertAdjacentElement("afterend", btn);
        log("Bulk move button injected successfully");
      } catch (error) {
        logError("Error injecting bulk move button:", error);
        // Fallback: try appending to parent
        try {
          header.parentElement?.appendChild(btn);
          log("Bulk move button injected (fallback method)");
        } catch (e2) {
          logError("Fallback injection also failed:", e2);
        }
      }
    },

    /**
     * Inject bulk move buttons for Agents Budget Details
     */
    injectBulkMoveButtons_AgentsBudgetDetails() {
      if (!PageDetector.isAgentsBudgetDetails) return;

      // First targeting area
      const header1 = document.getElementById("UpdateTargetTypeHeader");
      const example1 = document.getElementById("UpdateExample");

      if (header1 && example1 && !document.getElementById(ELEMENT_IDS.BULK_MOVE_BTN_UPDATE1)) {
        const btn1 = document.createElement("button");
        btn1.id = ELEMENT_IDS.BULK_MOVE_BTN_UPDATE1;
        btn1.type = "button";
        btn1.className = "btn btn--small melon-green";
        btn1.style.marginLeft = "10px";
        btn1.textContent = "Bulk Move";

        btn1.addEventListener("click", () =>
          openBulkMoveModal(
            () => BulkMoveOperations.getListBoxPairBySelectIds("UpdateBudgetTargetId", "UpdateListbox2"),
            ($) => {
              const ddl = $("#UpdateBudgetTargetTypeId").data("kendoDropDownList");
              const raw = ddl ? String(ddl.text() || "").trim() : "Targeting";
              const label = raw.charAt(0).toUpperCase() + raw.slice(1).toLowerCase();
              return `Bulk Move ${label}`;
            }
          )
        );

        header1.insertAdjacentElement("afterend", btn1);
      }

      // Second targeting area (if visible)
      const container2 = document.getElementById("UpdateTargetTypeContainer2");
      const header2 = document.getElementById("UpdateTargetTypeHeader2");
      const example2 = document.getElementById("UpdateExample2");

      const visible2 =
        container2 &&
        container2.style.display !== "none" &&
        container2.offsetParent !== null;

      if (
        visible2 &&
        header2 &&
        example2 &&
        !document.getElementById(ELEMENT_IDS.BULK_MOVE_BTN_UPDATE2)
      ) {
        const btn2 = document.createElement("button");
        btn2.id = ELEMENT_IDS.BULK_MOVE_BTN_UPDATE2;
        btn2.type = "button";
        btn2.className = "btn btn--small melon-green";
        btn2.style.marginLeft = "10px";
        btn2.textContent = "Bulk Move (2nd Target)";

        btn2.addEventListener("click", () =>
          openBulkMoveModal(
            () => BulkMoveOperations.getListBoxPairBySelectIds("UpdateBudgetTargetId2", "UpdateListbox22"),
            ($) => {
              const ddl2 = $("#UpdateBudgetTargetTypeId2").data("kendoDropDownList");
              const raw2 = ddl2 ? String(ddl2.text() || "").trim() : "Targeting";
              const label2 = raw2.charAt(0).toUpperCase() + raw2.slice(1).toLowerCase();
              return `Bulk Move ${label2} (2nd Target)`;
            }
          )
        );

        header2.insertAdjacentElement("afterend", btn2);
      }
    },
  };

  // ============================================================
  // INITIALIZATION & POLLING
  // ============================================================

  const AppController = {
    mutationObserver: null,
    beforeUnloadHandler: null,
    _waiting: false,

    /**
     * Wait for edit target buttons to appear. Guarded so overlapping
     * invocations can't stack concurrent poll loops.
     * @returns {Promise<void>}
     */
    async waitForEditTargetsButtons() {
      if (this._waiting) return;
      this._waiting = true;
      try {
        for (let i = 0; i < TIMING.MAX_WAIT_ITERATIONS; i++) {
          const countyBtn = document.querySelector('button[onclick="NewTargetCounty()"]');
          const cityBtn = document.querySelector('button[onclick="NewTargetCity()"]');
          const zipBtn = document.querySelector('button[onclick="NewTargetZip()"]');

          if (countyBtn || cityBtn || zipBtn) {
            ButtonInjector.injectTargetButtons();
            return;
          }

          await sleep(TIMING.WAIT_ITERATION_DELAY_MS);
        }
      } finally {
        this._waiting = false;
      }
    },

    /**
     * Initial injection pass.
     */
    tick() {
      this.waitForEditTargetsButtons().catch((error) => {
        logError("Error waiting for edit target buttons:", error);
      });
      ButtonInjector.injectBulkMoveButton_MelonMax();
      ButtonInjector.injectBulkMoveButtons_AgentsBudgetDetails();
    },

    /**
     * Debounced mutation handler — the observer is the only steady-state
     * driver for re-injection, so make sure all injectors run here.
     */
    handleMutations: debounce(function () {
      ButtonInjector.injectBulkMoveButton_MelonMax();
      ButtonInjector.injectBulkMoveButtons_AgentsBudgetDetails();
      ButtonInjector.injectTargetButtons();
    }, 300),

    /**
     * Initialize the application
     */
    init() {
      log("Initializing", VERSION);

      // Inject global styles
      injectGlobalStyles();

      // Initial tick
      this.tick();

      // Setup mutation observer — this is the sole re-injection trigger
      // after init (setInterval polling was removed to avoid double-firing).
      this.mutationObserver = new MutationObserver(() => {
        this.handleMutations();
      });

      this.mutationObserver.observe(document.documentElement || document.body, {
        childList: true,
        subtree: true,
      });

      // Fill in the shared _debug object so window.PatchTargetingHelper
      // exposes the same references that PatchTargetingHelperDebug sees.
      Object.assign(_debug, {
        FocusManager,
        BulkAddOperations,
        BulkRemoveOperations,
        BulkMoveOperations,
        ButtonInjector,
        PageDetector,
      });

      window.PatchTargetingHelper = {
        version: VERSION,
        _debug,
      };

      log("Loaded successfully");
    },

    /**
     * Cleanup on unload
     */
    cleanup() {
      if (this.mutationObserver) {
        this.mutationObserver.disconnect();
        this.mutationObserver = null;
      }

      if (this.beforeUnloadHandler) {
        window.removeEventListener("beforeunload", this.beforeUnloadHandler);
        this.beforeUnloadHandler = null;
      }

      FocusManager.restoreDashboardFocusManagement();

      log("Cleaned up");
    },
  };

  // Initialize when DOM is ready
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => AppController.init());
  } else {
    AppController.init();
  }

  // Cleanup on page unload — keep a reference so cleanup() can unbind it.
  AppController.beforeUnloadHandler = () => AppController.cleanup();
  window.addEventListener("beforeunload", AppController.beforeUnloadHandler);
})();