// ==UserScript==
// @name         MelonPatch - Remove Ex-Employees from System Groups
// @namespace    http://tampermonkey.net/
// @version      2.3.0
// @description  Highlights ex-employees on System Group & Teams Detail pages, plus inline panels to bulk-remove ex-employees and bulk-add active Melons.
// @match        https://thepatch.melonlocal.com/SystemGroups/Edit*
// @match        https://thepatch.melonlocal.com/Teams/Details*
// @run-at       document-end
// @grant        none
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/MelonPatch%20-%20Remove%20Ex-Employees%20from%20System%20Groups-1.0.user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/MelonPatch%20-%20Remove%20Ex-Employees%20from%20System%20Groups-1.0.user.js
// ==/UserScript==

(function () {
  "use strict";

  const VERSION = "patch-ex-employees-v2.3.0";
  const DEBUG = false;
  const debug = { DEBUGMODE: DEBUG };

  const TIMING = {
    DEFAULT_DELAY_MS: 120,
    MAX_WAIT_ITERATIONS: 200,
    WAIT_ITERATION_DELAY_MS: 150,
    FOCUS_DELAY_MS: 0,
    RETRY_BASE_BACKOFF_MS: 200,
    RETRY_MAX_ATTEMPTS: 2,
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
    SG_REMOVE_MEMBER: "/SystemGroups/Edit?handler=RemoveTargetSystemGroupMember",
    SG_ADD_MEMBER: "/SystemGroups/Edit",
    TEAMS_ADD_MEMBER: "/Teams/AddTeamMember",
    TEAMS_REMOVE_MEMBER: "/Teams/RemoveTeamMember",
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

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  function debounce(func, wait) {
    let t;
    return function (...args) {
      clearTimeout(t);
      t = setTimeout(() => func(...args), wait);
    };
  }

  function normalizeText(s) {
    return String(s || "").replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
  }

  function insertAfter(newNode, ref) {
    const p = ref?.parentNode;
    if (p) p.insertBefore(newNode, ref.nextSibling);
  }

  function log(msg, ...a) {
    if (DEBUG || debug.DEBUGMODE) console.log(`[PatchExEmployees] ${msg}`, ...a);
  }
  function logError(msg, ...a) {
    console.error(`[PatchExEmployees] ${msg}`, ...a);
  }
  // ── Page Detection ───────────────────────────────────────────
  const PageDetector = {
    get hrefLower() { return String(location.href || "").toLowerCase(); },
    get hostLower() { return String(location.hostname || "").toLowerCase(); },
    get isPatch() { return this.hostLower === "thepatch.melonlocal.com"; },
    get isSystemGroupEdit() { return this.isPatch && this.hrefLower.includes("/systemgroups/edit"); },
    get isTeamsDetails() { return this.isPatch && this.hrefLower.includes("/teams/details"); },
    get isSupported() { return this.isSystemGroupEdit || this.isTeamsDetails; },
  };

  // ── Global Styles ────────────────────────────────────────────
  function injectGlobalStyles() {
    if (document.getElementById("patch-ex-employees-styles")) return;
    const s = document.createElement("style");
    s.id = "patch-ex-employees-styles";
    s.textContent = `
      .patch-inline-panel{margin-top:12px;padding:12px 14px;background:${COLORS.alpine};border:1px solid ${COLORS.mojave};border-radius:8px;font-family:system-ui,-apple-system,"Segoe UI",sans-serif;color:${COLORS.coconut};box-shadow:0 1px 2px rgba(100,68,20,.08)}
      .patch-inline-panel[hidden]{display:none!important}
      .patch-inline-panel-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:8px}
      .patch-inline-panel-title{font-size:14px;font-weight:600;color:${COLORS.coconut}}
      .patch-inline-panel-close{background:none;border:none;cursor:pointer;font-size:20px;line-height:1;color:${COLORS.coconut};padding:0 6px}
      .patch-inline-panel-close:hover{color:${COLORS.cranberry}}
      .patch-inline-status{margin-top:6px;font-size:12px;min-height:16px;color:${COLORS.coconut}}
      .patch-inline-status.is-error{color:${COLORS.cranberry}}
      .patch-inline-status.is-warn{color:${COLORS.mustardSeed}}
      .patch-inline-status .badge{display:inline-block;padding:1px 6px;border-radius:10px;background:${COLORS.sand};color:${COLORS.coconut};margin-left:4px;font-weight:600}
      .patch-inline-panel-button-row{display:flex;gap:6px;justify-content:flex-end;margin-top:10px;flex-wrap:wrap}
      .patch-inline-btn{border:none;padding:7px 14px;border-radius:4px;cursor:pointer;font-size:13px;font-weight:500;transition:background-color 120ms ease}
      .patch-inline-btn[disabled]{cursor:progress;opacity:.85}
      .patch-inline-btn--primary{background:${COLORS.cactus};color:#fff}
      .patch-inline-btn--secondary{background:${COLORS.sand};color:${COLORS.coconut}}
      .patch-inline-btn--danger{background:${COLORS.watermelonSugar};color:#fff}
      .patch-inline-btn--busy{background:${COLORS.mustardSeed}!important;color:#fff!important}
      .patch-ex-employee-list{margin-top:4px;padding:8px;background:#fff;border:1px solid ${COLORS.mojave};border-radius:4px;max-height:180px;overflow-y:auto;font-size:12px;color:${COLORS.coconut}}
      .patch-ex-employee-list-empty{color:${COLORS.mustardSeed};font-style:italic}
      .patch-ex-employee-list-item{padding:2px 0;font-family:ui-monospace,SFMono-Regular,Menlo,monospace}
      .patch-ex-employee-chip{background-color:${COLORS.whitneyPink}!important;border-color:${COLORS.watermelonSugar}!important;color:${COLORS.cranberry}!important}
      .patch-ex-employee-row{background:${COLORS.whitneyPink}!important;border-radius:4px;padding:2px 4px;color:${COLORS.cranberry}!important}
      .patch-add-members-search{width:100%;box-sizing:border-box;padding:7px 10px;font-size:13px;border:1px solid ${COLORS.mojave};border-radius:4px;margin-bottom:6px;background:#fff;color:${COLORS.coconut}}
      .patch-add-members-search:focus{outline:none;border-color:${COLORS.cactus};box-shadow:0 0 0 2px rgba(71,183,79,.15)}
      .patch-add-members-list{margin-top:4px;padding:4px;background:#fff;border:1px solid ${COLORS.mojave};border-radius:4px;max-height:280px;overflow-y:auto;font-size:13px;color:${COLORS.coconut}}
      .patch-add-members-list-item{display:flex;align-items:center;gap:8px;padding:4px 6px;border-radius:3px;cursor:pointer;user-select:none}
      .patch-add-members-list-item:hover{background:${COLORS.alpine}}
      .patch-add-members-list-item input[type="checkbox"]{margin:0;cursor:pointer;accent-color:${COLORS.cactus}}
      .patch-add-members-list-item-name{flex:1}
      .patch-add-members-list-item-id{color:${COLORS.mustardSeed};font-family:ui-monospace,SFMono-Regular,Menlo,monospace;font-size:11px}
      .patch-add-members-list-empty{padding:8px;color:${COLORS.mustardSeed};font-style:italic}
    `;
    document.head.appendChild(s);
  }

  // ── API Client ───────────────────────────────────────────────
  const ApiClient = {
    getAntiforgeryToken() {
      return document.querySelector('input[name="__RequestVerificationToken"]')?.value || "";
    },
    getSystemGroupId() {
      const h = document.querySelector('input[name="systemGroupId"]');
      if (h?.value) return String(h.value);
      return new URLSearchParams(window.location.search).get("systemGroupId") || null;
    },
    getTeamId() {
      return new URLSearchParams(window.location.search).get("id") || null;
    },
    async getActiveMelons() {
      const r = await fetch(ENDPOINTS.GET_MELONS, { credentials: "same-origin" });
      if (!r.ok) throw new Error(`HTTP ${r.status} fetching active melons`);
      return r.json();
    },
    async getActiveMelonIds() {
      return new Set((await this.getActiveMelons()).map((m) => m.MelonId));
    },
    async addMember({ systemGroupId, melonId, token }, opts = {}) {
      return this._post(ENDPOINTS.SG_ADD_MEMBER, { SystemGroupId: Number(systemGroupId), MelonId: Number(melonId) }, token, opts);
    },
    async removeMember({ systemGroupId, melonId, token }, opts = {}) {
      return this._post(ENDPOINTS.SG_REMOVE_MEMBER, { SystemGroupId: Number(systemGroupId), MelonId: Number(melonId) }, token, opts);
    },
    async addTeamMember({ teamId, melonId, token }, opts = {}) {
      return this._post(ENDPOINTS.TEAMS_ADD_MEMBER, { TeamId: Number(teamId), MemberId: Number(melonId) }, token, opts);
    },
    async removeTeamMember({ teamId, melonId, token }, opts = {}) {
      return this._post(ENDPOINTS.TEAMS_REMOVE_MEMBER, { TeamId: Number(teamId), MemberId: Number(melonId) }, token, opts);
    },
    async _post(url, bodyObj, token, opts = {}) {
      const maxRetries = opts.maxRetries ?? TIMING.RETRY_MAX_ATTEMPTS;
      const baseBackoff = opts.baseBackoffMs ?? TIMING.RETRY_BASE_BACKOFF_MS;
      const body = JSON.stringify(bodyObj);
      let lastError;
      for (let attempt = 0; attempt <= maxRetries; attempt++) {
        try {
          const res = await fetch(url, {
            method: "POST",
            credentials: "same-origin",
            headers: {
              "Content-Type": "application/json; charset=utf-8",
              "X-Requested-With": "XMLHttpRequest",
              RequestVerificationToken: token,
            },
            body,
          });
          if (!res.ok) {
            if (res.status >= 500 || res.status === 429) throw new Error(`HTTP ${res.status}`);
            throw Object.assign(new Error(`HTTP ${res.status}`), { terminal: true });
          }
          return;
        } catch (err) {
          lastError = err;
          if (err?.terminal || attempt === maxRetries) break;
          await sleep(baseBackoff * Math.pow(2, attempt));
        }
      }
      throw lastError ?? new Error("Unknown error in _post");
    },
  };
  // ── Ex-Employee Operations ───────────────────────────────────
  const ExEmployeeOperations = {
    async scan() {
      let activeMelonIds;
      try {
        activeMelonIds = await ApiClient.getActiveMelonIds();
      } catch (e) {
        logError("Failed to load active melons:", e);
        return { members: [], exEmployees: [], reason: "fetch-failed", error: e };
      }

      if (PageDetector.isSystemGroupEdit) {
        const container = document.getElementById("systemGroupMembersChipList");
        const chips = container
          ? Array.from(container.querySelectorAll(".k-chip[data-chip-id],[data-chip-id]"))
          : [];
        if (!chips.length) return { members: [], exEmployees: [], reason: "no-chips" };
        const exEmployees = [];
        for (const chip of chips) {
          const melonId = Number(chip.getAttribute("data-chip-id"));
          if (!Number.isFinite(melonId)) continue;
          if (activeMelonIds.has(melonId)) {
            chip.classList.remove(PANEL_IDS.chipMarkClass);
            if (chip.title === "Ex-employee (not in active Melons list)") chip.removeAttribute("title");
            continue;
          }
          chip.classList.add(PANEL_IDS.chipMarkClass);
          chip.title = "Ex-employee (not in active Melons list)";
          exEmployees.push({ element: chip, melonId, name: normalizeText(chip.textContent) || `MelonId ${melonId}` });
        }
        log("Scan complete (SystemGroups)", { total: chips.length, exEmployees: exEmployees.length });
        return { members: chips, exEmployees, reason: "ok" };
      }

      if (PageDetector.isTeamsDetails) {
        const btns = Array.from(document.querySelectorAll('button[id^="removeMember_"]'));
        if (!btns.length) return { members: [], exEmployees: [], reason: "no-chips" };
        const exEmployees = [];
        for (const btn of btns) {
          const melonId = parseInt(btn.id.replace("removeMember_", ""), 10);
          if (!Number.isFinite(melonId)) continue;
          const rowEl = btn.parentElement;
          if (activeMelonIds.has(melonId)) {
            rowEl?.classList.remove("patch-ex-employee-row");
            if (rowEl?.title === "Ex-employee (not in active Melons list)") rowEl.removeAttribute("title");
            continue;
          }
          rowEl?.classList.add("patch-ex-employee-row");
          if (rowEl) rowEl.title = "Ex-employee (not in active Melons list)";
          const name = normalizeText(rowEl?.childNodes[0]?.textContent || "") || `MelonId ${melonId}`;
          exEmployees.push({ element: btn, rowEl, melonId, name });
        }
        log("Scan complete (Teams)", { total: btns.length, exEmployees: exEmployees.length });
        return { members: btns, exEmployees, reason: "ok" };
      }

      return { members: [], exEmployees: [], reason: "no-chips" };
    },

    async run(exEmployees, delayMs = TIMING.DEFAULT_DELAY_MS, opts = {}) {
      if (!exEmployees.length) return { ok: false, reason: "empty", message: "Nothing to remove." };
      const token = ApiClient.getAntiforgeryToken();
      let groupId, removeOneFn;
      if (PageDetector.isSystemGroupEdit) {
        groupId = ApiClient.getSystemGroupId();
        if (!token || !groupId) return { ok: false, reason: "no-api", message: "Cannot remove: missing token or systemGroupId." };
        removeOneFn = (e) => ApiClient.removeMember({ systemGroupId: groupId, melonId: e.melonId, token });
      } else if (PageDetector.isTeamsDetails) {
        groupId = ApiClient.getTeamId();
        if (!token || !groupId) return { ok: false, reason: "no-api", message: "Cannot remove: missing token or teamId." };
        removeOneFn = (e) => ApiClient.removeTeamMember({ teamId: groupId, melonId: e.melonId, token });
      } else {
        return { ok: false, reason: "no-api", message: "Unsupported page." };
      }
      let removed = 0;
      const failed = [];
      for (let i = 0; i < exEmployees.length; i++) {
        const entry = exEmployees[i];
        try {
          await removeOneFn(entry);
          try { (entry.rowEl || entry.element).remove(); } catch (_) {}
          removed++;
          await sleep(delayMs);
        } catch (e) {
          logError(`Remove failed MelonId ${entry.melonId}:`, e?.message);
          failed.push(entry);
        }
        opts.onProgress?.({ completed: i + 1, total: exEmployees.length, phase: "Removing" });
      }
      return {
        ok: true, completed: removed, total: exEmployees.length, failed, reload: removed > 0,
        message: `Removed ${removed} of ${exEmployees.length} ex-employee(s)` + (failed.length ? `  •  Failed: ${failed.length}` : ""),
      };
    },
  };

  // ── Add-Member Operations ────────────────────────────────────
  const AddMemberOperations = {
    getCurrentMemberIds() {
      const ids = new Set();
      if (PageDetector.isSystemGroupEdit) {
        const c = document.getElementById("systemGroupMembersChipList");
        if (c) for (const chip of c.querySelectorAll(".k-chip[data-chip-id],[data-chip-id]")) {
          const id = Number(chip.getAttribute("data-chip-id"));
          if (Number.isFinite(id)) ids.add(id);
        }
      } else if (PageDetector.isTeamsDetails) {
        for (const btn of document.querySelectorAll('button[id^="removeMember_"]')) {
          const id = parseInt(btn.id.replace("removeMember_", ""), 10);
          if (Number.isFinite(id)) ids.add(id);
        }
      }
      return ids;
    },
    async getCandidates() {
      const all = await ApiClient.getActiveMelons();
      const current = this.getCurrentMemberIds();
      const candidates = all
        .filter((m) => Number.isFinite(Number(m?.MelonId)) && !current.has(Number(m.MelonId)));
      candidates.sort((a, b) => melonDisplayName(a).localeCompare(melonDisplayName(b)));
      log("Add-member candidates", { total: all.length, alreadyMembers: current.size, candidates: candidates.length });
      return candidates;
    },
    async run(melons, delayMs = TIMING.DEFAULT_DELAY_MS, opts = {}) {
      if (!melons.length) return { ok: false, reason: "empty", message: "Nothing to add." };
      const token = ApiClient.getAntiforgeryToken();
      let groupId, addOneFn;
      if (PageDetector.isSystemGroupEdit) {
        groupId = ApiClient.getSystemGroupId();
        if (!token || !groupId) return { ok: false, reason: "no-api", message: "Cannot add: missing token or systemGroupId." };
        addOneFn = (m) => ApiClient.addMember({ systemGroupId: groupId, melonId: m.MelonId, token });
      } else if (PageDetector.isTeamsDetails) {
        groupId = ApiClient.getTeamId();
        if (!token || !groupId) return { ok: false, reason: "no-api", message: "Cannot add: missing token or teamId." };
        addOneFn = (m) => ApiClient.addTeamMember({ teamId: groupId, melonId: m.MelonId, token });
      } else {
        return { ok: false, reason: "no-api", message: "Unsupported page." };
      }
      let added = 0;
      const failed = [];
      for (let i = 0; i < melons.length; i++) {
        const m = melons[i];
        try { await addOneFn(m); added++; await sleep(delayMs); }
        catch (e) { logError(`Add failed MelonId ${m.MelonId}:`, e?.message); failed.push(m); }
        opts.onProgress?.({ completed: i + 1, total: melons.length, phase: "Adding" });
      }
      return {
        ok: true, completed: added, total: melons.length, failed, reload: added > 0,
        message: `Added ${added} of ${melons.length} member(s)` + (failed.length ? `  •  Failed: ${failed.length}` : ""),
      };
    },
  };

  // ── Inline Dashboard Helpers ─────────────────────────────────
  const InlineDashboard = {
    makeButton(text, variant = "primary") {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = text;
      btn.className = `patch-inline-btn patch-inline-btn--${variant}`;
      btn.dataset.originalText = text;
      return btn;
    },
    setBusy(btn, busy, label) {
      btn.disabled = busy;
      btn.classList.toggle("patch-inline-btn--busy", busy);
      if (busy && label != null) btn.textContent = label;
      else if (!busy) btn.textContent = btn.dataset.originalText || btn.textContent;
    },
  };
  // ── Ex-Employee Panel ────────────────────────────────────────
  function buildExEmployeePanel() {
    const panel = document.createElement("div");
    panel.id = PANEL_IDS.panelId;
    panel.className = PANEL_IDS.inlinePanelClass;
    panel.hidden = true;
    panel.dataset.bulkPanel = "ex-employees";

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
    header.append(titleEl, closeBtn);

    const list = document.createElement("div");
    list.className = "patch-ex-employee-list";
    const status = document.createElement("div");
    status.className = `${PANEL_IDS.inlinePanelStatusClass} is-info`;
    const btnRow = document.createElement("div");
    btnRow.className = `${PANEL_IDS.inlinePanelClass}-button-row`;
    const cancelBtn = InlineDashboard.makeButton("Cancel", "secondary");
    const rescanBtn = InlineDashboard.makeButton("Rescan", "secondary");
    const removeBtn = InlineDashboard.makeButton("Remove Ex-Employees", "danger");
    btnRow.append(cancelBtn, rescanBtn, removeBtn);
    panel.append(header, list, status, btnRow);

    let lastScan = { exEmployees: [], reason: "pending" };

    const countMembers = () => {
      if (PageDetector.isSystemGroupEdit)
        return document.querySelectorAll("#systemGroupMembersChipList .k-chip[data-chip-id],#systemGroupMembersChipList [data-chip-id]").length;
      if (PageDetector.isTeamsDetails)
        return document.querySelectorAll('button[id^="removeMember_"]').length;
      return 0;
    };

    const renderList = () => {
      list.innerHTML = "";
      if (!lastScan.exEmployees.length) {
        const e = document.createElement("div");
        e.className = "patch-ex-employee-list-empty";
        e.textContent = lastScan.reason === "pending"
          ? "Click Rescan to detect ex-employees."
          : "No ex-employees detected.";
        list.appendChild(e);
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
      const onPage = countMembers();
      const badge = `<span class="badge">${onPage} on page</span>`;
      if (lastScan.reason === "fetch-failed") {
        status.classList.add("is-error");
        status.innerHTML = `Could not fetch active melons. ${badge}`;
        return;
      }
      if (lastScan.reason === "no-chips") {
        status.classList.add("is-warn");
        status.textContent = "No members visible yet.";
        return;
      }
      if (lastScan.reason === "pending") {
        status.classList.add("is-info");
        status.innerHTML = `Ready. ${badge}`;
        return;
      }
      status.classList.add("is-info");
      status.innerHTML = lastScan.exEmployees.length
        ? `Found <strong>${lastScan.exEmployees.length}</strong> ex-employee(s). ${badge}`
        : `No ex-employees found. ${badge}`;
    };

    const updateRemoveBtnState = () => { removeBtn.disabled = !lastScan.exEmployees.length; };

    const rescan = async () => {
      InlineDashboard.setBusy(rescanBtn, true, "Scanning...");
      try { lastScan = await ExEmployeeOperations.scan(); }
      catch (e) { logError("Rescan failed:", e); lastScan = { exEmployees: [], reason: "fetch-failed", error: e }; }
      InlineDashboard.setBusy(rescanBtn, false);
      renderList(); renderStatus(); updateRemoveBtnState();
    };

    const close = () => { panel.hidden = true; };
    cancelBtn.addEventListener("click", close);
    closeBtn.addEventListener("click", close);
    rescanBtn.addEventListener("click", rescan);

    removeBtn.addEventListener("click", async () => {
      if (removeBtn.disabled) return;
      const count = lastScan.exEmployees.length;
      if (!count || !confirm(`Remove ${count} ex-employee(s) from this group? This cannot be undone.`)) return;
      cancelBtn.disabled = rescanBtn.disabled = true;
      InlineDashboard.setBusy(removeBtn, true, `(0/${count}) Removing...`);
      const result = await ExEmployeeOperations.run(lastScan.exEmployees, TIMING.DEFAULT_DELAY_MS, {
        onProgress: ({ completed, total }) => { removeBtn.textContent = `(${completed}/${total}) Removing...`; },
      });
      InlineDashboard.setBusy(removeBtn, false);
      cancelBtn.disabled = rescanBtn.disabled = false;
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

    renderList(); renderStatus(); updateRemoveBtnState();

    return {
      panel, rescan,
      show: async () => { panel.hidden = false; await rescan(); },
      hide: close,
      toggle: async () => { if (panel.hidden) { panel.hidden = false; await rescan(); } else close(); },
      isVisible: () => !panel.hidden,
    };
  }
  // ── Add Members Panel ────────────────────────────────────────
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
    header.append(titleEl, closeBtn);

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
    const cancelBtn        = InlineDashboard.makeButton("Cancel", "secondary");
    const refreshBtn       = InlineDashboard.makeButton("Refresh", "secondary");
    const clearBtn         = InlineDashboard.makeButton("Clear", "secondary");
    const selectVisibleBtn = InlineDashboard.makeButton("Select visible", "secondary");
    const addBtn           = InlineDashboard.makeButton("Add Selected", "primary");
    btnRow.append(cancelBtn, refreshBtn, clearBtn, selectVisibleBtn, addBtn);
    panel.append(header, search, list, status, btnRow);

    let allCandidates = [], filtered = [], selected = new Map(), loadState = "pending";

    const renderStatus = () => {
      status.classList.remove("is-error", "is-warn", "is-info");
      if (loadState === "error") {
        status.classList.add("is-error");
        status.textContent = "Could not load active Melons.";
        return;
      }
      status.classList.add("is-info");
      status.innerHTML =
        `<strong>${selected.size}</strong> selected  ` +
        `<span class="badge">${filtered.length} shown</span>  ` +
        `<span class="badge">${allCandidates.length} addable</span>`;
    };

    const updateAddBtnState = () => { addBtn.disabled = selected.size === 0; };

    const renderList = () => {
      list.innerHTML = "";
      const makeEmpty = (text) => {
        const d = document.createElement("div");
        d.className = "patch-add-members-list-empty";
        d.textContent = text;
        list.appendChild(d);
      };
      if (loadState === "pending") { makeEmpty("Loading active Melons..."); return; }
      if (loadState === "error")   { makeEmpty("Failed to load. Try Refresh."); return; }
      if (!filtered.length) {
        makeEmpty(allCandidates.length
          ? "No matches for that search."
          : "No melons available to add (everyone is already a member).");
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
          if (cb.checked) selected.set(m.MelonId, m); else selected.delete(m.MelonId);
          renderStatus(); updateAddBtnState();
        });
        const nameEl = document.createElement("span");
        nameEl.className = "patch-add-members-list-item-name";
        nameEl.textContent = melonDisplayName(m);
        const idEl = document.createElement("span");
        idEl.className = "patch-add-members-list-item-id";
        idEl.textContent = `#${m.MelonId}`;
        row.append(cb, nameEl, idEl);
        frag.appendChild(row);
      }
      list.appendChild(frag);
    };

    const applyFilter = () => {
      const q = (search.value || "").toLowerCase().trim();
      filtered = q ? allCandidates.filter((m) => melonMatchesQuery(m, q)) : allCandidates.slice();
      renderList(); renderStatus();
    };

    const refresh = async () => {
      loadState = "pending";
      InlineDashboard.setBusy(refreshBtn, true, "Loading...");
      renderList(); renderStatus();
      try {
        allCandidates = await AddMemberOperations.getCandidates();
        const stillValid = new Set(allCandidates.map((m) => m.MelonId));
        for (const id of [...selected.keys()]) { if (!stillValid.has(id)) selected.delete(id); }
        loadState = "ok";
      } catch (e) {
        logError("Failed to load candidates:", e);
        allCandidates = [];
        loadState = "error";
      }
      InlineDashboard.setBusy(refreshBtn, false);
      applyFilter(); updateAddBtnState();
    };

    const close = () => { panel.hidden = true; };
    cancelBtn.addEventListener("click", close);
    closeBtn.addEventListener("click", close);
    refreshBtn.addEventListener("click", refresh);
    clearBtn.addEventListener("click", () => { selected.clear(); renderList(); renderStatus(); updateAddBtnState(); });
    selectVisibleBtn.addEventListener("click", () => {
      for (const m of filtered) selected.set(m.MelonId, m);
      renderList(); renderStatus(); updateAddBtnState();
    });
    search.addEventListener("input", debounce(applyFilter, 100));

    addBtn.addEventListener("click", async () => {
      if (addBtn.disabled) return;
      const picks = [...selected.values()];
      if (!picks.length || !confirm(`Add ${picks.length} member(s) to this group?`)) return;
      [cancelBtn, refreshBtn, clearBtn, selectVisibleBtn].forEach((b) => (b.disabled = true));
      search.disabled = true;
      InlineDashboard.setBusy(addBtn, true, `(0/${picks.length}) Adding...`);
      const result = await AddMemberOperations.run(picks, TIMING.DEFAULT_DELAY_MS, {
        onProgress: ({ completed, total }) => { addBtn.textContent = `(${completed}/${total}) Adding...`; },
      });
      InlineDashboard.setBusy(addBtn, false);
      [cancelBtn, refreshBtn, clearBtn, selectVisibleBtn].forEach((b) => (b.disabled = false));
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

    renderList(); renderStatus(); updateAddBtnState();

    return {
      panel, refresh,
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
  // ── Button Injector ──────────────────────────────────────────
  const ButtonInjector = {
    _dashboard: null,
    _addDashboard: null,

    inject() {
      if (!PageDetector.isSupported) return;

      const haveExBtn  = !!document.getElementById(PANEL_IDS.triggerBtnId);
      const haveAddBtn = !!document.getElementById(PANEL_IDS.addTriggerBtnId);
      if (haveExBtn && haveAddBtn) return;

      // Find the anchor element to inject next to
      let anchor = null;
      if (PageDetector.isSystemGroupEdit) {
        anchor = document.getElementById("systemGroupMembersChipList");
      } else if (PageDetector.isTeamsDetails) {
        // Use the first removeMember button's parent's parent (the .dashboard-panel)
        const firstBtn = document.querySelector('button[id^="removeMember_"]');
        anchor = firstBtn?.closest(".dashboard-panel") || firstBtn?.parentElement || null;
      }

      if (!anchor) return;

      const exDashboard  = haveExBtn  ? this._dashboard    : buildExEmployeePanel();
      const addDashboard = haveAddBtn ? this._addDashboard  : buildAddMembersPanel();
      this._dashboard    = exDashboard;
      this._addDashboard = addDashboard;

      const anchorParent = anchor.parentElement;

      const buildTriggerBtn = (id, text, onClick) => {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.id = id;
        btn.className = "btn btn--small melon-green";
        btn.style.marginBottom = "10px";
        btn.style.marginRight  = "10px";
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
        if (anchorParent) anchorParent.insertBefore(addTrigger, anchor);
        else anchor.before(addTrigger);
        insertAfter(addDashboard.panel, addTrigger);
      }

      // Remove Ex-Employees trigger
      if (!haveExBtn) {
        const exTrigger = buildTriggerBtn(PANEL_IDS.triggerBtnId, "Remove Ex-Employees", () => {
          addDashboard.hide();
          exDashboard.toggle();
        });
        const addTrigger = document.getElementById(PANEL_IDS.addTriggerBtnId);
        if (addTrigger) insertAfter(exTrigger, addTrigger);
        else if (anchorParent) anchorParent.insertBefore(exTrigger, anchor);
        else anchor.before(exTrigger);
        insertAfter(exDashboard.panel, exTrigger);

        // Auto-scan on injection so ex-employee rows get marked immediately
        exDashboard.rescan().catch((e) => logError("Initial rescan failed:", e));
      }
    },

    rescan() {
      if (this._dashboard) this._dashboard.rescan().catch((e) => logError("Manual rescan failed:", e));
    },

    injectAllInlineUI() { this.inject(); },
  };

  // ── App Controller ───────────────────────────────────────────
  const _debug = {};

  const AppController = {
    mutationObserver: null,
    beforeUnloadHandler: null,
    reinjectIntervalId: null,
    _waiting: false,

    async waitForAnchor() {
      if (this._waiting) return;
      this._waiting = true;
      try {
        for (let i = 0; i < TIMING.MAX_WAIT_ITERATIONS; i++) {
          const found = PageDetector.isSystemGroupEdit
            ? document.getElementById("systemGroupMembersChipList")
            : document.querySelector('button[id^="removeMember_"]');
          if (found) { ButtonInjector.inject(); return; }
          await sleep(TIMING.WAIT_ITERATION_DELAY_MS);
        }
      } finally {
        this._waiting = false;
      }
    },

    tick() { this.waitForAnchor().catch((e) => logError("waitForAnchor:", e)); },

    handleMutations: debounce(function () { ButtonInjector.injectAllInlineUI(); }, 300),

    rescan() { ButtonInjector.rescan(); },

    init() {
      if (!PageDetector.isSupported) return;
      log("Initializing", VERSION);
      injectGlobalStyles();
      this.tick();

      this.mutationObserver = new MutationObserver(() => this.handleMutations());
      this.mutationObserver.observe(document.documentElement || document.body, {
        childList: true,
        subtree: true,
      });

      this.reinjectIntervalId = setInterval(() => {
        if (PageDetector.isSupported) ButtonInjector.injectAllInlineUI();
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
        enableDebug:    () => { debug.DEBUGMODE = true; },
        disableDebug:   () => { debug.DEBUGMODE = false; },
        injectButtons:  () => { ButtonInjector.injectAllInlineUI(); console.log("[PatchExEmployees] Manual UI injection triggered"); },
        rescan:         () => AppController.rescan(),
      };

      log("Loaded successfully");
    },

    cleanup() {
      if (this.mutationObserver)    { this.mutationObserver.disconnect(); this.mutationObserver = null; }
      if (this.reinjectIntervalId)  { clearInterval(this.reinjectIntervalId); this.reinjectIntervalId = null; }
      if (this.beforeUnloadHandler) { window.removeEventListener("beforeunload", this.beforeUnloadHandler); this.beforeUnloadHandler = null; }
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