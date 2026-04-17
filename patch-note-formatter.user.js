// ==UserScript==
// @name         Patch Note Formatter (Preserve Lists) + Full Term Normalizer + Preview Diff
// @namespace    http://tampermonkey.net/
// @version      1.8.0
// @description  Formats Patch notes with bold headings and preserves numbered/bulleted lists. Normalizes alternative transcriptions to canonical terms. Adds a before/after preview diff with confirm/cancel and optional auto-save. Draggable UI + minimize. v1.8.0: extracted modal wiring helper, O(1) term lookup via Map, fixed nested-list bug, waitFor polling instead of fixed sleeps, user-editable term overrides in localStorage, focus trap in modal, injected stylesheet with CSS vars, history-patched route detection.
// @match        https://thepatch.melonlocal.com/*
// @run-at       document-end
// @grant        none
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/patch-note-formatter.user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/patch-note-formatter.user.js
// ==/UserScript==

(function () {
  "use strict";

  // =============================
  // CONSTANTS & CONFIGURATION
  // =============================
  const CONFIG = {
    DEFAULT_AUTO_SAVE: false,
    DEFAULT_START_MINIMIZED: true,
    HOTKEY: { altKey: true, shiftKey: true, code: "KeyF" },
    MAX_PREVIEW_CHARS: 25000,
    MAX_DIFF_LINES: 600,
    MAX_CHANGE_ENTRIES: 250,
    UI_POS_KEY: "patchNoteFormatter_uiPos_v1",
    UI_MIN_KEY: "patchNoteFormatter_uiMin_v1",
    TERMS_KEY: "patchNoteFormatter_terms_v1",
    WAIT_DEFAULT_MS: 2500,
    WAIT_POLL_MS: 60,
    EDITOR_SLEEP_MS: 250,
  };

  // =============================
  // DEFAULT TERM NORMALIZATION (shipped defaults; user overrides merged in via TermStore)
  // =============================
  const DEFAULT_TERM_NORMALIZATION = {
    ABS: ["abs", "A B S", "a b s", "A.B.S.", "a.b.s", "ABS.", "A B S.", "A-B-S"],
    ACC: ["A C C", "a c c", "A.C.C.", "ACC.", "A C C."],
    AEO: ["A E O", "a e o", "answer engine optimization", "Answer Engine Optimization", "answer-engine optimization", "answer engine optimisation"],
    "Agent Tagged Media": ["Agent Tag Media", "agent tag media", "AgentTagMedia", "agenttagmedia", "Agent tagmedia", "agent tagmedia", "tag media"],
    "analytics studio": [
      ":Analook studio", "analook studio", "analic studio", "analic Studio", "Analic studio", "Analytic studio", "analytic studio",
      "Analyze Studio", "analyze studio", "Analyx Studio", "analyx studio", "Analyte studio", "analyte studio", "Analyst Studio",
      "analyst studio", "Analyst studio", "Analys(t) Studio", "Analyze Studio Dash", "analyze studio dash", "Analytics Studio Dash",
      "analytics studio dash", "Animal Lake Studio", "animal lake studio", "Analyze Studio dashboard", "analyze studio dashboard",
      "Analytics Studio dashboard", "analytics studio dashboard", "Analytics Studio", "analytics Studio", "Analytics studio",
      "Analect studio", "analect studio", "Analog studio", "analog studio", "Antelope Studio", "antelope studio", "Inlet studio",
      "inlet studio", "Inlick studio", "inlick studio", "Intellect studio", "intellect studio", "Alex studio", "alex studio",
      "Lake studio", "lake studio", "Analyz(e) Studio"
    ],
    ARS: ["ars", "A R S", "a r s", "AR's", "R S", "our S", "ours", "Rs"],
    ATM: [],
    "auction insights": ["Auction insights", "auction-insights", "auction insight", "auctioninsights", "Auction Insights", "Auction Insight", "auction insights report", "auction insight report"],
    "Butler/Till": ["Butler Till", "Butler and Till", "Butler & Till", "Butler n Till", "Butler till marketing", "Butler till agency", "Till Butler", "till butler"],
    "Chairman's Circle": ["chairman's circle", "Chairmans Circle", "chairmans circle", "Chairman Circle", "chairman circle", "Chairmens Circle", "chairmens circle"],
    ChatGPT: ["chat gpt", "Chat gpt", "chat-gpt", "chatgpt", "ChatGPTs", "Chat PBT", "Chat TPT"],
    Christel: ["Crystal"],
    "closed won": ["close won", "closed one", "closed-won", "close one", "close-on", "close won status", "closed won status"],
    CSM: ["csm", "C S M", "c s m", "C.S.M."],
    DAC: [],
    Darwyn: ["Darwin"],
    "Direct Clicks": ["direct clicks", "DirectClicks", "directclicks", "Direct Click", "direct click", "DC", "Dee Cee"],
    ECRM: ["ecr m", "e c r m", "eC R M", "E C Rm", "Ecr m", "Ec r m", "E.CRM", "e.CRM", "E.C.R.M", "e.C.R.M", "CRM", "e-crm", "e CRM.", "ECRM.", "E C R M.", "e c r m.", "ecrm", "eCRM"],
    "Electronic Library": ["Electronic Librarian", "electronic librarian", "electronic library", "Electronic library"],
    EverQuote: ["everquote", "Ever Quote", "ever quote", "EverQuote.", "Ever Quotes", "ever quotes", "Everquote", "EverQuot", "Everett quote", "Everett Quote"],
    FCC: ["fcc", "F C C", "f c c", "F.C.C.", "f.c.c", "Federal Communications Commission", "federal communications commission"],
    FIC: ["Field Implementation Coach"],
    FTC: ["ftc", "F T C", "f t c", "F.T.C.", "f.t.c", "Federal Trade Commission", "federal trade commission"],
    GBP: [],
    "Google Business Profile": [],
    ILP: [],
    "impression share": ["Impression share", "impression-share", "impression shares", "Impression Share", "Impression Shares", "impression-share percentage"],
    "internet lead provider": [],
    IPS: ["I P S", "I.P.S", "ips", "IPS case"],
    Korry: ["Korrey"],
    lapscan: ["LAPScan", "lap scan", "LAP scan", "lapse scan", "laps scan", "laps camera", "laps cam", "labs scan", "lab scan", "labscan", "lap-scan"],
    M1: ["Em 1", "Em I", "Em one", "Em won", "em 1", "em i", "em one", "em won", "M 1", "M I", "M one", "M won", "M-1", "MI", "m 1", "m i", "m one", "m won", "m-1"],
    M2: ["Em 2", "Em II", "Em I I", "Em two", "Em too", "em 2", "em ii", "em i i", "em two", "em too", "M 2", "M II", "M I I", "M two", "M too", "M-2", "MII", "MI I", "Mii", "m 2", "m ii", "m i i", "m two", "m too", "m-2", "m2"],
    Melon: ["mellon", "Mellon", "mellow", "Mellow", "mello", "Mello", "melow", "Melow", "Melan", "Mellan", "Melonn", "Melon.", "MelonInc", "Melon Inc", "MelonIncorporated", "Melon Incorporated", "melo"],
    "Melon Local": [
      "Melonlocal", "melonlocal", "Mellon Local", "MellonLocal", "Mellow Local", "mellow local", "Mello Local", "mello local",
      "Melow Local", "melow local", "Melon Loca", "MelonLocl", "Melon Locl", "MelonLocal.", "Melon Locals", "Melon Local marketing",
      "Melon local marketing", "Melon Local agency", "Melon local agency", "Mellon local", "local melon"
    ],
    "Melon Max": [
      "all of macs", "Lmax", "lmax", "Mail on Max", "Mealon Max", "Meh-lon Max", "Mehlon Max", "Mello Max", "MelloMax", "Mellow Max",
      "Mel and max", "Melmax", "melmax", "Mellon Max", "Mellon max", "Mellonmax", "Melomax", "Melon Local Max", "Melon Local max",
      "Melon local max", "Melon-Mac", "Melon Mack", "Melon Macs", "Melon Maps", "Melon Maxx", "Melon match", "Melon macks",
      "Melon marks", "Melon mask", "Melon math", "Melon max", "Melon maxed", "Melon maxes", "Melon packs", "MelonMax",
      "MelonLocal Max", "MelonLocal max", "Melonmax", "Millon Max", "Mel and Max", "Melamax"
    ],
    "Melon Max Calls": [
      "Melon Max Call", "Melon max call", "MelonMax call", "MelloMax Calls", "Mello Max Calls", "Mellow Max Calls", "Mellon Max Calls",
      "Mellonmax Calls", "Melonmax Calls", "MelonMax Calls", "Melon Max Live Leads", "Melon max live leads", "MelonMax live leads",
      "MelloMax Live Leads", "Mello Max Live Leads", "Mellow Max Live Leads", "Mellon Max Live Leads", "Mellonmax live leads",
      "Melonmax live leads", "Melon Max Live Transfers", "Melon max live transfers", "MelonMax live transfers", "MelloMax Live Transfers",
      "Mello Max Live Transfers", "Mellow Max Live Transfers", "Mellon Max Live Transfers", "Mellonmax live transfers",
      "Melonmax live transfers", "Mellon max calls"
    ],
    "Melon Max Clicks": [
      "MelloMax Clicks", "Mello Max Clicks", "Mellow Max Clicks", "Mellon Max Clicks", "Mellonmax Clicks", "Melonmax Clicks",
      "MelonMax Clicks", "Melon Max clicks", "Melon max clicks", "Melon Max click", "Melon max click", "MelonMax clicks",
      "Mellon max clicks", "Melon Local max clicks", "Melon Local Max clicks", "Melon local max clicks", "MelonLocal max clicks",
      "MelonLocal Max clicks"
    ],
    "Melon Max Leads": [
      "MelloMax Leads", "Mello Max Leads", "Mellow Max Leads", "Mellon Max Leads", "Mellonmax Leads", "Melonmax Leads", "MelonMax Leads",
      "Melon Max leads", "Melon max leads", "Melon lead", "Melon leads", "Internet leads", "Melon Max Internet Leads",
      "Melon max internet leads", "MelonMax internet leads", "Mellon Max Internet Leads", "Mellon max internet leads",
      "MelloMax Internet Leads", "Mello Max Internet Leads", "Mellow Max Internet Leads", "Melon Max internet lead",
      "Melon max internet lead", "MelonMax internet lead", "Melon Max internet", "Melon max internet", "MelonMax internet",
      "Mellon max leads"
    ],
    Mirus: [
      "Mirus Research", "MirusResearch", "Mirus.", "Myros", "Myris", "Miris", "Merus", "Mearus", "Myrrus", "Mirrus", "Meeres",
      "mirus", "MyRisk", "Myrisk", "my risk", "Miras", "miras", "Iris", "iris", "My risk"
    ],
    MOA: [
      "M O A", "MoA", "moa", "MOA office", "M O A office", "Memorandum of Agreement", "memorandum of agreement",
      "Memorandum Of Agreement", "Memorandum of agreement", "memorandum Of agreement", "Memorandum Agreement",
      "memorandum agreement", "Memorandum Of Agreement office", "Memorandum of Agreement office"
    ],
    Modernization: ["modernization", "Modernisation", "modernisation", "modernization program", "modernisation program"],
    MVR: ["MVRs", "mvr", "mvrs", "M V R", "m v r"],
    MySFDomain: [
      "My SF Domain", "MySF Domain", "My S F Domain", "My S.F. Domain", "MySFDomain.", "My SFDomain", "MySF Domains", "My SF Domains",
      "MySFdomain", "MySFdomian", "My SF domian", "MyEssFDomain", "mysfdomain", "mysf domain", "my sf domain", "SFDomain",
      "sf domain", "SF domain", "sfdomain", "sfdomain's", "SF domain's", "myself domain", "my self domain", "Myself domain",
      "My self domain", "my set domain", "myset domain"
    ],
    NerdWallet: ["nerdwallet", "Nerd Wallet", "nerd wallet", "NerdWallet.", "NerdWallet.com", "nerdwallet dot com", "Nerdwallet", "nerdwallet.com", "nerd wallet dot com"],
    NIL: ["nil", "N I L", "n i l", "N.I.L.", "n.i.l", "name image likeness", "Name Image Likeness"],
    "PC Mod": ["PC mod", "P C Mod", "P C mod", "PCMOD", "pcmod"],
    PMAX: [
      "Performance Max", "performance max", "PerformanceMax", "performancemax", "Perf Max", "perf max", "P Max", "p max", "P-Max",
      "p-max", "Google Performance Max", "google performance max", "Google Performance max", "Keymax", "keymax", "Key Max", "key max",
      "Kaymax", "kaymax", "K max", "k max", "Keymax performance max", "pmacs"
    ],
    PNC: [],
    "property and casualty": [],
    QLP: ["Q L P", "q l p", "Q.L.P.", "cue el pee", "queue el pee", "qlp"],
    QuoteWizard: ["Quote Wizard", "quote wizard", "Quote wizzard", "quote wizzard", "QuoteWizard.", "Quote Wizards", "quote wizards", "QuoteWizard's", "Quote Wizards."],
    ROI: ["Roi", "r o i", "R O I", "R.O.I.", "r.o.i.", "R O I.", "return on investment", "return on investments"],
    "round robin": ["Round robin", "round-robin", "roundrobin", "Round Robin"],
    Scorecard: ["scorecard", "score card", "Score card", "score-card", "Scorecards", "scorecards"],
    "SF layout": ["SF Layout", "State Farm layout", "state farm layout"],
    "SF.com": [],
    Sidekick: ["sidekick", "Side kick", "side kick", "Side-kick", "side-kick", "Side Kick admin assistant", "side kick admin assistant", "Sidekick administrative assistant", "sidekick administrative assistant"],
    SmartFinancial: [
      "Smart Financial", "smart financial", "SmartFinancials", "Smart financials", "Smart finanical", "smart finanical",
      "Smart finantial", "smart finantial", "Smart financal", "smart financal", "Smart financel", "smart financel",
      "SmartFinancial.", "Smart Finacial", "smart finacial"
    ],
    "State Farm": ["StateFarm"],
    "StateFarm.com": ["State Farm .com", "State Farm.com", "State Farm com", "State Farm dot com", "StateFarm com", "State Farm and icom", "State Farm at Conley"],
    "StateFarm.com leads": ["State Farm .com leads", "State Farm dot com leads", "State Farm com leads", "StateFarm com leads", "State Farm website leads", "State Farm internet leads", "StateFarm leads"],
    TikTok: ["Tik Tok", "tik tok", "tiktok"],
    USAA: ["usaa", "U S A A", "u s a a", "U.S.A.A.", "u.s.a.a"],
    YPC: ["y p c", "Y P C", "ypc"],
  };

  // =============================
  // UTILITIES
  // =============================
  const Utils = {
    sleep(ms) {
      return new Promise((resolve) => setTimeout(resolve, ms));
    },

    escapeRegExp(str) {
      return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    },

    escapeHtml(str) {
      return String(str)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
    },

    decodeHtmlEntities(str) {
      if (!str || typeof str !== "string") return str || "";
      try {
        const txt = new DOMParser().parseFromString(str, "text/html");
        return txt.documentElement.textContent || "";
      } catch (e) {
        console.error("[PatchNoteFormatter] HTML decode error:", e);
        return str;
      }
    },

    encodeHtmlEntities(str) {
      if (!str || typeof str !== "string") return str || "";
      const ta = document.createElement("textarea");
      ta.textContent = str;
      return ta.innerHTML;
    },

    looksLikeEscapedHtml(str) {
      if (!str || typeof str !== "string") return false;
      return /&lt;\s*\/?\s*(p|ul|ol|li|strong|br|div|span)\b/i.test(str);
    },

    // Poll until predicate returns a truthy value or timeout. Returns the truthy
    // value (so the same call can act as both wait + find) or null on timeout.
    async waitFor(predicate, timeoutMs = CONFIG.WAIT_DEFAULT_MS, pollMs = CONFIG.WAIT_POLL_MS) {
      const deadline = Date.now() + timeoutMs;
      while (Date.now() < deadline) {
        try {
          const result = predicate();
          if (result) return result;
        } catch (_) {
          // keep polling
        }
        await this.sleep(pollMs);
      }
      return null;
    },
  };

  // =============================
  // STORAGE HELPERS
  // =============================
  const Storage = {
    get(key, defaultValue = null) {
      try {
        const raw = localStorage.getItem(key);
        return raw ? JSON.parse(raw) : defaultValue;
      } catch (e) {
        console.error(`[PatchNoteFormatter] Storage read error for ${key}:`, e);
        return defaultValue;
      }
    },

    set(key, value) {
      try {
        localStorage.setItem(key, JSON.stringify(value));
        return true;
      } catch (e) {
        console.error(`[PatchNoteFormatter] Storage write error for ${key}:`, e);
        return false;
      }
    },
  };

  // =============================
  // ROUTING HELPERS
  // =============================
  const Router = {
    isCallsHash() {
      return (location.hash || "").toLowerCase() === "#calls";
    },

    isOnCallsView() {
      return /\/agents\/dashboard\/\d+/i.test(location.pathname) && this.isCallsHash();
    },
  };

  // =============================
  // STYLE INJECTOR
  // =============================
  const StyleInjector = {
    injected: false,

    inject() {
      if (this.injected) return;
      if (document.getElementById("pnfStyles")) {
        this.injected = true;
        return;
      }

      const style = document.createElement("style");
      style.id = "pnfStyles";
      style.textContent = `
        :root {
          --pnf-alpine: #FEF8E9;
          --pnf-sand: #EDDFDB;
          --pnf-mojave: #CFBA97;
          --pnf-cactus: #47B74F;
          --pnf-clover: #40A74C;
          --pnf-clover-hover: #368E40;
          --pnf-pine: #114E38;
          --pnf-pine-hover: #0D3D2B;
          --pnf-coconut: #644414;
          --pnf-cranberry: #6C2126;
          --pnf-cranberry-hover: #551A1E;
        }

        .pnf-btn {
          padding: 8px 10px;
          border: none;
          border-radius: 7px;
          cursor: pointer;
          color: #fff;
          font-weight: 700;
          font-size: 12px;
          transition: background 0.12s ease;
        }
        .pnf-btn-clover { background: var(--pnf-clover); }
        .pnf-btn-clover:hover { background: var(--pnf-clover-hover); }
        .pnf-btn-pine { background: var(--pnf-pine); }
        .pnf-btn-pine:hover { background: var(--pnf-pine-hover); }
        .pnf-btn-cranberry { background: var(--pnf-cranberry); }
        .pnf-btn-cranberry:hover { background: var(--pnf-cranberry-hover); }
        .pnf-btn-coconut { background: var(--pnf-coconut); }
        .pnf-btn-coconut:hover { filter: brightness(1.1); }
        .pnf-btn-icon {
          width: 28px;
          height: 26px;
          font-size: 16px;
          line-height: 1;
        }
        .pnf-btn-flex { flex: 1; padding: 9px 10px; }

        #patchNoteFormatterUI {
          position: fixed;
          right: 18px;
          bottom: 18px;
          z-index: 99999;
          background: #fff;
          border: 1px solid var(--pnf-mojave);
          border-radius: 10px;
          box-shadow: 0 10px 25px rgba(0,0,0,.16);
          width: 320px;
          font-family: sans-serif;
          overflow: hidden;
        }
        .pnf-ui-title {
          padding: 10px 12px;
          background: var(--pnf-pine);
          color: #fff;
          display: flex;
          align-items: center;
          justify-content: space-between;
          cursor: move;
          gap: 10px;
        }
        .pnf-ui-title-text { font-weight: 700; font-size: 13px; }
        .pnf-ui-title-btns { display: flex; align-items: center; gap: 8px; }
        .pnf-ui-body {
          padding: 12px;
          display: flex;
          flex-direction: column;
          gap: 10px;
        }
        .pnf-row { display: flex; gap: 8px; }
        .pnf-row-space {
          display: flex;
          gap: 10px;
          align-items: center;
          justify-content: space-between;
        }
        .pnf-status {
          font-size: 12px;
          min-height: 18px;
          opacity: 0.9;
        }
        .pnf-hint {
          font-size: 11px;
          opacity: 0.75;
          color: var(--pnf-coconut);
        }
        .pnf-check-label {
          display: flex;
          align-items: center;
          gap: 8px;
          font-size: 12px;
          cursor: pointer;
        }

        #patchPreviewModal, #patchTermsModal {
          position: fixed;
          inset: 0;
          z-index: 100000;
          background: rgba(17,78,56,0.55);
          display: flex;
          align-items: center;
          justify-content: center;
          padding: 18px;
        }
        .pnf-modal {
          background: #fff;
          border: 1px solid var(--pnf-mojave);
          border-radius: 10px;
          box-shadow: 0 10px 30px rgba(0,0,0,.25);
          display: flex;
          flex-direction: column;
          overflow: hidden;
          font-family: sans-serif;
          color: var(--pnf-pine);
        }
        .pnf-modal-preview {
          width: min(1200px, 95vw);
          height: min(820px, 92vh);
        }
        .pnf-modal-terms {
          width: min(860px, 95vw);
          height: min(700px, 92vh);
        }
        .pnf-modal-header {
          padding: 10px 14px;
          border-bottom: 1px solid var(--pnf-mojave);
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 10px;
          background: var(--pnf-pine);
          color: #fff;
        }
        .pnf-modal-header-title { font-weight: 700; }
        .pnf-modal-header-btns { display: flex; gap: 8px; align-items: center; }
        .pnf-modal-body {
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 12px;
          padding: 12px;
          flex: 1;
          min-height: 0;
        }
        .pnf-pane {
          border: 1px solid var(--pnf-mojave);
          border-radius: 8px;
          overflow: hidden;
          display: flex;
          flex-direction: column;
          min-height: 0;
        }
        .pnf-pane-head {
          padding: 8px 10px;
          border-bottom: 1px solid var(--pnf-mojave);
          font-weight: 700;
          background: var(--pnf-alpine);
          color: var(--pnf-pine);
        }
        .pnf-pane-body {
          padding: 10px;
          overflow: auto;
          min-height: 0;
          font-size: 12px;
          line-height: 1.4;
        }
        .pnf-pane-body-html {
          padding: 10px;
          overflow: auto;
          min-height: 0;
        }
        .pnf-change-row {
          display: flex;
          justify-content: space-between;
          gap: 12px;
        }
        .pnf-change-count { opacity: 0.7; }
        .pnf-no-changes { opacity: 0.75; }

        .pnf-diff-wrap {
          border-top: 1px solid var(--pnf-mojave);
          padding: 10px;
          overflow: auto;
          min-height: 0;
        }
        .pnf-diff-title { font-weight: 700; margin-bottom: 8px; }
        .pnf-diff-box {
          border: 1px solid var(--pnf-mojave);
          border-radius: 6px;
          overflow: hidden;
        }
        .pnf-diff-legend {
          display: flex;
          gap: 8px;
          padding: 6px 10px;
          background: var(--pnf-alpine);
          border-bottom: 1px solid var(--pnf-mojave);
          font-size: 12px;
        }
        .pnf-diff-legend-item { opacity: 0.7; }
        .pnf-diff {
          font-family: ui-monospace, Menlo, Monaco, Consolas, monospace;
          font-size: 12px;
          line-height: 1.35;
          max-height: 240px;
          overflow: auto;
        }
        .pnf-diff .row { display: flex; gap: 10px; padding: 2px 6px; }
        .pnf-diff .ln { min-width: 36px; opacity: 0.5; user-select: none; }
        .pnf-diff .cell { flex: 1; white-space: pre-wrap; word-break: break-word; }
        .pnf-diff .row.add { background: rgba(64,167,76,.12); }
        .pnf-diff .row.del { background: rgba(108,33,38,.12); }
        .pnf-diff .row.same { opacity: 0.75; }

        .pnf-after-editor {
          width: 100%;
          min-height: 420px;
          resize: vertical;
          overflow: auto;
          box-sizing: border-box;
          padding: 10px;
          border: 1px solid var(--pnf-mojave);
          border-radius: 6px;
          background: #fff;
          font-size: 13px;
          line-height: 1.45;
          outline: none;
          font-family: ui-monospace, Menlo, Monaco, Consolas, monospace;
          color: var(--pnf-pine);
        }
        .pnf-modal-footer {
          padding: 10px 14px;
          border-top: 1px solid var(--pnf-mojave);
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 10px;
          background: var(--pnf-alpine);
          color: var(--pnf-pine);
        }
        .pnf-footer-text { font-size: 12px; opacity: 0.8; }
        .pnf-help-text { font-size: 11px; opacity: 0.7; }

        .pnf-terms-help {
          font-size: 12px;
          color: var(--pnf-coconut);
          padding: 10px 14px;
          background: var(--pnf-alpine);
          border-bottom: 1px solid var(--pnf-mojave);
          line-height: 1.5;
        }
        .pnf-terms-body {
          flex: 1;
          padding: 12px;
          display: flex;
          flex-direction: column;
          gap: 8px;
          min-height: 0;
        }
        .pnf-terms-textarea {
          flex: 1;
          width: 100%;
          font-family: ui-monospace, Menlo, Monaco, Consolas, monospace;
          font-size: 12px;
          box-sizing: border-box;
          padding: 10px;
          border: 1px solid var(--pnf-mojave);
          border-radius: 6px;
          resize: none;
          outline: none;
          color: var(--pnf-pine);
        }
        .pnf-terms-textarea:focus { border-color: var(--pnf-pine); }
        .pnf-terms-status {
          font-size: 12px;
          min-height: 18px;
          padding: 0 2px;
        }
        .pnf-terms-status.error { color: var(--pnf-cranberry); }
        .pnf-terms-status.ok { color: var(--pnf-clover); }

        #patchAddTermDialog {
          position: fixed;
          inset: 0;
          z-index: 100001;
          background: rgba(0,0,0,0.3);
          display: flex;
          align-items: center;
          justify-content: center;
          padding: 18px;
        }
        .pnf-small-panel {
          background: #fff;
          border: 1px solid var(--pnf-mojave);
          border-radius: 10px;
          box-shadow: 0 10px 30px rgba(0,0,0,.25);
          width: min(480px, 92vw);
          padding: 16px;
          color: var(--pnf-pine);
          font-family: sans-serif;
          display: flex;
          flex-direction: column;
          gap: 10px;
        }
        .pnf-small-title { font-weight: 700; font-size: 14px; }
        .pnf-small-hint { font-size: 11px; opacity: 0.75; }
        .pnf-small-label { font-size: 12px; font-weight: 600; }
        .pnf-small-input {
          padding: 8px 10px;
          border: 1px solid var(--pnf-mojave);
          border-radius: 6px;
          font-size: 13px;
          outline: none;
          font-family: ui-monospace, Menlo, Monaco, Consolas, monospace;
          color: var(--pnf-pine);
          box-sizing: border-box;
          width: 100%;
        }
        .pnf-small-input:focus { border-color: var(--pnf-pine); }
        .pnf-small-status { font-size: 12px; min-height: 16px; }
        .pnf-small-status.error { color: var(--pnf-cranberry); }
        .pnf-small-status.ok { color: var(--pnf-clover); }
        .pnf-small-btns {
          display: flex;
          justify-content: flex-end;
          gap: 8px;
          margin-top: 4px;
        }
      `;
      document.head.appendChild(style);
      this.injected = true;
    },
  };

  // =============================
  // TERM STORE (merges defaults with user overrides)
  // =============================
  const TermStore = {
    getUserOverrides() {
      return Storage.get(CONFIG.TERMS_KEY, {}) || {};
    },

    setUserOverrides(overrides) {
      return Storage.set(CONFIG.TERMS_KEY, overrides || {});
    },

    getMerged() {
      const overrides = this.getUserOverrides();
      const merged = {};
      for (const [canon, alts] of Object.entries(DEFAULT_TERM_NORMALIZATION)) {
        merged[canon] = [...alts];
      }
      for (const [canon, alts] of Object.entries(overrides)) {
        if (!Array.isArray(alts)) continue;
        merged[canon] = [...(merged[canon] || []), ...alts];
      }
      return merged;
    },

    validateOverrides(obj) {
      if (!obj || typeof obj !== "object" || Array.isArray(obj)) {
        return "Top-level must be a JSON object mapping canonical → array of alternates.";
      }
      for (const [canon, alts] of Object.entries(obj)) {
        if (typeof canon !== "string" || !canon) {
          return `Canonical keys must be non-empty strings (got ${JSON.stringify(canon)}).`;
        }
        if (!Array.isArray(alts)) {
          return `Value for "${canon}" must be an array of strings.`;
        }
        for (const alt of alts) {
          if (typeof alt !== "string") {
            return `All alternates for "${canon}" must be strings.`;
          }
        }
      }
      return null;
    },
  };

  // =============================
  // TERM NORMALIZATION ENGINE (O(1) lookup via Map; rebuildable)
  // =============================
  const TermNormalizer = (() => {
    let altIndex = new Map();
    let normalizationRegex = null;

    function rebuild() {
      const merged = TermStore.getMerged();
      const termEntries = Object.entries(merged).flatMap(([canonical, alts]) =>
        alts
          .filter((alt) => alt !== canonical)
          .map((alt) => ({ canonical, alt }))
      );

      altIndex = new Map(termEntries.map((e) => [e.alt.toLowerCase(), e.canonical]));

      const parts = termEntries
        .slice()
        .sort((a, b) => b.alt.length - a.alt.length)
        .map(({ alt }) => Utils.escapeRegExp(alt));

      normalizationRegex = parts.length
        ? new RegExp(`\\b(?:${parts.join("|")})\\b(?!\\.[a-z]{2,})`, "gi")
        : null;
    }

    rebuild();

    return {
      rebuild,

      normalize(input) {
        if (!normalizationRegex || !input) {
          return { text: input || "", changeList: [] };
        }

        const changesMap = new Map();

        const output = input.replace(normalizationRegex, (match) => {
          const canonical = altIndex.get(match.toLowerCase());
          if (!canonical) return match;

          if (match !== canonical) {
            const key = `${match} → ${canonical}`;
            changesMap.set(key, (changesMap.get(key) || 0) + 1);
          }

          return canonical;
        });

        const changeList = Array.from(changesMap.entries())
          .map(([k, v]) => ({ k, v }))
          .sort((a, b) => b.v - a.v || a.k.localeCompare(b.k));

        return { text: output, changeList };
      },
    };
  })();

  // =============================
  // HTML PROCESSING
  // =============================
  const HtmlProcessor = {
    htmlToStructuredText(html) {
      const root = document.createElement("div");
      root.innerHTML = html || "";

      const out = [];
      const olStack = [];
      let insideListItem = false;

      const pushText = (text) => {
        if (!text) return;
        out.push(text);
      };

      const normalizeSpace = (str) =>
        str.replace(/[ \t]+\n/g, "\n").replace(/[ \t]{2,}/g, " ");

      const walk = (node) => {
        if (!node) return;

        if (node.nodeType === Node.TEXT_NODE) {
          const text = (node.nodeValue || "").replace(/\s+/g, " ");
          pushText(text);
          return;
        }

        if (node.nodeType !== Node.ELEMENT_NODE) return;

        const tag = (node.tagName || "").toUpperCase();

        if (tag === "BR") {
          pushText("\n");
          return;
        }

        if (tag === "OL") {
          olStack.push(0);
          pushText("\n");
          Array.from(node.childNodes).forEach(walk);
          pushText("\n");
          olStack.pop();
          return;
        }

        if (tag === "UL") {
          pushText("\n");
          Array.from(node.childNodes).forEach(walk);
          pushText("\n");
          return;
        }

        if (tag === "LI") {
          pushText("\n");
          if (olStack.length) {
            olStack[olStack.length - 1] += 1;
            pushText(`${olStack[olStack.length - 1]}. `);
          } else {
            pushText("- ");
          }
          // Save/restore so nested LIs don't flip the flag off for the outer LI's
          // remaining children.
          const prevInsideLI = insideListItem;
          insideListItem = true;
          Array.from(node.childNodes).forEach(walk);
          insideListItem = prevInsideLI;
          return;
        }

        const isBlock = ["P", "DIV", "SECTION", "ARTICLE", "HEADER", "FOOTER", "H1", "H2", "H3", "H4", "H5", "H6"].includes(
          tag
        );

        if (isBlock && !insideListItem) pushText("\n");

        Array.from(node.childNodes).forEach(walk);

        if (isBlock && !insideListItem) pushText("\n");
      };

      Array.from(root.childNodes).forEach(walk);

      const joined = out.join("");
      return normalizeSpace(joined)
        .replace(/\n{3,}/g, "\n\n")
        .replace(/[ \t]+\n/g, "\n")
        .trim();
    },
  };

  // =============================
  // EDITOR INTERFACE
  // =============================
  const EditorInterface = {
    findKendoEditorTextarea() {
      return document.querySelector('textarea[aria-label="editor"]');
    },

    getKendoEditor() {
      if (!window.kendo?.ui?.Editor) return null;
      const ta = this.findKendoEditorTextarea();
      if (!ta) return null;
      const $ = window.jQuery || window.$;
      if (!$) return null;
      return $(ta).data("kendoEditor") || null;
    },

    isEditorOpen() {
      return !!this.findKendoEditorTextarea();
    },

    getEditorRawText() {
      try {
        const editor = this.getKendoEditor();
        if (editor) {
          const html = editor.value() || "";
          return HtmlProcessor.htmlToStructuredText(html);
        }

        const ta = this.findKendoEditorTextarea();
        if (!ta) return "";

        const raw = ta.value || "";

        if (Utils.looksLikeEscapedHtml(raw)) {
          const decodedHtml = Utils.decodeHtmlEntities(raw);
          return HtmlProcessor.htmlToStructuredText(decodedHtml);
        }

        return raw;
      } catch (e) {
        console.error("[PatchNoteFormatter] Error reading editor:", e);
        return "";
      }
    },

    async setEditorHtml(html) {
      try {
        const editor = this.getKendoEditor();
        if (editor) {
          editor.value(html);
          return { ok: true };
        }

        const ta = this.findKendoEditorTextarea();
        if (ta) {
          const existing = ta.value || "";
          if (Utils.looksLikeEscapedHtml(existing)) {
            ta.value = Utils.encodeHtmlEntities(html);
          } else {
            ta.value = html;
          }
          return { ok: true };
        }

        return { ok: false, error: "Editor not found." };
      } catch (e) {
        console.error("[PatchNoteFormatter] Error setting editor:", e);
        return { ok: false, error: e.message || "Unknown error" };
      }
    },
  };

  // =============================
  // TEXT FORMATTING
  // =============================
  const TextFormatter = {
    parseBlocksPreserveLists(lines) {
      const blocks = [];
      let currentPara = [];
      let currentList = null;

      const flushPara = () => {
        if (!currentPara.length) return;
        const content = currentPara.join(" ").trim();
        if (content) blocks.push({ type: "p", text: content });
        currentPara = [];
      };

      const flushList = () => {
        if (!currentList) return;
        blocks.push(currentList);
        currentList = null;
      };

      const isBullet = (line) => /^\s*[-*•]\s+/.test(line);
      const isNumbered = (line) => /^\s*\d+[\.\)]\s+/.test(line);

      const isHeadingLike = (line) => {
        const trimmed = line.trim();
        if (!trimmed) return false;
        if (isBullet(trimmed) || isNumbered(trimmed)) return false;
        if (trimmed.length > 80) return false;
        if (/[.:]$/.test(trimmed)) return true;
        if (/^[A-Z][A-Za-z0-9\s/&'\-]{2,40}$/.test(trimmed) && !/[.?!]$/.test(trimmed)) return true;
        return false;
      };

      for (let i = 0; i < lines.length; i++) {
        const raw = lines[i];
        const line = raw.trimEnd();
        const trimmed = line.trim();

        if (!trimmed) {
          flushPara();
          flushList();
          continue;
        }

        if (isBullet(trimmed) || isNumbered(trimmed)) {
          flushPara();

          const listType = isNumbered(trimmed) ? "ol" : "ul";
          if (!currentList || currentList.listType !== listType) {
            flushList();
            currentList = { type: "list", listType, items: [] };
          }

          const itemText = trimmed.replace(/^\s*([-*•]|\d+[\.\)])\s+/, "");
          currentList.items.push(itemText);
          continue;
        }

        if (currentList) {
          const last = currentList.items.length - 1;
          currentList.items[last] = `${currentList.items[last]} ${trimmed}`.trim();
          continue;
        }

        if (isHeadingLike(trimmed)) {
          flushPara();
          flushList();
          blocks.push({ type: "h", text: trimmed.replace(/[:.]$/, "") });
          continue;
        }

        currentPara.push(trimmed);
      }

      flushPara();
      flushList();
      return blocks;
    },

    blocksToHtml(blocks) {
      const out = [];
      for (const block of blocks) {
        if (block.type === "h") {
          out.push(`<p><strong>${Utils.escapeHtml(block.text)}</strong></p>`);
        } else if (block.type === "p") {
          out.push(`<p>${Utils.escapeHtml(block.text)}</p>`);
        } else if (block.type === "list") {
          const tag = block.listType === "ol" ? "ol" : "ul";
          const items = block.items.map((item) => `<li>${Utils.escapeHtml(item)}</li>`).join("");
          out.push(`<${tag}>${items}</${tag}>`);
        }
      }
      return out.join("");
    },

    buildFormattedHtmlFromText(text) {
      const lines = text.replace(/\r\n/g, "\n").split("\n");
      const blocks = this.parseBlocksPreserveLists(lines);
      return this.blocksToHtml(blocks);
    },
  };

  // =============================
  // DIFF GENERATOR
  // =============================
  const DiffGenerator = {
    buildLineDiffHtml(before, after) {
      const a = before.replace(/\r\n/g, "\n").split("\n");
      const b = after.replace(/\r\n/g, "\n").split("\n");
      const dp = Array(a.length + 1)
        .fill(null)
        .map(() => Array(b.length + 1).fill(0));

      for (let i = a.length - 1; i >= 0; i--) {
        for (let j = b.length - 1; j >= 0; j--) {
          dp[i][j] = a[i] === b[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
        }
      }

      const rows = [];
      let i = 0;
      let j = 0;
      while (i < a.length && j < b.length) {
        if (a[i] === b[j]) {
          rows.push({ type: "same", text: a[i] });
          i++;
          j++;
        } else if (dp[i + 1][j] >= dp[i][j + 1]) {
          rows.push({ type: "del", text: a[i] });
          i++;
        } else {
          rows.push({ type: "add", text: b[j] });
          j++;
        }
      }
      while (i < a.length) rows.push({ type: "del", text: a[i++] });
      while (j < b.length) rows.push({ type: "add", text: b[j++] });

      let lineNum = 1;
      return rows
        .slice(0, CONFIG.MAX_DIFF_LINES)
        .map((row) => {
          const cls = ["add", "del", "same"].includes(row.type) ? row.type : "same";
          const ln = String(lineNum++).padStart(3, " ");
          return `<div class="row ${cls}"><div class="ln">${ln}</div><div class="cell">${Utils.escapeHtml(
            row.text
          )}</div></div>`;
        })
        .join("");
    },
  };

  // =============================
  // UI COMPONENTS
  // =============================
  const UIComponents = {
    createButton({ text, variant = "pine", onClick, ariaLabel, extraClass = "" }) {
      const btn = document.createElement("button");
      btn.textContent = text;
      if (ariaLabel) btn.setAttribute("aria-label", ariaLabel);
      btn.className = `pnf-btn pnf-btn-${variant}${extraClass ? " " + extraClass : ""}`;
      if (onClick) btn.onclick = onClick;
      return btn;
    },
  };

  // =============================
  // BUTTON HELPERS
  // =============================
  const ButtonHelpers = {
    findByText(matcher, buttons = null) {
      const list = buttons || document.querySelectorAll("button");
      for (const btn of list) {
        const text = (btn.textContent || "").trim().toLowerCase();
        if (matcher(text)) return btn;
      }
      return null;
    },

    findSaveButton() {
      return this.findByText((t) => t === "save");
    },

    async clickSave() {
      try {
        const btn = this.findSaveButton();
        if (!btn) return { ok: false, error: "Save button not found." };
        btn.click();
        await Utils.sleep(CONFIG.EDITOR_SLEEP_MS);
        return { ok: true };
      } catch (e) {
        console.error("[PatchNoteFormatter] Save click error:", e);
        return { ok: false, error: e.message || "Save failed" };
      }
    },
  };

  // =============================
  // APPLICATION STATE
  // =============================
  const AppState = {
    autoSave: CONFIG.DEFAULT_AUTO_SAVE,
  };

  // =============================
  // PREVIEW MODAL
  // =============================
  const PreviewModal = {
    remove() {
      const el = document.getElementById("patchPreviewModal");
      if (el) el.remove();
    },

    build({ beforeText, normalizedText, formattedHtml, changes }) {
      this.remove();

      const limitedBefore =
        beforeText.length > CONFIG.MAX_PREVIEW_CHARS
          ? beforeText.slice(0, CONFIG.MAX_PREVIEW_CHARS) + "\n\n[TRUNCATED]"
          : beforeText;
      const limitedAfter =
        normalizedText.length > CONFIG.MAX_PREVIEW_CHARS
          ? normalizedText.slice(0, CONFIG.MAX_PREVIEW_CHARS) + "\n\n[TRUNCATED]"
          : normalizedText;

      const overlay = document.createElement("div");
      overlay.id = "patchPreviewModal";

      const modal = document.createElement("div");
      modal.className = "pnf-modal pnf-modal-preview";
      modal.setAttribute("role", "dialog");
      modal.setAttribute("aria-modal", "true");
      modal.setAttribute("aria-labelledby", "previewModalTitle");
      modal.tabIndex = -1;

      const header = this._buildHeader();
      const btnRow = this._buildButtonRow();
      header.appendChild(btnRow.container);

      const body = this._buildBody({ limitedBefore, limitedAfter, changes, formattedHtml });
      const footer = this._buildFooter();

      modal.appendChild(header);
      modal.appendChild(body.container);
      modal.appendChild(footer);
      overlay.appendChild(modal);

      const handleEsc = (e) => {
        if (e.key === "Escape") {
          this.remove();
          document.removeEventListener("keydown", handleEsc);
        }
      };
      document.addEventListener("keydown", handleEsc);

      this._attachFocusTrap(modal);

      return [overlay, btnRow.btnCancel, btnRow.btnRenormalize, btnRow.btnApply, body.afterEditor];
    },

    _attachFocusTrap(modal) {
      modal.addEventListener("keydown", (e) => {
        if (e.key !== "Tab") return;
        const focusable = modal.querySelectorAll(
          'button:not([disabled]), [href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"]), [contenteditable="true"]'
        );
        if (!focusable.length) return;
        const first = focusable[0];
        const last = focusable[focusable.length - 1];
        if (e.shiftKey && document.activeElement === first) {
          e.preventDefault();
          last.focus();
        } else if (!e.shiftKey && document.activeElement === last) {
          e.preventDefault();
          first.focus();
        }
      });
      // Focus the modal shell on open so Tab traps from the start.
      setTimeout(() => modal.focus(), 0);
    },

    _buildHeader() {
      const header = document.createElement("div");
      header.className = "pnf-modal-header";

      const title = document.createElement("div");
      title.id = "previewModalTitle";
      title.className = "pnf-modal-header-title";
      title.textContent = "Preview: Normalize terms + Format note";
      header.appendChild(title);

      return header;
    },

    _buildButtonRow() {
      const container = document.createElement("div");
      container.className = "pnf-modal-header-btns";

      const btnCancel = UIComponents.createButton({
        text: "Cancel",
        variant: "cranberry",
        ariaLabel: "Cancel preview",
      });

      const btnRenormalize = UIComponents.createButton({
        text: "Re-normalize",
        variant: "coconut",
        ariaLabel: "Re-normalize edited content",
      });

      const btnAddTerm = UIComponents.createButton({
        text: "+ Add term",
        variant: "pine",
        ariaLabel: "Add term normalization from selection",
      });

      const btnCopy = UIComponents.createButton({
        text: "Copy",
        variant: "pine",
        ariaLabel: "Copy to clipboard",
      });

      const btnApply = UIComponents.createButton({
        text: "Apply to Editor",
        variant: "clover",
        ariaLabel: "Apply changes to editor",
      });

      container.appendChild(btnCancel);
      container.appendChild(btnRenormalize);
      container.appendChild(btnAddTerm);
      container.appendChild(btnCopy);
      container.appendChild(btnApply);

      return { container, btnCancel, btnRenormalize, btnAddTerm, btnCopy, btnApply };
    },

    _buildBody({ limitedBefore, limitedAfter, changes, formattedHtml }) {
      const body = document.createElement("div");
      body.className = "pnf-modal-body";

      const left = this._buildLeftPane({ limitedBefore, limitedAfter, changes });
      const right = this._buildRightPane({ formattedHtml });

      body.appendChild(left);
      body.appendChild(right.container);

      return { container: body, afterEditor: right.afterEditor };
    },

    _buildLeftPane({ limitedBefore, limitedAfter, changes }) {
      const left = document.createElement("div");
      left.className = "pnf-pane";

      const leftHead = document.createElement("div");
      leftHead.className = "pnf-pane-head";
      leftHead.textContent = "Term changes (before → normalized)";

      const leftBody = document.createElement("div");
      leftBody.id = "patchChangesList";
      leftBody.className = "pnf-pane-body";

      this._renderChanges(leftBody, changes);

      const diffWrap = this._buildDiffSection({ limitedBefore, limitedAfter });

      left.appendChild(leftHead);
      left.appendChild(leftBody);
      left.appendChild(diffWrap);

      return left;
    },

    _renderChanges(leftBody, changes) {
      if (!changes.length) {
        leftBody.innerHTML = '<div class="pnf-no-changes">No term normalization changes detected.</div>';
        return;
      }
      leftBody.innerHTML = changes
        .slice(0, CONFIG.MAX_CHANGE_ENTRIES)
        .map(
          (c) =>
            `<div class="pnf-change-row"><div>${Utils.escapeHtml(
              c.k
            )}</div><div class="pnf-change-count">${c.v}×</div></div>`
        )
        .join("");
    },

    _buildDiffSection({ limitedBefore, limitedAfter }) {
      const diffWrap = document.createElement("div");
      diffWrap.className = "pnf-diff-wrap";

      const diffTitle = document.createElement("div");
      diffTitle.className = "pnf-diff-title";
      diffTitle.textContent = "Line diff (raw text)";
      diffWrap.appendChild(diffTitle);

      const diffBox = document.createElement("div");
      diffBox.className = "pnf-diff-box";

      const legend = document.createElement("div");
      legend.className = "pnf-diff-legend";

      const legendDel = document.createElement("div");
      legendDel.className = "pnf-diff-legend-item";
      legendDel.textContent = "- removed";

      const legendAdd = document.createElement("div");
      legendAdd.className = "pnf-diff-legend-item";
      legendAdd.textContent = "+ added";

      legend.appendChild(legendDel);
      legend.appendChild(legendAdd);

      const diffPane = document.createElement("div");
      diffPane.className = "pnf-diff";
      diffPane.innerHTML = DiffGenerator.buildLineDiffHtml(limitedBefore, limitedAfter);

      diffBox.appendChild(legend);
      diffBox.appendChild(diffPane);
      diffWrap.appendChild(diffBox);

      return diffWrap;
    },

    _buildRightPane({ formattedHtml }) {
      const right = document.createElement("div");
      right.className = "pnf-pane";

      const rightHead = document.createElement("div");
      rightHead.className = "pnf-pane-head";
      rightHead.textContent = "Formatted preview (what will be inserted)";

      const rightBody = document.createElement("div");
      rightBody.className = "pnf-pane-body-html";

      const afterEditor = document.createElement("div");
      afterEditor.className = "pnf-after-editor";
      afterEditor.contentEditable = "true";
      afterEditor.spellcheck = false;
      afterEditor.setAttribute("role", "textbox");
      afterEditor.setAttribute("aria-label", "Formatted output preview");

      afterEditor.innerHTML = formattedHtml;

      afterEditor.addEventListener("paste", (e) => {
        e.preventDefault();

        const pastedText = (e.clipboardData || window.clipboardData).getData("text/plain");
        if (!pastedText) return;

        const formattedPaste = TextFormatter.buildFormattedHtmlFromText(pastedText);

        const selection = window.getSelection();
        if (!selection.rangeCount) return;

        const range = selection.getRangeAt(0);
        range.deleteContents();

        const temp = document.createElement("div");
        temp.innerHTML = formattedPaste;
        const frag = document.createDocumentFragment();
        Array.from(temp.childNodes).forEach((node) => frag.appendChild(node));
        range.insertNode(frag);

        range.collapse(false);
        selection.removeAllRanges();
        selection.addRange(range);
      });

      rightBody.appendChild(afterEditor);
      right.appendChild(rightHead);
      right.appendChild(rightBody);

      return { container: right, afterEditor };
    },

    _buildFooter() {
      const footer = document.createElement("div");
      footer.className = "pnf-modal-footer";

      const footerText = document.createElement("div");
      footerText.className = "pnf-footer-text";
      footerText.appendChild(document.createTextNode("Auto-save is "));

      const strong = document.createElement("strong");
      strong.textContent = AppState.autoSave ? "ON" : "OFF";
      footerText.appendChild(strong);

      const helpText = document.createElement("div");
      helpText.className = "pnf-help-text";
      helpText.textContent = 'Edit above as needed, then "Apply"';

      footer.appendChild(footerText);
      footer.appendChild(helpText);

      return footer;
    },

    renderChangesInto(leftBody, changes) {
      this._renderChanges(leftBody, changes);
    },
  };

  // =============================
  // MANAGE TERMS MODAL (user-editable overrides merged with defaults)
  // =============================
  const ManageTermsModal = {
    remove() {
      const el = document.getElementById("patchTermsModal");
      if (el) el.remove();
    },

    open() {
      this.remove();

      const overlay = document.createElement("div");
      overlay.id = "patchTermsModal";

      const modal = document.createElement("div");
      modal.className = "pnf-modal pnf-modal-terms";
      modal.setAttribute("role", "dialog");
      modal.setAttribute("aria-modal", "true");
      modal.setAttribute("aria-labelledby", "termsModalTitle");
      modal.tabIndex = -1;

      // Header
      const header = document.createElement("div");
      header.className = "pnf-modal-header";

      const title = document.createElement("div");
      title.id = "termsModalTitle";
      title.className = "pnf-modal-header-title";
      title.textContent = "Manage term normalization (user overrides)";
      header.appendChild(title);

      const headerBtns = document.createElement("div");
      headerBtns.className = "pnf-modal-header-btns";

      const btnCancel = UIComponents.createButton({
        text: "Close",
        variant: "cranberry",
        ariaLabel: "Close terms editor",
      });
      const btnReset = UIComponents.createButton({
        text: "Reset to defaults",
        variant: "coconut",
        ariaLabel: "Clear all user overrides",
      });
      const btnSave = UIComponents.createButton({
        text: "Save",
        variant: "clover",
        ariaLabel: "Save overrides",
      });

      headerBtns.appendChild(btnCancel);
      headerBtns.appendChild(btnReset);
      headerBtns.appendChild(btnSave);
      header.appendChild(headerBtns);

      // Help
      const help = document.createElement("div");
      help.className = "pnf-terms-help";
      help.innerHTML =
        'Paste a JSON object of <strong>canonical → array of alternates</strong>. Entries here are <em>added to</em> the shipped defaults (duplicates are harmless). Example: <code>{ "Scorecard": ["scoreboard", "score-cards"] }</code>';

      // Body
      const body = document.createElement("div");
      body.className = "pnf-terms-body";

      const textarea = document.createElement("textarea");
      textarea.className = "pnf-terms-textarea";
      textarea.spellcheck = false;
      textarea.value = JSON.stringify(TermStore.getUserOverrides(), null, 2);

      const status = document.createElement("div");
      status.className = "pnf-terms-status";

      body.appendChild(textarea);
      body.appendChild(status);

      modal.appendChild(header);
      modal.appendChild(help);
      modal.appendChild(body);
      overlay.appendChild(modal);

      const cleanup = () => this.remove();

      btnCancel.onclick = cleanup;
      overlay.addEventListener("click", (e) => {
        if (e.target === overlay) cleanup();
      });

      const handleEsc = (e) => {
        if (e.key === "Escape") {
          cleanup();
          document.removeEventListener("keydown", handleEsc);
        }
      };
      document.addEventListener("keydown", handleEsc);

      btnReset.onclick = () => {
        if (!confirm("Clear all user overrides? Shipped defaults will remain.")) return;
        TermStore.setUserOverrides({});
        TermNormalizer.rebuild();
        textarea.value = "{}";
        status.textContent = "User overrides cleared.";
        status.className = "pnf-terms-status ok";
      };

      btnSave.onclick = () => {
        let parsed;
        try {
          parsed = JSON.parse(textarea.value || "{}");
        } catch (e) {
          status.textContent = `Invalid JSON: ${e.message}`;
          status.className = "pnf-terms-status error";
          return;
        }
        const err = TermStore.validateOverrides(parsed);
        if (err) {
          status.textContent = err;
          status.className = "pnf-terms-status error";
          return;
        }
        TermStore.setUserOverrides(parsed);
        TermNormalizer.rebuild();
        status.textContent = "Saved. Normalizer reloaded.";
        status.className = "pnf-terms-status ok";
      };

      // Focus trap
      modal.addEventListener("keydown", (e) => {
        if (e.key !== "Tab") return;
        const focusable = modal.querySelectorAll(
          'button:not([disabled]), input:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])'
        );
        if (!focusable.length) return;
        const first = focusable[0];
        const last = focusable[focusable.length - 1];
        if (e.shiftKey && document.activeElement === first) {
          e.preventDefault();
          last.focus();
        } else if (!e.shiftKey && document.activeElement === last) {
          e.preventDefault();
          first.focus();
        }
      });

      document.body.appendChild(overlay);
      setTimeout(() => textarea.focus(), 0);
    },
  };

  // =============================
  // ADD TERM DIALOG (small popover inside preview flow)
  // =============================
  const AddTermDialog = {
    remove() {
      const el = document.getElementById("patchAddTermDialog");
      if (el) el.remove();
    },

    open({ prefillAlt = "", onSave } = {}) {
      this.remove();

      const overlay = document.createElement("div");
      overlay.id = "patchAddTermDialog";

      const panel = document.createElement("div");
      panel.className = "pnf-small-panel";
      panel.setAttribute("role", "dialog");
      panel.setAttribute("aria-modal", "true");
      panel.setAttribute("aria-labelledby", "addTermTitle");
      panel.tabIndex = -1;

      const title = document.createElement("div");
      title.id = "addTermTitle";
      title.className = "pnf-small-title";
      title.textContent = "Add term normalization";

      const hint = document.createElement("div");
      hint.className = "pnf-small-hint";
      hint.textContent =
        "Saves to user overrides and re-normalizes the preview immediately. Tip: select text in the preview before clicking to prefill.";

      const altLbl = document.createElement("div");
      altLbl.className = "pnf-small-label";
      altLbl.textContent = "Alternate (the variant to normalize)";
      const altInp = document.createElement("input");
      altInp.type = "text";
      altInp.value = prefillAlt;
      altInp.className = "pnf-small-input";
      altInp.placeholder = "e.g. multi line";

      const canonLbl = document.createElement("div");
      canonLbl.className = "pnf-small-label";
      canonLbl.textContent = "Canonical (what it becomes)";
      const canonInp = document.createElement("input");
      canonInp.type = "text";
      canonInp.className = "pnf-small-input";
      canonInp.placeholder = "e.g. multi-line";

      const status = document.createElement("div");
      status.className = "pnf-small-status";

      const btnRow = document.createElement("div");
      btnRow.className = "pnf-small-btns";

      const btnCancel = UIComponents.createButton({
        text: "Cancel",
        variant: "cranberry",
        ariaLabel: "Cancel add term",
      });
      const btnSave = UIComponents.createButton({
        text: "Add",
        variant: "clover",
        ariaLabel: "Save new term",
      });

      btnRow.appendChild(btnCancel);
      btnRow.appendChild(btnSave);

      panel.appendChild(title);
      panel.appendChild(hint);
      panel.appendChild(altLbl);
      panel.appendChild(altInp);
      panel.appendChild(canonLbl);
      panel.appendChild(canonInp);
      panel.appendChild(status);
      panel.appendChild(btnRow);
      overlay.appendChild(panel);

      const cleanup = () => this.remove();

      btnCancel.onclick = cleanup;
      overlay.addEventListener("click", (e) => {
        if (e.target === overlay) cleanup();
      });

      const escHandler = (e) => {
        if (e.key === "Escape") {
          e.stopPropagation();
          cleanup();
          document.removeEventListener("keydown", escHandler, true);
        }
      };
      // Capture-phase so the preview modal's own ESC handler doesn't also fire.
      document.addEventListener("keydown", escHandler, true);

      const submit = () => {
        const alt = altInp.value.trim();
        const canon = canonInp.value.trim();
        if (!alt || !canon) {
          status.textContent = "Both fields are required.";
          status.className = "pnf-small-status error";
          return;
        }
        if (alt === canon) {
          status.textContent = "Alternate must differ from canonical.";
          status.className = "pnf-small-status error";
          return;
        }

        const overrides = TermStore.getUserOverrides();
        const existingAlts = Array.isArray(overrides[canon]) ? overrides[canon] : [];
        if (!existingAlts.includes(alt)) {
          overrides[canon] = [...existingAlts, alt];
          TermStore.setUserOverrides(overrides);
        }
        TermNormalizer.rebuild();

        cleanup();
        document.removeEventListener("keydown", escHandler, true);
        if (onSave) onSave({ alt, canon });
      };

      btnSave.onclick = submit;

      // Enter in either field submits
      [altInp, canonInp].forEach((inp) => {
        inp.addEventListener("keydown", (e) => {
          if (e.key === "Enter") {
            e.preventDefault();
            submit();
          }
        });
      });

      // Focus trap
      panel.addEventListener("keydown", (e) => {
        if (e.key !== "Tab") return;
        const focusable = panel.querySelectorAll(
          'button:not([disabled]), input:not([disabled]), [tabindex]:not([tabindex="-1"])'
        );
        if (!focusable.length) return;
        const first = focusable[0];
        const last = focusable[focusable.length - 1];
        if (e.shiftKey && document.activeElement === first) {
          e.preventDefault();
          last.focus();
        } else if (!e.shiftKey && document.activeElement === last) {
          e.preventDefault();
          first.focus();
        }
      });

      document.body.appendChild(overlay);
      // If alternate was prefilled from selection, focus canonical next.
      setTimeout(() => (prefillAlt ? canonInp : altInp).focus(), 0);
    },
  };

  // =============================
  // MAIN ACTIONS
  // =============================
  const Actions = {
    async applyHtmlToEditorAndMaybeSave(formattedHtml) {
      try {
        const setRes = await EditorInterface.setEditorHtml(formattedHtml);
        if (!setRes.ok) return { ok: false, info: setRes.error };

        if (AppState.autoSave) {
          const saved = await ButtonHelpers.clickSave();
          if (!saved.ok) return { ok: false, info: saved.error };
          return { ok: true, info: "Applied and saved." };
        }
        return { ok: true, info: "Applied to editor (not saved)." };
      } catch (e) {
        console.error("[PatchNoteFormatter] Apply error:", e);
        return { ok: false, info: e.message || "Error applying changes" };
      }
    },

    // Build the modal, wire all buttons, and resolve when the user
    // applies/cancels/closes. Used by both the editor-preview and clipboard-preview flows.
    async _showPreviewAndApply({ beforeText, normalizedText, formattedHtml, changeList }) {
      const [overlay, btnCancel, btnRenormalize, btnApply, afterEditor] = PreviewModal.build({
        beforeText,
        normalizedText,
        formattedHtml,
        changes: changeList,
      });

      document.body.appendChild(overlay);

      btnRenormalize.onclick = () => {
        try {
          const currentHtml = afterEditor.innerHTML || "";
          const currentText = HtmlProcessor.htmlToStructuredText(currentHtml);
          const { text: renormalizedText, changeList: newChanges } = TermNormalizer.normalize(currentText);
          afterEditor.innerHTML = TextFormatter.buildFormattedHtmlFromText(renormalizedText);

          const leftBody = document.getElementById("patchChangesList");
          if (leftBody) PreviewModal.renderChangesInto(leftBody, newChanges);

          const originalText = btnRenormalize.textContent;
          btnRenormalize.textContent = "Re-normalized!";
          setTimeout(() => (btnRenormalize.textContent = originalText), 1500);
        } catch (e) {
          console.error("Re-normalize failed:", e);
          alert("Re-normalize failed. See console for details.");
        }
      };

      const btnCopy = overlay.querySelector('button[aria-label="Copy to clipboard"]');
      if (btnCopy) {
        btnCopy.onclick = async () => {
          try {
            const structuredText = HtmlProcessor.htmlToStructuredText(afterEditor.innerHTML);
            await navigator.clipboard.writeText(structuredText);
            btnCopy.textContent = "Copied!";
            setTimeout(() => (btnCopy.textContent = "Copy"), 1500);
          } catch (e) {
            console.error("Copy failed:", e);
            alert("Copy failed. See console for details.");
          }
        };
      }

      const btnAddTerm = overlay.querySelector(
        'button[aria-label="Add term normalization from selection"]'
      );
      if (btnAddTerm) {
        btnAddTerm.onclick = () => {
          // Pull selection from the preview contenteditable if present.
          let selectedText = "";
          const sel = window.getSelection();
          if (sel && sel.toString) {
            selectedText = sel.toString().trim();
          }
          // Only prefill if the selection actually lives inside our modal.
          if (selectedText && sel.rangeCount) {
            const range = sel.getRangeAt(0);
            if (!overlay.contains(range.commonAncestorContainer)) {
              selectedText = "";
            }
          }

          AddTermDialog.open({
            prefillAlt: selectedText,
            onSave: ({ alt, canon }) => {
              // Trigger a re-normalize so the new rule takes effect on the live preview.
              btnRenormalize.click();
              const orig = btnAddTerm.textContent;
              btnAddTerm.textContent = `Added "${alt}" → "${canon}"`;
              setTimeout(() => (btnAddTerm.textContent = orig), 1800);
            },
          });
        };
      }

      return new Promise((resolve) => {
        const cleanup = () => PreviewModal.remove();

        btnCancel.onclick = () => {
          cleanup();
          resolve({ ok: true, info: "Canceled. No changes applied." });
        };

        btnApply.onclick = async () => {
          try {
            const res = await this.applyHtmlToEditorAndMaybeSave(afterEditor.innerHTML);
            cleanup();
            resolve({ ok: res.ok, info: res.info });
          } catch (e) {
            cleanup();
            resolve({ ok: false, info: e.message || "Error applying changes." });
          }
        };

        overlay.addEventListener("click", (e) => {
          if (e.target === overlay) {
            cleanup();
            resolve({ ok: true, info: "Closed preview. No changes applied." });
          }
        });
      });
    },

    async previewNormalizeAndFormatFromEditor() {
      try {
        if (!EditorInterface.isEditorOpen()) {
          return {
            ok: false,
            info: "No note editor open. Click a note (or Add Note) first, or use 'Open latest + Preview'.",
          };
        }

        const beforeText = EditorInterface.getEditorRawText();
        const { text: normalizedText, changeList } = TermNormalizer.normalize(beforeText);
        const formattedHtml = TextFormatter.buildFormattedHtmlFromText(normalizedText);

        return await this._showPreviewAndApply({ beforeText, normalizedText, formattedHtml, changeList });
      } catch (e) {
        console.error("[PatchNoteFormatter] Preview error:", e);
        return { ok: false, info: e.message || "Error creating preview" };
      }
    },

    async openLatestAndPreviewNormalizeAndFormat() {
      try {
        const res = await this.openLatestCallNoteEditor();
        if (!res.ok) return res;
        // Editor is already confirmed open by waitFor inside openLatestCallNoteEditor.
        return await this.previewNormalizeAndFormatFromEditor();
      } catch (e) {
        console.error("[PatchNoteFormatter] Open latest error:", e);
        return { ok: false, info: e.message || "Error opening latest note" };
      }
    },

    async openLatestCallNoteEditor() {
      try {
        const callTab = document.querySelector('a[href$="#calls"], a[href$="#Calls"], a[href$="#CALLS"]');
        if (callTab) callTab.click();

        // Wait for the View activity button to appear after switching tabs.
        const viewActivityBtn = await Utils.waitFor(() =>
          ButtonHelpers.findByText((t) => t === "view activity")
        );
        if (viewActivityBtn) viewActivityBtn.click();

        // Try to open an existing note first.
        const noteBtn = await Utils.waitFor(() =>
          ButtonHelpers.findByText(
            (t) => t.includes("note") && t !== "add note" && t !== "note details"
          ),
          1500
        );

        if (noteBtn) {
          noteBtn.click();
          await Utils.waitFor(() => EditorInterface.isEditorOpen(), 1500);
        }

        if (!EditorInterface.isEditorOpen()) {
          const addNoteBtn = await Utils.waitFor(() =>
            ButtonHelpers.findByText((t) => t === "add note"),
            1500
          );
          if (addNoteBtn) {
            addNoteBtn.click();
            await Utils.waitFor(() => EditorInterface.isEditorOpen(), 2000);
          }
        }

        if (!EditorInterface.isEditorOpen()) {
          return {
            ok: false,
            info: "Could not open the latest note editor. Try opening a note first, then run preview.",
          };
        }

        return { ok: true, info: "Opened editor." };
      } catch (e) {
        console.error("[PatchNoteFormatter] Open editor error:", e);
        return { ok: false, info: e.message || "Error opening editor" };
      }
    },

    async clipboardToEditorPreview() {
      try {
        if (!EditorInterface.isEditorOpen()) {
          return {
            ok: false,
            info: "No note editor open. Click a note (or Add Note) first, or use 'Open latest + Preview'.",
          };
        }

        if (!navigator.clipboard?.readText) {
          return { ok: false, info: "Clipboard readText not available in this browser/context." };
        }

        let beforeText = "";
        try {
          beforeText = (await navigator.clipboard.readText()) || "";
        } catch (e) {
          return { ok: false, info: "Clipboard access denied or failed." };
        }

        if (!beforeText.trim()) return { ok: false, info: "Clipboard was empty." };

        const { text: normalizedText, changeList } = TermNormalizer.normalize(beforeText);
        const formattedHtml = TextFormatter.buildFormattedHtmlFromText(normalizedText);

        return await this._showPreviewAndApply({ beforeText, normalizedText, formattedHtml, changeList });
      } catch (e) {
        console.error("[PatchNoteFormatter] Clipboard preview error:", e);
        return { ok: false, info: e.message || "Error processing clipboard" };
      }
    },
  };

  // =============================
  // MAIN UI
  // =============================
  const MainUI = {
    remove() {
      const ui = document.getElementById("patchNoteFormatterUI");
      if (ui) ui.remove();
      const modal = document.getElementById("patchPreviewModal");
      if (modal) modal.remove();
      const terms = document.getElementById("patchTermsModal");
      if (terms) terms.remove();
    },

    create() {
      const existing = document.getElementById("patchNoteFormatterUI");
      if (existing) return;

      const ui = document.createElement("div");
      ui.id = "patchNoteFormatterUI";

      const title = this._buildTitle();
      const body = this._buildBody();

      ui.appendChild(title.container);
      ui.appendChild(body.container);
      document.body.appendChild(ui);

      this._setupPositioning(ui);
      this._setupMinimize(ui, body.container, title.btnMin);
      this._setupDragging(ui, title.container, title.btnMin, title.btnClose);
      this._setupEventHandlers(body);
    },

    _buildTitle() {
      const container = document.createElement("div");
      container.className = "pnf-ui-title";

      const titleText = document.createElement("div");
      titleText.className = "pnf-ui-title-text";
      titleText.textContent = "Patch Note Formatter";
      container.appendChild(titleText);

      const titleBtns = document.createElement("div");
      titleBtns.className = "pnf-ui-title-btns";

      const btnMin = UIComponents.createButton({
        text: "–",
        variant: "pine",
        extraClass: "pnf-btn-icon",
        ariaLabel: "Minimize",
      });

      const btnClose = UIComponents.createButton({
        text: "×",
        variant: "cranberry",
        extraClass: "pnf-btn-icon",
        ariaLabel: "Close",
      });

      titleBtns.appendChild(btnMin);
      titleBtns.appendChild(btnClose);
      container.appendChild(titleBtns);

      return { container, btnMin, btnClose };
    },

    _buildBody() {
      const container = document.createElement("div");
      container.className = "pnf-ui-body";

      const row1 = document.createElement("div");
      row1.className = "pnf-row";

      const btnPreview = UIComponents.createButton({
        text: "Preview Normalize + Format",
        variant: "clover",
        extraClass: "pnf-btn-flex",
      });

      const btnClipboard = UIComponents.createButton({
        text: "Clipboard → Preview",
        variant: "pine",
        extraClass: "pnf-btn-flex",
      });

      row1.appendChild(btnPreview);
      row1.appendChild(btnClipboard);

      const row2 = document.createElement("div");
      row2.className = "pnf-row";

      const btnOpenLatest = UIComponents.createButton({
        text: "Open latest + Preview",
        variant: "coconut",
        extraClass: "pnf-btn-flex",
      });

      const btnManageTerms = UIComponents.createButton({
        text: "Manage terms",
        variant: "pine",
        extraClass: "pnf-btn-flex",
        ariaLabel: "Edit user-defined term overrides",
      });

      row2.appendChild(btnOpenLatest);
      row2.appendChild(btnManageTerms);

      const row3 = document.createElement("div");
      row3.className = "pnf-row-space";

      const chk = document.createElement("input");
      chk.type = "checkbox";
      chk.checked = AppState.autoSave;
      chk.id = "patchFormatterAutoSave";

      const chkLbl = document.createElement("label");
      chkLbl.htmlFor = "patchFormatterAutoSave";
      chkLbl.className = "pnf-check-label";
      chkLbl.appendChild(chk);
      chkLbl.appendChild(document.createTextNode("Auto-save after apply"));

      const hint = document.createElement("div");
      hint.className = "pnf-hint";
      hint.textContent = "Hotkey: Alt+Shift+F";

      row3.appendChild(chkLbl);
      row3.appendChild(hint);

      const status = document.createElement("div");
      status.className = "pnf-status";

      container.appendChild(row1);
      container.appendChild(row2);
      container.appendChild(row3);
      container.appendChild(status);

      return {
        container,
        btnPreview,
        btnClipboard,
        btnOpenLatest,
        btnManageTerms,
        chk,
        status,
      };
    },

    _setupPositioning(ui) {
      const savedPos = Storage.get(CONFIG.UI_POS_KEY);
      if (savedPos?.left && savedPos?.top) {
        Object.assign(ui.style, {
          left: `${savedPos.left}px`,
          top: `${savedPos.top}px`,
          right: "auto",
          bottom: "auto",
        });
      }
    },

    _setupMinimize(ui, body, btnMin) {
      let minimized = Storage.get(CONFIG.UI_MIN_KEY);
      if (minimized === null) minimized = CONFIG.DEFAULT_START_MINIMIZED;

      const setMin = (min) => {
        minimized = min;
        Storage.set(CONFIG.UI_MIN_KEY, minimized);
        body.style.display = minimized ? "none" : "flex";
        btnMin.textContent = minimized ? "+" : "–";
      };

      setMin(minimized);
      btnMin.onclick = () => setMin(!minimized);
    },

    _setupDragging(ui, titleContainer, btnMin, btnClose) {
      let drag = null;

      titleContainer.addEventListener("mousedown", (e) => {
        if (e.target === btnMin || e.target === btnClose) return;
        const rect = ui.getBoundingClientRect();
        drag = {
          startX: e.clientX,
          startY: e.clientY,
          startLeft: rect.left,
          startTop: rect.top,
        };
        e.preventDefault();
      });

      window.addEventListener("mousemove", (e) => {
        if (!drag) return;
        const dx = e.clientX - drag.startX;
        const dy = e.clientY - drag.startY;
        const left = drag.startLeft + dx;
        const top = drag.startTop + dy;
        Object.assign(ui.style, {
          left: `${left}px`,
          top: `${top}px`,
          right: "auto",
          bottom: "auto",
        });
      });

      window.addEventListener("mouseup", () => {
        if (!drag) return;
        const rect = ui.getBoundingClientRect();
        Storage.set(CONFIG.UI_POS_KEY, { left: rect.left, top: rect.top });
        drag = null;
      });

      btnClose.onclick = () => this.remove();
    },

    _setupEventHandlers(body) {
      const showStatus = (text) => {
        body.status.textContent = text;
        setTimeout(() => (body.status.textContent = ""), 2200);
      };

      body.chk.onchange = () => {
        AppState.autoSave = body.chk.checked;
        showStatus(`Auto-save: ${AppState.autoSave ? "ON" : "OFF"}`);
      };

      body.btnPreview.onclick = async () => {
        try {
          body.status.textContent = "Working…";
          const res = await Actions.previewNormalizeAndFormatFromEditor();
          showStatus(res.info);
        } catch (e) {
          console.error("[PatchNoteFormatter] Preview error:", e);
          showStatus("Error occurred");
        }
      };

      body.btnClipboard.onclick = async () => {
        try {
          body.status.textContent = "Reading clipboard…";
          const res = await Actions.clipboardToEditorPreview();
          showStatus(res.info);
        } catch (e) {
          console.error("[PatchNoteFormatter] Clipboard error:", e);
          showStatus("Error occurred");
        }
      };

      body.btnOpenLatest.onclick = async () => {
        try {
          body.status.textContent = "Opening latest…";
          const res = await Actions.openLatestAndPreviewNormalizeAndFormat();
          showStatus(res.info);
        } catch (e) {
          console.error("[PatchNoteFormatter] Open latest error:", e);
          showStatus("Error occurred");
        }
      };

      body.btnManageTerms.onclick = () => {
        try {
          ManageTermsModal.open();
        } catch (e) {
          console.error("[PatchNoteFormatter] Manage terms error:", e);
          showStatus("Error opening terms editor");
        }
      };
    },
  };

  // =============================
  // HOTKEY HANDLER
  // =============================
  window.addEventListener("keydown", async (e) => {
    if (
      e.altKey === CONFIG.HOTKEY.altKey &&
      e.shiftKey === CONFIG.HOTKEY.shiftKey &&
      e.code === CONFIG.HOTKEY.code
    ) {
      e.preventDefault();
      if (!Router.isOnCallsView()) return;

      try {
        MainUI.create();
        const res = await Actions.previewNormalizeAndFormatFromEditor();
        const ui = document.getElementById("patchNoteFormatterUI");
        const status = ui?.querySelector(".pnf-status");
        if (status) {
          status.textContent = res.info;
          setTimeout(() => (status.textContent = ""), 2200);
        }
      } catch (err) {
        console.error("[PatchNoteFormatter] Hotkey error:", err);
      }
    }
  });

  // =============================
  // ROUTING & INITIALIZATION
  // =============================
  function handleRouteChange() {
    if (Router.isOnCallsView()) {
      StyleInjector.inject();
      MainUI.create();
      console.log("[PatchNoteFormatter] Initialized on Calls view.");
    } else {
      MainUI.remove();
    }
  }

  // Inject styles as early as possible so any modal created programmatically
  // (e.g. via hotkey on first route hit) is styled.
  StyleInjector.inject();

  // Patch history API so SPA navigations trigger our route check immediately
  // (no polling needed).
  (function patchHistory() {
    const origPush = history.pushState;
    history.pushState = function (...args) {
      const ret = origPush.apply(this, args);
      handleRouteChange();
      return ret;
    };
    const origReplace = history.replaceState;
    history.replaceState = function (...args) {
      const ret = origReplace.apply(this, args);
      handleRouteChange();
      return ret;
    };
  })();

  window.addEventListener("popstate", handleRouteChange);
  window.addEventListener("hashchange", handleRouteChange);

  handleRouteChange();

  console.log("[PatchNoteFormatter v1.8.0] Loaded successfully.");
})();
