# Melon Local Userscripts — Cheat Sheet

One-line reference for each script. Full how-to and install instructions: [README](README.md).

## Patch — notes & calls

- **Note Tools (Enhanced)** — Copy / Edit / Duplicate / Copy-to-Tasks / Delete buttons on every note, plus saved Note Templates. Fixes stale comment-count bubbles.
- **Note Formatter** — Auto-bolds headings, preserves lists, normalizes term variants. Click **Format** in the sidebar dock.
- **Paste Next Gong Note** — After **Add Note**, click **Paste Next Gong Note** to drop in the next un-pasted row from the Gong sheet.
- **Gong Workflow Helper** — Drives the full Gong follow-up workflow off a Google Sheet. **Config…** → **Start**. Survives page navigation.
- **Monthly Auto-Creator** — On an agent's Calls tab, one click creates next month's Monthly call from the most recent one.
- **Monthly Notes Gaps** — On an agent's Calls tab, shows which months are missing a final-status call for a year. Bulk-archive by year available.

## Patch — tasks

- **Smart Archive** — One click to bulk-archive Done / Unsuccessful / Both. Hold the confirm button 1.5 seconds to commit.

## Patch — campaigns & targeting

- **Daily Cap Calculator** — Paces budgets through end of month. Shortcuts: `⌘/Ctrl + Enter` calculate · `⌘/Ctrl + K` add product · `⌘/Ctrl + M` minimize · `⌘/Ctrl + E` export.
- **Bulk Campaign Patch Status** — On BudgetDetails, sticky toolbar to bulk Active/Inactive with undo buffer and per-row verification.
- **Targeting Helper** — Bulk Add/Remove on Edit Advertising Targets (City/County/Zip) + Bulk Move on BudgetDetails Kendo list boxes. Optional zip coverage analysis.

## Patch — admin & navigation

- **Remove Ex-Employees from System Groups** — Highlights ex-employees on System Groups + Teams pages; bulk-remove and bulk-add panels.
- **CIB Report Quick Filters** — Multi-column quick-filter chips on the CIB Report grid. Save Named Presets.
- **Search Agent Shortcut** — `⌘ + B` (Mac) / `Ctrl + B` (Win) opens a command-palette agent search from anywhere on Patch.

## Yext

- **Media Autofill** — Generates Description / Details / Alt Text for every photo via Claude or Gemini. Requires API key set once via `GM_setValue` (see README).
- **Photo Text Generator** — Lightweight no-AI version of Media Autofill. Uses page context + templates. No API key needed.

## Microsoft Ads

- **Dismiss All Recommendations** — On the recommendations page, one click dismisses all (never applies). Negative-keyword cards prompt before dismiss.

---

*Install: see <https://github.com/mikemelonlocal/melon-userscripts> → install Tampermonkey, click any `.user.js` Raw button.*
