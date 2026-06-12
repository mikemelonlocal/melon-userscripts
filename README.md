# Melon Local Userscripts

A collection of Tampermonkey scripts that add productivity tools to Patch, Yext, and Microsoft Ads. Each script adds buttons, panels, or shortcuts to existing pages — nothing here replaces the underlying apps; they make day-to-day work faster.

## Getting set up (one time)

1. **Install Tampermonkey** in your browser:
   - [Chrome / Edge](https://chromewebstore.google.com/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo)
   - [Firefox](https://addons.mozilla.org/firefox/addon/tampermonkey/)
   - [Safari](https://apps.apple.com/app/tampermonkey/id1482490089)
2. Open this repo on GitHub: <https://github.com/mikemelonlocal/melon-userscripts>
3. Click any `.user.js` file, then click the **Raw** button. Tampermonkey will detect it and prompt to install. Click **Install**.
4. Repeat for the scripts you want. Auto-updates are on by default — Tampermonkey will pull new versions in the background.

To uninstall: open the Tampermonkey dashboard (browser toolbar icon → Dashboard) and toggle or delete a script.

---

## Table of contents

**Patch — notes & calls**
- [MelonPatch - Note Tools (Enhanced)](#melonpatch---note-tools-enhanced)
- [Patch Note Formatter](#patch-note-formatter)
- [Patch - Paste Next Gong Note](#patch---paste-next-gong-note)
- [Patch Gong Workflow Helper](#patch-gong-workflow-helper)
- [Patch Monthly Auto-Creator](#patch-monthly-auto-creator)
- [Patch Monthly Notes Gaps](#patch-monthly-notes-gaps)

**Patch — tasks**
- [MelonPatch Smart Archive](#melonpatch-smart-archive)

**Patch — campaigns & targeting**
- [Daily Cap Calculator](#daily-cap-calculator)
- [Bulk Campaign Patch Status](#bulk-campaign-patch-status)
- [Patch Targeting Helper](#patch-targeting-helper)
- [Searchable & Sortable Budget Tables](#melonpatch---searchable--sortable-budget-tables)

**Patch — admin & navigation**
- [Remove Ex-Employees from System Groups](#remove-ex-employees-from-system-groups)
- [CIB Report Quick Filters](#cib-report-quick-filters)
- [Search Agent Shortcut](#search-agent-shortcut)

**Yext**
- [Yext Media Autofill](#yext-media-autofill)
- [Yext Photo Text Generator](#yext-photo-text-generator)

**Microsoft Ads**
- [Microsoft Ads - Dismiss All Recommendations](#microsoft-ads---dismiss-all-recommendations)

---

## MelonPatch - Note Tools (Enhanced)

**What it does:** Adds quick-action buttons (Copy, Edit, Duplicate, Copy to Tasks, Delete) and saved Note Templates to every note on Patch. Also fixes the stale comment-count bubbles that don't update when replies come in.

**Where it shows up:** Anywhere notes are displayed on `thepatch.melonlocal.com` — grid views, tree views, agent dashboards.

**How to use:**
- Hover or click on any note to see the new action buttons.
- **Edit** lets you change a note in place. **Duplicate** copies the full note as a new entry. **Copy to Tasks** pushes the note text into a new task.
- For templates: click the templates icon to open the panel, paste in your boilerplate, give it a name, and reuse from any note editor.
- Comment-count bubbles now refresh automatically as people reply.

---

## Patch Note Formatter

**What it does:** Auto-formats notes with bold headings and proper bullet lists, and normalizes common transcription variants to the canonical terms your team uses ("Melon Local" instead of "melon-local", etc.).

**Where it shows up:** Any note editor across `thepatch.melonlocal.com`. A dock panel appears in the sidebar.

**How to use:**
- Type your note normally (or paste in a Gong transcript).
- Open the dock panel and click **Format** to apply heading + list styling.
- The **Auditor view** shows you a before/after diff with the changed terms highlighted so you can confirm nothing important got rewritten.
- Manage your custom term rules in the **Simple-view terms editor** inside the dock — no script editing required.

---

## Patch - Paste Next Gong Note

**What it does:** One-click paste of the next unprocessed Gong call note from a shared Google Sheet directly into the Patch note editor. Skips notes already pasted.

**Where it shows up:** Patch note editor. After you click **Add Note**, a **Paste Next Gong Note** button appears next to it.

**How to use:**
- Click **Add Note** as usual.
- Click **Paste Next Gong Note**. The next un-pasted row from the source Google Sheet is dropped into the editor.
- Edit/save as you would any note. The script remembers which rows have been used so you don't repeat.

**Setup:** The source sheet URL is currently hardcoded. If the sheet moves or you need to swap it, edit `patch-paste-next-gong-note.user.js` line 18.

---

## Patch Gong Workflow Helper

**What it does:** Drives Patch through a full Gong-call follow-up workflow using a Google Sheet as the work queue. Resumes across page navigations so you can move agent-by-agent without losing your place.

**Where it shows up:** A floating panel on `thepatch.melonlocal.com` with **Start**, **Abort**, **Hard Reset**, **Config**, and a live log.

**How to use:**
- Click **Config…** to set the Google Sheet ID and (optionally) a webhook URL for completion notifications.
- Click **Start** to begin processing the queue. The script navigates between agents, opens the right tabs, and pastes notes from the sheet.
- **Abort** stops at the next safe checkpoint. **Hard Reset** clears progress and reloads the sheet from scratch.
- The log panel shows what the script is doing in real time — no alert spam.

**Setup:** Run **Config…** the first time you use it to point at your team's Gong workflow sheet.

---

## Patch Monthly Auto-Creator

**What it does:** Creates next month's Monthly call on the current agent, copying details from the most recent Monthly so you don't have to retype them.

**Where it shows up:** Agent Dashboard → Calls tab. Adds a button to the Calls grid.

**How to use:**
- Open any agent's Dashboard and click into the **Calls** tab.
- Click the **Create Next Monthly** button.
- A confirmation dialog shows what will be created. Click OK; the new call is drafted from the most recent Monthly's details.

---

## Patch Monthly Notes Gaps

**What it does:** Scans the agent's Calls grid and tells you which months are missing a final-status call (Completed, Canceled, or Agent No Show) for a given year. Optionally bulk-archives an entire year's calls by title.

**Where it shows up:** Agent Dashboard → Calls tab only. Adds a panel near the grid.

**How to use:**
- Open an agent's Calls tab.
- The panel lists months with missing final-status calls, color-coded.
- To clean up a past year, use **Bulk Archive by Year**. A confirmation dialog shows the exact count + month breakdown before any archive happens.

---

## MelonPatch Smart Archive

**What it does:** One-click bulk archive of Done and/or Unsuccessful tasks, instead of archiving them one at a time.

**Where it shows up:** Tasks pages (`Tasks/MyTasks`, `Tasks/TaskSearch`, and agent dashboards). Adds a **Smart Archive** button next to the standard Archive button.

**How to use:**
- Click **Smart Archive**.
- Choose **Done**, **Unsuccessful**, or **Both**.
- **Hold the confirm button for 1.5 seconds** to commit — this prevents accidental archives. Release before then to cancel.

---

## Daily Cap Calculator

**What it does:** A floating panel that paces ad budgets evenly through the rest of the month. Auto-pulls product/budget data from the page, lets you freeze line items, and saves your work across sessions.

**Where it shows up:** Floating panel on `thepatch.melonlocal.com`. Drag the panel anywhere on screen; minimize when not in use.

**How to use:**
- The panel auto-fills from page data when possible. Adjust per-product budgets and freeze any rows that shouldn't change.
- Click **Calculate** (or press **Ctrl/Cmd + Enter**) to run the pacing math.
- **Export / Import** lets you share a snapshot with a teammate.

**Keyboard shortcuts:**
- `Ctrl/Cmd + Enter` — Calculate
- `Ctrl/Cmd + K` — Add product row
- `Ctrl/Cmd + M` — Minimize panel
- `Ctrl/Cmd + E` — Export

---

## Bulk Campaign Patch Status

**What it does:** Bulk Active/Inactive toggling for campaign patch statuses, with an undo buffer in case you click the wrong thing.

**Where it shows up:** `Agents/BudgetDetails` pages. Adds a sticky toolbar with **Action FAB** (floating action button) and per-row pill counts.

**How to use:**
- Select the rows you want to change using the toolbar's checkboxes.
- Click the action button and choose Active or Inactive.
- A spinner shows progress per row, then verification confirms each change actually committed (not just queued).
- If something looks off, the **Undo** button rolls back the last batch.

---

## Patch Targeting Helper

**What it does:** Bulk Add / Remove / Move tools for advertising targets (City, County, Zip), plus an optional coverage analysis that tells you which ZIPs are missing from your City/County/DMA/State targeting.

**Where it shows up:**
- **Edit Advertising Targets** page: bulk Add/Remove panels for County, City, and Zip.
- **BudgetDetails** screens with Kendo list boxes: Bulk Move panel.

**How to use:**
- Paste a list of geos (one per line) into the relevant panel.
- The script validates each entry, shows live progress, and retries failed API calls.
- For coverage analysis: click **Analyze Zip Coverage** to see which ZIPs your current targeting misses.
- Bulk Remove requires a confirmation dialog showing the exact list before deletion.

---

## MelonPatch - Searchable & Sortable Budget Tables

**What it does:** Turns the Legacy Office budget tables into DataTables — adds a live search box, per-column sorting, paging, and checkbox filter panes (SearchPanes), so you can find or narrow line items without scrolling the whole sheet.

**Where it shows up:** Agent Dashboard (`Agents/Dashboard/*`). Applies to every `.budget-details-table` across all three tab panels.

**How to use:**
- Open an agent's Dashboard. Each budget table gains a **Search** box and a **Show N entries** selector.
- Click any column header to sort by it. The action-button column (last column) stays unsortable.
- **Checkbox filters:** categorical columns (Budget Type, Platform, etc.) get checkbox filter panes above the table — check values to narrow rows, and the panes cascade so each selection refines the others. Columns with mostly-unique values (Patch ID, descriptions, dates) are skipped automatically.
- Tables load in their original row order; sorting/searching/filtering is opt-in per table.

---

## Remove Ex-Employees from System Groups

**What it does:** Highlights ex-employees on System Group and Teams Detail pages, with inline panels to bulk-remove them and bulk-add active Melons in their place.

**Where it shows up:** `SystemGroups/Edit` and `Teams/Details` pages.

**How to use:**
- Open any System Group or Team Details page. Ex-employees are highlighted automatically.
- Use the **Bulk Remove Ex-Employees** panel to clear them all in one click (with confirmation).
- Use the **Bulk Add Active Melons** panel to paste in names and add them en masse.

---

## CIB Report Quick Filters

**What it does:** Adds a quick-filter panel to the CIB Report so you can filter multiple columns at once without dropping into Kendo's per-column filter menus.

**Where it shows up:** `Reports/CIB`. Adds a filter panel above the grid.

**How to use:**
- Type or pick values in the filter chips — multiple columns can be filtered simultaneously.
- Save commonly-used filter combos as **Named Presets** so you can re-apply them with one click.
- Presets are saved per browser, so each team member builds their own set.

---

## Search Agent Shortcut

**What it does:** Adds a command-palette style search for finding agents fast, wrapping the existing Search Agent autocomplete in a clean modal.

**Where it shows up:** Anywhere on `thepatch.melonlocal.com`. Triggered by keyboard shortcut.

**How to use:**
- Press **Cmd + B** (Mac) or **Ctrl + B** (Windows/Linux) from any Patch page.
- A modal opens with the search field focused. Type to filter, click or press Enter to navigate.
- Press **Esc** or the same shortcut again to close.

---

## Yext Media Autofill

**What it does:** Generates SEO-focused Description, Details, and Alt Text for every photo on a Yext entity, using Claude or Gemini vision AI. Processes media in parallel and respects Yext's field-length constraints.

**Where it shows up:** Yext entity editor (`www.yext.com/s/*/entity/edit3*`). Adds a panel near the media section with **Suggest** and **Apply** controls.

**How to use:**
- Open any entity's media editor.
- Click **Suggest** to generate text for all visible media items. Review the results in the panel.
- Click **Apply** to fill the fields, or accept/reject suggestions one at a time.

**Setup (one time):** API keys must be set via the browser console — they're never embedded in the script. Open DevTools → Console on any Yext page and run:
```js
GM_setValue("anthropicApiKey", "sk-ant-...")  // for Claude
GM_setValue("geminiApiKey", "AIza...")         // for Gemini
```
You only need the provider you plan to use. The script defaults to Claude; switch via the `provider` field in the script's `CONFIG` block (`"claude"`, `"gemini"`, or `"hybrid"`).

---

## Yext Photo Text Generator

**What it does:** A lightweight, no-AI version of Yext Media Autofill. Generates Description/Details/Alt text from the entity's existing page context (name, category, address fields) using templates — no API keys required, no calls to outside services.

**Where it shows up:** Yext entity editor, same pages as the Autofill script.

**How to use:**
- Open an entity's media editor.
- Click the generator button to fill in templated text for each photo.
- Best for entities where you want consistent, fast metadata without using AI credits.

**When to use this vs. Yext Media Autofill:** Use this for routine fill where speed and consistency matter. Use the AI Autofill when you want unique, photo-specific descriptions.

---

## Microsoft Ads - Dismiss All Recommendations

**What it does:** Adds a **Dismiss All** button to the Microsoft Ads recommendations page, with a progress bar and Fluent UI toasts. Protects against accidentally applying recommendations instead of dismissing them.

**Where it shows up:** `ui.ads.microsoft.com` recommendations page.

**How to use:**
- Click **Dismiss All**. A progress bar shows how many cards are being processed.
- The script *never* clicks **Apply** or **Yes** buttons — only **Dismiss** — so it can't accidentally accept recommendations.
- Negative-keyword recommendations get a confirmation prompt before dismiss, since those are sometimes worth reviewing individually.

---

## Reporting issues / requesting features

File an issue at <https://github.com/mikemelonlocal/melon-userscripts/issues>, or message Mike directly.
