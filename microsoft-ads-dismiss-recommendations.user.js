// ==UserScript==
// @name         Microsoft Ads - Dismiss All Recommendations
// @namespace    http://tampermonkey.net/
// @version      2.1
// @description  Adds a "Dismiss All" button to the Microsoft Advertising recommendations page, with progress bar, negative-keyword guard, and Fluent UI toast feedback.
// @author       You
// @match        https://ui.ads.microsoft.com/*
// @run-at       document-start
// @grant        none
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/microsoft-ads-dismiss-recommendations.user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/microsoft-ads-dismiss-recommendations.user.js
// ==/UserScript==

(function () {
  'use strict';

  console.info('[DismissAll] userscript loaded (v2.1)');

  // ---------------------------------------------------------------------------
  // Constants
  // ---------------------------------------------------------------------------

  // Microsoft A/B-tests button labels. Anywhere we identify a control by text,
  // accept any of these synonyms (compared case-insensitively, EXACT match).
  // IMPORTANT: do NOT include 'Apply' in SUBMIT — Microsoft Ads recommendation
  // cards have "Apply all" buttons that do the opposite of dismissing. Likewise
  // 'Yes' alone is too ambiguous (used by other confirmation dialogs).
  const TEXTS = {
    DISMISS_MENU_ITEM: ['Dismiss all', 'Dismiss', 'Dismiss recommendation', 'Discard', 'Discard all'],
    SUBMIT:            ['Submit', 'Confirm', 'OK', 'Done', 'Save'],
    CLOSE:             ['Close', 'Cancel'],
  };

  // CSS-in-JS-resistant overrides — Microsoft's stylesheets are aggressive.
  const HIGHLIGHT_OUTLINE = '4px solid #d13438';

  // ---------------------------------------------------------------------------
  // Module state
  // ---------------------------------------------------------------------------

  let dismissObserver = null;
  let isRunning = false;
  let attachScheduled = false;

  // ---------------------------------------------------------------------------
  // Async primitives
  // ---------------------------------------------------------------------------

  /**
   * Poll for an element matching `selector` (and optional `filterFn`) until it
   * appears and is visible (`offsetParent !== null`). Returns null on timeout.
   */
  async function waitFor(selector, filterFn, timeout = 3000) {
    const start = performance.now();
    while (performance.now() - start < timeout) {
      const els = Array.from(document.querySelectorAll(selector));
      const match = filterFn ? els.find(filterFn) : els[0];
      if (match && match.offsetParent !== null) return match;
      await sleep(20);
    }
    return null;
  }

  /**
   * Same as waitFor, but matches against any of `texts` (case-insensitive,
   * EXACT match after trim). This is the resilience layer for Microsoft's UI
   * A/B variants. `extraFilter` lets callers add constraints like `!disabled`.
   *
   * Exact match (no prefix matching) is intentional: prefix matching let
   * "Apply" match "Apply all" buttons elsewhere on the page, triggering
   * the wrong action.
   */
  async function waitForByText(selector, texts, timeout = 3000, extraFilter = null) {
    const lowered = texts.map(t => t.toLowerCase());
    return waitFor(selector, el => {
      if (extraFilter && !extraFilter(el)) return false;
      const t = (el.textContent || '').trim().toLowerCase();
      return lowered.includes(t);
    }, timeout);
  }

  async function waitForCondition(condition, timeout = 5000) {
    const start = performance.now();
    while (performance.now() - start < timeout) {
      if (condition()) return true;
      await sleep(20);
    }
    return false;
  }

  function sleep(ms) {
    return new Promise(r => setTimeout(r, ms));
  }

  function pressEscape() {
    document.dispatchEvent(new KeyboardEvent('keydown', {
      key: 'Escape', code: 'Escape', keyCode: 27, which: 27, bubbles: true
    }));
  }

  // ---------------------------------------------------------------------------
  // Card identification
  // ---------------------------------------------------------------------------

  function getCardElement(moreBtn) {
    // Prefer explicit class hints first.
    const byClass = moreBtn.closest(
      '[class*="recommendation" i], [class*="card" i], [class*="tile" i], ' +
      '[class*="blade" i], [class*="insight" i], [data-testid*="rec" i], ' +
      'li, article, section'
    );
    if (byClass && (byClass.textContent || '').trim().length > 40) return byClass;
    // Fallback: walk up until we find an ancestor with substantial text.
    let el = moreBtn.parentElement;
    for (let i = 0; i < 15 && el; i++, el = el.parentElement) {
      const text = (el.textContent || '').trim();
      if (text.length > 80) return el;
    }
    return byClass || moreBtn.parentElement;
  }

  function isNegativeKeywordCard(moreBtn) {
    const card = getCardElement(moreBtn);
    if (!card) return false;
    const text = (card.textContent || '').toLowerCase();
    const hit = /negative\s+(keyword|kw)/.test(text);
    console.debug(
      '[DismissAll] card check:',
      hit ? 'NEGATIVE KW' : 'ok',
      '—',
      text.slice(0, 120).replace(/\s+/g, ' ')
    );
    return hit;
  }

  // ---------------------------------------------------------------------------
  // Visual: progress bar (3px, top of viewport)
  // ---------------------------------------------------------------------------

  function ensureProgressBar() {
    let bar = document.getElementById('tamper-dismiss-progress');
    if (bar) return bar;
    bar = document.createElement('div');
    bar.id = 'tamper-dismiss-progress';
    bar.style.cssText =
      'position:fixed !important;top:0 !important;left:0 !important;' +
      'height:3px !important;width:0% !important;' +
      'background:linear-gradient(90deg,#0078d4,#40e0d0) !important;' +
      'box-shadow:0 0 6px rgba(0,120,212,0.6) !important;' +
      'z-index:2147483647 !important;pointer-events:none !important;' +
      'transition:width 0.25s ease, opacity 0.4s ease !important;' +
      'opacity:1 !important;';
    document.documentElement.appendChild(bar);
    return bar;
  }

  function updateProgress(processed, total) {
    const bar = ensureProgressBar();
    const pct = total > 0 ? Math.min(100, (processed / total) * 100) : 0;
    bar.style.setProperty('width', pct + '%', 'important');
    bar.style.setProperty('opacity', '1', 'important');
  }

  function hideProgressBar() {
    const bar = document.getElementById('tamper-dismiss-progress');
    if (!bar) return;
    bar.style.setProperty('opacity', '0', 'important');
    setTimeout(() => bar.remove(), 500);
  }

  // ---------------------------------------------------------------------------
  // Visual: card highlight (interlock during askUser)
  // ---------------------------------------------------------------------------

  /**
   * Outlines the card belonging to `moreBtn` in red and scrolls it into view.
   * Returns a function that restores the card's prior styling.
   */
  function highlightCard(moreBtn) {
    const card = getCardElement(moreBtn);
    if (!card) return () => {};
    const prevOutline = card.style.outline;
    const prevOffset = card.style.outlineOffset;
    const prevTransition = card.style.transition;
    card.style.setProperty('outline', HIGHLIGHT_OUTLINE, 'important');
    card.style.setProperty('outline-offset', '2px', 'important');
    card.style.setProperty('transition', 'outline 0.2s ease', 'important');
    try {
      card.scrollIntoView({ behavior: 'smooth', block: 'center' });
    } catch (_) { /* ignore */ }
    return () => {
      card.style.outline = prevOutline;
      card.style.outlineOffset = prevOffset;
      card.style.transition = prevTransition;
    };
  }

  // ---------------------------------------------------------------------------
  // Visual: Fluent UI–style toast
  // ---------------------------------------------------------------------------

  const TOAST_PALETTE = {
    success: { stripe: '#107c10', icon: '✓' },
    info:    { stripe: '#0078d4', icon: 'i' },
    warning: { stripe: '#797775', icon: '!' },
    error:   { stripe: '#a4262c', icon: '✕' },
  };

  function showToast(message, { kind = 'success', duration = 6000 } = {}) {
    const palette = TOAST_PALETTE[kind] || TOAST_PALETTE.info;

    const toast = document.createElement('div');
    toast.setAttribute('role', 'status');
    toast.style.cssText =
      'position:fixed !important;top:24px !important;right:24px !important;' +
      'min-width:280px;max-width:420px;padding:12px 14px !important;' +
      'background:#ffffff !important;border-radius:2px !important;' +
      'border-left:4px solid ' + palette.stripe + ' !important;' +
      'box-shadow:0 6.4px 14.4px rgba(0,0,0,0.18),0 1.2px 3.6px rgba(0,0,0,0.11) !important;' +
      'display:flex !important;align-items:flex-start !important;gap:10px !important;' +
      'font-family:"Segoe UI",-apple-system,BlinkMacSystemFont,sans-serif !important;' +
      'font-size:13px !important;line-height:1.4 !important;color:#323130 !important;' +
      'z-index:2147483647 !important;' +
      'transform:translateY(-12px) !important;opacity:0 !important;' +
      'transition:opacity 0.25s ease, transform 0.25s ease !important;';

    const icon = document.createElement('div');
    icon.textContent = palette.icon;
    icon.style.cssText =
      'flex:0 0 20px;width:20px;height:20px;border-radius:50%;' +
      'background:' + palette.stripe + ';color:#fff;font-weight:700;' +
      'display:flex;align-items:center;justify-content:center;font-size:12px;' +
      'margin-top:1px;';

    const body = document.createElement('div');
    body.style.cssText = 'flex:1;white-space:pre-line;';
    body.textContent = message;

    const close = document.createElement('button');
    close.textContent = '✕';
    close.setAttribute('aria-label', 'Close');
    close.style.cssText =
      'flex:0 0 auto;background:transparent;border:none;cursor:pointer;' +
      'color:#605e5c;font-size:14px;padding:0 2px;line-height:1;';

    const dismiss = () => {
      toast.style.setProperty('opacity', '0', 'important');
      toast.style.setProperty('transform', 'translateY(-12px)', 'important');
      setTimeout(() => toast.remove(), 260);
    };
    close.addEventListener('click', dismiss);

    toast.appendChild(icon);
    toast.appendChild(body);
    toast.appendChild(close);
    document.documentElement.appendChild(toast);

    // Slide-in
    requestAnimationFrame(() => {
      toast.style.setProperty('opacity', '1', 'important');
      toast.style.setProperty('transform', 'translateY(0)', 'important');
    });

    if (duration > 0) setTimeout(dismiss, duration);
  }

  // ---------------------------------------------------------------------------
  // Visual: confirmation modal
  // ---------------------------------------------------------------------------

  function askUser(message, choices) {
    return new Promise(resolve => {
      const overlay = document.createElement('div');
      overlay.style.cssText =
        'position:fixed;inset:0;background:rgba(0,0,0,0.30);z-index:2147483647;' +
        'display:flex;align-items:center;justify-content:center;' +
        'font-family:"Segoe UI",-apple-system,BlinkMacSystemFont,sans-serif;';
      const modal = document.createElement('div');
      modal.style.cssText =
        'background:#fff;padding:20px 22px;border-radius:8px;max-width:480px;' +
        'box-shadow:0 8px 24px rgba(0,0,0,0.25);';
      const title = document.createElement('div');
      title.textContent = 'Negative keywords recommendation';
      title.style.cssText = 'font-weight:600;font-size:15px;margin-bottom:8px;color:#323130;';
      const msg = document.createElement('div');
      msg.textContent = message;
      msg.style.cssText = 'font-size:13px;color:#323130;margin-bottom:16px;line-height:1.4;';
      modal.appendChild(title);
      modal.appendChild(msg);
      const buttons = document.createElement('div');
      buttons.style.cssText = 'display:flex;gap:8px;justify-content:flex-end;flex-wrap:wrap;';
      const close = value => { overlay.remove(); resolve(value); };
      choices.forEach(choice => {
        const b = document.createElement('button');
        b.textContent = choice.label;
        b.style.cssText =
          'padding:6px 14px;border:1px solid #0078d4;border-radius:4px;cursor:pointer;' +
          'font-size:13px;font-weight:500;' +
          (choice.primary
            ? 'background:#0078d4;color:#fff;'
            : 'background:#fff;color:#0078d4;');
        b.addEventListener('click', () => close(choice.value));
        buttons.appendChild(b);
      });
      modal.appendChild(buttons);
      overlay.appendChild(modal);
      overlay.addEventListener('keydown', e => {
        if (e.key === 'Escape') close('skip');
      });
      document.body.appendChild(overlay);
    });
  }

  // ---------------------------------------------------------------------------
  // Core dismissal flow
  // ---------------------------------------------------------------------------

  /**
   * Returns a CSS selector scoped to inside an open dialog, so we never
   * accidentally pick up buttons elsewhere on the page (e.g., "Apply all"
   * buttons on other cards).
   */
  function dialogScoped(innerSelector) {
    const scopes = ['[role="dialog"]', '.ms-Dialog', '.modal', '.ms-Modal'];
    return scopes.map(s => s + ' ' + innerSelector).join(', ');
  }

  async function dismissOneCard(moreBtn) {
    if (!moreBtn) return false;
    moreBtn.click();

    // Menu item: scope to actual menu containers; never plain `button`.
    const dismissBtn = await waitForByText(
      'button.btn-link.btn-block, [role="menuitem"], [role="menu"] button, .ms-ContextualMenu button',
      TEXTS.DISMISS_MENU_ITEM,
      3000
    );
    if (!dismissBtn) {
      console.debug('[DismissAll] dismiss menu item not found (tried texts:', TEXTS.DISMISS_MENU_ITEM, ')');
      pressEscape();
      return false;
    }
    dismissBtn.click();

    const radio = await waitFor('div.radio input[type="radio"], input[type="radio"]');
    if (!radio) {
      console.debug('[DismissAll] dismiss-reason radio not found');
      pressEscape();
      return false;
    }
    radio.click();

    // Submit: try class-based first, then fall back to text matching INSIDE
    // an open dialog only. Never search the whole page — too many "Apply all"
    // and "Yes" buttons on Microsoft Ads cards that would trigger wrong actions.
    let submitBtn = await waitFor('button.submit, button[type="submit"]', b => !b.disabled, 1500);
    if (!submitBtn) {
      submitBtn = await waitForByText(
        dialogScoped('button'),
        TEXTS.SUBMIT,
        2000,
        b => !b.disabled
      );
    }
    if (!submitBtn) {
      console.debug('[DismissAll] submit button not found (tried texts:', TEXTS.SUBMIT, ')');
      document.querySelector('button[aria-label*="Close"], .ms-Dialog-button--close')?.click();
      return false;
    }

    const beforeCount = document.querySelectorAll('button.iconba-More').length;
    submitBtn.click();
    await waitForCondition(
      () => document.querySelectorAll('button.iconba-More').length < beforeCount,
      5000
    );
    await sleep(100);
    return true;
  }

  function updateButton(text, { running } = {}) {
    const btn = document.getElementById('tamper-dismiss-all-btn');
    if (!btn) return;
    if (text != null) btn.textContent = text;
    if (running != null) {
      btn.disabled = running;
      btn.dataset.running = running ? 'true' : 'false';
    }
  }

  async function runDismissAll() {
    if (isRunning) return;
    isRunning = true;
    updateButton('Dismissing...', { running: true });

    // Clear any prior skip marks so a fresh run re-evaluates cards.
    document.querySelectorAll('button.iconba-More[data-dismiss-skipped]')
      .forEach(b => b.removeAttribute('data-dismiss-skipped'));

    // Snapshot the total at run-start so the progress bar has a stable
    // denominator. Cards added later (rare) won't shift the percentage.
    const total = document.querySelectorAll('button.iconba-More').length;
    let count = 0;
    let skipped = 0;
    let bulkChoice = null; // 'dismiss-all' | 'skip-all' | null
    let cancelled = false;

    if (total === 0) {
      isRunning = false;
      updateButton('Dismiss All', { running: false });
      showToast('No recommendations to dismiss.', { kind: 'info', duration: 3500 });
      return;
    }

    updateProgress(0, total);

    try {
      while (count + skipped < total + 50) { // small slack for late-arriving cards
        const moreBtn = document.querySelector('button.iconba-More:not([data-dismiss-skipped])');
        if (!moreBtn) break;

        updateButton(
          'Dismissing... (' + count + ' done' + (skipped ? ', ' + skipped + ' skipped' : '') + ')',
          { running: true }
        );

        if (isNegativeKeywordCard(moreBtn)) {
          let action;
          if (bulkChoice === 'dismiss-all') action = 'dismiss';
          else if (bulkChoice === 'skip-all') action = 'skip';
          else {
            // Visual interlock: outline the card so the user can see what
            // the modal is asking about.
            const restoreHighlight = highlightCard(moreBtn);
            try {
              action = await askUser(
                'This recommendation involves negative keywords. Dismissing it means the suggested negative-keyword changes will not be applied.',
                [
                  { label: 'Skip', value: 'skip' },
                  { label: 'Skip all', value: 'skip-all' },
                  { label: 'Dismiss', value: 'dismiss', primary: true },
                  { label: 'Dismiss all', value: 'dismiss-all' },
                  { label: 'Cancel run', value: 'cancel' },
                ]
              );
            } finally {
              restoreHighlight();
            }
            if (action === 'dismiss-all') { bulkChoice = 'dismiss-all'; action = 'dismiss'; }
            else if (action === 'skip-all') { bulkChoice = 'skip-all'; action = 'skip'; }
          }

          if (action === 'cancel') { cancelled = true; break; }
          if (action === 'skip') {
            moreBtn.setAttribute('data-dismiss-skipped', 'true');
            skipped++;
            updateProgress(count + skipped, total);
            continue;
          }
        }

        if (await dismissOneCard(moreBtn)) {
          count++;
          updateProgress(count + skipped, total);
        } else {
          break;
        }
      }
    } finally {
      isRunning = false;
      // If the toolbar re-rendered mid-run, attach a fresh button.
      attachButton();
      updateButton('Dismiss All', { running: false });
      hideProgressBar();

      // Activity toast summarizing the run.
      const lines = [];
      if (count) lines.push('✓ Dismissed: ' + count);
      if (skipped) lines.push('Skipped: ' + skipped + ' (negative keywords preserved)');
      if (cancelled) lines.push('Run cancelled by user');
      if (!count && !skipped && !cancelled) lines.push('Nothing to dismiss');

      const kind = cancelled ? 'warning' : (count ? 'success' : 'info');
      showToast(lines.join('\n'), { kind, duration: 6000 });
    }
  }

  // ---------------------------------------------------------------------------
  // Mount: button + observer + safety net
  // ---------------------------------------------------------------------------

  function attachButton() {
    if (!location.href.toLowerCase().includes('recommendations')) {
      // User navigated away — drop any stale button.
      document.getElementById('tamper-dismiss-all-btn')?.remove();
      return;
    }
    if (document.getElementById('tamper-dismiss-all-btn')) return;
    const container = document.querySelector('.download-button-container');
    if (!container) return;
    const wrapper = document.createElement('div');
    wrapper.className = 'toolbar-item';
    const btn = document.createElement('button');
    btn.id = 'tamper-dismiss-all-btn';
    btn.textContent = isRunning ? 'Dismissing...' : 'Dismiss All';
    btn.disabled = isRunning;
    btn.dataset.running = isRunning ? 'true' : 'false';
    btn.style.cssText =
      'padding:4px 10px;background:#fff;border:1px solid #0078d4;color:#0078d4;' +
      'border-radius:4px;cursor:pointer;font-size:13px;font-weight:600;';
    btn.addEventListener('mouseenter', () => { btn.style.background = '#e6f2fb'; });
    btn.addEventListener('mouseleave', () => { btn.style.background = '#fff'; });
    btn.addEventListener('mousedown', e => { e.preventDefault(); runDismissAll(); });
    wrapper.appendChild(btn);
    container.appendChild(wrapper);
  }

  function scheduleAttach() {
    if (attachScheduled) return;
    attachScheduled = true;
    requestAnimationFrame(() => {
      attachScheduled = false;
      attachButton();
    });
  }

  // Handle SPA (pushState/replaceState) navigation
  const origPush = history.pushState.bind(history);
  const origReplace = history.replaceState.bind(history);
  history.pushState = (...args) => { origPush(...args); setTimeout(attachButton, 500); };
  history.replaceState = (...args) => { origReplace(...args); setTimeout(attachButton, 500); };
  window.addEventListener('popstate', () => setTimeout(attachButton, 500));

  // Core MutationObserver — detects toolbar re-renders so we can re-attach.
  dismissObserver = new MutationObserver(scheduleAttach);
  dismissObserver.observe(document.documentElement, { childList: true, subtree: true });

  // Safety net: if the SPA re-renders the toolbar after the last mutation we
  // saw, this periodic check re-attaches the button.
  setInterval(() => {
    if (!location.href.toLowerCase().includes('recommendations')) return;
    if (document.getElementById('tamper-dismiss-all-btn')) return;
    attachButton();
  }, 1000);

  attachButton();
})();
