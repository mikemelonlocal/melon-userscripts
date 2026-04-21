// ==UserScript==
// @name         Microsoft Ads - Dismiss All Recommendations
// @namespace    http://tampermonkey.net/
// @version      1.9.1
// @description  Adds a "Dismiss All" button to the Microsoft Advertising recommendations page
// @author       You
// @match        https://ui.ads.microsoft.com/*recommendations*
// @run-at       document-start
// @grant        none
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/microsoft-ads-dismiss-recommendations.user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/microsoft-ads-dismiss-recommendations.user.js
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/microsoft-ads-dismiss-recommendations.user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/microsoft-ads-dismiss-recommendations.user.js
// ==/UserScript==

(function () {
  'use strict';

  let dismissObserver = null;
  let isRunning = false;

  async function waitFor(selector, filterFn, timeout = 3000) {
    const start = performance.now();
    while (performance.now() - start < timeout) {
      const els = Array.from(document.querySelectorAll(selector));
      const match = filterFn ? els.find(filterFn) : els[0];
      if (match && match.offsetParent !== null) return match;
      await new Promise(r => setTimeout(r, 20));
    }
    return null;
  }

  async function waitForCondition(condition, timeout = 5000) {
    const start = performance.now();
    while (performance.now() - start < timeout) {
      if (condition()) return true;
      await new Promise(r => setTimeout(r, 20));
    }
    return false;
  }

  function pressEscape() {
    document.dispatchEvent(new KeyboardEvent('keydown', {
      key: 'Escape', code: 'Escape', keyCode: 27, which: 27, bubbles: true
    }));
  }

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

  function askUser(message, choices) {
    return new Promise(resolve => {
      const overlay = document.createElement('div');
      overlay.style.cssText =
        'position:fixed;inset:0;background:rgba(0,0,0,0.45);z-index:2147483647;' +
        'display:flex;align-items:center;justify-content:center;' +
        'font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;';
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

  async function dismissOneCard(moreBtn) {
    if (!moreBtn) return false;
    moreBtn.click();
    const dismissBtn = await waitFor('button.btn-link.btn-block', b => b.textContent.trim() === 'Dismiss all');
    if (!dismissBtn) {
      console.debug('[DismissAll] "Dismiss all" menu item not found');
      pressEscape();
      return false;
    }
    dismissBtn.click();
    const radio = await waitFor('div.radio input[type="radio"]');
    if (!radio) {
      console.debug('[DismissAll] dismiss-reason radio not found');
      return false;
    }
    radio.click();
    const submitBtn = await waitFor('button.submit', b => !b.disabled);
    if (!submitBtn) {
      console.debug('[DismissAll] submit button not found');
      document.querySelector('button[aria-label*="Close"], .ms-Dialog-button--close')?.click();
      return false;
    }
    const beforeCount = document.querySelectorAll('button.iconba-More').length;
    submitBtn.click();
    await waitForCondition(
      () => document.querySelectorAll('button.iconba-More').length < beforeCount,
      5000
    );
    await new Promise(r => setTimeout(r, 100));
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
    updateButton(null, { running: true });

    // Clear any prior skip marks so a fresh run re-evaluates cards.
    document.querySelectorAll('button.iconba-More[data-dismiss-skipped]')
      .forEach(b => b.removeAttribute('data-dismiss-skipped'));

    let count = 0;
    let skipped = 0;
    let bulkChoice = null; // 'dismiss-all' | 'skip-all' | null
    let cancelled = false;

    try {
      while (count + skipped < 200) {
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
            if (action === 'dismiss-all') { bulkChoice = 'dismiss-all'; action = 'dismiss'; }
            else if (action === 'skip-all') { bulkChoice = 'skip-all'; action = 'skip'; }
          }

          if (action === 'cancel') { cancelled = true; break; }
          if (action === 'skip') {
            moreBtn.setAttribute('data-dismiss-skipped', 'true');
            skipped++;
            continue;
          }
        }

        if (await dismissOneCard(moreBtn)) count++;
        else break;
      }
    } finally {
      isRunning = false;
      // If the toolbar re-rendered mid-run, attach a fresh button.
      attachButton();
      const parts = [];
      if (count) parts.push('✓ Dismissed ' + count);
      if (skipped) parts.push(skipped + ' skipped');
      if (cancelled) parts.push('cancelled');
      updateButton(parts.length ? parts.join(', ') : 'Dismiss All', { running: false });
      setTimeout(() => {
        if (!isRunning) updateButton('Dismiss All', { running: false });
      }, 4000);
    }
  }

  function attachButton() {
    if (!location.href.includes('recommendations')) return;
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
    btn.style.cssText = 'padding:4px 10px;background:#fff;border:1px solid #0078d4;color:#0078d4;border-radius:4px;cursor:pointer;font-size:13px;font-weight:600;';
    btn.addEventListener('mouseenter', () => { btn.style.background = '#e6f2fb'; });
    btn.addEventListener('mouseleave', () => { btn.style.background = '#fff'; });
    btn.addEventListener('mousedown', e => { e.preventDefault(); runDismissAll(); });
    wrapper.appendChild(btn);
    container.appendChild(wrapper);
  }

  let attachScheduled = false;
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

  dismissObserver = new MutationObserver(scheduleAttach);
  dismissObserver.observe(document.documentElement, { childList: true, subtree: true });

  attachButton();
})();
