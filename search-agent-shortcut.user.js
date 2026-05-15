// ==UserScript==
// @name         Search Agent Shortcut
// @namespace    https://thepatch.melonlocal.com/
// @version      1.5
// @description  Modern Command Palette for Search Agent
// @match        https://thepatch.melonlocal.com/*
// @grant        none
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/search-agent-shortcut.user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/search-agent-shortcut.user.js
// ==/UserScript==

(function () {
  'use strict';

  const INPUT_ID = 'AgentQuickFind';
  const MODAL_ID = 'melon-search-agent-modal';
  const STYLE_ID = 'melon-search-agent-modal-style';

  const isMac = /Mac|iPhone|iPad/.test(navigator.platform);

  let origParent = null;
  let origNextSibling = null;
  let wrapperEl = null;
  let attachedWidget = null;
  let widgetSelectHandler = null;
  let widgetDataBoundHandler = null;

  function focusFirstSuggestion(widget) {
    if (widget.listView && typeof widget.listView.focus === 'function') {
      try { widget.listView.focus(0); return; } catch (_) {}
    }
    if (window.$ && window.$.Event) {
      try {
        window.$(widget.element).trigger(
          window.$.Event('keydown', { keyCode: 40, which: 40, key: 'ArrowDown' })
        );
      } catch (_) {}
    }
  }

  function injectStyles() {
    if (document.getElementById(STYLE_ID)) return;
    const style = document.createElement('style');
    style.id = STYLE_ID;
    style.textContent = `
      #${MODAL_ID} {
        position: fixed;
        inset: 0;
        z-index: 2147483640;
        background: rgba(15, 23, 42, 0.4);
        backdrop-filter: blur(8px);
        display: flex;
        align-items: flex-start;
        justify-content: center;
        padding-top: 15vh;
        opacity: 0;
        transition: opacity 150ms ease-out;
      }
      #${MODAL_ID}.is-open { opacity: 1; }

      #${MODAL_ID} .melon-sa-card {
        background: #ffffff;
        border-radius: 14px;
        border: 1px solid rgba(0, 0, 0, 0.08);
        box-shadow: 0 20px 50px -12px rgba(0, 0, 0, 0.25), 0 0 1px rgba(0,0,0,0.2);
        width: min(680px, 94vw);
        padding: 16px;
        transform: scale(0.96) translateY(-10px);
        transition: transform 150ms cubic-bezier(0.16, 1, 0.3, 1);
      }
      #${MODAL_ID}.is-open .melon-sa-card { transform: scale(1) translateY(0); }

      #${MODAL_ID} .melon-sa-label {
        font: 600 12px/1 Inter, -apple-system, sans-serif;
        color: #94a3b8;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-bottom: 12px;
        display: flex;
        justify-content: space-between;
        align-items: center;
      }

      .melon-sa-kbd {
        background: #f1f5f9;
        border: 1px solid #e2e8f0;
        border-radius: 4px;
        padding: 2px 5px;
        font-size: 10px;
        color: #64748b;
        box-shadow: 0 1px 0 rgba(0,0,0,0.1);
      }

      /* Fix Kendo Popup Z-Index */
      .k-animation-container { z-index: 2147483647 !important; }

      /* Highlight color for the focused/selected suggestion while palette is open */
      body.melon-sa-locked .k-animation-container .k-list-item.k-focus,
      body.melon-sa-locked .k-animation-container .k-list-item.k-selected,
      body.melon-sa-locked .k-animation-container .k-list-item.k-state-focused,
      body.melon-sa-locked .k-animation-container li.k-item.k-state-focused,
      body.melon-sa-locked .k-animation-container li.k-item.k-state-selected {
        background-color: #40A74C !important;
        color: #ffffff !important;
      }

      body.melon-sa-locked { overflow: hidden !important; }
    `;
    document.head.appendChild(style);
  }

  function openModal() {
    if (!!document.getElementById(MODAL_ID)) return;
    const input = document.getElementById(INPUT_ID);
    wrapperEl = input ? input.closest('.k-autocomplete') : null;
    if (!wrapperEl) return;

    injectStyles();
    document.body.classList.add('melon-sa-locked');

    origParent = wrapperEl.parentNode;
    origNextSibling = wrapperEl.nextSibling;

    const modal = document.createElement('div');
    modal.id = MODAL_ID;
    modal.innerHTML = `
      <div class="melon-sa-card" role="dialog">
        <div class="melon-sa-label">
          <span>Search Agent</span>
          <div><kbd class="melon-sa-kbd">Esc</kbd> <span style="font-size: 10px; text-transform:none">to close</span></div>
        </div>
        <div class="melon-sa-slot"></div>
      </div>`;

    modal.addEventListener('mousedown', (e) => { if (e.target === modal) closeModal(); });
    document.body.appendChild(modal);
    modal.querySelector('.melon-sa-slot').appendChild(wrapperEl);

    requestAnimationFrame(() => modal.classList.add('is-open'));

    if (input) {
      input.focus();
      input.select();
    }

    // Bind Kendo events to close on selection
    const widget = window.$ && window.$('#' + INPUT_ID).data('kendoAutoComplete');
    if (widget) {
      widgetSelectHandler = () => setTimeout(closeModal, 50);
      widget.bind('select', widgetSelectHandler);

      // Auto-focus the only suggestion when results narrow to one,
      // so pressing Enter selects it.
      widgetDataBoundHandler = () => {
        const view = widget.dataSource && widget.dataSource.view();
        if (view && view.length === 1) focusFirstSuggestion(widget);
      };
      widget.bind('dataBound', widgetDataBoundHandler);

      attachedWidget = widget;
      try { widget.popup.position(); } catch(e) {}
    }
  }

  function closeModal() {
    const modal = document.getElementById(MODAL_ID);
    if (!modal) return;

    document.body.classList.remove('melon-sa-locked');
    if (attachedWidget) {
      if (widgetSelectHandler) {
        try { attachedWidget.unbind('select', widgetSelectHandler); } catch (_) {}
      }
      if (widgetDataBoundHandler) {
        try { attachedWidget.unbind('dataBound', widgetDataBoundHandler); } catch (_) {}
      }
    }
    widgetSelectHandler = null;
    widgetDataBoundHandler = null;

    if (wrapperEl && origParent) {
      if (origNextSibling && origNextSibling.parentNode === origParent) {
        origParent.insertBefore(wrapperEl, origNextSibling);
      } else {
        origParent.appendChild(wrapperEl);
      }
    }

    modal.classList.remove('is-open');
    setTimeout(() => modal.remove(), 150);
  }

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && document.getElementById(MODAL_ID)) {
      e.preventDefault();
      closeModal();
    }
    const mod = isMac ? e.metaKey : e.ctrlKey;
    if (mod && e.code === 'KeyB') {
      e.preventDefault();
      document.getElementById(MODAL_ID) ? closeModal() : openModal();
    }
  });
})();
