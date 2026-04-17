// ==UserScript==
// @name         Patch - Paste Next Gong Note
// @namespace    http://tampermonkey.net/
// @version      0.3
// @description  Adds "Paste Next Gong Note" button after clicking "Add Note"
// @match        https://thepatch.melonlocal.com/*
// @grant        GM_xmlhttpRequest
// @grant        GM_setValue
// @grant        GM_getValue
// @connect      docs.google.com
// @updateURL    https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/patch-paste-next-gong-note.user.js
// @downloadURL  https://raw.githubusercontent.com/mikemelonlocal/melon-userscripts/main/patch-paste-next-gong-note.user.js
// ==/UserScript==

(function() {
    'use strict';

    const NOTES_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRj7xJnqmEb24fPpHaJJHfbfgytTCBTXwEatai6fKJzTuwdhIMXvj4vWGwh0elsueJ_Ap7EmaTdQhgV/pub?gid=63251332&single=true&output=csv';
    const BTN_ID = 'paste-gong-btn';
    const STATUS_ID = 'paste-gong-status';
    const DONE_KEY = 'pastedRows';

    function setStatus(msg) {
        const el = document.getElementById(STATUS_ID);
        if (el) el.textContent = msg;
        console.log('[PasteGong]', msg);
    }

    function getDoneSet() {
        return new Set(GM_getValue(DONE_KEY, []));
    }

    function markDone(key) {
        const set = getDoneSet();
        set.add(key);
        GM_setValue(DONE_KEY, [...set]);
    }

    function rowKey(r) {
        return `${r.agentName}|${r.callDate}`;
    }

    function fetchText(url) {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                method: 'GET',
                url,
                onload: (res) => resolve(res.responseText),
                onerror: reject
            });
        });
    }

    function parseCsv(text) {
        return text
            .trim()
            .split(/\r?\n/)
            .map(r => r.split(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/))
            .map(cells => cells.map(cell => cell.replace(/^"|"$/g, '')));
    }

    async function getNextNoteRow() {
        const rows = parseCsv(await fetchText(NOTES_URL));
        const header = rows[0];
        const idxAgentName    = header.indexOf('Agent Name');
        const idxCallDate     = header.indexOf('Call Date');
        const idxMonthYear    = header.indexOf('Month Year');
        const idxNotesToPaste = header.indexOf('Notes to Paste');
        const idxDone         = header.indexOf('Done?');
        const done = getDoneSet();

        for (let i = 1; i < rows.length; i++) {
            const r = rows[i];
            const sheetDone = (r[idxDone] || '').toString().trim().toLowerCase();
            if (sheetDone === 'true') continue;

            const row = {
                rowIndex: i + 1,
                agentName: r[idxAgentName],
                callDate:  r[idxCallDate],
                monthYear: r[idxMonthYear],
                notes:     r[idxNotesToPaste]
            };
            if (done.has(rowKey(row))) continue;
            return row;
        }
        return null;
    }

    async function pasteNextNote() {
        try {
            const noteRow = await getNextNoteRow();
            if (!noteRow) {
                setStatus('No unprocessed rows.');
                return;
            }

            const editor = document.querySelector('body[contenteditable="true"]');
            if (!editor) {
                setStatus('Editor not found — click "Add Note" first.');
                return;
            }

            editor.focus();
            editor.innerHTML = '';
            editor.innerText = noteRow.notes;
            editor.dispatchEvent(new InputEvent('input', { bubbles: true }));

            markDone(rowKey(noteRow));
            setStatus(`Pasted row ${noteRow.rowIndex}: ${noteRow.agentName} (${noteRow.callDate})`);
        } catch (err) {
            console.error(err);
            setStatus('Error — check console (F12).');
        }
    }

    function addButton() {
        if (document.getElementById(BTN_ID)) return;
        const saveBtn = document.querySelector('button.saveNote, button.melon-green');
        if (!saveBtn) return;

        const myBtn = document.createElement('button');
        myBtn.id = BTN_ID;
        myBtn.textContent = 'Paste Next Gong Note';
        myBtn.type = 'button';
        myBtn.style.marginLeft = '8px';
        myBtn.style.fontSize = '14px';
        myBtn.className = 'btn melon-green';
        myBtn.onclick = pasteNextNote;

        const status = document.createElement('span');
        status.id = STATUS_ID;
        status.style.marginLeft = '8px';
        status.style.fontSize = '13px';
        status.style.color = '#555';

        saveBtn.parentNode.insertBefore(myBtn, saveBtn.nextSibling);
        myBtn.parentNode.insertBefore(status, myBtn.nextSibling);
        return true;
    }

    let pending = false;
    const observer = new MutationObserver(() => {
        if (pending) return;
        if (!document.querySelector('body[contenteditable="true"]')) return;
        pending = true;
        setTimeout(() => {
            pending = false;
            if (addButton()) observer.disconnect();
        }, 300);
    });
    observer.observe(document.body, { childList: true, subtree: true });

    window.addEventListener('load', () => setTimeout(addButton, 1000));
})();
