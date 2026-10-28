// ==UserScript==
// @name         تصدير عهدة المختبر - عتاد
// @namespace    mulawah-school
// @author       سيف جمال المعولي
// @updateURL    https://mawals.github.io/gdg/atad.user.js
// @downloadURL  https://mawals.github.io/gdg/atad.user.js
// @version      3.7
// @description  أعمالي ← عتاد ← المدارس ← المدرسة ← تقارير المختبر ← عهدة الأصناف المخبرية، ثم يمر على كل نوع رئيسي وفرعي وكل الصفحات ويجمعها في Excel واحد (أو يضغط زر التصدير لكل صفحة)
// @match        *://*.moe.gov.om/*
// @include      /^https?:\/\/([^\/]*\.)?moe\.gov\.om(:\d+)?\/.*$/
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_deleteValue
// @grant        GM_notification
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';
  if (window.top !== window.self) return;

  // ================= الإعدادات =================
  // المدرسة: تُحفظ تلقائيًا أول مرة يختارها الفني في عتاد
  const SCHK = 'atad3_school', SCHN = 'atad3_schoolname';
  const DEPT = 'المدارس';                     // القسم
  const MENU = 'تقارير المختبر';
  const REPORT = 'عهدة الأصناف المخبرية';
  // =============================================

  const KEY = 'atad2_state', LAST = 'atad2_last', MODE = 'atad2_mode', WAITK = 'atad2_wait';
  const WAIT = () => Math.max(20, +GM_getValue(WAITK, 90) || 90) * 1000; // مهلة انتظار استجابة عتاد
  const ORDER = ['portal', 'selplace', 'menu', 'report'];
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const load = () => GM_getValue(KEY, null);
  const save = s => GM_setValue(KEY, s);
  const clear = () => GM_deleteValue(KEY);
  const norm = t => (t || '').replace(/\s+/g, ' ').trim();
  const AR = t => norm(t).replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي').replace(/[\u064B-\u0652\u0640]/g, '');
  const today = () => new Date().toISOString().slice(0, 10);
  const esc = v => String(v ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  let stopFlag = false, running = false, leaving = false;
  // عند بدء تحميل صفحة جديدة لا نضغط أي شيء (حتى لا نلغي طلبًا جاريًا)
  window.addEventListener('beforeunload', () => { leaving = true; });
  window.addEventListener('pagehide', () => { leaving = true; });

  function urlStage() {
    const p = location.pathname.toLowerCase();
    if (p.includes('lab_report_stock')) return 'report';
    if (p.includes('login_selplace')) return 'selplace';
    if (p.includes('/atad/')) return 'menu';
    if (p.includes('/portal')) return 'portal';
    return null;
  }

  // ================= اللوحة الفاخرة (Obsidian & Amber Gold UI) =================
  if (!document.getElementById('ax-cairo-font')) {
    const link = document.createElement('link');
    link.id = 'ax-cairo-font';
    link.rel = 'stylesheet';
    link.href = 'https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800;900&display=swap';
    document.head.appendChild(link);
  }

  const panel = document.createElement('div');
  panel.id = 'ax-panel';
  panel.dir = 'rtl';
  panel.innerHTML = `
    <style>
      #ax-panel {
        position: fixed;
        bottom: 18px;
        left: 18px;
        z-index: 2147483647;
        background: #1c1917;
        color: #f5f5f4;
        border: 1px solid #292524;
        border-radius: 22px;
        padding: 15px;
        width: 300px;
        font-family: 'Cairo', Tahoma, Arial, sans-serif;
        font-size: 12.5px;
        box-shadow: 0 22px 45px -12px rgba(0,0,0,0.65), 0 0 0 1px rgba(245,158,11,0.16);
        overflow: hidden;
        transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
        line-height: 1.5;
        box-sizing: border-box;
      }
      #ax-panel * { box-sizing: border-box; font-family: 'Cairo', Tahoma, Arial, sans-serif; }
      #ax-panel::before {
        content: "";
        position: absolute;
        top: -50px;
        left: -50px;
        width: 130px;
        height: 130px;
        background: radial-gradient(circle, rgba(245,158,11,0.18), transparent 70%);
        pointer-events: none;
      }
      #ax-head {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding-bottom: 11px;
        margin-bottom: 11px;
        border-bottom: 1px solid #292524;
        position: relative;
        z-index: 1;
      }
      .ax-brand {
        display: flex;
        align-items: center;
        gap: 9px;
      }
      .ax-badge {
        position: relative;
        width: 34px;
        height: 34px;
        border-radius: 11px;
        background: linear-gradient(135deg, #f59e0b, #ea580c);
        color: #0c0a09;
        font-weight: 900;
        font-size: 12px;
        display: flex;
        align-items: center;
        justify-content: center;
        box-shadow: 0 4px 12px rgba(234, 88, 12, 0.35);
      }
      .ax-badge::after {
        content: "";
        position: absolute;
        bottom: -2px;
        left: -2px;
        width: 9px;
        height: 9px;
        border-radius: 50%;
        background: #10b981;
        border: 2px solid #1c1917;
      }
      .ax-title-wrap { line-height: 1.25; }
      .ax-title { font-weight: 900; font-size: 13.5px; color: #ffffff; }
      .ax-sub { font-size: 10.5px; color: #a8a29e; font-weight: 600; }
      #ax-min {
        width: 28px;
        height: 28px;
        border-radius: 9px;
        background: #292524;
        color: #d6d3d1;
        display: flex;
        align-items: center;
        justify-content: center;
        cursor: pointer;
        font-weight: 900;
        user-select: none;
        transition: 0.2s;
      }
      #ax-min:hover { background: #44403c; color: #fbbf24; }
      #ax-body { position: relative; z-index: 1; }
      #ax-panel button {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 6px;
        width: 100%;
        margin: 5px 0;
        padding: 9px 12px;
        border: 1px solid transparent;
        border-radius: 12px;
        cursor: pointer;
        font-weight: 800;
        font-size: 12.5px;
        transition: all 0.2s ease;
      }
      #ax-panel button:active { transform: scale(0.98); }
      #ax-panel button[data-a="start"] {
        background: linear-gradient(135deg, #f59e0b, #ea580c);
        color: #0c0a09;
        font-size: 13.5px;
        font-weight: 900;
        box-shadow: 0 8px 20px -4px rgba(234, 88, 12, 0.45);
        margin-bottom: 9px;
      }
      #ax-panel button[data-a="start"]:hover {
        background: linear-gradient(135deg, #fbbf24, #f97316);
        transform: translateY(-1px);
      }
      .ax-options {
        background: #0c0a09;
        border: 1px solid #292524;
        border-radius: 14px;
        padding: 9px 11px;
        margin-bottom: 8px;
      }
      .ax-radio {
        display: flex;
        align-items: center;
        gap: 7px;
        margin: 4px 0;
        color: #d6d3d1;
        font-size: 11.5px;
        font-weight: 600;
        cursor: pointer;
      }
      .ax-radio input[type="radio"] {
        accent-color: #f59e0b;
        width: 14px;
        height: 14px;
        margin: 0;
        cursor: pointer;
      }
      .ax-wait-row {
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-top: 7px;
        padding-top: 7px;
        border-top: 1px dashed #292524;
        color: #a8a29e;
        font-size: 11.5px;
        font-weight: 700;
      }
      #ax-wait {
        width: 58px;
        background: #1c1917;
        color: #fbbf24;
        border: 1px solid #44403c;
        border-radius: 8px;
        padding: 2px 6px;
        text-align: center;
        font-weight: 800;
        font-size: 12px;
      }
      #ax-wait:focus { outline: none; border-color: #f59e0b; }
      .ax-grid-btns {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 6px;
        margin: 5px 0;
      }
      .ax-grid-btns button { margin: 0 !important; padding: 7.5px 8px !important; font-size: 11.5px !important; }
      #ax-panel button[data-a="stop"] {
        background: rgba(225, 29, 72, 0.14);
        color: #fda4af;
        border-color: rgba(225, 29, 72, 0.3);
      }
      #ax-panel button[data-a="stop"]:hover { background: rgba(225, 29, 72, 0.25); color: #fff; }
      #ax-panel button[data-a="xlsx"], #ax-panel button[data-a="pdf"] {
        background: #292524;
        color: #e7e5e4;
        border-color: #3f3a36;
      }
      #ax-panel button[data-a="xlsx"]:hover, #ax-panel button[data-a="pdf"]:hover {
        background: #3f3a36;
        border-color: #f59e0b;
        color: #fbbf24;
      }
      #ax-panel button[data-a="diag"] {
        background: transparent;
        color: #a8a29e;
        border: 1px dashed #3f3a36;
        padding: 6px !important;
        font-size: 11.5px !important;
      }
      #ax-panel button[data-a="diag"]:hover { color: #e7e5e4; border-color: #78716c; background: rgba(255,255,255,0.03); }
      #ax-status {
        margin-top: 8px;
        padding: 9px 11px;
        background: #0c0a09;
        border: 1px solid #292524;
        border-right: 3px solid #f59e0b;
        border-radius: 11px;
        min-height: 22px;
        white-space: pre-line;
        color: #fafaf9;
        font-size: 11.5px;
        font-weight: 700;
        line-height: 1.55;
      }
      #ax-timer {
        display: none;
        margin-top: 8px;
        padding: 8px 10px;
        border-radius: 11px;
        background: rgba(245, 158, 11, 0.1);
        border: 1px solid rgba(245, 158, 11, 0.3);
        text-align: center;
      }
      #ax-timer-txt { font-weight: 800; color: #fbbf24; font-size: 11.5px; }
      .ax-timer-track {
        height: 6px;
        background: #292524;
        border-radius: 99px;
        margin-top: 6px;
        overflow: hidden;
      }
      #ax-timer-bar {
        height: 100%;
        width: 0;
        background: linear-gradient(90deg, #f59e0b, #ea580c);
        transition: width 0.4s ease;
      }
      .ax-footer {
        margin-top: 10px;
        padding-top: 8px;
        border-top: 1px solid #292524;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 8px;
      }
      .ax-designer {
        display: flex;
        align-items: center;
        gap: 7px;
      }
      .ax-flask-icon {
        width: 26px;
        height: 26px;
        border-radius: 8px;
        background: rgba(245, 158, 11, 0.12);
        border: 1px solid rgba(245, 158, 11, 0.25);
        color: #fbbf24;
        display: flex;
        align-items: center;
        justify-content: center;
        flex-shrink: 0;
      }
      .ax-credit-txt {
        font-size: 10.5px;
        color: #a8a29e;
        line-height: 1.35;
      }
      .ax-credit-txt b { color: #f5f5f4; font-weight: 800; display: block; }
    </style>

    <div id="ax-head">
      <div class="ax-brand">
        <div class="ax-badge">عتاد</div>
        <div class="ax-title-wrap">
          <div class="ax-title">مُصدِّر عهدة المختبر</div>
          <div class="ax-sub">الأتمتة الذكية • الإصدار ${(typeof GM_info !== 'undefined' && GM_info.script && GM_info.script.version) || ''}</div>
        </div>
      </div>
      <span id="ax-min" title="تصغير / تكبير">—</span>
    </div>

    <div id="ax-body">
      <button data-a="start">▶️ ابدأ الفحص والتصدير</button>

      <div class="ax-options">
        <label class="ax-radio">
          <input type="radio" name="ax-mode" value="collect">
          <span>تجميع الكل في ملف Excel واحد</span>
        </label>
        <label class="ax-radio">
          <input type="radio" name="ax-mode" value="export">
          <span>الضغط على زر التصدير لكل صفحة</span>
        </label>
        <div class="ax-wait-row">
          <span>⏱️ انتظار استجابة عتاد (ثانية)</span>
          <input type="number" id="ax-wait" min="20" max="600">
        </div>
      </div>

      <button data-a="stop">⏹ إيقاف (وتصدير ما جُمع)</button>

      <div class="ax-grid-btns">
        <button data-a="xlsx">📥 تنزيل Excel</button>
        <button data-a="pdf">🖨️ كشف PDF</button>
      </div>

      <button data-a="diag">🔍 فحص عناصر الصفحة</button>

      <div id="ax-status">جاهز للفحص والتصدير.</div>

      <div id="ax-timer">
        <div id="ax-timer-txt"></div>
        <div class="ax-timer-track"><div id="ax-timer-bar"></div></div>
      </div>

      <div class="ax-footer">
        <div class="ax-designer">
          <div class="ax-flask-icon">
            <svg width="15" height="15" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2">
              <path stroke-linecap="round" stroke-linejoin="round" d="M9.75 3.104v5.714a2.25 2.25 0 01-.659 1.591L5 14.5M9.75 3.104c-.251.023-.501.05-.75.082m.75-.082a24.301 24.301 0 014.5 0m0 0v5.714c0 .597.237 1.17.659 1.591L19.8 15.3M14.25 3.104c.251.023.501.05.75.082M19.8 15.3l-1.57.393A9.065 9.065 0 0112 15a9.065 9.065 0 00-6.23-.693L5 14.5m14.8.8l1.402 1.402c1.232 1.232.65 3.318-1.067 3.611A48.309 48.309 0 0112 21c-2.773 0-5.491-.235-8.135-.687-1.718-.293-2.3-2.379-1.067-3.61L5 14.5" />
            </svg>
          </div>
          <div class="ax-credit-txt">
            <b>سيف جمال المعولي</b>
            <span>معولة بن شمس (5-12)</span>
          </div>
        </div>
      </div>
    </div>`;

  document.body.appendChild(panel);
  const status = t => (panel.querySelector('#ax-status').textContent = t);
  panel.querySelector('#ax-min').onclick = () => {
    const b = panel.querySelector('#ax-body');
    const isHidden = b.style.display === 'none';
    b.style.display = isHidden ? '' : 'none';
    panel.querySelector('#ax-min').textContent = isHidden ? '—' : '+';
  };
  const mode = GM_getValue(MODE, 'collect');
  panel.querySelector(`input[value=${mode}]`).checked = true;
  panel.querySelectorAll('input[name=ax-mode]').forEach(r => r.onchange = () => GM_setValue(MODE, r.value));
  const waitBox = panel.querySelector('#ax-wait');
  waitBox.value = WAIT() / 1000;
  waitBox.onchange = () => { const v = Math.min(600, Math.max(20, +waitBox.value || 90)); waitBox.value = v; GM_setValue(WAITK, v); };

  panel.addEventListener('click', e => {
    const btn = e.target.closest('button[data-a]');
    const a = btn && btn.dataset.a;
    if (!a) return;
    if (a === 'start') {
      const st = urlStage();
      if (!st) { status('افتح البوابة التعليمية أو صفحة عتاد أولًا.'); return; }
      stopFlag = false;
      save({ stage: st, mode: GM_getValue(MODE, 'collect'), mi: 0, si: 0, phase: 'sub', target: 1, done: [], rows: [], headers: null, files: 0, retryKey: '', retry: 0, log: {}, rec: {}, cnt: {}, pg: {} });
      dispatch();
    }
    if (a === 'stop') { stopFlag = true; const s = load(); clear(); s && s.stage === 'report' ? finish(s, true) : status('⏹ تم الإيقاف.'); }
    if (a === 'xlsx') { const d = GM_getValue(LAST, null); d ? exportExcel(d) : status('لا توجد نتيجة محفوظة بعد.'); }
    if (a === 'pdf') { const d = GM_getValue(LAST, null); d && d.rows && d.rows.length ? printPDF(d) : status('لا توجد نتيجة محفوظة بعد.'); }
    if (a === 'diag') diagnose();
  });

  // ================= أدوات =================
  const isVisible = el => !!(el && el.getClientRects().length && getComputedStyle(el).visibility !== 'hidden');
  const textOf = el => norm(el.innerText || el.value || el.getAttribute('title') || el.getAttribute('alt') || '');
  const isLink = el => !!el.closest('a[href],button,input,[onclick],[role=button]');
  const timerBox = panel.querySelector('#ax-timer'), timerTxt = panel.querySelector('#ax-timer-txt'), timerBar = panel.querySelector('#ax-timer-bar');
  function showTimer(sec, total) {
    timerBox.style.display = 'block';
    timerTxt.textContent = `⏳ بانتظار عتاد: ${sec} / ${total} ثانية`;
    timerBar.style.width = Math.min(100, sec / total * 100) + '%';
  }
  const hideTimer = () => { timerBox.style.display = 'none'; };
  async function waitFor(fn, ms, label) {
    const t0 = Date.now(), end = t0 + ms, total = Math.round(ms / 1000);
    let lastSec = -1;
    if (label) status(label);
    try {
      while (Date.now() < end && !stopFlag && !leaving) {
        const v = fn(); if (v) return v;
        if (label) { const sec = Math.floor((Date.now() - t0) / 1000); if (sec !== lastSec) { lastSec = sec; showTimer(sec, total); } }
        await sleep(400);
      }
      return null;
    } finally { if (label && !leaving) hideTimer(); }
  }
  function clickEl(el) { if (el.tagName === 'A') el.removeAttribute('target'); el.click(); }

  function findByText(txt, allowHidden) {
    const want = AR(txt);
    const els = [...document.querySelectorAll('a,button,input[type=button],input[type=submit],input[type=image],li,span,div,td,img,h1,h2,h3,h4,h5,p,label')]
      .filter(el => !panel.contains(el) && (allowHidden || isVisible(el)));
    let best = null, bestScore = Infinity;
    for (const el of els) {
      const t = AR(textOf(el));
      if (!t || !t.includes(want)) continue;
      const score = (t === want ? 0 : 1000) + t.length + el.querySelectorAll('*').length;
      if (score < bestScore) { best = el; bestScore = score; }
    }
    return best ? (best.closest('a,button,[onclick],[role=button],input') || best) : null;
  }
  function findButton(words) {
    const els = [...document.querySelectorAll('input[type=submit],input[type=button],input[type=image],button,a')]
      .filter(el => !panel.contains(el) && isVisible(el));
    for (const w of words) {
      const hit = els.find(el => AR(textOf(el)) === AR(w)) || els.find(el => AR(textOf(el) + ' ' + (el.id || '')).includes(AR(w)));
      if (hit) return hit;
    }
    return null;
  }
  const realOptions = sel => [...sel.options].filter(o => o.value && !['0', '-1'].includes(o.value) && !/اختر|^-+|الكل|select/i.test(o.text));
  function labelFor(sel) {
    let t = (sel.id || '') + ' ' + (sel.name || '');
    const lab = sel.id && document.querySelector(`label[for="${sel.id}"]`);
    if (lab) t += ' ' + lab.innerText;
    const td = sel.closest('td');
    if (td && td.previousElementSibling) t += ' ' + td.previousElementSibling.innerText;
    let p = sel.previousSibling, n = 0;
    while (p && n++ < 4) { t += ' ' + (p.textContent || ''); p = p.previousSibling; }
    return AR(t);
  }
  function findSelect(words, idx) {
    const sels = [...document.querySelectorAll('select')].filter(s => !panel.contains(s) && isVisible(s));
    for (const s of sels) { const l = labelFor(s); if (words.some(w => l.includes(AR(w)))) return s; }
    const cand = sels.filter(s => realOptions(s).length && !realOptions(s).every(o => /\d{4}/.test(o.text)));
    return cand[idx] || null;
  }
  async function setSelect(sel, value) {
    sel.value = value;
    sel.dispatchEvent(new Event('change', { bubbles: true }));
    await sleep(2500);
  }

  // ================= الجدول =================
  const dataRows = t => [...t.rows].filter(r => r.cells.length >= 2 && !r.querySelector('table') && !r.querySelector('th'));
  function findGrid() {
    let best = null, n = 0;
    for (const t of document.querySelectorAll('table')) {
      if (panel.contains(t) || !isVisible(t)) continue;
      const rows = dataRows(t);
      const looks = t.querySelector('th') || rows.some(r => AR(r.innerText).includes('التفاصيل'));
      if (!looks) continue;
      if (rows.length > n) { n = rows.length; best = t; }
    }
    return best;
  }
  const gridSig = () => { const g = findGrid(); if (!g) return 'none'; const r = dataRows(g); return r.length + '|' + r.slice(0, 2).map(x => norm(x.innerText)).join('|'); };
  function cellVal(td) {
    let v = norm(td.innerText);
    if (!v) { const i = td.querySelector('input:not([type=checkbox]),select,textarea'); if (i) v = norm(i.value); }
    if (/^-?(0|[1-9]\d*)(\.\d+)?$/.test(v)) return Number(v);
    return v;
  }
  function readGrid(t) {
    const hr = [...t.rows].find(r => r.querySelector('th'));
    let headers = hr ? [...hr.cells].map(c => norm(c.innerText)) : null;
    let rows = dataRows(t).map(r => [...r.cells].map(cellVal));
    const drop = new Set();
    const width = Math.max(headers ? headers.length : 0, ...rows.map(r => r.length));
    for (let i = 0; i < width; i++) {
      if ((headers && AR(headers[i] || '').includes('التفاصيل')) || (rows.length && rows.every(r => AR(String(r[i] ?? '')).includes('التفاصيل')))) drop.add(i);
    }
    if (headers) headers = headers.filter((_, i) => !drop.has(i));
    rows = rows.map(r => r.filter((_, i) => !drop.has(i)));
    return { headers, rows };
  }

  // ================= أرقام الصفحات =================
  function getPager() {
    const cands = [...document.querySelectorAll('a,span,button,li,input[type=submit],input[type=button]')]
      .filter(el => !panel.contains(el) && isVisible(el) && /^\d{1,4}$/.test(textOf(el)) && !el.querySelector('a,span,button'));
    const groups = new Map();
    for (const el of cands) {
      let p = el.parentElement;
      for (let i = 0; i < 5 && p; i++, p = p.parentElement) { if (!groups.has(p)) groups.set(p, []); groups.get(p).push(el); }
    }
    let best = null;
    for (const [box, els] of groups) {
      if (els.length < 2 || !els.some(isLink)) continue;
      const nums = els.map(e => textOf(e)).join('').length;
      const all = norm(box.innerText).replace(/[\s.…<>«»‹›]|التالي|السابق|Next|Prev|First|Last/gi, '').length || 1;
      if (nums / all < 0.6) continue;
      if (!best || els.length > best.els.length) best = { box, els };
    }
    if (!best) return null;
    const items = best.els.map(el => ({ n: +textOf(el), el, link: isLink(el) }));
    const cur = items.find(i => !i.link) ||
      items.find(i => /active|current|selected/i.test((i.el.className || '') + ' ' + ((i.el.closest('li') || {}).className || '')) || i.el.getAttribute('aria-current'));
    return { box: best.box, items, current: cur ? cur.n : null };
  }
  function findNext(pager, cur) {
    const direct = pager.items.find(i => i.n === cur + 1 && i.link);
    if (direct) return direct.el;
    const last = pager.items.reduce((a, b) => (b.n > a.n ? b : a)).el;
    return [...pager.box.querySelectorAll('a,button,input,span[onclick]')].find(el =>
      /^(\.\.\.|…|>|›|»|التالي|Next)$/i.test(textOf(el)) && (last.compareDocumentPosition(el) & Node.DOCUMENT_POSITION_FOLLOWING)) || null;
  }

  // ================= المراحل =================
  async function dispatch() {
    if (running) return;
    let s = load();
    if (!s) return;
    const st = urlStage();
    if (!st) return;
    if (ORDER.indexOf(st) > ORDER.indexOf(s.stage)) { s.stage = st; save(s); }
    if (st !== s.stage) return; // هذا التبويب ليس مكان المرحلة الحالية
    running = true;
    try {
      if (st === 'portal') await doPortal();
      else if (st === 'selplace') await doSelPlace();
      else if (st === 'menu') await doMenu();
      else if (st === 'report') await doReport();
    } finally { running = false; }
  }

  async function doPortal() {
    status('١- أعمالي ← عتاد');
    let el = findByText('عتاد');
    if (!el) {
      const my = findByText('أعمالي');
      if (my) { my.click(); await sleep(1200); }
      el = (await waitFor(() => findByText('عتاد'), 5000)) || findByText('عتاد', true);
    }
    if (!el) { status('⚠️ لم أجد «عتاد». افتحه يدويًا، والأداة تكمل من هناك.'); return; }
    const s = load(); s.stage = 'selplace'; save(s);
    clickEl(el);
    status('تم فتح عتاد — تابع في تبويب عتاد.');
  }

  async function doSelPlace() {
    status('٢- اختيار القسم والمدرسة');
    // نعيد البحث عن القوائم في كل مرة لأن عتاد يعيد رسم الصفحة بعد كل اختيار
    const findDept = () => [...document.querySelectorAll('select')].find(x => isVisible(x) && [...x.options].some(o => AR(o.text) === AR(DEPT)));
    const deptOk = d => d && AR((d.options[d.selectedIndex] || {}).text) === AR(DEPT);
    const findSch = () => [...document.querySelectorAll('select')].find(x => isVisible(x) && !panel.contains(x) &&
      ![...x.options].some(o => AR(o.text) === AR(DEPT)) && realOptions(x).length);

    let sch = null;
    for (let tries = 0; tries < 4 && !sch; tries++) {
      const dept = await waitFor(findDept, WAIT(), 'تحميل قائمة الأقسام');
      if (stopFlag || leaving) return;
      if (!dept) { status(`⚠️ لم أجد قائمة القسم «${DEPT}».`); return; }
      if (!deptOk(dept)) {
        const opt = [...dept.options].find(o => AR(o.text) === AR(DEPT));
        await setSelect(dept, opt.value);
        if (leaving) return;
      }
      // ننتظر ظهور قائمة المدارس مع بقاء القسم على «المدارس»
      sch = await waitFor(() => { const d = findDept(); if (d && !deptOk(d)) return null; return findSch(); }, WAIT(), 'تحميل قائمة المدارس');
      if (stopFlag || leaving) return;
    }
    if (!sch) { status('⚠️ لم تظهر قائمة المدارس. اختر القسم والمدرسة يدويًا والأداة تكمل.'); return; }

    const opts = realOptions(sch);
    const key = GM_getValue(SCHK, '');
    let so = (key && opts.find(o => AR(o.text) === AR(key))) || (opts.length === 1 ? opts[0] : null);
    const remember = o => { GM_setValue(SCHK, norm(o.text)); GM_setValue(SCHN, norm(o.text)); };
    if (!so) {
      status('🏫 اختر مدرستك من القائمة — مرة واحدة فقط، والأداة تتذكرها.');
      // نحفظ الاختيار فورًا لأن عتاد قد ينتقل للصفحة التالية مباشرة
      document.addEventListener('change', e => {
        if (e.target.tagName !== 'SELECT') return;
        const o = realOptions(e.target).find(x => x.value === e.target.value);
        if (o && !AR(o.text).includes(AR(DEPT))) remember(o);
      }, true);
      const v0 = sch.value;
      so = await waitFor(() => { const x = findSch(); return x && x.value !== v0 ? realOptions(x).find(o => o.value === x.value) || null : null; }, 300000);
      if (!so) { if (!leaving) status('⚠️ لم يتم اختيار المدرسة.'); return; }
    }
    remember(so);
    const st = load(); st.stage = 'menu'; save(st);
    const cur = findSch();
    if (cur && cur.value !== so.value) { await setSelect(cur, so.value); if (leaving) return; }
    await sleep(1500);
    if (!leaving && urlStage() === 'selplace') { const b = findButton(['دخول', 'موافق', 'انتقال', 'اختيار', 'متابعة']); if (b) clickEl(b); }
  }

  async function doMenu() {
    status('٣- تقارير المختبر ← عهدة الأصناف المخبرية');
    let el = findByText(REPORT);
    if (!el) { const m = await waitFor(() => findByText(MENU), 10000); if (m) { m.click(); await sleep(1500); } el = (await waitFor(() => findByText(REPORT), 5000)) || findByText(REPORT, true); }
    if (!el) { status(`⚠️ لم أجد «${REPORT}». افتحه يدويًا، والأداة تكمل.`); return; }
    const s = load(); s.stage = 'report'; save(s);
    clickEl(el);
  }

  const ST_OK = '✅ تم', ST_FAIL = '❌ لم يتم التنزيل', ST_EMPTY = '⚠️ لا توجد أصناف', ST_PART = '⚠️ تنزيل ناقص';
  const MAINW = ['النوع الرئيسي', 'رئيسي', 'Main'], SUBW = ['النوع الفرعي', 'فرعي', 'Sub'];
  const lkey = s => `${s.mi}|${s.si}`;
  function logSub(s, main, sub, st, note) {
    const k = lkey(s);
    s.log[k] = { main, sub, status: st, items: s.cnt[k] || 0, pages: s.pg[k] || 0, note: note || '' };
  }
  function bodyText() { return AR(document.body.innerText.replace(panel.innerText, '')); }
  const hasNoDataMsg = () => /لا توجد|لا يوجد|no data|no records/i.test(bodyText());
  function hardReload(s) { save(s); setTimeout(() => { location.href = location.pathname + location.search; }, 300); }
  // إعادة تحميل صفحة العهدة (حتى مرتين لكل مشكلة) ثم المتابعة من نفس المكان
  function tryRecover(s, k) {
    s.rec[k] = (s.rec[k] || 0) + 1;
    if (s.rec[k] > 2) return false;
    s.phase = 'sub'; s.target = 1; s.retry = 0; s.retryKey = ''; s.searchTry = 0; s.navTry = 0;
    status(`⚠️ عتاد لم يستجب — إعادة تحميل الصفحة (محاولة ${s.rec[k]})...`);
    hardReload(s);
    return true;
  }
  function nextSub(s) { s.si++; s.phase = 'sub'; s.target = 1; s.searchTry = 0; s.navTry = 0; save(s); }
  function nextMain(s) { s.mi++; s.si = 0; s.phase = 'sub'; s.target = 1; s.searchTry = 0; s.navTry = 0; save(s); }

  async function doReport() {
    let s = load();
    if (!s) return;
    s.log = s.log || {}; s.rec = s.rec || {}; s.cnt = s.cnt || {}; s.pg = s.pg || {};
    while (s && s.stage === 'report' && !stopFlag && !leaving) {
      const mainSel = await waitFor(() => findSelect(MAINW, 0), WAIT(), 'تحميل صفحة العهدة');
      if (stopFlag || leaving) return;
      if (!mainSel) {
        if (tryRecover(s, 'page')) return;
        s.log['page' + Date.now()] = { main: '—', sub: '—', status: ST_FAIL, items: 0, pages: 0, note: 'تعذر فتح صفحة العهدة في عتاد — توقف التشغيل هنا' };
        finish(s, true); return;
      }
      const mains = realOptions(mainSel);
      if (s.mi >= mains.length) { finish(s); return; }
      const m = mains[s.mi];

      // --- النوع الرئيسي ---
      if (mainSel.value !== m.value) {
        if (bump(s, 'm' + s.mi)) {
          if (tryRecover(s, 'm' + s.mi)) return;
          s.log['m' + s.mi] = { main: m.text, sub: '(كل الفروع)', status: ST_FAIL, items: 0, pages: 0, note: 'تعذر فتح النوع الرئيسي' };
          nextMain(s); continue;
        }
        status(`النوع الرئيسي ${s.mi + 1}/${mains.length}: ${m.text}`);
        s.si = 0; s.phase = 'sub'; save(s);
        await setSelect(mainSel, m.value); if (leaving) return; s = load(); continue;
      }

      // --- الأنواع الفرعية ---
      const subSel = await waitFor(() => { const x = findSelect(SUBW, 1); return x && x !== mainSel && realOptions(x).length ? x : null; }, WAIT(), `تحميل الأنواع الفرعية: ${m.text}`);
      if (stopFlag || leaving) return;
      const subs = subSel ? realOptions(subSel) : [];
      if (!subs.length) {
        if (tryRecover(s, 'ms' + s.mi)) return;
        s.log['m' + s.mi] = { main: m.text, sub: '(كل الفروع)', status: ST_FAIL, items: 0, pages: 0, note: 'لم تظهر الأنواع الفرعية لهذا النوع' };
        nextMain(s); continue;
      }
      if (s.si >= subs.length) { nextMain(s); continue; }
      const sb = subs[s.si];
      const k = lkey(s);
      const tag = `${s.mi + 1}/${mains.length} ${m.text} ← ${s.si + 1}/${subs.length} ${sb.text}`;

      if (subSel.value !== sb.value) {
        if (bump(s, 's' + k)) {
          if (tryRecover(s, k)) return;
          logSub(s, m.text, sb.text, ST_FAIL, 'تعذر اختيار النوع الفرعي'); nextSub(s); continue;
        }
        s.phase = 'sub'; save(s);
        await setSelect(subSel, sb.value); if (leaving) return; s = load(); continue;
      }

      // --- البحث ---
      if (s.phase === 'sub') {
        const btn = findButton(['بحث', 'عرض', 'Search']);
        if (!btn) {
          if (tryRecover(s, k)) return;
          logSub(s, m.text, sb.text, ST_FAIL, 'لم يظهر زر البحث'); nextSub(s); continue;
        }
        status(`🔎 بحث: ${tag}`);
        const sig = gridSig();
        s.phase = 'page'; s.target = 1; s.searchSig = sig; save(s);
        clickEl(btn);
        const ok = await waitFor(() => gridSig() !== sig, WAIT(), `🔎 بحث: ${tag}`);
        if (stopFlag || leaving) return;
        s = load();
        if (!ok) {
          s.searchTry = (s.searchTry || 0) + 1;
          if (s.searchTry <= 1) { s.phase = 'sub'; save(s); continue; } // إعادة البحث مرة
          if (hasNoDataMsg()) { logSub(s, m.text, sb.text, ST_EMPTY, 'لا توجد أصناف في هذا الفرع'); nextSub(s); continue; }
          if (tryRecover(s, k)) return;
          logSub(s, m.text, sb.text, ST_FAIL, 'لم يستجب عتاد عند البحث'); nextSub(s); continue;
        }
        s.searchTry = 0; save(s);
        await sleep(800); continue;
      }

      // --- الصفحات ---
      const grid = findGrid();
      if (!grid || !dataRows(grid).length) {
        if (s.pg[k]) { logSub(s, m.text, sb.text, ST_PART, `اختفى الجدول بعد الصفحة ${s.target - 1}`); nextSub(s); continue; }
        logSub(s, m.text, sb.text, ST_EMPTY, hasNoDataMsg() ? 'لا توجد أصناف في هذا الفرع' : 'لم تظهر نتائج (قد يكون الفرع فارغًا أو لم يستجب عتاد)');
        nextSub(s); continue;
      }
      if (!s.pg[k] && s.searchSig && s.searchSig !== 'none' && gridSig() === s.searchSig) {
        // الجدول ما زال جدول الفرع السابق
        s.searchTry = (s.searchTry || 0) + 1;
        if (s.searchTry <= 1) { s.phase = 'sub'; save(s); continue; }
        if (tryRecover(s, k)) return;
        logSub(s, m.text, sb.text, ST_FAIL, 'لم يتحدث الجدول بعد البحث'); nextSub(s); continue;
      }
      const pager = getPager();
      const cur = (pager && pager.current) || s.target || 1;
      const totalPages = pager ? Math.max(...pager.items.map(i => i.n)) : 1;
      const key = `${k}|${cur}`;
      if (!s.done.includes(key)) {
        if (s.mode === 'export') {
          const ex = findButton(['تصدير', 'Excel', 'اكسل', 'Export']);
          if (!ex) { logSub(s, m.text, sb.text, ST_FAIL, 'لم يظهر زر التصدير'); nextSub(s); continue; }
          s.done.push(key); s.files++; s.pg[k] = (s.pg[k] || 0) + 1; save(s);
          clickEl(ex);
          await sleep(4000);
          leaving = false; // تنزيل الملف لا يغيّر الصفحة
        } else {
          const { headers, rows } = readGrid(grid);
          if (!s.headers && headers) s.headers = ['النوع الرئيسي', 'النوع الفرعي', ...headers];
          rows.forEach(r => s.rows.push([m.text, sb.text, ...r]));
          s.cnt[k] = (s.cnt[k] || 0) + rows.length;
          s.pg[k] = (s.pg[k] || 0) + 1;
          s.done.push(key); save(s);
        }
      }
      const fails = Object.values(s.log).filter(x => x.status !== ST_OK).length;
      status(`${tag}\nصفحة ${cur}${pager ? '' : ' (صفحة واحدة)'} — ${s.mode === 'export' ? 'ملفات: ' + s.files : 'أصناف: ' + s.rows.length}${fails ? `\n⚠️ فروع بها مشكلة: ${fails}` : ''}`);

      const next = pager && findNext(pager, cur);
      if (!next) { logSub(s, m.text, sb.text, ST_OK, ''); s.rec.page = 0; nextSub(s); continue; }
      s.target = cur + 1; save(s);
      const sig = gridSig();
      clickEl(next);
      const moved = await waitFor(() => gridSig() !== sig, WAIT(), `${tag}\nفتح الصفحة ${cur + 1}`);
      if (stopFlag || leaving) return;
      s = load();
      if (!moved) {
        s.navTry = (s.navTry || 0) + 1;
        if (s.navTry <= 1) { save(s); continue; } // إعادة الضغط على الرقم مرة
        if (tryRecover(s, k + 'n')) return;
        logSub(s, m.text, sb.text, ST_PART, `توقف عند الصفحة ${cur} — لم تفتح الصفحة ${cur + 1} (آخر رقم ظاهر ${totalPages})`);
        nextSub(s); continue;
      }
      s.navTry = 0; save(s);
      await sleep(800);
    }
  }
  function bump(s, k) { // حماية من التكرار اللانهائي
    if (s.retryKey === k) s.retry++; else { s.retryKey = k; s.retry = 0; }
    save(s);
    return s.retry > 3;
  }

  function finish(s, partial) {
    clear();
    const log = Object.values(s.log || {});
    const bad = log.filter(x => x.status !== ST_OK);
    const d = { headers: s.headers, rows: s.rows || [], date: today(), log, partial: !!partial, mode: s.mode, files: s.files };
    GM_setValue(LAST, d);
    exportExcel(d);
    const head = partial ? '⏹ توقف قبل الاكتمال' : '✅ انتهى';
    const what = s.mode === 'export' ? `ملفات التصدير: ${s.files}` : `الأصناف: ${d.rows.length}`;
    status(`${head} — ${what}\nفروع تمت: ${log.length - bad.length}${bad.length ? `\n❌ فروع بها مشكلة: ${bad.length} (انظر ورقة «تقرير التنزيل»)` : ''}`);
    hideTimer();
    showDoneBubble(d, log.length - bad.length, bad.length, !!partial);
  }

  // ================= فقاعة انتهاء التصدير الفاخرة =================
  function chime() {
    try {
      const ac = new (window.AudioContext || window.webkitAudioContext)();
      [523, 659, 784, 1047].forEach((f, i) => {
        const o = ac.createOscillator(), g = ac.createGain();
        o.type = 'sine'; o.frequency.value = f;
        g.gain.setValueAtTime(0.0001, ac.currentTime + i * 0.15);
        g.gain.exponentialRampToValueAtTime(0.25, ac.currentTime + i * 0.15 + 0.02);
        g.gain.exponentialRampToValueAtTime(0.0001, ac.currentTime + i * 0.15 + 0.5);
        o.connect(g).connect(ac.destination); o.start(ac.currentTime + i * 0.15); o.stop(ac.currentTime + i * 0.15 + 0.55);
      });
    } catch (e) {}
  }
  function showDoneBubble(d, okCount, badCount, partial) {
    const old = document.getElementById('ax-done'); if (old) old.remove();
    const clean = !badCount && !partial;
    const accent = clean ? '#10b981' : '#f59e0b';
    const title = partial ? 'توقف تصدير العهدة قبل الاكتمال' : clean ? 'اكتمل تصدير عهدة المختبر بنجاح' : 'اكتمل تصدير العهدة — مع ملاحظات';
    const what = d.mode === 'export' ? `${d.files} ملف تصدير` : `${d.rows.length} صنفًا`;
    const box = document.createElement('div');
    box.id = 'ax-done'; box.dir = 'rtl';
    box.innerHTML = `
      <style>
        #ax-done {
          position: fixed;
          inset: 0;
          z-index: 2147483647;
          background: rgba(12, 10, 9, 0.72);
          backdrop-filter: blur(6px);
          display: flex;
          align-items: center;
          justify-content: center;
          font-family: 'Cairo', Tahoma, Arial, sans-serif;
          animation: axFade .3s ease;
        }
        #ax-done * { box-sizing: border-box; font-family: 'Cairo', Tahoma, Arial, sans-serif; }
        #ax-done .card {
          position: relative;
          background: #1c1917;
          color: #f5f5f4;
          border-radius: 26px;
          padding: 32px 28px 22px;
          width: min(410px, 92vw);
          text-align: center;
          box-shadow: 0 25px 70px rgba(0,0,0,0.65);
          border: 1px solid #292524;
          border-top: 6px solid ${accent};
          animation: axPop .45s cubic-bezier(.2,1.4,.4,1);
        }
        #ax-done .flask {
          position: relative;
          width: 88px;
          height: 88px;
          margin: -74px auto 12px;
          background: #0c0a09;
          border: 3px solid ${accent};
          border-radius: 24px;
          display: grid;
          place-items: center;
          font-size: 46px;
          box-shadow: 0 12px 28px rgba(0,0,0,0.45);
          animation: axBob 1.8s ease-in-out infinite;
        }
        #ax-done .b {
          position: absolute;
          bottom: 62px;
          border-radius: 50%;
          background: ${clean ? '#34d399' : '#fbbf24'};
          opacity: 0;
          animation: axRise 2.4s ease-in infinite;
        }
        #ax-done h2 { margin: 4px 0; font-size: 19px; font-weight: 900; color: #ffffff; }
        #ax-done p.sch { margin: 2px 0 12px; color: #fbbf24; font-size: 13.5px; font-weight: 700; }
        #ax-done .stats { display: flex; gap: 8px; justify-content: center; margin: 14px 0 10px; }
        #ax-done .st {
          flex: 1;
          background: #0c0a09;
          border: 1px solid #292524;
          border-radius: 14px;
          padding: 10px 6px;
          font-size: 12px;
          color: #a8a29e;
          font-weight: 700;
        }
        #ax-done .st b { display: block; font-size: 20px; font-weight: 900; color: #ffffff; margin-bottom: 2px; }
        #ax-done .st.bad b { color: #f59e0b; }
        #ax-done .note {
          font-size: 12px;
          padding: 8px 12px;
          border-radius: 10px;
          background: rgba(245, 158, 11, 0.1);
          color: #fde68a;
          margin: 10px 0;
        }
        #ax-done .btns { display: flex; gap: 8px; margin-top: 16px; }
        #ax-done button {
          flex: 1;
          border: 1px solid #3f3a36;
          border-radius: 12px;
          padding: 10px 8px;
          font-weight: 800;
          font-size: 13px;
          cursor: pointer;
          background: #292524;
          color: #e7e5e4;
          transition: 0.2s;
        }
        #ax-done button:hover { background: #3f3a36; }
        #ax-done button.main {
          background: linear-gradient(135deg, #f59e0b, #ea580c);
          color: #0c0a09;
          border-color: transparent;
          font-weight: 900;
        }
        #ax-done .credit {
          margin-top: 14px;
          padding-top: 10px;
          border-top: 1px solid #292524;
          font-size: 11px;
          color: #78716c;
          font-weight: 600;
        }
        @keyframes axFade { from { opacity: 0; } }
        @keyframes axPop { from { transform: scale(.7); opacity: 0; } }
        @keyframes axBob { 50% { transform: translateY(-5px) rotate(-4deg); } }
        @keyframes axRise { 0% { transform: translateY(0) scale(.6); opacity: 0; } 20% { opacity: .85; } 100% { transform: translateY(-110px) scale(1.2); opacity: 0; } }
      </style>
      <div class="card">
        <span class="b" style="left:44%;width:10px;height:10px;animation-delay:0s"></span>
        <span class="b" style="left:52%;width:7px;height:7px;animation-delay:.6s"></span>
        <span class="b" style="left:48%;width:12px;height:12px;animation-delay:1.2s"></span>
        <span class="b" style="left:56%;width:8px;height:8px;animation-delay:1.8s"></span>
        <div class="flask">${clean ? '🧪' : '⚗️'}</div>
        <h2>${title}</h2>
        <p class="sch">${esc(GM_getValue(SCHN, ''))}</p>
        <div class="stats">
          <div class="st"><b>${what.split(' ')[0]}</b>${what.split(' ').slice(1).join(' ')}</div>
          <div class="st"><b>${okCount}</b>فرع تم تنزيله</div>
          <div class="st ${badCount ? 'bad' : ''}"><b>${badCount}</b>فرع به ملاحظة</div>
        </div>
        <div class="note">
          ${badCount || partial ? '⚠️ راجع ورقة «تقرير التنزيل» في ملف الإكسل للتفاصيل.' : '✅ تم حفظ ملف الإكسل الكامل في مجلد التنزيلات.'}
        </div>
        <div class="btns">
          ${d.rows.length ? '<button class="main" data-x="pdf">🖨️ كشف PDF</button><button data-x="xlsx">📥 تنزيل Excel</button>' : ''}
          <button data-x="close">إغلاق</button>
        </div>
        <div class="credit">تصميم وتطوير: سيف جمال المعولي • معولة بن شمس (5-12)</div>
      </div>`;
    box.addEventListener('click', e => {
      const btn = e.target.closest('button[data-x]');
      const x = btn && btn.dataset.x;
      if (e.target === box || x === 'close') box.remove();
      if (x === 'pdf') printPDF(d);
      if (x === 'xlsx') exportExcel(d);
    });
    document.body.appendChild(box);
    chime();
    try { GM_notification({ title: '🧪 ' + title, text: `${what} — ${okCount} فرع تم${badCount ? ` — ${badCount} فرع به مشكلة` : ''}`, timeout: 15000 }); } catch (e) {}
  }

  // ================= التصدير =================
  function sheet(head, rows) {
    const ws = XLSX.utils.aoa_to_sheet([head, ...rows]);
    ws['!cols'] = head.map((h, i) => ({ wch: Math.min(45, Math.max(8, ...[h, ...rows.slice(0, 300).map(r => r[i])].map(v => String(v ?? '').length + 2))) }));
    ws['!autofilter'] = { ref: ws['!ref'] };
    return ws;
  }
  function reportSheet(d) {
    const log = d.log || [];
    const bad = log.filter(x => x.status !== ST_OK);
    const aoa = [
      ['تقرير تنزيل عهدة الأصناف المخبرية', '', '', '', '', '', ''],
      ['المدرسة', GM_getValue(SCHN, '')],
      ['التاريخ', d.date],
      ['فروع تم تنزيلها', log.length - bad.length],
      ['فروع لم يتم تنزيلها أو ناقصة أو فارغة', bad.length],
    ];
    if (d.partial) aoa.push(['⏹ تنبيه', 'تم إيقاف التشغيل قبل المرور على كل الفروع — الفروع غير المذكورة لم تُفحص']);
    if (bad.length) aoa.push(['❌ تنبيه', 'توجد فروع لم يتم تنزيلها بالكامل — أعد تشغيل الأداة لها أو نزّلها يدويًا']);
    aoa.push([], ['م', 'النوع الرئيسي', 'النوع الفرعي', 'الحالة', 'عدد الأصناف', 'عدد الصفحات', 'ملاحظة']);
    // الفروع ذات المشكلة أولًا ثم الناجحة
    [...bad, ...log.filter(x => x.status === ST_OK)].forEach((x, i) =>
      aoa.push([i + 1, x.main, x.sub, x.status, d.mode === 'export' ? '—' : x.items, x.pages, x.note]));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 30 }, { wch: 28 }, { wch: 28 }, { wch: 18 }, { wch: 12 }, { wch: 12 }, { wch: 50 }];
    return { ws, bad: bad.length };
  }
  function exportExcel(d) {
    const wb = XLSX.utils.book_new();
    wb.Workbook = { Views: [{ RTL: true }] };
    const rep = reportSheet(d);
    if (rep.bad || d.partial || !d.rows.length) XLSX.utils.book_append_sheet(wb, rep.ws, 'تقرير التنزيل');
    if (d.rows.length) {
      const head = d.headers || (d.rows[0] || []).map((_, i) => 'عمود ' + (i + 1));
      XLSX.utils.book_append_sheet(wb, sheet(head, d.rows), 'كل العهدة');
      const used = new Set(['كل العهدة', 'تقرير التنزيل']);
      const byMain = new Map();
      d.rows.forEach(r => { if (!byMain.has(r[0])) byMain.set(r[0], []); byMain.get(r[0]).push(r); });
      for (const [name, rows] of byMain) {
        let n = String(name).replace(/[\\\/?*\[\]:]/g, ' ').slice(0, 28).trim() || 'نوع';
        let k = n, i = 2; while (used.has(k)) k = n + ' ' + i++;
        used.add(k);
        XLSX.utils.book_append_sheet(wb, sheet(head, rows), k);
      }
    }
    if (!(rep.bad || d.partial || !d.rows.length)) XLSX.utils.book_append_sheet(wb, rep.ws, 'تقرير التنزيل');
    const suffix = rep.bad || d.partial ? '_ناقص' : '';
    const base = d.mode === 'export' ? 'تقرير_تنزيل_العهدة' : 'عهدة_المختبر';
    XLSX.writeFile(wb, `${base}_${d.date}${suffix}.xlsx`);
  }

  function printPDF(d) {
    const head = d.headers || (d.rows[0] || []).map((_, i) => 'عمود ' + (i + 1));
    const w = window.open('', '_blank');
    if (!w) { status('⚠️ المتصفح منع النافذة. اسمح بالنوافذ المنبثقة لهذا الموقع.'); return; }
    w.document.write(`<!doctype html><html dir="rtl" lang="ar"><head><meta charset="utf-8"><title>كشف عهدة المختبر ${d.date}</title>
      <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap" rel="stylesheet">
      <style>@page{size:A4 landscape;margin:10mm}body{font-family:'Cairo',Tahoma,Arial;font-size:11px;color:#1c1917}
      .top{display:flex;justify-content:space-between;border-bottom:2px solid #f59e0b;padding-bottom:8px;margin-bottom:10px;font-weight:700}
      h1{text-align:center;font-size:18px;font-weight:800;margin:6px 0 12px;color:#0c0a09}
      table{width:100%;border-collapse:collapse}th,td{border:1px solid #57534e;padding:4px 6px;text-align:center}
      th{background:#f5f5f4;font-weight:800}thead{display:table-header-group}tr{page-break-inside:avoid}
      .sign{display:flex;justify-content:space-around;margin-top:36px;font-weight:700}.sign div{text-align:center;width:30%}</style></head><body>
      <div class="top"><div>سلطنة عُمان<br>وزارة التعليم<br>${esc(GM_getValue(SCHN, ''))}</div><div>التاريخ: ${esc(d.date)}<br>إجمالي الأصناف: ${d.rows.length}</div></div>
      <h1>كشف عهدة الأصناف المخبرية</h1>
      <table><thead><tr><th>م</th>${head.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead>
      <tbody>${d.rows.map((r, i) => `<tr><td>${i + 1}</td>${head.map((_, j) => `<td>${esc(r[j])}</td>`).join('')}</tr>`).join('')}</tbody></table>
      ${(() => { const b = (d.log || []).filter(x => x.status !== ST_OK); return b.length || d.partial ? `<h3 style="color:#b91c1c;margin-top:16px">تنبيه: فروع لم يتم تنزيلها أو ناقصة${d.partial ? ' (وتم إيقاف التشغيل قبل الاكتمال)' : ''}</h3>
      <table><thead><tr><th>النوع الرئيسي</th><th>النوع الفرعي</th><th>الحالة</th><th>ملاحظة</th></tr></thead><tbody>${b.map(x => `<tr><td>${esc(x.main)}</td><td>${esc(x.sub)}</td><td>${esc(x.status)}</td><td>${esc(x.note)}</td></tr>`).join('')}</tbody></table>` : ''; })()}
      <div class="sign"><div>فني المختبر<br><br>الاسم: ................<br><br>التوقيع: ............</div>
      <div>مدير المدرسة<br><br>الاسم: ................<br><br>التوقيع: ............</div></div>
      <script>window.onload=()=>setTimeout(()=>window.print(),400)<\/script></body></html>`);
    w.document.close();
  }

  function diagnose() {
    const sels = [...document.querySelectorAll('select')].filter(s => !panel.contains(s) && isVisible(s));
    const g = findGrid(), p = getPager();
    const lines = [
      'المرحلة حسب الرابط: ' + (urlStage() || 'غير معروفة'),
      'القوائم: ' + (sels.map((s, i) => `[${i + 1}] ${labelFor(s).slice(-25)} (${realOptions(s).length} خيار)`).join(' | ') || 'لا يوجد'),
      'الجدول: ' + (g ? dataRows(g).length + ' صف' : 'لا يوجد'),
      'الصفحات: ' + (p ? p.items.map(i => i.n).join(',') + ' الحالية ' + p.current : 'لا يوجد'),
      'زر البحث: ' + (findButton(['بحث', 'عرض', 'Search']) ? 'موجود' : 'غير موجود'),
      'زر التصدير: ' + (findButton(['تصدير', 'Excel', 'اكسل', 'Export']) ? 'موجود' : 'غير موجود'),
    ];
    status(lines.join('\n'));
  }

  // استئناف تلقائي بعد كل تحميل للصفحة
  if (load()) setTimeout(dispatch, 1500);
})();
