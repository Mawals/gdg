// ==UserScript==
// @name         تصدير عهدة المختبر - عتاد
// @namespace    mulawah-school
// @author       سيف جمال المعولي
// @updateURL    https://mawals.github.io/gdg/atad.user.js
// @downloadURL  https://mawals.github.io/gdg/atad.user.js
// @version      3.2
// @description  أعمالي ← عتاد ← المدارس ← المدرسة ← تقارير المختبر ← عهدة الأصناف المخبرية، ثم يمر على كل نوع رئيسي وفرعي وكل الصفحات ويجمعها في Excel واحد (أو يضغط زر التصدير لكل صفحة)
// @match        *://*.moe.gov.om/*
// @include      /^https?:\/\/([^\/]*\.)?moe\.gov\.om(:\d+)?\/.*$/
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_deleteValue
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';
  if (window.top !== window.self) return;

  // ================= الإعدادات =================
  // المدرسة: تُكتب في خانة «المدرسة» في اللوحة، أو تُحفظ تلقائيًا أول مرة يختارها الفني
  const SCHK = 'atad2_school', SCHN = 'atad2_schoolname';
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

  // ================= اللوحة =================
  const panel = document.createElement('div');
  panel.dir = 'rtl';
  panel.style.cssText = 'position:fixed;bottom:16px;left:16px;z-index:2147483647;background:#fff;border:2px solid #0f766e;border-radius:12px;padding:10px;width:270px;font:13px Tahoma,Arial;box-shadow:0 4px 18px rgba(0,0,0,.25);color:#111';
  panel.innerHTML = `
    <div style="font-weight:bold;color:#0f766e;margin-bottom:6px;display:flex;justify-content:space-between">
      <span>📦 تصدير عهدة المختبر</span><span id="ax-min" style="cursor:pointer">—</span></div>
    <div id="ax-body">
      <button data-a="start">▶️ ابدأ</button>
      <label style="display:block;margin:4px 0"><input type="radio" name="ax-mode" value="collect"> تجميع الكل في ملف Excel واحد</label>
      <label style="display:block;margin:4px 0"><input type="radio" name="ax-mode" value="export"> الضغط على زر التصدير لكل صفحة</label>
      <label style="display:block;margin:4px 0">🏫 المدرسة: <input type="text" id="ax-school" placeholder="جزء من اسم المدرسة" style="width:150px"></label>
      <label style="display:block;margin:4px 0">⏱️ انتظار عتاد حتى <input type="number" id="ax-wait" min="20" max="600" style="width:55px"> ثانية</label>
      <button data-a="stop" style="background:#b91c1c">⏹ إيقاف (وتصدير ما جُمع)</button>
      <button data-a="xlsx">📥 تنزيل Excel لآخر نتيجة</button>
      <button data-a="pdf">🖨️ كشف PDF / طباعة</button>
      <button data-a="diag" style="background:#475569">🔍 فحص الصفحة</button>
      <div id="ax-status" style="margin-top:6px;padding:6px;background:#f0fdfa;border-radius:6px;min-height:18px;white-space:pre-line">جاهز.</div>
      <div style="margin-top:6px;text-align:center;font-size:11px;color:#64748b">تصميم: سيف جمال المعولي</div>
    </div>`;
  panel.querySelectorAll('button').forEach(b => b.style.cssText += ';display:block;width:100%;margin:3px 0;padding:6px;border:0;border-radius:6px;background:#0f766e;color:#fff;cursor:pointer;font:13px Tahoma');
  document.body.appendChild(panel);
  const status = t => (panel.querySelector('#ax-status').textContent = t);
  panel.querySelector('#ax-min').onclick = () => { const b = panel.querySelector('#ax-body'); b.style.display = b.style.display === 'none' ? '' : 'none'; };
  const mode = GM_getValue(MODE, 'collect');
  panel.querySelector(`input[value=${mode}]`).checked = true;
  panel.querySelectorAll('input[name=ax-mode]').forEach(r => r.onchange = () => GM_setValue(MODE, r.value));
  const schoolBox = panel.querySelector('#ax-school');
  schoolBox.value = GM_getValue(SCHK, '');
  schoolBox.onchange = () => GM_setValue(SCHK, norm(schoolBox.value));
  const waitBox = panel.querySelector('#ax-wait');
  waitBox.value = WAIT() / 1000;
  waitBox.onchange = () => { const v = Math.min(600, Math.max(20, +waitBox.value || 90)); waitBox.value = v; GM_setValue(WAITK, v); };

  panel.addEventListener('click', e => {
    const a = e.target.dataset && e.target.dataset.a;
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
  async function waitFor(fn, ms, label) {
    const t0 = Date.now(), end = t0 + ms;
    let lastSec = -1;
    while (Date.now() < end && !stopFlag && !leaving) {
      const v = fn(); if (v) return v;
      if (label) { const sec = Math.round((Date.now() - t0) / 1000); if (sec !== lastSec) { lastSec = sec; status(`${label}\n⏳ بانتظار عتاد: ${sec} / ${Math.round(ms / 1000)} ثانية`); } }
      await sleep(400);
    }
    return null;
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
    const dept = await waitFor(() => [...document.querySelectorAll('select')].find(x => [...x.options].some(o => AR(o.text) === AR(DEPT))), 15000);
    if (!dept) { status(`⚠️ لم أجد قائمة القسم «${DEPT}».`); return; }
    const opt = [...dept.options].find(o => AR(o.text) === AR(DEPT));
    if (dept.value !== opt.value) { await setSelect(dept, opt.value); }
    const sch = await waitFor(() => [...document.querySelectorAll('select')].find(x => x !== dept && isVisible(x) && realOptions(x).length), WAIT(), 'تحميل قائمة المدارس');
    if (!sch) { status('⚠️ لم تظهر قائمة المدارس.'); return; }
    const opts = realOptions(sch);
    const key = GM_getValue(SCHK, '');
    let so = key ? opts.find(o => AR(o.text).includes(AR(key))) : null;
    if (!so && opts.length === 1) so = opts[0];
    if (!so && opts.some(o => o.value === sch.value)) so = opts.find(o => o.value === sch.value);
    if (!so) {
      status('🏫 اختر مدرستك من القائمة مرة واحدة، والأداة ستتذكرها وتكمل.');
      const v0 = sch.value;
      // نحفظ الاختيار فورًا لأن عتاد قد ينتقل للصفحة التالية مباشرة بعد الاختيار
      sch.addEventListener('change', () => { const o = opts.find(x => x.value === sch.value); if (o) { GM_setValue(SCHK, norm(o.text)); GM_setValue(SCHN, norm(o.text)); } }, true);
      const picked = await waitFor(() => (sch.value !== v0 && opts.find(o => o.value === sch.value)) || null, 300000);
      if (!picked) { status('⚠️ لم يتم اختيار المدرسة.'); return; }
      so = picked;
    }
    GM_setValue(SCHK, norm(so.text)); GM_setValue(SCHN, norm(so.text)); schoolBox.value = norm(so.text);
    const s = load(); s.stage = 'menu'; save(s);
    if (sch.value !== so.value) await setSelect(sch, so.value);
    await sleep(1500);
    if (urlStage() === 'selplace') { const b = findButton(['دخول', 'موافق', 'انتقال', 'اختيار', 'متابعة']); if (b) clickEl(b); }
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
      <style>@page{size:A4 landscape;margin:10mm}body{font-family:Tahoma,Arial;font-size:11px}
      .top{display:flex;justify-content:space-between}h1{text-align:center;font-size:17px;margin:6px 0 10px}
      table{width:100%;border-collapse:collapse}th,td{border:1px solid #444;padding:3px 5px;text-align:center}
      th{background:#e5e7eb}thead{display:table-header-group}tr{page-break-inside:avoid}
      .sign{display:flex;justify-content:space-around;margin-top:36px}.sign div{text-align:center;width:30%}</style></head><body>
      <div class="top"><div>سلطنة عُمان<br>وزارة التعليم<br>${esc(GM_getValue(SCHN, ''))}</div><div>التاريخ: ${esc(d.date)}<br>عدد الأصناف: ${d.rows.length}</div></div>
      <h1>كشف عهدة الأصناف المخبرية</h1>
      <table><thead><tr><th>م</th>${head.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead>
      <tbody>${d.rows.map((r, i) => `<tr><td>${i + 1}</td>${head.map((_, j) => `<td>${esc(r[j])}</td>`).join('')}</tr>`).join('')}</tbody></table>
      ${(() => { const b = (d.log || []).filter(x => x.status !== ST_OK); return b.length || d.partial ? `<h3 style="color:#b91c1c;margin-top:16px">تنبيه: فروع لم يتم تنزيلها أو ناقصة${d.partial ? ' (وتم إيقاف التشغيل قبل الاكتمال)' : ''}</h3>
      <table><thead><tr><th>النوع الرئيسي</th><th>النوع الفرعي</th><th>الحالة</th><th>ملاحظة</th></tr></thead><tbody>${b.map(x => `<tr><td>${esc(x.main)}</td><td>${esc(x.sub)}</td><td>${esc(x.status)}</td><td>${esc(x.note)}</td></tr>`).join('')}</tbody></table>` : ''; })()}
      <div class="sign"><div>فني المختبر<br><br>الاسم: ................<br><br>التوقيع: ............</div>
      <div>مدير المدرسة<br><br>الاسم: ................<br><br>التوقيع: ............</div></div>
      <script>window.onload=()=>setTimeout(()=>window.print(),300)<\/script></body></html>`);
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
