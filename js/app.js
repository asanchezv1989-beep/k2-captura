import * as db from './db.js';
import {
  state, TM, EQLOG, KINDS, orderOf, val, baseVal, shown, setVal, isFormula, isEdited, num,
  workers, namedWorkers, addWorker, days, dayType, holidays, cellOf, hoursByOrder, split, crewOf,
  buildCapture, reconcile,
} from './model.js';

const $ = (s, el = document) => el.querySelector(s);
const view = $('#view');
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const fmt = (n) => (n === null || n === undefined || n === '' ? '' : String(+(+n).toFixed(2)));
const MES = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
const DIA = { Mon: 'Lun', Tue: 'Mar', Wed: 'Mié', Thu: 'Jue', Fri: 'Vie', Sat: 'Sáb', Sun: 'Dom' };
const dLabel = (d) => { const [, m, dd] = d.date.split('-'); return `${DIA[d.dow]} ${+dd} ${MES[+m - 1]}`; };
const todayIso = () => { const t = new Date(); return `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, '0')}-${String(t.getDate()).padStart(2, '0')}`; };

const ui = { tab: 'hours', day: 0, order: null, mode: 'order', person: null, extra: {}, undo: null };

// ------------------------------------------------------------- persistencia
let saveT;
function persist() {
  clearTimeout(saveT);
  saveT = setTimeout(() => db.set('edits', [...state.edits.values()]), 250);
  db.set('ui', { tab: ui.tab, day: ui.day, order: ui.order, mode: ui.mode, person: ui.person });
  renderTop();
}

function toast(msg, undo) {
  const t = $('#toast');
  t.innerHTML = esc(msg) + (undo ? '<button>Deshacer</button>' : '');
  t.hidden = false;
  if (undo) t.querySelector('button').onclick = () => { undo(); t.hidden = true; };
  clearTimeout(t._h);
  t._h = setTimeout(() => (t.hidden = true), undo ? 6000 : 2600);
}

function snapshot() {
  const copy = new Map(state.edits);
  return () => { state.edits = copy; persist(); render(); };
}

// ------------------------------------------------------------- arranque
async function init() {
  if ('serviceWorker' in navigator && location.protocol !== 'file:') navigator.serviceWorker.register('sw.js').catch(() => {});
  const base = await db.get('base');
  if (base) {
    state.base = base;
    for (const e of (await db.get('edits')) || []) state.edits.set(`${e.s}!${e.a}`, e);
    Object.assign(ui, (await db.get('ui')) || {});
    fixUi();
  }
  $('#tabs').addEventListener('click', (e) => { const b = e.target.closest('[data-tab]'); if (b) go(b.dataset.tab); });
  $('#hdr-dirty').onclick = () => go('send');
  $('#file').addEventListener('change', onFile);
  render();
}

function fixUi() {
  const ds = days();
  const t = todayIso();
  if (ui._fresh || !(ui.day >= 0 && ui.day < ds.length)) {
    let i = ds.findIndex((d) => d.date === t);
    if (i < 0) i = ds[ds.length - 1].date < t ? ds.length - 1 : 0;
    ui.day = i;
  }
  if (!ui.order || !orderOf(ui.order)) {
    const busy = state.base.orders.filter((o) => o.visible).sort((a, b) => Object.keys(b.cells).length - Object.keys(a.cells).length);
    ui.order = (busy[0] || state.base.orders[0]).sheet;
  }
  delete ui._fresh;
}

function go(tab) {
  if (!state.base) return;
  ui.tab = tab;
  persist();
  render();
  window.scrollTo(0, 0);
}

// ------------------------------------------------------------- cargar base
async function onFile(e) {
  const f = e.target.files[0];
  e.target.value = '';
  if (!f) return;
  let data;
  try { data = JSON.parse(await f.text()); } catch { return toast('Ese archivo no es un JSON válido.'); }
  if (data.app === 'k2-captura') return toast('Ese es un archivo de captura; ese se suelta en la PC, no aquí.');
  if (data.app !== 'k2-base') return toast('Ese archivo no es una base de K2 Captura.');
  if (state.base && state.base.file.name !== data.file.name && state.edits.size) {
    if (!confirm(`La base es de otro libro (${data.file.name}). Se borrarán ${state.edits.size} cambios sin enviar. ¿Continuar?`)) return;
    state.edits.clear();
  }
  const first = !state.base;
  const before = state.edits.size;
  state.base = data;
  reconcile();
  await db.set('base', data);
  await db.set('edits', [...state.edits.values()]);
  if (first) ui._fresh = true;
  fixUi();
  persist();
  render();
  const gone = before - state.edits.size;
  toast(gone ? `Base cargada. ${gone} cambios ya están en el Excel.` : 'Base cargada.');
}

// ------------------------------------------------------------- render general
function renderTop() {
  const b = state.base;
  $('#hdr-job').textContent = b ? `${b.project.customer || ''} · ${b.project.number || ''}` : 'K2 Captura';
  $('#hdr-sub').textContent = b ? `Base del ${b.stamp.slice(0, 16).replace('T', ' ')}` : 'Sin base';
  const n = state.edits.size;
  const d = $('#hdr-dirty');
  d.hidden = !n;
  d.textContent = `${n} sin enviar`;
  for (const t of document.querySelectorAll('#tabs button')) {
    t.classList.toggle('on', t.dataset.tab === ui.tab);
    t.disabled = !b;
  }
}

function render() {
  renderTop();
  if (!state.base) return renderWelcome();
  ({ hours: renderHours, equip: renderEquip, log: renderLog, send: renderSend })[ui.tab]();
}

function renderWelcome() {
  view.innerHTML = `
    <h2>K2 Captura</h2>
    <p class="muted">Captura horas, equipo y log de equipo en el celular. Al final la PC escribe todo en tu Excel con Excel mismo, sin tocar ninguna fórmula.</p>
    <div class="drop" id="pick">
      <svg viewBox="0 0 24 24"><path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6M12 18v-6M9 15l3-3 3 3"/></svg>
      <p><b>Cargar base</b><br><span class="small muted">el archivo <i>.base.json</i> que hizo la PC</span></p>
    </div>
    <h3>Cómo se usa</h3>
    <div class="card pad"><ol class="steps">
      <li>En la PC suelta el <b>.xlsb</b> sobre el icono <b>K2 Captura</b>. Sale un <i>.base.json</i> junto al Excel.</li>
      <li>Pásalo al celular (OneDrive, AirDrop o correo) y cárgalo aquí.</li>
      <li>Captura. Todo se guarda en el celular, aunque no haya señal.</li>
      <li>En <b>Enviar</b> exporta la captura y en la PC suéltala sobre el mismo icono.</li>
    </ol></div>`;
  $('#pick').onclick = () => $('#file').click();
}

// ------------------------------------------------------------- barra de día
function dayHasData(sheet, i, rows) {
  const col = days()[i].col;
  if (sheet === EQLOG) return false;
  for (const r of rows) for (const k of KINDS.slice(0, 4)) if (num(shown(sheet, col + (r + KINDS.indexOf(k))))) return true;
  return false;
}

function dayBar(sheet, rows) {
  const ds = days();
  const d = ds[ui.day];
  const ws = ui.day - (ui.day % 7);
  const hol = holidays();
  const t = todayIso();
  const type = dayType(ui.day);
  let week = '';
  for (let i = ws; i < Math.min(ws + 7, ds.length); i++) {
    const x = ds[i];
    const cls = ['wd', i === ui.day && 'on', x.date === t && 'today', hol.has(x.date) && 'hol'].filter(Boolean).join(' ');
    week += `<button class="${cls}" data-day="${i}">${DIA[x.dow].slice(0, 2)}<b>${+x.date.slice(8)}</b>${rows && dayHasData(sheet, i, rows) ? '<i class="dot"></i>' : ''}</button>`;
  }
  return `
    <div class="daynav">
      <button class="btn icon" data-day="${Math.max(0, ui.day - 1)}" ${ui.day === 0 ? 'disabled' : ''} aria-label="Día anterior"><svg viewBox="0 0 24 24"><path d="M15 6l-6 6 6 6"/></svg></button>
      <button class="daybtn ${type}" id="daypick"><b>${dLabel(d)}</b><span>Semana ${Math.floor(ui.day / 7) + 1}${hol.has(d.date) ? ' · festivo (DT)' : type === 'sat' ? ' · sábado (OT)' : type === 'sun' ? ' · domingo (DT)' : ''}</span></button>
      <button class="btn icon" data-day="${Math.min(ds.length - 1, ui.day + 1)}" ${ui.day === ds.length - 1 ? 'disabled' : ''} aria-label="Día siguiente"><svg viewBox="0 0 24 24"><path d="M9 6l6 6-6 6"/></svg></button>
    </div>
    <div class="week">${week}</div>`;
}

function bindDayBar() {
  for (const b of view.querySelectorAll('[data-day]')) b.onclick = () => { ui.day = +b.dataset.day; persist(); render(); };
  const p = $('#daypick');
  if (p) p.onclick = pickDay;
}

function pickDay() {
  const ds = days();
  let h = '';
  for (let w = 0; w < ds.length; w += 7) {
    h += `<h3 style="margin-top:${w ? 14 : 0}px">Semana ${w / 7 + 1}</h3><div class="wkgrid">`;
    for (let i = w; i < Math.min(w + 7, ds.length); i++) {
      const x = ds[i];
      h += `<button class="wd ${i === ui.day ? 'on' : ''} ${x.date === todayIso() ? 'today' : ''} ${holidays().has(x.date) ? 'hol' : ''}" data-i="${i}">${DIA[x.dow].slice(0, 2)}<b>${+x.date.slice(8)}</b>${MES[+x.date.slice(5, 7) - 1]}</button>`;
    }
    h += '</div>';
  }
  modal('Elegir día', h, (body) => {
    for (const b of body.querySelectorAll('[data-i]')) b.onclick = () => { ui.day = +b.dataset.i; closeModal(); persist(); render(); };
  });
}

// ------------------------------------------------------------- orden
function orderLabel(o) { return o.desc?.trim() || `Hoja ${o.sheet} (sin nombre)`; }

function orderHours(sheet, i) {
  let h = 0;
  for (const r of crewOf(sheet)) for (const k of ['REG', 'OT', 'DT']) h += num(shown(sheet, cellOf(r, k, i)));
  return h;
}

function orderButton() {
  const o = orderOf(ui.order);
  const h = orderHours(o.sheet, ui.day);
  return `<button class="orderbtn" id="orderpick"><span class="tag">${esc(o.sheet)}</span>
    <span class="ttl"><b>${esc(orderLabel(o))}</b><span>${h ? `${fmt(h)} horas este día` : 'sin horas este día'}</span></span>
    <svg viewBox="0 0 24 24"><path d="M6 9l6 6 6-6"/></svg></button>`;
}

function pickOrder(onPick) {
  const os = state.base.orders;
  const row = (o) => {
    const h = orderHours(o.sheet, ui.day);
    return `<button data-s="${esc(o.sheet)}" class="${o.sheet === ui.order ? 'on' : ''}"><span class="tag">${esc(o.sheet)}</span>
      <span class="grow"><b>${esc(orderLabel(o))}</b><span>${h ? fmt(h) + ' h este día' : ''}${o.visible ? '' : ' · hoja oculta'}</span></span></button>`;
  };
  const vis = os.filter((o) => o.visible), hid = os.filter((o) => !o.visible);
  modal('Orden de trabajo', `<div class="pick">${vis.map(row).join('')}</div>${hid.length ? `<h3>Hojas ocultas</h3><div class="pick">${hid.map(row).join('')}</div>` : ''}`, (body) => {
    for (const b of body.querySelectorAll('[data-s]')) b.onclick = () => { closeModal(); onPick(b.dataset.s); };
  });
}

// ------------------------------------------------------------- horas
function renderHours() {
  if (ui.mode === 'person') return renderPerson();
  const sheet = ui.order;
  const crew = crewOf(sheet);
  for (const r of ui.extra[sheet] || []) crew.add(r);
  // primero quien ya tiene algo este día; el resto en el orden del Excel
  const today = (r) => KINDS.slice(0, 4).some((k) => val(sheet, cellOf(r, k, ui.day)) !== null);
  const ws = workers().filter((w) => crew.has(w.r) && (w.name || hasAny(sheet, w.r)))
    .sort((a, b) => today(b.r) - today(a.r) || a.r - b.r);
  view.innerHTML = `
    ${dayBar(sheet, [...crew])}
    ${orderButton()}
    <div class="toolbar">
      <div class="seg"><button class="on">Por orden</button><button data-mode="person">Por persona</button></div>
      <span class="spacer"></span>
      <button class="btn sm" id="copyprev">Copiar día anterior</button>
    </div>
    ${ws.length > 8 ? `<div class="search" style="margin-bottom:8px"><svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg><input class="input" id="filter" placeholder="Filtrar ${ws.length} trabajadores" autocomplete="off"></div>` : ''}
    <div class="wk-list" id="list">${ws.map((w) => workerCard(sheet, w)).join('') || '<div class="card empty">Nadie tiene horas en esta orden todavía. Agrega trabajadores abajo.</div>'}</div>
    ${addBox()}`;
  bindDayBar();
  $('#orderpick').onclick = () => pickOrder((s) => { ui.order = s; persist(); render(); });
  view.querySelector('[data-mode]').onclick = () => { ui.mode = 'person'; persist(); render(); };
  $('#copyprev').onclick = () => copyPrev(sheet);
  const f = $('#filter');
  if (f) f.oninput = () => {
    const t = f.value.trim().toLowerCase();
    for (const c of view.querySelectorAll('#list .wrk')) c.hidden = !!t && !c.querySelector('.nm').textContent.toLowerCase().includes(t);
  };
  bindCards(sheet);
  bindAdd((r) => { (ui.extra[sheet] ||= []).push(r); render(); setTimeout(() => $(`[data-r="${r}"] .tot`)?.focus(), 50); }, [...crew]);
}

function hasAny(sheet, r) {
  for (let i = 0; i < days().length; i++) for (const k of KINDS) if (val(sheet, cellOf(r, k, i)) !== null) return true;
  return false;
}

function inputCell(sheet, a, cls = '') {
  if (isFormula(sheet, a)) return `<input class="ci fx" value="${esc(fmt(shown(sheet, a)))}ƒ" disabled title="Fórmula en Excel">`;
  const v = val(sheet, a);
  const c = isEdited(sheet, a) ? 'chg' : v !== null ? 'has' : '';
  return `<input class="ci ${c} ${cls}" data-a="${a}" inputmode="decimal" enterkeyhint="next" value="${esc(typeof v === 'string' && v.startsWith('=') ? fmt(shown(sheet, a)) + 'ƒ' : fmt(v))}">`;
}

function workerCard(sheet, w, title) {
  const i = ui.day;
  const hrs = ['REG', 'OT', 'DT'].reduce((s, k) => s + num(shown(sheet, cellOf(w.r, k, i))), 0);
  const other = hoursByOrder(w.r, i).filter((x) => x.sheet !== sheet);
  const tot = hrs + other.reduce((s, x) => s + x.h, 0);
  const pdElse = other.some((x) => x.pd);
  const mlA = cellOf(w.r, 'ML', i);
  const showMiles = w.miles && val(sheet, mlA) === null && !isFormula(sheet, mlA) && /mob/i.test(orderOf(sheet).desc);
  const meta = [
    w.trade && `<span class="badge">${esc(w.trade)}</span>`,
    other.length && !title ? `<span>también ${other.map((x) => `${x.sheet} ${x.h ? fmt(x.h) + 'h' : 'PD'}`).join(', ')}</span>` : '',
    tot > 16 ? `<span class="badge bad">${fmt(tot)} h en el día</span>` : other.length && tot && !title ?`<span class="badge">${fmt(tot)} h total</span>` : '',
    pdElse ? '<span class="badge ok">PD en otra orden</span>' : '',
    showMiles ? `<button class="link" data-miles="${mlA}">ML = millas (${fmt(w.miles)})</button>` : '',
  ].filter(Boolean).join('');
  return `<div class="wrk ${hrs ? 'has' : ''}" data-r="${w.r}">
    <div class="head"><span class="nm">${esc(title || w.name || '(sin nombre)')}</span></div>
    ${meta ? `<div class="meta">${meta}</div>` : ''}
    <div class="g6">
      <div><label>HORAS</label><input class="ci tot" data-tot="1" inputmode="decimal" enterkeyhint="next" value="${hrs ? fmt(hrs) : ''}"></div>
      ${KINDS.map((k) => `<div><label>${k}</label>${inputCell(sheet, cellOf(w.r, k, i))}</div>`).join('')}
    </div></div>`;
}

function bindCards(sheet) {
  for (const card of view.querySelectorAll('.wrk')) bindCard(card, card.dataset.sheet || sheet);
}

function bindCard(card, sheet) {
  const r = +card.dataset.r;
  const i = ui.day;
  const refresh = () => {
    const w = workers().find((x) => x.r === r);
    const tmp = document.createElement('div');
    tmp.innerHTML = workerCard(sheet, w, card.dataset.title);
    const nc = tmp.firstElementChild;
    if (card.dataset.sheet) { nc.dataset.sheet = card.dataset.sheet; nc.dataset.title = card.dataset.title || ''; }
    card.replaceWith(nc);
    bindCard(nc, sheet);
    persist();
    return nc;
  };
  for (const inp of card.querySelectorAll('input.ci:not([disabled])')) {
    inp.addEventListener('focus', () => inp.select());
    inp.addEventListener('change', () => {
      const raw = inp.value.replace(',', '.').replace('ƒ', '').trim();
      if (raw !== '' && (isNaN(+raw) || +raw < 0)) { toast('Solo números.'); inp.value = ''; return; }
      if (inp.dataset.tot) {
        if (raw !== '' && +raw > 24) { toast('Más de 24 horas en un día, revisa.'); return; }
        const before = hoursByOrder(r, i).filter((x) => x.sheet !== sheet).reduce((s, x) => s + x.h, 0);
        const h = raw === '' ? 0 : +raw;
        const sp = split(h, dayType(i), before);
        for (const k of ['REG', 'OT', 'DT']) setVal(sheet, cellOf(r, k, i), sp[k] || null);
        const pdA = cellOf(r, 'PD', i);
        const pdElse = hoursByOrder(r, i).some((x) => x.sheet !== sheet && x.pd);
        if (h && !pdElse && val(sheet, pdA) === null) setVal(sheet, pdA, 1);
        if (!h && isEdited(sheet, pdA)) setVal(sheet, pdA, null);
      } else {
        setVal(sheet, inp.dataset.a, raw === '' ? null : +raw);
      }
      // 'change' llega antes de que el foco pase al campo que tocó el usuario:
      // se redibuja después y se devuelve el foco a ese campo.
      setTimeout(() => {
        const ins = [...card.querySelectorAll('input.ci')];
        const at = ins.indexOf(document.activeElement);
        const nc = refresh();
        if (at >= 0) nc.querySelectorAll('input.ci')[at]?.focus();
      }, 0);
    });
  }
  const m = card.querySelector('[data-miles]');
  if (m) m.onclick = () => { setVal(sheet, m.dataset.miles, `=B${r + 4}`); refresh(); };
}

function copyPrev(sheet) {
  const i = ui.day;
  const crew = [...crewOf(sheet)];
  let p = i - 1;
  while (p >= 0 && !dayHasData(sheet, p, crew)) p--;
  if (p < 0) return toast('No hay un día anterior con horas en esta orden.');
  const undo = snapshot();
  const sameType = dayType(p) === dayType(i);
  let n = 0;
  for (const r of crew) {
    const filled = KINDS.slice(0, 4).some((k) => val(sheet, cellOf(r, k, i)) !== null || isFormula(sheet, cellOf(r, k, i)));
    if (filled || hoursByOrder(r, i).some((x) => x.sheet !== sheet)) continue;
    const pv = Object.fromEntries(KINDS.slice(0, 4).map((k) => [k, num(shown(sheet, cellOf(r, k, p)))]));
    if (!pv.REG && !pv.OT && !pv.DT && !pv.PD) continue;
    const hv = sameType ? pv : { ...split(pv.REG + pv.OT + pv.DT, dayType(i)), PD: pv.PD };
    for (const k of KINDS.slice(0, 4)) if (hv[k]) setVal(sheet, cellOf(r, k, i), hv[k]);
    n++;
  }
  persist();
  render();
  toast(n ? `Copiados ${n} trabajadores del ${dLabel(days()[p])}.` : 'Todos ya tenían algo este día; no se copió nada.', n && undo);
}

// --- agregar trabajador
function addBox() {
  return `<h3>Agregar trabajador</h3>
    <div class="search"><svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg>
    <input class="input" id="q" placeholder="Buscar por nombre u oficio" autocomplete="off"></div>
    <div class="card pick" id="qres" style="margin-top:8px" hidden></div>
    <button class="btn sm" id="newwk" style="margin-top:10px">+ Trabajador nuevo</button>`;
}

function bindAdd(onPick, exclude = []) {
  const ex = new Set(exclude);
  const q = $('#q'), res = $('#qres');
  q.oninput = () => {
    const t = q.value.trim().toLowerCase();
    if (!t) { res.hidden = true; return; }
    const ws = namedWorkers().filter((w) => !ex.has(w.r) && (`${w.name} ${w.trade}`.toLowerCase().includes(t))).slice(0, 12);
    res.hidden = false;
    res.innerHTML = ws.map((w) => `<button data-r="${w.r}"><span class="tag">${esc(w.trade || '·')}</span><span class="grow"><b>${esc(w.name)}</b></span><span class="link">Agregar</span></button>`).join('')
      || '<div class="pad small muted">Nadie con ese nombre. Usa “Trabajador nuevo”.</div>';
    for (const b of res.querySelectorAll('[data-r]')) b.onclick = () => onPick(+b.dataset.r);
  };
  $('#newwk').onclick = () => newWorker(onPick, q.value.trim());
}

function newWorker(onDone, name = '') {
  const free = workers().filter((w) => !w.name).length;
  if (!free) return toast('Ya no hay bloques vacíos en T&M - All. Inserta uno en Excel y genera la base de nuevo.');
  const opts = state.base.trades.map((t) => `<option value="${esc(t.code)}">${esc(t.code)} · ${esc(t.desc)}</option>`).join('');
  modal('Trabajador nuevo', `
    <div class="form">
      <div class="field full"><label>Nombre</label><input class="input" id="nw-name" value="${esc(name)}" autocomplete="off"></div>
      <div class="field full"><label>Oficio (Trade)</label><select class="input" id="nw-trade">${opts}</select></div>
      <div class="field"><label>Ciudad de origen</label><input class="input" id="nw-city" placeholder="Opcional"></div>
      <div class="field"><label>Millas</label><input class="input" id="nw-miles" inputmode="decimal" placeholder="Opcional"></div>
    </div>
    <p class="small muted">Quedan ${free} bloques vacíos en T&M - All. Se escribe el nombre ahí (y en Payroll si está vacío).</p>
    <button class="btn primary" id="nw-ok" style="width:100%">Agregar</button>`, (body) => {
    $('#nw-ok', body).onclick = () => {
      const nm = $('#nw-name', body).value.trim();
      if (!nm) return toast('Falta el nombre.');
      if (namedWorkers().some((w) => w.name.toLowerCase() === nm.toLowerCase())) return toast('Ya existe alguien con ese nombre.');
      const r = addWorker({ name: nm, trade: $('#nw-trade', body).value, city: $('#nw-city', body).value.trim(), miles: $('#nw-miles', body).value.trim() });
      closeModal();
      persist();
      onDone(r);
    };
  });
}

// --- por persona
function renderPerson() {
  const ws = namedWorkers();
  const w = ws.find((x) => x.r === ui.person);
  const sheets = w ? new Set(hoursByOrder(w.r, ui.day).map((x) => x.sheet)) : new Set();
  for (const s of (w && ui.extra['p' + w.r]) || []) sheets.add(s);
  view.innerHTML = `
    ${dayBar(null, null)}
    <div class="toolbar">
      <div class="seg"><button data-mode="order">Por orden</button><button class="on">Por persona</button></div>
    </div>
    ${w ? `<button class="orderbtn" id="chgperson" style="margin-top:0"><span class="tag">${esc(w.trade || '·')}</span><span class="ttl"><b>${esc(w.name)}</b><span>${fmt(hoursByOrder(w.r, ui.day).reduce((s, x) => s + x.h, 0))} h este día</span></span><svg viewBox="0 0 24 24"><path d="M6 9l6 6 6-6"/></svg></button>
      <div class="wk-list" style="margin-top:10px">${[...sheets].sort((a, b) => a.length - b.length || a.localeCompare(b)).map((s) => {
        const o = orderOf(s);
        return workerCard(s, w, `${s} · ${orderLabel(o)}`).replace('class="wrk', `data-sheet="${esc(s)}" data-title="${esc(`${s} · ${orderLabel(o)}`)}" class="wrk`);
      }).join('') || '<div class="card empty">Sin horas este día.</div>'}</div>
      <button class="btn" id="addorder" style="margin-top:10px;width:100%">+ Agregar orden</button>`
    : `${addBox().replace('Agregar trabajador', 'Elegir trabajador')}`}`;
  bindDayBar();
  view.querySelector('[data-mode]').onclick = () => { ui.mode = 'order'; persist(); render(); };
  if (w) {
    $('#chgperson').onclick = () => { ui.person = null; render(); };
    bindCards(null);
    $('#addorder').onclick = () => pickOrder((s) => { (ui.extra['p' + w.r] ||= []).push(s); render(); setTimeout(() => $(`[data-sheet="${s}"] .tot`)?.focus(), 50); });
  } else {
    bindAdd((r) => { ui.person = r; persist(); render(); });
    $('#q').focus();
  }
}

// ------------------------------------------------------------- equipo por orden
function eqRowsOf(sheet) {
  return state.base.eqRows.filter((r) => val(sheet, 'A' + r));
}

function eqList(sheet, rows, dcol, weekCols) {
  return rows.map((r) => {
    const a = dcol + r;
    const wk = weekCols.reduce((s, c) => s + num(shown(sheet, c + r)), 0);
    return `<div class="eqrow"><span class="nm">${esc(val(sheet, 'A' + r))}<span>semana: ${fmt(wk) || 0}</span></span>${inputCell(sheet, a)}</div>`;
  }).join('');
}

function bindEq(sheet) {
  for (const inp of view.querySelectorAll('.eqrow input.ci:not([disabled])')) {
    inp.addEventListener('focus', () => inp.select());
    inp.addEventListener('change', () => {
      const raw = inp.value.replace(',', '.').trim();
      if (raw !== '' && (isNaN(+raw) || +raw < 0)) { toast('Solo números.'); inp.value = ''; return; }
      setVal(sheet, inp.dataset.a, raw === '' ? null : +raw);
      inp.className = 'ci ' + (isEdited(sheet, inp.dataset.a) ? 'chg' : raw !== '' ? 'has' : '');
      persist();
    });
  }
}

function pickEquipment(onPick) {
  const list = state.base.equipment;
  modal('Agregar equipo', `
    <div class="search"><svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg><input class="input" id="eq-q" placeholder="Buscar equipo" autocomplete="off"></div>
    <div class="pick" id="eq-res" style="margin-top:8px"></div>`, (body) => {
    const q = $('#eq-q', body), res = $('#eq-res', body);
    const draw = () => {
      const t = q.value.trim().toLowerCase();
      res.innerHTML = list.filter((e) => !t || `${e.name} ${e.desc}`.toLowerCase().includes(t)).slice(0, 40)
        .map((e) => `<button data-n="${esc(e.name)}"><span class="tag">${esc(e.type || '·')}</span><span class="grow"><b>${esc(e.name)}</b><span>${esc(e.desc)}</span></span></button>`).join('');
      for (const b of res.querySelectorAll('[data-n]')) b.onclick = () => { closeModal(); onPick(b.dataset.n); };
    };
    q.oninput = draw;
    draw();
  });
}

function weekCols(i, ds) {
  const ws = i - (i % 7);
  return ds.slice(ws, ws + 7).map((d) => d.col);
}

function copyPrevEq(sheet, rows, ds, i) {
  if (i === 0) return toast('Es el primer día.');
  const undo = snapshot();
  const p = i - 1;
  let n = 0;
  for (const r of rows) {
    const a = ds[i].col + r;
    const pv = num(shown(sheet, ds[p].col + r));
    if (pv && val(sheet, a) === null && !isFormula(sheet, a)) { setVal(sheet, a, pv); n++; }
  }
  persist();
  render();
  toast(n ? `Copiados ${n} equipos del ${dLabel(ds[p])}.` : 'No había nada que copiar del día anterior.', n && undo);
}

function renderEquip() {
  const sheet = ui.order;
  const ds = days();
  const rows = eqRowsOf(sheet);
  view.innerHTML = `
    ${dayBar(null, null)}
    ${orderButton()}
    <div class="toolbar"><span class="muted small">Equipo (Both Shifts) · turnos/horas</span><span class="spacer"></span><button class="btn sm" id="copyprev">Copiar día anterior</button></div>
    <div class="card">${eqList(sheet, rows, ds[ui.day].col, weekCols(ui.day, ds)) || '<div class="empty">Esta orden no tiene equipo todavía.</div>'}</div>
    <button class="btn" id="addeq" style="margin-top:10px;width:100%">+ Agregar equipo</button>`;
  bindDayBar();
  $('#orderpick').onclick = () => pickOrder((s) => { ui.order = s; persist(); render(); });
  $('#copyprev').onclick = () => copyPrevEq(sheet, rows, ds, ui.day);
  bindEq(sheet);
  $('#addeq').onclick = () => pickEquipment((name) => {
    const had = rows.find((r) => val(sheet, 'A' + r) === name);
    if (had) return $(`[data-a="${ds[ui.day].col}${had}"]`)?.focus();
    const free = state.base.eqRows.find((r) => !val(sheet, 'A' + r) && !isFormula(sheet, 'A' + r));
    if (!free) return toast('Ya no hay filas de equipo libres en esta hoja.');
    setVal(sheet, 'A' + free, name);
    persist();
    render();
    setTimeout(() => $(`[data-a="${ds[ui.day].col}${free}"]`)?.focus(), 50);
  });
}

// ------------------------------------------------------------- equipment & material log
function renderLog() {
  const el = state.base.eqlog;
  const d = days()[ui.day];
  const li = el.days.findIndex((x) => x.date === d.date);
  const rows = el.rows.filter((r) => val(EQLOG, 'A' + r));
  view.innerHTML = `
    ${dayBar(null, null)}
    <div class="toolbar"><span class="muted small">Equipment &amp; Material Log (Both Shifts)</span><span class="spacer"></span>${li >= 0 ? '<button class="btn sm" id="copyprev">Copiar día anterior</button>' : ''}</div>
    ${li < 0 ? '<div class="card empty">Este día no está en el Equipment &amp; Material Log.</div>' : `
    <div class="card">${eqList(EQLOG, rows, el.days[li].col, weekCols(li, el.days)) || '<div class="empty">El log no tiene equipo todavía.</div>'}</div>
    <button class="btn" id="addeq" style="margin-top:10px;width:100%">+ Agregar equipo / material</button>`}`;
  bindDayBar();
  if (li < 0) return;
  $('#copyprev').onclick = () => copyPrevEq(EQLOG, rows, el.days, li);
  bindEq(EQLOG);
  $('#addeq').onclick = () => pickEquipment((name) => {
    const had = rows.find((r) => val(EQLOG, 'A' + r) === name);
    if (had) return $(`[data-a="${el.days[li].col}${had}"]`)?.focus();
    const free = el.rows.find((r) => !val(EQLOG, 'A' + r) && !isFormula(EQLOG, 'A' + r));
    if (!free) return toast('Ya no hay filas libres en el log.');
    setVal(EQLOG, 'A' + free, name);
    persist();
    render();
    setTimeout(() => $(`[data-a="${el.days[li].col}${free}"]`)?.focus(), 50);
  });
}

// ------------------------------------------------------------- enviar
function describe(e) {
  const m = /^([A-Z]+)(\d+)$/.exec(e.a);
  const c = m[1], r = +m[2];
  if (e.s === TM || e.s === 'Payroll') return `${e.s} · ${c === 'A' ? (workers().some((w) => w.r === r) ? 'nombre' : 'ciudad') : workers().some((w) => w.r === r) ? 'oficio' : 'millas'}`;
  const isLog = e.s === EQLOG;
  const ds = isLog ? state.base.eqlog.days : days();
  const d = ds.find((x) => x.col === c);
  if (c === 'A') return `${isLog ? 'Log' : e.s} · equipo nuevo`;
  if (isLog || r >= state.base.eqRows[0]) return `${isLog ? 'Log' : e.s} · ${val(e.s, 'A' + r) || 'equipo'} · ${d ? dLabel(d) : c}`;
  const wr = r - ((r - 12) % 5);
  const w = workers().find((x) => x.r === wr);
  return `${e.s} · ${w?.name || 'fila ' + wr} · ${KINDS[(r - 12) % 5]} · ${d ? dLabel(d) : c}`;
}

function renderSend() {
  const es = [...state.edits.values()];
  const show = (v, s, a) => (v === null || v === undefined || v === '' ? '—' : typeof v === 'string' && v.startsWith('=') ? `${fmt(shown(s, a))} (${v})` : esc(v));
  const bySheet = {};
  for (const e of es) (bySheet[e.s] ||= []).push(e);
  view.innerHTML = `
    <h2>Enviar al Excel</h2>
    <div class="stats"><div class="stat"><b>${es.length}</b><span>celdas sin enviar</span></div><div class="stat"><b>${Object.keys(bySheet).length}</b><span>hojas</span></div><div class="stat"><b>${state.base.file.name.replace(/\.xls[bxm]$/i, '').slice(0, 14)}</b><span>libro</span></div></div>
    <button class="btn primary" id="export" style="width:100%;height:48px" ${es.length ? '' : 'disabled'}>Exportar captura</button>
    <div class="card pad small" style="margin-top:12px"><ol class="steps">
      <li>Exporta y guárdala en OneDrive, mándala por AirDrop o correo.</li>
      <li>En la PC suelta el archivo sobre el icono <b>K2 Captura</b>. Hace respaldo, escribe con Excel y guarda el .xlsb.</li>
      <li>La PC genera una base nueva: cárgala aquí y los cambios enviados se limpian solos.</li>
    </ol></div>
    ${es.length ? `<h3>Cambios</h3>${Object.entries(bySheet).map(([s, list]) => `<div class="card changes" style="margin-bottom:8px">${list.map((e) => `<div class="chg"><span>${esc(describe(e))}</span><span><s>${show(baseVal(e.s, e.a), e.s, e.a)}</s> → <b>${show(e.v, e.s, e.a)}</b></span></div>`).join('')}</div>`).join('')}` : ''}
    <h3>Base</h3>
    <div class="card pad"><dl class="kv">
      <dt>Libro</dt><dd>${esc(state.base.file.name)}</dd>
      <dt>Proyecto</dt><dd>${esc(state.base.project.name)}</dd>
      <dt>Generada</dt><dd>${esc(state.base.stamp.slice(0, 16).replace('T', ' '))}</dd>
      <dt>Días</dt><dd>${dLabel(days()[0])} a ${dLabel(days()[days().length - 1])}</dd>
    </dl>
    <div class="row wrap" style="margin-top:12px"><button class="btn" id="loadbase">Cargar base nueva</button>${es.length ? '<button class="btn danger" id="discard">Descartar cambios</button>' : ''}</div></div>`;
  $('#export').onclick = exportCapture;
  $('#loadbase').onclick = () => $('#file').click();
  const d = $('#discard');
  if (d) d.onclick = () => {
    if (!confirm(`¿Borrar los ${es.length} cambios sin enviar?`)) return;
    const undo = snapshot();
    state.edits.clear();
    persist();
    render();
    toast('Cambios descartados.', undo);
  };
}

async function exportCapture() {
  const cap = buildCapture();
  const t = new Date();
  const p = (n) => String(n).padStart(2, '0');
  const name = `K2 captura ${t.getFullYear()}-${p(t.getMonth() + 1)}-${p(t.getDate())} ${p(t.getHours())}${p(t.getMinutes())}.json`;
  const blob = new Blob([JSON.stringify(cap)], { type: 'application/json' });
  const file = new File([blob], name, { type: 'application/json' });
  try {
    if (navigator.canShare?.({ files: [file] })) {
      await navigator.share({ files: [file], title: name });
      return toast('Captura lista. Suéltala en la PC sobre el icono K2 Captura.');
    }
  } catch (err) {
    if (err.name === 'AbortError') return;
  }
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  toast('Captura descargada. Suéltala en la PC sobre el icono K2 Captura.');
}

// ------------------------------------------------------------- modal
function modal(title, body, bind) {
  const m = $('#modal');
  $('#modal-body').innerHTML = `<div class="m-head"><h2>${esc(title)}</h2><button class="btn icon sm" id="m-x" aria-label="Cerrar"><svg viewBox="0 0 24 24"><path d="M6 6l12 12M18 6L6 18"/></svg></button></div><div class="m-body">${body}</div>`;
  $('#m-x').onclick = closeModal;
  m.onclick = (e) => { if (e.target === m) closeModal(); };
  m.showModal();
  bind?.($('#modal-body'));
}
function closeModal() { $('#modal').close(); }

init();
