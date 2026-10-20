// Modelo: la base que exportó la PC + los cambios capturados en el celular.
// Nunca se edita una celda que en la base tiene fórmula.

export const TM = 'T&M - All';
export const PAYROLL = 'Payroll';
export const EQLOG = 'Equipment & Material Log';
export const KINDS = ['REG', 'OT', 'DT', 'PD', 'ML'];

export const state = { base: null, edits: new Map() };

const key = (s, a) => `${s}!${a}`;

export function orderOf(sheet) { return state.base.orders.find((o) => o.sheet === sheet); }

function baseRaw(s, a) {
  const b = state.base;
  if (s === EQLOG) return b.eqlog.cells[a];
  if (s === TM || s === PAYROLL) {
    const m = /^([AB])(\d+)$/.exec(a);
    if (!m) return undefined;
    const r = +m[2];
    const w = b.workers.find((x) => x.r === r || x.r + 4 === r);
    if (!w) return undefined;
    if (s === PAYROLL) return w.r === r ? (m[1] === 'A' ? w.pname : w.ptrade) || undefined : undefined;
    if (w.r === r) return (m[1] === 'A' ? w.name : w.trade) || undefined;
    return (m[1] === 'A' ? w.city : w.miles) ?? undefined;
  }
  return orderOf(s)?.cells[a];
}

export const isFormula = (s, a) => typeof baseRaw(s, a) === 'object' && baseRaw(s, a) !== null;
export const formulaOf = (s, a) => (isFormula(s, a) ? baseRaw(s, a).f : null);

/** Valor plano de la base (sin cambios). */
export function baseVal(s, a) {
  const v = baseRaw(s, a);
  if (v && typeof v === 'object') return v.v;
  return v ?? null;
}

export const isEdited = (s, a) => state.edits.has(key(s, a));

/** Valor vigente: cambio del celular si existe, si no el de la base. */
export function val(s, a) {
  const e = state.edits.get(key(s, a));
  return e ? e.v : baseVal(s, a);
}

export const num = (v) => (typeof v === 'number' ? v : v && !isNaN(+v) ? +v : 0);

function same(a, b) {
  if ((a === null || a === '' || a === undefined) && (b === null || b === '' || b === undefined)) return true;
  return String(a) === String(b);
}

/** Escribe un valor. Regresa false si la celda está protegida por fórmula. */
export function setVal(s, a, v) {
  if (isFormula(s, a)) {
    if (typeof v === 'string' && v === formulaOf(s, a)) { state.edits.delete(key(s, a)); return true; }
    return false;
  }
  if (v === '' || v === undefined) v = null;
  if (typeof v === 'string' && v.trim() !== '' && !v.startsWith('=') && !isNaN(+v)) v = +v;
  if (same(v, baseVal(s, a))) state.edits.delete(key(s, a));
  else state.edits.set(key(s, a), { s, a, v });
  return true;
}

/** Valor que se ve en pantalla para una celda que puede ser fórmula (=B16 → millas). */
export function shown(s, a) {
  const v = val(s, a);
  if (typeof v === 'string' && v.startsWith('=')) {
    const m = /^=B(\d+)$/.exec(v);
    if (m) return num(val(TM, 'B' + m[1]));
  }
  return v;
}

// ------------------------------------------------------------- trabajadores
export function workers() {
  return state.base.workers.map((w) => ({
    r: w.r,
    name: val(TM, 'A' + w.r) || '',
    trade: val(TM, 'B' + w.r) || '',
    city: val(TM, 'A' + (w.r + 4)) || '',
    miles: val(TM, 'B' + (w.r + 4)),
  }));
}
export const namedWorkers = () => workers().filter((w) => w.name);

export function addWorker({ name, trade, city, miles }) {
  const free = workers().find((w) => !w.name);
  if (!free) return null;
  const r = free.r;
  setVal(TM, 'A' + r, name);
  setVal(TM, 'B' + r, trade);
  if (city) setVal(TM, 'A' + (r + 4), city);
  if (miles !== '' && miles != null) setVal(TM, 'B' + (r + 4), +miles);
  const bw = state.base.workers.find((w) => w.r === r);
  if (!bw.pform && !bw.pname) {
    setVal(PAYROLL, 'A' + r, name);
    setVal(PAYROLL, 'B' + r, trade);
  }
  return r;
}

// ------------------------------------------------------------- días
export const days = () => state.base.days;
export const holidays = () => new Set(state.base.holidays);

/** 'wk' lunes-viernes, 'sat', 'sun' (domingo o festivo federal). */
export function dayType(i) {
  const d = days()[i];
  if (holidays().has(d.date) || d.dow === 'Sun') return 'sun';
  return d.dow === 'Sat' ? 'sat' : 'wk';
}

export const cellOf = (r, kind, i) => days()[i].col + (r + KINDS.indexOf(kind));

/** Horas (REG+OT+DT) de un trabajador en un día, por orden. */
export function hoursByOrder(r, i) {
  const out = [];
  for (const o of state.base.orders) {
    let h = 0, pd = 0;
    for (const k of ['REG', 'OT', 'DT']) h += num(shown(o.sheet, cellOf(r, k, i)));
    pd = num(shown(o.sheet, cellOf(r, 'PD', i)));
    if (h || pd) out.push({ sheet: o.sheet, h, pd });
  }
  return out;
}

/**
 * Reparte horas totales según las reglas de Rates: entre semana las primeras 8 del día
 * son REG, de 8 a 12 OT y más de 12 DT; sábado todo OT; domingo y festivos todo DT.
 * `before` = horas que el trabajador ya tiene ese día en otras órdenes.
 */
export function split(h, type, before = 0) {
  if (type === 'sat') return { REG: 0, OT: h, DT: 0 };
  if (type === 'sun') return { REG: 0, OT: 0, DT: h };
  const seg = (lo, hi) => Math.max(0, Math.min(before + h, hi) - Math.max(before, lo));
  return { REG: seg(0, 8), OT: seg(8, 12), DT: seg(12, Infinity) };
}

export function crewOf(sheet) {
  const o = orderOf(sheet);
  const rows = new Set();
  const lim = state.base.eqRows[0];
  const take = (a) => {
    const m = /^[A-Z]+(\d+)$/.exec(a);
    const r = +m[1];
    if (r >= lim) return;
    rows.add(r - ((r - 12) % 5));
  };
  for (const a of Object.keys(o.cells)) take(a);
  for (const e of state.edits.values()) if (e.s === sheet) take(e.a);
  return rows;
}

// ------------------------------------------------------------- exportar
export function buildCapture() {
  const b = state.base;
  const writes = [...state.edits.values()].map((e) => ({ s: e.s, a: e.a, v: e.v, p: baseVal(e.s, e.a) }));
  return {
    app: 'k2-captura', v: 1, created: new Date().toISOString(),
    base: { stamp: b.stamp, name: b.file.name, path: b.file.path },
    writes,
  };
}

/** Al cargar una base nueva, quita los cambios que ya están en el Excel. */
export function reconcile() {
  for (const [k, e] of state.edits) {
    if (isFormula(e.s, e.a) || same(e.v, baseVal(e.s, e.a))) state.edits.delete(k);
  }
}
