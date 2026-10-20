"""K2 Captura - puente PC <-> celular para el libro TS-CEM-BRO-K2 (.xlsb).

Soltar sobre el icono:
  * el .xlsb            -> genera "<libro>.base.json" para cargarlo en el celular.
  * una captura .json   -> escribe esos valores en el .xlsb con Excel mismo
                           (respaldo antes, nunca toca celdas con formula),
                           guarda y vuelve a generar la base.

Todo lo hace Excel por COM, asi el .xlsb conserva formulas, formato y macros.
"""
import datetime as dt
import hashlib
import json
import os
import shutil
import sys
import traceback

import pythoncom
import win32com.client as win32

ORDER_SHEETS = [chr(65 + i) for i in range(26)] + ['AA', 'AB', 'AC', 'AD']
TM = 'T&M - All'
PAYROLL = 'Payroll'
EQLOG = 'Equipment & Material Log'
DAYNAMES = {'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'}
FIRST_BLOCK = 12          # fila 7 es el bloque "DAYS", no es trabajador
GRID_COLS = 47            # A..AU (5 semanas de dias + totales)
XL_MANUAL, XL_AUTO = -4135, -4105


def col_letter(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def addr(r, c):
    return f'{col_letter(c)}{r}'


def serial_to_iso(v):
    return (dt.date(1899, 12, 30) + dt.timedelta(days=int(v))).isoformat()


def clean(v):
    """Valor de Excel -> JSON (numeros enteros sin .0, sin None raros)."""
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, str):
        return v
    if isinstance(v, (int, float)):
        return v
    if v is None:
        return None
    return str(v)


def grid(ws, rows, cols):
    rng = ws.Range(ws.Cells(1, 1), ws.Cells(rows, cols))
    vals = rng.Value2
    forms = rng.Formula
    return vals, forms


def cell_entry(vals, forms, r, c):
    v = vals[r - 1][c - 1]
    f = forms[r - 1][c - 1]
    if isinstance(f, str) and f.startswith('='):
        return {'f': f, 'v': clean(v)}
    v = clean(v)
    if v is None or v == '':
        return None
    return v


# ---------------------------------------------------------------- Excel
class Excel:
    """Usa el libro si ya esta abierto en Excel; si no, lo abre oculto."""

    def __init__(self, path, readonly):
        pythoncom.CoInitialize()
        self.path = os.path.abspath(path)
        self.own_app = False
        self.own_wb = False
        self.app = None
        self.wb = None
        try:
            active = win32.GetActiveObject('Excel.Application')
            for wb in active.Workbooks:
                if os.path.normcase(wb.FullName) == os.path.normcase(self.path):
                    self.app, self.wb = active, wb
                    break
        except Exception:
            pass
        if self.wb is None:
            # instancia propia y oculta: no se mete con el Excel que tengas abierto
            self.app = win32.DispatchEx('Excel.Application')
            self.own_app = True
            self.app.Visible = False
            self.app.DisplayAlerts = False
            self.wb = self.app.Workbooks.Open(self.path, 0, readonly)
            self.own_wb = True
            if self.wb is None:
                raise RuntimeError(f'Excel no pudo abrir {self.path}')

    def close(self, save=False):
        try:
            if self.own_wb:
                self.wb.Close(save)
        finally:
            if self.own_app:
                self.app.Quit()


# ---------------------------------------------------------------- base
def build_base(wb, path):
    tm = wb.Worksheets(TM)
    vals, forms = grid(tm, 708, GRID_COLS)

    days = []
    for c in range(5, GRID_COLS + 1):
        name = vals[5][c - 1]
        f6 = forms[5][c - 1]
        d = vals[4][c - 1]
        if name in DAYNAMES and isinstance(f6, str) and f6.startswith('=') and isinstance(d, (int, float)):
            days.append({'col': col_letter(c), 'c': c, 'date': serial_to_iso(d), 'dow': name})

    # fila "Man Hours" marca el final de los bloques de trabajadores
    a_sheet = wb.Worksheets('A')
    av, af = grid(a_sheet, 708, 6)
    end = next(r for r in range(FIRST_BLOCK, 708) if av[r - 1][3] == 'Man Hours')
    eq_hdr = next(r for r in range(end, 708) if isinstance(av[r - 1][0], str) and av[r - 1][0].startswith('Equipment'))
    eq_rows = [r for r in range(eq_hdr + 1, 709) if isinstance(af[r - 1][1], str) and 'EqRange' in af[r - 1][1]]

    pr = wb.Worksheets(PAYROLL)
    pv, pf = grid(pr, end, 2)
    workers = []
    for r in range(FIRST_BLOCK, end, 5):
        workers.append({
            'r': r,
            'name': clean(vals[r - 1][0]) or '',
            'trade': clean(vals[r - 1][1]) or '',
            'city': clean(vals[r + 3][0]) or '',
            'miles': clean(vals[r + 3][1]),
            'pname': clean(pv[r - 1][0]) or '',
            'ptrade': clean(pv[r - 1][1]) or '',
            'pform': isinstance(pf[r - 1][0], str) and pf[r - 1][0].startswith('='),
        })

    day_cols = [d['c'] for d in days]
    orders = []
    for s in ORDER_SHEETS:
        ws = wb.Worksheets(s)
        v, f = grid(ws, 708, GRID_COLS)
        cells = {}
        for r in list(range(FIRST_BLOCK, end)) + eq_rows:
            for c in day_cols:
                e = cell_entry(v, f, r, c)
                if e is not None:
                    cells[addr(r, c)] = e
        for r in eq_rows:
            e = cell_entry(v, f, r, 1)
            if e is not None:
                cells[addr(r, 1)] = e
        orders.append({'sheet': s, 'desc': clean(v[2][1]) or '', 'visible': ws.Visible == -1, 'cells': cells})

    el = wb.Worksheets(EQLOG)
    ev, ef = grid(el, 120, GRID_COLS)
    el_days = []
    for c in range(5, GRID_COLS + 1):
        name, d = ev[4][c - 1], ev[3][c - 1]
        if name in DAYNAMES and isinstance(d, (int, float)):
            el_days.append({'col': col_letter(c), 'c': c, 'date': serial_to_iso(d), 'dow': name})
    el_rows = [r for r in range(6, 121) if isinstance(ef[r - 1][1], str) and 'EqRange' in ef[r - 1][1]]
    el_cells = {}
    for r in el_rows:
        for c in [1] + [d['c'] for d in el_days]:
            e = cell_entry(ev, ef, r, c)
            if e is not None:
                el_cells[addr(r, c)] = e

    er = wb.Worksheets('Equipment Rates')
    rv, _ = grid(er, 100, 3)
    equipment = [{'name': clean(rv[r][0]), 'desc': clean(rv[r][1]) or '', 'type': clean(rv[r][2]) or ''}
                 for r in range(1, 100) if rv[r][0]]

    rt = wb.Worksheets('Rates')
    tv, _ = grid(rt, 30, 3)
    trades = [{'code': clean(tv[r][1]), 'desc': clean(tv[r][2]) or ''} for r in range(3, 30) if tv[r][1]]

    info = wb.Worksheets('Info')
    iv, _ = grid(info, 30, 10)
    holidays = [serial_to_iso(iv[r][8]) for r in range(2, 30) if isinstance(iv[r][8], (int, float))]

    base = {
        'app': 'k2-base', 'v': 1,
        'file': {'name': os.path.basename(path), 'path': os.path.abspath(path)},
        'project': {'customer': clean(iv[2][2]), 'plant': clean(iv[3][2]), 'name': clean(iv[5][2]), 'number': clean(iv[6][2])},
        'days': days, 'holidays': holidays,
        'workers': workers, 'trades': trades, 'equipment': equipment,
        'eqRows': eq_rows, 'orders': orders,
        'eqlog': {'days': el_days, 'rows': el_rows, 'cells': el_cells},
    }
    body = json.dumps(base, ensure_ascii=False, sort_keys=True)
    base['stamp'] = dt.datetime.now().strftime('%Y-%m-%dT%H:%M:%S') + '-' + hashlib.sha1(body.encode()).hexdigest()[:8]
    return base


def write_base(wb, path):
    base = build_base(wb, path)
    out = os.path.splitext(path)[0] + '.base.json'
    with open(out, 'w', encoding='utf-8') as fh:
        json.dump(base, fh, ensure_ascii=False, separators=(',', ':'))
    n = sum(len(o['cells']) for o in base['orders'])
    named = sum(1 for w in base['workers'] if w['name'])
    print(f'  Base lista: {out}')
    print(f'  {named} trabajadores, {len(base["orders"])} ordenes, {len(base["days"])} dias '
          f'({base["days"][0]["date"]} a {base["days"][-1]["date"]}), {n} celdas con datos.')
    return out


# ---------------------------------------------------------------- captura
def same(a, b):
    if a in (None, '') and b in (None, ''):
        return True
    try:
        return abs(float(a) - float(b)) < 1e-9
    except (TypeError, ValueError):
        return str(a) == str(b)


def apply_capture(cap, xlsb, auto_yes=False):
    xl = Excel(xlsb, readonly=False)
    wb = xl.wb
    ok = False
    try:
        bdir = os.path.join(os.path.dirname(xl.path), 'respaldos K2')
        os.makedirs(bdir, exist_ok=True)
        stem, ext = os.path.splitext(os.path.basename(xl.path))
        stamp = dt.datetime.now().strftime('%Y%m%d-%H%M%S')
        backup = os.path.join(bdir, f'{stem} {stamp}{ext}')
        if xl.own_wb:
            shutil.copy2(xl.path, backup)
        else:
            wb.SaveCopyAs(backup)
        print(f'  Respaldo: {backup}')

        writes = cap.get('writes', [])
        plan, skipped, conflicts = [], [], []
        for w in writes:
            cell = wb.Worksheets(w['s']).Range(w['a'])
            new = w.get('v')
            if cell.HasFormula:
                # el celular nunca edita celdas con formula; si ahora tiene una, se respeta
                if cell.Formula != new:
                    skipped.append((w, f'tiene formula {cell.Formula}'))
                continue
            cur_v = clean(cell.Value2)
            if not same(cur_v, w.get('p')):
                if same(cur_v, new):
                    continue
                conflicts.append((w, cur_v))
            plan.append((w, cell))

        if conflicts:
            print(f'\n  {len(conflicts)} celdas cambiaron en la PC desde que se hizo la base:')
            for w, cur in conflicts[:25]:
                print(f'    {w["s"]}!{w["a"]}: en Excel={cur!r}  base={w.get("p")!r}  celular={w.get("v")!r}')
            if len(conflicts) > 25:
                print(f'    ... y {len(conflicts) - 25} mas')
            ans = 's' if auto_yes else input('  ¿Sobrescribirlas con lo del celular? (s/n): ').strip().lower()
            if ans != 's':
                keep = {id(w) for w, _ in conflicts}
                plan = [(w, c) for w, c in plan if id(w) not in keep]
                print('  Esas celdas se dejan como estan en Excel.')

        app = xl.app
        calc = app.Calculation
        app.ScreenUpdating = False
        app.Calculation = XL_MANUAL
        try:
            for w, cell in plan:
                v = w.get('v')
                if v is None or v == '':
                    cell.ClearContents()
                elif isinstance(v, str) and v.startswith('='):
                    cell.Formula = v
                else:
                    cell.Value2 = v
        finally:
            app.Calculation = calc if not xl.own_app else XL_AUTO
            app.ScreenUpdating = True
        app.Calculate()
        wb.Save()
        ok = True

        log = os.path.join(bdir, f'{stem} {stamp} captura.txt')
        with open(log, 'w', encoding='utf-8') as fh:
            fh.write(f'Captura del {cap.get("created")} sobre base {cap.get("base", {}).get("stamp")}\n')
            for w, _ in plan:
                fh.write(f'escrita  {w["s"]}!{w["a"]} = {w.get("v")!r} (antes {w.get("p")!r})\n')
            for w, why in skipped:
                fh.write(f'omitida  {w["s"]}!{w["a"]}: {why}\n')
        print(f'\n  Escritas {len(plan)} celdas en {os.path.basename(xl.path)} y guardado.')
        if skipped:
            print(f'  Omitidas {len(skipped)} celdas porque tienen formula (ver {log}).')
        write_base(wb, xl.path)
    finally:
        xl.close(save=False)
    return ok


def find_xlsb(cap, dropped):
    for p in dropped:
        if p.lower().endswith(('.xlsb', '.xlsx', '.xlsm')):
            return p
    p = cap.get('base', {}).get('path')
    if p and os.path.exists(p):
        return p
    here = os.path.dirname(os.path.abspath(dropped[0])) if dropped else os.getcwd()
    name = cap.get('base', {}).get('name')
    if name and os.path.exists(os.path.join(here, name)):
        return os.path.join(here, name)
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    p = filedialog.askopenfilename(title='¿En qué Excel escribo la captura?',
                                   filetypes=[('Excel', '*.xlsb *.xlsm *.xlsx')])
    root.destroy()
    return p or None


def main(argv):
    files = [a for a in argv if not a.startswith('--')]
    auto_yes = '--yes' in argv
    if not files:
        print('Suelta sobre este icono el .xlsb (para mandar la base al celular)\n'
              'o la captura .json que exportó el celular (para escribirla en el Excel).')
        return
    caps = [f for f in files if f.lower().endswith('.json')]
    books = [f for f in files if f.lower().endswith(('.xlsb', '.xlsx', '.xlsm'))]
    if caps:
        for c in caps:
            with open(c, encoding='utf-8') as fh:
                cap = json.load(fh)
            if cap.get('app') != 'k2-captura':
                print(f'  {os.path.basename(c)} no es una captura de K2 Captura.')
                continue
            xlsb = find_xlsb(cap, books + [c])
            if not xlsb:
                print('  No se eligió Excel; no se escribió nada.')
                continue
            print(f'Captura {os.path.basename(c)}: {len(cap.get("writes", []))} cambios -> {os.path.basename(xlsb)}')
            if apply_capture(cap, xlsb, auto_yes):
                done = os.path.join(os.path.dirname(c), 'capturas aplicadas')
                os.makedirs(done, exist_ok=True)
                shutil.move(c, os.path.join(done, os.path.basename(c)))
    else:
        for b in books:
            print(f'Generando base de {os.path.basename(b)} ...')
            xl = Excel(b, readonly=True)
            try:
                out = write_base(xl.wb, xl.path)
            finally:
                xl.close(save=False)
            print('  Pásala al celular (OneDrive, AirDrop o correo) y cárgala en K2 Captura.')
            if '--no-open' not in argv:
                os.startfile(os.path.dirname(out))


if __name__ == '__main__':
    pause = '--pause' in sys.argv
    try:
        main(sys.argv[1:])
    except Exception:
        traceback.print_exc()
        print('\nAlgo falló; no se guardó nada que no hayas visto arriba.')
    if pause:
        input('\nPresiona Enter para cerrar...')
