#!/usr/bin/env python3
"""Monthly supplier spend & rate tracker for AUW camp sites.

Reads the purchase ledger (sheet "Accounts_ Inventory" of the Accounts Inventory
workbook) and, optionally, an export of the AUW Financial System (sheet
"Transactions") for purchases entered only in the system, and writes an Excel
report that separates price changes from volume changes, supplier by supplier.

Usage:
    python tools/supplier_tracker.py LEDGER.xlsx [--system SYSTEM.xlsx]
        [--sites Misfa,Nizwa] [--base 2025-10:2026-02] [--norms NORMS.csv]
        [--out tracker.xlsx]

--norms (default: consumption_norms.csv next to this script) lists the kitchen's
weekly usage per site and item group; the "Expected vs Purchased" sheet compares
it with purchased kilograms each month.

--system adds only rows whose reference starts with "INV/" (entered directly in
the system). Drop it once the ledger has been re-imported, or those purchases
will be counted twice.
"""
import argparse
import calendar
import collections
import csv
import datetime as dt
import os
import re
import statistics

import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SITE_CODES = {'MSFA': 'Misfa', 'NZWA': 'Nizwa', 'AVN': 'Avano', 'POE': 'Poesia'}
SUPPLIER_ALIASES = {'cash': 'Cash', 'AIN': 'Al-Ainkawi L.L.C'}

# Flag thresholds
SPEND_SPIKE = 1.5          # month spend vs supplier's trailing 6-month median
SPEND_MIN_RO = 150         # ignore suppliers below this median monthly spend
RATE_MOM = 0.10            # item rate change vs previous month
RATE_VS_BASE = 0.20        # item rate change vs base period
INDEX_GAP = 10             # supplier price index points above all-supplier index
UNIT_ERROR = 5.0           # item rate this far from base (either way) is treated as a unit-entry error

HDR_FONT = Font(bold=True, color='FFFFFF')
HDR_FILL = PatternFill('solid', fgColor='1F3A5F')
RED = PatternFill('solid', fgColor='F8D7DA')
AMBER = PatternFill('solid', fgColor='FFF3CD')
GREEN = PatternFill('solid', fgColor='D4EDDA')


def site_of(text):
    t = str(text or '').lower()
    for name in ('Misfa', 'Nizwa', 'Avano', 'Poesia'):
        if name.lower()[:4] in t:
            return name
    return None


def num(x):
    try:
        return float(x or 0)
    except (TypeError, ValueError):
        return None


def load_ledger(path, sites):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = []
    for r in wb['Accounts_ Inventory'].iter_rows(min_row=4, values_only=True):
        r = r[1:22]
        if r[0] != 'Invoice' or not isinstance(r[2], dt.datetime):
            continue
        site = site_of(r[3])
        q, rate, amt = num(r[9]), num(r[10]), num(r[13])
        if site not in sites or None in (q, rate, amt):
            continue
        sup = str(r[18] or '').strip() or '(no supplier)'
        out.append(dict(month=r[2].strftime('%Y-%m'), site=site, supplier=SUPPLIER_ALIASES.get(sup, sup),
                        category=str(r[7] or '').strip().upper() or '?', item=str(r[6] or '').strip(),
                        uom=str(r[8] or '').strip().lower(), qty=q, rate=rate, amount=amt))
    return out


def load_system(path, sites):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = []
    for r in wb['Transactions'].iter_rows(min_row=4, values_only=True):
        if r[2] != 'INV' or not str(r[0]).startswith('INV/') or not isinstance(r[1], dt.datetime):
            continue
        site = SITE_CODES.get(r[4])
        if site not in sites:
            continue
        sup = str(r[5] or '').strip()
        out.append(dict(month=r[1].strftime('%Y-%m'), site=site, supplier=SUPPLIER_ALIASES.get(sup, sup),
                        category=str(r[9] or '').strip().upper() or '?', item=str(r[7] or '').strip(),
                        uom=str(r[10] or '').strip().lower(), qty=num(r[11]) or 0, rate=num(r[12]) or 0,
                        amount=num(r[15]) or 0))
    return out


def month_range(a, b):
    y, m = map(int, a.split('-'))
    out = []
    while f'{y}-{m:02d}' <= b:
        out.append(f'{y}-{m:02d}')
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return out


class Rates:
    """Quantity-weighted average rate per (supplier, item, uom) and month."""

    def __init__(self, lines):
        self.d = collections.defaultdict(lambda: collections.defaultdict(lambda: [0.0, 0.0]))
        self.cat = {}
        for x in lines:
            if x['qty'] > 0 and x['rate'] > 0:
                k = (x['supplier'], x['item'], x['uom'])
                a = self.d[k][x['month']]
                a[0] += x['qty']
                a[1] += x['qty'] * x['rate']
                self.cat[k] = x['category']

    def rate(self, key, months):
        d = self.d[key]
        q = sum(d[m][0] for m in months if m in d)
        return (sum(d[m][1] for m in months if m in d) / q, q) if q else (None, 0.0)

    def index(self, months_now, base, keyfilter):
        """Base-quantity-weighted price index (base = 100) over items bought in both periods.

        Item-months whose rate is more than UNIT_ERROR times (or less than 1/UNIT_ERROR of) the base
        rate are skipped: they are carton-vs-kg style entry errors, not price changes.
        """
        num_ = den = 0.0
        for k in self.d:
            if not keyfilter(k):
                continue
            br, bq = self.rate(k, base)
            cr, _ = self.rate(k, months_now)
            if br and cr and 1 / UNIT_ERROR <= cr / br <= UNIT_ERROR:
                num_ += cr * bq
                den += br * bq
        return 100 * num_ / den if den else None


def load_norms(path, sites):
    """Weekly usage norms: site, group, weekly_kg, yield (usable share of purchased kg), match keywords."""
    if not path or not os.path.exists(path):
        return []
    out = []
    with open(path, newline='', encoding='utf-8') as f:
        for row in csv.DictReader(f):
            if row['site'].strip() not in sites:
                continue
            out.append(dict(site=row['site'].strip(), group=row['group'].strip(),
                            weekly_kg=float(row['weekly_kg']), yld=float(row.get('yield') or 1.0),
                            match=[m.strip().lower() for m in row['match'].split(';') if m.strip()],
                            source=(row.get('source') or '').strip()))
    return out


def line_kg(x):
    """Purchased weight in kg: kg lines as-is; 'pcs' of e.g. 'Whole Chicken 10*1100' as 1.1 kg each."""
    if x['uom'] == 'kg':
        return x['qty']
    m = re.search(r'\*(\d{3,4})\b', x['item'])
    if x['uom'] in ('pcs', 'pc', 'no', 'nos') and m:
        return x['qty'] * int(m.group(1)) / 1000
    return None


def expected_vs_purchased(lines, norms, months):
    """Rows of (site, group, month, norm, expected kg, purchased kg, ratio, gap kg, RO/kg, gap RO, 3m ratio)."""
    rows, skipped = [], collections.Counter()
    for n in norms:
        kg = collections.Counter()
        ro = collections.Counter()
        for x in lines:
            if x['site'] != n['site'] or not any(k in x['item'].lower() for k in n['match']):
                continue
            w = line_kg(x)
            if w is None:
                skipped[(n['site'], n['group'])] += 1
                continue
            kg[x['month']] += w
            ro[x['month']] += x['amount']
        for i, m in enumerate(months):
            y, mo = map(int, m.split('-'))
            exp = n['weekly_kg'] * calendar.monthrange(y, mo)[1] / 7 / n['yld']
            got = kg[m]
            rate = ro[m] / got if got else None
            w3 = months[max(0, i - 2):i + 1]
            e3 = sum(n['weekly_kg'] * calendar.monthrange(*map(int, mm.split('-')))[1] / 7 / n['yld'] for mm in w3)
            rows.append(dict(site=n['site'], group=n['group'], month=m, norm=n['weekly_kg'], yld=n['yld'],
                             expected=exp, purchased=got, ratio=got / exp if exp else None, gap=got - exp,
                             rate=rate, gap_ro=(got - exp) * rate if rate else None,
                             ratio3=sum(kg[mm] for mm in w3) / e3 if e3 else None, source=n['source']))
    return rows, skipped


def style_header(ws, row=1):
    for c in ws[row]:
        c.font, c.fill = HDR_FONT, HDR_FILL
        c.alignment = Alignment(wrap_text=True, vertical='center')


def widths(ws, w):
    for i, x in enumerate(w, 1):
        ws.column_dimensions[get_column_letter(i)].width = x


def build(lines, base, out_path, sites, norms=()):
    months = sorted({x['month'] for x in lines if x['month'] >= base[0]})
    last = months[-1]
    R = Rates(lines)
    spend = collections.defaultdict(collections.Counter)
    for x in lines:
        spend[x['supplier']][x['month']] += x['amount']
    suppliers = sorted(spend, key=lambda s: -sum(spend[s][m] for m in months))
    cats = sorted({R.cat[k] for k in R.cat})

    wb = openpyxl.Workbook()
    # ---- README
    ws = wb.active
    ws.title = 'Read me'
    for line in [
        ['Supplier spend & rate tracker'],
        [f'Sites: {", ".join(sites)} | Base period: {base[0]} to {base[-1]} | Months: {months[0]} to {last} | '
         f'Generated {dt.date.today()}'],
        [],
        ['Sheet', 'What it answers'],
        ['Flags', 'What changed this month that needs a question asked'],
        ['Expected vs Purchased', 'Kilograms bought vs what the kitchen says it uses (weekly norms per site). '
                                  'A ratio well above 1 every month means food bought but not used in the menu'],
        ['Price vs Volume', 'Is a supplier costing more because prices rose or because we bought more?'],
        ['Price Index', 'Each supplier\'s prices vs base period (100 = base). Compare with "ALL SUPPLIERS" '
                        'to see whether a rise is market-wide or supplier-specific'],
        ['Category Index', 'Price movement by account category (FRZ, VEG, RIC...)'],
        ['Item Rates', 'Rate per unit, per item, per supplier, per month'],
        ['Supplier Spend', 'RO spent per supplier per month'],
        [],
        ['Reading the index', 'Price index = what this month\'s prices would cost for the base period\'s basket. '
                              'Volume index = spend index / price index. Items bought only in one period are '
                              'excluded from the price index, so check "coverage".'],
        ['Limits', 'Purchases are not consumption: stock build-ups and draw-downs show up as volume swings. '
                   'Cash purchases have no supplier statement and rely on receipts.'],
    ]:
        ws.append(line)
    ws['A1'].font = Font(bold=True, size=14)
    for c in ws[4]:
        c.font, c.fill = HDR_FONT, HDR_FILL
    widths(ws, [22, 120])

    # ---- Price index rows
    all_idx = {m: R.index([m], base, lambda k: True) for m in months}

    # ---- Flags
    flags = []
    prev = months[-2] if len(months) > 1 else None
    trail = months[-7:-1]
    for s in suppliers:
        med = statistics.median([spend[s][m] for m in trail]) if trail else 0
        cur = spend[s][last]
        if med >= SPEND_MIN_RO and cur >= SPEND_SPIKE * med:
            flags.append(('High', 'Spend spike', s, '', f'{cur:,.0f} RO vs 6-month median {med:,.0f} ({cur/med:.1f}x)'))
        if med >= SPEND_MIN_RO and cur == 0:
            flags.append(('High', 'Supplier stopped', s, '', f'Nothing this month; 6-month median {med:,.0f} RO - '
                                                            'supplier change or missing invoices?'))
        if cur >= SPEND_MIN_RO and all(spend[s][m] == 0 for m in trail):
            flags.append(('Medium', 'New supplier', s, '', f'{cur:,.0f} RO this month, none in previous 6 months'))
        si = R.index([last], base, lambda k, s=s: k[0] == s)
        if si and all_idx[last] and si - all_idx[last] >= INDEX_GAP:
            flags.append(('High', 'Prices above market', s, '',
                          f'Price index {si:.0f} vs all-supplier {all_idx[last]:.0f} - rise is supplier-specific'))
    for k in R.d:
        cr, cq = R.rate(k, [last])
        if not cr:
            continue
        val = cr * cq
        if val < 20:
            continue
        pr, _ = R.rate(k, [prev]) if prev else (None, 0)
        br, _ = R.rate(k, base)
        if pr and abs(cr / pr - 1) >= RATE_MOM and 1 / UNIT_ERROR <= cr / pr <= UNIT_ERROR:
            flags.append(('High' if cr > pr else 'Low', 'Rate change MoM', k[0], f'{k[1]} ({k[2]})',
                          f'{pr:.3f} -> {cr:.3f} ({cr/pr-1:+.0%}); {val:,.0f} RO bought this month'))
        if br and RATE_VS_BASE <= cr / br - 1 < UNIT_ERROR - 1:
            flags.append(('Medium', 'Rate vs base', k[0], f'{k[1]} ({k[2]})',
                          f'{br:.3f} -> {cr:.3f} ({cr/br-1:+.0%} vs base); {val:,.0f} RO bought this month'))
    evp, evp_skipped = expected_vs_purchased(lines, norms, months) if norms else ([], {})
    for r in evp:
        if r['month'] != last or r['ratio3'] is None:
            continue
        if r['ratio3'] >= 1.5:
            flags.append(('High', 'Bought > usage norm', r['site'], r['group'],
                          f"3-month purchases are {r['ratio3']:.1f}x the weekly norm ({r['norm']:g} kg/wk); "
                          f"{last}: {r['purchased']:,.0f} kg bought vs {r['expected']:,.0f} kg expected"))
        elif r['ratio3'] <= 0.7:
            flags.append(('Medium', 'Bought < usage norm', r['site'], r['group'],
                          f"3-month purchases are {r['ratio3']:.1f}x the norm - missing invoices, menu change "
                          f"or norm too high?"))
    ws = wb.create_sheet('Flags')
    ws.append(['Severity', 'Type', 'Supplier', 'Item', f'Detail ({last})'])
    style_header(ws)
    for f in sorted(flags, key=lambda f: ({'High': 0, 'Medium': 1, 'Low': 2}[f[0]], f[1], f[2])):
        ws.append(list(f))
        ws.cell(ws.max_row, 1).fill = {'High': RED, 'Medium': AMBER, 'Low': GREEN}[f[0]]
    widths(ws, [10, 20, 28, 34, 90])
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions

    # ---- Expected vs Purchased
    if evp:
        ws = wb.create_sheet('Expected vs Purchased', 2)
        ws.append(['Site', 'Item group', 'Weekly norm kg', 'Usable share', 'Months',
                   'Expected kg', 'Purchased kg', 'Ratio', 'Gap kg', 'Gap RO (at avg rate)'])
        style_header(ws)
        for span in (months[-3:], months[-12:]):
            label = f'Last {len(span)} months'
            for site in sites:
                for g in dict.fromkeys(r['group'] for r in evp if r['site'] == site):
                    rs = [r for r in evp if r['site'] == site and r['group'] == g and r['month'] in span]
                    e = sum(r['expected'] for r in rs)
                    p = sum(r['purchased'] for r in rs)
                    gro = sum(r['gap_ro'] or 0 for r in rs)
                    ws.append([site, g, rs[0]['norm'], rs[0]['yld'], f'{label} ({span[0]} to {span[-1]})',
                               round(e), round(p), round(p / e, 2) if e else None, round(p - e), round(gro)])
        top = ws.max_row
        ws.append([])
        ws.append(['Site', 'Item group', 'Month', 'Weekly norm kg', 'Expected kg', 'Purchased kg', 'Ratio',
                   'Gap kg', 'Avg RO/kg', 'Gap RO', '3-month ratio'])
        style_header(ws, ws.max_row)
        first = ws.max_row + 1
        for r in evp:
            ws.append([r['site'], r['group'], r['month'], r['norm'], round(r['expected'], 1),
                       round(r['purchased'], 1), round(r['ratio'], 2) if r['ratio'] is not None else None,
                       round(r['gap'], 1), round(r['rate'], 3) if r['rate'] else None,
                       round(r['gap_ro'], 1) if r['gap_ro'] is not None else None,
                       round(r['ratio3'], 2) if r['ratio3'] is not None else None])
        for rng in (f'H2:H{top}', f'G{first}:G{ws.max_row}', f'K{first}:K{ws.max_row}'):
            ws.conditional_formatting.add(rng, CellIsRule(operator='greaterThanOrEqual', formula=['1.5'], fill=RED))
            ws.conditional_formatting.add(rng, CellIsRule(operator='between', formula=['1.2', '1.4999'], fill=AMBER))
            ws.conditional_formatting.add(rng, CellIsRule(operator='lessThanOrEqual', formula=['0.7'], fill=AMBER))
        ws.append([])
        ws.append(['Notes'])
        ws.cell(ws.max_row, 1).font = Font(bold=True)
        for note in [
            'Expected kg = weekly norm x days in month / 7 / usable share. Usable share = 1.0 means the norm is '
            'already in purchased (raw, bone-in) weight; set e.g. 0.65 in consumption_norms.csv if the norm is '
            'cleaned/cooked weight.',
            'Whole chickens bought in pieces are converted with the bird size in the item name (10*1100 = 1.1 kg).',
            'Purchases are not consumption: a stock build-up shows as a high month followed by low months, so '
            'judge on the 3-month ratio. A ratio far above 1 for months on end is not stock.',
            'A ratio below 1 usually means invoices are missing (e.g. Misfa meat, March 2026) or the menu changed.',
        ] + [f'Norm source - {n["site"]} {n["group"]}: {n["source"]}' for n in norms if n['source']] + [
            f'{c} purchase lines for {s} {g} had a unit that could not be converted to kg and were left out'
            for (s, g), c in evp_skipped.items()]:
            ws.append([note])
        widths(ws, [10, 14, 12, 12, 30, 12, 13, 8, 9, 18, 13])
        ws.freeze_panes = 'A2'

    # ---- Price vs Volume
    ws = wb.create_sheet('Price vs Volume')
    ws.append(['Supplier', 'Base avg RO/month'] + [f'{m}\nspend idx' for m in months[-6:]] +
              [f'{m}\nprice idx' for m in months[-6:]] + [f'{m}\nvolume idx' for m in months[-6:]])
    style_header(ws)
    for s in [s for s in suppliers if sum(spend[s][m] for m in base) > 0][:25] + ['ALL SUPPLIERS']:
        if s == 'ALL SUPPLIERS':
            b = sum(sum(spend[x][m] for m in base) for x in suppliers) / len(base)
            sp = [sum(spend[x][m] for x in suppliers) for m in months[-6:]]
            pi = [all_idx[m] for m in months[-6:]]
        else:
            b = sum(spend[s][m] for m in base) / len(base)
            sp = [spend[s][m] for m in months[-6:]]
            pi = [R.index([m], base, lambda k, s=s: k[0] == s) for m in months[-6:]]
        si = [100 * v / b if b else None for v in sp]
        vi = [100 * a / p if a is not None and p else None for a, p in zip(si, pi)]
        ws.append([s, round(b, 1)] + [round(x) if x is not None else None for x in si + pi + vi])
    n = len(months[-6:])
    rng = f'{get_column_letter(3 + n)}2:{get_column_letter(2 + 2 * n)}{ws.max_row}'
    ws.conditional_formatting.add(rng, CellIsRule(operator='greaterThanOrEqual', formula=['120'], fill=RED))
    ws.conditional_formatting.add(rng, CellIsRule(operator='between', formula=['110', '119.99'], fill=AMBER))
    widths(ws, [28, 12] + [10] * (3 * n))
    ws.freeze_panes = 'C2'

    # ---- Price Index (supplier x month)
    ws = wb.create_sheet('Price Index')
    ws.append(['Supplier', 'Base spend RO'] + months)
    style_header(ws)
    ws.append(['ALL SUPPLIERS', None] + [round(all_idx[m]) if all_idx[m] else None for m in months])
    for c in ws[2]:
        c.font = Font(bold=True)
    for s in suppliers:
        b = sum(spend[s][m] for m in base)
        if b <= 0:
            continue
        ws.append([s, round(b)] + [round(v) if v else None for v in
                                   (R.index([m], base, lambda k, s=s: k[0] == s) for m in months)])
    rng = f'C3:{get_column_letter(2 + len(months))}{ws.max_row}'
    ws.conditional_formatting.add(rng, CellIsRule(operator='greaterThanOrEqual', formula=['120'], fill=RED))
    ws.conditional_formatting.add(rng, CellIsRule(operator='between', formula=['110', '119.99'], fill=AMBER))
    widths(ws, [28, 12] + [9] * len(months))
    ws.freeze_panes = 'C2'

    # ---- Category Index
    ws = wb.create_sheet('Category Index')
    ws.append(['Category'] + months)
    style_header(ws)
    for c in cats:
        vals = [R.index([m], base, lambda k, c=c: R.cat[k] == c) for m in months]
        if any(vals):
            ws.append([c] + [round(v) if v else None for v in vals])
    rng = f'B2:{get_column_letter(1 + len(months))}{ws.max_row}'
    ws.conditional_formatting.add(rng, CellIsRule(operator='greaterThanOrEqual', formula=['120'], fill=RED))
    ws.conditional_formatting.add(rng, CellIsRule(operator='between', formula=['110', '119.99'], fill=AMBER))
    widths(ws, [12] + [9] * len(months))

    # ---- Item Rates
    ws = wb.create_sheet('Item Rates')
    ws.append(['Supplier', 'Item', 'UoM', 'Category', 'RO last 6m', 'Base rate'] + months +
              ['vs base', 'vs prev month'])
    style_header(ws)
    keys = []
    for k in R.d:
        v6 = sum(R.d[k][m][1] for m in months[-6:] if m in R.d[k])
        if v6 >= 20:
            keys.append((v6, k))
    for v6, k in sorted(keys, reverse=True):
        br, _ = R.rate(k, base)
        rs = [R.rate(k, [m])[0] for m in months]
        cur = rs[-1] or next((x for x in reversed(rs) if x), None)
        prv = next((x for x in reversed(rs[:-1]) if x), None) if rs[-1] else None
        ws.append([k[0], k[1], k[2], R.cat[k], round(v6), round(br, 3) if br else None] +
                  [round(x, 3) if x else None for x in rs] +
                  [round(cur / br - 1, 3) if br and cur else None,
                   round(rs[-1] / prv - 1, 3) if rs[-1] and prv else None])
    c1, c2 = get_column_letter(7 + len(months)), get_column_letter(8 + len(months))
    for col in (c1, c2):
        ws.conditional_formatting.add(f'{col}2:{col}{ws.max_row}',
                                      CellIsRule(operator='greaterThanOrEqual', formula=['0.2'], fill=RED))
        ws.conditional_formatting.add(f'{col}2:{col}{ws.max_row}',
                                      CellIsRule(operator='between', formula=['0.1', '0.1999'], fill=AMBER))
        for row in ws.iter_rows(min_row=2, min_col=7 + len(months), max_col=8 + len(months)):
            for c in row:
                c.number_format = '+0%;-0%;0%'
    widths(ws, [26, 30, 6, 9, 10, 9] + [8] * len(months) + [9, 11])
    ws.freeze_panes = 'D2'
    ws.auto_filter.ref = ws.dimensions

    # ---- Supplier Spend
    ws = wb.create_sheet('Supplier Spend')
    ws.append(['Supplier'] + months + ['Total', '6m median', f'{last} vs median'])
    style_header(ws)
    for s in suppliers:
        row = [round(spend[s][m], 2) for m in months]
        med = statistics.median(row[-7:-1]) if len(row) > 1 else 0
        ws.append([s] + row + [round(sum(row), 2), round(med, 2), round(row[-1] / med, 2) if med else None])
    ws.append(['TOTAL'] + [round(sum(spend[s][m] for s in suppliers), 2) for m in months])
    for c in ws[ws.max_row]:
        c.font = Font(bold=True)
    widths(ws, [28] + [9] * len(months) + [10, 10, 12])
    ws.freeze_panes = 'B2'

    wb.save(out_path)
    return flags, all_idx


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('ledger')
    ap.add_argument('--system', help='AUW Financial System export (adds rows entered only in the system)')
    ap.add_argument('--sites', default='Misfa,Nizwa')
    ap.add_argument('--base', default='2025-10:2026-02', help='base period FROM:TO (YYYY-MM)')
    ap.add_argument('--to', help='last month to include (YYYY-MM); default = last complete month in data')
    ap.add_argument('--norms', default=os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                                    'consumption_norms.csv'),
                    help='weekly kitchen usage per site & item group (CSV); "" to skip')
    ap.add_argument('--out', default='supplier_tracker.xlsx')
    a = ap.parse_args()
    sites = [s.strip() for s in a.sites.split(',')]
    lines = load_ledger(a.ledger, sites)
    if a.system:
        lines += load_system(a.system, sites)
    if a.to:
        lines = [x for x in lines if x['month'] <= a.to]
    base = month_range(*a.base.split(':'))
    norms = load_norms(a.norms, sites)
    flags, idx = build(lines, base, a.out, sites, norms)
    print(f'Wrote {a.out}: {len(lines)} purchase lines, {len(flags)} flags')


if __name__ == '__main__':
    main()
