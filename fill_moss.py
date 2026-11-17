"""Fill a monthly MOSS workbook from a CULT workbook.

Usage:
    python3 fill_moss.py <cult.xlsx> <template_moss.xlsx> <out.xlsx>

Source: the CULT ' расчет' sheet, which always holds the month the CULT file
is named after (A1/B1 = "EXACTLY" or "Stripe" / the Russian month name,
B2 = the ECB reference date, B3 = the USD rate). The ' Table August' /
' Table September' sheets in the same workbook are stale legacy tables --
do not read them.

Target: the single MOSS sheet of last month's report. Two layouts exist,
told apart by the template's E6 header:

  exactly  (E6 = 'Chargebacks'): sheet 'exactly MM.YYYY', label I5 'MOSS MM/YY'
      D Income, E Chargebacks, F Total Income, G Net Income, H VAT,
      I/J Net Income/VAT in EUR, K Total with Tax.
  Stripe   (E6 = 'Total Income'): sheet 'STRIPE MM.YYYY', label H5 'OSS MM/YY'
      D Income, E Total Income, F Net Income, G VAT,
      H/I Net Income/VAT in EUR, J Total with Tax. No chargeback column.

Column mapping (CULT ' расчет' -> MOSS):
    D (code)            -> B (code)         : join key
    M (USD Income)      -> D (Income)
    O (USD Chargebacks) -> E (Chargebacks), exactly layout only
    everything else in MOSS is derived by formula from D/E/C and the USD rate.

The script refuses to run if CULT reports turnover MOSS has nowhere to put:
refunds, chargebacks in the Stripe layout, or income in RUB or EUR.
"""
import shutil
import sys

import openpyxl

EU_FIRST = 7                  # Austria
CZ_ROW = 33                   # Czech Republic, reported separately
CULT_LAST = 260               # last CULT row that may hold a country

# CULT ' расчет' columns that MOSS has no column for; all must be zero.
UNSUPPORTED = {6: 'RUB income', 7: 'RUB refunds', 8: 'RUB chargebacks',
               14: 'USD refunds', 20: 'EUR income', 21: 'EUR refunds',
               22: 'EUR chargebacks'}

LAYOUTS = {
    'exactly': dict(title='exactly', label='I5', label_text='MOSS',
                    non_eu_last=255, total_eu_row=257, unsupported={}),
    'stripe': dict(title='STRIPE', label='H5', label_text='OSS',
                   non_eu_last=256, total_eu_row=None,
                   unsupported={15: 'USD chargebacks'}),
}


def read_cult(path, unsupported):
    ws = openpyxl.load_workbook(path, data_only=True)[' расчет']
    asof = ws['B2'].value
    rate = ws['B3'].value

    data = {}
    for r in range(EU_FIRST, CULT_LAST + 1):
        code = ws.cell(r, 4).value
        if not code:
            continue
        # A broken formula (e.g. a #REF! in the country lookup) caches an
        # error string; never let it pass as zero or leak into MOSS.
        for col in (5, 13, 15, *unsupported):
            v = ws.cell(r, col).value
            if isinstance(v, str) and v.startswith('#'):
                raise SystemExit(
                    f'{path} row {r} ({code}) column '
                    f'{openpyxl.utils.get_column_letter(col)} is {v}; '
                    'fix the formula in CULT and recalculate'
                )
        for col, label in unsupported.items():
            v = ws.cell(r, col).value
            if isinstance(v, (int, float)) and v:
                raise SystemExit(
                    f'{path} row {r} ({code}) has {label} = {v}; '
                    'MOSS has no column for it'
                )
        data[code] = (ws.cell(r, 13).value or 0,   # M  income USD
                      ws.cell(r, 15).value or 0,   # O  chargebacks USD
                      ws.cell(r, 5).value)         # E  VAT rate
    return asof, rate, data


def detect_layout(ws):
    header = (ws['E6'].value or '').strip()
    if header == 'Chargebacks':
        return 'exactly'
    if header == 'Total Income':
        return 'stripe'
    raise SystemExit(f'unknown MOSS layout: E6 = {header!r}')


def fill(cult_path, template, out):
    shutil.copyfile(template, out)
    wb = openpyxl.load_workbook(out)
    ws = wb.worksheets[0]
    kind = detect_layout(ws)
    lay = LAYOUTS[kind]

    asof, rate, cult = read_cult(cult_path, {**UNSUPPORTED, **lay['unsupported']})

    ws.title = f"{lay['title']} {asof:%m.%Y}"
    ws['A2'] = f'Exchange rate ECB {asof:%d.%m.%Y}'
    ws['B3'] = rate
    ws[lay['label']] = f"{lay['label_text']} {asof:%m/%y}"

    written = 0
    seen = set()
    for r in range(EU_FIRST, lay['non_eu_last'] + 1):
        code = ws.cell(r, 2).value
        if not code:
            continue
        if code not in cult:
            raise SystemExit(f'{code} (MOSS row {r}) is missing from CULT')
        seen.add(code)
        income, chargeback, vat = cult[code]
        ws.cell(r, 4).value = income                      # D Income (USD)
        written += 1

        if r > CZ_ROW:                                    # non-EU: income only
            continue

        if vat != ws.cell(r, 3).value:
            raise SystemExit(f'VAT rate for {code} differs: CULT {vat}, '
                             f'MOSS {ws.cell(r, 3).value}')
        if kind == 'exactly':
            ws.cell(r, 5).value = chargeback or None      # E Chargebacks
            ws.cell(r, 6).value = f'=D{r}-E{r}'           # F Total Income
            first = 7                                     # G Net Income
        else:
            ws.cell(r, 5).value = f'=D{r}'                # E Total Income
            first = 6                                     # F Net Income
        net, vat_c, net_eur, vat_eur, total = (
            openpyxl.utils.get_column_letter(first + i) for i in range(5))
        tot_inc = 'F' if kind == 'exactly' else 'E'
        ws[f'{net}{r}'] = f'={tot_inc}{r}/(100+C{r})*100'
        ws[f'{vat_c}{r}'] = f'={net}{r}*C{r}/100'
        ws[f'{net_eur}{r}'] = f'={net}{r}/$B$3'
        ws[f'{vat_eur}{r}'] = f'={vat_c}{r}/$B$3'
        ws[f'{total}{r}'] = f'={net_eur}{r}+{vat_eur}{r}'

    lost = {c: v[0] for c, v in cult.items() if c not in seen and v[0]}
    if lost:
        raise SystemExit(f'CULT income for codes absent from MOSS: {lost}')

    # The shipped exactly template summed E:H of the Total EU row from row 8,
    # silently excluding Austria; sum from row 7 like D/I/J/K do.
    if lay['total_eu_row']:
        for col in 'EFGH':
            ws[f"{col}{lay['total_eu_row']}"] = \
                f'=SUM({col}{EU_FIRST}:{col}{CZ_ROW - 1})'

    wb.save(out)
    total = sum(v[0] for v in cult.values())
    print(f'{out}: {kind}, {written} countries, {asof:%Y-%m-%d}, '
          f'USD rate {rate}, income {total:.2f}')


if __name__ == '__main__':
    fill(*sys.argv[1:4])
