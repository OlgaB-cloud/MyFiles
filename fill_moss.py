"""Fill a monthly MOSS workbook from a CULT workbook.

Usage:
    python3 fill_moss.py <cult.xlsx> <template_moss.xlsx> <out.xlsx>

Source: the CULT ' расчет' sheet, which always holds the month the CULT file
is named after (A1/B1 = "EXACTLY" / the Russian month name, B2 = the ECB
reference date, B3 = the USD rate). The ' Table August' / ' Table September'
sheets in the same workbook are stale legacy tables -- do not read them.

Target: the single MOSS sheet, renamed 'exactly MM.YYYY' for the new month.

Column mapping (CULT ' расчет' -> MOSS):
    D (code)            -> B (code)         : join key
    M (USD Income)      -> D (Income)
    O (USD Chargebacks) -> E (Chargebacks)
    everything else in MOSS is derived by formula from D/E/C and the USD rate.

The script refuses to run if CULT reports turnover MOSS has nowhere to put:
refunds, or income in RUB or EUR.
"""
import shutil
import sys

import openpyxl

EU_FIRST = 7                  # Austria
CZ_ROW = 33                   # Czech Republic, reported separately
NON_EU_LAST = 255
TOTAL_EU_ROW = 257

# CULT ' расчет' columns that MOSS has no column for; all must be zero.
UNSUPPORTED = {6: 'RUB income', 7: 'RUB refunds', 8: 'RUB chargebacks',
               14: 'USD refunds', 20: 'EUR income', 21: 'EUR refunds',
               22: 'EUR chargebacks'}


def read_cult(path):
    ws = openpyxl.load_workbook(path, data_only=True)[' расчет']
    asof = ws['B2'].value
    rate = ws['B3'].value

    data = {}
    for r in range(EU_FIRST, 259):
        code = ws.cell(r, 4).value
        if not code:
            continue
        for col, label in UNSUPPORTED.items():
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


def fill(cult_path, template, out):
    asof, rate, cult = read_cult(cult_path)

    shutil.copyfile(template, out)
    wb = openpyxl.load_workbook(out)
    ws = wb.worksheets[0]

    ws.title = f'exactly {asof:%m.%Y}'
    ws['A2'] = f'Exchange rate ECB {asof:%d.%m.%Y}'
    ws['B3'] = rate
    ws['I5'] = f'MOSS {asof:%m/%y}'

    written = 0
    for r in range(EU_FIRST, NON_EU_LAST + 1):
        code = ws.cell(r, 2).value
        if not code:
            continue
        if code not in cult:
            raise SystemExit(f'{code} (MOSS row {r}) is missing from CULT')
        income, chargeback, vat = cult[code]
        ws.cell(r, 4).value = income                      # D Income (USD)
        written += 1

        if r > CZ_ROW:                                    # non-EU: income only
            continue

        if vat != ws.cell(r, 3).value:
            raise SystemExit(f'VAT rate for {code} differs: CULT {vat}, '
                             f'MOSS {ws.cell(r, 3).value}')
        ws.cell(r, 5).value = chargeback or None          # E Chargebacks
        ws.cell(r, 6).value = f'=D{r}-E{r}'               # F Total Income
        ws.cell(r, 7).value = f'=F{r}/(100+C{r})*100'     # G Net Income
        ws.cell(r, 8).value = f'=G{r}*C{r}/100'           # H VAT
        ws.cell(r, 9).value = f'=G{r}/$B$3'               # I Net Income Total (EUR)
        ws.cell(r, 10).value = f'=H{r}/$B$3'              # J VAT Total (EUR)
        ws.cell(r, 11).value = f'=I{r}+J{r}'              # K Total with Tax

    # The shipped template summed E:H of the Total EU row from row 8, silently
    # excluding Austria; sum from row 7 like D/I/J/K do.
    for col in 'EFGH':
        ws[f'{col}{TOTAL_EU_ROW}'] = f'=SUM({col}{EU_FIRST}:{col}{CZ_ROW - 1})'

    wb.save(out)
    print(f'{out}: {written} countries, {asof:%Y-%m-%d}, USD rate {rate}')


if __name__ == '__main__':
    fill(*sys.argv[1:4])
