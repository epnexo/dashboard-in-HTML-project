"""One-off helper: turns the original (Nigerian) sample file into an Australian-looking demo file.

    python scripts/make_demo_au.py

Reads  data/source/telecoms_transactions_dirty.csv
Writes data/raw/telecom_transactions_au_dirty.csv

It keeps every data-quality problem of the original (mixed date formats, aliases,
symbols, blanks, duplicates) and only changes what makes it look Nigerian:
  - regions: Nigerian cities and their aliases -> Australian cities and aliases
  - amounts: divided by 100 and written with $ / AUD instead of the naira sign / NGN
  - "Airtime N" bundles -> "Recharge $N/100"; product "Airtime" -> "Recharge"
This is demo data, not real Australian figures. It is not part of the normal run.
"""
import csv
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / 'data' / 'source' / 'telecoms_transactions_dirty.csv'
DST = ROOT / 'data' / 'raw' / 'telecom_transactions_au_dirty.csv'

REGIONS = {
    'lagos': 'sydney', 'Lagos': 'Sydney', 'Lasgidi': 'Harbour City', 'Lagos State': 'Sydney NSW', 'LAG': 'SYD',
    'Abuja': 'Melbourne', 'FCT': 'VIC', 'Abuja FCT': 'Melbourne VIC', 'ABV': 'MEL',
    'kano': 'brisbane', 'Kano State': 'Brisbane QLD', 'KAN': 'BNE',
    'Ibadan': 'Adelaide', 'IB': 'ADL', 'Ibadan Oyo': 'Adelaide SA',
    'Kaduna': 'Canberra', 'KAD': 'CBR',
    'port harcourt': 'perth', 'Port-Harcourt': 'Perth-WA', 'PHC': 'PER', 'PH': 'PTH',
    'enugu': 'hobart', 'ENU': 'HBA',
}
LOWER = {k.lower(): v for k, v in REGIONS.items()}


def keep_spaces(raw, new):
    """Puts back the leading/trailing spaces of the original value."""
    core = raw.strip()
    return raw.replace(core, new) if core else raw


def amount(raw):
    s = raw.strip()
    if not s:
        return raw
    value = float(re.sub(r'[^\d.]', '', s)) / 100
    text = f'{value:,.2f}' if ('.' in s or value != int(value)) else f'{value:,.0f}'
    if '₦' in s:
        new = '$' + text
    elif 'NGN' in s:
        new = ('-' if s.startswith('-') else '') + 'AUD ' + text
    elif s.startswith('('):
        new = f'({text})'
    else:
        new = ('-' if s.startswith('-') else '') + text
    return keep_spaces(raw, new)


with open(SRC, encoding='utf-8-sig', newline='') as f:
    reader = csv.DictReader(f)
    rows, header = list(reader), reader.fieldnames

for r in rows:
    region = r['region'].strip()
    if region:  # exact spelling first; otherwise match ignoring case and keep the original's capitalisation style
        new = REGIONS.get(region) or LOWER[region.lower()]
        r['region'] = keep_spaces(r['region'], new if region in REGIONS else (new.title() if region[0].isupper() else new.lower()))
    r['amount'] = amount(r['amount'])
    if m := re.fullmatch(r'Airtime (\d+)', r['bundle_name'].strip()):
        r['bundle_name'] = f'Recharge ${int(m[1]) // 100}'
    if r['product_type'].strip() == 'Airtime':
        r['product_type'] = 'Recharge'

DST.parent.mkdir(parents=True, exist_ok=True)
with open(DST, 'w', encoding='utf-8', newline='') as f:
    w = csv.DictWriter(f, fieldnames=header)
    w.writeheader()
    w.writerows(rows)
print(f'{len(rows)} rows written to {DST.relative_to(ROOT)}')
