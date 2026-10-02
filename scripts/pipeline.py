"""Telecom transactions: cleaning + HTML and Excel reports.

Usage:
    python scripts/pipeline.py                   # newest CSV in data/raw
    python scripts/pipeline.py --input path.csv  # a specific CSV
    python scripts/pipeline.py --open            # open the report when done

Outputs (relative to the project folder):
    data/clean/<name>_clean.csv        clean data
    cleaning/cleaning_report.html      what was found and what was done
    cleaning/findings_detail.csv       every correction, row by row
    cleaning/cleaning_summary.json     counts
    reports/telecom_report.html        dashboard with KPIs and charts
    reports/telecom_report.xlsx        the same report in Excel (scripts/excel.py, needs openpyxl)
    runs/history.csv + .log            record of every run
Except for the Excel file, only the Python standard library is used.
"""
import argparse
import base64
import csv
import html
import json
import re
import sys
import time
import webbrowser
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / 'scripts'
RAW_DIR, CLEAN_DIR = ROOT / 'data' / 'raw', ROOT / 'data' / 'clean'
CLEANING_DIR, REPORTS_DIR, RUNS_DIR = ROOT / 'cleaning', ROOT / 'reports', ROOT / 'runs'
REPORT = REPORTS_DIR / 'telecom_report.html'
EXCEL = REPORTS_DIR / 'telecom_report.xlsx'
CLEANING_REPORT = CLEANING_DIR / 'cleaning_report.html'
LOGO = SCRIPTS / 'logo.png'

COLUMNS = ['transaction_id', 'txn_date', 'customer_id', 'region', 'customer_segment', 'product_type',
           'bundle_name', 'channel', 'data_volume_mb', 'amount', 'status']
MONTHS = {m: i + 1 for i, m in enumerate('jan feb mar apr may jun jul aug sep oct nov dec'.split())}


# ---------- value parsing ----------
def parse_date(raw):
    """Returns (ISO date or None, detected format)."""
    s = raw.strip()
    if not s:
        return None, 'empty'
    try:
        if m := re.fullmatch(r'(\d{2})/(\d{2})/(\d{4})', s):
            return date(int(m[3]), int(m[2]), int(m[1])).isoformat(), 'DD/MM/YYYY'
        if m := re.fullmatch(r'(\d{2})-(\d{2})-(\d{4})', s):
            return date(int(m[3]), int(m[1]), int(m[2])).isoformat(), 'MM-DD-YYYY'
        if m := re.fullmatch(r'(\d{4})\.(\d{2})\.(\d{2})', s):
            return date(int(m[1]), int(m[2]), int(m[3])).isoformat(), 'YYYY.MM.DD'
        if m := re.fullmatch(r'(\d{4})-(\d{2})-(\d{2})', s):
            return date(int(m[1]), int(m[2]), int(m[3])).isoformat(), 'YYYY-MM-DD'
        if m := re.fullmatch(r'(\d{1,2}) ([A-Za-z]{3}) (\d{4})', s):
            return date(int(m[3]), MONTHS[m[2].lower()], int(m[1])).isoformat(), 'DD Mon YYYY'
        if re.fullmatch(r'\d{5}', s):  # Excel serial (epoch 1899-12-30)
            return (date(1899, 12, 30) + timedelta(days=int(s))).isoformat(), 'Excel serial'
    except (ValueError, KeyError):
        return None, 'invalid'
    return None, 'invalid'


def parse_amount(raw):
    s = raw.strip()
    if not s:
        return None
    digits = re.sub(r'[^\d.]', '', s)
    if not digits:
        return None
    n = float(digits)
    return -n if s[0] in '(-' else n


def fmt_num(n):
    return str(int(n)) if float(n).is_integer() else str(n)


# ---------- cleaning ----------
def clean(source):
    rules = json.loads((SCRIPTS / 'rules.json').read_text(encoding='utf-8'))
    regions, channels = rules['regions'], rules['channels']

    with open(source, encoding='utf-8-sig', newline='') as f:
        raw_rows = list(csv.DictReader(f))
    if not raw_rows:
        raise ValueError('The file has no rows.')
    missing = [c for c in COLUMNS if c not in raw_rows[0]]
    if missing:
        raise ValueError('Missing columns in the file: ' + ', '.join(missing))

    # List price: most frequent amount of each bundle among successful transactions
    counts = defaultdict(Counter)
    for r in raw_rows:
        m = parse_amount(r['amount'] or '')
        if (r['status'] or '').strip() == 'Successful' and m is not None:
            counts[(r['bundle_name'] or '').strip()][abs(m)] += 1
    price = {b: c.most_common(1)[0][0] for b, c in counts.items()}

    log = {'file': Path(source).name, 'rows_in': len(raw_rows), 'duplicates_removed': 0, 'rows_out': 0,
           'date_formats': {}, 'dates_missing': 0, 'dates_invalid': 0, 'customer_id_spaces': 0,
           'customer_id_no_prefix': 0, 'regions_normalized': 0, 'regions_empty': 0, 'channels_normalized': 0,
           'amounts_reformatted': 0, 'amounts_imputed': 0, 'amounts_outliers': 0, 'outliers': []}
    findings = []  # (row, id, column, issue, original, corrected, action)
    seen, out = set(), []

    for n, r in enumerate(raw_rows, start=2):  # 2 = first data row in the file
        o = {c: (r[c] or '') for c in COLUMNS}
        tid = o['transaction_id'].strip()

        def add(col, issue, corrected, action, original=None):
            findings.append((n, tid, col, issue, o[col] if original is None else original, corrected, action))

        if tid in seen:
            log['duplicates_removed'] += 1
            findings.append((n, tid, 'row', 'Duplicate row', tid, '', 'Row removed (the first one is kept)'))
            continue
        seen.add(tid)

        iso, fmt = parse_date(o['txn_date'])
        log['date_formats'][fmt] = log['date_formats'].get(fmt, 0) + 1
        if fmt == 'empty':
            log['dates_missing'] += 1
            add('txn_date', 'Empty date', '', 'Left without a date; only excluded from monthly series')
        elif fmt == 'invalid':
            log['dates_missing'] += 1
            log['dates_invalid'] += 1
            add('txn_date', 'Impossible or unrecognised date', '', 'Left without a date; only excluded from monthly series')
        elif iso != o['txn_date']:
            add('txn_date', 'Non-standard date format', iso, f'Converted to YYYY-MM-DD (read as {fmt})')

        cust = o['customer_id'].strip()
        if cust != o['customer_id']:
            log['customer_id_spaces'] += 1
            add('customer_id', 'Customer ID with extra spaces', cust, 'Spaces removed')
        if cust.isdigit():
            log['customer_id_no_prefix'] += 1
            add('customer_id', 'Customer ID without CUST prefix', 'CUST' + cust, 'CUST prefix added')
            cust = 'CUST' + cust

        key = o['region'].strip().lower()
        if not key:
            region = 'Unknown'
            log['regions_empty'] += 1
            add('region', 'Empty region', region, 'Marked as "Unknown"')
        else:
            region = regions.get(key) or o['region'].strip().title()
            if region != o['region']:
                log['regions_normalized'] += 1
                extra = '' if key in regions else ' (alias not listed in rules.json)'
                add('region', 'Region written as alias, abbreviation, different case or with spaces', region, 'Unified' + extra)

        key = o['channel'].strip().lower()
        channel = channels.get(key) or (o['channel'].strip().title() or 'Unknown')
        if channel != o['channel']:
            log['channels_normalized'] += 1
            extra = '' if key in channels else ' (alias not listed in rules.json)'
            add('channel', 'Channel written with variants', channel, 'Unified' + extra)

        bundle, status = o['bundle_name'].strip(), o['status'].strip()
        lst = price.get(bundle)
        expected = None if lst is None else 0.0 if status == 'Failed' else -lst if status == 'Reversed' else lst
        amount, flag = parse_amount(o['amount']), ''
        if amount is None:
            amount, flag = (expected or 0.0), 'imputed'
            log['amounts_imputed'] += 1
            add('amount', 'Empty amount', fmt_num(amount), "Imputed with the bundle's list price according to status")
        else:
            if fmt_num(amount) != o['amount']:
                log['amounts_reformatted'] += 1
                add('amount', 'Amount with symbols, commas, spaces or parentheses', fmt_num(amount), 'Converted to a number')
            if expected is not None and amount != expected:
                log['amounts_outliers'] += 1
                log['outliers'].append({'transaction_id': tid, 'bundle': bundle, 'status': status,
                                        'original_amount': amount, 'corrected_amount': expected})
                add('amount', 'Outlier amount (does not match the list price)', fmt_num(expected),
                    'Replaced with the list price', original=fmt_num(amount))
                amount, flag = expected, 'outlier_corrected'

        out.append({'transaction_id': tid, 'txn_date': iso or '', 'customer_id': cust, 'region': region,
                    'customer_segment': o['customer_segment'].strip(), 'product_type': o['product_type'].strip(),
                    'bundle_name': bundle, 'channel': channel, 'data_volume_mb': o['data_volume_mb'].strip(),
                    'amount': fmt_num(amount), 'status': status, 'amount_flag': flag})

    out.sort(key=lambda r: (r['txn_date'] or '9999', r['transaction_id']))
    log['rows_out'] = len(out)
    return out, log, findings


# ---------- outputs ----------
def write_csv(path, header, rows):
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def group_findings(findings):
    """Summarises findings by (column, issue), from most to fewest rows."""
    groups = {}
    for _, _, col, issue, original, _, action in findings:
        g = groups.setdefault((col, issue), {'column': col, 'issue': issue, 'n': 0, 'examples': [], 'action': ''})
        g['n'] += 1
        g['action'] = action.split(' (')[0]
        v = original if original.strip() == original and original else f'"{original}"'
        if v not in g['examples'] and len(g['examples']) < 5:
            g['examples'].append(v)
    return sorted(groups.values(), key=lambda g: -g['n'])


def write_cleaning_report(log, findings, when):
    e = html.escape
    rows = ''.join(
        f'<tr><td>{e(g["column"])}</td><td>{e(g["issue"])}</td><td class="n">{g["n"]:,}</td>'
        f'<td class="ej">{e(" · ".join(g["examples"]))}</td><td>{e(g["action"])}</td></tr>'
        for g in group_findings(findings))
    outliers = ''.join(
        f'<tr><td>{e(a["transaction_id"])}</td><td>{e(a["bundle"])}</td><td>{e(a["status"])}</td>'
        f'<td class="n">{fmt_num(a["original_amount"])}</td><td class="n">{fmt_num(a["corrected_amount"])}</td></tr>'
        for a in log['outliers']) or '<tr><td colspan="5">None</td></tr>'
    formats = ''.join(f'<tr><td>{e(k)}</td><td class="n">{v:,}</td></tr>' for k, v in sorted(log['date_formats'].items(), key=lambda kv: -kv[1]))
    doc = f"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Cleaning report</title><style>
body{{font:14px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif;margin:0;background:#f4f6f8;color:#0b1220}}
.w{{max-width:1100px;margin:0 auto;padding:28px 16px 48px}} h1{{font-size:24px;margin:0 0 4px}} h2{{font-size:16px;margin:28px 0 10px}}
p{{color:#4a5568;margin:0 0 8px}} .kp{{display:flex;flex-wrap:wrap;gap:12px;margin:18px 0}}
.k{{background:#fff;border:1px solid rgba(9,18,41,.1);border-radius:12px;padding:14px 18px;min-width:160px}}
.k b{{display:block;font-size:24px}} .k span{{color:#4a5568;font-size:12.5px}}
.t{{background:#fff;border:1px solid rgba(9,18,41,.1);border-radius:12px;overflow-x:auto}}
table{{width:100%;border-collapse:collapse;font-size:13px}} th{{text-align:left;color:#4a5568;font-weight:500;font-size:12px}}
th,td{{padding:8px 12px;border-bottom:1px solid #e4e8ee;vertical-align:top}} tr:last-child td{{border-bottom:0}}
.n{{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}} .ej{{color:#4a5568;font-family:Consolas,monospace;font-size:12px}}
</style></head><body><div class="w">
<h1>Cleaning report</h1>
<p>File: <b>{e(log['file'])}</b> · Run: {when}</p>
<div class="kp">
<div class="k"><b>{log['rows_in']:,}</b><span>rows in the original file</span></div>
<div class="k"><b>{log['duplicates_removed']:,}</b><span>duplicates removed</span></div>
<div class="k"><b>{log['rows_out']:,}</b><span>rows in the clean file</span></div>
<div class="k"><b>{len(findings):,}</b><span>corrections recorded</span></div>
</div>
<h2>What was found and what was done</h2>
<div class="t"><table><thead><tr><th>Column</th><th>Issue found</th><th class="n">Rows</th><th>Examples of the original value</th><th>What was done</th></tr></thead><tbody>{rows}</tbody></table></div>
<h2>Date formats detected</h2>
<p>Dates with slashes are read as day/month and dates with dashes as month-day, as shown by the values above 12 in each format.</p>
<div class="t"><table><thead><tr><th>Format</th><th class="n">Rows</th></tr></thead><tbody>{formats}</tbody></table></div>
<h2>Outlier amounts corrected</h2>
<p>Each bundle's list price is its most frequent amount among successful transactions. Amounts that do not match are replaced with that price.</p>
<div class="t"><table><thead><tr><th>Transaction</th><th>Bundle</th><th>Status</th><th class="n">Original amount</th><th class="n">Corrected amount</th></tr></thead><tbody>{outliers}</tbody></table></div>
<h2>Detail</h2>
<p>Every correction, row by row (row number in the original file, value before and after), is in <b>findings_detail.csv</b>, in this same folder.</p>
</div></body></html>"""
    CLEANING_REPORT.write_text(doc, encoding='utf-8')


def load_settings():
    """Currency, targets and thresholds (scripts/settings.json); keys starting with _ are notes."""
    path = SCRIPTS / 'settings.json'
    if not path.exists():
        return {}
    return {k: v for k, v in json.loads(path.read_text(encoding='utf-8')).items() if not k.startswith('_')}


def build_report(out, log):
    dim_cols = {'cust': 'customer_id', 'region': 'region', 'seg': 'customer_segment', 'prod': 'product_type',
                'bundle': 'bundle_name', 'chan': 'channel', 'status': 'status'}
    dims = {k: sorted({r[c] for r in out}) for k, c in dim_cols.items()}
    idx = {k: {v: i for i, v in enumerate(vals)} for k, vals in dims.items()}

    def num(s):
        f = float(s) if s else 0
        return int(f) if f.is_integer() else f

    rows = [[r['txn_date'], idx['cust'][r['customer_id']], idx['region'][r['region']], idx['seg'][r['customer_segment']],
             idx['prod'][r['product_type']], idx['bundle'][r['bundle_name']], idx['chan'][r['channel']],
             num(r['data_volume_mb']), num(r['amount']), idx['status'][r['status']]] for r in out]
    data = json.dumps({'dims': dims, 'rows': rows, 'log': log, 'settings': load_settings()},
                      ensure_ascii=False, separators=(',', ':')).replace('<', '\\u003c')
    template = (SCRIPTS / 'report_template.html').read_text(encoding='utf-8')
    logo = ''  # scripts/logo.png is embedded so the HTML stays a single file
    if LOGO.exists():
        logo = 'data:image/png;base64,' + base64.b64encode(LOGO.read_bytes()).decode('ascii')
    REPORT.write_text(template.replace('__LOGO__', logo).replace('/*__DATA__*/null', data), encoding='utf-8')


def latest_csv():
    files = sorted(RAW_DIR.glob('*.csv'), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files:
        raise FileNotFoundError(f'No .csv files in {RAW_DIR}')
    return files[0]


def run(source=None, notify=print):
    """Runs the whole flow. `notify` receives each progress line (the UI shows it on screen)."""
    start, when = time.time(), datetime.now()
    for d in (RAW_DIR, CLEAN_DIR, CLEANING_DIR, REPORTS_DIR, RUNS_DIR):
        d.mkdir(parents=True, exist_ok=True)
    lines = []

    def step(txt):
        lines.append(f'[{datetime.now():%H:%M:%S}] {txt}')
        notify(txt)

    hist = RUNS_DIR / 'history.csv'
    new = not hist.exists()
    record = [f'{when:%Y-%m-%d %H:%M:%S}', '', '', '', '', '', '', '', 'ERROR', '']
    try:
        source = Path(source) if source else latest_csv()
        record[1] = source.name
        step(f'1/4 Reading and cleaning {source.name}')
        out, log, findings = clean(source)
        clean_path = CLEAN_DIR / (re.sub(r'_?dirty$', '', source.stem) + '_clean.csv')
        write_csv(clean_path, list(out[0].keys()), [list(r.values()) for r in out])
        step(f'    {log["rows_in"]:,} rows read, {log["duplicates_removed"]} duplicates removed, {log["rows_out"]:,} clean rows')

        step('2/4 Writing cleaning report')
        write_csv(CLEANING_DIR / 'findings_detail.csv',
                  ['original_row', 'transaction_id', 'column', 'issue', 'original_value', 'corrected_value', 'action'], findings)
        (CLEANING_DIR / 'cleaning_summary.json').write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding='utf-8')
        write_cleaning_report(log, findings, f'{when:%Y-%m-%d %H:%M}')
        step(f'    {len(findings):,} corrections recorded')

        step('3/4 Building HTML report')
        build_report(out, log)

        try:
            from excel import build_excel
            build_excel(out, log, group_findings(findings), EXCEL, LOGO, load_settings())
            step('    Excel report built')
        except ImportError:
            step('    WARNING: Excel not built; openpyxl is missing (python -m pip install openpyxl)')
        except PermissionError:
            step(f'    WARNING: Excel not updated; close {EXCEL.name} and run again')

        ok = [r for r in out if r['status'] == 'Successful']
        revenue = sum(float(r['amount']) for r in ok)
        summary = {'input': str(source), 'clean': str(clean_path), 'report': str(REPORT), 'cleaning_report': str(CLEANING_REPORT),
                   'rows_in': log['rows_in'], 'rows_out': log['rows_out'], 'duplicates': log['duplicates_removed'],
                   'no_date': log['dates_missing'], 'amounts_imputed': log['amounts_imputed'],
                   'amounts_outliers': log['amounts_outliers'], 'corrections': len(findings), 'revenue': round(revenue),
                   'successful_transactions': len(ok), 'success_rate_pct': round(len(ok) / len(out) * 100, 1),
                   'customers': len({r['customer_id'] for r in out})}
        record[2:] = [log['rows_in'], log['rows_out'], log['duplicates_removed'], len(findings),
                      round(revenue), f'{time.time() - start:.1f}', 'OK', '']
        step(f'4/4 Done in {time.time() - start:.1f} s')
        step(f'    Revenue {revenue:,.2f} | success {summary["success_rate_pct"]}% | {summary["customers"]} customers')
        step(f'    Report:   {REPORT.relative_to(ROOT)}')
        step(f'    Excel:    {EXCEL.relative_to(ROOT)}')
        step(f'    Cleaning: {CLEANING_REPORT.relative_to(ROOT)}')
        return summary
    except Exception as ex:
        record[9] = str(ex)
        step(f'ERROR: {ex}')
        raise
    finally:
        with open(hist, 'a', encoding='utf-8-sig', newline='') as f:
            w = csv.writer(f)
            if new:
                w.writerow(['datetime', 'input_file', 'rows_in', 'rows_out', 'duplicates',
                            'corrections', 'revenue', 'seconds', 'status', 'error'])
            w.writerow(record)
        (RUNS_DIR / f'{when:%Y-%m-%d_%H%M%S}.log').write_text('\n'.join(lines) + '\n', encoding='utf-8')


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    ap = argparse.ArgumentParser(description='Cleans the transactions CSV and builds the HTML and Excel reports.')
    ap.add_argument('--input', help='CSV to process (default: the newest one in data/raw)')
    ap.add_argument('--open', action='store_true', help='open the report in the browser when done')
    a = ap.parse_args()
    try:
        run(a.input)
    except Exception:
        sys.exit(1)
    if a.open:
        webbrowser.open(REPORT.as_uri())
