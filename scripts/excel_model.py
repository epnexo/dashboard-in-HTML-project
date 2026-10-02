"""Excel report built the Power Query + Power Pivot way.

Drives the installed Excel (pywin32) to create a workbook with:
  - Power Query: SourceFile parameter, Transactions and Calendar queries
  - Data Model (Power Pivot): both tables, a relationship and DAX measures
  - Dashboard: KPI cards (CUBEVALUE), PivotCharts and slicers wired to every pivot
  - Detail, Pivots, Cleaning and Read me sheets
After a new run of the pipeline, or after editing the SourceFile parameter, use
Data > Refresh All inside Excel. Needs Excel 2016+ on Windows and `pip install pywin32`.
"""
import os
import shutil
import signal
import time
import tempfile
from pathlib import Path

NAVY, TEAL, GREEN, RED, ORANGE, GREY, LINE = '0F1D37', '0D9488', '0CA30C', 'D03B3B', 'EC835A', '4A5568', 'E4E8EE'
MODEL = 'ThisWorkbookDataModel'

# Excel constants
XL_EXTERNAL, XL_CMD_SQL, XL_ROW, XL_COLUMN, XL_DESC, XL_XLSX = 2, 2, 1, 2, 2, 51
XL_LINE_MARKERS, XL_BAR, XL_COLUMN_CHART, XL_BAR_STACKED_100 = 65, 57, 51, 59
XL_CENTER, XL_LEFT, XL_RIGHT, XL_TOP = -4108, -4131, -4152, -4160


def rgb(hex_color):
    """Excel wants BGR integers."""
    return int(hex_color[4:6] + hex_color[2:4] + hex_color[0:2], 16)


def m_queries(csv_path):
    path = str(csv_path).replace('"', '""')
    source = f'"{path}" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]'
    transactions = '''let
    Source = Csv.Document(File.Contents(SourceFile), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Headers, {{"txn_date", type date}, {"data_volume_mb", Int64.Type}, {"amount", type number}}, "en-US"),
    ListPrices = Table.Group(Table.SelectRows(Typed, each [status] = "Successful"), {"bundle_name"}, {{"list_price", each List.Max([amount]), type number}}),
    Merged = Table.NestedJoin(Typed, {"bundle_name"}, ListPrices, {"bundle_name"}, "prices", JoinKind.LeftOuter),
    WithPrice = Table.ExpandTableColumn(Merged, "prices", {"list_price"})
in
    WithPrice'''
    calendar = '''let
    Dates = List.RemoveNulls(Transactions[txn_date]),
    First = Date.StartOfYear(List.Min(Dates)),
    Last = Date.EndOfYear(List.Max(Dates)),
    Days = List.Dates(First, Duration.Days(Last - First) + 1, #duration(1, 0, 0, 0)),
    AsTable = Table.FromList(Days, Splitter.SplitByNothing(), type table [Date = date]),
    Year = Table.AddColumn(AsTable, "Year", each Date.Year([Date]), Int64.Type),
    Quarter = Table.AddColumn(Year, "Quarter", each "Q" & Text.From(Date.QuarterOfYear([Date])), type text),
    MonthNo = Table.AddColumn(Quarter, "MonthNo", each Date.Month([Date]), Int64.Type),
    YearMonth = Table.AddColumn(MonthNo, "YearMonth", each Date.ToText([Date], "yyyy-MM"), type text),
    Weekday = Table.AddColumn(YearMonth, "Weekday", each Text.From(Date.DayOfWeek([Date], Day.Monday) + 1) & " " & Date.ToText([Date], "ddd", "en-US"), type text)
in
    Weekday'''
    return source, transactions, calendar


MEASURES = [  # name, DAX, format (money / int / pct), description
    ('Revenue', 'CALCULATE(SUM(Transactions[amount]), Transactions[status] = "Successful")', 'money', 'Amounts of successful transactions'),
    ('Transaction Count', 'COUNTROWS(Transactions)', 'int', 'All statuses'),
    ('Successful Txns', 'CALCULATE(COUNTROWS(Transactions), Transactions[status] = "Successful")', 'int', ''),
    ('Failed Txns', 'CALCULATE(COUNTROWS(Transactions), Transactions[status] = "Failed")', 'int', ''),
    ('Reversed Txns', 'CALCULATE(COUNTROWS(Transactions), Transactions[status] = "Reversed")', 'int', ''),
    ('Success Rate', 'DIVIDE([Successful Txns], [Transaction Count])', 'pct', ''),
    ('Failure Rate', 'DIVIDE([Failed Txns], [Transaction Count])', 'pct', ''),
    ('Reversal Rate', 'DIVIDE([Reversed Txns], [Transaction Count])', 'pct', ''),
    ('Average Ticket', 'DIVIDE([Revenue], [Successful Txns])', 'money', 'Revenue per successful transaction'),
    ('Active Customers', 'DISTINCTCOUNT(Transactions[customer_id])', 'int', 'Customers with any transaction'),
    ('Buying Customers', 'CALCULATE(DISTINCTCOUNT(Transactions[customer_id]), Transactions[status] = "Successful")', 'int', ''),
    ('Revenue per Customer', 'DIVIDE([Revenue], [Buying Customers])', 'money', 'ARPU'),
    ('Revenue Lost', 'CALCULATE(SUM(Transactions[list_price]), Transactions[status] = "Failed")', 'money', 'Failed transactions valued at the bundle list price'),
    ('Reversed Amount', '-CALCULATE(SUM(Transactions[amount]), Transactions[status] = "Reversed")', 'money', ''),
    ('Revenue Prev Month', 'CALCULATE([Revenue], DATEADD(Calendar[Date], -1, MONTH))', 'money', ''),
    ('Revenue MoM %', 'DIVIDE([Revenue] - [Revenue Prev Month], [Revenue Prev Month])', 'pct', 'Change vs. previous month'),
]


def build_excel_model(clean_csv, out, log, groups, target, logo=None, settings=None):
    import pythoncom
    import win32com.client as win32
    import win32process

    settings = settings or {}
    symbol = settings.get('currency_symbol', '$')
    money_fmt = f'"{symbol}"#,##0.00'
    target = Path(target)

    pythoncom.CoInitialize()
    xl = win32.DispatchEx('Excel.Application')
    xl.Visible, xl.DisplayAlerts, xl.ScreenUpdating = False, False, False
    pid = win32process.GetWindowThreadProcessId(xl.Hwnd)[1]  # to make sure this hidden Excel never lingers
    wb = None
    try:
        wb = xl.Workbooks.Add()
        while wb.Worksheets.Count > 1:
            wb.Worksheets(wb.Worksheets.Count).Delete()
        dash = wb.Worksheets(1)
        dash.Name = 'Dashboard'
        add = lambda name: wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))  # noqa: E731
        detail, pivots, cln, readme = add(0), add(0), add(0), add(0)
        detail.Name, pivots.Name, cln.Name, readme.Name = 'Detail', 'Pivots', 'Cleaning', 'Read me'

        # ---------- Power Query + Data Model ----------
        source, transactions, calendar = m_queries(clean_csv)
        wb.Queries.Add('SourceFile', source)
        wb.Queries.Add('Transactions', transactions)
        wb.Queries.Add('Calendar', calendar)
        for q in ('Transactions', 'Calendar'):
            # command type 6 (table collection) with the quoted query name gives the model table that same name
            wb.Connections.Add2(f'Query - {q}', f"Connection to the '{q}' query in the workbook.",
                                f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={q};Extended Properties=""',
                                f'"{q}"', 6, True, False)
        model = wb.Model
        model.Refresh()
        tables = {mt.Name: mt for mt in model.ModelTables}
        if 'Transactions' not in tables or 'Calendar' not in tables:
            raise RuntimeError(f'Data model tables not loaded: found {sorted(tables)}')
        t, c = tables['Transactions'], tables['Calendar']
        model.ModelRelationships.Add(t.ModelTableColumns('txn_date'), c.ModelTableColumns('Date'))

        def fmt(kind):
            if kind == 'pct':
                f = model.ModelFormatPercentageNumber
                f.DecimalPlaces = 1
            elif kind == 'int':
                f = model.ModelFormatWholeNumber
                f.UseThousandSeparator = True
            else:
                f = model.ModelFormatDecimalNumber
                f.DecimalPlaces, f.UseThousandSeparator = 2, True
            return f

        for name, dax, kind, desc in MEASURES:
            model.ModelMeasures.Add(name, t, dax, fmt(kind), desc)

        # ---------- pivots ----------
        cache = wb.PivotCaches().Create(XL_EXTERNAL, wb.Connections(MODEL), 6)
        made = []

        def pivot(sheet, cell, name, rows, measures, cols=(), sort=None):
            cache = wb.PivotCaches().Create(XL_EXTERNAL, wb.Connections(MODEL), 6)
            rng = sheet.Range(cell)
            pt = cache.CreatePivotTable(f"'{sheet.Name}'!R{rng.Row}C{rng.Column}", name, False, 6)
            for f in rows:
                pt.CubeFields(f).Orientation = XL_ROW
            for f in cols:
                pt.CubeFields(f).Orientation = XL_COLUMN
            for m in measures:
                pt.AddDataField(pt.CubeFields(f'[Measures].[{m}]'), m + ' ')
            if sort:
                field = rows[0] + '.' + rows[0].split('.')[1]
                pt.PivotFields(field).AutoSort(XL_DESC, f'[Measures].[{sort}]')
            pt.RowAxisLayout(1)  # tabular
            pt.TableStyle2 = 'PivotStyleLight1'
            made.append(pt)
            return pt

        T = lambda col: f'[Transactions].[{col}]'  # noqa: E731
        p_month = pivot(pivots, 'A3', 'ptMonth', ['[Calendar].[YearMonth]'], ['Revenue'])
        p_region = pivot(pivots, 'E3', 'ptRegion', [T('region')], ['Revenue'], sort='Revenue')
        p_product = pivot(pivots, 'I3', 'ptProduct', [T('product_type')], ['Revenue'], sort='Revenue')
        p_channel = pivot(pivots, 'M3', 'ptChannel', [T('channel')], ['Revenue'], sort='Revenue')
        p_fail = pivot(pivots, 'Q3', 'ptFailure', [T('region')], ['Failure Rate'], sort='Failure Rate')
        p_status = pivot(pivots, 'U3', 'ptStatus', [T('channel')], ['Transaction Count'], cols=[T('status')])
        p_bundle = pivot(detail, 'A5', 'ptBundles', [T('product_type'), T('bundle_name')],
                         ['Successful Txns', 'Average Ticket', 'Revenue', 'Revenue Lost'])
        p_regions = pivot(detail, 'H5', 'ptRegions', [T('region')],
                          ['Transaction Count', 'Failed Txns', 'Failure Rate', 'Reversal Rate', 'Revenue Lost', 'Revenue'], sort='Revenue')
        # rows without a date have no month: keep only real months in the monthly pivot
        months = sorted({r['txn_date'][:7] for r in out if r['txn_date']})
        try:
            p_month.PivotFields('[Calendar].[YearMonth].[YearMonth]').VisibleItemsList = [f'[Calendar].[YearMonth].&[{m}]' for m in months]
        except Exception:
            pass
        pivots.Range('A1').Value = 'Pivot tables that feed the charts on the Dashboard sheet. They follow the slicers.'
        for sheet, title in ((detail, 'Detail tables'),):
            sheet.Range('A1').Value = title
            sheet.Range('A1').Font.Size, sheet.Range('A1').Font.Bold = 16, True
            sheet.Range('A2').Value = 'PivotTables on the data model. They follow the slicers on the Dashboard sheet.'
            sheet.Range('A2').Font.Color = rgb(GREY)
        detail.Range('A4').Value, detail.Range('H4').Value = 'Bundle performance', 'Regions: volume, failures and revenue'
        detail.Range('A4,H4').Font.Bold = True
        detail.Columns('A:N').ColumnWidth = 16
        xl.ActiveWindow.DisplayGridlines = True

        # ---------- dashboard ----------
        dash.Activate()
        xl.ActiveWindow.DisplayGridlines = False
        xl.ActiveWindow.Zoom = 90
        dash.Cells.Font.Name = 'Arial'
        dash.Columns('A:Z').ColumnWidth = 9
        dash.Range('A1:Z3').Interior.Color = rgb(NAVY)
        dash.Rows('1:3').RowHeight = 19
        left = 8
        if logo and Path(logo).exists():
            pic = dash.Shapes.AddPicture(str(logo), False, True, 6, 2, -1, -1)
            pic.LockAspectRatio = True
            pic.Height = 53
            left = pic.Width + 20
        title = dash.Shapes.AddTextbox(1, left, 4, 520, 26)
        title.TextFrame2.TextRange.Text = 'Telecom Transactions Report'
        sub = dash.Shapes.AddTextbox(1, left, 30, 700, 20)
        dates = sorted(r['txn_date'] for r in out if r['txn_date'])
        sub.TextFrame2.TextRange.Text = (f'Source: {log["file"]} · Period: {dates[0]} to {dates[-1]} · '
                                         f'Currency: {settings.get("currency_name", "")} ({symbol}) · Power Query + Power Pivot')
        for box, size, color, bold in ((title, 16, 'FFFFFF', True), (sub, 9, 'A9B6CC', False)):
            box.Fill.Visible, box.Line.Visible = False, False
            f = box.TextFrame2.TextRange.Font
            f.Size, f.Bold, f.Name = size, bold, 'Arial'
            f.Fill.ForeColor.RGB = rgb(color)

        # slicers (left column), connected to every pivot
        slicer_names, top = [], 70
        for field, caption, height in ((T('region'), 'Region', 178), (T('customer_segment'), 'Segment', 72),
                                       (T('product_type'), 'Product', 112), (T('channel'), 'Channel', 112),
                                       ('[Calendar].[Quarter]', 'Quarter', 112)):
            name = 'Slicer_' + caption
            sc = wb.SlicerCaches.Add2(p_region, field, name)
            sl = sc.Slicers.Add(dash, field + '.' + field.split('.')[1], caption, caption, top, 8, 132, height)
            sl.RowHeight = 15
            sl.Style = 'SlicerStyleLight6'
            for pt in made:
                if pt.Name != p_region.Name:
                    sc.PivotTables.AddPivotTable(pt)
            slicer_names.append(name)
            top += height + 6

        # KPI cards: CUBEVALUE on the model, filtered by the slicers
        cards = [('Revenue', 'Revenue', money_fmt), ('Transactions', 'Transaction Count', '#,##0'),
                 ('Success rate', 'Success Rate', '0.0%'), ('Active customers', 'Active Customers', '#,##0'),
                 ('Average ticket', 'Average Ticket', money_fmt), ('Revenue lost to failures', 'Revenue Lost', money_fmt)]
        col = 4  # column D
        for label, measure, number_format in cards:
            lab = dash.Range(dash.Cells(5, col), dash.Cells(5, col + 2))
            val = dash.Range(dash.Cells(6, col), dash.Cells(7, col + 2))
            lab.Merge()
            val.Merge()
            lab.Value = label
            lab.Font.Size, lab.Font.Color = 9, rgb(GREY)
            val.Formula = f'=CUBEVALUE("{MODEL}","[Measures].[{measure}]",{",".join(slicer_names)})'
            val.NumberFormat = number_format
            val.Font.Size, val.Font.Bold = 18, True
            val.HorizontalAlignment = XL_LEFT
            box = dash.Range(dash.Cells(5, col), dash.Cells(7, col + 2))
            box.Interior.Color = rgb('FFFFFF')
            for edge in (7, 9, 10):  # left, bottom, right
                box.Borders(edge).LineStyle = 1
                box.Borders(edge).Color = rgb(LINE)
            dash.Range(dash.Cells(5, col), dash.Cells(5, col + 2)).Borders(8).Color = rgb(TEAL)
            dash.Range(dash.Cells(5, col), dash.Cells(5, col + 2)).Borders(8).Weight = 3
            col += 3
        dash.Range('A4:Z60').Interior.Color = rgb('F4F6F8')
        for i in range(6):
            dash.Range(dash.Cells(5, 4 + i * 3), dash.Cells(7, 6 + i * 3)).Interior.Color = rgb('FFFFFF')
        dash.Rows('4:4').RowHeight = 10
        dash.Rows('8:8').RowHeight = 10

        def chart(pt, kind, title_text, x, y, w, h, colors=(TEAL,), reverse=False, legend=False, number_format=None):
            shape = dash.Shapes.AddChart2(-1, kind, x, y, w, h)
            ch = shape.Chart
            ch.SetSourceData(pt.TableRange1)
            ch.ChartType = kind
            ch.HasTitle = True
            ch.ChartTitle.Text = title_text
            ch.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 11
            ch.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = True
            ch.HasLegend = legend
            if legend:
                ch.Legend.Position = -4107  # bottom
            ch.ShowAllFieldButtons = False
            for i, color in enumerate(colors, start=1):
                if i <= ch.SeriesCollection().Count:
                    s = ch.SeriesCollection(i)
                    s.Format.Fill.ForeColor.RGB = rgb(color)
                    s.Format.Line.ForeColor.RGB = rgb(color)
            if reverse:  # first category on top, value axis at the bottom
                ch.Axes(1).ReversePlotOrder = True
                ch.Axes(1).Crosses = 2
            if number_format:
                ch.Axes(2).TickLabels.NumberFormat = number_format
            try:
                ch.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB = rgb(LINE)
            except Exception:
                pass
            shape.Line.ForeColor.RGB = rgb(LINE)
            return ch

        x0, y0, w, h, gap = 150, 136, 331, 200, 8
        chart(p_month, XL_LINE_MARKERS, f'Monthly revenue ({symbol})', x0, y0, w * 2 + gap, h, number_format='#,##0')
        chart(p_region, XL_BAR, f'Revenue by region ({symbol})', x0, y0 + h + gap, w, h, reverse=True, number_format='#,##0')
        chart(p_fail, XL_BAR, 'Failure rate by region', x0 + w + gap, y0 + h + gap, w, h, colors=(RED,), reverse=True, number_format='0%')
        chart(p_product, XL_BAR, f'Revenue by product type ({symbol})', x0, y0 + 2 * (h + gap), w, h, reverse=True, number_format='#,##0')
        chart(p_status, XL_BAR_STACKED_100, 'Transaction outcome by channel', x0 + w + gap, y0 + 2 * (h + gap), w, h,
              colors=(RED, ORANGE, GREEN), reverse=True, legend=True, number_format='0%')
        dash.Range('A1').Select()

        # ---------- cleaning ----------
        cln.Range('A1').Value = 'Cleaning applied'
        cln.Range('A1').Font.Size, cln.Range('A1').Font.Bold = 16, True
        cln.Range('A2').Value = (f'{log["rows_in"]:,} rows in the original file, {log["duplicates_removed"]} duplicates removed, '
                                 f'{log["rows_out"]:,} clean rows. Row-by-row detail in cleaning/findings_detail.csv.')
        cln.Range('A2').Font.Color = rgb(GREY)

        def table(sheet, row, heading, header, rows):
            sheet.Cells(row, 1).Value = heading
            sheet.Cells(row, 1).Font.Bold, sheet.Cells(row, 1).Font.Size = True, 12
            head = sheet.Range(sheet.Cells(row + 1, 1), sheet.Cells(row + 1, len(header)))
            head.Value = [header]
            head.Font.Bold, head.Font.Color, head.Interior.Color = True, rgb('FFFFFF'), rgb('162C4C')
            if rows:
                sheet.Range(sheet.Cells(row + 2, 1), sheet.Cells(row + 1 + len(rows), len(header))).Value = rows
            return row + len(rows) + 4

        nxt = table(cln, 4, 'What was found and what was done', ['Column', 'Issue found', 'Rows', 'Examples of the original value', 'What was done'],
                    [[g['column'], g['issue'], g['n'], "'" + ' · '.join(g['examples']), g['action']] for g in groups])
        table(cln, nxt, 'Outlier amounts corrected', ['Transaction', 'Bundle', 'Status', 'Original amount', 'Corrected amount'],
              [[a['transaction_id'], a['bundle'], a['status'], a['original_amount'], a['corrected_amount']] for a in log['outliers']])
        for letter, width in zip('ABCDE', (16, 52, 12, 60, 52)):
            cln.Columns(letter).ColumnWidth = width
        cln.Cells.WrapText = True
        cln.Cells.VerticalAlignment = XL_TOP

        # ---------- read me ----------
        notes = [
            ('How this workbook is built', None),
            ('Power Query', 'Three queries (Data > Queries & Connections): SourceFile (parameter with the path of the clean CSV), '
                            'Transactions (typed columns + list price per bundle) and Calendar (one row per day).'),
            ('Data Model', 'Both tables are loaded to the Power Pivot data model only, related by Transactions[txn_date] -> Calendar[Date]. '
                           'Open it with Power Pivot > Manage.'),
            ('Dashboard', 'KPI cards are CUBEVALUE formulas on the model; charts are PivotCharts fed by the Pivots sheet. '
                          'All of them follow the slicers.'),
            ('Refresh', 'Run the pipeline (run.bat) to rebuild this file, or use Data > Refresh All after the clean CSV changes. '
                        'If the project folder moves, edit the SourceFile parameter.'),
            ('Snapshot', 'The numbers reflect the clean CSV at the last refresh. Transactions without a date count in the totals '
                         'but have no month or quarter.'),
            ('DAX measures', None),
        ] + [(name, '= ' + dax) for name, dax, _, _ in MEASURES]
        for i, (a, b) in enumerate(notes, start=1):
            readme.Cells(i, 1).Value = a
            if b is None:
                readme.Cells(i, 1).Font.Size, readme.Cells(i, 1).Font.Bold = 14, True
            else:
                readme.Cells(i, 1).Font.Bold = True
                readme.Cells(i, 2).Value = "'" + b
        readme.Columns('A').ColumnWidth = 24
        readme.Columns('B').ColumnWidth = 120
        readme.Cells.WrapText = True
        readme.Cells.VerticalAlignment = XL_TOP

        for sheet in (readme, cln, pivots, detail, dash):  # final order: Dashboard, Detail, Pivots, Cleaning, Read me
            sheet.Move(wb.Worksheets(1))
        dash.Activate()
        xl.ScreenUpdating = True
        # cube formulas resolve asynchronously: recalculate and wait so the saved file opens with values, not errors
        for _ in range(40):
            xl.CalculateFull()
            if not any(str(dash.Cells(6, 4 + 3 * i).Text).startswith('#') for i in range(len(cards))):
                break
            time.sleep(0.5)
        # save to a local temp file first: Excel's SaveAs can fail on synced (OneDrive) folders
        tmp = Path(tempfile.gettempdir()) / f'telecom_report_{os.getpid()}.xlsx'
        if tmp.exists():
            tmp.unlink()
        wb.SaveAs(str(tmp), XL_XLSX)
        wb.Close(False)
        wb = None
        shutil.move(str(tmp), str(target))  # raises PermissionError if the target is open in Excel
    finally:
        try:
            if wb is not None:
                wb.Close(False)
            xl.Quit()
        except Exception:
            pass
        del xl
        pythoncom.CoUninitialize()
        try:  # Quit normally ends the process; if it is still there, end it
            os.kill(pid, signal.SIGTERM)
        except OSError:
            pass
