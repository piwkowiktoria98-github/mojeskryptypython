import pandas as pd
import re
import os
import glob
from html import escape as html_escape

FOLDER = r'C:\Users\wpiwko\Desktop\UK formatowanie'

# --- Auto-wykrywanie plików ---
statement_files = [f for f in glob.glob(os.path.join(FOLDER, 'Save_As_Account_Statement*.xls'))
                   if '_sformatowany' not in f]
if not statement_files:
    print('BŁĄD: Nie znaleziono plików wyciągów w folderze.')
    exit(1)

batch_files = glob.glob(os.path.join(FOLDER, 'Macro batch upload *.xlsm'))
batch_map = {}
for bf in batch_files:
    match = re.search(r'(\d{2})(\d{2})(\d{4})\.xlsm$', bf)
    if match:
        day, month, year = match.groups()
        date_str = f'{day}/{month}/{year}'
        batch_map[date_str] = bf

print(f'Znalezione wyciągi: {len(statement_files)}')
print(f'Znalezione batche: {len(batch_map)}')
for date, path in sorted(batch_map.items()):
    print(f'  {date} -> {os.path.basename(path)}')
print()

# --- Funkcje ---

def format_pl(val):
    num = float(val)
    sign = '-' if num < 0 else ''
    num = abs(num)
    integer_part = int(num)
    decimal_part = round(num - integer_part, 2)
    decimal_str = f'{decimal_part:.2f}'[2:]
    int_str = f'{integer_part:,}'.replace(',', ' ')
    return f'{sign}{int_str},{decimal_str}'

def parse_amount(val):
    if pd.isna(val):
        return 0.0
    return float(str(val).replace(' ', '').replace(',', ''))

def format_amount_en(num):
    return f'{num:,.2f}'

def detect_account_type(filepath):
    """Rozpoznaje typ konta z nagłówka: zwraca (waluta, typ) np. ('GBP', 'retail') lub ('GBP', 'isa')."""
    dfs = pd.read_html(filepath)
    header = dfs[0]
    row0_vals = [str(v) for v in header.iloc[0] if pd.notna(v)]
    row0_text = ' '.join(row0_vals).upper()

    currency = 'GBP'
    for cur in ['EUR', 'USD', 'GBP']:
        if cur in row0_vals:
            currency = cur
            break

    acc_type = 'retail'
    if 'ISA' in row0_text:
        acc_type = 'isa'

    return currency, acc_type

SHEET_NAME_XML = """<xml>
<x:ExcelWorkbook>
<x:ExcelWorksheets>
<x:ExcelWorksheet>
<x:Name>Sheet1</x:Name>
<x:WorksheetOptions>
<x:Print><x:ValidPrinterInfo/></x:Print>
</x:WorksheetOptions>
</x:ExcelWorksheet>
</x:ExcelWorksheets>
</x:ExcelWorkbook>
</xml>"""

def set_sheet_name(html):
    """Wstawia nazwę arkusza Sheet1 do bloku mso."""
    return html.replace(
        '<!--[if gte mso 9]>\n\n<![endif]-->',
        '<!--[if gte mso 9]>\n' + SHEET_NAME_XML + '\n<![endif]-->'
    )

def format_dates_only(filepath, output_path):
    """Tylko formatuje kwoty EN→PL w pliku HTML. Bez rozbijania batchy."""
    with open(filepath, 'r', encoding='utf-8') as f:
        html = f.read()

    # Kwoty w komórkach
    html = re.sub(
        r'>([\d,]+\.\d{2})</td>',
        lambda m: '>' + format_pl(float(m.group(1).replace(',', ''))) + '</td>',
        html
    )
    # Kwoty w nagłówku z &nbsp;As At:
    html = re.sub(
        r'>([\d,]+\.\d{2})(&nbsp;)',
        lambda m: '>' + format_pl(float(m.group(1).replace(',', ''))) + m.group(2),
        html
    )
    # Kwoty w nagłówku Total Payment/Receipt
    html = re.sub(
        r'>([\d,]+\.\d{2})(/\d+</td>)',
        lambda m: '>' + format_pl(float(m.group(1).replace(',', ''))) + m.group(2),
        html
    )

    html = set_sheet_name(html)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

def process_gbp_retail(filepath, output_path):
    """Pełne przetwarzanie GBP Retail: rozbijanie batchy + formatowanie kwot."""
    with open(filepath, 'r', encoding='utf-8') as f:
        html = f.read()

    dfs = pd.read_html(filepath)
    header = dfs[0]
    df = dfs[1]

    # Oczekiwane salda
    raw_total_payment = header.iloc[6, 1]
    raw_total_receipt = header.iloc[6, 2]
    expected_payment = float(re.match(r'([\d,]+\.\d+)', raw_total_payment).group(1).replace(',', ''))
    expected_receipt = float(re.match(r'([\d,]+\.\d+)', raw_total_receipt).group(1).replace(',', ''))

    ROW_EVEN = '''	</tr><tr style="font-size:8pt;height:35px;font-family:Arial;text-align:left;">
		<td style="mso-number-format:\\@;border:1px solid #808080;">{date}</td><td style="mso-number-format:\\@;border:1px solid #808080;">

                                                                    {details}
                                                                </td><td style="mso-number-format:\\@;border:1px solid #808080;">Bill Payment via FP</td><td class="cmb_blackcolor" align="right" style="mso-number-format:\\#\\,\\#\\#0.00;border:1px solid #808080;">{amount}</td><td class="cmb_blackcolor" style="mso-number-format:;border:1px solid #808080;"></td><td class="cmb_blackcolor" align="right" style="mso-number-format:;border:1px solid #808080;"></td>'''

    ROW_ODD = '''	</tr><tr style="font-size:8pt;height:35px;font-family:Arial;text-align:left;background-color:#F1F1F1;">
		<td style="mso-number-format:\\@;border:1px solid #808080;">{date}</td><td style="mso-number-format:\\@;border:1px solid #808080;">

                                                                    {details}
                                                                </td><td style="mso-number-format:\\@;border:1px solid #808080;">Bill Payment via FP</td><td class="cmb_blackcolor" align="right" style="mso-number-format:\\#\\,\\#\\#0.00;border:1px solid #808080;">{amount}</td><td class="cmb_blackcolor" style="mso-number-format:;border:1px solid #808080;"></td><td class="cmb_blackcolor" align="right" style="mso-number-format:;border:1px solid #808080;"></td>'''

    # Rozbij batche
    for date, batch_path in sorted(batch_map.items()):
        batch_df = pd.read_excel(batch_path, sheet_name='Import file', engine='openpyxl', header=None)
        # Kolumny: 0=sort code, 1=name, 2=bank account, 3=amount, 4=details (Outgoing XXXXX), 5=code
        if batch_df.shape[1] < 5:
            print(f'  UWAGA: Za malo kolumn w {os.path.basename(batch_path)} - pomijam')
            continue

        batch_mask = (df['1Entry Date'] == date) & (df['Transaction Details'].str.contains(r'Outgoing\s+FT$', case=False, na=False, regex=True))
        batch_indices = df.index[batch_mask].tolist()

        if len(batch_indices) != 1:
            print(f'  UWAGA: {len(batch_indices)} batchy dla {date} - pomijam')
            continue

        batch_idx = batch_indices[0]
        batch_amount = df.loc[batch_idx, 'Payment Amount']
        batch_sum = batch_df.iloc[:, 3].apply(parse_amount).sum()

        if abs(batch_sum - batch_amount) >= 0.01:
            print(f'  UWAGA: Suma batch {date} ({batch_sum:.2f}) != wyciąg ({batch_amount:.2f}) - pomijam')
            continue

        batch_amount_str = format_amount_en(batch_amount)

        pattern = (
            r'(</tr>)?<tr[^>]*>\s*'
            r'<td[^>]*>' + re.escape(date) + r'</td>'
            r'<td[^>]*>\s*\n?\s*\n?\s*[^<]*Outgoing\s+FT\s*\n?\s*</td>'
            r'<td[^>]*>[^<]*</td>'
            r'<td[^>]*>' + re.escape(batch_amount_str) + r'</td>'
            r'<td[^>]*>[^<]*</td>'
            r'<td[^>]*>[^<]*</td>'
        )

        match = re.search(pattern, html, re.DOTALL | re.IGNORECASE)
        if not match:
            print(f'  UWAGA: Nie znaleziono batch w HTML dla {date} ({batch_amount_str}) - pomijam')
            continue

        new_rows_html = []
        for i, (_, row) in enumerate(batch_df.iterrows()):
            amount_val = parse_amount(row.iloc[3])
            amount_str = format_amount_en(amount_val)
            details = html_escape(str(row.iloc[4]))
            template = ROW_ODD if i % 2 == 1 else ROW_EVEN
            new_rows_html.append(template.format(date=date, details=details, amount=amount_str))

        replacement = ''.join(new_rows_html)
        html = html[:match.start()] + replacement + html[match.end():]

        new_rows_data = []
        for _, row in batch_df.iterrows():
            new_rows_data.append({
                '1Entry Date': date,
                'Transaction Details': str(row.iloc[4]),
                '2Transaction Type': 'Bill Payment via FP',
                'Payment Amount': parse_amount(row.iloc[3]),
                'Receipt Amount': float('nan'),
                'Ledger Balance': float('nan'),
            })
        new_df = pd.DataFrame(new_rows_data)
        df_before = df.loc[:batch_idx - 1] if batch_idx > 0 else pd.DataFrame(columns=df.columns)
        df_after = df.loc[batch_idx + 1:]
        df = pd.concat([df_before, new_df, df_after], ignore_index=True)

        print(f'  {date}: Zamieniono 1 batch ({format_pl(batch_amount)}) na {len(batch_df)} transakcji')

    # Weryfikacja salda
    payments = df['Payment Amount'].dropna()
    receipts = df['Receipt Amount'].dropna()
    sum_payments = payments.sum()
    sum_receipts = receipts.sum()

    print(f'  Wierszy po rozbiciu: {len(df)}')
    print(f'  Payment: {sum_payments:.2f} (oczekiwane {expected_payment:.2f}) {"OK" if abs(sum_payments - expected_payment) < 0.01 else "NIEZGODNE!"}')
    print(f'  Receipt: {sum_receipts:.2f} (oczekiwane {expected_receipt:.2f}) {"OK" if abs(sum_receipts - expected_receipt) < 0.01 else "NIEZGODNE!"}')

    # Formatowanie kwot EN→PL
    html = re.sub(
        r'>([\d,]+\.\d{2})</td>',
        lambda m: '>' + format_pl(float(m.group(1).replace(',', ''))) + '</td>',
        html
    )
    html = re.sub(
        r'>([\d,]+\.\d{2})(&nbsp;)',
        lambda m: '>' + format_pl(float(m.group(1).replace(',', ''))) + m.group(2),
        html
    )
    html = re.sub(
        r'>([\d,]+\.\d{2})(/\d+</td>)',
        lambda m: '>' + format_pl(float(m.group(1).replace(',', ''))) + m.group(2),
        html
    )

    html = set_sheet_name(html)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)


# --- MAIN ---
for filepath in sorted(statement_files):
    basename = os.path.basename(filepath)
    currency, acc_type = detect_account_type(filepath)
    label = f'{currency} {"ISA" if acc_type == "isa" else "Retail"}'

    dfs = pd.read_html(filepath)
    account_num = str(dfs[0].iloc[0, 0]).strip()
    output_path = os.path.join(FOLDER, f'{account_num}.xls')

    print(f'{"="*60}')
    print(f'{basename} -> {label}')
    print(f'{"="*60}')

    if currency == 'GBP' and acc_type == 'retail':
        process_gbp_retail(filepath, output_path)
    else:
        format_dates_only(filepath, output_path)
        print(f'  Tylko formatowanie kwot (bez batchy)')

    print(f'  Zapisano: {account_num}.xls')
    print()

print('Gotowe! Wszystkie pliki przetworzone.')
