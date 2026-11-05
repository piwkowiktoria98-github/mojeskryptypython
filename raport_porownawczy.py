import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

OUT_FILE = r"C:\Users\wpiwko\Desktop\KP2.00 x DAX.xlsx"

def parse_num(v):
    if v is None or str(v).strip() in ('', 'nan', 'None'):
        return 0.0
    return float(str(v).replace(' ', '').replace(',', '.'))

# === WCZYTAJ DANE Z ISTNIEJACEGO PLIKU ===
wb_src = openpyxl.load_workbook(OUT_FILE, data_only=True)

# KP 2.00 balance z Arkusza 1
ws_kp = wb_src['KP 2.00 balance']
kp_headers = [ws_kp.cell(1, j).value for j in range(1, ws_kp.max_column + 1)]
kp_df = pd.DataFrame([
    {kp_headers[j]: ws_kp.cell(i, j + 1).value for j in range(len(kp_headers))}
    for i in range(2, ws_kp.max_row + 1)
    if ws_kp.cell(i, 1).value is not None
])
kp_data = {}
for _, row in kp_df.iterrows():
    acct = str(row['LEDGER_ACCOUNT_CODE']).strip()
    kp_data[acct] = {
        'name': str(row.get('LEDGER_ACCOUNT_NAME', '')).strip(),
        'currency': str(row.get('TRANSACTION_CURRENCY', '')).strip(),
        'debit_base': parse_num(row.get('DEBIT_IN_COMPANY_CURRENCY_TURNOVER', 0)),    # M
        'credit_base': parse_num(row.get('CREDIT_IN_COMPANY_CURRENCY_TURNOVER', 0)),  # N
        'balance_base': parse_num(row.get('SALDO_IN_COMPANY_CURRENCY_TURNOVER', 0)),  # O
        'debit': parse_num(row.get('DEBIT_TURNOVER', 0)),                       # P
        'credit': parse_num(row.get('CREDIT_TURNOVER', 0)),                     # Q
        'balance': parse_num(row.get('SALDO_TURNOVER', 0)),                     # R
    }

# DAX balance z Arkusza 2
ws_dax = wb_src['DAX balance']
dax_data = {}
for i in range(2, ws_dax.max_row + 1):
    acct = str(ws_dax.cell(i, 2).value or '').strip()
    if not acct:
        continue
    dax_data[acct] = {
        'debit_base': ws_dax.cell(i, 6).value or 0,    # F - Turnover db currency DT
        'credit_base': ws_dax.cell(i, 7).value or 0,   # G - Turnover db currency CT
        'balance_base': ws_dax.cell(i, 8).value or 0,   # H - Turnover db currency
        'debit': ws_dax.cell(i, 12).value or 0,          # L - Turnover trans currency DT
        'credit': ws_dax.cell(i, 13).value or 0,         # M - Turnover trans currency CT
        'balance': ws_dax.cell(i, 14).value or 0,        # N - Turnover trans currency
    }

all_accounts = sorted(set(list(kp_data.keys()) + list(dax_data.keys())))
print(f"Konta KP: {len(kp_data)}, DAX: {len(dax_data)}, Razem: {len(all_accounts)}")

# === BUDUJ RAPORT ===
wb = openpyxl.load_workbook(OUT_FILE)

# Usun stary arkusz Porownanie jesli istnieje, stworz nowy
if 'Porownanie' in wb.sheetnames:
    del wb['Porownanie']
ws3 = wb.create_sheet('Porownanie')

header_font = Font(bold=True, color='FFFFFF', size=10)
header_fill = PatternFill('solid', fgColor='2F5496')
acct_font = Font(bold=True, size=11)
acct_fill = PatternFill('solid', fgColor='D6E4F0')
ok_fill = PatternFill('solid', fgColor='C6EFCE')
bad_fill = PatternFill('solid', fgColor='FFC7CE')
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin'))
num_fmt = '#,##0.00'

row_idx = 1
ws3.cell(row_idx, 1, 'RAPORT POROWNAWCZY SALD: KP 2.00 vs DAX').font = Font(bold=True, size=14)
ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
row_idx += 1
ws3.cell(row_idx, 1, 'Data: 31.08.2026').font = Font(italic=True, size=10, color='666666')
row_idx += 2

for acct in all_accounts:
    kp = kp_data.get(acct, {})
    dax = dax_data.get(acct, {})
    name = kp.get('name', '')
    currency = kp.get('currency', '')

    # Naglowek konta
    acct_label = f"{acct} - {name}" if name else acct
    if currency:
        acct_label += f" ({currency})"
    ws3.cell(row_idx, 1, acct_label).font = acct_font
    for c in range(1, 9):
        ws3.cell(row_idx, c).fill = acct_fill
        ws3.cell(row_idx, c).border = thin_border
    ws3.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=8)
    row_idx += 1

    # Naglowki kolumn
    headers = ['', 'Debit (wal. trans.)', 'Credit (wal. trans.)', 'Saldo (wal. trans.)',
               '', 'Debit (wal. bazy PLN)', 'Credit (wal. bazy PLN)', 'Saldo (wal. bazy PLN)']
    for c, h in enumerate(headers, 1):
        cell = ws3.cell(row_idx, c, h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
    row_idx += 1

    # KP 2.00
    ws3.cell(row_idx, 1, 'KP 2.00').font = Font(bold=True)
    ws3.cell(row_idx, 1).border = thin_border
    kp_vals = [
        kp.get('debit', 0), kp.get('credit', 0), kp.get('balance', 0),
        None,
        kp.get('debit_base', 0), kp.get('credit_base', 0), kp.get('balance_base', 0)
    ]
    for c, v in enumerate(kp_vals, 2):
        cell = ws3.cell(row_idx, c, v)
        cell.border = thin_border
        if isinstance(v, (int, float)):
            cell.number_format = num_fmt
    row_idx += 1

    # DAX
    ws3.cell(row_idx, 1, 'DAX').font = Font(bold=True)
    ws3.cell(row_idx, 1).border = thin_border
    dax_vals = [
        float(dax.get('debit', 0)), float(dax.get('credit', 0)), float(dax.get('balance', 0)),
        None,
        float(dax.get('debit_base', 0)), float(dax.get('credit_base', 0)), float(dax.get('balance_base', 0))
    ]
    for c, v in enumerate(dax_vals, 2):
        cell = ws3.cell(row_idx, c, v)
        cell.border = thin_border
        if isinstance(v, (int, float)):
            cell.number_format = num_fmt
    row_idx += 1

    # Roznica
    ws3.cell(row_idx, 1, 'ROZNICA').font = Font(bold=True, color='C00000')
    ws3.cell(row_idx, 1).border = thin_border
    for c in range(len(kp_vals)):
        if kp_vals[c] is None or dax_vals[c] is None:
            ws3.cell(row_idx, c + 2).border = thin_border
            continue
        diff = round(float(kp_vals[c] or 0) - float(dax_vals[c] or 0), 2)
        cell = ws3.cell(row_idx, c + 2, diff)
        cell.number_format = num_fmt
        cell.border = thin_border
        cell.font = Font(bold=True)
        if abs(diff) < 0.01:
            cell.fill = ok_fill
        else:
            cell.fill = bad_fill
    row_idx += 2

# Szerokosci kolumn
ws3.column_dimensions['A'].width = 18
for col in ['B', 'C', 'D', 'F', 'G', 'H']:
    ws3.column_dimensions[col].width = 24
ws3.column_dimensions['E'].width = 3

wb.save(OUT_FILE)

# Podsumowanie
ok_count = 0
bad_count = 0
for acct in all_accounts:
    kp = kp_data.get(acct, {})
    dax = dax_data.get(acct, {})
    bal_kp = float(kp.get('balance', 0))
    bal_dax = float(dax.get('balance', 0))
    bal_base_kp = float(kp.get('balance_base', 0))
    bal_base_dax = float(dax.get('balance_base', 0))
    if abs(bal_kp - bal_dax) < 0.01 and abs(bal_base_kp - bal_base_dax) < 0.01:
        ok_count += 1
    else:
        bad_count += 1
        name = kp.get('name', '')
        print(f"  ROZNICA: {acct} {name}")
        print(f"    Saldo trans: KP={bal_kp:.2f} DAX={bal_dax:.2f} diff={bal_kp-bal_dax:.2f}")
        print(f"    Saldo bazy:  KP={bal_base_kp:.2f} DAX={bal_base_dax:.2f} diff={bal_base_kp-bal_base_dax:.2f}")

print(f"\nZapisano: {OUT_FILE}")
print(f"  Arkusz 1: KP 2.00 balance ({len(kp_df)} wierszy)")
print(f"  Arkusz 2: DAX balance ({ws_dax.max_row-1} wierszy)")
print(f"  Arkusz 3: Porownanie ({len(all_accounts)} kont)")
print(f"\n  Zgodne: {ok_count}  |  Roznice: {bad_count}")
