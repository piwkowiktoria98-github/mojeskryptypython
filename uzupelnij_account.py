"""
Uzupełnia kolumnę AccountNum w DAX_uzupelnij.csv na podstawie rapot_DWH.xlsx.

Logika:
1. Szukamy dokładnego dopasowania kwoty (CreditAmount == Amount).
   - Dokładnie 1 rekord z tą kwotą → jednoznaczne dopasowanie.
   - Więcej rekordów z tą kwotą → zawężamy do nich i weryfikujemy po imieniu/nazwisku. 
2. Weryfikacja po imieniu/nazwisku: wyciągamy pierwsze 3 słowa z Description (bez Pan/Pani po rumuńsku),
   myślnik traktujemy jak spację, szukamy co najmniej 2 z 3 słów w kolumnie Name.
3. Wynik zapisujemy do DAX_uzupelnij_WYNIK.csv (oryginał zostaje nienaruszony).
"""

import re
import glob
import os
import sys
import unicodedata
import pandas as pd
import openpyxl
import win32com.client

import zaznacz_bledy_dax


# konfiguracja

FOLDER = r"C:\Users\wpiwko\Desktop\Uzupelnij loginy"
ZWROT_ACCOUNT = "240/04"
RAPORT_FILE = "Raport porównawczy DAX !!!!.xlsx"
RAPORT_SHEET = "DepositWithdrawal_DAX"
DWH_DATA_START_ROW = 8


def find_file(pattern: str, label: str) -> str:
    matches = sorted(glob.glob(os.path.join(FOLDER, pattern)))
    if not matches:
        print(f"BŁĄD: Nie znaleziono pliku {label} (wzorzec: {pattern})")
        sys.exit(1)
    if len(matches) > 1:
        print(f"UWAGA: Znaleziono {len(matches)} plików {label}:")
        for m in matches:
            print(f"  - {os.path.basename(m)}")
        print(f"  Używam: {os.path.basename(matches[0])}")
    return matches[0]

# Rząd nagłówkowy w Excelu (0-based); wiersz 4 zawiera Branch Name / Login / Name / Amount…
# Pierwsze trzy rzędy są puste
DWH_HEADER_ROW = 4

# Minimalna liczba słów z Description, które muszą trafić w Name
MIN_WORD_MATCHES = 2

# Rumuńskie/bankowe stopwords — NIE są częścią imienia/nazwiska
STOPWORDS = {
    'PLATA', 'RECURENTA', 'TRANSFER', 'TRIMIS', 'PRIN', 'REVOLUT',
    'BANK', 'REFERENCE', 'INVESTITII', 'ALIMENTARE', 'LUNA', 'AUGUST',
    'PT', 'RATA', 'FOND', 'CONT', 'CATRE', 'ALTA', 'BANCA', 'INTERNAL',
    'VIRAMENT', 'DEPUNERE', 'XTB', 'IULIE', 'SEPTEMBRIE', 'OCTOMBRIE',
    'NOIEMBRIE', 'DECEMBRIE', 'IANUARIE', 'FEBRUARIE', 'MARTIE', 'APRILIE',
    'MAI', 'IUNIE', 'TRANZACTIE', 'ORDIN', 'PLATI', 'INCASARE',
}

# ──────────────────────────────────────────────────────────────────────────────


def normalize(text: str) -> str:
    """Zamienia na wielkie litery i usuwa znaki diakrytyczne (ą→A, ș→S itp.)."""
    text = text.upper()
    # Normalizacja NFD rozbija litery złożone na bazową + akcentową
    nfd = unicodedata.normalize("NFD", text)
    return "".join(c for c in nfd if unicodedata.category(c) != "Mn")


def extract_name_words(description: str) -> list[str]:
    """
    Zwraca pierwsze 3 słowa z Description po:
    - usunięciu prefixu Pan/Pani/Dl/Dnl (rumuński odpowiednik)
    - normalizacji (wielkie litery, bez diakrytyków)
    """
    if not isinstance(description, str) or not description.strip():
        return []

    text = description.strip()

    # Utnij przy "Bank reference" / "Internal reference" — reszta to UUID, nie imię
    for marker in ['Bank reference', 'Internal reference', 'bank reference', 'internal reference']:
        idx = text.lower().find(marker.lower())
        if idx > 0:
            text = text[:idx]
            break

    # Usuń typowe prefixy honoryfikatywne (case-insensitive)
    prefixes = r"^(pan|pani|dl\.?|dnl\.?|dna\.?|dnei\.?|doamna|domnul|mr\.?|mrs\.?|ms\.?|dr\.?|dra\.?)\s+"
    text = re.sub(prefixes, "", text, flags=re.IGNORECASE).strip()

    # Myślnik traktujemy jak spację (np. "Jan-Kowalski" → "Jan Kowalski")
    text = text.replace("-", " ")

    words = text.split()
    # Normalizuj i filtruj stopwords — bierzemy do 4 słów imienia/nazwiska
    name_words = []
    for w in words:
        if not re.search(r"[a-zA-ZÀ-ÿ]", w):
            continue
        if re.search(r"\d", w):
            continue
        clean = re.sub(r"[^a-zA-ZÀ-ÿ]", "", w)
        if len(clean) < 2:
            continue
        nw = normalize(clean)
        if nw in STOPWORDS:
            continue
        name_words.append(nw)
        if len(name_words) >= 4:
            break
    return name_words


def parse_amount(value) -> float | None:
    """Konwertuje kwotę z różnych formatów (100,89 / 100.89 / NaN) na float."""
    if pd.isna(value):
        return None
    text = str(value).strip().replace(" ", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def load_zwroty(path: str, dax_date: str) -> tuple[pd.DataFrame, float]:
    """
    Wczytuje zwroty z pliku Excel dla danej daty DAX.
    Szuka we wszystkich arkuszach wierszy z datą pasującą do dax_date.
    dax_date w formacie "09.07.2026" → szukamy daty 2026-07-09 w kolumnie Data.
    Zwraca (DataFrame z kolumnami Beneficiar/Suma, oczekiwana suma z nagłówka).
    """
    if "-" in dax_date and len(dax_date.split("-")) == 3:
        iso_date = dax_date.strip()
    else:
        parts = dax_date.split(".")
        if len(parts) == 3:
            iso_date = f"{parts[2]}-{parts[1]}-{parts[0]}"
        else:
            return pd.DataFrame(columns=["Beneficiar", "Suma"]), 0.0

    xls = pd.ExcelFile(path, engine="openpyxl")
    all_zwroty = []

    for sheet in xls.sheet_names:
        raw = pd.read_excel(path, sheet_name=sheet, engine="openpyxl", header=None, dtype=str)
        if raw.shape[0] < 2 or raw.shape[1] < 6:
            continue

        for i in range(1, raw.shape[0]):
            date_val = ""
            for dc in [6, 5, 7]:
                if dc < raw.shape[1] and pd.notna(raw.iloc[i, dc]):
                    val = str(raw.iloc[i, dc]).strip()
                    if iso_date in val:
                        date_val = val
                        break
            if iso_date in date_val:
                name = str(raw.iloc[i, 4]).strip() if pd.notna(raw.iloc[i, 4]) else ""
                amount = raw.iloc[i, 2]
                if name:
                    all_zwroty.append({"Beneficiar": name, "Suma": amount, "Arkusz": sheet})

    if not all_zwroty:
        return pd.DataFrame(columns=["Beneficiar", "Suma"]), 0.0

    zwroty_df = pd.DataFrame(all_zwroty)
    total = zwroty_df["Suma"].apply(parse_amount).sum()
    return zwroty_df, total


def match_zwrot(desc_words: list[str], amount_dax: float, zwroty: pd.DataFrame, used: set) -> str | None:
    """
    Sprawdza czy transakcja pasuje do zwrotu.
    Każdy zwrot (wiersz) może być użyty tylko raz — indeksy zużytych są w `used`.
    Zwraca nazwę arkusza (datę zwrotu) lub None.
    """
    for zidx, zrow in zwroty.iterrows():
        if zidx in used:
            continue

        name_zwrot = normalize(str(zrow["Beneficiar"])).replace("-", " ")
        name_zwrot = re.sub(r"^(DL|DRA|DNA|DL\.|DRA\.|DNA\.)\s+", "", name_zwrot).strip()
        zwrot_words = [w for w in name_zwrot.split() if re.search(r"[A-Z]", w)]
        amount_zwrot = parse_amount(zrow["Suma"])

        if amount_zwrot is None or abs(amount_dax - amount_zwrot) >= 0.01:
            continue

        matches = sum(1 for w in desc_words if w in zwrot_words)
        if matches >= MIN_WORD_MATCHES:
            used.add(zidx)
            return str(zrow.get("Arkusz", ""))

    return None


def load_dax(path: str) -> pd.DataFrame:
    try:
        df = pd.read_csv(path, encoding="utf-8-sig", sep=";", dtype=str)
    except UnicodeDecodeError:
        df = pd.read_csv(path, encoding="windows-1250", sep=";", dtype=str)
    first_col = df.columns[0]
    clean_name = first_col
    for prefix in ['ď»ż', '﻿', '﻿']:
        clean_name = clean_name.replace(prefix, '')
    idx = clean_name.find('DocumentDate')
    if idx > 0:
        clean_name = clean_name[idx:]
    df.columns = [clean_name] + list(df.columns[1:])
    df.columns = [c.strip() for c in df.columns]
    return df


def load_dwh(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, header=DWH_HEADER_ROW, dtype=str)
    df.columns = df.columns.str.strip()
    # Usuń puste wiersze (bez wartości w kolumnie Name)
    df = df.dropna(subset=["Name"])
    return df


def _match_by_name(desc_words: list[str], pool: pd.DataFrame) -> tuple[str | None, str]:
    """Szuka w pool wierszy, gdzie co najmniej MIN_WORD_MATCHES słów trafia w Name_norm."""
    candidates = []
    for _, row in pool.iterrows():
        name_words = str(row.get("Name_norm", "")).split()
        matches = sum(1 for w in desc_words if w in name_words)
        if matches >= MIN_WORD_MATCHES:
            candidates.append((matches, row["Login"], row["Name"]))

    if not candidates:
        return None, f"brak dopasowania po imieniu dla słów: {desc_words}"

    candidates.sort(key=lambda c: c[0], reverse=True)

    if len(candidates) == 1:
        return candidates[0][1], f"dopasowano po imieniu ({candidates[0][0]} słów): {candidates[0][2]}"

    # Jeśli najlepszy kandydat ma wyraźnie więcej trafień niż drugi — wybierz go
    if candidates[0][0] > candidates[1][0]:
        return candidates[0][1], f"dopasowano po imieniu (najlepszy {candidates[0][0]} vs {candidates[1][0]} słów): {candidates[0][2]}"

    # Równa liczba trafień — sprawdź czy ten sam login
    top_score = candidates[0][0]
    top_candidates = [c for c in candidates if c[0] == top_score]
    logins = [c[1] for c in top_candidates]
    names  = [c[2] for c in top_candidates]
    if len(set(logins)) == 1:
        return top_candidates[0][1], f"dopasowano po imieniu ({top_candidates[0][0]} słów, powtórzony login): {names}"
    return None, f"NIEJEDNOZNACZNE po imieniu ({len(top_candidates)} kandydatów z {top_score} słowami): {names}"


def find_login(desc_words: list[str], amount_dax: float | None,
               dwh: pd.DataFrame) -> tuple[str | None, str]:
    """
    Szuka pasującego wiersza w DWH.
    Zwraca (login, status_info).
    """
    if not desc_words:
        return None, "brak słów w Description"

    # Krok 1: dokładne dopasowanie po kwocie (wymagane, jeśli kwota jest znana)
    if amount_dax is not None:
        exact = dwh[dwh["Amount_parsed"] == amount_dax]

        if len(exact) == 0:
            return None, f"brak rekordu z kwotą {amount_dax}"

        if len(exact) == 1:
            row = exact.iloc[0]
            return row["Login"], f"dopasowano po kwocie: {row['Name']}"

        # Wiele rekordów z tą kwotą → weryfikuj po imieniu tylko wśród nich
        result = _match_by_name(desc_words, exact)
        if result[0] is not None:
            return result[0], f"[kwota={amount_dax}, {len(exact)} kandydatów] " + result[1]
        return None, f"NIEJEDNOZNACZNE po kwocie ({len(exact)} kandydatów), " + result[1]

    # Brak kwoty w DAX → dopasowanie wyłącznie po imieniu/nazwisku
    return _match_by_name(desc_words, dwh)


def raport_otwarty(raport_path: str) -> bool:
    lock = os.path.join(os.path.dirname(raport_path), "~$" + os.path.basename(raport_path))
    return os.path.exists(lock)


def otworz_excel():
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    return xl


def append_dwh_to_raport(dwh_file: str):
    """
    Dopisuje wiersze z DWH (od wiersza DWH_DATA_START_ROW, kolumny B-P) na koniec arkusza
    DepositWithdrawal_DAX w Raporcie porównawczym, z formułą =H&L w kolumnie A.
    Zapis przez Excela - openpyxl przy tak dużym pliku psuł tabele, kasował formatowanie
    warunkowe i raz zapisał plik w połowie.
    """
    raport_path = os.path.join(FOLDER, RAPORT_FILE)
    if not os.path.exists(raport_path):
        print(f"BŁĄD: Nie znaleziono pliku raportu: {RAPORT_FILE}")
        return

    print("=== Kopiowanie DWH do Raportu porównawczego ===")

    marker_path = os.path.join(FOLDER, ".last_dwh.txt")
    dwh_name = os.path.basename(dwh_file)

    if os.path.exists(marker_path):
        with open(marker_path, "r", encoding="utf-8") as f:
            last_dwh = f.read().strip()
        if last_dwh == dwh_name:
            print(f"  Plik {dwh_name} już był skopiowany. Pomijam.")
            print()
            return

    if raport_otwarty(raport_path):
        print(f"BŁĄD: {RAPORT_FILE} jest otwarty w Excelu - zamknij go i uruchom ponownie.")
        sys.exit(1)

    # Dane DWH czytamy openpyxl (tylko odczyt - bezpieczne), bez scaleń ani formatowania
    dwh_wb = openpyxl.load_workbook(dwh_file, read_only=True, data_only=True)
    dwh_ws = dwh_wb.active
    wiersze = [list(row) for row in dwh_ws.iter_rows(min_row=DWH_DATA_START_ROW, min_col=2, max_col=16,
                                                     values_only=True) if row[0] is not None]
    dwh_wb.close()
    if not wiersze:
        print(f"  Brak danych w {dwh_name} od wiersza {DWH_DATA_START_ROW}. Pomijam.")
        print()
        return

    xl = otworz_excel()
    try:
        wb = xl.Workbooks.Open(raport_path)
        ws = wb.Worksheets(RAPORT_SHEET)
        last_row = ws.Cells(ws.Rows.Count, 8).End(-4162).Row  # ostatni wiersz w kolumnie H (Login)
        print(f"  Raport: {last_row} wierszy (przed dopisaniem)")

        first, last = last_row + 1, last_row + len(wiersze)
        ws.Range(ws.Cells(first, 2), ws.Cells(last, 16)).Value = wiersze
        ws.Range(f"A{first}:A{last}").Formula = [(f"=H{r}&L{r}",) for r in range(first, last + 1)]
        print(f"  Dopisano {len(wiersze)} wierszy (nowy zakres: {first} - {last})")

        wb.Save()
        wb.Close()
    finally:
        xl.Quit()

    with open(marker_path, "w", encoding="utf-8") as f:
        f.write(dwh_name)

    print(f"  Zapisano: {RAPORT_FILE}")
    print()


def main():
    dax_file = find_file("Export_*.csv", "DAX (Export CSV)")
    dwh_file = find_file("DepositWithdrawal_DAX*.xlsx", "DWH (DepositWithdrawal)")
    zwroty_file = find_file("RO zwroty*.xlsx", "Zwroty")

    print(f"Pliki:")
    print(f"  DAX:    {os.path.basename(dax_file)}")
    print(f"  DWH:    {os.path.basename(dwh_file)}")
    print(f"  Zwroty: {os.path.basename(zwroty_file)}")
    print()

    append_dwh_to_raport(dwh_file)

    print("Wczytuję pliki…")
    dax = load_dax(dax_file)
    dwh = load_dwh(dwh_file)

    print(f"  DAX:  {len(dax)} wierszy")
    print(f"  DWH:  {len(dwh)} wierszy")

    # Usuń wiersze z pustym Description
    empty_desc = dax["Description"].isna() | (dax["Description"].str.strip() == "")
    if empty_desc.any():
        print(f"  Usunięto {empty_desc.sum()} wierszy z pustym Description")
        dax = dax[~empty_desc].reset_index(drop=True)

    # Wykryj datę z DAX
    date_col = [c for c in dax.columns if "Date" in c][0]
    dax_date = dax[date_col].dropna().iloc[0]
    print(f"  Data wyciągu: {dax_date}")

    # Wczytaj zwroty
    zwroty, zwroty_total = load_zwroty(zwroty_file, dax_date)
    print(f"  Zwroty: {len(zwroty)} pozycji (suma: {zwroty_total:.2f})")
    print()

    # Preprzetworz DWH raz (dla wydajności)
    dwh = dwh.copy()
    dwh["Name_norm"] = dwh["Name"].apply(lambda x: normalize(str(x)).replace("-", " "))
    dwh["Amount_parsed"] = dwh["Amount"].apply(parse_amount)

    # KROK 1: Oznacz zwroty — każdy zwrot z pliku może być użyty tylko raz
    zwrot_count = 0
    zwrot_sum = 0.0
    used_zwroty = set()

    if len(zwroty) > 0:
        print("=== KROK 1: Oznaczanie zwrotów ===")
        for idx, row in dax.iterrows():
            desc_words = extract_name_words(row.get("Description", ""))
            amount_dax = parse_amount(row.get("CreditAmount"))

            if not desc_words or amount_dax is None:
                continue

            arkusz = match_zwrot(desc_words, amount_dax, zwroty, used_zwroty)
            if arkusz is not None:
                dax.at[idx, "AccountNum"] = ZWROT_ACCOUNT
                date_parts = arkusz.replace(".2026", "").replace(".2025", "")
                orig_desc = str(row.get("Description", ""))
                orig_desc = re.sub(r"^zwrocone dn\.\s*[\d.]+\s*", "", orig_desc)
                dax.at[idx, "Description"] = f"zwrocone dn. {date_parts} " + orig_desc
                zwrot_count += 1
                zwrot_sum += amount_dax
                print(f"  ZWROT wiersz {idx:>3} | {str(row.get('Description',''))[:50]:<50} | kwota={amount_dax} | arkusz={arkusz}")

        print(f"\n  Oznaczono zwrotów: {zwrot_count}, suma: {zwrot_sum:.2f}")
        if abs(zwrot_sum - zwroty_total) < 0.01:
            print(f"  Suma zwrotów ZGADZA SIĘ z oczekiwaną ({zwroty_total:.2f})")
        else:
            print(f"  UWAGA: Suma zwrotów ({zwrot_sum:.2f}) != oczekiwana ({zwroty_total:.2f})!")
        print()

    # KROK 1b: Pozycje po debecie (nadpisuje istniejący AccountNum):
    #   z "Retragere" w opisie -> login z tytułu (liczba po słowie Retragere)
    #   bez "Retragere"        -> zwrot 240/04
    print("=== KROK 1b: Pozycje po debecie ===")
    debit_zwrot = 0
    debit_retragere = 0
    retragere_bez_loginu = 0
    for idx, row in dax.iterrows():
        debit = parse_amount(row.get("DebitAmount"))
        if not debit or debit <= 0:
            continue
        desc = str(row.get("Description", ""))
        retragere = re.search(r"retragere\s+(\d{5,8})\b", desc, flags=re.IGNORECASE)
        if retragere:
            dax.at[idx, "AccountNum"] = retragere.group(1)
            debit_retragere += 1
        elif "retragere" in desc.lower():
            retragere_bez_loginu += 1
            print(f"  RETRAGERE BEZ LOGINU wiersz {idx:>3} | {desc[:60]:<60} | debit={debit} - zostawiam")
        else:
            stary = row.get("AccountNum")
            dax.at[idx, "AccountNum"] = ZWROT_ACCOUNT
            debit_zwrot += 1
            if pd.notna(stary) and str(stary).strip() and str(stary).strip() != ZWROT_ACCOUNT:
                print(f"  DEBIT→240/04 wiersz {idx:>3} | {desc[:50]:<50} | debit={debit} | był login {stary}")

    print(f"\n  Debet → 240/04: {debit_zwrot}")
    print(f"  Debet Retragere → login z tytułu: {debit_retragere}")
    if retragere_bez_loginu:
        print(f"  UWAGA: Retragere bez loginu w tytule (do sprawdzenia ręcznie): {retragere_bez_loginu}")
    print()

    # KROK 2: Uzupełnij loginy (tylko te bez AccountNum)
    print("=== KROK 2: Uzupełnianie loginów ===")
    filled = 0
    skipped_already = 0
    not_found = 0

    for idx, row in dax.iterrows():
        if pd.notna(row.get("AccountNum")) and str(row["AccountNum"]).strip():
            skipped_already += 1
            continue

        desc_words = extract_name_words(row.get("Description", ""))
        amount_dax = parse_amount(row.get("CreditAmount"))

        login, info = find_login(desc_words, amount_dax, dwh)

        print(f"  Wiersz {idx:>3} | Description: {str(row.get('Description',''))[:50]:<50} | {info}")

        if login:
            dax.at[idx, "AccountNum"] = login
            filled += 1
        else:
            not_found += 1

    print()
    print(f"Wyniki:")
    print(f"  Zwroty z pliku (240/04):   {zwrot_count}")
    print(f"  Debet → 240/04:            {debit_zwrot}")
    print(f"  Debet Retragere → login:   {debit_retragere}")
    print(f"  Uzupełniono loginy:        {filled}")
    print(f"  Nie znaleziono:            {not_found}")
    print(f"  Już miało AccountNum:      {skipped_already - zwrot_count - debit_zwrot - debit_retragere}")

    dax["BranchCC"] = "20024"
    dax.to_csv(dax_file, sep=";", index=False, encoding="utf-8-sig")
    print(f"\nZapisano wynik do: {dax_file}")

    rozszerz_axtable1()

    print("\n=== Zaznaczanie błędów w raporcie ===")
    zaznacz_bledy_dax.PLIK = os.path.join(FOLDER, RAPORT_FILE)
    zaznacz_bledy_dax.main()


def rozszerz_axtable1():
    """Rozszerza tabelę AxTable1 w arkuszu z wyciągiem do ostatniego wiersza z danymi (przez Excela)."""
    raport_path = os.path.join(FOLDER, RAPORT_FILE)
    if not os.path.exists(raport_path):
        return

    print("\n=== Rozszerzanie tabeli AxTable1 ===")
    if raport_otwarty(raport_path):
        print(f"BŁĄD: {RAPORT_FILE} jest otwarty w Excelu - zamknij go i uruchom ponownie.")
        sys.exit(1)

    xl = otworz_excel()
    try:
        wb = xl.Workbooks.Open(raport_path)
        ws = wb.Worksheets(1)
        tabele = [t for t in ws.ListObjects if t.Name == "AxTable1"]
        if not tabele:
            print("  Nie znaleziono tabeli AxTable1. Pomijam.")
            wb.Close(False)
            return
        table = tabele[0]
        last_data = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        old_ref = table.Range.Address.replace("$", "")
        old_end_row = table.Range.Row + table.Range.Rows.Count - 1

        if last_data <= old_end_row:
            print(f"  AxTable1 ({old_ref}) już obejmuje wszystkie dane (ostatni wiersz: {last_data}). OK.")
            wb.Close(False)
            return

        ostatnia_kol = table.Range.Column + table.Range.Columns.Count - 1
        table.Resize(ws.Range(ws.Cells(1, 1), ws.Cells(last_data, ostatnia_kol)))
        print(f"  AxTable1: {old_ref} → {table.Range.Address.replace('$', '')}")
        wb.Save()
        wb.Close()
        print(f"  Zapisano: {RAPORT_FILE}")
    finally:
        xl.Quit()


if __name__ == "__main__":
    main()
