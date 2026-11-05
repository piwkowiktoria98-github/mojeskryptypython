"""
Porównuje dzienne sumy transakcji: raport PayPal vs zapisy na koncie w DAX.

Mapowanie:
  PayPal Express Checkout Payment - Gross  <->  DAX 0135 Depozyt
  PayPal Express Checkout Payment - Fee    <->  DAX 0136 prowizja
  PayPal General Hold                      <->  DAX 0139 HOLD (kwoty ujemne)
  PayPal General Hold Release              <->  DAX 0139 HOLD (kwoty dodatnie)

Wynik: plik "PayPal vs DAX - porownanie.xlsx" z arkuszami:
  - Porownanie dni: sumy dzienne, różnice, dni niezgodne na czerwono
  - Brakujace transakcje: transakcje, które są tylko po jednej stronie
"""

import os
import sys
import glob
import pandas as pd
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8")

FOLDER = r"C:\Users\wpiwko\Desktop\Paypal sprawdzanie"
OUTPUT_NAME = "PayPal vs DAX - porownanie.xlsx"

KATEGORIE = ["Depozyt", "Prowizja", "Hold", "Hold Release"]
RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
BOLD = Font(bold=True)


def znajdz_pliki():
    """Rozpoznaje pliki po zawartości: PayPal ma kolumnę Gross, DAX ma kolumnę Opis."""
    paypal = dax = None
    for path in glob.glob(os.path.join(FOLDER, "*.xlsx")):
        name = os.path.basename(path)
        if name.startswith("~$") or name == OUTPUT_NAME:
            continue
        for sheet, df in pd.read_excel(path, sheet_name=None, nrows=1).items():
            if "Gross" in df.columns and "Transaction ID" in df.columns:
                paypal = (path, sheet)
            elif "Opis" in df.columns and "Dokument" in df.columns:
                dax = (path, sheet)
    if paypal is None or dax is None:
        print("BŁĄD: Nie znaleziono raportu PayPal (kolumna Gross) lub DAX (kolumna Opis).")
        sys.exit(1)
    return paypal, dax


def wczytaj_paypal(path, sheet):
    df = pd.read_excel(path, sheet_name=sheet)
    df["Data"] = pd.to_datetime(df["Date"]).dt.date
    df["ID"] = df["Transaction ID"].astype(str).str.strip()
    wiersze = []
    for _, r in df.iterrows():
        if r["Type"] == "Express Checkout Payment":
            wiersze.append((r["Data"], r["ID"], "Depozyt", r["Gross"], r["Name"]))
            wiersze.append((r["Data"], r["ID"], "Prowizja", r["Fee"], r["Name"]))
        elif r["Type"] == "General Hold":
            wiersze.append((r["Data"], r["ID"], "Hold", r["Gross"], r["Name"]))
        elif r["Type"] == "General Hold Release":
            wiersze.append((r["Data"], r["ID"], "Hold Release", r["Gross"], r["Name"]))
        else:
            print(f"  UWAGA: nieznany typ PayPal '{r['Type']}' ({r['ID']}) - pomijam")
    return pd.DataFrame(wiersze, columns=["Data", "ID", "Kategoria", "Kwota", "Nazwa"])


def wczytaj_dax(path, sheet):
    df = pd.read_excel(path, sheet_name=sheet)
    df["Data"] = pd.to_datetime(df["Data"]).dt.date
    # Dokument: "D" + Transaction ID PayPal + "_MMDD_kod", np. D9UT75335DV164015X_1001_0135
    df["ID"] = df["Dokument"].astype(str).str.split("_").str[0].str[1:]
    kwota = df["Kwota w walucie transakcji"]
    opis = df["Opis"].astype(str)
    df["Kategoria"] = None
    df.loc[opis.str.startswith("0135"), "Kategoria"] = "Depozyt"
    df.loc[opis.str.startswith("0136"), "Kategoria"] = "Prowizja"
    df.loc[opis.str.startswith("0139") & (kwota < 0), "Kategoria"] = "Hold"
    df.loc[opis.str.startswith("0139") & (kwota >= 0), "Kategoria"] = "Hold Release"
    nieznane = df[df["Kategoria"].isna()]
    for _, r in nieznane.iterrows():
        print(f"  UWAGA: nieznany opis DAX '{r['Opis']}' ({r['Dokument']}) - pomijam")
    df = df[df["Kategoria"].notna()]
    out = df[["Data", "ID", "Kategoria"]].copy()
    out["Kwota"] = kwota[df.index]
    out["Nazwa"] = df["Odbiorca"].astype(str)
    return out


def sumy_dzienne(df):
    return df.pivot_table(index="Data", columns="Kategoria", values="Kwota",
                          aggfunc="sum", fill_value=0).reindex(columns=KATEGORIE, fill_value=0)


def main():
    (pp_path, pp_sheet), (dax_path, dax_sheet) = znajdz_pliki()
    print(f"PayPal: {os.path.basename(pp_path)} [{pp_sheet}]")
    print(f"DAX:    {os.path.basename(dax_path)} [{dax_sheet}]\n")

    pp = wczytaj_paypal(pp_path, pp_sheet)
    dax = wczytaj_dax(dax_path, dax_sheet)

    s_pp = sumy_dzienne(pp)
    s_dax = sumy_dzienne(dax)
    dni = sorted(set(s_pp.index) | set(s_dax.index))
    s_pp = s_pp.reindex(dni, fill_value=0)
    s_dax = s_dax.reindex(dni, fill_value=0)

    porownanie = pd.DataFrame({"Data": dni})
    for k in KATEGORIE:
        porownanie[f"{k} PayPal"] = s_pp[k].values
        porownanie[f"{k} DAX"] = s_dax[k].values
        porownanie[f"{k} Roznica"] = (s_pp[k] - s_dax[k]).round(2).values
    roznice = porownanie[[f"{k} Roznica" for k in KATEGORIE]]
    porownanie["Status"] = (roznice.abs() >= 0.01).any(axis=1).map({True: "NIEZGODNE", False: "OK"})

    # Transakcje tylko po jednej stronie (dopasowanie po ID + kategorii).
    # DAX ma storna (+X i -X na tym samym ID) - sumujemy netto i pomijamy zerowe.
    klucz = ["ID", "Kategoria"]

    def netto(df):
        g = df.groupby(klucz, as_index=False).agg(Data=("Data", "min"), Kwota=("Kwota", "sum"),
                                                   Nazwa=("Nazwa", "first"))
        return g[g["Kwota"].abs() >= 0.01]

    scalone = netto(pp).merge(netto(dax), on=klucz, how="outer",
                              suffixes=(" PayPal", " DAX"), indicator=True)
    brakujace = scalone[scalone["_merge"] != "both"].copy()
    brakujace["Gdzie jest"] = brakujace["_merge"].map({"left_only": "tylko PayPal", "right_only": "tylko DAX"})
    # Ta sama transakcja po obu stronach, ale z inną kwotą albo datą
    obie = scalone[scalone["_merge"] == "both"]
    rozne = obie[((obie["Kwota PayPal"] - obie["Kwota DAX"]).abs() >= 0.01) |
                 (obie["Data PayPal"] != obie["Data DAX"])].copy()
    rozne["Gdzie jest"] = "obie strony - inna kwota/data"
    problemy = pd.concat([brakujace, rozne])
    problemy = problemy[["Gdzie jest", "Kategoria", "ID", "Data PayPal", "Kwota PayPal",
                         "Nazwa PayPal", "Data DAX", "Kwota DAX", "Nazwa DAX"]]
    problemy = problemy.sort_values(["Gdzie jest", "Data PayPal", "Data DAX"])

    # --- Konsola ---
    niezgodne = porownanie[porownanie["Status"] == "NIEZGODNE"]
    print(f"Dni łącznie: {len(porownanie)}, niezgodne: {len(niezgodne)}")
    for _, r in niezgodne.iterrows():
        opis = ", ".join(f"{k}: {r[f'{k} Roznica']:,.2f}".replace(",", " ")
                         for k in KATEGORIE if abs(r[f"{k} Roznica"]) >= 0.01)
        print(f"  {r['Data']}  {opis}")
    print(f"\nTransakcji do sprawdzenia: {len(problemy)}")
    for g, n in problemy["Gdzie jest"].value_counts().items():
        print(f"  {g}: {n}")

    # --- Excel ---
    out_path = os.path.join(FOLDER, OUTPUT_NAME)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        porownanie.to_excel(writer, sheet_name="Porownanie dni", index=False)
        problemy.to_excel(writer, sheet_name="Brakujace transakcje", index=False)

        ws = writer.sheets["Porownanie dni"]
        status_col = porownanie.columns.get_loc("Status") + 1
        for r in range(2, ws.max_row + 1):
            if ws.cell(r, status_col).value == "NIEZGODNE":
                for c in range(1, ws.max_column + 1):
                    ws.cell(r, c).fill = RED_FILL
            ws.cell(r, 1).number_format = "yyyy-mm-dd"
            for c in range(2, status_col):
                ws.cell(r, c).number_format = "#,##0.00"

        ws2 = writer.sheets["Brakujace transakcje"]
        for r in range(2, ws2.max_row + 1):
            for c in (4, 7):
                ws2.cell(r, c).number_format = "yyyy-mm-dd"
            for c in (5, 8):
                ws2.cell(r, c).number_format = "#,##0.00"

        for ws_ in (ws, ws2):
            ws_.freeze_panes = "B2"
            for c in range(1, ws_.max_column + 1):
                ws_.cell(1, c).font = BOLD
                ws_.column_dimensions[get_column_letter(c)].width = 16

    print(f"\nZapisano: {out_path}")


if __name__ == "__main__":
    main()
