"""
Analizza il turno 46 nel file ufficiale
"""
import pandas as pd

file_ufficiale = r"D:\Users\gcaravel\OneDrive - ram.it\Desktop\TURNO COMPLETO 2025 con modifiche da CCNL.xls"

excel_uff = pd.ExcelFile(file_ufficiale, engine='xlrd')

print("="*80)
print("ANALISI TURNO 46 - FILE UFFICIALE")
print("="*80)

mesi = ['GENNAIO', 'FEBBRAIO', 'MARZO', 'APRILE', 'MAGGIO', 'GIUGNO',
        'LUGLIO', 'AGOSTO', 'SETTEMBRE', 'OTTOBRE', 'NOVEMBRE', 'DICEMBRE']

turno = 46

print("\nPattern Turno 46 per mese (primi 15 giorni):")
print("-" * 80)

for mese in mesi:
    df = pd.read_excel(excel_uff, sheet_name=mese, header=1)
    riga = df[df.iloc[:, 0] == turno]

    if not riga.empty:
        # Primi 15 giorni
        giorni = []
        for col in range(1, min(16, len(riga.columns))):
            val = riga.iloc[0, col]
            if pd.isna(val):
                giorni.append('-')
            else:
                giorni.append(str(val).strip())

        print(f"{mese:10s}: {' '.join(giorni)}")

# Conta totale G per turno 46
print("\n" + "="*80)
print("CONTEGGIO GIORNI TURNO 46:")
print("-" * 80)

totale_g = 0
totale_riposi = 0
totale_ferie = 0
totale_lavorativi = 0

for mese in mesi:
    df = pd.read_excel(excel_uff, sheet_name=mese, header=1)
    riga = df[df.iloc[:, 0] == turno]

    if not riga.empty:
        for col in range(1, len(riga.columns)):
            val = riga.iloc[0, col]
            if pd.notna(val):
                val_str = str(val).strip()
                if val_str == 'G':
                    totale_g += 1
                    totale_lavorativi += 1
                elif val_str == '-':
                    totale_riposi += 1
                elif val_str.startswith('F'):
                    totale_ferie += 1
                    totale_lavorativi += 1
                else:
                    # Altri codici (A, B, C se presenti)
                    totale_lavorativi += 1

print(f"Giorni G: {totale_g}")
print(f"Riposi (-): {totale_riposi}")
print(f"Ferie (F*): {totale_ferie}")
print(f"TOTALE LAVORATIVI: {totale_lavorativi}")

print("\n" + "="*80)
