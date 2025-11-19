"""
Analizza la distribuzione dei giorni G nel file ufficiale
"""
import pandas as pd

file_ufficiale = r"D:\Users\gcaravel\OneDrive - ram.it\Desktop\TURNO COMPLETO 2025 con modifiche da CCNL.xls"

print("="*80)
print("ANALISI DISTRIBUZIONE GIORNI G - FILE UFFICIALE")
print("="*80)

excel_uff = pd.ExcelFile(file_ufficiale, engine='xlrd')

mesi = ['GENNAIO', 'FEBBRAIO', 'MARZO', 'APRILE', 'MAGGIO', 'GIUGNO',
        'LUGLIO', 'AGOSTO', 'SETTEMBRE', 'OTTOBRE', 'NOVEMBRE', 'DICEMBRE']

turni = [41, 42, 43, 44, 45, 47, 48, 49, 50, 51]

# Conteggio G per turno per mese
print("\nDISTRIBUZIONE G PER MESE:")
print("-" * 80)

g_per_turno_per_mese = {}

for turno in turni:
    g_per_turno_per_mese[turno] = {}

    for mese in mesi:
        try:
            df = pd.read_excel(excel_uff, sheet_name=mese, header=1)
            riga = df[df.iloc[:, 0] == turno]

            if not riga.empty:
                g_count = 0
                for col in range(1, len(riga.columns)):
                    val = riga.iloc[0, col]
                    if val == 'G':
                        g_count += 1

                g_per_turno_per_mese[turno][mese] = g_count
            else:
                g_per_turno_per_mese[turno][mese] = 0
        except:
            g_per_turno_per_mese[turno][mese] = 0

# Stampa tabella
header = "Turno | " + " | ".join([m[:3] for m in mesi]) + " | TOT"
print(header)
print("-" * len(header))

for turno in turni:
    row = f"  {turno}  |"
    total = 0
    for mese in mesi:
        count = g_per_turno_per_mese[turno].get(mese, 0)
        total += count
        row += f" {count:3d} |"
    row += f" {total:3d}"
    print(row)

print("\n" + "="*80)
print("TOTALE GIORNI LAVORATIVI PER TURNO:")
print("-" * 80)

for turno in turni:
    total_giorni = 0

    for mese in mesi:
        try:
            df = pd.read_excel(excel_uff, sheet_name=mese, header=1)
            riga = df[df.iloc[:, 0] == turno]

            if not riga.empty:
                for col in range(1, len(riga.columns)):
                    val = riga.iloc[0, col]
                    if pd.notna(val):
                        val_str = str(val).strip()
                        # Conta tutti i codici tranne le ferie
                        if val_str and val_str not in ['F', 'F1', 'F2', 'F3', 'F4', 'F5', 'F6']:
                            total_giorni += 1
        except:
            pass

    print(f"Turno {turno}: {total_giorni} giorni lavorativi")

print("\n" + "="*80)
print("MESI CON ALMENO 1 G:")
print("-" * 80)

for turno in turni:
    mesi_con_g = [m for m in mesi if g_per_turno_per_mese[turno].get(m, 0) > 0]
    print(f"Turno {turno}: {', '.join(mesi_con_g)}")

print("\n" + "="*80)
