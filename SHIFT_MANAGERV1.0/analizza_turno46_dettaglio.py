"""
Analisi dettagliata turno 46 con conteggio corretto
"""
import pandas as pd

file_ufficiale = r"D:\Users\gcaravel\OneDrive - ram.it\Desktop\TURNO COMPLETO 2025 con modifiche da CCNL.xls"

excel_uff = pd.ExcelFile(file_ufficiale, engine='xlrd')

print("="*80)
print("ANALISI DETTAGLIATA TURNO 46 - FILE UFFICIALE")
print("="*80)

mesi = ['GENNAIO', 'FEBBRAIO', 'MARZO', 'APRILE', 'MAGGIO', 'GIUGNO',
        'LUGLIO', 'AGOSTO', 'SETTEMBRE', 'OTTOBRE', 'NOVEMBRE', 'DICEMBRE']

turno = 46

# Conteggio mese per mese
print("\nCONTEGGIO PER MESE:")
print("-" * 80)
print(f"{'Mese':<12} | {'G':>3} | {'A':>3} | {'B':>3} | {'C':>3} | {'F*':>3} | {'Totale Lav':>11} | {'Riposi':>7}")
print("-" * 80)

totale_g = 0
totale_abc = 0
totale_ferie = 0
totale_riposi = 0

for mese in mesi:
    df = pd.read_excel(excel_uff, sheet_name=mese, header=1)
    riga = df[df.iloc[:, 0] == turno]

    if not riga.empty:
        g_count = 0
        a_count = 0
        b_count = 0
        c_count = 0
        f_count = 0
        riposi = 0

        for col in range(1, len(riga.columns)):
            val = riga.iloc[0, col]
            if pd.notna(val):
                val_str = str(val).strip()
                if val_str == 'G':
                    g_count += 1
                elif val_str == 'A':
                    a_count += 1
                elif val_str == 'B':
                    b_count += 1
                elif val_str == 'C':
                    c_count += 1
                elif val_str.startswith('F'):
                    f_count += 1
                elif val_str == '-':
                    riposi += 1

        lav_mese = g_count + a_count + b_count + c_count + f_count
        print(f"{mese:<12} | {g_count:3d} | {a_count:3d} | {b_count:3d} | {c_count:3d} | {f_count:3d} | {lav_mese:11d} | {riposi:7d}")

        totale_g += g_count
        totale_abc += (a_count + b_count + c_count)
        totale_ferie += f_count
        totale_riposi += riposi

print("-" * 80)
totale_lavorativi = totale_g + totale_abc + totale_ferie
print(f"{'TOTALE':<12} | {totale_g:3d} | {totale_abc:3d} (ABC) | {totale_ferie:3d} | {totale_lavorativi:11d} | {totale_riposi:7d}")

print("\n" + "="*80)
print("RIEPILOGO:")
print("-" * 80)
print(f"Giorni G totali: {totale_g}")
print(f"Giorni A+B+C totali: {totale_abc}")
print(f"Ferie totali: {totale_ferie}")
print(f"TOTALE GIORNI LAVORATIVI: {totale_lavorativi}")
print(f"Riposi totali: {totale_riposi}")

# Leggi il progressivo da DICEMBRE
df_dic = pd.read_excel(excel_uff, sheet_name='DICEMBRE', header=1)
riga_dic = df_dic[df_dic.iloc[:, 0] == turno]
if not riga_dic.empty:
    # Il progressivo è nell'ultima colonna
    progressivo = riga_dic.iloc[0, -1]
    print(f"\nProgressivo da colonna 'Progr.' in DICEMBRE: {progressivo}")

print("\n" + "="*80)

# Verifica ferie settembre
print("\nVERIFICA FERIE SETTEMBRE:")
print("-" * 80)
df_set = pd.read_excel(excel_uff, sheet_name='SETTEMBRE', header=1)
riga_set = df_set[df_set.iloc[:, 0] == turno]
if not riga_set.empty:
    print("Giorni di SETTEMBRE per Turno 46:")
    for giorno in range(1, 31):
        val = riga_set.iloc[0, giorno]
        if pd.notna(val):
            val_str = str(val).strip()
            if val_str.startswith('F'):
                print(f"  Giorno {giorno}: {val_str}")

print("\n" + "="*80)
