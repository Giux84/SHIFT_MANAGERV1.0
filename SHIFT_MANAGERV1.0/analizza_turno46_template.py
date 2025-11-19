"""
Analizza il turno 46 nel TEMPLATE 2024
"""
import pandas as pd

file_template = r"D:\Users\gcaravel\OneDrive - ram.it\Desktop\TURNO COMPLETO 2025 con modifiche da CCNL.xls"

excel_template = pd.ExcelFile(file_template, engine='xlrd')

print("="*80)
print("ANALISI TURNO 46 - TEMPLATE 2024 (DICEMBRE)")
print("="*80)

# Leggi dicembre 2024 (che uso come template)
df = pd.read_excel(excel_template, sheet_name='DICEMBRE', header=1)

turno = 46
riga = df[df.iloc[:, 0] == turno]

if not riga.empty:
    print("\nPATTERN TURNO 46 - DICEMBRE 2024 (ultimi 10 giorni):")
    print("-" * 80)

    # Ultimi 10 giorni di dicembre (22-31)
    print("Giorni 22-31:")
    ultimi_giorni = []
    for giorno in range(22, 32):
        val = riga.iloc[0, giorno]
        if pd.isna(val):
            ultimi_giorni.append('-')
        else:
            ultimi_giorni.append(str(val).strip())

    print(f"  {' '.join(ultimi_giorni)}")

    # Ultimi 3 giorni (che uso per calcolare offset)
    print("\nUltimi 3 giorni (29-31) per calcolo offset:")
    ultimi_3 = []
    for giorno in range(29, 32):
        val = riga.iloc[0, giorno]
        if pd.isna(val):
            ultimi_3.append('-')
        else:
            ultimi_3.append(str(val).strip())
    print(f"  {' '.join(ultimi_3)}")

    # Conta G in dicembre
    g_count = 0
    for col in range(1, len(riga.columns)):
        val = riga.iloc[0, col]
        if pd.notna(val) and str(val).strip() == 'G':
            g_count += 1

    print(f"\nGiorni G in DICEMBRE 2024: {g_count}")

    # Mostra pattern completo dicembre
    print("\nPattern completo DICEMBRE 2024:")
    tutti_giorni = []
    for giorno in range(1, 32):
        val = riga.iloc[0, giorno]
        if pd.isna(val):
            tutti_giorni.append('-')
        else:
            tutti_giorni.append(str(val).strip())

    # Stampa in gruppi di 10
    for i in range(0, len(tutti_giorni), 10):
        gruppo = tutti_giorni[i:i+10]
        print(f"  Giorni {i+1:2d}-{min(i+10, 31):2d}: {' '.join(gruppo)}")

print("\n" + "="*80)
print("PATTERN_46_NORMALE nel codice:")
print("-" * 80)
print("Attuale: ['G', 'G', '-', 'G', 'G', '-', 'G', 'G', '-', '-']")
print("\nVerifica se il pattern nel template corrisponde a quello nel codice.")
print("="*80)
