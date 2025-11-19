"""
Verifica in che giorni del mese sono i G nel file ufficiale
"""
import pandas as pd

file_ufficiale = r"D:\Users\gcaravel\OneDrive - ram.it\Desktop\TURNO COMPLETO 2025 con modifiche da CCNL.xls"

excel_uff = pd.ExcelFile(file_ufficiale, engine='xlrd')

print("="*80)
print("VERIFICA POSIZIONE G NEL MESE - FILE UFFICIALE")
print("="*80)

mesi = ['GENNAIO', 'FEBBRAIO', 'MARZO', 'APRILE', 'MAGGIO', 'OTTOBRE']
turni = [41, 42, 43, 44, 45, 47, 48, 49, 50, 51]

g_prima_15 = []
g_dopo_15 = []

for mese in mesi:
    df = pd.read_excel(excel_uff, sheet_name=mese, header=1)

    print(f"\n{mese}:")
    print("-" * 80)

    for turno in turni:
        riga = df[df.iloc[:, 0] == turno]

        if not riga.empty:
            posizioni_g = []

            for giorno in range(1, 32):
                try:
                    val = riga.iloc[0, giorno]
                    if pd.notna(val) and str(val).strip() == 'G':
                        posizioni_g.append(giorno)
                        if giorno < 15:
                            g_prima_15.append((mese, turno, giorno))
                        else:
                            g_dopo_15.append((mese, turno, giorno))
                except:
                    break

            if posizioni_g:
                print(f"  Turno {turno}: G nei giorni {posizioni_g}")

print("\n" + "="*80)
print("STATISTICHE:")
print("-" * 80)
print(f"G prima del giorno 15: {len(g_prima_15)}")
print(f"G dal giorno 15 in poi: {len(g_dopo_15)}")

if g_prima_15:
    print("\nEsempi G prima del giorno 15:")
    for mese, turno, giorno in g_prima_15[:10]:
        print(f"  {mese} - Turno {turno} - Giorno {giorno}")

print("\n" + "="*80)
if g_prima_15:
    print("CONCLUSIONE: Ci sono G PRIMA del giorno 15!")
    print("La regola 'G solo dopo giorno 14' è troppo restrittiva.")
else:
    print("CONCLUSIONE: Tutti i G sono dal giorno 15 in poi.")
    print("La regola 'G solo dopo giorno 14' è corretta.")
print("="*80)
