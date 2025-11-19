"""
Verifica se nel file ufficiale ci sono G in weekend o festivi
"""
import pandas as pd
from datetime import date

file_ufficiale = r"D:\Users\gcaravel\OneDrive - ram.it\Desktop\TURNO COMPLETO 2025 con modifiche da CCNL.xls"

# Festività italiane 2025
FESTIVITA_2025 = [
    (1, 1),    # Capodanno
    (1, 6),    # Epifania
    (4, 20),   # Pasqua
    (4, 21),   # Pasquetta
    (4, 25),   # Liberazione
    (5, 1),    # Festa dei Lavoratori
    (6, 2),    # Festa della Repubblica
    (8, 15),   # Ferragosto
    (11, 1),   # Tutti i Santi
    (12, 8),   # Immacolata
    (12, 25),  # Natale
    (12, 26),  # Santo Stefano
]

def is_weekend_or_festivo(d: date) -> bool:
    """Verifica se una data è weekend o festivo"""
    if d.weekday() in [5, 6]:  # Sabato=5, Domenica=6
        return True
    if (d.month, d.day) in FESTIVITA_2025:
        return True
    return False

print("="*80)
print("VERIFICA G IN WEEKEND/FESTIVI - FILE UFFICIALE")
print("="*80)

excel_uff = pd.ExcelFile(file_ufficiale, engine='xlrd')

mesi = ['GENNAIO', 'FEBBRAIO', 'MARZO', 'APRILE', 'MAGGIO', 'GIUGNO',
        'LUGLIO', 'AGOSTO', 'SETTEMBRE', 'OTTOBRE', 'NOVEMBRE', 'DICEMBRE']

turni = [41, 42, 43, 44, 45, 47, 48, 49, 50, 51]

g_in_weekend = []

for mese_idx, mese in enumerate(mesi, 1):
    df = pd.read_excel(excel_uff, sheet_name=mese, header=1)

    for turno in turni:
        riga = df[df.iloc[:, 0] == turno]
        if not riga.empty:
            for giorno in range(1, 32):
                try:
                    val = riga.iloc[0, giorno]
                    if val == 'G':
                        data = date(2025, mese_idx, giorno)
                        if is_weekend_or_festivo(data):
                            g_in_weekend.append({
                                'turno': turno,
                                'data': data,
                                'giorno_settimana': data.strftime('%A')
                            })
                except:
                    break

print(f"\nG in weekend o festivi: {len(g_in_weekend)}")

if g_in_weekend:
    for item in g_in_weekend:
        print(f"  Turno {item['turno']}: {item['data']} ({item['giorno_settimana']})")
else:
    print("  Nessun G in weekend o festivi trovato!")

print("\n" + "="*80)
print("CONCLUSIONE:")
if g_in_weekend:
    print("  Il filtro weekend/festivi è troppo restrittivo!")
else:
    print("  Il filtro weekend/festivi è corretto - nessun G in weekend/festivi nel file ufficiale")
print("="*80)
