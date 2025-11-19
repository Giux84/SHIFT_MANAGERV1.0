"""
Conta i giorni lavorativi effettivi (A+B+C+G+F*) nel file ufficiale
"""
import pandas as pd

file_ufficiale = r"D:\Users\gcaravel\OneDrive - ram.it\Desktop\TURNO COMPLETO 2025 con modifiche da CCNL.xls"

print("="*80)
print("CONTEGGIO GIORNI LAVORATIVI EFFETTIVI - FILE UFFICIALE")
print("="*80)

excel_uff = pd.ExcelFile(file_ufficiale, engine='xlrd')

mesi = ['GENNAIO', 'FEBBRAIO', 'MARZO', 'APRILE', 'MAGGIO', 'GIUGNO',
        'LUGLIO', 'AGOSTO', 'SETTEMBRE', 'OTTOBRE', 'NOVEMBRE', 'DICEMBRE']

turni = [41, 42, 43, 44, 45, 47, 48, 49, 50, 51]

print("\nCONTEGGIO PER TURNO:")
print("-" * 80)

for turno in turni:
    total_lavorativi = 0  # A+B+C+G+F*
    total_solo_shifts = 0  # A+B+C+G (no ferie)
    total_ferie = 0  # F*

    for mese in mesi:
        try:
            df = pd.read_excel(excel_uff, sheet_name=mese, header=1)
            riga = df[df.iloc[:, 0] == turno]

            if not riga.empty:
                for col in range(1, len(riga.columns)):
                    val = riga.iloc[0, col]
                    if pd.notna(val):
                        val_str = str(val).strip()
                        # Conta shifts effettivi
                        if val_str in ['A', 'B', 'C', 'G']:
                            total_solo_shifts += 1
                            total_lavorativi += 1
                        # Conta ferie
                        elif val_str.startswith('F'):
                            total_ferie += 1
                            total_lavorativi += 1
        except:
            pass

    print(f"Turno {turno}:")
    print(f"  Shifts (A+B+C+G): {total_solo_shifts}")
    print(f"  Ferie (F*): {total_ferie}")
    print(f"  TOTALE LAVORATIVI: {total_lavorativi}")
    print()

print("="*80)
