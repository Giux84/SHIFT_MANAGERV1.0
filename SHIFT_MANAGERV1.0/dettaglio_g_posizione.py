"""
Analizza DOVE sono posizionati i G nel file ufficiale
"""
import pandas as pd

file_ufficiale = r"D:\Users\gcaravel\OneDrive - ram.it\Desktop\TURNO COMPLETO 2025 con modifiche da CCNL.xls"

print("="*80)
print("ANALISI POSIZIONE GIORNI G - FILE UFFICIALE")
print("="*80)

excel_uff = pd.ExcelFile(file_ufficiale, engine='xlrd')

mesi = ['GENNAIO', 'FEBBRAIO', 'MARZO', 'APRILE', 'MAGGIO']
turni = [41, 42, 43, 44, 45]

for mese in mesi:
    print(f"\n{'='*80}")
    print(f"MESE: {mese}")
    print(f"{'='*80}")

    df = pd.read_excel(excel_uff, sheet_name=mese, header=1)

    for turno in turni:
        riga = df[df.iloc[:, 0] == turno]

        if not riga.empty:
            # Estrai tutti i giorni del mese
            giorni = []
            for col in range(1, len(riga.columns)):
                val = riga.iloc[0, col]
                if pd.isna(val):
                    giorni.append('-')
                else:
                    giorni.append(str(val).strip())

            # Trova posizioni G
            posizioni_g = [i+1 for i, g in enumerate(giorni) if g == 'G']

            if posizioni_g:
                # Mostra contesto (3 giorni prima e dopo)
                print(f"\nTurno {turno}: G nei giorni {posizioni_g}")
                for pos in posizioni_g:
                    idx = pos - 1
                    start = max(0, idx - 3)
                    end = min(len(giorni), idx + 4)

                    contesto = giorni[start:end]
                    # Evidenzia il G
                    ctx_str = []
                    for i, c in enumerate(contesto):
                        if start + i == idx:
                            ctx_str.append(f"[{c}]")
                        else:
                            ctx_str.append(c)

                    print(f"  Giorno {pos}: {' '.join(ctx_str)}")
            else:
                print(f"\nTurno {turno}: NESSUN G")

print("\n" + "="*80)
