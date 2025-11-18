# QUICK START - Generatore Turni Raffineria

## AVVIO RAPIDO

### 1. Prima Esecuzione (Anno 2025)

```bash
cd "D:\Users\gcaravel\OneDrive - ram.it\Desktop\PROGETTI_PYTHON_OFFICE\SHIFT_MANAGERV1.0"
python turni_generator.py
```

**Output generato**:
- `TURNO_COMPLETO_2025.xlsx` - File Excel con tutti i turni
- `report_verifica.txt` - Report anomalie copertura

### 2. Verifica Risultati

1. **Apri** `TURNO_COMPLETO_2025.xlsx`
2. **Controlla** i 12 fogli mensili (GENNAIO, FEBBRAIO, ...)
3. **Verifica** il report: `report_verifica.txt`
4. **Attenzione**: Periodo giugno-settembre richiede verifica manuale

### 3. Anni Successivi (2026, 2027, ...)

**Passo 1**: Aggiorna anno target in `turni_generator.py`

```python
# Linea ~598 nella funzione main()
anno_target = 2026  # Cambia qui
```

**Passo 2**: Aggiorna path file template

```python
# Linea ~597
file_template = r"D:\Users\gcaravel\OneDrive - ram.it\Desktop\TURNO COMPLETO 2025.xls"
```

**Passo 3**: Aggiungi configurazione ferie per il nuovo anno

```python
# Dopo FERIE_2025 (linea ~100), aggiungi:
FERIE_2026 = {
    1: [42, 48],  # Era periodo 4 nel 2025
    2: [41, 47],  # Era periodo 5 nel 2025
    3: [44, 50],  # Era periodo 1 nel 2025
    4: [43, 49],  # Era periodo 2 nel 2025
    5: [45, 51],  # Era periodo 3 nel 2025
    6: [46]       # Sempre fisso
}
```

**Passo 4**: Usa FERIE_2026 nel codice

```python
# Linea ~262, sostituisci:
def applica_ferie(self):
    """Applica le ferie secondo FERIE_2026"""  # Cambia qui
    print("\nApplicazione ferie...")

    for periodo, turni_in_ferie in FERIE_2026.items():  # Cambia qui
```

**Passo 5**: Esegui

```bash
python turni_generator.py
```

## COSA CONTROLLARE

### ✅ Controlli Automatici OK
- Continuità ciclo 2024→2025
- Pattern normale (10 giorni)
- Pattern estivo (9 giorni)
- Formattazione Excel
- Conteggio giorni lavorativi

### ⚠️ Controlli Manuali Richiesti

**1. Periodo Estivo (Giugno-Settembre)**
   - Date: 20/06 - 13/09
   - Verificare transizioni pattern
   - Controllare sovrapposizioni ferie

**2. Copertura 2-2-2**
   - Ogni giorno deve avere: 2A, 2B, 2C
   - Il report segnala le anomalie
   - Anomalie estive sono normali (ferie)

**3. Turno 46**
   - Periodo normale: solo G
   - Periodo estivo: A/B/C
   - Ferie: FG (periodo 6)

**4. Festività**
   - G mai in weekend/festivi
   - Aggiornare lista festività per anno

## TABELLA ROTAZIONE RAPIDA

| Anno | 41-47 | 42-48 | 43-49 | 44-50 | 45-51 | 46 |
|------|-------|-------|-------|-------|-------|----|
| 2024 | P3    | P2    | P5    | P4    | P1    | P6 |
| 2025 | P5    | P4    | P2    | P1    | P3    | P6 |
| 2026 | P2    | P1    | P4    | P3    | P5    | P6 |
| 2027 | P4    | P3    | P1    | P5    | P2    | P6 |
| 2028 | P1    | P5    | P3    | P2    | P4    | P6 |
| 2029 | P3    | P2    | P5    | P4    | P1    | P6 |

**P1**: 20/06-08/07 | **P2**: 08/07-25/07 | **P3**: 25/07-11/08
**P4**: 11/08-28/08 | **P5**: 28/08-14/09 | **P6**: 14/09-30/09

## RISOLUZIONE PROBLEMI

### Problema: "File template non trovato"
**Soluzione**: Verifica percorso file in `main()` (linea ~597)

### Problema: "Troppe anomalie nel report"
**Soluzione**: Normale per giugno-settembre (ferie). Controlla date ferie.

### Problema: "G in weekend"
**Soluzione**: Aggiorna `FESTIVITA_2025` con festività corrette

### Problema: "Offset sbagliati"
**Soluzione**: Verifica che il file template sia dell'anno precedente corretto

## FILE GENERATI

```
📁 SHIFT_MANAGERV1.0/
├── 📄 turni_generator.py         [Script principale]
├── 📄 README.md                  [Documentazione completa]
├── 📄 ROTAZIONE_FERIE.md         [Guida rotazione quinquennale]
├── 📄 QUICK_START.md             [Questa guida]
├── 📊 TURNO_COMPLETO_2025.xlsx   [Output turni 2025]
└── 📋 report_verifica.txt        [Report anomalie]
```

## WORKFLOW ANNUALE

```
1. Fine anno precedente (Dicembre)
   └─> Eseguire generatore con anno+1

2. Inizio anno nuovo (Gennaio)
   └─> Verificare output manualmente
   └─> Correggere anomalie se necessarie
   └─> Distribuire file finale

3. Durante l'anno
   └─> Monitorare eventuali modifiche
   └─> Mantenere backup file originale
```

## CONTATTI E SUPPORTO

Per problemi tecnici:
1. Consultare `README.md` per documentazione completa
2. Consultare `ROTAZIONE_FERIE.md` per logica rotazione
3. Verificare `report_verifica.txt` per anomalie specifiche

## CHECKLIST DISTRIBUZIONE

Prima di distribuire il file Excel:

- [ ] Eseguito generatore senza errori
- [ ] Verificato report_verifica.txt
- [ ] Controllati periodi ferie (giugno-settembre)
- [ ] Verificato turno 46 speciale
- [ ] Controllate festività anno corrente
- [ ] Verificata formattazione Excel
- [ ] Backup file precedente effettuato
- [ ] Anomalie documentate e approvate

---

**Versione**: 1.0
**Ultimo aggiornamento**: 2025
**Tempo esecuzione**: ~5-10 secondi
