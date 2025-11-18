# GENERATORE TURNI RAFFINERIA v1.0

Sistema automatico per la generazione dei turni annuali della raffineria, con gestione della continuità del ciclo e rotazione quinquennale delle ferie.

## 🎯 DUE MODALITÀ DISPONIBILI

### 🖥️ **INTERFACCIA GRAFICA (CONSIGLIATA)**
- **Eseguibile standalone**: `dist/ShiftManager.exe` (non serve Python!)
- Interfaccia moderna e intuitiva
- Anteprima calendario integrata
- Opzioni avanzate personalizzabili
- **👉 Leggi**: `QUICK_START_GUI.md` per iniziare subito!

### 💻 **Linea di Comando (Avanzata)**
- Script Python: `turni_generator.py`
- Per utenti tecnici o automazione
- **👉 Continua a leggere** questa guida

## CARATTERISTICHE

- **Continuità automatica**: Calcola automaticamente l'offset di partenza per il nuovo anno basandosi sull'ultimo giorno dell'anno precedente
- **Gestione pattern multipli**:
  - Periodo normale (Gen-Mag, Ott-Dic): ciclo 10 giorni
  - Periodo estivo (20 Giu - 13 Set): ciclo 9 giorni
  - Turno 46 speciale: pattern dedicato
- **Rotazione ferie quinquennale**: Applica automaticamente la rotazione 1→3→5→2→4
- **Verifica copertura**: Controlla automaticamente la copertura 2-2-2 (2 turni A, 2 B, 2 C per ogni giorno)
- **Output formattato**: Genera file Excel con formattazione professionale e report di verifica

## REQUISITI

```bash
pip install pandas openpyxl xlrd
```

## STRUTTURA FILE

```
SHIFT_MANAGERV1.0/
├── dist/
│   └── ShiftManager.exe              # 🎯 ESEGUIBILE GUI (avvia questo!)
├── turni_generator.py                # Core: generatore turni
├── shift_manager_gui.py              # GUI: applicazione principale
├── config.py                         # Gestione configurazioni
├── advanced_options_dialog.py        # Dialog opzioni avanzate
├── preview_dialog.py                 # Dialog anteprima calendario
├── build_exe.py                      # Script build eseguibile
├── requirements.txt                  # Dipendenze Python
├── README.md                         # Guida completa
├── README_GUI.md                     # Guida dettagliata GUI
├── QUICK_START_GUI.md                # 🚀 Avvio rapido GUI
├── ROTAZIONE_FERIE.md                # Documentazione ferie
└── .gitignore                        # Git ignore rules
```

## UTILIZZO

### 1. Preparazione

Assicurarsi che il file template dell'anno precedente sia disponibile:
```
D:\Users\gcaravel\OneDrive - ram.it\Desktop\TURNO COMPLETO 2024.xls
```

### 2. Esecuzione

```bash
cd "D:\Users\gcaravel\OneDrive - ram.it\Desktop\PROGETTI_PYTHON_OFFICE\SHIFT_MANAGERV1.0"
python turni_generator.py
```

### 3. Output

Il programma genera:
- `TURNO_COMPLETO_2025.xlsx` - File Excel con i turni dell'anno
- `report_verifica.txt` - Report delle anomalie di copertura

## STRUTTURA TURNI

### Turni attivi
- **11 turni totali**: 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51
- **5 coppie gemelle**: (41,47), (42,48), (43,49), (44,50), (45,51)
- **1 turno speciale**: 46 (solo giornaliero in periodo normale)

### Codici
- **A**: Mattina
- **B**: Pomeriggio
- **C**: Notte
- **G**: Giornaliero (solo Gen, Feb, Mar, Apr, Ott, Nov - mai weekend/festivi)
- **-**: Riposo
- **FA/FB/FC/FG**: Ferie (12 giorni lavorativi per turno)

## PERIODI

### Periodo Normale (Gen-Mag, Ott-Dic)
- Ciclo: 10 giorni
- Pattern: `['B', '-', 'A', 'A', '-', '-', 'C', 'C', '-', 'B']`
- G inserito in posizione 5 (dopo AA) nei mesi autorizzati

### Periodo Estivo (20 Giu - 13 Set)
- Ciclo: 9 giorni
- Pattern: `['B', '-', 'A', 'A', '-', 'C', 'C', '-', 'B']`
- Include periodi ferie obbligatorie

## FERIE 2025

### Periodi ferie
1. **Periodo 1**: 20/06 - 08/07 → Turni 44, 50
2. **Periodo 2**: 08/07 - 25/07 → Turni 43, 49
3. **Periodo 3**: 25/07 - 11/08 → Turni 45, 51
4. **Periodo 4**: 11/08 - 28/08 → Turni 42, 48
5. **Periodo 5**: 28/08 - 14/09 → Turni 41, 47
6. **Periodo 6**: 14/09 - 30/09 → Turno 46

### Rotazione quinquennale
La rotazione segue lo schema: **1→3→5→2→4**

Anno 2024 → Anno 2025:
- Periodo 1 → Periodo 3
- Periodo 2 → Periodo 4
- Periodo 3 → Periodo 5
- Periodo 4 → Periodo 1
- Periodo 5 → Periodo 2
- Periodo 6 → Periodo 6 (fisso)

## VERIFICA COPERTURA

Il programma verifica automaticamente che ogni giorno mantenga la copertura **2-2-2**:
- 2 turni in Mattina (A)
- 2 turni in Pomeriggio (B)
- 2 turni in Notte (C)

### Anomalie attese
Durante i periodi di transizione (specialmente giugno-settembre) potrebbero verificarsi anomalie dovute a:
- Sovrapposizione periodi ferie
- Cambio pattern normale → estivo
- Gestione turni gemelli

⚠️ **IMPORTANTE**: Queste anomalie richiedono **VERIFICA MANUALE** come indicato nelle specifiche.

## OUTPUT EXCEL

### Struttura fogli
- 12 fogli mensili (GENNAIO, FEBBRAIO, ..., DICEMBRE)
- Ogni foglio contiene:
  - Riga 1: Titolo mese/anno
  - Riga 2: Intestazioni (Turno, 1-31, Turno, gg., Progr.)
  - Righe 3+: Dati turni

### Formattazione
- **Celle G**: Sfondo verde chiaro (giornaliero)
- **Celle F***: Sfondo giallo (ferie)
- **Weekend**: Sfondo grigio chiaro
- **Bordi**: Tutti i bordi visibili

## PERSONALIZZAZIONE

### Modificare anno target
Nel file `turni_generator.py`, sezione `main()`:

```python
anno_target = 2026  # Cambia l'anno qui
```

### Modificare percorsi file
```python
file_template = r"percorso\al\template\2025.xls"
output_dir = r"percorso\output"
```

### Aggiornare festività
Modificare la lista `FESTIVITA_2025` per l'anno desiderato:

```python
FESTIVITA_2025 = [
    (1, 1),    # Capodanno
    (1, 6),    # Epifania
    # ... aggiungi altre festività
]
```

### Modificare rotazione ferie
Per generare gli anni successivi, aggiornare `FERIE_2026`, `FERIE_2027`, etc.:

```python
# Rotazione 2026 (applica 1→3→5→2→4 a FERIE_2025)
FERIE_2026 = {
    1: [42, 48],  # Era periodo 4 nel 2025
    2: [41, 47],  # Era periodo 5 nel 2025
    # ...
}
```

## REPORT VERIFICA

Il file `report_verifica.txt` contiene:
- Numero totale giorni verificati OK
- Numero anomalie rilevate
- Lista dettagliata di tutte le anomalie con conteggi A/B/C/G

Esempio:
```
======================================================================
REPORT VERIFICA TURNI 2025
======================================================================

Giorni verificati OK: 300
Giorni con anomalie: 65

ANOMALIE RILEVATE:
----------------------------------------------------------------------
Data: 20/06/2025 - A=2, B=3, C=2, G=0
Data: 21/06/2025 - A=2, B=0, C=2, G=0
...
```

## LIMITAZIONI E NOTE

1. **Verifica manuale obbligatoria**: Il programma genera una bozza che deve essere verificata manualmente
2. **Periodi transizione**: Le settimane 16-23 giugno e 14-22 settembre hanno pattern irregolari
3. **Turno 46 speciale**: Comportamento diverso tra periodo normale (solo G) e periodo estivo (A/B/C)
4. **Leap year**: Il programma gestisce automaticamente gli anni bisestili per il calcolo dell'offset
5. **Template dependency**: Richiede il file template dell'anno precedente per calcolare la continuità

## TROUBLESHOOTING

### Errore: "File template non trovato"
Verificare che il percorso del file template sia corretto e che il file esista.

### Errore: "File is not a zip file"
Il file .xls potrebbe essere corrotto. Provare a:
1. Aprire il file con Excel
2. Salvarlo come .xlsx
3. Aggiornare il percorso nel codice

### Molte anomalie nel periodo estivo
Questo è normale. Le anomalie durante giugno-settembre sono dovute ai periodi ferie sovrapposti e richiedono verifica manuale secondo le specifiche.

### Celle G in weekend/festivi
Verificare la lista `FESTIVITA_2025` e la logica in `_codice_periodo_normale()`.

## SUPPORTO

Per problemi o domande:
1. Verificare che tutti i requisiti siano installati
2. Controllare i percorsi dei file
3. Consultare il report di verifica
4. Verificare manualmente i periodi con anomalie

## VERSIONE

- **Versione**: 1.0
- **Data**: 2025
- **Autore**: Generato secondo specifiche raffineria
- **Python**: >= 3.7

## LICENSE

Uso interno raffineria.
