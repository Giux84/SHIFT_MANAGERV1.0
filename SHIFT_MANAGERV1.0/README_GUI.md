# 🔄 Shift Manager - Interfaccia Grafica Utente

## Introduzione

**Shift Manager** è un'applicazione desktop con interfaccia grafica moderna per generare automaticamente i turni annuali della raffineria.

L'applicazione legge il file Excel dell'anno precedente e genera automaticamente il calendario turni per l'anno successivo, mantenendo la continuità del ciclo e applicando la rotazione quinquennale delle ferie.

## 🚀 Avvio Rapido

### Modalità 1: Eseguibile Standalone (.exe)

**La modalità più semplice - non richiede installazione di Python!**

1. Vai nella cartella `dist/`
2. Esegui `ShiftManager.exe`
3. L'applicazione si aprirà immediatamente

**Nota**: La prima volta potrebbe richiedere qualche secondo in più per caricare tutte le librerie.

### Modalità 2: Esecuzione da Python

Se preferisci eseguire il codice sorgente:

```bash
python shift_manager_gui.py
```

**Requisiti**:
- Python 3.8 o superiore
- Dipendenze: `pip install -r requirements.txt`

## 📋 Guida all'Uso

### 1. Schermata Principale

#### 📁 **File Anno Precedente**
- Clicca su **"Sfoglia"** per selezionare il file Excel dell'anno precedente (`.xls` o `.xlsx`)
- Esempio: `TURNO_COMPLETO_2024.xls`
- Il file viene validato automaticamente

#### 📅 **Anno da Generare**
- Inserisci l'anno per cui vuoi generare i turni
- Esempio: `2025`, `2026`, ecc.
- L'anno deve essere compreso tra 2020 e 10 anni nel futuro

#### 💾 **Cartella Output**
- Clicca su **"Sfoglia"** per selezionare dove salvare il file generato
- Di default usa la cartella corrente dell'applicazione
- Il file sarà salvato come `TURNO_COMPLETO_[ANNO].xlsx`

#### ▶️ **GENERA TURNI**
- Il grande bottone verde genera i turni
- Durante la generazione vedrai:
  - Progress bar che mostra l'avanzamento
  - Log in tempo reale delle operazioni
- Al termine viene chiesto se vuoi aprire la cartella del file generato

### 2. ⚙️ Opzioni Avanzate

Clicca il bottone **"Opzioni Avanzate"** per personalizzare:

#### Tab "Regole Turni"
- **Massimo giorni lavorati/anno**: Default 231
- **Giorni ciclo periodo normale**: Default 10 giorni
- **Giorni ciclo periodo estivo**: Default 9 giorni
- **Periodo estivo**: Date inizio/fine (default: 20/6 - 13/9)

#### Tab "Colori"
- Personalizza i colori delle celle Excel:
  - 🟥 **Domeniche e Festivi**: Default rosso (#FF0000)
  - 🟦 **Sabati**: Default azzurro (#4472C4)
  - 🟨 **Ferie**: Default giallo chiaro (#FFFFCC)
  - 🟩 **Giornate G** (no T46): Default verde chiaro (#C6EFCE)
- Usa il **color picker** visuale cliccando "Seleziona"

#### Tab "Festività"
- Le festività italiane standard sono automatiche
- Puoi aggiungere festività extra personalizzate
- Formato: una per riga, `GG/MM` (es. `15/08`)

#### Tab "Avanzate"
- ✅ **Verifica anomalie copertura 2-2-2**
- ✅ **Genera report dettagliato**
- ✅ **Mostra progressivo giorni per turno**
- Personalizza font Excel (nome, dimensione titolo, dimensione normale)

**Nota**: Le configurazioni vengono salvate automaticamente in `shift_manager_config.json` e ricaricate al prossimo avvio.

### 3. 👁️ Anteprima

Dopo aver generato i turni:

- Clicca **"Anteprima"** per visualizzare il calendario
- Naviga tra i mesi con i bottoni **◀ Mese Precedente** / **Mese Successivo ▶**
- Visualizza:
  - Tabella completa di tutti i turni per il mese
  - Colori esattamente come nell'Excel finale
  - Totali e progressivi per turno
  - Legenda colori

### 4. ℹ️ Aiuto

Clicca **"Aiuto"** in qualsiasi momento per vedere la guida rapida integrata.

## 📊 Output Generato

### File Excel: `TURNO_COMPLETO_[ANNO].xlsx`

Il file generato contiene:

- **12 fogli** (uno per ogni mese)
- **Formattazione colori automatica**:
  - Rosso: domeniche e festività
  - Azzurro: sabati
  - Giallo chiaro: ferie (FA, FB, FC, FG)
  - Verde chiaro: giornate G (eccetto turno 46)
- **Colonne finali**:
  - **Turno**: Numero turno
  - **gg.**: Giorni lavorativi del mese
  - **Progr.**: Progressivo giorni per turno (max 231)
- **Righe bianche** separatrici prima e dopo il turno 46

### File Report: `report_verifica_[ANNO].txt`

Report testuale con:
- Statistiche verifica copertura 2-2-2
- Lista anomalie rilevate (se presenti)
- Data e orari delle anomalie

## 🔧 Risoluzione Problemi

### L'exe non si avvia
- **Antivirus**: Alcuni antivirus possono bloccare eseguibili non firmati. Aggiungi `ShiftManager.exe` alle eccezioni.
- **Windows Defender**: Potrebbe mostrare un warning "App non riconosciuta". Clicca "Maggiori informazioni" → "Esegui comunque"

### Errore "File template non trovato"
- Verifica che il file Excel esista nel percorso selezionato
- Controlla che l'estensione sia `.xls` o `.xlsx`

### Errore durante la generazione
- Controlla il **Log** nella schermata principale per dettagli
- Verifica che il file template sia valido e leggibile
- Controlla che Excel non abbia il file aperto (chiudilo prima di rigenerare)

### Le modifiche alle opzioni non vengono salvate
- Assicurati di cliccare **"Salva"** nel dialog Opzioni Avanzate
- Verifica i permessi di scrittura nella cartella dell'applicazione
- Il file di configurazione è `shift_manager_config.json` nella cartella dell'app

### L'anteprima non si apre
- Devi prima generare i turni con il bottone "GENERA TURNI"
- L'anteprima funziona solo se il file Excel esiste nella cartella output

## 🎨 Personalizzazione

### Modificare i colori di default
1. Apri **Opzioni Avanzate**
2. Vai al tab **"Colori"**
3. Clicca **"Seleziona"** per usare il color picker
4. Oppure inserisci manualmente il codice HEX (senza #)
5. Clicca **"Salva"**

### Salvare preset personalizzati
Le configurazioni vengono salvate automaticamente in `shift_manager_config.json`.

Per usare preset diversi:
- Copia il file `shift_manager_config.json` e rinominalo (es. `preset_standard.json`)
- Quando vuoi cambiare preset, sostituisci il file `shift_manager_config.json` con quello desiderato

### Ripristinare impostazioni di default
1. Apri **Opzioni Avanzate**
2. Clicca **"Ripristina Default"**
3. Conferma l'operazione

## 📝 Note Tecniche

### Funzionalità Automatiche

- ✅ **Continuità ciclo**: Calcola automaticamente l'offset tra anni
- ✅ **Rotazione ferie quinquennale**: 1→3→5→2→4
- ✅ **Verifica copertura 2-2-2**: Controlla che ogni giorno abbia 2 turni A, 2 B, 2 C
- ✅ **Pattern normale**: Ciclo 10 giorni (Gen-Mag, Ott-Dic)
- ✅ **Pattern estivo**: Ciclo 9 giorni (20 Giu - 13 Set)
- ✅ **Calcolo automatico Pasqua**: Aggiornato ogni anno
- ✅ **Progressivo per turno**: Ogni turno ha il suo progressivo (max 231 giorni)

### Limiti Noti

- Il file template deve essere un Excel valido dell'anno precedente
- L'anno da generare deve essere l'anno successivo a quello del template
- I pattern dei turni sono fissi (non modificabili dalla GUI)
- La rotazione ferie è fissa (1→3→5→2→4)

## 🆘 Supporto

Per problemi o domande:
- Controlla il file `report_verifica_[ANNO].txt` per anomalie
- Consulta il log nella schermata principale
- Contatta l'amministratore del sistema

## 📄 Licenza

© 2025 Shift Manager v1.0
Generato con [Claude Code](https://claude.com/claude-code)

---

**Ultima versione**: 1.0
**Data**: Novembre 2025
**Compatibilità**: Windows 10/11, Python 3.8+
