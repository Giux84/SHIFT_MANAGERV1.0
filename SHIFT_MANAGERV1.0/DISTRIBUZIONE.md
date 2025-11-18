# 📦 Guida Distribuzione ShiftManager.exe

## ✅ L'exe è COMPLETAMENTE STANDALONE

**NON serve installare NULLA!**

- ✅ Python incluso nell'exe
- ✅ Tutte le librerie incluse (pandas, openpyxl, customtkinter, ecc.)
- ✅ Funziona su qualsiasi PC Windows 10/11
- ✅ Dimensione: ~50 MB

---

## 🚀 Come Distribuire su Chiavetta USB

### **Step 1: Copia l'exe**

1. Vai nella cartella `dist/`
2. Copia il file **`ShiftManager.exe`**
3. Incolla su chiavetta USB

### **Step 2: Utilizzo su Altro PC**

1. Inserisci la chiavetta
2. **Doppio click** su `ShiftManager.exe`
3. **Fatto!** L'applicazione si avvia

---

## ⚠️ Possibili Avvisi (NORMALI)

### Windows Defender SmartScreen

**Al primo avvio potrebbe apparire:**
```
Windows ha protetto il PC
L'esecuzione di questa app potrebbe mettere a rischio il PC
```

**Soluzione:**
1. Clicca **"Maggiori informazioni"**
2. Clicca **"Esegui comunque"**

**Motivo:** L'exe non è firmato digitalmente (richiede certificato Microsoft da €300/anno)

### Antivirus

Alcuni antivirus potrebbero bloccare l'exe perché "non riconosciuto".

**Soluzione:**
- Aggiungi `ShiftManager.exe` alle **eccezioni** dell'antivirus
- Oppure disattiva temporaneamente l'antivirus durante il primo avvio

---

## 💡 Suggerimenti

### **Primo Avvio Lento**

Il primo avvio può richiedere 5-10 secondi perché:
- L'exe estrae le librerie in una cartella temporanea
- Windows scansiona il file

I successivi avvii saranno più veloci.

### **Dove Mettere l'exe**

**Opzione 1 - Chiavetta USB (Portable):**
- ✅ Funziona senza installazione
- ✅ Può essere usato su più PC
- ❌ Primo avvio sempre lento

**Opzione 2 - Cartella Locale (Consigliato):**
- Copia l'exe in: `C:\Programmi\ShiftManager\` o `Desktop`
- ✅ Avvio più veloce
- ✅ Nessuna estrazione ad ogni avvio

### **File di Configurazione**

L'app crea un file **`shift_manager_config.json`** nella stessa cartella dell'exe per salvare le preferenze.

Se usi la chiavetta, il file verrà creato lì e le impostazioni saranno portabili.

---

## 🔧 Risoluzione Problemi

### "Impossibile avviare l'applicazione"

**Causa:** Windows 10/11 non aggiornato

**Soluzione:**
- Aggiorna Windows
- Oppure installa **Visual C++ Redistributable** (scarica da Microsoft)

### "Errore mancano DLL"

**Causa:** Sistema Windows danneggiato

**Soluzione:**
- Esegui `sfc /scannow` nel Prompt dei Comandi (come admin)

### L'exe viene eliminato automaticamente

**Causa:** Antivirus troppo aggressivo

**Soluzione:**
1. Ripristina il file dal Quarantena dell'antivirus
2. Aggiungi eccezione permanente per `ShiftManager.exe`

---

## 📊 Requisiti Sistema

### **Minimi:**
- Windows 10 (64-bit) o superiore
- 2 GB RAM
- 100 MB spazio disco (per file temporanei)

### **Consigliati:**
- Windows 11 (64-bit)
- 4 GB RAM
- SSD per caricamento rapido

---

## 🔐 Sicurezza

### **L'exe è Sicuro?**

✅ **SÌ!** L'exe contiene:
- Python ufficiale
- Librerie open-source verificate
- Il tuo codice sorgente

**NON contiene:**
- Virus
- Malware
- Spyware
- Connessioni internet non autorizzate

### **Verifica Hash (Opzionale)**

Per verificare che l'exe non sia stato modificato:

```powershell
Get-FileHash dist\ShiftManager.exe -Algorithm SHA256
```

Salva questo hash e confrontalo quando ridistribuisci.

---

## 📝 Note Finali

1. **L'exe funziona offline** - non serve connessione internet
2. **I file Excel** generati sono standard - possono essere aperti con qualsiasi Excel
3. **Le configurazioni** sono salvate localmente nel file JSON
4. **Non c'è telemetria** - nessun dato viene inviato online

---

## 🆘 Supporto

Per problemi o domande:
- Leggi questa guida
- Controlla il file `README_GUI.md` per l'uso dell'applicazione
- Contatta l'amministratore del sistema

---

**Versione:** 1.0 (Novembre 2025)
**Build:** PyInstaller 6.16.0
**Compatibilità:** Windows 10/11 (64-bit)
