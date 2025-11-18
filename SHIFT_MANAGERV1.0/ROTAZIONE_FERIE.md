# GUIDA ROTAZIONE FERIE QUINQUENNALE

## SCHEMA ROTAZIONE

La rotazione delle ferie segue uno schema quinquennale (5 anni) con la formula:
**1 → 3 → 5 → 2 → 4 → (ritorna a 1)**

## PERIODI FERIE (Date Fisse)

| Periodo | Date             | Durata      |
|---------|------------------|-------------|
| 1       | 20/06 - 08/07    | ~18 giorni  |
| 2       | 08/07 - 25/07    | ~17 giorni  |
| 3       | 25/07 - 11/08    | ~17 giorni  |
| 4       | 11/08 - 28/08    | ~17 giorni  |
| 5       | 28/08 - 14/09    | ~17 giorni  |
| 6       | 14/09 - 30/09    | ~16 giorni  |

**Nota**: Periodo 6 è riservato sempre al turno 46 (speciale)

## ASSEGNAZIONE PER ANNO

### Anno 2024 (Base di partenza)
| Periodo | Date             | Turni    |
|---------|------------------|----------|
| 1       | 20/06 - 08/07    | 45, 51   |
| 2       | 08/07 - 25/07    | 42, 48   |
| 3       | 25/07 - 11/08    | 41, 47   |
| 4       | 11/08 - 28/08    | 44, 50   |
| 5       | 28/08 - 14/09    | 43, 49   |
| 6       | 14/09 - 30/09    | **46**   |

### Anno 2025 (Rotazione applicata)
| Periodo | Date             | Turni    | Provenienza 2024 |
|---------|------------------|----------|------------------|
| 1       | 20/06 - 08/07    | 44, 50   | Era periodo 4    |
| 2       | 08/07 - 25/07    | 43, 49   | Era periodo 5    |
| 3       | 25/07 - 11/08    | 45, 51   | Era periodo 1    |
| 4       | 11/08 - 28/08    | 42, 48   | Era periodo 2    |
| 5       | 28/08 - 14/09    | 41, 47   | Era periodo 3    |
| 6       | 14/09 - 30/09    | **46**   | Sempre fisso     |

### Anno 2026 (Rotazione successiva)
| Periodo | Date             | Turni    | Provenienza 2025 |
|---------|------------------|----------|------------------|
| 1       | 20/06 - 08/07    | 42, 48   | Era periodo 4    |
| 2       | 08/07 - 25/07    | 41, 47   | Era periodo 5    |
| 3       | 25/07 - 11/08    | 44, 50   | Era periodo 1    |
| 4       | 11/08 - 28/08    | 43, 49   | Era periodo 2    |
| 5       | 28/08 - 14/09    | 45, 51   | Era periodo 3    |
| 6       | 14/09 - 30/09    | **46**   | Sempre fisso     |

### Anno 2027 (Rotazione successiva)
| Periodo | Date             | Turni    | Provenienza 2026 |
|---------|------------------|----------|------------------|
| 1       | 20/06 - 08/07    | 43, 49   | Era periodo 4    |
| 2       | 08/07 - 25/07    | 45, 51   | Era periodo 5    |
| 3       | 25/07 - 11/08    | 42, 48   | Era periodo 1    |
| 4       | 11/08 - 28/08    | 41, 47   | Era periodo 2    |
| 5       | 28/08 - 14/09    | 44, 50   | Era periodo 3    |
| 6       | 14/09 - 30/09    | **46**   | Sempre fisso     |

### Anno 2028 (Rotazione successiva)
| Periodo | Date             | Turni    | Provenienza 2027 |
|---------|------------------|----------|------------------|
| 1       | 20/06 - 08/07    | 41, 47   | Era periodo 4    |
| 2       | 08/07 - 25/07    | 44, 50   | Era periodo 5    |
| 3       | 25/07 - 11/08    | 43, 49   | Era periodo 1    |
| 4       | 11/08 - 28/08    | 45, 51   | Era periodo 2    |
| 5       | 28/08 - 14/09    | 42, 48   | Era periodo 3    |
| 6       | 14/09 - 30/09    | **46**   | Sempre fisso     |

### Anno 2029 (Ritorno al ciclo)
| Periodo | Date             | Turni    | Provenienza 2028 |
|---------|------------------|----------|------------------|
| 1       | 20/06 - 08/07    | 45, 51   | Era periodo 4    |
| 2       | 08/07 - 25/07    | 42, 48   | Era periodo 5    |
| 3       | 25/07 - 11/08    | 41, 47   | Era periodo 1    |
| 4       | 11/08 - 28/08    | 44, 50   | Era periodo 2    |
| 5       | 28/08 - 14/09    | 43, 49   | Era periodo 3    |
| 6       | 14/09 - 30/09    | **46**   | Sempre fisso     |

**Nota**: 2029 torna alla configurazione 2024! Il ciclo si ripete ogni 5 anni.

## LOGICA ROTAZIONE

### Formula Matematica
```
nuovo_periodo = ((vecchio_periodo - 1 + spostamento) % 5) + 1
```

Dove `spostamento` per la rotazione 1→3→5→2→4 è equivalente a +2 posizioni (modulo 5).

### Mappatura Esplicita
```
Vecchio → Nuovo
1       → 3
2       → 4
3       → 5
4       → 1
5       → 2
6       → 6  (invariante)
```

## AGGIORNAMENTO CODICE

Per generare l'anno successivo, aggiornare il dizionario `FERIE_XXXX` nel file `turni_generator.py`:

```python
# Esempio per anno 2026
FERIE_2026 = {
    1: [42, 48],  # Applica rotazione a FERIE_2025[4]
    2: [41, 47],  # Applica rotazione a FERIE_2025[5]
    3: [44, 50],  # Applica rotazione a FERIE_2025[1]
    4: [43, 49],  # Applica rotazione a FERIE_2025[2]
    5: [45, 51],  # Applica rotazione a FERIE_2025[3]
    6: [46]       # Sempre fisso
}
```

## COPPIE GEMELLE

Le coppie gemelle vanno SEMPRE in ferie nello stesso periodo:
- **(41, 47)** - Coppia gemella
- **(42, 48)** - Coppia gemella
- **(43, 49)** - Coppia gemella
- **(44, 50)** - Coppia gemella
- **(45, 51)** - Coppia gemella
- **46** - Turno singolo speciale

**Importante**: Il turno gemello inizia le ferie **1 giorno prima** del turno principale per garantire la continuità operativa.

## DURATA FERIE

- **12 giorni lavorativi** per ogni turno
- Non contano i giorni di riposo (-)
- Il periodo di ferie può estendersi oltre 12 giorni di calendario per includere i riposi

## CODICI FERIE

- **FA**: Ferie durante turno mattina (A)
- **FB**: Ferie durante turno pomeriggio (B)
- **FC**: Ferie durante turno notte (C)
- **FG**: Ferie durante turno giornaliero (G) - solo turno 46

## VERIFICA ROTAZIONE

Per verificare che la rotazione sia corretta:

1. Ogni coppia gemella deve essere assegnata a un periodo unico
2. Nessun turno (eccetto 46) deve avere lo stesso periodo per due anni consecutivi
3. Il turno 46 deve sempre avere il periodo 6
4. Dopo 5 anni, ogni coppia torna al periodo di partenza

## ESEMPIO TRACCIAMENTO TURNO 41-47

| Anno | Periodo | Settimane      |
|------|---------|----------------|
| 2024 | 3       | 25/07 - 11/08  |
| 2025 | 5       | 28/08 - 14/09  |
| 2026 | 2       | 08/07 - 25/07  |
| 2027 | 4       | 11/08 - 28/08  |
| 2028 | 1       | 20/06 - 08/07  |
| 2029 | 3       | 25/07 - 11/08  | ← Ritorna al periodo iniziale

## TOOL DI SUPPORTO

Per facilitare il calcolo, usare questa tabella di lookup:

| Coppia | 2024 | 2025 | 2026 | 2027 | 2028 | 2029 |
|--------|------|------|------|------|------|------|
| 41-47  | P3   | P5   | P2   | P4   | P1   | P3   |
| 42-48  | P2   | P4   | P1   | P3   | P5   | P2   |
| 43-49  | P5   | P2   | P4   | P1   | P3   | P5   |
| 44-50  | P4   | P1   | P3   | P5   | P2   | P4   |
| 45-51  | P1   | P3   | P5   | P2   | P4   | P1   |
| 46     | P6   | P6   | P6   | P6   | P6   | P6   |

## MODIFICHE FUTURE

Se in futuro si dovesse modificare:
- **Date periodi**: Aggiornare `PERIODI_FERIE`
- **Schema rotazione**: Aggiornare la logica in `FERIE_XXXX`
- **Numero periodi**: Modificare anche la logica di verifica

---

**Data ultima revisione**: 2025
**Versione documento**: 1.0
