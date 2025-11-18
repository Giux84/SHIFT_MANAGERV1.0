"""
CONFIGURAZIONI APPLICAZIONE SHIFT MANAGER
Gestisce le configurazioni, preset e costanti dell'applicazione
"""

import json
import os
from typing import Dict, Any
from datetime import date

# ============================================================================
# CONFIGURAZIONI DEFAULT
# ============================================================================

DEFAULT_CONFIG = {
    # Regole turni
    "max_giorni_lavorati": 231,  # Massimo giorni lavorabili in un anno (365 - 134 riposi)
    "pattern_normale_giorni": 10,
    "pattern_estivo_giorni": 9,
    "periodo_estivo_inizio": (6, 20),  # 20 giugno
    "periodo_estivo_fine": (9, 13),    # 13 settembre

    # Colori formattazione (RGB hex)
    "colore_domeniche_festivi": "FF0000",  # Rosso
    "colore_sabati": "4472C4",             # Azzurro
    "colore_ferie": "FFFFCC",              # Giallo chiaro
    "colore_giornate_g": "C6EFCE",         # Verde chiaro

    # Font
    "font_nome": "Calibri",
    "font_dimensione_titolo": 14,
    "font_dimensione_normale": 11,

    # Opzioni avanzate
    "verifica_anomalie": True,
    "genera_report": True,
    "mostra_progressivo_per_turno": True,

    # Festività personalizzate (oltre a quelle standard)
    "festivita_extra": []
}

# ============================================================================
# FESTIVITÀ ITALIANE BASE (template per ogni anno)
# ============================================================================

def get_festivita_anno(anno: int) -> list:
    """
    Restituisce le festività per un anno specifico.
    Nota: Pasqua varia ogni anno, va calcolata separatamente
    """
    return [
        (1, 1),    # Capodanno
        (1, 6),    # Epifania
        # Pasqua e Lunedì dell'Angelo: da calcolare
        (4, 25),   # Liberazione
        (5, 1),    # Festa del Lavoro
        (6, 2),    # Repubblica
        (8, 15),   # Ferragosto
        (11, 1),   # Tutti i Santi
        (12, 8),   # Immacolata
        (12, 25),  # Natale
        (12, 26),  # Santo Stefano
    ]

def calcola_pasqua(anno: int) -> tuple:
    """
    Calcola la data della Pasqua per un anno usando l'algoritmo di Meeus/Jones/Butcher
    Restituisce (mese, giorno)
    """
    a = anno % 19
    b = anno // 100
    c = anno % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mese = (h + l - 7 * m + 114) // 31
    giorno = ((h + l - 7 * m + 114) % 31) + 1
    return (mese, giorno)

# ============================================================================
# GESTIONE CONFIGURAZIONI
# ============================================================================

class ConfigManager:
    """Gestisce il caricamento e salvataggio delle configurazioni"""

    def __init__(self, config_file: str = "shift_manager_config.json"):
        self.config_file = config_file
        self.config = DEFAULT_CONFIG.copy()
        self.load()

    def load(self) -> bool:
        """Carica la configurazione da file"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                    # Aggiorna solo le chiavi presenti, mantieni i default per le nuove
                    self.config.update(loaded_config)
                return True
            except Exception as e:
                print(f"Errore caricamento configurazione: {e}")
                return False
        return False

    def save(self) -> bool:
        """Salva la configurazione su file"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Errore salvataggio configurazione: {e}")
            return False

    def get(self, key: str, default=None) -> Any:
        """Ottiene un valore di configurazione"""
        return self.config.get(key, default)

    def set(self, key: str, value: Any):
        """Imposta un valore di configurazione"""
        self.config[key] = value

    def reset_to_default(self):
        """Ripristina le configurazioni di default"""
        self.config = DEFAULT_CONFIG.copy()

    def export_preset(self, preset_name: str, preset_file: str) -> bool:
        """Esporta la configurazione corrente come preset"""
        try:
            preset_data = {
                "name": preset_name,
                "config": self.config.copy()
            }
            with open(preset_file, 'w', encoding='utf-8') as f:
                json.dump(preset_data, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Errore esportazione preset: {e}")
            return False

    def import_preset(self, preset_file: str) -> bool:
        """Importa un preset"""
        try:
            with open(preset_file, 'r', encoding='utf-8') as f:
                preset_data = json.load(f)
                if "config" in preset_data:
                    self.config = preset_data["config"]
                    return True
            return False
        except Exception as e:
            print(f"Errore importazione preset: {e}")
            return False

# ============================================================================
# VALIDATORI
# ============================================================================

def valida_anno(anno: int) -> tuple[bool, str]:
    """
    Valida che l'anno sia ragionevole
    Restituisce (valido, messaggio_errore)
    """
    anno_corrente = date.today().year
    if anno < 2020:
        return False, "L'anno deve essere >= 2020"
    if anno > anno_corrente + 10:
        return False, f"L'anno non può essere oltre {anno_corrente + 10}"
    return True, ""

def valida_file_excel(file_path: str) -> tuple[bool, str]:
    """
    Valida che il file Excel esista e sia leggibile
    Restituisce (valido, messaggio_errore)
    """
    if not file_path:
        return False, "Seleziona un file"

    if not os.path.exists(file_path):
        return False, "Il file non esiste"

    ext = os.path.splitext(file_path)[1].lower()
    if ext not in ['.xls', '.xlsx']:
        return False, "Il file deve essere .xls o .xlsx"

    return True, ""

def valida_percorso_output(percorso: str) -> tuple[bool, str]:
    """
    Valida che il percorso di output sia valido
    Restituisce (valido, messaggio_errore)
    """
    if not percorso:
        return False, "Seleziona una cartella di output"

    # Se è una cartella, verifica che esista
    if os.path.isdir(percorso):
        return True, ""

    # Se è un file, verifica che la cartella parent esista
    parent_dir = os.path.dirname(percorso)
    if not os.path.exists(parent_dir):
        return False, "La cartella di destinazione non esiste"

    return True, ""
