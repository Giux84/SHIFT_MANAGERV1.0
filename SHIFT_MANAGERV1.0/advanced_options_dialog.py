"""
ADVANCED OPTIONS DIALOG
Dialog per configurare opzioni avanzate del generatore turni
"""

import customtkinter as ctk
from tkinter import colorchooser, messagebox
from typing import Dict, Any
from config import ConfigManager, calcola_pasqua, get_festivita_anno

class AdvancedOptionsDialog(ctk.CTkToplevel):
    """Dialog opzioni avanzate con tabs"""

    def __init__(self, parent, config_manager: ConfigManager):
        super().__init__(parent)

        self.config_manager = config_manager
        self.modified = False

        # Configurazione finestra
        self.title("⚙️ Opzioni Avanzate")
        self.geometry("800x750")
        self.minsize(750, 700)

        # Crea UI
        self.create_widgets()

        # Carica valori correnti
        self.load_config()

        # Centra finestra
        self.center_window()

    def center_window(self):
        """Centra la finestra sullo schermo"""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')

    def create_widgets(self):
        """Crea tutti i widget"""

        # Header
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=(20, 10))

        title_label = ctk.CTkLabel(
            header_frame,
            text="⚙️ Opzioni Avanzate",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        title_label.pack()

        # TabView
        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(fill="both", expand=True, padx=20, pady=10)

        # Crea tabs
        self.tabview.add("Regole Turni")
        self.tabview.add("Colori")
        self.tabview.add("Festività")
        self.tabview.add("Avanzate")

        # Popola tabs
        self.create_rules_tab()
        self.create_colors_tab()
        self.create_holidays_tab()
        self.create_advanced_tab()

        # Footer buttons
        footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        footer_frame.pack(fill="x", padx=20, pady=(0, 20))

        reset_btn = ctk.CTkButton(
            footer_frame,
            text="Ripristina Default",
            command=self.reset_to_default,
            width=150,
            height=40,
            fg_color="gray",
            hover_color="darkgray"
        )
        reset_btn.pack(side="left", padx=(0, 10))

        cancel_btn = ctk.CTkButton(
            footer_frame,
            text="Annulla",
            command=self.cancel,
            width=100,
            height=40
        )
        cancel_btn.pack(side="right")

        save_btn = ctk.CTkButton(
            footer_frame,
            text="Salva",
            command=self.save,
            width=100,
            height=40,
            fg_color="#2B7A0B",
            hover_color="#1F5A08"
        )
        save_btn.pack(side="right", padx=(0, 10))

    # ========================================================================
    # TAB: REGOLE TURNI
    # ========================================================================

    def create_rules_tab(self):
        """Crea tab regole turni"""
        tab = self.tabview.tab("Regole Turni")

        # Scrollable frame
        scroll_frame = ctk.CTkScrollableFrame(tab)
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Max giorni lavorati
        self.create_labeled_entry(
            scroll_frame,
            "Massimo giorni lavorati/anno:",
            "max_giorni_lavorati",
            "231"
        )

        # Pattern normale
        self.create_labeled_entry(
            scroll_frame,
            "Giorni ciclo periodo normale:",
            "pattern_normale_giorni",
            "10"
        )

        # Pattern estivo
        self.create_labeled_entry(
            scroll_frame,
            "Giorni ciclo periodo estivo:",
            "pattern_estivo_giorni",
            "9"
        )

        # Periodo estivo inizio
        periodo_frame = ctk.CTkFrame(scroll_frame)
        periodo_frame.pack(fill="x", pady=10)

        label = ctk.CTkLabel(
            periodo_frame,
            text="Periodo estivo - Inizio:",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        label.pack(anchor="w", padx=10, pady=(10, 5))

        date_frame = ctk.CTkFrame(periodo_frame, fg_color="transparent")
        date_frame.pack(fill="x", padx=10, pady=(0, 10))

        ctk.CTkLabel(date_frame, text="Giorno:").pack(side="left", padx=(0, 5))
        self.estivo_inizio_giorno = ctk.CTkEntry(date_frame, width=60)
        self.estivo_inizio_giorno.pack(side="left", padx=(0, 15))

        ctk.CTkLabel(date_frame, text="Mese:").pack(side="left", padx=(0, 5))
        self.estivo_inizio_mese = ctk.CTkEntry(date_frame, width=60)
        self.estivo_inizio_mese.pack(side="left")

        # Periodo estivo fine
        periodo_frame2 = ctk.CTkFrame(scroll_frame)
        periodo_frame2.pack(fill="x", pady=10)

        label2 = ctk.CTkLabel(
            periodo_frame2,
            text="Periodo estivo - Fine:",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        label2.pack(anchor="w", padx=10, pady=(10, 5))

        date_frame2 = ctk.CTkFrame(periodo_frame2, fg_color="transparent")
        date_frame2.pack(fill="x", padx=10, pady=(0, 10))

        ctk.CTkLabel(date_frame2, text="Giorno:").pack(side="left", padx=(0, 5))
        self.estivo_fine_giorno = ctk.CTkEntry(date_frame2, width=60)
        self.estivo_fine_giorno.pack(side="left", padx=(0, 15))

        ctk.CTkLabel(date_frame2, text="Mese:").pack(side="left", padx=(0, 5))
        self.estivo_fine_mese = ctk.CTkEntry(date_frame2, width=60)
        self.estivo_fine_mese.pack(side="left")

    # ========================================================================
    # TAB: COLORI
    # ========================================================================

    def create_colors_tab(self):
        """Crea tab colori"""
        tab = self.tabview.tab("Colori")

        scroll_frame = ctk.CTkScrollableFrame(tab)
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Dizionario per memorizzare i colori
        self.color_vars = {}

        # Colori da configurare
        colors = [
            ("Domeniche e Festivi", "colore_domeniche_festivi", "FF0000"),
            ("Sabati", "colore_sabati", "4472C4"),
            ("Ferie", "colore_ferie", "FFFFCC"),
            ("Giornate G (non T46)", "colore_giornate_g", "C6EFCE")
        ]

        for label_text, key, default in colors:
            self.create_color_picker(scroll_frame, label_text, key, default)

        # Info
        info_frame = ctk.CTkFrame(scroll_frame)
        info_frame.pack(fill="x", pady=20)

        info_label = ctk.CTkLabel(
            info_frame,
            text="ℹ️ I colori sono in formato esadecimale RGB (es. FF0000 = rosso)\nClicca 'Seleziona' per usare il color picker visuale.",
            font=ctk.CTkFont(size=11),
            text_color="gray",
            justify="left"
        )
        info_label.pack(padx=15, pady=15)

    def create_color_picker(self, parent, label_text: str, key: str, default: str):
        """Crea un color picker con preview"""
        frame = ctk.CTkFrame(parent)
        frame.pack(fill="x", pady=10)

        # Label
        label = ctk.CTkLabel(
            frame,
            text=label_text + ":",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        label.pack(anchor="w", padx=10, pady=(10, 5))

        # Input frame
        input_frame = ctk.CTkFrame(frame, fg_color="transparent")
        input_frame.pack(fill="x", padx=10, pady=(0, 10))

        # Entry
        entry = ctk.CTkEntry(input_frame, width=100, placeholder_text=default)
        entry.pack(side="left", padx=(0, 10))
        self.color_vars[key] = entry

        # Preview
        preview = ctk.CTkLabel(
            input_frame,
            text="     ",
            width=50,
            height=30,
            fg_color=f"#{default}",
            corner_radius=5
        )
        preview.pack(side="left", padx=(0, 10))

        # Button
        def pick_color():
            # Converti hex in rgb per il color picker
            current_color = entry.get() or default
            if not current_color.startswith('#'):
                current_color = f"#{current_color}"

            color_code = colorchooser.askcolor(
                title=f"Seleziona colore per {label_text}",
                color=current_color
            )

            if color_code[1]:  # Se l'utente ha selezionato un colore
                # Rimuovi # dall'inizio
                hex_color = color_code[1][1:]
                entry.delete(0, "end")
                entry.insert(0, hex_color)
                preview.configure(fg_color=f"#{hex_color}")

        pick_btn = ctk.CTkButton(
            input_frame,
            text="Seleziona",
            command=pick_color,
            width=80
        )
        pick_btn.pack(side="left")

    # ========================================================================
    # TAB: FESTIVITÀ
    # ========================================================================

    def create_holidays_tab(self):
        """Crea tab festività"""
        tab = self.tabview.tab("Festività")

        # Info
        info_frame = ctk.CTkFrame(tab, fg_color="transparent")
        info_frame.pack(fill="x", padx=10, pady=10)

        info_label = ctk.CTkLabel(
            info_frame,
            text="Festività italiane standard (automatiche):\n"
                 "Capodanno, Epifania, Pasqua, Lunedì dell'Angelo,\n"
                 "Liberazione, Festa del Lavoro, Repubblica, Ferragosto,\n"
                 "Tutti i Santi, Immacolata, Natale, Santo Stefano",
            font=ctk.CTkFont(size=11),
            text_color="gray",
            justify="left"
        )
        info_label.pack(anchor="w", padx=5)

        # Festività extra
        extra_frame = ctk.CTkFrame(tab)
        extra_frame.pack(fill="both", expand=True, padx=10, pady=(10, 10))

        label = ctk.CTkLabel(
            extra_frame,
            text="Festività aggiuntive:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        label.pack(anchor="w", padx=15, pady=(15, 5))

        # Lista festività extra
        list_frame = ctk.CTkFrame(extra_frame, fg_color="transparent")
        list_frame.pack(fill="both", expand=True, padx=15, pady=(0, 10))

        self.holidays_textbox = ctk.CTkTextbox(
            list_frame,
            height=200,
            font=ctk.CTkFont(size=11)
        )
        self.holidays_textbox.pack(fill="both", expand=True)

        # Info formato
        format_label = ctk.CTkLabel(
            extra_frame,
            text="Formato: una festività per riga, formato GG/MM (es. 15/08)\nLascia vuoto se non ci sono festività aggiuntive",
            font=ctk.CTkFont(size=10),
            text_color="gray",
            justify="left"
        )
        format_label.pack(anchor="w", padx=15, pady=(0, 15))

    # ========================================================================
    # TAB: AVANZATE
    # ========================================================================

    def create_advanced_tab(self):
        """Crea tab opzioni avanzate"""
        tab = self.tabview.tab("Avanzate")

        scroll_frame = ctk.CTkScrollableFrame(tab)
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Verifica anomalie
        self.verifica_anomalie_var = ctk.BooleanVar(value=True)
        check1 = ctk.CTkCheckBox(
            scroll_frame,
            text="Verifica anomalie copertura 2-2-2",
            variable=self.verifica_anomalie_var,
            font=ctk.CTkFont(size=13)
        )
        check1.pack(anchor="w", padx=10, pady=10)

        # Genera report
        self.genera_report_var = ctk.BooleanVar(value=True)
        check2 = ctk.CTkCheckBox(
            scroll_frame,
            text="Genera report dettagliato",
            variable=self.genera_report_var,
            font=ctk.CTkFont(size=13)
        )
        check2.pack(anchor="w", padx=10, pady=10)

        # Progressivo per turno
        self.progressivo_var = ctk.BooleanVar(value=True)
        check3 = ctk.CTkCheckBox(
            scroll_frame,
            text="Mostra progressivo giorni per turno",
            variable=self.progressivo_var,
            font=ctk.CTkFont(size=13)
        )
        check3.pack(anchor="w", padx=10, pady=10)

        # Font
        font_frame = ctk.CTkFrame(scroll_frame)
        font_frame.pack(fill="x", pady=20)

        font_label = ctk.CTkLabel(
            font_frame,
            text="Formato Excel:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        font_label.pack(anchor="w", padx=10, pady=(10, 10))

        # Nome font
        name_frame = ctk.CTkFrame(font_frame, fg_color="transparent")
        name_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(name_frame, text="Font:").pack(side="left", padx=(0, 10))
        self.font_nome = ctk.CTkEntry(name_frame, width=150)
        self.font_nome.pack(side="left")

        # Dimensione titolo
        size_frame = ctk.CTkFrame(font_frame, fg_color="transparent")
        size_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(size_frame, text="Dimensione titolo:").pack(side="left", padx=(0, 10))
        self.font_size_title = ctk.CTkEntry(size_frame, width=60)
        self.font_size_title.pack(side="left")

        # Dimensione normale
        size_frame2 = ctk.CTkFrame(font_frame, fg_color="transparent")
        size_frame2.pack(fill="x", padx=10, pady=(5, 10))

        ctk.CTkLabel(size_frame2, text="Dimensione normale:").pack(side="left", padx=(0, 10))
        self.font_size_normal = ctk.CTkEntry(size_frame2, width=60)
        self.font_size_normal.pack(side="left")

    # ========================================================================
    # UTILITY
    # ========================================================================

    def create_labeled_entry(self, parent, label_text: str, key: str, placeholder: str):
        """Crea un entry con label"""
        frame = ctk.CTkFrame(parent)
        frame.pack(fill="x", pady=10)

        label = ctk.CTkLabel(
            frame,
            text=label_text,
            font=ctk.CTkFont(size=13, weight="bold")
        )
        label.pack(anchor="w", padx=10, pady=(10, 5))

        entry = ctk.CTkEntry(
            frame,
            placeholder_text=placeholder,
            width=200
        )
        entry.pack(anchor="w", padx=10, pady=(0, 10))

        # Salva riferimento
        setattr(self, key, entry)

    # ========================================================================
    # GESTIONE CONFIGURAZIONE
    # ========================================================================

    def load_config(self):
        """Carica la configurazione corrente"""
        config = self.config_manager.config

        # Regole
        self.max_giorni_lavorati.insert(0, str(config.get("max_giorni_lavorati", 231)))
        self.pattern_normale_giorni.insert(0, str(config.get("pattern_normale_giorni", 10)))
        self.pattern_estivo_giorni.insert(0, str(config.get("pattern_estivo_giorni", 9)))

        # Periodo estivo
        inizio = config.get("periodo_estivo_inizio", (6, 20))
        self.estivo_inizio_mese.insert(0, str(inizio[0]))
        self.estivo_inizio_giorno.insert(0, str(inizio[1]))

        fine = config.get("periodo_estivo_fine", (9, 13))
        self.estivo_fine_mese.insert(0, str(fine[0]))
        self.estivo_fine_giorno.insert(0, str(fine[1]))

        # Colori
        self.color_vars["colore_domeniche_festivi"].insert(0, config.get("colore_domeniche_festivi", "FF0000"))
        self.color_vars["colore_sabati"].insert(0, config.get("colore_sabati", "4472C4"))
        self.color_vars["colore_ferie"].insert(0, config.get("colore_ferie", "FFFFCC"))
        self.color_vars["colore_giornate_g"].insert(0, config.get("colore_giornate_g", "C6EFCE"))

        # Festività extra
        festivita_extra = config.get("festivita_extra", [])
        if festivita_extra:
            text = "\n".join([f"{g:02d}/{m:02d}" for m, g in festivita_extra])
            self.holidays_textbox.insert("1.0", text)

        # Avanzate
        self.verifica_anomalie_var.set(config.get("verifica_anomalie", True))
        self.genera_report_var.set(config.get("genera_report", True))
        self.progressivo_var.set(config.get("mostra_progressivo_per_turno", True))

        self.font_nome.insert(0, config.get("font_nome", "Calibri"))
        self.font_size_title.insert(0, str(config.get("font_dimensione_titolo", 14)))
        self.font_size_normal.insert(0, str(config.get("font_dimensione_normale", 11)))

    def save(self):
        """Salva le modifiche"""
        try:
            # Regole
            self.config_manager.set("max_giorni_lavorati", int(self.max_giorni_lavorati.get()))
            self.config_manager.set("pattern_normale_giorni", int(self.pattern_normale_giorni.get()))
            self.config_manager.set("pattern_estivo_giorni", int(self.pattern_estivo_giorni.get()))

            # Periodo estivo
            self.config_manager.set("periodo_estivo_inizio", (
                int(self.estivo_inizio_mese.get()),
                int(self.estivo_inizio_giorno.get())
            ))
            self.config_manager.set("periodo_estivo_fine", (
                int(self.estivo_fine_mese.get()),
                int(self.estivo_fine_giorno.get())
            ))

            # Colori
            for key, entry in self.color_vars.items():
                color = entry.get().replace("#", "").upper()
                self.config_manager.set(key, color)

            # Festività extra
            festivita_text = self.holidays_textbox.get("1.0", "end").strip()
            festivita_extra = []
            if festivita_text:
                for line in festivita_text.split("\n"):
                    line = line.strip()
                    if line and "/" in line:
                        try:
                            g, m = line.split("/")
                            festivita_extra.append((int(m), int(g)))
                        except:
                            pass
            self.config_manager.set("festivita_extra", festivita_extra)

            # Avanzate
            self.config_manager.set("verifica_anomalie", self.verifica_anomalie_var.get())
            self.config_manager.set("genera_report", self.genera_report_var.get())
            self.config_manager.set("mostra_progressivo_per_turno", self.progressivo_var.get())

            self.config_manager.set("font_nome", self.font_nome.get())
            self.config_manager.set("font_dimensione_titolo", int(self.font_size_title.get()))
            self.config_manager.set("font_dimensione_normale", int(self.font_size_normal.get()))

            self.modified = True
            self.destroy()

        except ValueError as e:
            messagebox.showerror(
                "Errore",
                "Alcuni valori non sono validi. Controlla i campi numerici."
            )

    def cancel(self):
        """Annulla e chiudi"""
        self.modified = False
        self.destroy()

    def reset_to_default(self):
        """Ripristina impostazioni di default"""
        result = messagebox.askyesno(
            "Conferma",
            "Ripristinare tutte le impostazioni ai valori di default?\n\nQuesta operazione non può essere annullata."
        )

        if result:
            self.config_manager.reset_to_default()
            self.modified = True
            self.destroy()
