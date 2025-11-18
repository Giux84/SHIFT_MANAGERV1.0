"""
SHIFT MANAGER GUI
Interfaccia grafica per il generatore automatico di turni raffineria
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import threading
from datetime import date
from typing import Optional
import sys

# Import moduli locali
from config import ConfigManager, valida_anno, valida_file_excel, valida_percorso_output
from turni_generator import GeneratoreTurni

# ============================================================================
# CONFIGURAZIONE CUSTOMTKINTER
# ============================================================================

ctk.set_appearance_mode("System")  # "System", "Dark", "Light"
ctk.set_default_color_theme("blue")  # "blue", "green", "dark-blue"

# ============================================================================
# FINESTRA PRINCIPALE
# ============================================================================

class ShiftManagerApp(ctk.CTk):
    """Applicazione principale Shift Manager"""

    def __init__(self):
        super().__init__()

        # Configurazione finestra
        self.title("Shift Manager - Generatore Turni Raffineria")
        self.geometry("1100x950")
        self.minsize(1000, 850)

        # Massimizza la finestra all'avvio
        # self.state('zoomed')  # Decommentare per aprire massimizzata

        # Config manager
        self.config_manager = ConfigManager()

        # Variabili
        self.file_template_var = ctk.StringVar()
        self.anno_target_var = ctk.StringVar(value=str(date.today().year + 1))
        self.output_path_var = ctk.StringVar()

        # Variabili validazione
        self.file_template_valid = False
        self.anno_target_valid = True
        self.output_path_valid = False

        # Thread generazione
        self.generation_thread: Optional[threading.Thread] = None
        self.is_generating = False

        # Crea UI
        self.create_widgets()

        # Imposta path output di default
        self.set_default_output_path()

        # Bind eventi
        self.file_template_var.trace_add('write', self.validate_inputs)
        self.anno_target_var.trace_add('write', self.validate_inputs)
        self.output_path_var.trace_add('write', self.validate_inputs)

        # Validazione iniziale
        self.validate_inputs()

    def create_widgets(self):
        """Crea tutti i widget dell'interfaccia"""

        # ===== HEADER =====
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=(20, 10))

        title_label = ctk.CTkLabel(
            header_frame,
            text="🔄 Shift Manager",
            font=ctk.CTkFont(size=28, weight="bold")
        )
        title_label.pack()

        subtitle_label = ctk.CTkLabel(
            header_frame,
            text="Generatore Automatico Turni Raffineria",
            font=ctk.CTkFont(size=14),
            text_color="gray"
        )
        subtitle_label.pack()

        # ===== MAIN CONTENT =====
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # --- Sezione Input File ---
        self.create_file_section(main_frame)

        # --- Sezione Anno Target ---
        self.create_year_section(main_frame)

        # --- Sezione Output ---
        self.create_output_section(main_frame)

        # --- Sezione Pulsanti Azione ---
        self.create_action_section(main_frame)

        # --- Sezione Progress ---
        self.create_progress_section(main_frame)

        # --- Sezione Log ---
        self.create_log_section(main_frame)

        # ===== FOOTER =====
        footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        footer_frame.pack(fill="x", padx=20, pady=(0, 10))

        footer_label = ctk.CTkLabel(
            footer_frame,
            text="© 2025 Shift Manager v1.0 | Powered by CustomTkinter",
            font=ctk.CTkFont(size=10),
            text_color="gray"
        )
        footer_label.pack()

    def create_file_section(self, parent):
        """Crea la sezione per selezionare il file template"""
        section_frame = ctk.CTkFrame(parent)
        section_frame.pack(fill="x", padx=20, pady=(20, 10))

        # Label
        label = ctk.CTkLabel(
            section_frame,
            text="📁 File Anno Precedente",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        label.pack(anchor="w", padx=15, pady=(15, 5))

        # Input frame
        input_frame = ctk.CTkFrame(section_frame, fg_color="transparent")
        input_frame.pack(fill="x", padx=15, pady=(0, 15))

        self.file_entry = ctk.CTkEntry(
            input_frame,
            textvariable=self.file_template_var,
            placeholder_text="Seleziona il file .xls/.xlsx dell'anno precedente...",
            height=40,
            font=ctk.CTkFont(size=12)
        )
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        browse_btn = ctk.CTkButton(
            input_frame,
            text="Sfoglia",
            command=self.browse_file,
            width=100,
            height=40,
            font=ctk.CTkFont(size=12, weight="bold")
        )
        browse_btn.pack(side="left")

        # Validation label
        self.file_validation_label = ctk.CTkLabel(
            section_frame,
            text="",
            font=ctk.CTkFont(size=11),
            text_color="red"
        )
        self.file_validation_label.pack(anchor="w", padx=15, pady=(0, 10))

    def create_year_section(self, parent):
        """Crea la sezione per selezionare l'anno target"""
        section_frame = ctk.CTkFrame(parent)
        section_frame.pack(fill="x", padx=20, pady=10)

        # Label
        label = ctk.CTkLabel(
            section_frame,
            text="📅 Anno da Generare",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        label.pack(anchor="w", padx=15, pady=(15, 5))

        # Input frame
        input_frame = ctk.CTkFrame(section_frame, fg_color="transparent")
        input_frame.pack(fill="x", padx=15, pady=(0, 15))

        self.year_entry = ctk.CTkEntry(
            input_frame,
            textvariable=self.anno_target_var,
            placeholder_text="es. 2026",
            width=200,
            height=40,
            font=ctk.CTkFont(size=12)
        )
        self.year_entry.pack(side="left", padx=(0, 10))

        # Info label
        info_label = ctk.CTkLabel(
            input_frame,
            text="Inserisci l'anno per cui generare i turni",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        )
        info_label.pack(side="left")

        # Validation label
        self.year_validation_label = ctk.CTkLabel(
            section_frame,
            text="",
            font=ctk.CTkFont(size=11),
            text_color="red"
        )
        self.year_validation_label.pack(anchor="w", padx=15, pady=(0, 10))

    def create_output_section(self, parent):
        """Crea la sezione per selezionare il path di output"""
        section_frame = ctk.CTkFrame(parent)
        section_frame.pack(fill="x", padx=20, pady=10)

        # Label
        label = ctk.CTkLabel(
            section_frame,
            text="💾 Cartella Output",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        label.pack(anchor="w", padx=15, pady=(15, 5))

        # Input frame
        input_frame = ctk.CTkFrame(section_frame, fg_color="transparent")
        input_frame.pack(fill="x", padx=15, pady=(0, 15))

        self.output_entry = ctk.CTkEntry(
            input_frame,
            textvariable=self.output_path_var,
            placeholder_text="Seleziona la cartella dove salvare il file generato...",
            height=40,
            font=ctk.CTkFont(size=12)
        )
        self.output_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        browse_btn = ctk.CTkButton(
            input_frame,
            text="Sfoglia",
            command=self.browse_output,
            width=100,
            height=40,
            font=ctk.CTkFont(size=12, weight="bold")
        )
        browse_btn.pack(side="left")

        # Validation label
        self.output_validation_label = ctk.CTkLabel(
            section_frame,
            text="",
            font=ctk.CTkFont(size=11),
            text_color="red"
        )
        self.output_validation_label.pack(anchor="w", padx=15, pady=(0, 10))

    def create_action_section(self, parent):
        """Crea la sezione con i pulsanti di azione"""
        section_frame = ctk.CTkFrame(parent, fg_color="transparent")
        section_frame.pack(fill="x", padx=20, pady=20)

        # Bottone Genera (grande)
        self.generate_btn = ctk.CTkButton(
            section_frame,
            text="▶️  GENERA TURNI",
            command=self.generate_shifts,
            height=50,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color="#2B7A0B",  # Verde
            hover_color="#1F5A08"
        )
        self.generate_btn.pack(fill="x", padx=15, pady=(0, 10))

        # Bottoni secondari
        buttons_frame = ctk.CTkFrame(section_frame, fg_color="transparent")
        buttons_frame.pack(fill="x", padx=15)

        self.options_btn = ctk.CTkButton(
            buttons_frame,
            text="⚙️  Opzioni Avanzate",
            command=self.open_advanced_options,
            height=40,
            font=ctk.CTkFont(size=13)
        )
        self.options_btn.pack(side="left", fill="x", expand=True, padx=(0, 5))

        self.preview_btn = ctk.CTkButton(
            buttons_frame,
            text="👁️  Anteprima",
            command=self.open_preview,
            height=40,
            font=ctk.CTkFont(size=13),
            state="disabled"
        )
        self.preview_btn.pack(side="left", fill="x", expand=True, padx=5)

        help_btn = ctk.CTkButton(
            buttons_frame,
            text="ℹ️  Aiuto",
            command=self.show_help,
            height=40,
            font=ctk.CTkFont(size=13)
        )
        help_btn.pack(side="left", fill="x", expand=True, padx=(5, 0))

    def create_progress_section(self, parent):
        """Crea la sezione progress bar"""
        section_frame = ctk.CTkFrame(parent, fg_color="transparent")
        section_frame.pack(fill="x", padx=20, pady=(0, 10))

        self.progress_bar = ctk.CTkProgressBar(section_frame, height=20)
        self.progress_bar.pack(fill="x", padx=15)
        self.progress_bar.set(0)

        self.progress_label = ctk.CTkLabel(
            section_frame,
            text="",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        )
        self.progress_label.pack(padx=15, pady=(5, 0))

    def create_log_section(self, parent):
        """Crea la sezione log"""
        section_frame = ctk.CTkFrame(parent)
        section_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        label = ctk.CTkLabel(
            section_frame,
            text="📋 Log",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        label.pack(anchor="w", padx=15, pady=(15, 5))

        self.log_textbox = ctk.CTkTextbox(
            section_frame,
            height=150,
            font=ctk.CTkFont(family="Consolas", size=10),
            wrap="word"
        )
        self.log_textbox.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        self.log_textbox.configure(state="disabled")

    # ========================================================================
    # METODI UTILITY
    # ========================================================================

    def set_default_output_path(self):
        """Imposta il path di output di default alla cartella corrente"""
        default_path = os.path.dirname(os.path.abspath(__file__))
        self.output_path_var.set(default_path)

    def log(self, message: str, level: str = "info"):
        """Aggiunge un messaggio al log"""
        self.log_textbox.configure(state="normal")

        # Aggiungi timestamp
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")

        # Emoji per livello
        emoji_map = {
            "info": "ℹ️",
            "success": "✅",
            "warning": "⚠️",
            "error": "❌"
        }
        emoji = emoji_map.get(level, "ℹ️")

        formatted_message = f"[{timestamp}] {emoji} {message}\n"
        self.log_textbox.insert("end", formatted_message)
        self.log_textbox.see("end")
        self.log_textbox.configure(state="disabled")

    def clear_log(self):
        """Pulisce il log"""
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")

    def update_progress(self, value: float, message: str = ""):
        """Aggiorna la progress bar"""
        self.progress_bar.set(value)
        if message:
            self.progress_label.configure(text=message)

    # ========================================================================
    # VALIDAZIONE INPUT
    # ========================================================================

    def validate_inputs(self, *args):
        """Valida tutti gli input e aggiorna UI"""
        # Valida file template
        valid, error = valida_file_excel(self.file_template_var.get())
        self.file_template_valid = valid
        if not valid and self.file_template_var.get():
            self.file_validation_label.configure(text=f"❌ {error}")
        else:
            self.file_validation_label.configure(text="")

        # Valida anno
        try:
            anno = int(self.anno_target_var.get())
            valid, error = valida_anno(anno)
            self.anno_target_valid = valid
            if not valid:
                self.year_validation_label.configure(text=f"❌ {error}")
            else:
                self.year_validation_label.configure(text="")
        except ValueError:
            self.anno_target_valid = False
            if self.anno_target_var.get():
                self.year_validation_label.configure(text="❌ Anno non valido")
            else:
                self.year_validation_label.configure(text="")

        # Valida output path
        valid, error = valida_percorso_output(self.output_path_var.get())
        self.output_path_valid = valid
        if not valid and self.output_path_var.get():
            self.output_validation_label.configure(text=f"❌ {error}")
        else:
            self.output_validation_label.configure(text="")

        # Abilita/disabilita bottone genera
        all_valid = self.file_template_valid and self.anno_target_valid and self.output_path_valid
        if all_valid and not self.is_generating:
            self.generate_btn.configure(state="normal")
        else:
            self.generate_btn.configure(state="disabled")

    # ========================================================================
    # GESTIONE FILE BROWSER
    # ========================================================================

    def browse_file(self):
        """Apre dialog per selezionare file template"""
        filename = filedialog.askopenfilename(
            title="Seleziona file anno precedente",
            filetypes=[
                ("File Excel", "*.xls *.xlsx"),
                ("Tutti i file", "*.*")
            ]
        )
        if filename:
            self.file_template_var.set(filename)

    def browse_output(self):
        """Apre dialog per selezionare cartella output"""
        directory = filedialog.askdirectory(
            title="Seleziona cartella output"
        )
        if directory:
            self.output_path_var.set(directory)

    # ========================================================================
    # GENERAZIONE TURNI
    # ========================================================================

    def generate_shifts(self):
        """Avvia la generazione dei turni"""
        if self.is_generating:
            return

        # Conferma se file esiste già
        anno = int(self.anno_target_var.get())
        output_file = os.path.join(
            self.output_path_var.get(),
            f"TURNO_COMPLETO_{anno}.xlsx"
        )

        if os.path.exists(output_file):
            result = messagebox.askyesno(
                "File esistente",
                f"Il file {os.path.basename(output_file)} esiste già.\n\nSovrascrivere?"
            )
            if not result:
                return

        # Pulisci log
        self.clear_log()
        self.log("Avvio generazione turni...", "info")

        # Disabilita UI
        self.is_generating = True
        self.generate_btn.configure(state="disabled", text="⏳ Generazione in corso...")
        self.options_btn.configure(state="disabled")
        self.file_entry.configure(state="disabled")
        self.year_entry.configure(state="disabled")
        self.output_entry.configure(state="disabled")

        # Reset progress
        self.update_progress(0, "Preparazione...")

        # Avvia thread di generazione
        self.generation_thread = threading.Thread(
            target=self._generate_shifts_thread,
            args=(self.file_template_var.get(), anno, output_file),
            daemon=True
        )
        self.generation_thread.start()

    def _generate_shifts_thread(self, file_template: str, anno_target: int, output_file: str):
        """Thread di generazione turni"""
        try:
            # Creazione generatore
            self.after(0, lambda: self.log(f"Lettura template {anno_target - 1}...", "info"))
            self.after(0, lambda: self.update_progress(0.1, "Lettura template..."))

            generatore = GeneratoreTurni(file_template, anno_target)

            # Step 1: Lettura template
            if not generatore.leggi_template():
                raise Exception("Impossibile leggere il template")

            self.after(0, lambda: self.log("Template letto con successo", "success"))
            self.after(0, lambda: self.update_progress(0.3, "Template caricato"))

            # Step 2: Generazione calendario
            self.after(0, lambda: self.log(f"Generazione calendario {anno_target}...", "info"))
            self.after(0, lambda: self.update_progress(0.4, "Generazione calendario..."))

            generatore.genera_calendario()

            self.after(0, lambda: self.log("Calendario base generato", "success"))
            self.after(0, lambda: self.update_progress(0.6, "Calendario generato"))

            # Step 3: Applicazione ferie
            self.after(0, lambda: self.log("Applicazione ferie...", "info"))
            self.after(0, lambda: self.update_progress(0.7, "Applicazione ferie..."))

            generatore.applica_ferie()

            self.after(0, lambda: self.log("Ferie applicate", "success"))
            self.after(0, lambda: self.update_progress(0.8, "Ferie applicate"))

            # Step 4: Verifica copertura
            self.after(0, lambda: self.log("Verifica copertura 2-2-2...", "info"))
            self.after(0, lambda: self.update_progress(0.85, "Verifica copertura..."))

            verifica = generatore.verifica_copertura()

            giorni_ok = verifica['statistiche']['ok']
            anomalie = verifica['statistiche']['anomalie']

            if anomalie > 0:
                self.after(0, lambda: self.log(
                    f"⚠️ Trovate {anomalie} anomalie (vedi report)",
                    "warning"
                ))
            else:
                self.after(0, lambda: self.log(
                    f"Verifica OK: {giorni_ok} giorni corretti",
                    "success"
                ))

            # Step 5: Generazione Excel
            self.after(0, lambda: self.log("Generazione file Excel...", "info"))
            self.after(0, lambda: self.update_progress(0.9, "Creazione file Excel..."))

            generatore.genera_excel(output_file)

            # Genera report
            report_file = os.path.join(
                os.path.dirname(output_file),
                f"report_verifica_{anno_target}.txt"
            )
            generatore.genera_report(verifica, report_file)

            self.after(0, lambda: self.log(f"File salvato: {os.path.basename(output_file)}", "success"))
            self.after(0, lambda: self.log(f"Report salvato: {os.path.basename(report_file)}", "success"))
            self.after(0, lambda: self.update_progress(1.0, "Completato!"))

            # Mostra messaggio di successo
            self.after(0, lambda: self._show_generation_complete(output_file, anomalie))

        except Exception as e:
            self.after(0, lambda: self.log(f"ERRORE: {str(e)}", "error"))
            self.after(0, lambda: self.update_progress(0, "Errore"))
            self.after(0, lambda: messagebox.showerror(
                "Errore",
                f"Si è verificato un errore durante la generazione:\n\n{str(e)}"
            ))

        finally:
            # Riabilita UI
            self.after(0, self._reset_ui_after_generation)

    def _show_generation_complete(self, output_file: str, anomalie: int):
        """Mostra dialog di generazione completata"""
        message = f"File generato con successo!\n\n{os.path.basename(output_file)}"

        if anomalie > 0:
            message += f"\n\n⚠️ Attenzione: trovate {anomalie} anomalie.\nControlla il report per dettagli."
            icon = "warning"
        else:
            message += "\n\n✅ Nessuna anomalia rilevata."
            icon = "info"

        result = messagebox.askquestion(
            "Generazione Completata",
            message + "\n\nVuoi aprire la cartella?",
            icon=icon
        )

        if result == 'yes':
            # Apri cartella in Windows Explorer
            import subprocess
            subprocess.Popen(f'explorer /select,"{output_file}"')

        # Abilita preview
        self.preview_btn.configure(state="normal")

    def _reset_ui_after_generation(self):
        """Ripristina UI dopo generazione"""
        self.is_generating = False
        self.generate_btn.configure(text="▶️  GENERA TURNI")
        self.options_btn.configure(state="normal")
        self.file_entry.configure(state="normal")
        self.year_entry.configure(state="normal")
        self.output_entry.configure(state="normal")
        self.validate_inputs()

    # ========================================================================
    # ALTRE FUNZIONI
    # ========================================================================

    def open_advanced_options(self):
        """Apre dialog opzioni avanzate"""
        # Importa qui per evitare import circolari
        from advanced_options_dialog import AdvancedOptionsDialog

        dialog = AdvancedOptionsDialog(self, self.config_manager)
        dialog.grab_set()  # Rendi modale
        self.wait_window(dialog)

        # Salva configurazione se modificata
        if dialog.modified:
            self.config_manager.save()
            self.log("Configurazione aggiornata", "success")

    def open_preview(self):
        """Apre finestra anteprima"""
        # Importa qui per evitare import circolari
        from preview_dialog import PreviewDialog

        anno = int(self.anno_target_var.get())
        output_file = os.path.join(
            self.output_path_var.get(),
            f"TURNO_COMPLETO_{anno}.xlsx"
        )

        if not os.path.exists(output_file):
            messagebox.showwarning(
                "File non trovato",
                "Devi prima generare i turni per vedere l'anteprima."
            )
            return

        dialog = PreviewDialog(self, output_file)
        dialog.grab_set()

    def show_help(self):
        """Mostra finestra aiuto"""
        help_text = """
SHIFT MANAGER - GUIDA RAPIDA

1. SELEZIONE FILE TEMPLATE
   Seleziona il file Excel dell'anno precedente (.xls o .xlsx)

2. ANNO DA GENERARE
   Inserisci l'anno per cui vuoi generare i turni

3. CARTELLA OUTPUT
   Seleziona dove salvare il file generato

4. OPZIONI AVANZATE
   Personalizza regole, colori e festività

5. GENERA TURNI
   Clicca il bottone verde per avviare la generazione

6. ANTEPRIMA
   Dopo la generazione, visualizza l'anteprima del calendario

---

FUNZIONALITÀ PRINCIPALI:
• Continuità automatica del ciclo tra anni
• Rotazione quinquennale ferie (1→3→5→2→4)
• Verifica automatica copertura 2-2-2
• Formattazione colori automatica
• Report anomalie dettagliato

Per supporto: contatta l'amministratore
        """

        messagebox.showinfo("Aiuto", help_text)

# ============================================================================
# MAIN
# ============================================================================

def main():
    """Funzione principale"""
    app = ShiftManagerApp()
    app.mainloop()

if __name__ == "__main__":
    main()
