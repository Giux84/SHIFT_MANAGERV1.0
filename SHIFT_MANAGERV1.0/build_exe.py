"""
Script per creare l'eseguibile di Shift Manager con PyInstaller
"""

import PyInstaller.__main__
import os
import sys

# Path corrente
current_dir = os.path.dirname(os.path.abspath(__file__))

# Parametri PyInstaller
PyInstaller.__main__.run([
    'shift_manager_gui.py',  # File principale
    '--name=ShiftManager',   # Nome eseguibile
    '--onefile',             # Singolo file .exe
    '--windowed',            # No console window
    # '--clean',             # Pulisci cache (commentato per evitare errori permessi)

    # Icona (opzionale - se non c'è, usa default)
    # '--icon=icon.ico',

    # Dati addizionali necessari per CustomTkinter
    '--collect-all=customtkinter',

    # Hidden imports
    '--hidden-import=openpyxl',
    '--hidden-import=pandas',
    '--hidden-import=xlrd',
    '--hidden-import=darkdetect',
    '--hidden-import=PIL',

    # Ottimizzazioni
    '--noupx',  # Non usare UPX compression (può causare falsi positivi antivirus)

    # Percorso output
    f'--distpath={os.path.join(current_dir, "dist")}',
    f'--workpath={os.path.join(current_dir, "build")}',
    f'--specpath={current_dir}',
])

print("\n" + "="*70)
print("ESEGUIBILE CREATO CON SUCCESSO!")
print("="*70)
print(f"\nTrova il file eseguibile in:")
print(f"{os.path.join(current_dir, 'dist', 'ShiftManager.exe')}")
print("\n" + "="*70)
