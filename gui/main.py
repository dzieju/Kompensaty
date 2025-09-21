"""
Punkt wejścia aplikacji Automat Poczty PDF Filter.
"""
import sys
import logging
import os
import traceback
from pathlib import Path

# Dodaj katalog główny do ścieżki Python
current_dir = Path(__file__).parent.absolute()
if str(current_dir) not in sys.path:
    sys.path.insert(0, str(current_dir))

# Sprawdź dostępność modułów
missing_modules = []
try:
    import tkinter as tk
except ImportError:
    missing_modules.append("tkinter")

try:
    import exchangelib
except ImportError:
    missing_modules.append("exchangelib")

try:
    import PyPDF2
except ImportError:
    missing_modules.append("PyPDF2")

try:
    import dotenv
except ImportError:
    missing_modules.append("python-dotenv")

if missing_modules:
    print("Brakujące moduły:")
    for module in missing_modules:
        print(f"  - {module}")
    print("\nZainstaluj je za pomocą:")
    print(f"pip install {' '.join(missing_modules)}")
    sys.exit(1)

try:
    from gui.main_window import MainWindow
except ImportError as e:
    print(f"Błąd importu GUI: {e}")
    print(f"Katalog roboczy: {os.getcwd()}")
    print(f"Ścieżka pliku: {__file__}")
    print(f"Python path: {sys.path}")
    traceback.print_exc()
    sys.exit(1)


def main():
    """Główna funkcja aplikacji."""
    try:
        # Sprawdź wersję Pythona
        if sys.version_info < (3, 8):
            print("Ta aplikacja wymaga Python 3.8 lub nowszego")
            print(f"Twoja wersja: {sys.version}")
            sys.exit(1)
        
        # Sprawdź czy tkinter działa
        try:
            root = tk.Tk()
            root.withdraw()  # Ukryj okno testowe
            root.destroy()
        except Exception as e:
            print(f"Błąd inicjalizacji tkinter: {e}")
            print("Sprawdź czy masz zainstalowany kompletny Python z tkinter")
            sys.exit(1)
        
        # Utwórz i uruchom główne okno
        print("Uruchamianie aplikacji...")
        app = MainWindow()
        app.run()
        
    except KeyboardInterrupt:
        print("\nAplikacja przerwana przez użytkownika")
        sys.exit(0)
    except Exception as e:
        logging.critical(f"Krytyczny błąd aplikacji: {e}")
        print(f"Krytyczny błąd: {e}")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()