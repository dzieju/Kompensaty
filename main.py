"""
Punkt wejścia aplikacji Automat Poczty PDF Filter.
"""
import sys
import logging
from pathlib import Path

# Dodaj katalog główny do ścieżki Python
sys.path.insert(0, str(Path(__file__).parent))

try:
    from gui.main_window import MainWindow
except ImportError as e:
    print(f"Błąd importu: {e}")
    print("Upewnij się, że wszystkie wymagane moduły są zainstalowane.")
    sys.exit(1)


def main():
    """Główna funkcja aplikacji."""
    try:
        # Utwórz i uruchom główne okno
        app = MainWindow()
        app.run()
        
    except Exception as e:
        logging.critical(f"Krytyczny błąd aplikacji: {e}")
        print(f"Krytyczny błąd: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()