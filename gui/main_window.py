"""
Główne okno aplikacji GUI.
"""
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import logging
import os
import sys
import json
from typing import Optional

from config import AppConfig
from models import SearchCriteria, EmailTemplateCollection
from email_processor import EmailProcessor
from exceptions import OperationCancelledError
from gui.config_tab import ConfigTab
from gui.search_criteria_tab import SearchCriteriaTab
from gui.templates_tab import TemplatesTab
from gui.logs_tab import LogsTab
from gui.about_tab import AboutTab


class MainWindow:
    """Główne okno aplikacji."""
    
    def __init__(self):
        """Inicjalizuje główne okno."""
        print("DEBUG: Inicjalizacja MainWindow...")
        
        self.root = tk.Tk()
        self.root.title("Automat Poczty PDF Filter v3.0")
        self.root.geometry("1000x800")
        self.root.minsize(800, 600)  # Minimalna wielkość okna
        print("DEBUG: Tkinter root utworzony")
        
        # Konfiguracja i dane
        self.config = AppConfig()
        self.search_criteria = SearchCriteria()
        self.templates = EmailTemplateCollection.get_default()
        self.processor: Optional[EmailProcessor] = None
        print("DEBUG: Obiekty danych utworzone")
        
        # Event do anulowania operacji
        self.stop_event = threading.Event()
        
        # Referencje do zakładek
        self.config_tab: Optional[ConfigTab] = None
        self.criteria_tab: Optional[SearchCriteriaTab] = None
        self.templates_tab: Optional[TemplatesTab] = None
        self.logs_tab: Optional[LogsTab] = None
        self.about_tab: Optional[AboutTab] = None
        
        # Przyciski sterujące
        self.run_button: Optional[ttk.Button] = None
        self.stop_button: Optional[ttk.Button] = None
        print("DEBUG: Zmienne składowe zainicjalizowane")
        
        try:
            print("DEBUG: Rozpoczynam setup_ui...")
            self._setup_ui()
            print("DEBUG: setup_ui zakończone")
            
            print("DEBUG: Rozpoczynam setup_logging...")
            self._setup_logging()
            print("DEBUG: setup_logging zakończone")
            
            print("DEBUG: Rozpoczynam load_all_configurations...")
            self._load_all_configurations()
            print("DEBUG: load_all_configurations zakończone")
            
        except Exception as e:
            print(f"DEBUG: BŁĄD podczas inicjalizacji: {e}")
            import traceback
            traceback.print_exc()
            raise
        
        # Obsługa zamykania okna
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        print("DEBUG: MainWindow zainicjalizowane pomyślnie")

    def _setup_ui(self) -> None:
        """Konfiguruje interfejs użytkownika używając Grid Layout."""
        try:
            print("DEBUG: Rozpoczynam setup_ui z Grid Layout")
            
            # Konfiguracja Grid - KLUCZOWE dla prawidłowego layoutu
            self.root.grid_rowconfigure(0, weight=1)  # Notebook rozciąga się
            self.root.grid_rowconfigure(1, weight=0)  # Przyciski stały rozmiar
            self.root.grid_rowconfigure(2, weight=0)  # Status bar stały rozmiar
            self.root.grid_columnconfigure(0, weight=1)
            print("DEBUG: Grid skonfigurowany")
            
            # Notebook w rzędzie 0 - zajmuje większość miejsca
            print("DEBUG: Tworzę notebook...")
            self.notebook = ttk.Notebook(self.root)
            self.notebook.grid(row=0, column=0, sticky='nsew', padx=10, pady=(10, 5))
            print("DEBUG: Notebook utworzony i umieszczony w grid")
            
            # Tworzenie zakładek
            print("DEBUG: Tworzę zakładki...")
            self._create_tabs()
            print("DEBUG: Zakładki utworzone")
            
            # Panel przycisków w rzędzie 1 - zawsze widoczny
            print("DEBUG: Tworzę panel przycisków...")
            self._create_button_panel()
            print("DEBUG: Panel przycisków utworzony")
            
            # Status bar w rzędzie 2 - na samym dole
            print("DEBUG: Tworzę status bar...")
            self._create_status_bar()
            print("DEBUG: Status bar utworzony")
            
            print("DEBUG: setup_ui zakończone pomyślnie")
            
        except Exception as e:
            print(f"DEBUG: BŁĄD w setup_ui: {e}")
            import traceback
            traceback.print_exc()
            raise

    def _create_tabs(self) -> None:
        """Tworzy wszystkie zakładki."""
        try:
            print("DEBUG: Tworzę zakładkę konfiguracji...")
            # Zakładka konfiguracji głównej
            config_frame = ttk.Frame(self.notebook)
            self.notebook.add(config_frame, text='Konfiguracja Główna')
            self.config_tab = ConfigTab(config_frame)
            print("DEBUG: Zakładka konfiguracji utworzona")
            
            print("DEBUG: Tworzę zakładkę kryteriów...")
            # Zakładka kryteriów wyszukiwania
            criteria_frame = ttk.Frame(self.notebook)
            self.notebook.add(criteria_frame, text='Kryteria Wyszukiwania')
            self.criteria_tab = SearchCriteriaTab(criteria_frame)
            print("DEBUG: Zakładka kryteriów utworzona")
            
            print("DEBUG: Tworzę zakładkę szablonów...")
            # Zakładka szablonów
            templates_frame = ttk.Frame(self.notebook)
            self.notebook.add(templates_frame, text='Szablony E-mail')
            self.templates_tab = TemplatesTab(templates_frame)
            print("DEBUG: Zakładka szablonów utworzona")
            
            print("DEBUG: Tworzę zakładkę logów...")
            # Zakładka logów
            logs_frame = ttk.Frame(self.notebook)
            self.notebook.add(logs_frame, text='Logi')
            self.logs_tab = LogsTab(logs_frame)
            print("DEBUG: Zakładka logów utworzona")
            
            print("DEBUG: Tworzę zakładkę o programie...")
            # Zakładka o programie
            about_frame = ttk.Frame(self.notebook)
            self.notebook.add(about_frame, text='O Programie')
            self.about_tab = AboutTab(about_frame)
            print("DEBUG: Zakładka o programie utworzona")
            
            print("DEBUG: Wszystkie zakładki utworzone")
            
        except Exception as e:
            print(f"DEBUG: BŁĄD w create_tabs: {e}")
            import traceback
            traceback.print_exc()
            raise

    def _create_button_panel(self) -> None:
        """Tworzy panel z przyciskami używając Grid Layout."""
        try:
            print("DEBUG: Rozpoczynam tworzenie panelu przycisków...")
            
            # Panel przycisków w rzędzie 1
            button_frame = ttk.Frame(self.root)
            button_frame.grid(row=1, column=0, sticky='ew', padx=10, pady=5)
            button_frame.grid_columnconfigure(8, weight=1)  # Ostatnia kolumna rozciąga się
            print("DEBUG: button_frame utworzony w grid")
            
            col = 0  # Licznik kolumn
            
            # Sekcja 1: Przyciski operacyjne
            print("DEBUG: Tworzę przyciski operacyjne...")
            self.run_button = ttk.Button(
                button_frame,
                text="🚀 Uruchom Przetwarzanie",
                command=self._run_processing_thread,
                width=20
            )
            self.run_button.grid(row=0, column=col, padx=5, pady=2)
            col += 1
            
            self.stop_button = ttk.Button(
                button_frame,
                text="⏹ Zatrzymaj",
                command=self._stop_processing,
                state=tk.DISABLED,
                width=15
            )
            self.stop_button.grid(row=0, column=col, padx=5, pady=2)
            col += 1
            print("DEBUG: Przyciski operacyjne utworzone")
            
            # Separator 1
            ttk.Separator(button_frame, orient='vertical').grid(row=0, column=col, sticky='ns', padx=10)
            col += 1
            
            # Sekcja 2: Przyciski zapisywania
            print("DEBUG: Tworzę przyciski zapisywania...")
            ttk.Button(
                button_frame,
                text="💾 Zapisz Konfigurację",
                command=self._save_main_config,
                width=18
            ).grid(row=0, column=col, padx=2, pady=2)
            col += 1
            
            ttk.Button(
                button_frame,
                text="🔍 Zapisz Kryteria", 
                command=self._save_search_criteria,
                width=15
            ).grid(row=0, column=col, padx=2, pady=2)
            col += 1
            
            ttk.Button(
                button_frame,
                text="📧 Zapisz Szablony",
                command=self._save_templates,
                width=15
            ).grid(row=0, column=col, padx=2, pady=2)
            col += 1
            print("DEBUG: Przyciski zapisywania utworzone")
            
            # Separator 2
            ttk.Separator(button_frame, orient='vertical').grid(row=0, column=col, sticky='ns', padx=10)
            col += 1
            
            # Sekcja 3: Przyciski testowe
            print("DEBUG: Tworzę przyciski testowe...")
            ttk.Button(
                button_frame,
                text="🔌 Test Połączenia",
                command=self._test_connection,
                width=15
            ).grid(row=0, column=col, padx=2, pady=2)
            col += 1
            
            ttk.Button(
                button_frame,
                text="📊 Statystyki",
                command=self._show_statistics,
                width=12
            ).grid(row=0, column=col, padx=2, pady=2)
            col += 1
            print("DEBUG: Przyciski testowe utworzone")
            
            print("DEBUG: Panel przycisków utworzony pomyślnie!")
            
        except Exception as e:
            print(f"DEBUG: BŁĄD w create_button_panel: {e}")
            import traceback
            traceback.print_exc()
            raise

    def _create_status_bar(self) -> None:
        """Tworzy pasek statusu używając Grid Layout."""
        try:
            print("DEBUG: Tworzę status bar...")
            
            # Status bar w rzędzie 2
            self.status_frame = ttk.Frame(self.root)
            self.status_frame.grid(row=2, column=0, sticky='ew', padx=10, pady=5)
            self.status_frame.grid_columnconfigure(0, weight=1)
            
            self.status_var = tk.StringVar(value="Gotowy")
            self.status_label = ttk.Label(
                self.status_frame,
                textvariable=self.status_var,
                relief='sunken',
                anchor='w'
            )
            self.status_label.grid(row=0, column=0, sticky='ew')
            
            # Progressbar (ukryty domyślnie) - będzie dodawany dynamicznie
            self.progress_var = tk.DoubleVar()
            self.progress_bar = ttk.Progressbar(
                self.status_frame,
                variable=self.progress_var,
                mode='indeterminate'
            )
            
            print("DEBUG: Status bar utworzony")
            
        except Exception as e:
            print(f"DEBUG: BŁĄD w create_status_bar: {e}")
            import traceback
            traceback.print_exc()
            raise

    def _setup_logging(self) -> None:
        """Konfiguruje system logowania."""
        try:
            print("DEBUG: Konfiguracja logowania...")
            
            # Konfiguracja głównego loggera
            logging.basicConfig(
                level=logging.INFO,
                format='%(asctime)s - %(levelname)s - %(name)s - %(message)s',
                handlers=[]
            )
            
            # File handler
            try:
                file_handler = logging.FileHandler(
                    self.config.log_file,
                    mode='a',
                    encoding='utf-8'
                )
                file_formatter = logging.Formatter(
                    '%(asctime)s - %(levelname)s - %(name)s - %(funcName)s - %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S'
                )
                file_handler.setFormatter(file_formatter)
                logging.getLogger('').addHandler(file_handler)
                print("DEBUG: File handler skonfigurowany")
            except Exception as e:
                print(f"DEBUG: Błąd konfiguracji logowania do pliku: {e}")
            
            # Wyłącz verbose logging dla exchangelib
            logging.getLogger('exchangelib').setLevel(logging.WARNING)
            
            logging.info("Aplikacja uruchomiona")
            print("DEBUG: Logowanie skonfigurowane")
            
        except Exception as e:
            print(f"DEBUG: BŁĄD w setup_logging: {e}")
            import traceback
            traceback.print_exc()

    def _load_all_configurations(self) -> None:
        """Ładuje wszystkie konfiguracje."""
        try:
            print("DEBUG: Ładowanie konfiguracji...")
            
            # Ładuj konfigurację główną
            env_path = self._get_env_file_path()
            print(f"DEBUG: Ścieżka .env: {env_path}")
            
            self.config = AppConfig.from_environment(env_path)
            if self.config_tab:
                self.config_tab.load_from_config(self.config)
            print("DEBUG: Konfiguracja główna załadowana")
            
            # Ładuj kryteria wyszukiwania
            self._load_search_criteria()
            print("DEBUG: Kryteria wyszukiwania załadowane")
            
            # Ładuj szablony
            self._load_templates()
            print("DEBUG: Szablony załadowane")
            
        except Exception as e:
            print(f"DEBUG: BŁĄD w load_all_configurations: {e}")
            import traceback
            traceback.print_exc()

    def _load_search_criteria(self) -> None:
        """Ładuje kryteria wyszukiwania z pliku."""
        criteria_path = self._get_search_criteria_file_path()
        
        try:
            if os.path.exists(criteria_path):
                with open(criteria_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.search_criteria = SearchCriteria.from_dict(data)
                    logging.info("Wczytano kryteria wyszukiwania z pliku")
            else:
                self.search_criteria = SearchCriteria()
                logging.info("Użyto domyślnych kryteriów wyszukiwania")
        except Exception as e:
            logging.error(f"Błąd wczytywania kryteriów: {e}")
            self.search_criteria = SearchCriteria()
        
        if self.criteria_tab:
            self.criteria_tab.load_from_search_criteria(self.search_criteria)

    def _load_templates(self) -> None:
        """Ładuje szablony emaili z pliku."""
        templates_path = self._get_templates_file_path()
        
        try:
            if os.path.exists(templates_path):
                with open(templates_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.templates = EmailTemplateCollection.from_dict(data)
                    logging.info("Wczytano szablony emaili z pliku")
            else:
                self.templates = EmailTemplateCollection.get_default()
                logging.info("Użyto domyślnych szablonów emaili")
        except Exception as e:
            logging.error(f"Błąd wczytywania szablonów: {e}")
            self.templates = EmailTemplateCollection.get_default()
        
        if self.templates_tab:
            self.templates_tab.load_from_templates(self.templates)

    def _save_main_config(self) -> None:
        """Zapisuje główną konfigurację."""
        try:
            print("DEBUG: Zapisywanie głównej konfiguracji...")
            
            if self.config_tab:
                self.config_tab.save_to_config(self.config)
            
            # Zapisz do pliku .env
            env_path = self._get_env_file_path()
            self._save_config_to_env_file(env_path)
            
            messagebox.showinfo("Sukces", f"Konfiguracja zapisana do {env_path}")
            logging.info("Konfiguracja główna zapisana pomyślnie")
            
        except Exception as e:
            print(f"DEBUG: Błąd zapisywania konfiguracji: {e}")
            logging.error(f"Błąd zapisywania konfiguracji: {e}")
            messagebox.showerror("Błąd", f"Nie udało się zapisać konfiguracji: {e}")

    def _save_search_criteria(self) -> None:
        """Zapisuje kryteria wyszukiwania."""
        try:
            print("DEBUG: Zapisywanie kryteriów wyszukiwania...")
            
            if self.criteria_tab:
                self.search_criteria = self.criteria_tab.save_to_search_criteria()
            
            criteria_path = self._get_search_criteria_file_path()
            with open(criteria_path, 'w', encoding='utf-8') as f:
                json.dump(self.search_criteria.to_dict(), f, indent=2, ensure_ascii=False)
            
            messagebox.showinfo("Sukces", f"Kryteria wyszukiwania zapisane do {criteria_path}")
            logging.info("Kryteria wyszukiwania zapisane pomyślnie")
            
        except Exception as e:
            print(f"DEBUG: Błąd zapisywania kryteriów: {e}")
            logging.error(f"Błąd zapisywania kryteriów: {e}")
            messagebox.showerror("Błąd", f"Nie udało się zapisać kryteriów: {e}")

    def _save_templates(self) -> None:
        """Zapisuje szablony emaili."""
        try:
            print("DEBUG: Zapisywanie szablonów emaili...")
            
            if self.templates_tab:
                self.templates = self.templates_tab.save_to_templates()
            
            templates_path = self._get_templates_file_path()
            with open(templates_path, 'w', encoding='utf-8') as f:
                json.dump(self.templates.to_dict(), f, indent=2, ensure_ascii=False)
            
            messagebox.showinfo("Sukces", f"Szablony emaili zapisane do {templates_path}")
            logging.info("Szablony emaili zapisane pomyślnie")
            
        except Exception as e:
            print(f"DEBUG: Błąd zapisywania szablonów: {e}")
            logging.error(f"Błąd zapisywania szablonów: {e}")
            messagebox.showerror("Błąd", f"Nie udało się zapisać szablonów: {e}")

    def _run_processing_thread(self) -> None:
        """Uruchamia przetwarzanie w osobnym wątku."""
        print("DEBUG: Uruchamianie przetwarzania...")
        
        # Walidacja konfiguracji
        if not self._validate_configuration():
            return
        
        # Przygotuj GUI do pracy
        self._set_processing_state(True)
        
        # Uruchom wątek przetwarzania
        processing_thread = threading.Thread(target=self._processing_task, daemon=True)
        processing_thread.start()

    def _processing_task(self) -> None:
        """Główne zadanie przetwarzania."""
        try:
            self.stop_event.clear()
            
            # Aktualizuj konfigurację z GUI
            self._update_config_from_gui()
            
            # Utwórz procesor
            self.processor = EmailProcessor(self.config, self.templates)
            
            # Uruchom przetwarzanie
            logging.info(">>> Rozpoczynanie głównej logiki przetwarzania emaili...")
            self._update_status("Przetwarzanie emaili w toku...")
            
            result = self.processor.process_emails(self.search_criteria, self.stop_event)
            
            if self.stop_event.is_set():
                result.message = "Przetwarzanie anulowane przez użytkownika"
                logging.warning(result.message)
            
            logging.info(f">>> Przetwarzanie zakończone. Status: {result.message}")
            
            # Pokaż wynik
            self.root.after(0, lambda: self._show_processing_result(result))
            
        except OperationCancelledError:
            logging.warning("Operacja przetwarzania została anulowana")
            self.root.after(0, lambda: messagebox.showwarning(
                "Anulowano", "Operacja przetwarzania została anulowana"
            ))
        except Exception as e:
            logging.error(f"Krytyczny błąd podczas przetwarzania: {e}")
            self.root.after(0, lambda: messagebox.showerror(
                "Błąd Krytyczny", f"Wystąpił błąd podczas przetwarzania:\n{e}"
            ))
        finally:
            self.root.after(0, lambda: self._set_processing_state(False))

    def _stop_processing(self) -> None:
        """Zatrzymuje przetwarzanie."""
        if not self.stop_event.is_set():
            logging.info("Żądanie zatrzymania przetwarzania wysłane")
            self.stop_event.set()
            self._update_status("Anulowanie operacji...")
            self.stop_button.config(state=tk.DISABLED)
        else:
            logging.info("Przetwarzanie już jest w trakcie zatrzymywania")

    def _set_processing_state(self, processing: bool) -> None:
        """Ustawia stan GUI podczas przetwarzania."""
        if processing:
            self.run_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            self.progress_bar.grid(row=0, column=1, sticky='e', padx=(10, 0))
            self.progress_bar.start(10)
            self._update_status("Przetwarzanie w toku...")
        else:
            self.run_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            self.progress_bar.grid_remove()
            self.progress_bar.stop()
            self._update_status("Gotowy")
            self.stop_event.clear()

    def _update_status(self, message: str) -> None:
        """Aktualizuje pasek statusu."""
        self.status_var.set(message)
        self.root.update_idletasks()

    def _validate_configuration(self) -> bool:
        """Waliduje konfigurację przed uruchomieniem."""
        if self.config_tab:
            values = self.config_tab.get_all_values()
            
            if not values.get("EXCHANGE_USER"):
                messagebox.showerror("Błąd konfiguracji", "Pole 'Użytkownik (Email)' jest wymagane")
                return False
                
            if not values.get("EXCHANGE_PASS"):
                messagebox.showerror("Błąd konfiguracji", "Pole 'Hasło' jest wymagane")
                return False
        
        return True

    def _update_config_from_gui(self) -> None:
        """Aktualizuje konfigurację z wartości z GUI."""
        if self.config_tab:
            self.config_tab.save_to_config(self.config)
        if self.criteria_tab:
            self.search_criteria = self.criteria_tab.save_to_search_criteria()
        if self.templates_tab:
            self.templates = self.templates_tab.save_to_templates()

    def _show_processing_result(self, result) -> None:
        """Pokazuje wynik przetwarzania."""
        if result.success:
            messagebox.showinfo("Sukces", result.message)
        else:
            messagebox.showwarning("Informacja", result.message)

    def _test_connection(self) -> None:
        """Testuje połączenie z Exchange."""
        try:
            print("DEBUG: Test połączenia...")
            self._update_config_from_gui()
            
            # Walidacja podstawowych danych
            if not self.config.exchange.username or not self.config.exchange.password:
                messagebox.showerror("Błąd", "Brak danych logowania Exchange")
                return
            
            self._update_status("Testowanie połączenia...")
            
            # Utwórz procesor i przetestuj połączenie
            processor = EmailProcessor(self.config, self.templates)
            success, message = processor.test_exchange_connection()
            
            if success:
                messagebox.showinfo("Test połączenia", f"✅ {message}")
                logging.info(f"Test połączenia pomyślny: {message}")
            else:
                messagebox.showerror("Test połączenia", f"❌ {message}")
                logging.error(f"Test połączenia nieudany: {message}")
                
        except Exception as e:
            error_msg = f"Błąd testu połączenia: {e}"
            messagebox.showerror("Błąd", error_msg)
            logging.error(error_msg)
        finally:
            self._update_status("Gotowy")

    def _show_statistics(self) -> None:
        """Pokazuje statystyki konfiguracji."""
        try:
            print("DEBUG: Pokazywanie statystyk...")
            self._update_config_from_gui()
            
            if not self.processor:
                self.processor = EmailProcessor(self.config, self.templates)
            
            stats = self.processor.get_processing_statistics()
            
            # Utwórz okno ze statystykami
            stats_window = tk.Toplevel(self.root)
            stats_window.title("Statystyki Konfiguracji")
            stats_window.geometry("500x400")
            stats_window.resizable(False, False)
            
            # Centrum względem głównego okna
            stats_window.transient(self.root)
            stats_window.grab_set()
            
            # Treść
            text_area = tk.Text(stats_window, wrap=tk.WORD, font=("Consolas", 10))
            text_area.pack(expand=True, fill='both', padx=10, pady=10)
            
            stats_text = "STATYSTYKI KONFIGURACJI\n" + "="*50 + "\n\n"
            
            for key, value in stats.items():
                formatted_key = key.replace('_', ' ').title()
                stats_text += f"{formatted_key:<30}: {value}\n"
            
            text_area.insert('1.0', stats_text)
            text_area.config(state='disabled')
            
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się pobrać statystyk: {e}")

    def _get_env_file_path(self) -> str:
        """Zwraca ścieżkę do pliku .env."""
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '.env')

    def _get_search_criteria_file_path(self) -> str:
        """Zwraca ścieżkę do pliku kryteriów wyszukiwania."""
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'search_criteria.json')

    def _get_templates_file_path(self) -> str:
        """Zwraca ścieżkę do pliku szablonów."""
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'email_templates.json')

    def _save_config_to_env_file(self, env_path: str) -> None:
        """Zapisuje konfigurację do pliku .env."""
        try:
            print("DEBUG: Sprawdzam dostępność modułu dotenv...")
            from dotenv import set_key
            print("DEBUG: Moduł dotenv dostępny")
            
            if not os.path.exists(os.path.dirname(env_path)) and os.path.dirname(env_path):
                os.makedirs(os.path.dirname(env_path))
            
            if not os.path.exists(env_path):
                open(env_path, 'a').close()
            
            # Zapisz wszystkie wartości konfiguracji
            config_items = [
                ('EXCHANGE_USER', self.config.exchange.username or ''),
                ('EXCHANGE_PASS', self.config.exchange.password or ''),
                ('EXCHANGE_SERVER_ADDR', self.config.exchange.server or ''),
                ('IGNORE_SSL_ERRORS', str(self.config.exchange.ignore_ssl).lower()),
                ('EMAIL_KSIEGOWEJ_GLOWNY', self.config.email.ksiegowa_glowny),
                ('EMAIL_KSIEGOWEJ_2', self.config.email.ksiegowa_2 or ''),
                ('EMAIL_KSIEGOWEJ_3', self.config.email.ksiegowa_3 or ''),
                ('FOLDER_PRZETWORZONE_KOMPENSATY', self.config.folders.przetworzone),
                ('MIESIACE_WSTECZ_KOMPENSATY', str(self.config.search.miesiace_wstecz_kompensaty)),
                ('MIESIACE_WSTECZ_FAKTURY', str(self.config.search.miesiace_wstecz_faktury))
            ]
            
            # Dodaj adresy powiadomień
            for i, email in enumerate(self.config.email.powiadomienia_brak_pdf[:4], 1):
                config_items.append((f'EMAIL_POWIADOMIENIA_BRAK_PDF_{i}', email))
            
            for i, email in enumerate(self.config.email.powiadomienia_czesciowe[:4], 1):
                config_items.append((f'EMAIL_POWIADOMIENIA_CZESCIOWE_{i}', email))
            
            for key, value in config_items:
                set_key(env_path, key, value, quote_mode="never")
                
            print("DEBUG: Konfiguracja zapisana do .env")
            
        except ImportError:
            print("DEBUG: BŁĄD - brak modułu python-dotenv")
            raise Exception("Brak modułu python-dotenv. Zainstaluj: pip install python-dotenv")
        except Exception as e:
            print(f"DEBUG: Błąd zapisywania do .env: {e}")
            raise

    def _on_closing(self) -> None:
        """Obsługuje zamykanie aplikacji."""
        if self.stop_event.is_set() or (self.run_button and self.run_button['state'] == tk.DISABLED):
            if messagebox.askokcancel("Zamykanie", "Przetwarzanie w toku. Czy na pewno zamknąć?"):
                self.stop_event.set()
                logging.info("Aplikacja zamykana przez użytkownika podczas przetwarzania")
                self.root.destroy()
        else:
            logging.info("Aplikacja zamykana")
            self.root.destroy()

    def run(self) -> None:
        """Uruchamia główną pętlę aplikacji."""
        try:
            print("DEBUG: Uruchamianie mainloop...")
            self.root.mainloop()
            print("DEBUG: Mainloop zakończony")
        except KeyboardInterrupt:
            logging.info("Aplikacja przerwana przez Ctrl+C")
        except Exception as e:
            logging.error(f"Nieoczekiwany błąd aplikacji: {e}")
            print(f"DEBUG: Błąd w mainloop: {e}")
            raise