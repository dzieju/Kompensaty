"""
Zakładka konfiguracji głównej aplikacji.
"""
import tkinter as tk
from tkinter import ttk
from gui.base_tab import BaseTab


class ConfigTab(BaseTab):
    """Zakładka konfiguracji głównej."""
    
    def setup_ui(self) -> None:
        """Konfiguruje interfejs użytkownika."""
        self._create_auth_section()
        self._create_accounting_emails_section()
        self._create_notification_sections()
        self._create_folders_section()
        self._create_time_params_section()

    def _create_auth_section(self) -> None:
        """Tworzy sekcję danych logowania."""
        frame = self.create_labeled_frame("Dane Logowania Exchange")
        
        self.create_config_entry(frame, "Użytkownik (Email):", "EXCHANGE_USER", 0)
        self.create_config_entry(frame, "Hasło:", "EXCHANGE_PASS", 1, show_char="*")
        self.create_config_entry(frame, "Serwer (opcjonalnie):", "EXCHANGE_SERVER_ADDR", 2)
        self.create_config_entry(frame, "Ignoruj błędy SSL:", "IGNORE_SSL_ERRORS", 3, is_bool=True)
        
        frame.grid_columnconfigure(1, weight=1)

    def _create_accounting_emails_section(self) -> None:
        """Tworzy sekcję adresów księgowości."""
        frame = self.create_labeled_frame("Adresy Email (Raport do Księgowości)")
        
        self.create_config_entry(frame, "Główny Odbiorca:", "EMAIL_KSIEGOWEJ_GLOWNY", 0, col=0, width=30)
        self.create_config_entry(frame, "Odbiorca 2:", "EMAIL_KSIEGOWEJ_2", 0, col=2, width=30)
        self.create_config_entry(frame, "Odbiorca 3:", "EMAIL_KSIEGOWEJ_3", 1, col=0, width=30)
        
        frame.grid_columnconfigure(1, weight=1)
        frame.grid_columnconfigure(3, weight=1)

    def _create_notification_sections(self) -> None:
        """Tworzy sekcje powiadomień."""
        # Powiadomienia o braku PDF
        frame_pdf = self.create_labeled_frame("Powiadomienie: Problem z ekstrakcją faktur (do 4 odbiorców)")
        
        for i in range(4):
            row_idx, col_idx = i // 2, (i % 2) * 2
            self.create_config_entry(
                frame_pdf, 
                f"Odbiorca {i+1}:", 
                f"EMAIL_POWIADOMIENIA_BRAK_PDF_{i+1}", 
                row_idx, 
                col_idx, 
                width=30
            )
        
        ttk.Label(
            frame_pdf, 
            text="(domyślnie: Użytkownik (Email), jeśli wszystkie puste)", 
            font=("Segoe UI", 8)
        ).grid(row=2, column=0, columnspan=4, sticky="w", padx=5, pady=(0,5))
        
        for i in range(4):
            frame_pdf.grid_columnconfigure(i, weight=1 if i % 2 != 0 else 0)

        # Powiadomienia o częściowych wynikach
        frame_partial = self.create_labeled_frame("Powiadomienie: Częściowo znalezione faktury (do 4 odbiorców)")
        
        for i in range(4):
            row_idx, col_idx = i // 2, (i % 2) * 2
            self.create_config_entry(
                frame_partial, 
                f"Odbiorca {i+1}:", 
                f"EMAIL_POWIADOMIENIA_CZESCIOWE_{i+1}", 
                row_idx, 
                col_idx, 
                width=30
            )
        
        ttk.Label(
            frame_partial, 
            text="(domyślnie: Użytkownik (Email), jeśli wszystkie puste)", 
            font=("Segoe UI", 8)
        ).grid(row=2, column=0, columnspan=4, sticky="w", padx=5, pady=(0,5))
        
        for i in range(4):
            frame_partial.grid_columnconfigure(i, weight=1 if i % 2 != 0 else 0)

    def _create_folders_section(self) -> None:
        """Tworzy sekcję ścieżek folderów."""
        frame = self.create_labeled_frame("Ścieżki Folderów")
        
        self.create_config_entry(frame, "Folder Przetworzonych:", "FOLDER_PRZETWORZONE_KOMPENSATY", 0)
        
        frame.grid_columnconfigure(1, weight=1)

    def _create_time_params_section(self) -> None:
        """Tworzy sekcję parametrów czasowych."""
        frame = self.create_labeled_frame("Parametry Czasowe Przeszukiwania")
        
        self.create_config_entry(frame, "Miesiące wstecz (główne emaile):", "MIESIACE_WSTECZ_KOMPENSATY", 0, is_int=True)
        self.create_config_entry(frame, "Miesiące wstecz (powiązane emaile):", "MIESIACE_WSTECZ_FAKTURY", 1, is_int=True)
        
        frame.grid_columnconfigure(1, weight=1)

    def load_from_config(self, config) -> None:
        """
        Ładuje wartości z konfiguracji.
        
        Args:
            config: Obiekt konfiguracji AppConfig
        """
        values = {
            "EXCHANGE_USER": config.exchange.username or "",
            "EXCHANGE_PASS": config.exchange.password or "",
            "EXCHANGE_SERVER_ADDR": config.exchange.server or "",
            "IGNORE_SSL_ERRORS": config.exchange.ignore_ssl,
            "EMAIL_KSIEGOWEJ_GLOWNY": config.email.ksiegowa_glowny,
            "EMAIL_KSIEGOWEJ_2": config.email.ksiegowa_2 or "",
            "EMAIL_KSIEGOWEJ_3": config.email.ksiegowa_3 or "",
            "FOLDER_PRZETWORZONE_KOMPENSATY": config.folders.przetworzone,
            "MIESIACE_WSTECZ_KOMPENSATY": config.search.miesiace_wstecz_kompensaty,
            "MIESIACE_WSTECZ_FAKTURY": config.search.miesiace_wstecz_faktury
        }

        # Dodaj adresy powiadomień
        for i in range(4):
            if i < len(config.email.powiadomienia_brak_pdf):
                values[f"EMAIL_POWIADOMIENIA_BRAK_PDF_{i+1}"] = config.email.powiadomienia_brak_pdf[i]
            else:
                values[f"EMAIL_POWIADOMIENIA_BRAK_PDF_{i+1}"] = ""
            
        for i in range(4):
            if i < len(config.email.powiadomienia_czesciowe):
                values[f"EMAIL_POWIADOMIENIA_CZESCIOWE_{i+1}"] = config.email.powiadomienia_czesciowe[i]
            else:
                values[f"EMAIL_POWIADOMIENIA_CZESCIOWE_{i+1}"] = ""

        self.set_all_values(values)

    def save_to_config(self, config) -> None:
        """
        Zapisuje wartości do konfiguracji.
        
        Args:
            config: Obiekt konfiguracji AppConfig do aktualizacji
        """
        values = self.get_all_values()
        
        config.exchange.username = values.get("EXCHANGE_USER", "")
        config.exchange.password = values.get("EXCHANGE_PASS", "")
        config.exchange.server = values.get("EXCHANGE_SERVER_ADDR", "")
        config.exchange.ignore_ssl = values.get("IGNORE_SSL_ERRORS", False)
        
        config.email.ksiegowa_glowny = values.get("EMAIL_KSIEGOWEJ_GLOWNY", "")
        config.email.ksiegowa_2 = values.get("EMAIL_KSIEGOWEJ_2", "")
        config.email.ksiegowa_3 = values.get("EMAIL_KSIEGOWEJ_3", "")
        
        config.folders.przetworzone = values.get("FOLDER_PRZETWORZONE_KOMPENSATY", "")
        
        config.search.miesiace_wstecz_kompensaty = values.get("MIESIACE_WSTECZ_KOMPENSATY", 8)
        config.search.miesiace_wstecz_faktury = values.get("MIESIACE_WSTECZ_FAKTURY", 4)
        
        # Aktualizuj listy powiadomień
        config.email.powiadomienia_brak_pdf = [
            values.get(f"EMAIL_POWIADOMIENIA_BRAK_PDF_{i}", "")
            for i in range(1, 5)
            if values.get(f"EMAIL_POWIADOMIENIA_BRAK_PDF_{i}", "").strip()
        ]
        
        config.email.powiadomienia_czesciowe = [
            values.get(f"EMAIL_POWIADOMIENIA_CZESCIOWE_{i}", "")
            for i in range(1, 5)
            if values.get(f"EMAIL_POWIADOMIENIA_CZESCIOWE_{i}", "").strip()
        ]