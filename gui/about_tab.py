"""
Zakładka informacji o programie.
"""
import tkinter as tk
from tkinter import ttk
from gui.base_tab import BaseTab


class AboutTab(BaseTab):
    """Zakładka z informacjami o programie."""
    
    def setup_ui(self) -> None:
        """Konfiguruje interfejs użytkownika."""
        # Główna ramka z paddingiem
        main_frame = ttk.Frame(self.parent_frame, padding="30")
        main_frame.pack(expand=True, fill=tk.BOTH)
        
        # Tytuł aplikacji
        title_label = ttk.Label(
            main_frame,
            text="Automat Poczty PDF Filter",
            font=("Segoe UI", 20, "bold")
        )
        title_label.pack(pady=(0, 20))
        
        # Opis aplikacji
        description_text = (
            "Automatyczne przetwarzanie emaili z kompensatami i fakturami.\n"
            "Aplikacja łączy się z serwerem Exchange, wyszukuje emaile z kompensatami,\n"
            "ekstraktuje numery faktur z PDF-ów i automatycznie wysyła raporty do księgowości."
        )
        description_label = ttk.Label(
            main_frame,
            text=description_text,
            font=("Segoe UI", 10),
            justify=tk.CENTER,
            wraplength=500
        )
        description_label.pack(pady=(0, 30))
        
        # Ramka z informacjami o autorze
        author_frame = ttk.LabelFrame(main_frame, text="Informacje o autorze", padding=20)
        author_frame.pack(fill="x", pady=10)
        
        # Informacje o autorze
        author_info = [
            ("Autor:", "Grzegorz Ciekot"),
            ("Email:", "grzegorz.ciekot@woox.pl"),
            ("Telefon:", "34 363 28 68"),
            ("Firma:", "Woox Technologies")
        ]
        
        for i, (label, value) in enumerate(author_info):
            ttk.Label(
                author_frame,
                text=label,
                font=("Segoe UI", 10, "bold")
            ).grid(row=i, column=0, sticky="w", padx=(0, 10), pady=5)
            
            ttk.Label(
                author_frame,
                text=value,
                font=("Segoe UI", 10)
            ).grid(row=i, column=1, sticky="w", pady=5)
        
        # Ramka z informacjami technicznymi
        tech_frame = ttk.LabelFrame(main_frame, text="Informacje techniczne", padding=20)
        tech_frame.pack(fill="x", pady=10)
        
        # Informacje techniczne
        tech_info = [
            ("Wersja:", "3.0"),
            ("Data wydania:", "2024-12-19"),
            ("Python:", "3.8+"),
            ("GUI:", "Tkinter"),
            ("Exchange:", "exchangelib"),
            ("PDF:", "PyPDF2")
        ]
        
        for i, (label, value) in enumerate(tech_info):
            ttk.Label(
                tech_frame,
                text=label,
                font=("Segoe UI", 10, "bold")
            ).grid(row=i, column=0, sticky="w", padx=(0, 10), pady=5)
            
            ttk.Label(
                tech_frame,
                text=value,
                font=("Segoe UI", 10)
            ).grid(row=i, column=1, sticky="w", pady=5)
        
        # Ramka z funkcjonalnościami
        features_frame = ttk.LabelFrame(main_frame, text="Główne funkcjonalności", padding=20)
        features_frame.pack(fill="x", pady=10)
        
        features_text = (
            "• Automatyczne wyszukiwanie emaili z kompensatami\n"
            "• Ekstraktowanie numerów faktur z plików PDF\n"
            "• Wyszukiwanie powiązanych faktur w skrzynce\n"
            "• Automatyczne wysyłanie raportów do księgowości\n"
            "• System powiadomień o problemach\n"
            "• Konfigurowalne szablony emaili\n"
            "• Zaawansowane kryteria wyszukiwania\n"
            "• Logowanie i monitorowanie operacji"
        )
        
        ttk.Label(
            features_frame,
            text=features_text,
            font=("Segoe UI", 10),
            justify=tk.LEFT
        ).pack(anchor="w")
        
        # Stopka z prawami autorskimi
        copyright_label = ttk.Label(
            main_frame,
            text="© 2024 Woox Technologies. Wszystkie prawa zastrzeżone.",
            font=("Segoe UI", 8, "italic")
        )
        copyright_label.pack(side=tk.BOTTOM, pady=(30, 0))

    def get_version_info(self) -> dict:
        """
        Zwraca informacje o wersji aplikacji.
        
        Returns:
            Słownik z informacjami o wersji
        """
        return {
            "version": "3.0",
            "release_date": "2024-12-19",
            "author": "Grzegorz Ciekot",
            "company": "Woox Technologies",
            "python_required": "3.8+",
            "main_dependencies": [
                "tkinter", "exchangelib", "PyPDF2", "python-dotenv"
            ]
        }