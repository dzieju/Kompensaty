"""
Zakładka szablonów emaili.
"""
import tkinter as tk
from tkinter import ttk, scrolledtext
from gui.base_tab import BaseTab
from models import EmailTemplateCollection


class TemplatesTab(BaseTab):
    """Zakładka szablonów emaili."""
    
    def __init__(self, parent_frame: ttk.Frame):
        """
        Inicjalizuje zakładkę szablonów.
        
        Args:
            parent_frame: Ramka rodzica
        """
        self.template_widgets = {}
        super().__init__(parent_frame)

    def setup_ui(self) -> None:
        """Konfiguruje interfejs użytkownika."""
        self._create_template_sections()
        self._create_placeholders_section()

    def _create_template_sections(self) -> None:
        """Tworzy sekcje dla każdego szablonu."""
        templates = [
            ("email_do_ksiegowej", "Email do Księgowej"),
            ("powiadomienie_brak_faktur_w_pdf", "Powiadomienie: Brak faktur w PDF"),
            ("powiadomienie_czesciowo_znalezione", "Powiadomienie: Częściowo znalezione faktury")
        ]
        
        for template_key, template_label in templates:
            self._create_single_template_section(template_key, template_label)

    def _create_single_template_section(self, key: str, label: str) -> None:
        """
        Tworzy sekcję dla pojedynczego szablonu.
        
        Args:
            key: Klucz szablonu
            label: Etykieta szablonu
        """
        frame = self.create_labeled_frame(label)
        
        # Temat
        ttk.Label(frame, text="Temat:").grid(row=0, column=0, sticky="nw", padx=5, pady=2)
        subject_var = tk.StringVar()
        subject_entry = ttk.Entry(frame, textvariable=subject_var, width=80)
        subject_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        
        # Treść
        ttk.Label(frame, text="Treść:").grid(row=1, column=0, sticky="nw", padx=5, pady=2)
        body_text = scrolledtext.ScrolledText(
            frame, 
            wrap=tk.WORD, 
            height=8, 
            width=80, 
            font=("Segoe UI", 9)
        )
        body_text.grid(row=1, column=1, sticky="nsew", padx=5, pady=2)
        
        frame.grid_columnconfigure(1, weight=1)
        frame.grid_rowconfigure(1, weight=1)
        
        self.template_widgets[key] = {
            "subject_var": subject_var,
            "body_text_widget": body_text
        }

    def _create_placeholders_section(self) -> None:
        """Tworzy sekcję z dostępnymi placeholderami."""
        frame = self.create_labeled_frame("Dostępne Placeholdery (do użycia w temacie i treści)")
        
        placeholders_text = (
            "Ogólne:\n"
            "  {nazwa_uzytkownika} - nazwa użytkownika\n"
            "  {temat_kompensaty} - temat emaila z kompensatą\n"
            "  {nadawca_kompensaty} - adres nadawcy kompensaty\n"
            "  {data_kompensaty} - data otrzymania kompensaty\n"
            "  {nazwa_pliku_pdf_kompensaty} - nazwa pliku PDF\n\n"
            "Statystyki:\n"
            "  {liczba_faktur_z_kompensaty} - liczba numerów w PDF\n"
            "  {numery_faktur_z_kompensaty_lista} - lista numerów (automatyczne formatowanie)\n"
            "  {liczba_znalezionych_faktur_zal} - liczba znalezionych załączników\n"
            "  {liczba_oczekiwanych} - liczba oczekiwanych faktur\n"
            "  {liczba_znalezionych} - liczba faktycznie znalezionych\n"
            "  {brakujace_faktury_info} - informacja o brakujących (automatyczne formatowanie)\n"
            "  {numery_brakujace} - lista brakujących numerów"
        )
        
        # UPROSZCZONY TEXT WIDGET BEZ OPCJI KOLORÓW
        text_widget = tk.Text(
            frame, 
            height=12, 
            wrap=tk.WORD, 
            font=("Consolas", 9),
            relief="sunken",
            state="disabled"
        )
        text_widget.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Wstaw tekst
        text_widget.config(state="normal")
        text_widget.insert("1.0", placeholders_text)
        text_widget.config(state="disabled")

    def load_from_templates(self, templates: EmailTemplateCollection) -> None:
        """
        Ładuje szablony do widgetów.
        
        Args:
            templates: Kolekcja szablonów
        """
        template_data = {
            "email_do_ksiegowej": templates.email_do_ksiegowej,
            "powiadomienie_brak_faktur_w_pdf": templates.powiadomienie_brak_faktur_w_pdf,
            "powiadomienie_czesciowo_znalezione": templates.powiadomienie_czesciowo_znalezione
        }
        
        for key, template in template_data.items():
            if key in self.template_widgets:
                widgets = self.template_widgets[key]
                widgets["subject_var"].set(template.subject)
                widgets["body_text_widget"].delete('1.0', tk.END)
                widgets["body_text_widget"].insert('1.0', template.body)

    def save_to_templates(self) -> EmailTemplateCollection:
        """
        Tworzy kolekcję szablonów z wartości z GUI.
        
        Returns:
            Kolekcja szablonów
        """
        templates_dict = {}
        
        for key, widgets in self.template_widgets.items():
            templates_dict[key] = {
                "subject": widgets["subject_var"].get(),
                "body": widgets["body_text_widget"].get("1.0", tk.END).strip()
            }
        
        return EmailTemplateCollection.from_dict(templates_dict)

    def reset_to_defaults(self) -> None:
        """Resetuje szablony do wartości domyślnych."""
        default_templates = EmailTemplateCollection.get_default()
        self.load_from_templates(default_templates)