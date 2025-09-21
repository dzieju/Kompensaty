"""
Moduł z modelami danych używanymi w aplikacji.
"""
from dataclasses import dataclass, field
from typing import Optional, List, Dict, Any
from collections import defaultdict


@dataclass
class MainEmailCriteria:
    """Kryteria wyszukiwania głównego emaila."""
    folder_path: str = "Skrzynka odbiorcza/Kompensaty Quadra"
    subject_contains: str = ""
    sender_contains: str = ""
    only_unread: bool = True
    require_attachments: bool = True


@dataclass
class AttachmentCriteria:
    """Kryteria załącznika."""
    name_contains: str = "kompensata"
    extension: str = "pdf"


@dataclass
class ExtractionCriteria:
    """Kryteria ekstrakcji danych z załącznika."""
    regex_pattern: str = r'\b(?:F/M|QN|F/)\s*[\w\s\/-]*\d[\w\s\/-]*\b'


@dataclass
class RelatedEmailCriteria:
    """Kryteria wyszukiwania powiązanych emaili."""
    folder_path: str = "Skrzynka odbiorcza/Faktury"
    file_extension: str = "pdf"


@dataclass
class SearchCriteria:
    """Kompletne kryteria wyszukiwania."""
    main_email: MainEmailCriteria = field(default_factory=MainEmailCriteria)
    main_attachment: AttachmentCriteria = field(default_factory=AttachmentCriteria)
    extraction: ExtractionCriteria = field(default_factory=ExtractionCriteria)
    related_emails: RelatedEmailCriteria = field(default_factory=RelatedEmailCriteria)
    months_back_compensation: int = 8
    months_back_invoices: int = 4

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'SearchCriteria':
        """
        Tworzy obiekt SearchCriteria z słownika.
        
        Args:
            data: Słownik z danymi
            
        Returns:
            Obiekt SearchCriteria
        """
        main_email_data = data.get("glowny_email", {})
        main_email = MainEmailCriteria(
            folder_path=main_email_data.get("folder_sciezka", "Skrzynka odbiorcza/Kompensaty Quadra"),
            subject_contains=main_email_data.get("temat_zawiera", ""),
            sender_contains=main_email_data.get("nadawca_zawiera", ""),
            only_unread=main_email_data.get("tylko_nieprzeczytane", True),
            require_attachments=main_email_data.get("wymagane_zalaczniki", True)
        )

        main_attachment_data = data.get("glowny_zalacznik", {})
        main_attachment = AttachmentCriteria(
            name_contains=main_attachment_data.get("nazwa_zawiera", "kompensata"),
            extension=main_attachment_data.get("rozszerzenie", "pdf")
        )

        extraction_data = data.get("ekstrakcja_z_zalacznika", {})
        extraction = ExtractionCriteria(
            regex_pattern=extraction_data.get("wzorzec_regex", 
                                            r'\b(?:F/M|QN|F/)\s*[\w\s\/-]*\d[\w\s\/-]*\b')
        )

        related_data = data.get("powiazane_emaile_faktury", {})
        related_emails = RelatedEmailCriteria(
            folder_path=related_data.get("folder_sciezka", "Skrzynka odbiorcza/Faktury"),
            file_extension=related_data.get("rozszerzenie_pliku", "pdf")
        )

        return cls(
            main_email=main_email,
            main_attachment=main_attachment,
            extraction=extraction,
            related_emails=related_emails,
            months_back_compensation=data.get("miesiace_wstecz_kompensaty", 8),
            months_back_invoices=data.get("miesiace_wstecz_faktury", 4)
        )

    def to_dict(self) -> Dict[str, Any]:
        """
        Konwertuje obiekt do słownika.
        
        Returns:
            Słownik z danymi
        """
        return {
            "glowny_email": {
                "folder_sciezka": self.main_email.folder_path,
                "temat_zawiera": self.main_email.subject_contains,
                "nadawca_zawiera": self.main_email.sender_contains,
                "tylko_nieprzeczytane": self.main_email.only_unread,
                "wymagane_zalaczniki": self.main_email.require_attachments
            },
            "glowny_zalacznik": {
                "nazwa_zawiera": self.main_attachment.name_contains,
                "rozszerzenie": self.main_attachment.extension
            },
            "ekstrakcja_z_zalacznika": {
                "wzorzec_regex": self.extraction.regex_pattern
            },
            "powiazane_emaile_faktury": {
                "folder_sciezka": self.related_emails.folder_path,
                "rozszerzenie_pliku": self.related_emails.file_extension
            },
            "miesiace_wstecz_kompensaty": self.months_back_compensation,
            "miesiace_wstecz_faktury": self.months_back_invoices
        }


@dataclass
class ProcessingResult:
    """Wynik przetwarzania emaila."""
    success: bool
    message: str
    email_subject: Optional[str] = None
    extracted_data: List[str] = field(default_factory=list)
    found_attachments_count: int = 0
    missing_invoices: List[str] = field(default_factory=list)


@dataclass
class EmailTemplate:
    """Szablon emaila."""
    subject: str
    body: str

    def format(self, **kwargs) -> 'EmailTemplate':
        """
        Formatuje szablon z podanymi argumentami.
        
        Args:
            **kwargs: Argumenty do formatowania
            
        Returns:
            Sformatowany szablon
        """
        # Użyj defaultdict żeby uniknąć KeyError dla brakujących kluczy
        format_dict = defaultdict(lambda: '', kwargs)
        
        # Specjalne formatowanie dla list
        if 'numery_faktur_z_kompensaty' in kwargs and isinstance(kwargs['numery_faktur_z_kompensaty'], list):
            format_dict['numery_faktur_z_kompensaty_lista'] = "\n".join([
                f"  - {nr}" for nr in kwargs['numery_faktur_z_kompensaty']
            ])
        
        if 'numery_brakujace' in kwargs and kwargs['numery_brakujace']:
            brakujace_lista = "\n".join([f"  - {nr}" for nr in kwargs['numery_brakujace']])
            format_dict['brakujace_faktury_info'] = f"\nNie udało się odnaleźć następujących faktur:\n{brakujace_lista}"
        else:
            format_dict['brakujace_faktury_info'] = ""

        # Dodaj nazwę użytkownika jeśli nie podana
        if 'nazwa_uzytkownika' not in format_dict and 'exchange_username' in kwargs:
            username = kwargs['exchange_username']
            format_dict['nazwa_uzytkownika'] = username.split('@')[0] if '@' in username else username

        return EmailTemplate(
            subject=self.subject.format_map(format_dict),
            body=self.body.format_map(format_dict)
        )


@dataclass
class EmailTemplateCollection:
    """Kolekcja szablonów emaili."""
    email_do_ksiegowej: EmailTemplate
    powiadomienie_brak_faktur_w_pdf: EmailTemplate
    powiadomienie_czesciowo_znalezione: EmailTemplate

    @classmethod
    def get_default(cls) -> 'EmailTemplateCollection':
        """
        Zwraca domyślne szablony emaili.
        
        Returns:
            Kolekcja domyślnych szablonów
        """
        return cls(
            email_do_ksiegowej=EmailTemplate(
                subject="Kompensata {temat_kompensaty} - faktury ({liczba_znalezionych_faktur_zal}/{liczba_faktur_z_kompensaty})",
                body=(
                    "Dzień dobry,\n\n"
                    "W załączeniu przesyłam dokument kompensaty (z e-maila '{temat_kompensaty}') "
                    "oraz odnalezione na skrzynce faktury.\n\n"
                    "Numery faktur odczytane z dokumentu kompensaty ({liczba_faktur_z_kompensaty}):\n"
                    "{numery_faktur_z_kompensaty_lista}\n\n"
                    "Załączono {liczba_znalezionych_faktur_zal} plików faktur.\n"
                    "{brakujace_faktury_info}\n\n"
                    "Z poważaniem,\n"
                    "Automatyczny Asystent Skrzynki Pocztowej"
                )
            ),
            powiadomienie_brak_faktur_w_pdf=EmailTemplate(
                subject="Problem z ekstrakcją faktur z kompensaty: {temat_kompensaty}",
                body=(
                    "Witaj {nazwa_uzytkownika},\n\n"
                    "Automat próbował przetworzyć email z kompensatą, ale nie udało się "
                    "wyekstrahować żadnych numerów faktur z załączonego dokumentu PDF.\n\n"
                    "Dotyczy to emaila:\n"
                    "  Temat: {temat_kompensaty}\n"
                    "  Nadawca: {nadawca_kompensaty}\n"
                    "  Data otrzymania: {data_kompensaty}\n"
                    "  Nazwa pliku PDF z kompensatą: {nazwa_pliku_pdf_kompensaty}\n\n"
                    "Prosimy o ręczne sprawdzenie tego dokumentu.\n\n"
                    "Z poważaniem,\n"
                    "Automatyczny Asystent Skrzynki Pocztowej"
                )
            ),
            powiadomienie_czesciowo_znalezione=EmailTemplate(
                subject="UWAGA: Nie wszystkie faktury znalezione dla kompensaty '{temat_kompensaty}'",
                body=(
                    "Cześć {nazwa_uzytkownika},\n\n"
                    "Automat przetworzył email z kompensatą:\n"
                    "  - Temat: '{temat_kompensaty}'\n"
                    "  - Otrzymano: {data_kompensaty}\n\n"
                    "Z dokumentu PDF wyciągnięto {liczba_oczekiwanych} numerów faktur:\n"
                    "{numery_faktur_z_kompensaty_lista}\n\n"
                    "Udało się odnaleźć {liczba_znalezionych} odpowiadających im plików faktur.\n"
                    "{brakujace_faktury_info}\n\n"
                    "Email z kompensatą i znalezionymi fakturami został wysłany do księgowej.\n"
                    "Prosimy o ewentualne sprawdzenie brakujących pozycji.\n\n"
                    "Z poważaniem,\n"
                    "Automatyczny Asystent Skrzynki Pocztowej"
                )
            )
        )

    def get_template(self, template_name: str) -> Optional[EmailTemplate]:
        """
        Zwraca szablon o podanej nazwie.
        
        Args:
            template_name: Nazwa szablonu
            
        Returns:
            Szablon lub None jeśli nie znaleziono
        """
        return getattr(self, template_name, None)

    @classmethod
    def from_dict(cls, data: Dict[str, Dict[str, str]]) -> 'EmailTemplateCollection':
        """
        Tworzy kolekcję z słownika.
        
        Args:
            data: Słownik z danymi szablonów
            
        Returns:
            Kolekcja szablonów
        """
        default = cls.get_default()
        
        email_do_ksiegowej_data = data.get("email_do_ksiegowej", {})
        powiadomienie_brak_data = data.get("powiadomienie_brak_faktur_w_pdf", {})
        powiadomienie_czesciowo_data = data.get("powiadomienie_czesciowo_znalezione", {})
        
        return cls(
            email_do_ksiegowej=EmailTemplate(
                subject=email_do_ksiegowej_data.get("subject", default.email_do_ksiegowej.subject),
                body=email_do_ksiegowej_data.get("body", default.email_do_ksiegowej.body)
            ),
            powiadomienie_brak_faktur_w_pdf=EmailTemplate(
                subject=powiadomienie_brak_data.get("subject", default.powiadomienie_brak_faktur_w_pdf.subject),
                body=powiadomienie_brak_data.get("body", default.powiadomienie_brak_faktur_w_pdf.body)
            ),
            powiadomienie_czesciowo_znalezione=EmailTemplate(
                subject=powiadomienie_czesciowo_data.get("subject", default.powiadomienie_czesciowo_znalezione.subject),
                body=powiadomienie_czesciowo_data.get("body", default.powiadomienie_czesciowo_znalezione.body)
            )
        )

    def to_dict(self) -> Dict[str, Dict[str, str]]:
        """
        Konwertuje kolekcję do słownika.
        
        Returns:
            Słownik z danymi
        """
        return {
            "email_do_ksiegowej": {
                "subject": self.email_do_ksiegowej.subject,
                "body": self.email_do_ksiegowej.body
            },
            "powiadomienie_brak_faktur_w_pdf": {
                "subject": self.powiadomienie_brak_faktur_w_pdf.subject,
                "body": self.powiadomienie_brak_faktur_w_pdf.body
            },
            "powiadomienie_czesciowo_znalezione": {
                "subject": self.powiadomienie_czesciowo_znalezione.subject,
                "body": self.powiadomienie_czesciowo_znalezione.body
            }
        }