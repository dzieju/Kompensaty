"""
Moduł zarządzania konfiguracją aplikacji.
"""
import os
import logging
from typing import Optional, List
from dataclasses import dataclass, field
from dotenv import load_dotenv


@dataclass
class ExchangeConfig:
    """Konfiguracja połączenia z Exchange Server."""
    username: Optional[str] = None
    password: Optional[str] = None
    server: Optional[str] = None
    ignore_ssl: bool = False


@dataclass
class EmailConfig:
    """Konfiguracja adresów email."""
    ksiegowa_glowny: str = 'crypto@woox.pl'
    ksiegowa_2: Optional[str] = None
    ksiegowa_3: Optional[str] = None
    powiadomienia_brak_pdf: List[str] = field(default_factory=list)
    powiadomienia_czesciowe: List[str] = field(default_factory=list)


@dataclass
class FolderConfig:
    """Konfiguracja ścieżek folderów."""
    kompensaty: str = "Skrzynka odbiorcza/Kompensaty Quadra"
    faktury: str = "Skrzynka odbiorcza/Faktury"
    przetworzone: str = "Skrzynka odbiorcza/Kompensaty Quadra/Wyslane do weryfikacji"


@dataclass
class SearchConfig:
    """Konfiguracja parametrów wyszukiwania."""
    miesiace_wstecz_kompensaty: int = 8
    miesiace_wstecz_faktury: int = 4
    tylko_nieprzeczytane: bool = True
    wzorzec_numeru_faktury: str = r'\b(?:F/M|QN|F/)\s*[\w\s\/-]*\d[\w\s\/-]*\b'


@dataclass
class AppConfig:
    """Główna konfiguracja aplikacji."""
    exchange: ExchangeConfig = field(default_factory=ExchangeConfig)
    email: EmailConfig = field(default_factory=EmailConfig)
    folders: FolderConfig = field(default_factory=FolderConfig)
    search: SearchConfig = field(default_factory=SearchConfig)
    log_file: str = 'kompensata_processor.log'

    @classmethod
    def from_environment(cls, env_path: Optional[str] = None) -> 'AppConfig':
        """
        Tworzy konfigurację na podstawie zmiennych środowiskowych.
        
        Args:
            env_path: Ścieżka do pliku .env
            
        Returns:
            Obiekt konfiguracji
        """
        if env_path and os.path.exists(env_path):
            load_dotenv(dotenv_path=env_path, override=True)
            logging.info(f"Wczytano konfigurację z {env_path}")

        exchange_config = ExchangeConfig(
            username=os.environ.get('EXCHANGE_USER'),
            password=os.environ.get('EXCHANGE_PASS'),
            server=os.environ.get('EXCHANGE_SERVER_ADDR'),
            ignore_ssl=os.environ.get('IGNORE_SSL_ERRORS', 'false').lower() == 'true'
        )

        email_config = EmailConfig(
            ksiegowa_glowny=os.environ.get('EMAIL_KSIEGOWEJ_GLOWNY', 'crypto@woox.pl'),
            ksiegowa_2=os.environ.get('EMAIL_KSIEGOWEJ_2'),
            ksiegowa_3=os.environ.get('EMAIL_KSIEGOWEJ_3'),
            powiadomienia_brak_pdf=cls._get_email_list('EMAIL_POWIADOMIENIA_BRAK_PDF', 4),
            powiadomienia_czesciowe=cls._get_email_list('EMAIL_POWIADOMIENIA_CZESCIOWE', 4)
        )

        folder_config = FolderConfig(
            kompensaty=os.environ.get('FOLDER_KOMPENSATY', "Skrzynka odbiorcza/Kompensaty Quadra"),
            faktury=os.environ.get('FOLDER_FAKTURY', "Skrzynka odbiorcza/Faktury"),
            przetworzone=os.environ.get('FOLDER_PRZETWORZONE_KOMPENSATY', 
                                     "Skrzynka odbiorcza/Kompensaty Quadra/Wyslane do weryfikacji")
        )

        search_config = SearchConfig(
            miesiace_wstecz_kompensaty=int(os.environ.get('MIESIACE_WSTECZ_KOMPENSATY', '8')),
            miesiace_wstecz_faktury=int(os.environ.get('MIESIACE_WSTECZ_FAKTURY', '4')),
            tylko_nieprzeczytane=os.environ.get('SZUKAJ_TYLKO_NIEPRZECZYTANYCH', 'true').lower() == 'true',
            wzorzec_numeru_faktury=os.environ.get('WZORZEC_NUMERU_FAKTURY', 
                                                r'\b(?:F/M|QN|F/)\s*[\w\s\/-]*\d[\w\s\/-]*\b')
        )

        return cls(
            exchange=exchange_config,
            email=email_config,
            folders=folder_config,
            search=search_config,
            log_file=os.environ.get('LOG_FILE', 'kompensata_processor.log')
        )

    @staticmethod
    def _get_email_list(base_key: str, max_count: int) -> List[str]:
        """
        Pobiera listę adresów email ze zmiennych środowiskowych.
        
        Args:
            base_key: Bazowa nazwa klucza
            max_count: Maksymalna liczba adresów
            
        Returns:
            Lista adresów email
        """
        emails = []
        for i in range(1, max_count + 1):
            email = os.environ.get(f'{base_key}_{i}')
            if email and email.strip():
                emails.append(email.strip())
        return emails

    def get_ksiegowa_emails(self) -> List[str]:
        """Zwraca listę adresów email księgowej."""
        emails = [self.email.ksiegowa_glowny]
        if self.email.ksiegowa_2:
            emails.append(self.email.ksiegowa_2)
        if self.email.ksiegowa_3:
            emails.append(self.email.ksiegowa_3)
        return emails

    def get_notification_emails_for_missing_pdf(self) -> List[str]:
        """Zwraca listę adresów dla powiadomień o braku PDF."""
        if self.email.powiadomienia_brak_pdf:
            return self.email.powiadomienia_brak_pdf
        return [self.exchange.username] if self.exchange.username else []

    def get_notification_emails_for_partial_results(self) -> List[str]:
        """Zwraca listę adresów dla powiadomień o częściowych wynikach."""
        if self.email.powiadomienia_czesciowe:
            return self.email.powiadomienia_czesciowe
        return [self.exchange.username] if self.exchange.username else []