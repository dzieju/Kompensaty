"""
Moduł do komunikacji z serwerem Exchange.
"""
import logging
import traceback
from typing import Optional, List, Tuple
from datetime import datetime, timedelta

from exchangelib import Credentials, Configuration, Account, DELEGATE, FileAttachment, Message, EWSDateTime
from exchangelib.folders import Folder
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter

from config import AppConfig, ExchangeConfig
from exceptions import ExchangeConnectionError, FolderNotFoundError, OperationCancelledError
from models import SearchCriteria

class ExchangeClient:
    """Klient do komunikacji z serwerem Exchange."""
    
    def __init__(self, config: ExchangeConfig):
        """
        Inicjalizuje klienta Exchange.
        
        Args:
            config: Konfiguracja połączenia Exchange
            
        Raises:
            ExchangeConnectionError: Błąd połączenia z serwerem
        """
        self.config = config
        self.account: Optional[Account] = None
        self._setup_ssl_handling()
        self._connect()

    def _setup_ssl_handling(self) -> None:
        """Konfiguruje obsługę SSL."""
        if self.config.ignore_ssl:
            if BaseProtocol.HTTP_ADAPTER_CLS != NoVerifyHTTPAdapter:
                logging.warning("Włączono ignorowanie błędów SSL")
                BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter
        else:
            if BaseProtocol.HTTP_ADAPTER_CLS == NoVerifyHTTPAdapter:
                logging.info("Ignorowanie SSL wyłączone")

    def _connect(self) -> None:
        """
        Nawiązuje połączenie z serwerem Exchange.
        
        Raises:
            ExchangeConnectionError: Błąd połączenia
        """
        if not self.config.username or not self.config.password:
            raise ExchangeConnectionError("Brak danych logowania Exchange")

        try:
            credentials = Credentials(self.config.username, self.config.password)
            
            if self.config.server:
                config = Configuration(server=self.config.server, credentials=credentials)
                self.account = Account(
                    primary_smtp_address=self.config.username,
                    config=config,
                    autodiscover=False,
                    access_type=DELEGATE
                )
            else:
                self.account = Account(
                    primary_smtp_address=self.config.username,
                    credentials=credentials,
                    autodiscover=True,
                    access_type=DELEGATE
                )
            
            logging.info(f"Połączono z kontem: {self.account.primary_smtp_address}")
            
        except Exception as e:
            logging.error(f"Błąd połączenia z Exchange: {e}")
            logging.error(traceback.format_exc())
            raise ExchangeConnectionError(f"Nie udało się połączyć z Exchange: {e}")

    def get_folder_by_path(self, folder_path: str) -> Optional[Folder]:
        """
        Znajduje folder na podstawie ścieżki.
        
        Args:
            folder_path: Ścieżka do folderu (np. "Skrzynka odbiorcza/Kompensaty")
            
        Returns:
            Obiekt folderu lub None jeśli nie znaleziono
            
        Raises:
            FolderNotFoundError: Folder nie istnieje
        """
        if not folder_path:
            raise FolderNotFoundError("Pusta ścieżka folderu")

        path_parts = folder_path.split('/')
        current_folder = None

        # Określ folder startowy
        first_part = path_parts[0].lower()
        if first_part in ['inbox', 'skrzynka odbiorcza']:
            current_folder = self.account.inbox
            path_parts = path_parts[1:]
        elif first_part in ['sent items', 'elementy wysłane']:
            current_folder = self.account.sent
            path_parts = path_parts[1:]
        else:
            current_folder = self.account.root

        if not current_folder:
            raise FolderNotFoundError(f"Nie można określić folderu startowego dla: {folder_path}")

        # Nawiguj do docelowego folderu
        for part in path_parts:
            if not part:
                continue
                
            try:
                next_folder = next(
                    (child for child in current_folder.children 
                     if child.name.lower() == part.lower()), 
                    None
                )
                
                if next_folder:
                    current_folder = next_folder
                    logging.debug(f"Znaleziono podfolder '{part}' w '{current_folder.name}'")
                else:
                    available_folders = [child.name for child in current_folder.children]
                    logging.warning(f"Nie znaleziono podfolderu '{part}' w '{current_folder.name}'. "
                                  f"Dostępne: {available_folders}")
                    raise FolderNotFoundError(f"Folder '{part}' nie istnieje w '{current_folder.name}'")
                    
            except Exception as e:
                logging.error(f"Błąd nawigacji do '{part}': {e}")
                raise FolderNotFoundError(f"Błąd dostępu do folderu '{part}': {e}")

        logging.info(f"Uzyskano dostęp do folderu: {current_folder.name}")
        return current_folder

    def find_compensation_email(self, criteria: SearchCriteria, stop_event) -> Tuple[Optional[Message], Optional[FileAttachment]]:
        """
        Znajduje email z kompensatą zgodnie z kryteriami.
        
        Args:
            criteria: Kryteria wyszukiwania
            stop_event: Event do anulowania operacji
            
        Returns:
            Krotka (email, załącznik) lub (None, None)
            
        Raises:
            OperationCancelledError: Operacja anulowana
            FolderNotFoundError: Folder nie istnieje
        """
        if stop_event.is_set():
            raise OperationCancelledError("Anulowano wyszukiwanie emaila")

        folder = self.get_folder_by_path(criteria.main_email.folder_path)
        if not folder:
            raise FolderNotFoundError(f"Nie znaleziono folderu: {criteria.main_email.folder_path}")

        # Przygotuj zakres czasowy
        end_datetime = EWSDateTime.now(tz=self.account.default_timezone)
        start_datetime = end_datetime - timedelta(days=criteria.months_back_compensation * 31)
        
        logging.info(f"Wyszukiwanie emaili od {start_datetime.strftime('%Y-%m-%d %H:%M:%S %Z')} "
                    f"do {end_datetime.strftime('%Y-%m-%d %H:%M:%S %Z')}")

        # Przygotuj filtry
        filter_kwargs = {
            'datetime_received__gte': start_datetime,
            'datetime_received__lt': end_datetime
        }
        
        if criteria.main_email.require_attachments:
            filter_kwargs['has_attachments'] = True
            
        if criteria.main_email.only_unread:
            filter_kwargs['is_read'] = False
            logging.info("Wyszukiwanie tylko nieprzeczytanych emaili")

        try:
            emails = folder.filter(**filter_kwargs).order_by('-datetime_received')
            
            for email in emails:
                if stop_event.is_set():
                    raise OperationCancelledError("Anulowano wyszukiwanie")

                # Sprawdź kryteria tematu
                if (criteria.main_email.subject_contains and 
                    criteria.main_email.subject_contains.lower() not in email.subject.lower()):
                    continue

                # Sprawdź kryteria nadawcy
                if (criteria.main_email.sender_contains and 
                    criteria.main_email.sender_contains.lower() not in email.sender.email_address.lower()):
                    continue

                logging.debug(f"Sprawdzam email: '{email.subject}' od {email.sender.email_address}")

                # Sprawdź załączniki
                if not email.attachments:
                    continue

                for attachment in email.attachments:
                    if stop_event.is_set():
                        raise OperationCancelledError("Anulowano sprawdzanie załączników")
                        
                    if isinstance(attachment, FileAttachment):
                        if self._attachment_matches_criteria(attachment, criteria.main_attachment):
                            logging.info(f"Znaleziono email: '{email.subject}' z załącznikiem: {attachment.name}")
                            return email, attachment

            logging.info("Nie znaleziono emaila spełniającego kryteria")
            return None, None

        except OperationCancelledError:
            raise
        except Exception as e:
            logging.error(f"Błąd wyszukiwania emaili: {e}")
            logging.error(traceback.format_exc())
            return None, None

    def _attachment_matches_criteria(self, attachment: FileAttachment, criteria) -> bool:
        """
        Sprawdza czy załącznik spełnia kryteria.
        
        Args:
            attachment: Załącznik do sprawdzenia
            criteria: Kryteria załącznika
            
        Returns:
            True jeśli załącznik spełnia kryteria
        """
        att_name_lower = attachment.name.lower()
        
        # Sprawdź nazwę
        if (criteria.name_contains and 
            criteria.name_contains.lower() not in att_name_lower):
            return False
            
        # Sprawdź rozszerzenie
        if (criteria.extension and 
            not att_name_lower.endswith(f".{criteria.extension.lower()}")):
            return False
            
        return True

    def send_email(self, recipients: List[str], subject: str, body: str, 
                   attachments: List[FileAttachment] = None) -> bool:
        """
        Wysyła email.
        
        Args:
            recipients: Lista odbiorców
            subject: Temat emaila
            body: Treść emaila
            attachments: Lista załączników
            
        Returns:
            True jeśli email został wysłany pomyślnie
        """
        if not recipients:
            logging.error("Brak odbiorców")
            return False

        try:
            email = Message(
                account=self.account,
                subject=subject,
                body=body,
                to_recipients=recipients
            )

            if attachments:
                for attachment in attachments:
                    email.attach(attachment)
                    logging.info(f"Dołączono załącznik: {attachment.name}")

            email.send_and_save()
            logging.info(f"Email wysłany do: {', '.join(recipients)}")
            return True

        except Exception as e:
            logging.error(f"Błąd wysyłania emaila: {e}")
            logging.error(traceback.format_exc())
            return False

    def mark_as_read(self, email: Message) -> bool:
        """
        Oznacza email jako przeczytany.
        
        Args:
            email: Email do oznaczenia
            
        Returns:
            True jeśli operacja się powiodła
        """
        try:
            if not email.is_read:
                email.is_read = True
                email.save(update_fields=['is_read'])
                logging.info(f"Oznaczono email '{email.subject}' jako przeczytany")
            return True
        except Exception as e:
            logging.error(f"Błąd oznaczania emaila jako przeczytany: {e}")
            return False

    def move_email(self, email: Message, destination_folder_path: str) -> bool:
        """
        Przenosi email do innego folderu.
        
        Args:
            email: Email do przeniesienia
            destination_folder_path: Ścieżka docelowego folderu
            
        Returns:
            True jeśli operacja się powiodła
        """
        try:
            destination_folder = self.get_folder_by_path(destination_folder_path)
            if not destination_folder:
                logging.error(f"Nie znaleziono folderu docelowego: {destination_folder_path}")
                return False

            email.move(to_folder=destination_folder)
            logging.info(f"Email '{email.subject}' przeniesiony do {destination_folder.name}")
            return True

        except Exception as e:
            logging.error(f"Błąd przenoszenia emaila: {e}")
            return False