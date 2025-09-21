"""
Moduł serwisu emaili - łączy Exchange z szablonami i logiką biznesową.
"""
import logging
import traceback
from typing import List, Optional, Tuple
from datetime import datetime, timedelta
import re

from exchangelib import FileAttachment, EWSDateTime

from config import AppConfig
from models import SearchCriteria, EmailTemplateCollection, ProcessingResult
from exchange_client import ExchangeClient
from pdf_processor import PDFProcessor
from exceptions import OperationCancelledError, EmailNotFoundError


class EmailService:
    """Serwis do zarządzania operacjami emailowymi."""
    
    def __init__(self, config: AppConfig, templates: EmailTemplateCollection):
        """
        Inicjalizuje serwis emaili.
        
        Args:
            config: Konfiguracja aplikacji
            templates: Szablony emaili
        """
        self.config = config
        self.templates = templates
        self.exchange_client = ExchangeClient(config.exchange)
        self.pdf_processor = PDFProcessor()

    def find_related_invoices(self, extracted_data: List[str], reference_date: datetime,
                             criteria: SearchCriteria, stop_event) -> List[FileAttachment]:
        """
        Znajduje powiązane faktury na podstawie wyekstrahowanych danych.
        
        Args:
            extracted_data: Lista wyekstrahowanych numerów faktur
            reference_date: Data referencyjna (data głównego emaila)
            criteria: Kryteria wyszukiwania
            stop_event: Event do anulowania operacji
            
        Returns:
            Lista znalezionych załączników faktur
            
        Raises:
            OperationCancelledError: Operacja anulowana
        """
        if stop_event.is_set():
            raise OperationCancelledError("Anulowano wyszukiwanie załączników faktur")

        found_attachments = []
        
        if not extracted_data:
            return found_attachments

        # Pobierz folder z fakturami
        folder = self.exchange_client.get_folder_by_path(criteria.related_emails.folder_path)
        if not folder:
            logging.error(f"Nie znaleziono folderu faktur: {criteria.related_emails.folder_path}")
            return found_attachments

        # Przygotuj zakres czasowy
        if isinstance(reference_date, datetime) and reference_date.tzinfo is None:
            ref_date_aware = EWSDateTime.combine(
                reference_date.date(), 
                reference_date.time(), 
                tzinfo=self.exchange_client.account.default_timezone
            )
        elif not isinstance(reference_date, EWSDateTime):
            ref_date_aware = EWSDateTime.from_datetime(reference_date)
        else:
            ref_date_aware = reference_date

        end_datetime = ref_date_aware + timedelta(days=2)
        start_datetime = ref_date_aware - timedelta(days=criteria.months_back_invoices * 31)
        
        logging.info(f"Wyszukiwanie powiązanych emaili od {start_datetime.strftime('%Y-%m-%d %H:%M:%S %Z')} "
                    f"do {end_datetime.strftime('%Y-%m-%d %H:%M:%S %Z')}")

        # Szukaj dla każdego numeru faktury
        for invoice_number in extracted_data:
            if stop_event.is_set():
                raise OperationCancelledError("Anulowano wyszukiwanie powiązanego pliku")

            attachment = self._find_single_invoice_attachment(
                folder, invoice_number, start_datetime, end_datetime, 
                criteria.related_emails.file_extension, stop_event
            )
            
            if attachment:
                found_attachments.append(attachment)

        logging.info(f"Znaleziono łącznie {len(found_attachments)} załączników powiązanych")
        return found_attachments

    def _find_single_invoice_attachment(self, folder, invoice_number: str, 
                                       start_datetime: EWSDateTime, end_datetime: EWSDateTime,
                                       file_extension: str, stop_event) -> Optional[FileAttachment]:
        """
        Znajduje pojedynczy załącznik faktury.
        
        Args:
            folder: Folder do przeszukania
            invoice_number: Numer faktury do znalezienia
            start_datetime: Początek zakresu czasowego
            end_datetime: Koniec zakresu czasowego
            file_extension: Rozszerzenie szukanego pliku
            stop_event: Event do anulowania
            
        Returns:
            Znaleziony załącznik lub None
        """
        # Przygotuj fragment nazwy do szukania
        search_fragment = re.sub(r'[^\w\d.-]', '_', invoice_number).lower()
        
        logging.info(f"Szukam powiązanego pliku dla faktury: '{invoice_number}' "
                    f"(fragment: '{search_fragment}')")

        try:
            emails = folder.filter(
                has_attachments=True,
                datetime_received__gte=start_datetime,
                datetime_received__lt=end_datetime
            ).order_by('-datetime_received')

            for email in emails:
                if stop_event.is_set():
                    raise OperationCancelledError("Anulowano przeszukiwanie emaili")

                logging.debug(f"Sprawdzam email: '{email.subject}'")
                
                if not email.attachments:
                    continue

                for attachment in email.attachments:
                    if stop_event.is_set():
                        raise OperationCancelledError("Anulowano sprawdzanie załączników")

                    if isinstance(attachment, FileAttachment):
                        att_name_lower = attachment.name.lower()
                        
                        if (att_name_lower.endswith(f".{file_extension.lower()}") and
                            search_fragment in att_name_lower):
                            
                            logging.info(f"Znaleziono załącznik: '{attachment.name}' "
                                       f"w emailu '{email.subject}'")
                            return attachment

            logging.warning(f"Nie znaleziono pliku dla faktury: '{invoice_number}'")
            return None

        except OperationCancelledError:
            raise
        except Exception as e:
            logging.error(f"Błąd wyszukiwania pliku dla faktury '{invoice_number}': {e}")
            return None

    def send_compensation_email_to_accounting(self, original_pdf: FileAttachment,
                                            found_invoices: List[FileAttachment],
                                            extracted_data: List[str],
                                            original_email, stop_event) -> bool:
        """
        Wysyła email z kompensatą do księgowości.
        
        Args:
            original_pdf: Oryginalny PDF z kompensatą
            found_invoices: Znalezione faktury
            extracted_data: Wyekstrahowane numery faktur
            original_email: Oryginalny email z kompensatą
            stop_event: Event do anulowania
            
        Returns:
            True jeśli email został wysłany pomyślnie
        """
        if stop_event.is_set():
            raise OperationCancelledError("Anulowano tworzenie emaila do księgowej")

        try:
            recipients = self.config.get_ksiegowa_emails()
            if not recipients:
                logging.error("Brak zdefiniowanych adresów email dla księgowości")
                return False

            # Przygotuj dane do szablonu
            template_data = {
                "temat_kompensaty": original_email.subject,
                "liczba_faktur_z_kompensaty": len(extracted_data),
                "numery_faktur_z_kompensaty": extracted_data,
                "liczba_znalezionych_faktur_zal": len(found_invoices),
                "exchange_username": self.config.exchange.username
            }

            # Dodaj informacje o brakujących fakturach
            if len(found_invoices) < len(extracted_data):
                missing_invoices = self._find_missing_invoices(extracted_data, found_invoices)
                template_data["numery_brakujace"] = missing_invoices

            # Formatuj szablon
            formatted_template = self.templates.email_do_ksiegowej.format(**template_data)

            # Przygotuj załączniki
            attachments = [original_pdf]
            added_names = {original_pdf.name.lower()}

            for invoice_attachment in found_invoices:
                if stop_event.is_set():
                    raise OperationCancelledError("Anulowano dołączanie załączników")

                if invoice_attachment.name.lower() not in added_names:
                    attachments.append(invoice_attachment)
                    added_names.add(invoice_attachment.name.lower())

            # Wyślij email
            success = self.exchange_client.send_email(
                recipients=recipients,
                subject=formatted_template.subject,
                body=formatted_template.body,
                attachments=attachments
            )

            if success:
                logging.info(f"Email do księgowej wysłany pomyślnie do: {', '.join(recipients)}")
            
            return success

        except OperationCancelledError:
            raise
        except Exception as e:
            logging.error(f"Błąd wysyłania emaila do księgowej: {e}")
            logging.error(traceback.format_exc())
            return False

    def send_missing_pdf_notification(self, original_email, pdf_filename: str, stop_event) -> bool:
        """
        Wysyła powiadomienie o problemie z ekstrakcją PDF.
        
        Args:
            original_email: Oryginalny email z kompensatą
            pdf_filename: Nazwa pliku PDF
            stop_event: Event do anulowania
            
        Returns:
            True jeśli powiadomienie zostało wysłane
        """
        if stop_event.is_set():
            raise OperationCancelledError("Anulowano wysyłanie powiadomienia")

        recipients = self.config.get_notification_emails_for_missing_pdf()
        if not recipients:
            logging.error("Brak odbiorców dla powiadomienia o problemie z PDF")
            return False

        try:
            template_data = {
                "temat_kompensaty": original_email.subject,
                "nadawca_kompensaty": original_email.sender.email_address,
                "data_kompensaty": original_email.datetime_received.strftime('%Y-%m-%d %H:%M:%S %Z'),
                "nazwa_pliku_pdf_kompensaty": pdf_filename,
                "exchange_username": self.config.exchange.username
            }

            formatted_template = self.templates.powiadomienie_brak_faktur_w_pdf.format(**template_data)

            success = self.exchange_client.send_email(
                recipients=recipients,
                subject=formatted_template.subject,
                body=formatted_template.body
            )

            if success:
                logging.info(f"Powiadomienie o problemie z PDF wysłane do: {', '.join(recipients)}")

            return success

        except OperationCancelledError:
            raise
        except Exception as e:
            logging.error(f"Błąd wysyłania powiadomienia o problemie z PDF: {e}")
            return False

    def send_partial_results_notification(self, original_email, extracted_data: List[str],
                                        found_attachments: List[FileAttachment], stop_event) -> bool:
        """
        Wysyła powiadomienie o częściowo znalezionych fakturach.
        
        Args:
            original_email: Oryginalny email z kompensatą
            extracted_data: Wyekstrahowane numery faktur
            found_attachments: Znalezione załączniki
            stop_event: Event do anulowania
            
        Returns:
            True jeśli powiadomienie zostało wysłane
        """
        if stop_event.is_set():
            raise OperationCancelledError("Anulowano wysyłanie powiadomienia")

        recipients = self.config.get_notification_emails_for_partial_results()
        if not recipients:
            logging.error("Brak odbiorców dla powiadomienia o częściowych wynikach")
            return False

        try:
            missing_invoices = self._find_missing_invoices(extracted_data, found_attachments)

            template_data = {
                "temat_kompensaty": original_email.subject,
                "data_kompensaty": original_email.datetime_received.strftime('%Y-%m-%d %H:%M'),
                "liczba_oczekiwanych": len(extracted_data),
                "numery_faktur_z_kompensaty": extracted_data,
                "liczba_znalezionych": len(found_attachments),
                "numery_brakujace": missing_invoices,
                "exchange_username": self.config.exchange.username
            }

            formatted_template = self.templates.powiadomienie_czesciowo_znalezione.format(**template_data)

            success = self.exchange_client.send_email(
                recipients=recipients,
                subject=formatted_template.subject,
                body=formatted_template.body
            )

            if success:
                logging.info(f"Powiadomienie o częściowych wynikach wysłane do: {', '.join(recipients)}")

            return success

        except OperationCancelledError:
            raise
        except Exception as e:
            logging.error(f"Błąd wysyłania powiadomienia o częściowych wynikach: {e}")
            return False

    def _find_missing_invoices(self, extracted_data: List[str], 
                              found_attachments: List[FileAttachment]) -> List[str]:
        """
        Znajduje numery faktur, których nie udało się odnaleźć.
        
        Args:
            extracted_data: Lista wyekstrahowanych numerów
            found_attachments: Lista znalezionych załączników
            
        Returns:
            Lista brakujących numerów faktur
        """
        found_names_lower = {att.name.lower() for att in found_attachments}
        
        missing = []
        for invoice_number in extracted_data:
            search_fragment = re.sub(r'[^\w\d.-]', '_', invoice_number).lower()
            if not any(search_fragment in filename for filename in found_names_lower):
                missing.append(invoice_number)
                
        return missing

    def move_email_to_processed_folder(self, email, stop_event) -> bool:
        """
        Przenosi email do folderu przetworzonych.
        
        Args:
            email: Email do przeniesienia
            stop_event: Event do anulowania
            
        Returns:
            True jeśli operacja się powiodła
        """
        try:
            success = self.exchange_client.move_email(email, self.config.folders.przetworzone)
            if not success:
                # Fallback - oznacz jako przeczytany
                self.exchange_client.mark_as_read(email)
            return success
        except Exception as e:
            logging.error(f"Błąd przenoszenia emaila: {e}")
            # Fallback - oznacz jako przeczytany
            return self.exchange_client.mark_as_read(email)