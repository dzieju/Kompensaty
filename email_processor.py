"""
Moduł głównej logiki przetwarzania emaili.
"""
import logging
from datetime import datetime
from typing import Optional

from config import AppConfig
from models import SearchCriteria, EmailTemplateCollection, ProcessingResult
from email_service import EmailService
from pdf_processor import PDFProcessor
from exceptions import (
    OperationCancelledError, 
    ExchangeConnectionError, 
    EmailNotFoundError,
    PDFProcessingError
)


class EmailProcessor:
    """Główny procesor emaili - orkiestruje cały proces przetwarzania."""
    
    def __init__(self, config: AppConfig, templates: EmailTemplateCollection):
        """
        Inicjalizuje procesor emaili.
        
        Args:
            config: Konfiguracja aplikacji
            templates: Szablony emaili
        """
        self.config = config
        self.templates = templates
        self.email_service = EmailService(config, templates)
        self.pdf_processor = PDFProcessor()

    def process_emails(self, search_criteria: SearchCriteria, stop_event) -> ProcessingResult:
        """
        Główna metoda przetwarzania emaili.
        
        Args:
            search_criteria: Kryteria wyszukiwania
            stop_event: Event do anulowania operacji
            
        Returns:
            Wynik przetwarzania
            
        Raises:
            OperationCancelledError: Operacja anulowana
            ExchangeConnectionError: Błąd połączenia z Exchange
        """
        logging.info("=" * 50)
        logging.info(f"Rozpoczynanie przetwarzania emaili - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        try:
            if stop_event.is_set():
                raise OperationCancelledError("Anulowano na starcie przetwarzania")

            # Krok 1: Znajdź główny email z kompensatą
            logging.info("Krok 1: Wyszukiwanie głównego emaila z kompensatą...")
            compensation_email, pdf_attachment = self._find_compensation_email(search_criteria, stop_event)
            
            if not compensation_email or not pdf_attachment:
                return ProcessingResult(
                    success=False,
                    message="Nie znaleziono głównego emaila z kompensatą zgodnie z kryteriami"
                )

            logging.info(f"Znaleziono email: '{compensation_email.subject}' "
                        f"(otrzymany: {compensation_email.datetime_received})")

            # Krok 2: Ekstraktuj dane z PDF
            logging.info("Krok 2: Ekstraktowanie danych z PDF...")
            extracted_data = self._extract_data_from_pdf(pdf_attachment, search_criteria, stop_event)
            
            if not extracted_data:
                return self._handle_no_data_extracted(compensation_email, pdf_attachment, stop_event)

            # Krok 3: Znajdź powiązane faktury
            logging.info("Krok 3: Wyszukiwanie powiązanych faktur...")
            found_invoices = self._find_related_invoices(
                extracted_data, compensation_email, search_criteria, stop_event
            )

            # Krok 4: Wyślij powiadomienie o częściowych wynikach (jeśli potrzeba)
            if len(found_invoices) < len(extracted_data):
                logging.info("Krok 4a: Wysyłanie powiadomienia o częściowych wynikach...")
                self.email_service.send_partial_results_notification(
                    compensation_email, extracted_data, found_invoices, stop_event
                )

            # Krok 5: Wyślij email do księgowej
            logging.info("Krok 5: Wysyłanie emaila do księgowej...")
            email_sent = self.email_service.send_compensation_email_to_accounting(
                pdf_attachment, found_invoices, extracted_data, compensation_email, stop_event
            )

            # Krok 6: Przenieś email do folderu przetworzonych
            if email_sent:
                logging.info("Krok 6: Przenoszenie emaila do folderu przetworzonych...")
                self.email_service.move_email_to_processed_folder(compensation_email, stop_event)
                
                result_message = (
                    f"Przetwarzanie zakończone pomyślnie. "
                    f"Wysłano email z {len(found_invoices)}/{len(extracted_data)} fakturami."
                )
            else:
                result_message = "Błąd wysyłania emaila do księgowej. Email nie został przeniesiony."

            return ProcessingResult(
                success=email_sent,
                message=result_message,
                email_subject=compensation_email.subject,
                extracted_data=extracted_data,
                found_attachments_count=len(found_invoices),
                missing_invoices=self.email_service._find_missing_invoices(extracted_data, found_invoices)
            )

        except OperationCancelledError:
            logging.warning("Przetwarzanie zostało anulowane przez użytkownika")
            return ProcessingResult(
                success=False,
                message="Przetwarzanie anulowane przez użytkownika"
            )
        except Exception as e:
            logging.error(f"Krytyczny błąd podczas przetwarzania: {e}")
            return ProcessingResult(
                success=False,
                message=f"Błąd krytyczny: {e}"
            )
        finally:
            logging.info("Zakończono przetwarzanie emaili")
            logging.info("=" * 50)

    def _find_compensation_email(self, criteria: SearchCriteria, stop_event):
        """
        Znajduje główny email z kompensatą.
        
        Args:
            criteria: Kryteria wyszukiwania
            stop_event: Event do anulowania
            
        Returns:
            Krotka (email, załącznik) lub (None, None)
        """
        return self.email_service.exchange_client.find_compensation_email(criteria, stop_event)

    def _extract_data_from_pdf(self, pdf_attachment, criteria: SearchCriteria, stop_event):
        """
        Ekstraktuje dane z załącznika PDF.
        
        Args:
            pdf_attachment: Załącznik PDF
            criteria: Kryteria ekstraktowania
            stop_event: Event do anulowania
            
        Returns:
            Lista wyekstrahowanych danych
        """
        return self.pdf_processor.extract_invoice_numbers(
            pdf_attachment.content,
            criteria.extraction.regex_pattern,
            stop_event
        )

    def _find_related_invoices(self, extracted_data, compensation_email, criteria: SearchCriteria, stop_event):
        """
        Znajduje powiązane faktury.
        
        Args:
            extracted_data: Wyekstrahowane numery faktur
            compensation_email: Główny email z kompensatą
            criteria: Kryteria wyszukiwania
            stop_event: Event do anulowania
            
        Returns:
            Lista znalezionych załączników
        """
        return self.email_service.find_related_invoices(
            extracted_data,
            compensation_email.datetime_received,
            criteria,
            stop_event
        )

    def _handle_no_data_extracted(self, compensation_email, pdf_attachment, stop_event) -> ProcessingResult:
        """
        Obsługuje sytuację gdy nie udało się wyekstrahować danych z PDF.
        
        Args:
            compensation_email: Email z kompensatą
            pdf_attachment: Załącznik PDF
            stop_event: Event do anulowania
            
        Returns:
            Wynik przetwarzania
        """
        logging.warning("Nie znaleziono danych w głównym załączniku PDF")
        
        # Wyślij powiadomienie o problemie
        self.email_service.send_missing_pdf_notification(
            compensation_email, pdf_attachment.name, stop_event
        )
        
        # Przenieś email do folderu przetworzonych (z problemem)
        self.email_service.move_email_to_processed_folder(compensation_email, stop_event)
        
        return ProcessingResult(
            success=False,
            message="Nie udało się wyekstrahować danych z PDF. Wysłano powiadomienie.",
            email_subject=compensation_email.subject
        )

    def validate_configuration(self) -> tuple[bool, str]:
        """
        Waliduje konfigurację przed rozpoczęciem przetwarzania.
        
        Returns:
            Krotka (czy_poprawna, komunikat_błędu)
        """
        if not self.config.exchange.username:
            return False, "Brak nazwy użytkownika Exchange"
            
        if not self.config.exchange.password:
            return False, "Brak hasła użytkownika Exchange"
            
        if not self.config.get_ksiegowa_emails():
            return False, "Brak adresów email księgowej"
            
        return True, "Konfiguracja poprawna"

    def test_exchange_connection(self) -> tuple[bool, str]:
        """
        Testuje połączenie z Exchange.
        
        Returns:
            Krotka (czy_połączono, komunikat)
        """
        try:
            # Próba dostępu do skrzynki odbiorczej
            inbox = self.email_service.exchange_client.account.inbox
            if inbox:
                return True, f"Połączono z kontem: {self.config.exchange.username}"
            else:
                return False, "Nie udało się uzyskać dostępu do skrzynki odbiorczej"
        except Exception as e:
            return False, f"Błąd połączenia: {e}"

    def get_processing_statistics(self) -> dict:
        """
        Zwraca statystyki przetwarzania.
        
        Returns:
            Słownik ze statystykami
        """
        return {
            "exchange_server": self.config.exchange.server or "autodiscover",
            "username": self.config.exchange.username,
            "compensation_folder": self.config.folders.kompensaty,
            "invoices_folder": self.config.folders.faktury,
            "processed_folder": self.config.folders.przetworzone,
            "months_back_compensation": self.config.search.miesiace_wstecz_kompensaty,
            "months_back_invoices": self.config.search.miesiace_wstecz_faktury,
            "accounting_emails": len(self.config.get_ksiegowa_emails()),
            "notification_emails_pdf": len(self.config.get_notification_emails_for_missing_pdf()),
            "notification_emails_partial": len(self.config.get_notification_emails_for_partial_results())
        }