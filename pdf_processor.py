"""
Moduł do przetwarzania plików PDF.
"""
import io
import re
import logging
import traceback
from typing import List, Set
import PyPDF2

from exceptions import PDFProcessingError, OperationCancelledError


class PDFProcessor:
    """Klasa do przetwarzania plików PDF."""
    
    def __init__(self):
        """Inicjalizuje procesor PDF."""
        pass

    def extract_invoice_numbers(self, pdf_content: bytes, regex_pattern: str, 
                               stop_event=None) -> List[str]:
        """
        Ekstraktuje numery faktur z zawartości PDF używając podanego wzorca regex.
        
        Args:
            pdf_content: Binarna zawartość pliku PDF
            regex_pattern: Wzorzec regex do ekstraktowania numerów
            stop_event: Event do anulowania operacji
            
        Returns:
            Lista unikalnych numerów faktur
            
        Raises:
            PDFProcessingError: Błąd przetwarzania PDF
            OperationCancelledError: Operacja anulowana
        """
        if stop_event and stop_event.is_set():
            raise OperationCancelledError("Anulowano odczyt numerów faktur z PDF")

        if not regex_pattern:
            raise PDFProcessingError("Brak wzorca Regex do ekstrakcji danych z PDF")

        invoice_numbers: Set[str] = set()
        
        try:
            pdf_file = io.BytesIO(pdf_content)
            reader = PyPDF2.PdfReader(pdf_file)
            
            logging.info(f"Przetwarzanie PDF, liczba stron: {len(reader.pages)}")
            
            for page_num, page in enumerate(reader.pages):
                if stop_event and stop_event.is_set():
                    raise OperationCancelledError("Anulowano przetwarzanie stron PDF")

                text = self._extract_text_from_page(page, page_num)
                if text:
                    numbers = self._find_invoice_numbers_in_text(text, regex_pattern, page_num)
                    invoice_numbers.update(numbers)

            if not invoice_numbers:
                logging.warning("Nie udało się wyekstrahować żadnych danych z PDF za pomocą podanego Regex")
            else:
                logging.info(f"Wyekstrahowano następujące unikalne dane: {list(invoice_numbers)}")

            return list(invoice_numbers)

        except OperationCancelledError:
            raise
        except Exception as e:
            logging.error(f"Błąd podczas przetwarzania PDF: {e}")
            logging.error(traceback.format_exc())
            raise PDFProcessingError(f"Błąd przetwarzania PDF: {e}")

    def _extract_text_from_page(self, page, page_num: int) -> str:
        """
        Ekstraktuje tekst ze strony PDF.
        
        Args:
            page: Obiekt strony PDF
            page_num: Numer strony (do logowania)
            
        Returns:
            Tekst ze strony
        """
        try:
            text = page.extract_text()
            if text:
                # Normalizacja tekstu - usunięcie nadmiarowych białych znaków
                normalized_text = " ".join(text.split())
                return normalized_text
            else:
                logging.warning(f"Strona {page_num + 1} PDF nie zawierała tekstu")
                return ""
        except Exception as e:
            logging.warning(f"Błąd ekstraktowania tekstu ze strony {page_num + 1}: {e}")
            return ""

    def _find_invoice_numbers_in_text(self, text: str, regex_pattern: str, 
                                     page_num: int) -> Set[str]:
        """
        Znajduje numery faktur w tekście używając regex.
        
        Args:
            text: Tekst do przeszukania
            regex_pattern: Wzorzec regex
            page_num: Numer strony (do logowania)
            
        Returns:
            Zbiór znalezionych numerów faktur
        """
        invoice_numbers: Set[str] = set()
        
        try:
            matches = re.findall(regex_pattern, text, re.IGNORECASE)
            
            if matches:
                logging.debug(f"Na stronie {page_num + 1} znaleziono potencjalne dane: {matches}")
                
                for raw_number in matches:
                    cleaned_number = self._clean_invoice_number(raw_number)
                    if cleaned_number:
                        invoice_numbers.add(cleaned_number)
                        
        except re.error as e:
            logging.error(f"Błąd wzorca regex '{regex_pattern}': {e}")
        except Exception as e:
            logging.warning(f"Błąd wyszukiwania na stronie {page_num + 1}: {e}")
            
        return invoice_numbers

    def _clean_invoice_number(self, raw_number: str) -> str:
        """
        Czyści numer faktury z niepotrzebnych elementów.
        
        Args:
            raw_number: Surowy numer faktury
            
        Returns:
            Oczyszczony numer faktury
        """
        cleaned = raw_number.strip()
        
        # Usuń końcówki typu " 12", " 01" itp.
        match_ending = re.search(r'(\s+\d{1,2})$', cleaned)
        if match_ending:
            cleaned = cleaned[:-len(match_ending.group(1))]
            
        return cleaned.strip()

    def validate_pdf_content(self, pdf_content: bytes) -> bool:
        """
        Sprawdza czy zawartość to poprawny plik PDF.
        
        Args:
            pdf_content: Binarna zawartość pliku
            
        Returns:
            True jeśli to poprawny PDF
        """
        try:
            pdf_file = io.BytesIO(pdf_content)
            reader = PyPDF2.PdfReader(pdf_file)
            # Próba dostępu do pierwszej strony
            if len(reader.pages) > 0:
                _ = reader.pages[0]
                return True
            return False
        except Exception:
            return False

    def get_pdf_info(self, pdf_content: bytes) -> dict:
        """
        Pobiera informacje o pliku PDF.
        
        Args:
            pdf_content: Binarna zawartość pliku
            
        Returns:
            Słownik z informacjami o PDF
        """
        try:
            pdf_file = io.BytesIO(pdf_content)
            reader = PyPDF2.PdfReader(pdf_file)
            
            info = {
                "page_count": len(reader.pages),
                "is_encrypted": reader.is_encrypted,
                "metadata": {}
            }
            
            if reader.metadata:
                info["metadata"] = {
                    "title": reader.metadata.get("/Title", ""),
                    "author": reader.metadata.get("/Author", ""),
                    "creator": reader.metadata.get("/Creator", ""),
                    "producer": reader.metadata.get("/Producer", ""),
                    "creation_date": reader.metadata.get("/CreationDate", ""),
                    "modification_date": reader.metadata.get("/ModDate", "")
                }
            
            return info
            
        except Exception as e:
            logging.warning(f"Nie udało się pobrać informacji o PDF: {e}")
            return {"error": str(e)}