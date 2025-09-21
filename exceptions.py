"""
Moduł z definicjami własnych wyjątków aplikacji.
"""


class EmailProcessorError(Exception):
    """Bazowy wyjątek dla błędów procesora emaili."""
    pass


class OperationCancelledError(EmailProcessorError):
    """Wyjątek wyrzucany gdy operacja zostanie anulowana przez użytkownika."""
    pass


class ExchangeConnectionError(EmailProcessorError):
    """Błąd połączenia z serwerem Exchange."""
    pass


class PDFProcessingError(EmailProcessorError):
    """Błąd przetwarzania pliku PDF."""
    pass


class EmailNotFoundError(EmailProcessorError):
    """Błąd gdy nie znaleziono oczekiwanego emaila."""
    pass


class FolderNotFoundError(EmailProcessorError):
    """Błąd gdy nie znaleziono folderu."""
    pass