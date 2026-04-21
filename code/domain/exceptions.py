"""
Domain exceptions.

Custom exceptions for business logic and domain validation errors.
These exceptions provide clear, actionable error messages to users.
"""

from typing import List, Tuple


class DatacolisaError(Exception):
    """Base exception for all DATACOLISA application errors."""
    pass


class ValidationError(DatacolisaError):
    """
    Business validation failed.

    Raised when data doesn't meet business validation rules.
    """

    def __init__(self, field: str, message: str):
        """
        Initialize validation error.

        Args:
            field: Field name that failed validation
            message: Human-readable error message
        """
        self.field = field
        self.message = message
        super().__init__(f"{field}: {message}")


class SheetNotFoundError(DatacolisaError):
    """
    Excel sheet not found in workbook.

    Raised when trying to access a sheet that doesn't exist.
    """

    def __init__(self, sheet_name: str, available_sheets: List[str] = None):
        """
        Initialize sheet not found error.

        Args:
            sheet_name: Name of sheet that wasn't found
            available_sheets: List of available sheet names (optional)
        """
        self.sheet_name = sheet_name
        self.available_sheets = available_sheets or []

        message = f"Sheet '{sheet_name}' not found"
        if self.available_sheets:
            message += f". Available sheets: {', '.join(self.available_sheets)}"

        super().__init__(message)


class DuplicateRowError(DatacolisaError):
    """
    Duplicate row detected during import.

    Raised when attempting to import a row that already exists
    (based on num_individu + type_echantillon key).
    """

    def __init__(self, key: Tuple[str, str], existing_row: int = None):
        """
        Initialize duplicate row error.

        Args:
            key: Tuple of (num_individu, code_type_echantillon)
            existing_row: Row number where duplicate was found (optional)
        """
        self.key = key
        self.existing_row = existing_row

        message = f"Duplicate row: {key}"
        if existing_row:
            message += f" (existing at row {existing_row})"

        super().__init__(message)


class FileAccessError(DatacolisaError):
    """
    File cannot be read or written.

    Raised when file operations fail (permissions, locked, not found, etc.).
    """

    def __init__(self, file_path: str, operation: str, reason: str = ""):
        """
        Initialize file access error.

        Args:
            file_path: Path to file that couldn't be accessed
            operation: Operation that failed (read, write, open, etc.)
            reason: Additional reason/details (optional)
        """
        self.file_path = file_path
        self.operation = operation
        self.reason = reason

        message = f"Cannot {operation} file: {file_path}"
        if reason:
            message += f" ({reason})"

        super().__init__(message)


class InvalidRefCodeError(DatacolisaError):
    """
    Reference code format is invalid.

    Raised when RefCode parsing fails due to invalid format.
    """

    def __init__(self, code: str, expected_format: str = "LETTER(S)+NUMBER"):
        """
        Initialize invalid ref code error.

        Args:
            code: Invalid code string
            expected_format: Description of expected format
        """
        self.code = code
        self.expected_format = expected_format

        super().__init__(
            f"Invalid reference code format: '{code}'. "
            f"Expected: {expected_format}"
        )


class ConfigurationError(DatacolisaError):
    """
    Configuration is invalid or missing.

    Raised when application configuration is incorrect.
    """

    def __init__(self, config_key: str, reason: str):
        """
        Initialize configuration error.

        Args:
            config_key: Configuration key that's invalid
            reason: Why it's invalid
        """
        self.config_key = config_key
        self.reason = reason

        super().__init__(f"Configuration error for '{config_key}': {reason}")


class DateParsingError(DatacolisaError):
    """
    Date value cannot be parsed.

    Raised when date parsing fails for all known formats.
    """

    def __init__(self, value: any, attempted_formats: List[str] = None):
        """
        Initialize date parsing error.

        Args:
            value: Value that couldn't be parsed
            attempted_formats: List of formats that were tried
        """
        self.value = value
        self.attempted_formats = attempted_formats or []

        message = f"Cannot parse date: '{value}'"
        if self.attempted_formats:
            message += f". Tried formats: {', '.join(self.attempted_formats)}"

        super().__init__(message)
