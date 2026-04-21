"""Domain layer - Business models and rules."""

from .models import (
    SourceRow,
    TargetRow,
    ImportConfig,
    ImportResult,
    ValidationResult,
    ExtractionResult,
)
from .value_objects import RefCode, DateCapture
from .business_rules import (
    TypePecheDeriver,
    PaysDeriver,
    ValidationRules,
    DataTransformations,
    ReferenceCodeGenerator,
)
from .exceptions import (
    DatacolisaError,
    ValidationError,
    SheetNotFoundError,
    DuplicateRowError,
    FileAccessError,
    InvalidRefCodeError,
    ConfigurationError,
    DateParsingError,
)

__all__ = [
    # Models
    "SourceRow",
    "TargetRow",
    "ImportConfig",
    "ImportResult",
    "ValidationResult",
    "ExtractionResult",
    # Value Objects
    "RefCode",
    "DateCapture",
    # Business Rules
    "TypePecheDeriver",
    "PaysDeriver",
    "ValidationRules",
    "DataTransformations",
    "ReferenceCodeGenerator",
    # Exceptions
    "DatacolisaError",
    "ValidationError",
    "SheetNotFoundError",
    "DuplicateRowError",
    "FileAccessError",
    "InvalidRefCodeError",
    "ConfigurationError",
    "DateParsingError",
]
