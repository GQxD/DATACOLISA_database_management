"""Configuration layer - Mappings and constants."""

from .mappings import SOURCE_POSITIONS, TARGET_HEADERS, TYPE_SHEET_CANDIDATES
from .constants import (
    DEFAULT_SOURCE_SHEET,
    DEFAULT_TARGET_SHEET,
    DuplicatePolicy,
    ImportStatus,
)

__all__ = [
    "SOURCE_POSITIONS",
    "TARGET_HEADERS",
    "TYPE_SHEET_CANDIDATES",
    "DEFAULT_SOURCE_SHEET",
    "DEFAULT_TARGET_SHEET",
    "DuplicatePolicy",
    "ImportStatus",
]
