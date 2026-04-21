"""
Value objects for domain concepts.

Value objects are immutable objects that represent domain concepts
with validation and behavior encapsulated.
"""

from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass
from typing import Optional

from config.constants import DATE_FORMATS


@dataclass(frozen=True)
class RefCode:
    """
    Represents a reference code (e.g., CA961, T00042).

    RefCodes consist of an alphabetic prefix and a numeric part.
    They are immutable and can be compared for range checks.
    """
    prefix: str
    number: int

    def __post_init__(self):
        """Validate RefCode components."""
        if not self.prefix.isalpha():
            raise ValueError(f"RefCode prefix must be alphabetic: {self.prefix}")
        if self.number < 0:
            raise ValueError(f"RefCode number must be non-negative: {self.number}")

    @classmethod
    def parse(cls, code: str) -> Optional[RefCode]:
        """
        Parse a string into a RefCode.

        Examples:
            >>> RefCode.parse("CA961")
            RefCode(prefix='CA', number=961)
            >>> RefCode.parse("T00042")
            RefCode(prefix='T', number=42)
            >>> RefCode.parse("invalid")
            None

        Args:
            code: String to parse

        Returns:
            RefCode object or None if parsing fails
        """
        if not code:
            return None

        code = code.strip().upper()
        # Match pattern: letters, optional whitespace, optional leading zeros, digits
        m = re.match(r"^([A-Z]+)\s*0*(\d+)$", code)
        if not m:
            return None

        try:
            return cls(prefix=m.group(1), number=int(m.group(2)))
        except (ValueError, AttributeError):
            return None

    def in_range(self, start: RefCode, end: RefCode) -> bool:
        """
        Check if this RefCode is within a range (inclusive).

        Args:
            start: Start of range
            end: End of range

        Returns:
            True if this code is in range [start, end]

        Raises:
            ValueError: If range has different prefixes or is invalid
        """
        # All codes must have same prefix
        if start.prefix != end.prefix or self.prefix != start.prefix:
            raise ValueError(
                f"RefCode range requires same prefix: "
                f"{start.prefix}, {self.prefix}, {end.prefix}"
            )

        # Start must be <= end
        if start.number > end.number:
            raise ValueError(
                f"Invalid range: start {start.number} > end {end.number}"
            )

        return start.number <= self.number <= end.number

    def __str__(self) -> str:
        """String representation: CA961"""
        return f"{self.prefix}{self.number}"

    def __lt__(self, other: RefCode) -> bool:
        """Enable sorting of RefCodes."""
        if self.prefix != other.prefix:
            return self.prefix < other.prefix
        return self.number < other.number


@dataclass(frozen=True)
class DateCapture:
    """
    Represents a capture date with parsing and formatting logic.

    Handles various Excel date formats and string representations.
    """
    date: dt.date

    @classmethod
    def from_excel(
        cls,
        value: any,
        datemode: int = 0
    ) -> Optional[DateCapture]:
        """
        Parse Excel date value into DateCapture.

        Handles multiple input types:
        - Excel numeric dates (float)
        - Python datetime/date objects
        - String dates in various formats

        Args:
            value: Excel cell value
            datemode: Excel datemode (0 or 1) for numeric dates

        Returns:
            DateCapture object or None if parsing fails
        """
        if value is None:
            return None

        # Already a date object
        if isinstance(value, dt.date) and not isinstance(value, dt.datetime):
            return cls(date=value)

        # datetime object
        if isinstance(value, dt.datetime):
            return cls(date=value.date())

        # Excel numeric date
        if isinstance(value, (int, float)):
            try:
                # Try xlrd conversion first
                import xlrd
                tup = xlrd.xldate_as_tuple(float(value), datemode)
                return cls(date=dt.date(tup[0], tup[1], tup[2]))
            except Exception:
                pass

            # Fallback: Excel date calculation (1899-12-30 as base)
            try:
                base = dt.date(1899, 12, 30)
                return cls(date=base + dt.timedelta(days=int(float(value))))
            except Exception:
                return None

        # String date
        if isinstance(value, str):
            s = value.strip()
            if not s:
                return None

            import re
            match = re.search(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b|\b\d{4}-\d{2}-\d{2}\b", s)
            if match:
                s = match.group(0)

            for fmt in DATE_FORMATS:
                try:
                    parsed = dt.datetime.strptime(s, fmt)
                    return cls(date=parsed.date())
                except ValueError:
                    continue

        return None

    def format_display(self) -> str:
        """
        Format date for display in UI/CSV.

        Returns:
            Date string in dd/mm/yy format
        """
        return self.date.strftime("%d/%m/%y")

    def format_iso(self) -> str:
        """
        Format date in ISO format (YYYY-MM-DD).

        Returns:
            Date string in ISO format
        """
        return self.date.isoformat()

    def __str__(self) -> str:
        """String representation: dd/mm/yy"""
        return self.format_display()


def parse_ref_code(code: str) -> Optional[tuple[str, int]]:
    """
    Legacy function for backward compatibility.

    Use RefCode.parse() for new code.

    Args:
        code: String to parse

    Returns:
        Tuple of (prefix, number) or None
    """
    ref = RefCode.parse(code)
    if ref:
        return (ref.prefix, ref.number)
    return None


def in_ref_range(code: str, start: str, end: str) -> bool:
    """
    Legacy function for backward compatibility.

    Use RefCode.in_range() for new code.

    Args:
        code: Code to check
        start: Start of range
        end: End of range

    Returns:
        True if code is in range
    """
    ref = RefCode.parse(code)
    start_ref = RefCode.parse(start)
    end_ref = RefCode.parse(end)

    if not ref or not start_ref or not end_ref:
        return False

    try:
        return ref.in_range(start_ref, end_ref)
    except ValueError:
        return False
