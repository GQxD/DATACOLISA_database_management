"""Extraction service for reading and filtering source data."""

from __future__ import annotations

from pathlib import Path
from typing import Any, List

from domain.models import SourceRow, ExtractionResult
from domain.value_objects import RefCode
from domain.business_rules import TypePecheDeriver, PaysDeriver
from infrastructure.excel_reader import ExcelReader
from infrastructure.csv_repository import CSVRepository


class ExtractionService:
    """
    Orchestrates extraction workflow from source Excel to selection CSV.

    This service coordinates:
    1. Reading source Excel file
    2. Filtering rows by REF range
    3. Deriving business values
    4. Writing selection CSV
    """

    def __init__(
        self,
        excel_reader: ExcelReader,
        csv_repo: CSVRepository
    ):
        """
        Initialize ExtractionService.

        Args:
            excel_reader: Excel reading operations
            csv_repo: CSV repository for saving selections
        """
        self.excel_reader = excel_reader
        self.csv_repo = csv_repo

    def extract_range(
        self,
        source_path: Path,
        sheet_name: str,
        start_ref: RefCode,
        end_ref: RefCode,
        output_csv: Path,
        default_type_echantillon: str = "EC"
    ) -> ExtractionResult:
        """
        Extract rows within REF range from source file to CSV.

        This replaces the cmd_extract() function logic.

        Args:
            source_path: Path to source .xls file
            sheet_name: Sheet name to read
            start_ref: Start of REF range (inclusive)
            end_ref: End of REF range (inclusive)
            output_csv: Path where to save selection CSV
            default_type_echantillon: Default type for all rows

        Returns:
            ExtractionResult with statistics and missing codes

        Raises:
            IOError: If file cannot be read/written
            ValueError: If sheet not found
        """
        # 1. Read source file
        source_rows, datemode = self.excel_reader.read_source_rows(
            source_path,
            sheet_name
        )

        # 2. Find candidate rows (rows with valid REF codes)
        candidates = self._find_candidate_rows(source_rows, datemode)

        # 3. Filter by REF range
        filtered = []
        for row in candidates:
            ref = RefCode.parse(row.ref)
            if ref:
                try:
                    if ref.in_range(start_ref, end_ref):
                        filtered.append(row)
                except ValueError:
                    # Different prefixes - skip
                    continue

        # 4. Sort by REF number
        filtered.sort(key=lambda r: RefCode.parse(r.ref).number if RefCode.parse(r.ref) else 0)

        # 5. Find missing codes in range
        found_codes = {row.ref.upper() for row in filtered}
        missing_codes = []
        for num in range(start_ref.number, end_ref.number + 1):
            code = f"{start_ref.prefix}{num}"
            if code not in found_codes:
                missing_codes.append(code)

        # 6. Save to CSV
        self.csv_repo.save_selection(filtered, output_csv, default_type_echantillon)

        # 7. Return result
        return ExtractionResult(
            rows=filtered,
            missing_codes=missing_codes,
            found_count=len(filtered),
            range_spec=f"{start_ref}..{end_ref}",
            extract_csv_path=str(output_csv)
        )

    def _find_candidate_rows(
        self,
        rows: List[List[Any]],
        datemode: int
    ) -> List[SourceRow]:
        """
        Find candidate rows with valid REF codes.

        Args:
            rows: Raw rows from Excel
            datemode: Excel datemode for date parsing

        Returns:
            List of SourceRow objects
        """
        from config.mappings import SOURCE_POSITIONS
        from domain.value_objects import DateCapture

        candidates = []
        for i, row in enumerate(rows, start=1):
            # Get REF from primary position
            ref = self._get_pos(row, SOURCE_POSITIONS["num_individu_primary"])
            ref = self._normalize_ref_code(ref)

            # Skip rows without REF or without digits
            if not ref or not any(c.isdigit() for c in ref):
                continue

            # Get or fallback to secondary REF position
            num_individu = ref or self._normalize_ref_code(
                self._get_pos(row, SOURCE_POSITIONS["num_individu_fallback"])
            )

            # Get raw engin for derivation
            engin_raw = self._get_pos(row, SOURCE_POSITIONS["engin_source"])
            contexte_raw = self._get_pos(row, SOURCE_POSITIONS["contexte"])

            # Derive business values
            type_peche, categorie = TypePecheDeriver.derive_type_and_categorie(
                str(engin_raw or "")
            )
            pays_capture = PaysDeriver.derive_country(str(contexte_raw or ""))

            # Parse date
            date_raw = self._get_pos(row, SOURCE_POSITIONS["date_capture"])
            date_capture = DateCapture.from_excel(date_raw, datemode)
            date_str = date_capture.format_display() if date_capture else self._normalize(date_raw)

            # Normalize placeholders
            poids_src = self._normalize(self._get_pos(row, SOURCE_POSITIONS["poids_g"]))
            sexe_src = self._normalize(self._get_pos(row, SOURCE_POSITIONS["sexe"]))
            if poids_src == "_":
                poids_src = ""
            if sexe_src == "_":
                sexe_src = ""

            candidates.append(SourceRow(
                source_row_index=i,
                ref=ref,
                code_espece=self._get_pos(row, SOURCE_POSITIONS["code_espece"]),
                date_capture=date_str,
                lac_riviere=self._get_pos(row, SOURCE_POSITIONS["lac_riviere"]),
                num_individu=num_individu,
                longueur_mm=self._get_pos(row, SOURCE_POSITIONS["longueur_mm"]),
                poids_g=poids_src if poids_src else "N",
                maturite=self._get_pos(row, SOURCE_POSITIONS["maturite"]),
                sexe=sexe_src if sexe_src else "N",
                age_total=self._get_pos(row, SOURCE_POSITIONS["age_total"]),
                type_peche=type_peche,
                categorie=categorie,
                pecheur=self._normalize(self._get_pos(row, SOURCE_POSITIONS["pecheur"])),
                pays_capture=pays_capture,
                pecheur_source=self._normalize(engin_raw),
                observation_disponibilite="",
            ))

        return candidates

    @staticmethod
    def _normalize_ref_code(value: Any) -> str:
        """Normalize refs like 'XY 0682' to 'XY682'."""
        from domain.value_objects import RefCode

        txt = ExtractionService._normalize(value)
        if not txt:
            return ""
        ref = RefCode.parse(txt)
        if ref:
            return str(ref)
        return txt.upper().replace(" ", "")

    @staticmethod
    def _get_pos(row: List[Any], one_based_col: int) -> Any:
        """Get value at 1-based column position."""
        idx = one_based_col - 1
        if idx < 0 or idx >= len(row):
            return None
        return row[idx]

    @staticmethod
    def _normalize(value: Any) -> str:
        """Normalize value to string."""
        if value is None:
            return ""
        if isinstance(value, float):
            if value.is_integer():
                return str(int(value))
            return format(value, "f").rstrip("0").rstrip(".")
        return str(value).strip()
