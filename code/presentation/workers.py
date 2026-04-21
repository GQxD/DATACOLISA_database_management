"""Background workers for long-running operations (threading)."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Dict, List, Optional
from PySide6.QtCore import QThread, Signal

logger = logging.getLogger(__name__)


class LoadRangeWorker(QThread):
    """
    Background worker for loading and filtering source data.

    Prevents UI freeze during:
    - Reading large Excel files
    - Filtering by REF range
    - Validating rows

    Signals:
        progress(int, int): Emits (current, total) for progress updates
        finished(dict): Emits result dict with rows, missing_codes, etc.
        error(str): Emits error message if operation fails
    """

    progress = Signal(int, int)  # current, total
    finished = Signal(dict)  # result dictionary
    error = Signal(str)  # error message

    def __init__(
        self,
        source_path: Path,
        source_sheet: str,
        source_mode: str,
        source_mapping: Optional[Dict[str, Any]],
        start_ref: str,
        end_ref: str,
        default_type_echantillon: str,
        parent=None
    ):
        super().__init__(parent)
        self.source_path = source_path
        self.source_sheet = source_sheet
        self.source_mode = source_mode
        self.source_mapping = source_mapping or {}
        self.start_ref = start_ref
        self.end_ref = end_ref
        self.default_type_echantillon = default_type_echantillon

    def run(self) -> None:
        """Execute load operation in background thread."""
        try:
            logger.info(f"LoadRangeWorker: Loading {self.source_path.name}, range {self.start_ref}-{self.end_ref}")

            # Import here to avoid circular dependencies
            import datacolisa_importer as core

            # Read source rows
            source_rows, datemode = core.read_any_source_rows(
                self.source_path, self.source_sheet
            )
            logger.debug(f"Read {len(source_rows)} rows from Excel")
            self.progress.emit(1, 3)  # Step 1/3

            # Find candidate rows
            if self.source_mode == "custom":
                candidates = core.find_candidate_rows_from_mapping(
                    source_rows,
                    datemode,
                    self.source_mapping,
                )
            else:
                candidates = core.find_candidate_rows(source_rows, datemode)
            self.progress.emit(2, 3)  # Step 2/3

            # Filter by REF range
            filtered = [
                r for r in candidates
                if core.in_ref_range(r.ref, self.start_ref, self.end_ref)
            ]

            # Sort by REF number
            filtered.sort(
                key=lambda r: (
                    core.parse_ref_parts(r.ref)[1]
                    if core.parse_ref_parts(r.ref) else 0
                )
            )

            # Find missing codes
            found_codes = {core.normalize(r.ref).upper() for r in filtered}
            missing_codes = []
            p_start = core.parse_ref_parts(self.start_ref)
            p_end = core.parse_ref_parts(self.end_ref)
            if p_start and p_end:
                prefix = p_start[0]
                for num in range(p_start[1], p_end[1] + 1):
                    code = f"{prefix}{num}"
                    if code not in found_codes:
                        missing_codes.append(code)

            # Convert SourceRow to dict format for table
            rows: List[Dict[str, Any]] = []
            for r in filtered:
                # Validate row
                errs = core.validate_row(r, self.default_type_echantillon)
                has_date = bool(core.normalize(r.date_capture))

                rows.append({
                    "selected": has_date,
                    "ref": core.normalize_ref_code(r.ref),
                    "code_type_echantillon": self.default_type_echantillon,
                    "categorie": core.normalize(r.categorie),
                    "type_peche": core.normalize(r.type_peche),
                    "autre_oss": "NON",
                    "ecailles_brutes": core.normalize(getattr(r, "ecailles_brutes", "")),
                    "montees": core.normalize(getattr(r, "montees", "")),
                    "empreintes": core.normalize(getattr(r, "empreintes", "")),
                    "otolithes": core.normalize(getattr(r, "otolithes", "")),
                    "observation_disponibilite": core.normalize(r.observation_disponibilite),
                    "source_row": r.source_row_index,
                    "num_individu": core.normalize_ref_code(r.num_individu),
                    "date_capture": core.normalize(r.date_capture),
                    "code_espece": core.normalize(r.code_espece),
                    "lac_riviere": core.normalize(r.lac_riviere),
                    "pays_capture": core.normalize(r.pays_capture),
                    "pecheur": core.normalize(r.pecheur),
                    "longueur_mm": core.normalize(r.longueur_mm),
                    "poids_g": core.normalize(r.poids_g),
                    "maturite": core.normalize(r.maturite),
                    "sexe": core.normalize(r.sexe),
                    "age_total": core.normalize(r.age_total),
                    "sous_espece": core.normalize(getattr(r, "sous_espece", "")),
                    "nom_operateur": core.normalize(getattr(r, "nom_operateur", "")),
                    "lieu_capture": core.normalize(getattr(r, "lieu_capture", "")),
                    "maille_mm": core.normalize(getattr(r, "maille_mm", "")),
                    "code_stade": core.normalize(getattr(r, "code_stade", "")),
                    "presence_otolithe_gauche": core.normalize(getattr(r, "presence_otolithe_gauche", "")),
                    "presence_otolithe_droite": core.normalize(getattr(r, "presence_otolithe_droite", "")),
                    "nb_opercules": core.normalize(getattr(r, "nb_opercules", "")),
                    "information_stockage": core.normalize(getattr(r, "information_stockage", "")),
                    "age_riviere": core.normalize(getattr(r, "age_riviere", "")),
                    "age_lac": core.normalize(getattr(r, "age_lac", "")),
                    "nb_fraie": core.normalize(getattr(r, "nb_fraie", "")),
                    "ecailles_regenerees": core.normalize(getattr(r, "ecailles_regenerees", "")),
                    "observations": core.normalize(getattr(r, "observations", "")),
                    "status": "a_reimporter" if errs else "pret",
                    "errors": " | ".join(errs) if errs else "",
                })

            self.progress.emit(3, 3)  # Step 3/3

            # Emit result
            result = {
                "rows": rows,
                "missing_codes": missing_codes,
                "found_count": len(filtered),
                "pending_count": sum(1 for row in rows if row["status"] == "a_reimporter"),
                "missing_date_count": sum(1 for row in rows if not row.get("date_capture")),
            }
            logger.info(f"LoadRangeWorker: Completed, found {len(filtered)} rows, {len(missing_codes)} missing")
            self.finished.emit(result)

        except Exception as e:
            logger.error(f"LoadRangeWorker failed: {e}", exc_info=True)
            self.error.emit(str(e))


class ImportWorker(QThread):
    """
    Background worker for importing data to target Excel.

    Prevents UI freeze during:
    - CSV loading
    - Excel workbook manipulation
    - Formula propagation
    - File saving

    Signals:
        progress(int, int): Emits (current, total) for progress updates
        finished(dict): Emits result dict with import statistics
        error(str): Emits error message if operation fails
    """

    progress = Signal(int, int)  # current, total
    finished = Signal(dict)  # result dictionary
    error = Signal(str)  # error message

    def __init__(
        self,
        selection_csv: Path,
        selection_rows: Optional[List[Dict[str, Any]]],
        target_path: Path,
        target_sheet: str,
        out_target: Path,
        history_path: Path,
        default_organisme: str,
        default_country: str,
        default_code_unite_gestionnaire: str,
        default_site_atelier: str,
        default_numero_correspondant: str,
        on_duplicate: str,
        start_numero: int = 0,
        parent=None
    ):
        super().__init__(parent)
        self.selection_csv = selection_csv
        self.selection_rows = selection_rows or []
        self.target_path = target_path
        self.target_sheet = target_sheet
        self.out_target = out_target
        self.history_path = history_path
        self.default_organisme = default_organisme
        self.default_country = default_country
        self.default_code_unite_gestionnaire = default_code_unite_gestionnaire
        self.default_site_atelier = default_site_atelier
        self.default_numero_correspondant = default_numero_correspondant
        self.on_duplicate = on_duplicate
        self.start_numero = start_numero

    def run(self) -> None:
        """Execute import operation in background thread."""
        try:
            if self.selection_rows:
                logger.info(f"ImportWorker: Starting import from in-memory table ({len(self.selection_rows)} rows)")
            else:
                logger.info(f"ImportWorker: Starting import from {self.selection_csv.name}")

            # Import here to avoid circular dependencies
            import datacolisa_importer as core
            from domain.models import ImportConfig

            # Ensure dependencies
            openpyxl, xlrd = core.ensure_deps()

            self.progress.emit(1, 5)  # Step 1/5

            # Create services
            from infrastructure.excel_reader import ExcelReader
            from infrastructure.excel_writer import ExcelWriter
            from infrastructure.csv_repository import CSVRepository
            from infrastructure.history_repository import HistoryRepository
            from application.import_service import ImportService

            excel_reader = ExcelReader(xlrd)
            excel_writer = ExcelWriter(openpyxl)
            csv_repo = CSVRepository()
            history_repo = HistoryRepository()

            service = ImportService(excel_reader, excel_writer, csv_repo, history_repo)

            self.progress.emit(2, 5)  # Step 2/5

            # Build configuration
            config = ImportConfig(
                selection_csv=self.selection_csv,
                target_path=self.target_path,
                target_sheet=self.target_sheet,
                output_path=self.out_target,
                history_path=self.history_path,
                default_organisme=self.default_organisme,
                default_country=self.default_country,
                default_code_unite_gestionnaire=self.default_code_unite_gestionnaire,
                default_site_atelier=self.default_site_atelier,
                default_numero_correspondant=self.default_numero_correspondant,
                on_duplicate=self.on_duplicate,
                selection_rows=self.selection_rows,
                start_numero=self.start_numero,
            )

            self.progress.emit(3, 5)  # Step 3/5

            # Execute import
            result = service.import_selection(config)

            self.progress.emit(4, 5)  # Step 4/5

            # Convert result to dict
            result_dict = result.to_summary()

            self.progress.emit(5, 5)  # Step 5/5

            # Emit result
            logger.info(f"ImportWorker: Completed, imported {result_dict.get('imported', 0)} rows")
            self.finished.emit(result_dict)

        except Exception as e:
            logger.error(f"ImportWorker failed: {e}", exc_info=True)
            self.error.emit(str(e))
