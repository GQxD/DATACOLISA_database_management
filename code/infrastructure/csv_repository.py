"""CSV file operations repository."""

from __future__ import annotations

import csv
from pathlib import Path
from typing import Any, Dict, List

from domain.models import SourceRow
from infrastructure.file_value_normalizer import coerce_internal_value


class CSVRepository:
    """
    Handles reading and writing CSV files for import selections.

    This class centralizes CSV operations, making them testable and
    reusable across the application.
    """

    def save_selection(
        self,
        rows: List[SourceRow],
        path: Path,
        default_type_echantillon: str = ""
    ) -> None:
        """
        Save source rows to a selection CSV file.

        Args:
            rows: List of SourceRow objects to save
            path: Path where to save the CSV
            default_type_echantillon: Default type for all rows

        Raises:
            IOError: If file cannot be written
        """
        try:
            path.parent.mkdir(parents=True, exist_ok=True)

            with path.open("w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)

                # Write header
                writer.writerow([
                    "include",
                    "status",
                    "source_row",
                    "ref",
                    "num_individu",
                    "date_capture",
                    "code_espece",
                    "lac_riviere",
                    "pays_capture",
                    "pecheur",
                    "pecheur_source",
                    "categorie",
                    "type_peche",
                    "observation_disponibilite",
                    "autre_oss",
                    "ecailles_brutes",
                    "montees",
                    "empreintes",
                    "otolithes",
                    "longueur_mm",
                    "poids_g",
                    "maturite",
                    "sexe",
                    "age_total",
                    "code_type_echantillon",
                    "errors",
                ])

                # Write data rows
                for row in rows:
                    # Validation will be done by validation service in Phase 4
                    # For now, just mark as "pret" (ready)
                    writer.writerow([
                        "1",  # include by default
                        "pret",  # status
                        coerce_internal_value("source_row", row.source_row_index),
                        row.ref,
                        row.num_individu,
                        row.date_capture,
                        row.code_espece,
                        row.lac_riviere,
                        row.pays_capture,
                        row.pecheur,
                        row.pecheur_source,
                        row.categorie,
                        row.type_peche,
                        row.observation_disponibilite,
                        "",  # autre_oss
                        coerce_internal_value("ecailles_brutes", row.ecailles_brutes),
                        coerce_internal_value("montees", row.montees),
                        coerce_internal_value("empreintes", row.empreintes),
                        coerce_internal_value("otolithes", row.otolithes),
                        coerce_internal_value("longueur_mm", row.longueur_mm),
                        coerce_internal_value("poids_g", row.poids_g),
                        row.maturite,
                        row.sexe,
                        row.age_total,
                        default_type_echantillon,
                        "",  # errors (will be filled by validation)
                    ])

        except Exception as e:
            raise IOError(f"Cannot write CSV to {path}: {e}") from e

    def load_selection(self, path: Path) -> List[Dict[str, Any]]:
        """
        Load selection CSV file as list of dictionaries.

        Args:
            path: Path to the CSV file

        Returns:
            List of dictionaries, one per row

        Raises:
            IOError: If file cannot be read
            ValueError: If CSV format is invalid
        """
        if not path.exists():
            raise FileNotFoundError(f"CSV file not found: {path}")

        try:
            with path.open("r", newline="", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                rows = list(reader)

            # Validate required columns exist
            if rows:
                required_cols = ["include", "num_individu", "code_type_echantillon"]
                first_row = rows[0]
                missing = [col for col in required_cols if col not in first_row]
                if missing:
                    raise ValueError(f"Missing required columns: {missing}")

            return rows

        except csv.Error as e:
            raise ValueError(f"Invalid CSV format in {path}: {e}") from e
        except Exception as e:
            raise IOError(f"Cannot read CSV from {path}: {e}") from e

    def update_row_status(
        self,
        path: Path,
        row_ref: str,
        new_status: str,
        errors: str = ""
    ) -> None:
        """
        Update status and errors for a specific row in CSV.

        Args:
            path: Path to the CSV file
            row_ref: REF code of the row to update
            new_status: New status value
            errors: Error messages to set

        Raises:
            IOError: If file cannot be read/written
            ValueError: If row not found
        """
        rows = self.load_selection(path)
        updated = False

        for row in rows:
            if row.get("ref") == row_ref:
                row["status"] = new_status
                row["errors"] = errors
                updated = True
                break

        if not updated:
            raise ValueError(f"Row with ref '{row_ref}' not found in CSV")

        # Rewrite entire CSV
        self._write_dict_rows(path, rows)

    def _write_dict_rows(self, path: Path, rows: List[Dict[str, Any]]) -> None:
        """Write list of dictionaries back to CSV."""
        if not rows:
            return

        try:
            with path.open("w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=rows[0].keys())
                writer.writeheader()
                for row in rows:
                    normalized_row = {
                        key: coerce_internal_value(str(key), value)
                        for key, value in row.items()
                    }
                    writer.writerow(normalized_row)
        except Exception as e:
            raise IOError(f"Cannot write CSV to {path}: {e}") from e
