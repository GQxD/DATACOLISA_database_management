"""Import history repository for tracking import operations."""

from __future__ import annotations

import datetime as dt
import json
from pathlib import Path
from typing import Any, Dict, List


class HistoryRepository:
    """
    Handles reading and writing import history files.

    History files track which rows were imported, skipped, or had errors
    during import operations.
    """

    def save_history(
        self,
        history_rows: List[Dict[str, Any]],
        path: Path
    ) -> None:
        """
        Save import history to JSON file.

        Args:
            history_rows: List of dicts with 'ref', 'status', 'reason' keys
            path: Path where to save the history file

        Raises:
            IOError: If file cannot be written
        """
        payload = {
            "updated_at": dt.datetime.now().isoformat(timespec="seconds"),
            "rows": history_rows,
        }

        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2),
                encoding="utf-8"
            )
        except Exception as e:
            raise IOError(f"Cannot write history to {path}: {e}") from e

    def load_history(self, path: Path) -> Dict[str, Any]:
        """
        Load import history from JSON file.

        Args:
            path: Path to the history file

        Returns:
            Dict with 'updated_at' and 'rows' keys

        Raises:
            FileNotFoundError: If file doesn't exist
            IOError: If file cannot be read
            ValueError: If JSON is invalid
        """
        if not path.exists():
            raise FileNotFoundError(f"History file not found: {path}")

        try:
            content = path.read_text(encoding="utf-8")
            payload = json.loads(content)

            # Validate structure
            if "rows" not in payload:
                raise ValueError("History file missing 'rows' key")

            return payload

        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON in history file {path}: {e}") from e
        except Exception as e:
            raise IOError(f"Cannot read history from {path}: {e}") from e

    def get_pending_rows(self, path: Path) -> List[Dict[str, Any]]:
        """
        Get rows that need to be re-imported.

        Returns rows with status 'a_reimporter' or 'non_importe_manuel'.

        Args:
            path: Path to the history file

        Returns:
            List of row dicts that need re-import

        Raises:
            FileNotFoundError: If file doesn't exist
            IOError: If file cannot be read
        """
        history = self.load_history(path)
        rows = history.get("rows", [])

        pending_statuses = ["a_reimporter", "non_importe_manuel"]
        return [
            row for row in rows
            if row.get("status") in pending_statuses
        ]

    def get_rows_by_status(
        self,
        path: Path,
        status: str
    ) -> List[Dict[str, Any]]:
        """
        Get all rows with a specific status.

        Args:
            path: Path to the history file
            status: Status to filter by

        Returns:
            List of row dicts with the specified status

        Raises:
            FileNotFoundError: If file doesn't exist
            IOError: If file cannot be read
        """
        history = self.load_history(path)
        rows = history.get("rows", [])

        return [
            row for row in rows
            if row.get("status") == status
        ]

    def append_rows(
        self,
        path: Path,
        new_rows: List[Dict[str, Any]]
    ) -> None:
        """
        Append new rows to existing history file.

        Args:
            path: Path to the history file
            new_rows: Rows to append

        Raises:
            IOError: If file cannot be read/written
        """
        try:
            history = self.load_history(path)
        except FileNotFoundError:
            # Create new history if doesn't exist
            history = {"updated_at": "", "rows": []}

        history["rows"].extend(new_rows)
        self.save_history(history["rows"], path)

    def get_statistics(self, path: Path) -> Dict[str, int]:
        """
        Get import statistics from history.

        Args:
            path: Path to the history file

        Returns:
            Dict with counts by status

        Raises:
            FileNotFoundError: If file doesn't exist
            IOError: If file cannot be read
        """
        history = self.load_history(path)
        rows = history.get("rows", [])

        stats: Dict[str, int] = {}
        for row in rows:
            status = row.get("status", "unknown")
            stats[status] = stats.get(status, 0) + 1

        return stats
