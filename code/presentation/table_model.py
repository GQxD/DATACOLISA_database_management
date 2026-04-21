"""Table model for displaying import data with high performance."""

from __future__ import annotations

from typing import Any, Dict, List, Optional
from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import QApplication


# Column definitions (must match UI expectations)
COLUMNS = [
    "selected",
    "ref",
    "code_type_echantillon",
    "categorie",
    "type_peche",
    "autre_oss",
    "ecailles_brutes",
    "montees",
    "empreintes",
    "otolithes",
    "observation_disponibilite",
    "num_individu",
    "date_capture",
    "code_espece",
    "lac_riviere",
    "pays_capture",
    "pecheur",
    "longueur_mm",
    "poids_g",
    "maturite",
    "sexe",
    "age_total",
    "status",
    "errors",
]

DISPLAY_HEADERS = {
    "selected": "Selection",
    "ref": "Ref",
    "code_type_echantillon": "Code type echantillon",
    "categorie": "Categorie",
    "type_peche": "Type peche",
    "autre_oss": "Autre oss",
    "ecailles_brutes": "Ecailles brutes",
    "montees": "Montees",
    "empreintes": "Empreintes",
    "otolithes": "Otolithes",
    "observation_disponibilite": "Observation disponibilite",
    "num_individu": "Numero individu",
    "date_capture": "Date capture",
    "code_espece": "Code espece",
    "lac_riviere": "Lac/riviere",
    "pays_capture": "Pays capture",
    "pecheur": "Pecheur",
    "longueur_mm": "Longueur (mm)",
    "poids_g": "Poids (g)",
    "maturite": "Maturite",
    "sexe": "Sexe",
    "age_total": "Age total",
    "status": "Statut",
    "errors": "Erreurs",
}

# Non-editable columns
NON_EDITABLE_COLUMNS = {"selected", "ref", "status", "errors"}


class ImportTableModel(QAbstractTableModel):
    """
    High-performance table model for import data.

    Replaces QTableWidget to fix performance issues with 100+ rows.
    Uses QAbstractTableModel pattern for efficient rendering:
    - Data stored once in model, not duplicated in widgets
    - Only visible cells are rendered (lazy rendering)
    - No QComboBox created per cell (only on edit)

    Performance: 100+ rows render in <1s instead of 10s+
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self._rows: List[Dict[str, Any]] = []
        self._columns = COLUMNS

    def rowCount(self, parent=QModelIndex()) -> int:
        """Return number of rows."""
        if parent.isValid():
            return 0
        return len(self._rows)

    def columnCount(self, parent=QModelIndex()) -> int:
        """Return number of columns."""
        if parent.isValid():
            return 0
        return len(self._columns)

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole) -> Any:
        """Return data for given index and role."""
        if not index.isValid():
            return None

        row = index.row()
        col = index.column()

        if row < 0 or row >= len(self._rows):
            return None
        if col < 0 or col >= len(self._columns):
            return None

        col_name = self._columns[col]
        row_data = self._rows[row]

        if role == Qt.DisplayRole or role == Qt.EditRole:
            value = row_data.get(col_name, "")

            # Special handling for selected column (checkbox)
            if col_name == "selected":
                return None  # Checkbox handles this

            # Convert to string for display
            return str(value) if value is not None else ""

        elif role == Qt.CheckStateRole:
            # Handle checkbox for "selected" column
            if col_name == "selected":
                is_selected = row_data.get(col_name, True)
                return Qt.Checked if is_selected else Qt.Unchecked
            return None

        elif role == Qt.BackgroundRole:
            # Highlight rows with errors using theme-aware colors
            errors = row_data.get("errors", "")
            if errors:
                app = QApplication.instance()
                dark = (
                    app is not None
                    and app.palette().color(app.palette().ColorRole.Window).lightness() < 128
                )
                # Mode nuit : rouge sombre lisible ; mode jour : rouge clair lisible
                return QBrush(QColor("#5c1a1a") if dark else QColor("#fde8e8"))
            return None

        return None

    def setData(self, index: QModelIndex, value: Any, role: int = Qt.EditRole) -> bool:
        """Set data for given index."""
        if not index.isValid():
            return False

        row = index.row()
        col = index.column()

        if row < 0 or row >= len(self._rows):
            return False
        if col < 0 or col >= len(self._columns):
            return False

        col_name = self._columns[col]

        if role == Qt.EditRole:
            self._rows[row][col_name] = value
            self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
            return True

        elif role == Qt.CheckStateRole:
            if col_name == "selected":
                self._rows[row][col_name] = (value == Qt.Checked)
                self.dataChanged.emit(index, index, [Qt.CheckStateRole, Qt.DisplayRole])
                return True

        return False

    def flags(self, index: QModelIndex) -> Qt.ItemFlags:
        """Return item flags for given index."""
        if not index.isValid():
            return Qt.NoItemFlags

        col_name = self._columns[index.column()]

        flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable

        # Selected column is checkable
        if col_name == "selected":
            return Qt.ItemIsEnabled | Qt.ItemIsUserCheckable
        # Some columns are non-editable
        elif col_name not in NON_EDITABLE_COLUMNS:
            flags |= Qt.ItemIsEditable

        return flags

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole) -> Any:
        """Return header data."""
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                if 0 <= section < len(self._columns):
                    return DISPLAY_HEADERS.get(self._columns[section], self._columns[section])
            elif orientation == Qt.Vertical:
                return str(section + 1)
        return None

    # Custom methods for data management

    def set_rows(self, rows: List[Dict[str, Any]]) -> None:
        """
        Replace all rows with new data.

        This is the main method to populate the table.
        Uses beginResetModel/endResetModel for efficient bulk updates.
        """
        self.beginResetModel()
        self._rows = rows
        self.endResetModel()

    def get_rows(self) -> List[Dict[str, Any]]:
        """Return all rows."""
        return self._rows

    def get_row(self, row_index: int) -> Optional[Dict[str, Any]]:
        """Return a single row by index."""
        if 0 <= row_index < len(self._rows):
            return self._rows[row_index]
        return None

    def update_row(self, row_index: int, row_data: Dict[str, Any]) -> bool:
        """Update a single row."""
        if 0 <= row_index < len(self._rows):
            self._rows[row_index] = row_data
            # Emit dataChanged for entire row
            left = self.index(row_index, 0)
            right = self.index(row_index, len(self._columns) - 1)
            self.dataChanged.emit(left, right, [Qt.DisplayRole, Qt.EditRole])
            return True
        return False

    def clear(self) -> None:
        """Clear all rows."""
        self.beginResetModel()
        self._rows = []
        self.endResetModel()

    def append_row(self, row_data: Dict[str, Any]) -> None:
        """Append a new row."""
        row_index = len(self._rows)
        self.beginInsertRows(QModelIndex(), row_index, row_index)
        self._rows.append(row_data)
        self.endInsertRows()

    def remove_rows(self, row_indices: List[int]) -> None:
        """Remove multiple rows by index."""
        # Sort in reverse to avoid index shifting issues
        for row_index in sorted(row_indices, reverse=True):
            if 0 <= row_index < len(self._rows):
                self.beginRemoveRows(QModelIndex(), row_index, row_index)
                self._rows.pop(row_index)
                self.endRemoveRows()

    def get_selected_rows(self) -> List[int]:
        """Return indices of selected rows."""
        selected = []
        for i, row in enumerate(self._rows):
            if row.get("selected", False):
                selected.append(i)
        return selected

    def set_all_selected(self, selected: bool) -> None:
        """Select or deselect all rows."""
        if not self._rows:
            return

        for row in self._rows:
            row["selected"] = selected

        # Emit dataChanged for the entire "selected" column
        top_left = self.index(0, 0)  # Assuming "selected" is first column
        bottom_right = self.index(len(self._rows) - 1, 0)
        self.dataChanged.emit(top_left, bottom_right, [Qt.CheckStateRole, Qt.DisplayRole])
