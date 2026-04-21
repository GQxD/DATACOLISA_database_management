"""Dialog for assigning sample types individually to selected rows."""

from __future__ import annotations

from typing import Any, Dict, List

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)


class SampleTypeAssignmentDialog(QDialog):
    """Allow editing sample type row by row for selected samples."""

    def __init__(
        self,
        rows: List[Dict[str, Any]],
        type_options: List[str],
        parent=None,
    ) -> None:
        super().__init__(parent)
        self._rows = rows
        self._type_options = [opt for opt in type_options if str(opt).strip()]
        self._type_labels = self._build_type_labels(self._type_options)
        self.setWindowTitle("Types d'echantillon")
        self.setMinimumSize(760, 520)

        if parent and parent.styleSheet():
            self.setStyleSheet(parent.styleSheet())

        self._build_ui()
        self._load_rows()

    def _build_ui(self) -> None:
        main = QVBoxLayout(self)
        main.setSpacing(10)

        help_label = QLabel(
            "Choisis le type d'echantillon pour chaque ligne selectionnee, puis enregistre."
        )
        help_label.setWordWrap(True)
        main.addWidget(help_label)

        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["Ref", "Numero individu", "Type echantillon"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        main.addWidget(self.table, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()

        btn_cancel = QPushButton("Annuler")
        btn_ok = QPushButton("Enregistrer")
        btn_cancel.clicked.connect(self.reject)
        btn_ok.clicked.connect(self.accept)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        main.addLayout(btn_row)

    def _load_rows(self) -> None:
        self.table.setRowCount(len(self._rows))
        for row_index, row_data in enumerate(self._rows):
            ref_item = QTableWidgetItem(str(row_data.get("ref", "")))
            num_item = QTableWidgetItem(str(row_data.get("num_individu", "")))
            ref_item.setFlags(ref_item.flags() & ~Qt.ItemIsEditable)
            num_item.setFlags(num_item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row_index, 0, ref_item)
            self.table.setItem(row_index, 1, num_item)

            combo = QComboBox()
            combo.setEditable(True)
            combo.setInsertPolicy(QComboBox.NoInsert)
            for value in self._type_options:
                label = self._type_labels.get(value, "")
                combo.addItem(f"{value} - {label}" if label else value, value)
            current_value = str(row_data.get("code_type_echantillon", ""))
            found = False
            for combo_index in range(combo.count()):
                if str(combo.itemData(combo_index) or "") == current_value:
                    combo.setCurrentIndex(combo_index)
                    found = True
                    break
            if not found:
                combo.setEditText(current_value)
            self.table.setCellWidget(row_index, 2, combo)

        self.table.resizeColumnsToContents()

    def get_updated_rows(self) -> List[Dict[str, Any]]:
        updated_rows: List[Dict[str, Any]] = []
        for row_index, row_data in enumerate(self._rows):
            updated = dict(row_data)
            combo = self.table.cellWidget(row_index, 2)
            if isinstance(combo, QComboBox):
                data = combo.currentData()
                if data is not None and str(data).strip():
                    updated["code_type_echantillon"] = str(data).strip()
                else:
                    text = combo.currentText().strip()
                    updated["code_type_echantillon"] = text.split(" - ", 1)[0].strip() if " - " in text else text
            else:
                updated["code_type_echantillon"] = ""
            updated_rows.append(updated)
        return updated_rows

    @staticmethod
    def _build_type_labels(type_options: List[str]) -> Dict[str, str]:
        labels = {
            "BI": "Bile",
            "GN": "Bouche de poisson",
            "BN": "Branchie de poisson",
            "VN": "Colonne vertebrale de poisson",
            "HN": "Dos de poisson",
            "EC": "Ecaille de poisson",
            "ES": "Estomac",
            "FO": "Foie de poisson",
            "FN": "Fraction inconnue de poisson",
            "GO": "Gonade de poisson",
            "GR": "Graisse de poisson",
            "HM": "Hemocytes de poissons",
            "LE": "Levre de poisson",
            "MN": "Machoire de poisson",
            "MU": "Muscle de poisson",
            "MP": "Muscle et peau de poisson",
            "MI": "Muscle, muscle+tissu adipeux face interne de peau",
            "QN": "Nageoire caudale de poisson",
            "NN": "Nageoire de poisson",
            "DN": "Nageoire dorsale de poisson",
            "PN": "Nageoire pectorale de poisson",
            "YN": "Oeil de poisson",
            "ON": "Opercules",
            "XN": "Orifice anal de poisson",
            "XU": "Orifice urogenital de poisson",
            "OT": "Otolithes",
            "KN": "Pedoncule caudal de poisson",
            "CN": "Poisson entier",
            "WE": "Poisson tete et queue",
            "WV": "Poisson sans visceres ni gonades",
            "RE": "Rein de poisson",
            "NF": "Systeme nerveux de poisson",
            "TN": "Tete de poisson",
            "WN": "Tronc de poisson",
        }
        return {code: labels.get(code, "") for code in type_options}
