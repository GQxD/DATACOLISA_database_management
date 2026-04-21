"""Dialog for choosing PAC final or mapping another Excel source format."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Dict

from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QGroupBox,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QFormLayout,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QAbstractItemView,
    QVBoxLayout,
    QWidget,
)

import datacolisa_importer as core


FIELD_LABELS = [
    ("code_unite_gestionnaire", "Code unite gestionnaire"),
    ("site_atelier", "Site Atelier"),
    ("numero_correspondant", "Numero du correspondant"),
    ("code_type_echantillon", "Code type echantillon"),
    ("code_echantillon", "Code echantillon"),
    ("code_espece", "Code espece"),
    ("sous_espece", "Sous-espece"),
    ("organisme", "Organisme preleveur"),
    ("pecheur", "Nom du pecheur"),
    ("pays_capture", "Pays capture"),
    ("date_capture", "Date capture"),
    ("lac_riviere", "Lac/riviere"),
    ("lieu_capture", "Lieu de capture / debarquement"),
    ("categorie", "Categorie pecheur"),
    ("type_peche", "Type peche/engin"),
    ("maille_mm", "Maille (mm)"),
    ("identifiant", "Identifiant"),
    ("num_individu", "Numero individu"),
    ("longueur_mm", "Longueur totale (mm)"),
    ("poids_g", "Poids (g)"),
    ("code_stade", "Code stade"),
    ("maturite", "Code maturite sexuelle"),
    ("sexe", "Code sexe"),
    ("presence_tache_gauche", "Presence de tache gauche (0 si absente)"),
    ("presence_tache_droite", "Presence de tache droite (0 si absente)"),
    ("nb_ecailles_stockage", "Nombre d'ecailles en etat stockage"),
    ("information_disponibilite", "Information disponibilite"),
    ("observation_disponibilite", "Observation disponibilite"),
    ("autre_echantillon_osseux", "Autre echantillon osseux collecte"),
    ("age_total", "Age total"),
    ("age_riviere", "Age riviere"),
    ("age_lac", "Age lac"),
    ("nombre_fraie", "Nombre de fraie"),
    ("ecailles_recuperees", "Ecailles recuperees"),
    ("observations", "Observations"),
    ("ecailles_brutes", "Ecailles brutes"),
    ("montees", "Montees"),
    ("empreintes", "Empreintes"),
    ("otolithes", "Otolithes"),
    ("engin_source", "Colonne source pour derivation type/categorie"),
    ("contexte", "Colonne source pour derivation pays"),
]


class SourceFormatDialog(QDialog):
    def __init__(
        self,
        source_path: Path,
        current_mode: str = "pac_final",
        current_sheet: str = "",
        current_mapping: Dict[str, Any] | None = None,
        force_custom: bool = False,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.source_path = source_path
        self.force_custom = force_custom
        self.current_mapping = current_mapping or {}
        self.header_choices: list[tuple[str, int | None]] = []
        self.assigned_columns: Dict[str, int | None] = {
            key: None for key, _label in FIELD_LABELS
        }
        self.assignment_labels: Dict[str, QLabel] = {}
        self.sheet_names = core.get_workbook_sheet_names(self.source_path)
        self.sheet_name = current_sheet if current_sheet in self.sheet_names else (self.sheet_names[0] if self.sheet_names else "")
        self.source_rows: list[list[Any]] = []
        self.setWindowTitle("Autre type de source")
        self.setMinimumSize(1100, 620)

        # Inherit parent's stylesheet for proper theme support (light/dark mode)
        if parent and parent.styleSheet():
            self.setStyleSheet(parent.styleSheet())

        self._build_ui(current_mode)
        self._reload_headers()

    def _build_ui(self, current_mode: str) -> None:
        main = QVBoxLayout(self)
        main.setSpacing(10)

        help_label = QLabel(
            "1. Choisis a droite le champ COLISA en cours a renseigner.\n"
            "2. Clique a gauche sur la colonne correspondante dans ton fichier source.\n"
            "3. Les champs de derivation servent a recalculer automatiquement le pays, le type de peche et la categorie si besoin."
        )
        help_label.setObjectName("stepHelp")
        help_label.setWordWrap(True)
        main.addWidget(help_label)

        if not self.force_custom:
            btn_open = QPushButton("Ouvrir le fichier Excel")
            btn_open.clicked.connect(self._open_source_file)
            main.addWidget(btn_open)

        self.mode_combo = QComboBox()
        self.mode_combo.addItem("PAC final", "pac_final")
        self.mode_combo.addItem("Autre type de source", "custom")
        idx = max(0, self.mode_combo.findData("custom" if self.force_custom else current_mode))
        self.mode_combo.setCurrentIndex(idx)

        if not self.force_custom:
            top_form = QFormLayout()
            top_form.addRow("Format", self.mode_combo)
            main.addLayout(top_form)

        if self.sheet_name:
            sheet_label = QLabel(f"Onglet : {self.sheet_name}")
            sheet_label.setObjectName("sheetInfo")
            main.addWidget(sheet_label)

        content = QWidget()
        content_layout = QHBoxLayout(content)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(12)

        preview_group = QGroupBox("Colonnes du fichier source")
        preview_layout = QVBoxLayout(preview_group)
        preview_layout.setContentsMargins(8, 8, 8, 8)
        self.preview_table = QTableWidget()
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.preview_table.setSelectionBehavior(QAbstractItemView.SelectColumns)
        self.preview_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.preview_table.horizontalHeader().sectionClicked.connect(self._apply_selected_preview_column)
        preview_help = QLabel("Clique sur le titre d'une colonne pour lier ce champ au fichier COLISA en cours.")
        preview_help.setWordWrap(True)
        preview_layout.addWidget(preview_help)
        preview_layout.addWidget(self.preview_table, 1)

        self.mapping_group = QGroupBox("Champs a envoyer vers COLISA en cours")
        mapping_layout = QVBoxLayout(self.mapping_group)
        mapping_form = QFormLayout()
        mapping_form.setSpacing(6)
        self.target_field_combo = QComboBox()
        self.target_field_combo.setMaxVisibleItems(20)
        self.target_field_combo.setStyleSheet("QComboBox { min-height: 30px; }")
        for key, label in FIELD_LABELS:
            self.target_field_combo.addItem(label, key)
        mapping_form.addRow("Champ COLISA", self.target_field_combo)
        mapping_layout.addLayout(mapping_form)

        assigned_group = QGroupBox("Correspondances retenues")
        assigned_grid = QGridLayout(assigned_group)
        assigned_grid.setContentsMargins(8, 8, 8, 8)
        assigned_grid.setHorizontalSpacing(10)
        assigned_grid.setVerticalSpacing(6)
        for row_index, (key, label) in enumerate(FIELD_LABELS):
            name_label = QLabel(label)
            value_label = QLabel("Aucune")
            value_label.setObjectName("chosenValue")
            self.assignment_labels[key] = value_label
            assigned_grid.addWidget(name_label, row_index, 0)
            assigned_grid.addWidget(value_label, row_index, 1)

        assigned_scroll = QScrollArea()
        assigned_scroll.setWidgetResizable(True)
        assigned_scroll.setWidget(assigned_group)
        assigned_scroll.setMinimumHeight(320)
        assigned_scroll.setStyleSheet("QScrollArea { border: none; }")
        mapping_layout.addWidget(assigned_scroll)
        content_layout.addWidget(preview_group, 3)
        content_layout.addWidget(self.mapping_group, 2)
        main.addWidget(content, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_cancel = QPushButton("Annuler")
        btn_ok = QPushButton("Enregistrer")
        btn_cancel.clicked.connect(self.reject)
        btn_ok.clicked.connect(self._accept)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        main.addLayout(btn_row)

        self.mode_combo.currentIndexChanged.connect(self._update_enabled_state)
        self._update_enabled_state()

    def _update_enabled_state(self) -> None:
        is_custom = self.force_custom or self.mode_combo.currentData() == "custom"
        self.mapping_group.setEnabled(is_custom)
        self.preview_table.setEnabled(is_custom)
        self.target_field_combo.setEnabled(is_custom)

    def _reload_headers(self) -> None:
        sheet_name = self.sheet_name.strip()
        if not sheet_name:
            return

        rows, _ = core.read_any_source_rows(self.source_path, sheet_name)
        self.source_rows = rows
        header_index = max(0, int(self.current_mapping.get("header_row", 1) or 1) - 1)
        headers = rows[header_index] if header_index < len(rows) else []
        self._fill_preview_table(rows, headers, header_index)
        previous_columns = (self.current_mapping.get("columns", {}) or {}).copy()
        for key in self.assigned_columns:
            self.assigned_columns[key] = previous_columns.get(key)
        self._refresh_assignment_labels()

    def _fill_preview_table(self, rows: list[list[Any]], headers: list[Any], header_index: int) -> None:
        preview_rows = rows[header_index + 1: header_index + 11]
        column_count = max(len(headers), max((len(r) for r in preview_rows), default=0))
        self.preview_table.clear()
        self.preview_table.setColumnCount(column_count)
        self.preview_table.setRowCount(len(preview_rows))

        horizontal_headers = []
        for idx in range(column_count):
            label = core.normalize(headers[idx]) if idx < len(headers) else ""
            horizontal_headers.append(label or f"Col {idx + 1}")
        self.preview_table.setHorizontalHeaderLabels(horizontal_headers)

        for row_index, row in enumerate(preview_rows):
            for col_index in range(column_count):
                value = row[col_index] if col_index < len(row) else ""
                item = QTableWidgetItem(core.normalize(value))
                self.preview_table.setItem(row_index, col_index, item)

        self.preview_table.resizeColumnsToContents()

    def _apply_selected_preview_column(self, column_index: int) -> None:
        field_key = self.target_field_combo.currentData()
        self.assigned_columns[str(field_key)] = column_index
        self._refresh_assignment_labels()

    def _refresh_assignment_labels(self) -> None:
        headers = self.preview_table.horizontalHeaderItem
        for key, _label in FIELD_LABELS:
            value_label = self.assignment_labels.get(key)
            if value_label is None:
                continue
            column_index = self.assigned_columns.get(key)
            if column_index is None:
                value_label.setText("Aucune")
                continue
            header_item = headers(column_index)
            header_text = header_item.text() if header_item else f"Col {column_index + 1}"
            value_label.setText(header_text)

    def _accept(self) -> None:
        if not self.force_custom and self.mode_combo.currentData() != "custom":
            self.accept()
            return

        ref_col = self.assigned_columns.get("num_individu")
        if ref_col is None:
            QMessageBox.warning(self, "Autre type de source", "Choisis au moins la colonne Numero.")
            return
        self.accept()

    def _open_source_file(self) -> None:
        try:
            os.startfile(str(self.source_path))
        except Exception as exc:
            QMessageBox.warning(self, "Fichier source", f"Impossible d'ouvrir le fichier:\n{exc}")

    def get_result(self) -> Dict[str, Any]:
        mode = "custom" if self.force_custom else self.mode_combo.currentData()
        result = {
            "mode": mode,
            "sheet_name": self.sheet_name,
            "mapping": {},
        }
        if mode == "custom":
            result["mapping"] = {
                "header_row": int(self.current_mapping.get("header_row", 1) or 1),
                "columns": {
                    key: value
                    for key, value in self.assigned_columns.items()
                    if value is not None
                },
            }
        return result
