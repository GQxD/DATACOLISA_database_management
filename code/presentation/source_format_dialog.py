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
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QAbstractItemView,
    QVBoxLayout,
    QWidget,
)
from PySide6.QtGui import QColor, QFont

import datacolisa_importer as core


HEADER_ALIASES = {
    "num_individu": ["numero individu", "num individu", "num", "nsach", "n sach", "ref", "reference"],
    "date_capture": ["date", "date capture", "date de capture"],
    "code_espece": ["code espece", "espece", "especes", "species"],
    "lac_riviere": ["lac riviere", "lac riviere", "riviere", "lieux secteurs", "lieu secteur"],
    "longueur_mm": ["longueur totale mm", "longueur mm", "lf", "lt"],
    "poids_g": ["poids g", "poids", "poidstot"],
    "maturite": ["maturite", "code maturite sexuelle"],
    "sexe": ["sexe", "code sexe"],
    "age_total": ["age total", "nht", "age"],
    "lieu_capture": ["lieu capture", "lieu de capture", "lieux secteurs", "lieu secteur"],
    "type_peche": ["type peche", "type peche engin", "engin"],
    "categorie": ["categorie", "categorie pecheur"],
    "pecheur": ["pecheur", "nom du pecheur"],
}


# Champs supplementaires (moins frequents, accessible via onglet dedie)
EXTRA_FIELD_LABELS = [
    ("code_unite_gestionnaire", "Code unite gestionnaire"),
    ("site_atelier", "Site Atelier"),
    ("numero_correspondant", "Numero du correspondant"),
    ("organisme", "Organisme preleveur"),
    ("categorie", "Categorie pecheur"),
    ("maille_mm", "Maille (mm)"),
    ("identifiant", "Identifiant"),
    ("presence_tache_gauche", "Presence de tache gauche"),
    ("presence_tache_droite", "Presence de tache droite"),
    ("nb_ecailles_stockage", "Nombre d'ecailles en stockage"),
    ("information_disponibilite", "Information disponibilite"),
    ("observation_disponibilite", "Observation disponibilite"),
    ("autre_echantillon_osseux", "Autre echantillon osseux"),
    ("ecailles_recuperees", "Ecailles recuperees"),
    ("observations", "Observations"),
    ("ecailles_brutes", "Ecailles brutes"),
    ("montees", "Montees"),
    ("empreintes", "Empreintes"),
    ("otolithes", "Otolithes"),
    ("engin_source", "Derivation type/categorie"),
    ("contexte", "Derivation pays"),
]

FIELD_LABELS = [
    ("num_individu", "Numero individu"),
    ("date_capture", "Date capture"),
    ("code_espece", "Code espece"),
    ("sous_espece", "Sous-espece"),
    ("pays_capture", "Pays capture"),
    ("lac_riviere", "Lac/riviere"),
    ("longueur_mm", "Longueur totale (mm)"),
    ("poids_g", "Poids (g)"),
    ("pecheur", "Nom du pecheur"),
    ("lieu_capture", "Lieu de capture / debarquement"),
    ("type_peche", "Type peche/engin"),
    ("code_stade", "Code stade"),
    ("maturite", "Code maturite sexuelle"),
    ("sexe", "Code sexe"),
    ("age_total", "Age total"),
    ("age_riviere", "Age riviere"),
    ("age_lac", "Age lac"),
    ("nombre_fraie", "Nombre de fraie"),
]

PRIMARY_FIELD_KEYS = [
    "num_individu",
    "date_capture",
    "code_espece",
    "pays_capture",
    "lac_riviere",
    "longueur_mm",
    "poids_g",
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
            key: None for key, _label in FIELD_LABELS + EXTRA_FIELD_LABELS
        }
        self.assignment_labels: Dict[str, QLabel] = {}
        self.sheet_names = core.get_workbook_sheet_names(self.source_path)
        self.sheet_name = current_sheet if current_sheet in self.sheet_names else (self.sheet_names[0] if self.sheet_names else "")
        self.source_rows: list[list[Any]] = []
        self.sheet_combo: QComboBox | None = None
        self.header_row_spin: QSpinBox | None = None
        self.setWindowTitle("Autre type de source")
        self.setMinimumSize(1280, 760)

        # Inherit parent's stylesheet for proper theme support (light/dark mode)
        if parent and parent.styleSheet():
            self.setStyleSheet(parent.styleSheet())

        self._build_ui(current_mode)
        self._reload_headers()

    def _guess_header_index(self, rows: list[list[Any]]) -> int:
        saved_header_row = self.current_mapping.get("header_row")
        if saved_header_row:
            saved_index = max(0, int(saved_header_row or 1) - 1)
            if saved_index < len(rows):
                normalized = [core.normalize_header_name(value) for value in rows[saved_index]]
                aliases = {alias for values in HEADER_ALIASES.values() for alias in values}
                if any(value in aliases for value in normalized):
                    return saved_index

        best_index = 0
        best_score = -1
        aliases = {alias for values in HEADER_ALIASES.values() for alias in values}
        for index, row in enumerate(rows[:50]):
            normalized = [core.normalize_header_name(value) for value in row]
            score = sum(1 for value in normalized if value in aliases)
            # Les anciens fichiers ont parfois "NSach" + "LF" + "NHT".
            if "nsach" in normalized:
                score += 4
            if "date" in normalized:
                score += 2
            if "lf" in normalized or "lt" in normalized:
                score += 1
            if score > best_score:
                best_score = score
                best_index = index
        return best_index

    def _auto_assign_columns(self, headers: list[Any], rows: list[list[Any]], header_index: int) -> None:
        existing_columns = (self.current_mapping.get("columns", {}) or {}).copy()
        if existing_columns:
            for key in self.assigned_columns:
                self.assigned_columns[key] = existing_columns.get(key)

        normalized_headers = [core.normalize_header_name(value) for value in headers]
        alias_columns: Dict[str, int] = {}
        for key, aliases in HEADER_ALIASES.items():
            for alias in aliases:
                if alias in normalized_headers:
                    alias_columns[key] = normalized_headers.index(alias)
                    break

        # Si le fichier contient un vrai en-tete ESPECE, on le prefere toujours
        # a une ancienne correspondance ou a une colonne sans titre contenant "Truite".
        if "code_espece" in alias_columns:
            self.assigned_columns["code_espece"] = alias_columns["code_espece"]

        for key, aliases in HEADER_ALIASES.items():
            if self.assigned_columns.get(key) is not None:
                continue
            if key in alias_columns:
                self.assigned_columns[key] = alias_columns[key]

        if self.assigned_columns.get("code_espece") is None:
            preview_rows = rows[header_index + 1: header_index + 31]
            for col_index in range(max((len(row) for row in preview_rows), default=0)):
                values = [
                    core.normalize(row[col_index]).lower()
                    for row in preview_rows
                    if col_index < len(row)
                ]
                species_hits = sum(1 for value in values if value in {"truite", "saumon", "ombre", "omble"})
                if species_hits >= 3:
                    self.assigned_columns["code_espece"] = col_index
                    break

    def _column_display_label(self, column_index: int) -> str:
        header_item = self.preview_table.horizontalHeaderItem(column_index)
        header_text = header_item.text() if header_item else ""
        if header_text:
            return header_text
        for row_index in range(self.preview_table.rowCount()):
            item = self.preview_table.item(row_index, column_index)
            value = item.text().strip() if item else ""
            if value:
                return f"{header_text or f'Col {column_index + 1}'} ({value})"
        return header_text or f"Col {column_index + 1}"

    def _excel_column_label(self, index: int) -> str:
        label = ""
        number = index + 1
        while number:
            number, remainder = divmod(number - 1, 26)
            label = chr(65 + remainder) + label
        return label

    def _build_ui(self, current_mode: str) -> None:
        main = QVBoxLayout(self)
        main.setSpacing(10)

        help_label = QLabel(
            "1. Choisis a droite le champ COLISA en cours a renseigner.\n"
            "2. Clique a gauche sur la colonne correspondante dans ton fichier source.\n"
            "3. Les champs Code type echantillon, Code echantillon et Numero individu sont automatiquement geres ou derives par l'application.\n"
            "4. Les champs de derivation servent a recalculer automatiquement le pays, le type de peche et la categorie si besoin."
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

        top_form = QFormLayout()
        if not self.force_custom:
            top_form.addRow("Format", self.mode_combo)

        self.sheet_combo = QComboBox()
        self.sheet_combo.addItems(self.sheet_names)
        if self.sheet_name:
            self.sheet_combo.setCurrentText(self.sheet_name)
        self.sheet_combo.currentTextChanged.connect(self._on_sheet_changed)
        top_form.addRow("Onglet Excel", self.sheet_combo)

        self.header_row_spin = QSpinBox()
        self.header_row_spin.setMinimum(1)
        self.header_row_spin.setMaximum(99999)
        self.header_row_spin.setValue(max(1, int(self.current_mapping.get("header_row", 1) or 1)))
        self.header_row_spin.valueChanged.connect(self._on_header_row_changed)
        top_form.addRow("Ligne d'en-tete", self.header_row_spin)
        main.addLayout(top_form)

        content = QWidget()
        content_layout = QHBoxLayout(content)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(12)

        preview_group = QGroupBox("Feuille Excel source")
        preview_layout = QVBoxLayout(preview_group)
        preview_layout.setContentsMargins(8, 8, 8, 8)
        self.preview_table = QTableWidget()
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.preview_table.setSelectionBehavior(QAbstractItemView.SelectColumns)
        self.preview_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.preview_table.horizontalHeader().sectionClicked.connect(self._apply_selected_preview_column)
        self.preview_table.cellClicked.connect(lambda _row, col: self._apply_selected_preview_column(col))
        preview_help = QLabel(
            "Clique sur le titre d'une colonne pour lier ce champ au fichier COLISA en cours. "
            "Si tu veux annuler une correspondance, sélectionne le champ puis clique sur Effacer."
        )
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
        for key, label in FIELD_LABELS + EXTRA_FIELD_LABELS:
            self.target_field_combo.addItem(label, key)
        mapping_form.addRow("Champ COLISA", self.target_field_combo)
        mapping_layout.addLayout(mapping_form)

        button_row = QHBoxLayout()
        btn_clear = QPushButton("Effacer la correspondance")
        btn_reset = QPushButton("Reinitialiser tout")
        btn_clear.clicked.connect(self._clear_current_mapping)
        btn_reset.clicked.connect(self._reset_mappings)
        button_row.addWidget(btn_clear)
        button_row.addWidget(btn_reset)
        button_row.addStretch()
        mapping_layout.addLayout(button_row)

        # ── Onglet 1 : champs principaux ────────────────────────────────────
        essential_group = QGroupBox("Champs essentiels")
        essential_layout = QGridLayout(essential_group)
        essential_layout.setContentsMargins(8, 8, 8, 8)
        essential_layout.setHorizontalSpacing(10)
        essential_layout.setVerticalSpacing(6)

        optional_group = QGroupBox("Autres champs principaux")
        optional_layout = QGridLayout(optional_group)
        optional_layout.setContentsMargins(8, 8, 8, 8)
        optional_layout.setHorizontalSpacing(10)
        optional_layout.setVerticalSpacing(6)

        essential_row = 0
        optional_row = 0
        for key, label in FIELD_LABELS:
            name_label = QLabel(label)
            value_label = QLabel("Aucune")
            value_label.setObjectName("chosenValue")
            self.assignment_labels[key] = value_label
            if key in PRIMARY_FIELD_KEYS:
                essential_layout.addWidget(name_label, essential_row, 0)
                essential_layout.addWidget(value_label, essential_row, 1)
                essential_row += 1
            else:
                optional_layout.addWidget(name_label, optional_row, 0)
                optional_layout.addWidget(value_label, optional_row, 1)
                optional_row += 1

        tab1_widget = QWidget()
        tab1_layout = QVBoxLayout(tab1_widget)
        tab1_layout.setContentsMargins(0, 4, 0, 0)
        tab1_layout.setSpacing(8)
        tab1_layout.addWidget(essential_group)
        tab1_layout.addWidget(optional_group)
        tab1_layout.addStretch()

        # ── Onglet 2 : champs supplementaires ───────────────────────────────
        extra_group = QGroupBox("Champs supplementaires")
        extra_layout = QGridLayout(extra_group)
        extra_layout.setContentsMargins(8, 8, 8, 8)
        extra_layout.setHorizontalSpacing(10)
        extra_layout.setVerticalSpacing(6)

        extra_row = 0
        for key, label in EXTRA_FIELD_LABELS:
            name_label = QLabel(label)
            value_label = QLabel("Aucune")
            value_label.setObjectName("chosenValue")
            self.assignment_labels[key] = value_label
            extra_layout.addWidget(name_label, extra_row, 0)
            extra_layout.addWidget(value_label, extra_row, 1)
            extra_row += 1

        tab2_widget = QWidget()
        tab2_layout = QVBoxLayout(tab2_widget)
        tab2_layout.setContentsMargins(0, 4, 0, 0)
        tab2_layout.addWidget(extra_group)
        tab2_layout.addStretch()

        # ── QTabWidget ───────────────────────────────────────────────────────
        tabs = QTabWidget()
        tabs.addTab(tab1_widget, "Principaux")
        tabs.addTab(tab2_widget, "Supplementaires")

        assigned_scroll = QScrollArea()
        assigned_scroll.setWidgetResizable(True)
        assigned_scroll.setWidget(tabs)
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

    def _on_sheet_changed(self, sheet_name: str) -> None:
        self.sheet_name = sheet_name.strip()
        self.current_mapping.pop("header_row", None)
        self.current_mapping["columns"] = {}
        for key in self.assigned_columns:
            self.assigned_columns[key] = None
        self._reload_headers()

    def _on_header_row_changed(self, value: int) -> None:
        self.current_mapping["header_row"] = int(value)
        self.current_mapping["columns"] = {
            key: value
            for key, value in self.assigned_columns.items()
            if value is not None
        }
        self._reload_headers()

    def _reload_headers(self) -> None:
        sheet_name = self.sheet_name.strip()
        if not sheet_name:
            return

        rows, _ = core.read_any_source_rows(self.source_path, sheet_name)
        self.source_rows = rows
        header_index = self._guess_header_index(rows)
        self.current_mapping["header_row"] = header_index + 1
        if self.header_row_spin is not None:
            self.header_row_spin.blockSignals(True)
            self.header_row_spin.setMaximum(max(1, len(rows)))
            self.header_row_spin.setValue(header_index + 1)
            self.header_row_spin.blockSignals(False)
        headers = rows[header_index] if header_index < len(rows) else []
        self._fill_preview_table(rows, headers, header_index)
        self._auto_assign_columns(headers, rows, header_index)
        self._refresh_assignment_labels()

    def _fill_preview_table(self, rows: list[list[Any]], headers: list[Any], header_index: int) -> None:
        preview_rows = rows
        column_count = max((len(r) for r in preview_rows), default=0)
        self.preview_table.clear()
        self.preview_table.setColumnCount(column_count)
        self.preview_table.setRowCount(len(preview_rows))

        horizontal_headers = []
        for idx in range(column_count):
            excel_label = self._excel_column_label(idx)
            label = core.normalize(headers[idx]) if idx < len(headers) else ""
            horizontal_headers.append(f"{excel_label} - {label}" if label else excel_label)
        self.preview_table.setHorizontalHeaderLabels(horizontal_headers)
        self.preview_table.setVerticalHeaderLabels([str(i + 1) for i in range(len(preview_rows))])

        header_font = QFont()
        header_font.setBold(True)
        header_background = QColor("#fff2cc")

        for row_index, row in enumerate(preview_rows):
            for col_index in range(column_count):
                value = row[col_index] if col_index < len(row) else ""
                item = QTableWidgetItem(core.normalize(value))
                if row_index == header_index:
                    item.setBackground(header_background)
                    item.setFont(header_font)
                self.preview_table.setItem(row_index, col_index, item)

        self.preview_table.resizeColumnsToContents()
        header_item = self.preview_table.item(header_index, 0)
        if header_item is not None:
            self.preview_table.scrollToItem(header_item, QAbstractItemView.PositionAtCenter)

    def _apply_selected_preview_column(self, column_index: int) -> None:
        field_key = self.target_field_combo.currentData()
        self.assigned_columns[str(field_key)] = column_index
        self._refresh_assignment_labels()

    def _clear_current_mapping(self) -> None:
        field_key = self.target_field_combo.currentData()
        if field_key is None:
            return
        self.assigned_columns[str(field_key)] = None
        self._refresh_assignment_labels()

    def _reset_mappings(self) -> None:
        for key in self.assigned_columns:
            self.assigned_columns[key] = None
        self._refresh_assignment_labels()

    def _refresh_assignment_labels(self) -> None:
        headers = self.preview_table.horizontalHeaderItem
        for key, _label in FIELD_LABELS + EXTRA_FIELD_LABELS:
            value_label = self.assignment_labels.get(key)
            if value_label is None:
                continue
            column_index = self.assigned_columns.get(key)
            if column_index is None:
                value_label.setText("Aucune")
                continue
            value_label.setText(self._column_display_label(column_index))

    def _accept(self) -> None:
        if not self.force_custom and self.mode_combo.currentData() != "custom":
            self.accept()
            return

        ref_col = self.assigned_columns.get("num_individu")
        if ref_col is None:
            QMessageBox.warning(self, "Autre type de source", "Choisis au moins la colonne Numero individu.")
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
