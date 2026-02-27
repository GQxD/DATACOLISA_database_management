#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import os
import shutil
import sys
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path
from typing import Any, Dict, List

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QColor, QIcon, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

import datacolisa_importer as core

TYPE_PECHE_OPTIONS = ["", "LIGNE", "FILET", "TRAINE", "SONDE"]
CATEGORIE_OPTIONS = ["", "PRO", "AMATEUR", "SCIENTIFIQUE"]
OUI_NON_OPTIONS = ["", "OUI", "NON"]
OBSERVATION_OPTIONS = ["", "+", "++", "+++"]
NUMERIC_OPTIONS = [""] + [str(i) for i in range(1, 11)]
ECAILLES_BRUTES_OPTIONS = [""] + [str(i) for i in range(1, 21)]
APP_VERSION = "1.3.0"
APP_ORGANISATION = "INRAE"
APP_AUTHOR = "Quentin Godeaux"
SETTINGS_DIR = Path(os.environ.get("APPDATA", str(Path.home()))) / "DATACOLISA"
SETTINGS_FILE = SETTINGS_DIR / "ui_pyside6_settings.json"

COLS = [
    "selected",
    "include",
    "ref",
    "code_type_echantillon",
    "categorie",
    "type_peche",
    "autre_oss",
    "ecailles_brutes",
    "montees",
    "otolithes",
    "observation_disponibilite",
    "source_row",
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


def qitem(text: Any, editable: bool = True) -> QTableWidgetItem:
    it = QTableWidgetItem(str(text) if text is not None else "")
    if not editable:
        it.setFlags(it.flags() & ~Qt.ItemIsEditable)
    return it


def _resource_root() -> Path:
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            return Path(meipass)
    return Path(__file__).resolve().parent.parent


def _app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("DATACOLISA - POC PySide6")
        self.resize(1680, 920)

        self.rows: List[Dict[str, Any]] = []
        self.type_options: List[str] = []
        self.missing_codes: List[str] = []
        self._table_sync_guard = False
        self.dark_mode = False

        app_base = _app_base_dir()
        resource_base = _resource_root()
        self.source_path = app_base / "PacFinalTL14novembrel2012.xls"
        self.source_sheet = core.DEFAULT_SOURCE_SHEET
        self.target_path = resource_base / "COLISA_template_interne.xlsx"
        self.target_sheet = core.DEFAULT_TARGET_SHEET
        self.out_path = app_base / "COLISA_imported.xlsx"
        self.history_path = app_base / "import_history.json"
        self.selection_csv = app_base / "selection_import.csv"

        self._build_menu()

        root = QWidget()
        self.setCentralWidget(root)
        lay = QVBoxLayout(root)

        lay.addWidget(self._build_context_strip())
        lay.addWidget(self._build_top_panel(), 0, Qt.AlignLeft)
        lay.addWidget(self._build_metrics_panel(), 0, Qt.AlignLeft)
        lay.addWidget(self._build_bulk_panel())
        lay.addWidget(self._build_table())
        lay.addWidget(self._build_bottom_panel())

        self._load_settings()
        self._refresh_type_options()
        self._refresh_context_labels()
        self._apply_theme(self.dark_mode)
        self.switch_theme.blockSignals(True)
        self.switch_theme.setChecked(self.dark_mode)
        self.switch_theme.blockSignals(False)

    def _build_menu(self) -> None:
        mb = self.menuBar()
        menu_file = mb.addMenu("Fichier")

        act_source = QAction("Selectionner source .xls", self)
        act_source.triggered.connect(self.select_source_file)
        menu_file.addAction(act_source)

        act_source_sheet = QAction("Onglet source...", self)
        act_source_sheet.triggered.connect(self.set_source_sheet)
        menu_file.addAction(act_source_sheet)

        menu_file.addSeparator()

        act_out = QAction("Fichier de sortie...", self)
        act_out.triggered.connect(self.set_output_file)
        menu_file.addAction(act_out)

        act_hist = QAction("Fichier historique...", self)
        act_hist.triggered.connect(self.set_history_file)
        menu_file.addAction(act_hist)

        act_csv = QAction("Fichier selection CSV...", self)
        act_csv.triggered.connect(self.set_selection_csv)
        menu_file.addAction(act_csv)

        menu_file.addSeparator()
        act_about = QAction("A propos", self)
        act_about.triggered.connect(self.show_about)
        menu_file.addAction(act_about)

        menu_file.addSeparator()
        act_quit = QAction("Quitter", self)
        act_quit.triggered.connect(self.close)
        menu_file.addAction(act_quit)

    def _build_context_strip(self) -> QWidget:
        w = QWidget()
        h = QHBoxLayout(w)
        h.setContentsMargins(0, 0, 0, 0)

        self.lbl_source = QLabel()
        self.lbl_target = QLabel()
        self.lbl_paths = QLabel()

        info = QWidget()
        info_v = QVBoxLayout(info)
        info_v.setContentsMargins(0, 0, 0, 0)
        info_v.addWidget(self.lbl_source)
        info_v.addWidget(self.lbl_target)
        info_v.addWidget(self.lbl_paths)

        self.switch_theme = QCheckBox("Mode nuit")
        self.switch_theme.setObjectName("themeSwitch")
        self.switch_theme.toggled.connect(self._on_theme_toggled)

        h.addWidget(info, 3)
        h.addWidget(self.switch_theme, 0)
        return w

    def _build_top_panel(self) -> QWidget:
        box = QGroupBox("Chargement plage")
        box.setMaximumWidth(560)
        box.setMaximumHeight(230)
        box.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Fixed)
        g = QGridLayout(box)
        g.setHorizontalSpacing(8)
        g.setVerticalSpacing(6)
        g.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        g.setColumnStretch(0, 0)
        g.setColumnStretch(1, 0)
        g.setColumnStretch(2, 0)
        g.setColumnStretch(3, 1)

        self.ed_start = QLineEdit("CA961")
        self.ed_end = QLineEdit("CA989")

        self.cb_duplicate = QComboBox(); self.cb_duplicate.addItems(["alert", "ignore", "replace"])

        self.ed_org = QLineEdit("INRAE")
        self.ed_country = QLineEdit("France")

        # Ajustements visuels: champs compacts, adaptes au contenu
        self.ed_start.setMaximumWidth(120)
        self.ed_end.setMaximumWidth(120)
        self.cb_duplicate.setMaximumWidth(140)
        self.ed_org.setMaximumWidth(180)
        self.ed_country.setMaximumWidth(180)

        btn_add_type = QPushButton("Ajouter type")
        btn_add_type.setMaximumWidth(130)
        btn_add_type.clicked.connect(self.add_new_type)

        btn_load = QPushButton("Charger plage")
        btn_load.setMaximumWidth(150)
        btn_load.clicked.connect(self.load_range)

        labels = [
            ("Code debut", self.ed_start),
            ("Code fin", self.ed_end),
            ("Doublons", self.cb_duplicate),
            ("Organisme", self.ed_org),
            ("Pays", self.ed_country),
        ]

        row = 0
        for lbl, w in labels:
            g.addWidget(QLabel(lbl), row, 0, alignment=Qt.AlignLeft)
            g.addWidget(w, row, 1, alignment=Qt.AlignLeft)
            row += 1

        g.addWidget(btn_add_type, 2, 2, alignment=Qt.AlignLeft)
        g.addWidget(btn_load, 6, 1, alignment=Qt.AlignLeft)
        return box

    def _build_metrics_panel(self) -> QWidget:
        box = QGroupBox("Resume")
        h = QHBoxLayout(box)
        self.lbl_count = QLabel("Lignes: 0")
        self.lbl_pending = QLabel("A reimporter: 0")
        self.lbl_missing = QLabel("Codes manquants: 0")
        btn_missing = QPushButton("Voir codes manquants")
        btn_missing.clicked.connect(self.show_missing_codes)
        h.addWidget(self.lbl_count)
        h.addWidget(self.lbl_pending)
        h.addWidget(self.lbl_missing)
        h.addWidget(btn_missing)
        return box

    def _build_bulk_panel(self) -> QWidget:
        box = QGroupBox("Actions multi-lignes (sur cases Selection)")
        h = QHBoxLayout(box)

        self.bulk_include_mode = QComboBox()
        self.bulk_include_mode.addItems(["", "Inclure", "Exclure"])
        self.bulk_type = QComboBox()
        self.bulk_categorie = QComboBox(); self.bulk_categorie.addItems(CATEGORIE_OPTIONS)
        self.bulk_type_peche = QComboBox(); self.bulk_type_peche.addItems(TYPE_PECHE_OPTIONS)
        self.bulk_autre = QComboBox(); self.bulk_autre.addItems(OUI_NON_OPTIONS)
        self.bulk_ecailles = QComboBox(); self.bulk_ecailles.addItems(ECAILLES_BRUTES_OPTIONS)
        self.bulk_montees = QComboBox(); self.bulk_montees.addItems(NUMERIC_OPTIONS)
        self.bulk_otolithes = QComboBox(); self.bulk_otolithes.addItems(NUMERIC_OPTIONS)
        self.bulk_observation = QComboBox(); self.bulk_observation.addItems(OBSERVATION_OPTIONS)

        for lbl, w in [
            ("Type", self.bulk_type),
            ("Categorie", self.bulk_categorie),
            ("Type peche", self.bulk_type_peche),
            ("Autre oss", self.bulk_autre),
            ("Ecailles", self.bulk_ecailles),
            ("Montees", self.bulk_montees),
            ("Otolithes", self.bulk_otolithes),
            ("Observation", self.bulk_observation),
        ]:
            h.addWidget(QLabel(lbl)); h.addWidget(w)

        h.addWidget(QLabel("Inclure?")); h.addWidget(self.bulk_include_mode)

        btn_sel_all = QPushButton("Tout selectionner")
        btn_sel_none = QPushButton("Vider selection")
        btn_apply = QPushButton("Appliquer")
        btn_sel_all.clicked.connect(lambda: self._select_all(True))
        btn_sel_none.clicked.connect(lambda: self._select_all(False))
        btn_apply.clicked.connect(self.apply_bulk)

        h.addWidget(btn_sel_all)
        h.addWidget(btn_sel_none)
        h.addWidget(btn_apply)
        return box

    def _build_table(self) -> QWidget:
        self.table = QTableWidget(0, len(COLS))
        self.table.setHorizontalHeaderLabels(COLS)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(False)
        self.table.itemChanged.connect(self._on_table_item_changed)
        return self.table

    def _build_bottom_panel(self) -> QWidget:
        w = QWidget()
        h = QHBoxLayout(w)

        btn_csv = QPushButton("Enregistrer selection CSV")
        btn_csv.clicked.connect(self.save_csv)
        btn_import = QPushButton("Lancer import")
        btn_import.clicked.connect(self.run_import)
        btn_hist = QPushButton("Voir historique")
        btn_hist.clicked.connect(self.show_history)

        h.addWidget(btn_csv)
        h.addWidget(btn_import)
        h.addWidget(btn_hist)
        self.lbl_status = QLabel("Pret")
        h.addWidget(self.lbl_status)
        return w

    def _refresh_context_labels(self) -> None:
        self.lbl_source.setText(f"Source: {self.source_path.name} | Onglet: {self.source_sheet}")
        self.lbl_target.setText(f"Template interne: {self.target_path.name} | Onglet: {self.target_sheet}")
        self.lbl_paths.setText(f"Sortie: {self.out_path.name} | Historique: {self.history_path.name}")

    def _load_settings(self) -> None:
        if not SETTINGS_FILE.exists():
            return
        try:
            payload = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        except Exception:
            return

        def path_if_set(key: str, current: Path) -> Path:
            val = payload.get(key)
            if isinstance(val, str) and val.strip():
                return Path(val)
            return current

        self.source_path = path_if_set("source_path", self.source_path)
        self.out_path = path_if_set("out_path", self.out_path)
        self.history_path = path_if_set("history_path", self.history_path)
        self.selection_csv = path_if_set("selection_csv", self.selection_csv)

        src_sheet = payload.get("source_sheet")
        if isinstance(src_sheet, str) and src_sheet.strip():
            self.source_sheet = src_sheet.strip()

        start_ref = payload.get("start_ref")
        end_ref = payload.get("end_ref")
        org = payload.get("default_org")
        country = payload.get("default_country")
        dup_mode = payload.get("duplicate_mode")

        if isinstance(start_ref, str):
            self.ed_start.setText(start_ref)
        if isinstance(end_ref, str):
            self.ed_end.setText(end_ref)
        if isinstance(org, str):
            self.ed_org.setText(org)
        if isinstance(country, str):
            self.ed_country.setText(country)
        if isinstance(dup_mode, str) and self.cb_duplicate.findText(dup_mode) >= 0:
            self.cb_duplicate.setCurrentText(dup_mode)

        self.dark_mode = bool(payload.get("dark_mode", False))

    def _save_settings(self) -> None:
        try:
            SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
            payload = {
                "source_path": str(self.source_path),
                "source_sheet": self.source_sheet,
                "out_path": str(self.out_path),
                "history_path": str(self.history_path),
                "selection_csv": str(self.selection_csv),
                "start_ref": self.ed_start.text().strip(),
                "end_ref": self.ed_end.text().strip(),
                "default_org": self.ed_org.text().strip(),
                "default_country": self.ed_country.text().strip(),
                "duplicate_mode": self.cb_duplicate.currentText(),
                "dark_mode": bool(self.switch_theme.isChecked()),
            }
            SETTINGS_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            return

    def _on_theme_toggled(self, checked: bool) -> None:
        self._apply_theme(checked)
        self._save_settings()
        if hasattr(self, "lbl_status"):
            self.lbl_status.setText("Theme: nuit" if checked else "Theme: jour")

    def _update_app_icon(self) -> None:
        pix = self._build_logo_pixmap(self.dark_mode)
        self.setWindowIcon(QIcon(pix))

    def _build_logo_pixmap(self, dark_mode: bool) -> QPixmap:
        pix = QPixmap(256, 256)
        pix.fill(Qt.transparent)
        p = QPainter(pix)
        p.setRenderHint(QPainter.Antialiasing)

        bg = QColor("#0f172a") if dark_mode else QColor("#f8fafc")
        card = QColor("#111827") if dark_mode else QColor("#ffffff")
        border = QColor("#334155") if dark_mode else QColor("#94a3b8")
        accent = QColor("#22c55e") if dark_mode else QColor("#166534")
        glass = QColor("#374151") if dark_mode else QColor("#f8fafc")

        p.setPen(QPen(border, 1.2))
        p.setBrush(card)
        p.drawRoundedRect(12, 12, 232, 232, 28, 28)
        p.fillRect(14, 14, 228, 36, bg)

        p.setPen(QPen(border, 6))
        p.setBrush(QColor("#1f2937") if dark_mode else QColor("#e2e8f0"))
        p.drawRoundedRect(42, 58, 96, 140, 12, 12)

        p.setBrush(glass)
        for y in (72, 112, 152):
            p.drawRoundedRect(50, y, 80, 30, 8, 8)
            p.setBrush(accent)
            p.drawEllipse(114, y + 10, 8, 8)
            p.setBrush(glass)

        p.end()
        return pix

    def _apply_theme(self, dark_mode: bool) -> None:
        self.dark_mode = dark_mode
        if dark_mode:
            self.setStyleSheet(
                """
                QMainWindow, QWidget { background: #0f172a; color: #e5e7eb; }
                QGroupBox { border: 1px solid #334155; border-radius: 8px; margin-top: 8px; padding-top: 10px; }
                QLineEdit, QComboBox, QTableWidget { background: #111827; color: #e5e7eb; border: 1px solid #334155; }
                QPushButton { background: #1d4ed8; color: #ffffff; border-radius: 6px; padding: 5px 10px; }
                QPushButton:hover { background: #2563eb; }
                QHeaderView::section { background: #1e293b; color: #e5e7eb; }
                QCheckBox#themeSwitch::indicator { width: 38px; height: 18px; border-radius: 9px; background: #334155; border: 1px solid #475569; }
                QCheckBox#themeSwitch::indicator:checked { background: #22c55e; border: 1px solid #16a34a; }
                """
            )
        else:
            self.setStyleSheet(
                """
                QMainWindow, QWidget { background: #f8fafc; color: #0f172a; }
                QGroupBox { border: 1px solid #cbd5e1; border-radius: 8px; margin-top: 8px; padding-top: 10px; }
                QLineEdit, QComboBox, QTableWidget { background: #ffffff; color: #0f172a; border: 1px solid #cbd5e1; }
                QPushButton { background: #166534; color: #ffffff; border-radius: 6px; padding: 5px 10px; }
                QPushButton:hover { background: #15803d; }
                QHeaderView::section { background: #e2e8f0; color: #0f172a; }
                QCheckBox#themeSwitch::indicator { width: 38px; height: 18px; border-radius: 9px; background: #cbd5e1; border: 1px solid #94a3b8; }
                QCheckBox#themeSwitch::indicator:checked { background: #22c55e; border: 1px solid #16a34a; }
                """
            )
        self._update_app_icon()

    def show_about(self) -> None:
        date_str = dt.date.today().strftime("%d/%m/%Y")
        text = (
            "DATACOLISA - Interface PySide6\n"
            f"Version: {APP_VERSION}\n"
            f"Auteur: {APP_AUTHOR}\n"
            f"Date: {date_str}\n"
            f"Organisation: {APP_ORGANISATION}\n"
            "Application metier de gestion de collection et import de donnees."
        )
        QMessageBox.information(self, "A propos", text)

    def _import_base_path(self) -> Path:
        return self.out_path if self.out_path.exists() else self.target_path

    def _ensure_output_initialized(self) -> Path | None:
        if self.out_path.exists():
            return self.out_path
        if not self.target_path.exists():
            return None
        try:
            self.out_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(self.target_path, self.out_path)
            return self.out_path
        except Exception:
            return None

    def _refresh_type_options(self) -> None:
        target = self._import_base_path()
        opts = core.load_type_echantillon_options(target)
        if "EC MONTEE" not in opts:
            opts.append("EC MONTEE")
        self.type_options = sorted(set(opts))

        self.bulk_type.clear(); self.bulk_type.addItems(self.type_options)

    def select_source_file(self) -> None:
        p, _ = QFileDialog.getOpenFileName(self, "Choisir source .xls", str(self.source_path.parent), "Excel (*.xls *.xlsx)")
        if p:
            self.source_path = Path(p)
            self._refresh_context_labels()
            self._save_settings()

    def set_source_sheet(self) -> None:
        txt, ok = QInputDialog.getText(self, "Onglet source", "Nom onglet source:", text=self.source_sheet)
        if ok and txt.strip():
            self.source_sheet = txt.strip()
            self._refresh_context_labels()
            self._save_settings()

    def select_target_file(self) -> None:
        p, _ = QFileDialog.getOpenFileName(self, "Choisir cible .xlsx", str(self.target_path.parent), "Excel (*.xlsx)")
        if p:
            self.target_path = Path(p)
            self._refresh_type_options()
            self._refresh_context_labels()
            self._save_settings()

    def set_target_sheet(self) -> None:
        txt, ok = QInputDialog.getText(self, "Onglet cible", "Nom onglet cible:", text=self.target_sheet)
        if ok and txt.strip():
            self.target_sheet = txt.strip()
            self._refresh_context_labels()
            self._save_settings()

    def set_output_file(self) -> None:
        p, _ = QFileDialog.getSaveFileName(self, "Fichier sortie", str(self.out_path), "Excel (*.xlsx)")
        if p:
            self.out_path = Path(p)
            self._refresh_type_options()
            self._refresh_context_labels()
            self._save_settings()

    def set_history_file(self) -> None:
        p, _ = QFileDialog.getSaveFileName(self, "Fichier historique", str(self.history_path), "JSON (*.json)")
        if p:
            self.history_path = Path(p)
            self._refresh_context_labels()
            self._save_settings()

    def set_selection_csv(self) -> None:
        p, _ = QFileDialog.getSaveFileName(self, "Fichier selection CSV", str(self.selection_csv), "CSV (*.csv)")
        if p:
            self.selection_csv = Path(p)
            self._refresh_context_labels()
            self._save_settings()

    def closeEvent(self, event: Any) -> None:
        self._save_settings()
        super().closeEvent(event)

    def add_new_type(self) -> None:
        target = self._ensure_output_initialized()
        if target is None:
            QMessageBox.warning(self, "Type", f"Template introuvable: {self.target_path}")
            return
        value, ok = QInputDialog.getText(self, "Nouveau type", "Code type echantillon:")
        if not ok or not value.strip():
            return
        if core.append_type_echantillon_option(target, value.strip()):
            self._refresh_type_options()
            self.lbl_status.setText(f"Type ajoute: {value.strip()}")
        else:
            QMessageBox.warning(self, "Type", "Impossible d'ajouter le type")

    def load_range(self) -> None:
        try:
            _, xlrd = core.ensure_deps()
            source_rows, datemode = core.read_source_rows(xlrd, self.source_path, self.source_sheet)
            candidates = core.find_candidate_rows(source_rows, datemode)
            start_ref = self.ed_start.text().strip()
            end_ref = self.ed_end.text().strip()
            default_type = self.bulk_type.currentText().strip() or "EC MONTEE"

            filtered = [r for r in candidates if core.in_ref_range(r.ref, start_ref, end_ref)]
            filtered.sort(key=lambda r: core.parse_ref_parts(r.ref)[1] if core.parse_ref_parts(r.ref) else 0)

            found_codes = {core.normalize(r.ref).upper() for r in filtered}
            self.missing_codes = []
            p_start = core.parse_ref_parts(start_ref)
            p_end = core.parse_ref_parts(end_ref)
            if p_start and p_end and p_start[0] == p_end[0]:
                for n in range(p_start[1], p_end[1] + 1):
                    code = f"{p_start[0]}{n}"
                    if code not in found_codes:
                        self.missing_codes.append(code)

            self.rows = []
            for r in filtered:
                errs = core.validate_row(r, default_type)
                self.rows.append({
                    "selected": False,
                    "include": True,
                    "ref": core.normalize(r.ref),
                    "code_type_echantillon": default_type,
                    "categorie": core.normalize(r.categorie),
                    "type_peche": core.normalize(r.type_peche),
                    "autre_oss": "",
                    "ecailles_brutes": "",
                    "montees": "",
                    "otolithes": "",
                    "observation_disponibilite": core.normalize(r.observation_disponibilite),
                    "source_row": r.source_row_index,
                    "num_individu": core.normalize(r.num_individu),
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
                    "status": "a_reimporter" if errs else "pret",
                    "errors": " | ".join(errs),
                })

            self._render_table()
            self.lbl_count.setText(f"Lignes: {len(self.rows)}")
            pending = sum(1 for rr in self.rows if rr.get("status") == "a_reimporter")
            self.lbl_pending.setText(f"A reimporter: {pending}")
            self.lbl_missing.setText(f"Codes manquants: {len(self.missing_codes)}")
            self.lbl_status.setText(f"Lignes chargees: {len(self.rows)}")
        except Exception as exc:
            QMessageBox.critical(self, "Chargement", str(exc))

    def show_missing_codes(self) -> None:
        if not self.missing_codes:
            QMessageBox.information(self, "Codes manquants", "Aucun code manquant sur cette plage")
            return
        QMessageBox.information(self, "Codes manquants", "\n".join(self.missing_codes))

    def _render_table(self) -> None:
        self._table_sync_guard = True
        self.table.setRowCount(len(self.rows))
        for r, row in enumerate(self.rows):
            for c, key in enumerate(COLS):
                val = row.get(key, "")

                if key in ("selected", "include"):
                    item = qitem("", editable=True)
                    item.setCheckState(Qt.Checked if bool(val) else Qt.Unchecked)
                    self.table.setItem(r, c, item)
                    continue

                if key == "code_type_echantillon":
                    cb = QComboBox(); cb.addItems(self.type_options); cb.setCurrentText(str(val)); self.table.setCellWidget(r, c, cb); continue
                if key == "categorie":
                    cb = QComboBox(); cb.addItems(CATEGORIE_OPTIONS); cb.setCurrentText(str(val)); self.table.setCellWidget(r, c, cb); continue
                if key == "type_peche":
                    cb = QComboBox(); cb.addItems(TYPE_PECHE_OPTIONS)
                    if str(val) and cb.findText(str(val)) < 0:
                        cb.addItem(str(val))
                    cb.setCurrentText(str(val)); self.table.setCellWidget(r, c, cb); continue
                if key == "autre_oss":
                    cb = QComboBox(); cb.addItems(OUI_NON_OPTIONS); cb.setCurrentText(str(val)); self.table.setCellWidget(r, c, cb); continue
                if key == "ecailles_brutes":
                    cb = QComboBox(); cb.addItems(ECAILLES_BRUTES_OPTIONS)
                    if str(val) and cb.findText(str(val)) < 0:
                        cb.addItem(str(val))
                    cb.setCurrentText(str(val)); self.table.setCellWidget(r, c, cb); continue
                if key in ("montees", "otolithes"):
                    cb = QComboBox(); cb.addItems(NUMERIC_OPTIONS); cb.setCurrentText(str(val));
                    if key == "otolithes":
                        cb.currentTextChanged.connect(lambda _text, row_index=r: self._sync_autre_oss_from_otolithes(row_index))
                    self.table.setCellWidget(r, c, cb); continue
                if key == "observation_disponibilite":
                    cb = QComboBox(); cb.addItems(OBSERVATION_OPTIONS); cb.setCurrentText(str(val)); self.table.setCellWidget(r, c, cb); continue

                editable = key not in ("ref", "source_row", "status", "errors", "pecheur", "pays_capture")
                self.table.setItem(r, c, qitem(val, editable=editable))
            self._set_selected_item_enabled(r, bool(row.get("include")))
        self._table_sync_guard = False

    def _set_selected_item_enabled(self, row_index: int, enabled: bool) -> None:
        it_sel = self.table.item(row_index, COLS.index("selected"))
        if not it_sel:
            return
        flags = Qt.ItemIsEditable | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable
        if enabled:
            it_sel.setFlags(flags | Qt.ItemIsEnabled)
        else:
            it_sel.setFlags(flags)
            it_sel.setCheckState(Qt.Unchecked)

    def _on_table_item_changed(self, item: QTableWidgetItem) -> None:
        if self._table_sync_guard:
            return
        if item.column() != COLS.index("include"):
            return
        row_index = item.row()
        include_enabled = item.checkState() == Qt.Checked
        self._table_sync_guard = True
        self._set_selected_item_enabled(row_index, include_enabled)
        self._table_sync_guard = False

    def _read_table(self) -> List[Dict[str, Any]]:
        out: List[Dict[str, Any]] = []
        for r in range(self.table.rowCount()):
            row: Dict[str, Any] = {}
            for c, key in enumerate(COLS):
                w = self.table.cellWidget(r, c)
                it = self.table.item(r, c)
                if key in ("selected", "include"):
                    row[key] = bool(it and it.checkState() == Qt.Checked)
                elif isinstance(w, QComboBox):
                    row[key] = w.currentText()
                else:
                    row[key] = it.text() if it else ""

            ot = core.normalize(row.get("otolithes", ""))
            row["autre_oss"] = "OUI" if (ot and ot != "0") else "NON"

            out.append(row)
        self.rows = out
        return out

    def _selected_row_indexes(self) -> List[int]:
        return [i for i, r in enumerate(self._read_table()) if bool(r.get("selected")) and bool(r.get("include"))]

    def _select_all(self, value: bool) -> None:
        for r in range(self.table.rowCount()):
            it_inc = self.table.item(r, COLS.index("include"))
            it = self.table.item(r, COLS.index("selected"))
            if it:
                include_enabled = bool(it_inc and it_inc.checkState() == Qt.Checked)
                if include_enabled:
                    it.setCheckState(Qt.Checked if value else Qt.Unchecked)

    def _sync_autre_oss_from_otolithes(self, row_index: int) -> None:
        ot_w = self.table.cellWidget(row_index, COLS.index("otolithes"))
        autre_w = self.table.cellWidget(row_index, COLS.index("autre_oss"))
        if not isinstance(ot_w, QComboBox) or not isinstance(autre_w, QComboBox):
            return
        ot = core.normalize(ot_w.currentText())
        if ot and ot != "0":
            autre_w.setCurrentText("OUI")
        else:
            autre_w.setCurrentText("NON")

    def apply_bulk(self) -> None:
        idxs = self._selected_row_indexes()
        if not idxs:
            QMessageBox.information(self, "Selection", "Aucune ligne selectionnee")
            return

        for r in idxs:
            it_inc = self.table.item(r, COLS.index("include"))
            mode = self.bulk_include_mode.currentText()
            if it_inc and mode:
                it_inc.setCheckState(Qt.Checked if mode == "Inclure" else Qt.Unchecked)

            for key, val in [
                ("code_type_echantillon", self.bulk_type.currentText()),
                ("categorie", self.bulk_categorie.currentText()),
                ("type_peche", self.bulk_type_peche.currentText()),
                ("autre_oss", self.bulk_autre.currentText()),
                ("ecailles_brutes", self.bulk_ecailles.currentText()),
                ("montees", self.bulk_montees.currentText()),
                ("otolithes", self.bulk_otolithes.currentText()),
                ("observation_disponibilite", self.bulk_observation.currentText()),
            ]:
                if val == "":
                    continue
                w = self.table.cellWidget(r, COLS.index(key))
                if isinstance(w, QComboBox):
                    w.setCurrentText(val)

            self._sync_autre_oss_from_otolithes(r)

        self.lbl_status.setText(f"Bulk applique sur {len(idxs)} ligne(s)")

    def save_csv(self) -> None:
        try:
            data = self._read_table()
            out = self.selection_csv
            out.parent.mkdir(parents=True, exist_ok=True)

            headers = [
                "include", "status", "source_row", "ref", "num_individu", "date_capture", "code_espece",
                "lac_riviere", "pays_capture", "pecheur", "categorie", "type_peche", "observation_disponibilite", "autre_oss",
                "ecailles_brutes", "montees", "otolithes", "longueur_mm", "poids_g", "maturite", "sexe",
                "age_total", "code_type_echantillon", "errors",
            ]
            with out.open("w", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
                w.writeheader()
                for row in data:
                    payload = dict(row)
                    payload["include"] = "1" if bool(payload.get("include")) else "0"
                    w.writerow(payload)

            self.lbl_status.setText(f"CSV enregistre: {out}")
        except Exception as exc:
            QMessageBox.critical(self, "CSV", str(exc))

    def run_import(self) -> None:
        try:
            self.save_csv()
            data = self._read_table()
            include_count = sum(1 for r in data if bool(r.get("include")))
            if include_count == 0:
                QMessageBox.warning(self, "Import", "Aucune ligne n'est cochee en import (colonne include).")
                return

            target = self._ensure_output_initialized()
            if target is None:
                QMessageBox.warning(self, "Import", f"Template introuvable: {self.target_path}")
                return
            args = argparse.Namespace(
                selection_csv=str(self.selection_csv),
                target=str(target),
                target_sheet=self.target_sheet,
                out_target=str(self.out_path),
                history=str(self.history_path),
                default_organisme=self.ed_org.text(),
                default_country=self.ed_country.text(),
                on_duplicate=self.cb_duplicate.currentText(),
            )

            buf = StringIO()
            with redirect_stdout(buf):
                core.cmd_import(args)
            raw = buf.getvalue().strip()
            payload = json.loads(raw) if raw else {"message": "Import termine"}

            QMessageBox.information(self, "Import", self._format_import_popup(payload))
            self.lbl_status.setText("Import termine")
        except SystemExit as exc:
            QMessageBox.warning(self, "Import", f"Arrete: {exc}")
        except Exception as exc:
            QMessageBox.critical(self, "Import", str(exc))



    def _format_import_popup(self, payload: Dict[str, Any]) -> str:
        imported = int(payload.get("imported", 0) or 0)
        skipped_manual = int(payload.get("skipped_manual", 0) or 0)
        skipped_validation = int(payload.get("skipped_validation", 0) or 0)
        duplicates = int(payload.get("duplicates", 0) or 0)
        target_out = str(payload.get("target_out", ""))

        lines = [
            "Import termin?.",
            f"Lignes import?es: {imported}",
            f"Lignes exclues: {skipped_manual}",
            f"Lignes ? corriger: {skipped_validation}",
            f"Doublons d?tect?s: {duplicates}",
        ]
        if target_out:
            lines.append(f"Fichier de sortie: {target_out}")
        return "\n".join(lines)

    def show_history(self) -> None:
        try:
            if not self.history_path.exists():
                QMessageBox.information(self, "Historique", f"Fichier introuvable: {self.history_path}")
                return
            payload = json.loads(self.history_path.read_text(encoding="utf-8"))
            rows = payload.get("rows", [])
            QMessageBox.information(
                self,
                "Historique",
                json.dumps({"updated_at": payload.get("updated_at"), "rows": rows[:50], "count": len(rows)}, ensure_ascii=False, indent=2),
            )
        except Exception as exc:
            QMessageBox.warning(self, "Historique", str(exc))


def main() -> None:
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
