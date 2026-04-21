"""Modern stylesheet definitions for the application."""

from infrastructure.embedded_assets import get_arrow_path as _get_arrow_path_embedded

# ─── Palette jour ───────────────────────────────────────────────────────────
COLORS = {
    "primary":          "#1a6b8a",   # bleu pétrole — en-têtes, boutons principaux
    "primary_hover":    "#135270",
    "primary_light":    "#d6edf5",

    "secondary":        "#3a7d6e",   # vert forêt — bouton Collec-Science
    "secondary_hover":  "#2d6358",
    "secondary_light":  "#d4ede9",

    "accent":           "#7b4f9e",   # violet — bouton COLISA logiciel
    "accent_hover":     "#5f3a7a",

    "pipeline":         "#c0550a",   # orange brûlé — Export complet
    "pipeline_hover":   "#963f06",

    "background":       "#f0f5f7",   # gris bleuté très clair
    "surface":          "#ffffff",
    "border":           "#b0c4cc",
    "border_dark":      "#4a6470",
    "text":             "#0f2830",
    "text_secondary":   "#4a6470",

    "success":          "#2e7d58",
    "warning":          "#a05c10",
    "error":            "#a83030",
    "info":             "#1a6b8a",

    "table_header":     "#bdd8e2",
    "table_row_alt":    "#eaf4f7",
    "table_hover":      "#d6edf5",
    "table_selected":   "#b8dde8",

    "import_btn":       "#1a6b8a",
    "import_btn_hover": "#135270",
}

# ─── Palette nuit ────────────────────────────────────────────────────────────
COLORS_DARK = {
    "primary":          "#4a9fc0",
    "primary_hover":    "#357fa0",
    "primary_light":    "#0d2d3d",

    "secondary":        "#3cb890",
    "secondary_hover":  "#2a9070",
    "secondary_light":  "#0a2820",

    "accent":           "#a07ac8",
    "accent_hover":     "#7d5aaa",

    "pipeline":         "#e0792a",
    "pipeline_hover":   "#c05a10",

    "background":       "#0c1a20",
    "surface":          "#162530",
    "border":           "#2a4050",
    "border_dark":      "#3a5565",
    "text":             "#e8f4f8",
    "text_secondary":   "#8ab0be",

    "success":          "#3dcb8a",
    "warning":          "#e09a30",
    "error":            "#e06060",
    "info":             "#4a9fc0",

    "table_header":     "#1e3a4a",
    "table_row_alt":    "#162530",
    "table_hover":      "#1e3a4a",
    "table_selected":   "#254555",

    "import_btn":       "#4a9fc0",
    "import_btn_hover": "#357fa0",
}


def _get_arrow_path(dark_mode: bool) -> str:
    return _get_arrow_path_embedded(dark_mode)


def get_stylesheet(dark_mode: bool = False) -> str:
    c = COLORS_DARK if dark_mode else COLORS
    arrow_path = _get_arrow_path(dark_mode)

    return f"""
    /* ===== GLOBAL ===== */
    QMainWindow {{
        background-color: {c['background']};
    }}

    QWidget {{
        color: {c['text']};
        font-family: 'Segoe UI', Arial, sans-serif;
        font-size: 9pt;
    }}

    /* ===== GROUP BOX ===== */
    QGroupBox {{
        background-color: {c['surface']};
        border: 1px solid {c['border']};
        border-radius: 10px;
        margin-top: 12px;
        padding: 12px 10px 10px 10px;
        font-weight: 600;
        font-size: 9pt;
    }}

    QGroupBox::title {{
        subcontrol-origin: margin;
        subcontrol-position: top left;
        padding: 3px 12px;
        background-color: {c['primary']};
        color: white;
        border-radius: 6px;
        left: 10px;
        font-size: 9pt;
        font-weight: 700;
    }}

    /* ===== BOUTONS PRINCIPAUX ===== */
    QPushButton {{
        background-color: {c['primary']};
        color: white;
        border: none;
        border-radius: 8px;
        padding: 6px 14px;
        font-weight: 600;
        font-size: 9pt;
    }}

    QPushButton:hover {{
        background-color: {c['primary_hover']};
    }}

    QPushButton:pressed {{
        background-color: {c['primary_hover']};
        padding: 8px 16px 6px 16px;
    }}

    QPushButton:disabled {{
        background-color: {c['border']};
        color: {c['text_secondary']};
    }}

    /* Bouton Import — bleu pétrole foncé */
    QPushButton#btn_import {{
        background-color: {c['import_btn']};
        font-size: 10pt;
        font-weight: 700;
        padding: 8px 18px;
        border-radius: 9px;
    }}
    QPushButton#btn_import:hover {{
        background-color: {c['import_btn_hover']};
    }}

    /* Bouton Collec-Science — vert */
    QPushButton#btn_collec {{
        background-color: {c['secondary']};
        font-size: 10pt;
        font-weight: 700;
        padding: 8px 18px;
        border-radius: 9px;
    }}
    QPushButton#btn_collec:hover {{
        background-color: {c['secondary_hover']};
    }}

    /* Bouton COLISA logiciel — violet */
    QPushButton#btn_colisa_logiciel {{
        background-color: {c['accent']};
        font-size: 10pt;
        font-weight: 700;
        padding: 8px 18px;
        border-radius: 9px;
    }}
    QPushButton#btn_colisa_logiciel:hover {{
        background-color: {c['accent_hover']};
    }}

    /* Bouton Export complet — orange */
    QPushButton#btn_pipeline {{
        background-color: {c['pipeline']};
        font-size: 10pt;
        font-weight: 700;
        padding: 8px 18px;
        border-radius: 9px;
    }}
    QPushButton#btn_pipeline:hover {{
        background-color: {c['pipeline_hover']};
    }}

    /* ===== LINE EDIT ===== */
    QLineEdit {{
        background-color: {c['surface']};
        border: 1px solid {c['border']};
        border-radius: 6px;
        padding: 4px 7px;
        font-size: 9pt;
    }}

    QTableWidget QLineEdit {{
        padding: 2px 4px;
        border: 1px solid {c['border']};
    }}

    QLineEdit:focus {{
        border: 2px solid {c['primary']};
    }}

    QLineEdit:disabled {{
        background-color: {c['background']};
        color: {c['text_secondary']};
    }}

    /* ===== COMBO BOX ===== */
    QComboBox {{
        background-color: {c['surface']};
        border: 1px solid {c['border']};
        border-radius: 6px;
        padding: 4px 7px;
        font-size: 9pt;
    }}

    QTableWidget QComboBox {{
        padding: 2px 4px;
        border: 1px solid {c['border']};
    }}

    QComboBox:hover {{
        border: 1px solid {c['primary']};
    }}

    QComboBox:focus {{
        border: 2px solid {c['primary']};
    }}

    QComboBox::drop-down {{
        border-left: 1px solid {c['border']};
        width: 28px;
        background-color: {c['surface']};
    }}

    QComboBox::down-arrow {{
        image: url("{arrow_path}");
        width: 12px;
        height: 8px;
    }}

    QComboBox QAbstractItemView {{
        background-color: {c['surface']};
        border: 1px solid {c['border']};
        border-radius: 6px;
        selection-background-color: {c['primary']};
        selection-color: white;
        color: {c['text']};
        padding: 4px;
    }}

    QComboBox QAbstractItemView::item {{
        padding: 6px;
        color: {c['text']};
    }}

    /* ===== CHECK BOX ===== */
    QCheckBox {{
        spacing: 8px;
        font-size: 9pt;
    }}

    QCheckBox::indicator {{
        width: 16px;
        height: 16px;
        border: 2px solid {c['border']};
        border-radius: 4px;
        background-color: {c['surface']};
    }}

    QCheckBox::indicator:hover {{
        border: 2px solid {c['primary']};
    }}

    QCheckBox::indicator:checked {{
        background-color: {c['primary']};
        border: 2px solid {c['primary']};
    }}

    /* ===== LABEL ===== */
    QLabel {{
        color: {c['text']};
        font-size: 9pt;
        padding: 2px;
    }}

    QLabel[class="header"] {{
        font-size: 12pt;
        font-weight: bold;
        color: {c['primary']};
        padding: 4px 0px;
    }}

    QLabel#introHelp, QLabel#panelHelp {{
        color: {c['text_secondary']};
        font-size: 9pt;
        padding: 0px;
    }}

    QLabel#workflowHelp {{
        color: {c['text']};
        background-color: {c['surface']};
        border: 2px solid {c['primary']};
        border-radius: 8px;
        padding: 7px 10px;
    }}

    QLabel#workflowNote {{
        color: {c['text_secondary']};
        background-color: {c['table_row_alt']};
        border: 1px solid {c['border']};
        border-radius: 8px;
        padding: 8px 10px;
    }}

    QFrame#contextCard {{
        background-color: {c['surface']};
        border: 2px solid {c['primary']};
        border-radius: 12px;
    }}

    QLabel[class="secondary"] {{
        color: {c['text_secondary']};
        font-size: 9pt;
    }}

    /* ===== TABLE VIEW ===== */
    QTableView {{
        background-color: {c['surface']};
        alternate-background-color: {c['table_row_alt']};
        border: 1px solid {c['border']};
        border-radius: 8px;
        gridline-color: {c['border']};
        selection-background-color: {c['table_selected']};
        selection-color: {c['text']};
    }}

    QTableView::item {{
        padding: 3px 6px;
        border: none;
    }}

    QTableView::item:hover {{
        background-color: {c['table_hover']};
    }}

    QTableView::item:selected {{
        background-color: {c['table_selected']};
    }}

    QHeaderView::section {{
        background-color: {c['table_header']};
        color: {c['text']};
        padding: 6px;
        border: none;
        border-right: 1px solid {c['border']};
        border-bottom: 2px solid {c['primary']};
        font-weight: 700;
        font-size: 9pt;
    }}

    QHeaderView::section:first {{
        border-top-left-radius: 6px;
    }}

    QHeaderView::section:last {{
        border-top-right-radius: 6px;
        border-right: none;
    }}

    /* ===== MENU BAR ===== */
    QMenuBar {{
        background-color: {c['primary']};
        color: white;
        padding: 3px;
        font-weight: 600;
    }}

    QMenuBar::item {{
        background-color: transparent;
        padding: 6px 14px;
        border-radius: 4px;
        color: white;
    }}

    QMenuBar::item:selected {{
        background-color: {c['primary_hover']};
        color: white;
    }}

    QMenu {{
        background-color: {c['surface']};
        border: 1px solid {c['border']};
        border-radius: 6px;
        padding: 4px;
    }}

    QMenu::item {{
        padding: 8px 24px 8px 12px;
        border-radius: 4px;
        color: {c['text']};
    }}

    QMenu::item:selected {{
        background-color: {c['primary']};
        color: white;
    }}

    /* ===== STATUS BAR ===== */
    QStatusBar {{
        background-color: {c['primary']};
        color: white;
        border-top: none;
        padding: 4px;
        font-weight: 600;
    }}

    /* ===== SCROLL BAR ===== */
    QScrollBar:vertical {{
        background-color: {c['background']};
        width: 10px;
        border-radius: 5px;
        margin: 0px;
    }}

    QScrollBar::handle:vertical {{
        background-color: {c['border']};
        min-height: 30px;
        border-radius: 5px;
        margin: 2px;
    }}

    QScrollBar::handle:vertical:hover {{
        background-color: {c['primary']};
    }}

    QScrollBar::add-line:vertical,
    QScrollBar::sub-line:vertical {{
        height: 0px;
    }}

    QScrollBar:horizontal {{
        background-color: {c['background']};
        height: 10px;
        border-radius: 5px;
        margin: 0px;
    }}

    QScrollBar::handle:horizontal {{
        background-color: {c['border']};
        min-width: 30px;
        border-radius: 5px;
        margin: 2px;
    }}

    QScrollBar::handle:horizontal:hover {{
        background-color: {c['primary']};
    }}

    QScrollBar::add-line:horizontal,
    QScrollBar::sub-line:horizontal {{
        width: 0px;
    }}

    /* ===== MESSAGE BOX ===== */
    QMessageBox {{
        background-color: {c['surface']};
    }}

    QMessageBox QPushButton {{
        min-width: 80px;
    }}

    /* ===== DIALOG ===== */
    QDialog {{
        background-color: {c['background']};
        color: {c['text']};
    }}

    /* ===== TEXT EDIT ===== */
    QTextEdit {{
        background-color: {c['surface']};
        color: {c['text']};
        border: 1px solid {c['border']};
        border-radius: 6px;
        padding: 8px;
        selection-background-color: {c['primary']};
        selection-color: white;
    }}

    QTextEdit:focus {{
        border: 2px solid {c['primary']};
    }}

    /* ===== TAB WIDGET ===== */
    QTabWidget::pane {{
        background-color: {c['surface']};
        border: 1px solid {c['border']};
        border-radius: 6px;
        top: -1px;
    }}

    QTabBar::tab {{
        background-color: {c['background']};
        color: {c['text']};
        padding: 8px 18px;
        border: 1px solid {c['border']};
        border-bottom: none;
        border-top-left-radius: 6px;
        border-top-right-radius: 6px;
        margin-right: 2px;
        font-weight: 600;
    }}

    QTabBar::tab:selected {{
        background-color: {c['surface']};
        color: {c['primary']};
        border-bottom: 2px solid {c['primary']};
        font-weight: 700;
    }}

    QTabBar::tab:hover:!selected {{
        background-color: {c['primary_light']};
    }}

    /* ===== PROGRESS BAR ===== */
    QProgressBar {{
        background-color: {c['background']};
        border: 1px solid {c['border']};
        border-radius: 6px;
        text-align: center;
        color: {c['text']};
        font-weight: 600;
    }}

    QProgressBar::chunk {{
        background-color: {c['primary']};
        border-radius: 5px;
    }}
    """


def get_button_icons() -> dict:
    return {
        "load": "📂",
        "import": "📥",
        "export": "📤",
        "save": "💾",
        "search": "🔍",
        "settings": "⚙️",
        "history": "📋",
        "add": "➕",
        "remove": "➖",
        "apply": "✓",
        "cancel": "✕",
        "refresh": "🔄",
        "help": "❓",
        "info": "ℹ️",
        "warning": "⚠️",
        "error": "❌",
    }
