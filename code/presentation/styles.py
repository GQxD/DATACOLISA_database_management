"""Modern stylesheet definitions for the application."""

from infrastructure.embedded_assets import get_arrow_path as _get_arrow_path_embedded

# ─── Palette jour ────────────────────────────────────────────────────────────
COLORS = {
    "primary":          "#1a6b8a",
    "primary_hover":    "#135270",
    "primary_light":    "#d6edf5",

    "secondary":        "#3a7d6e",
    "secondary_hover":  "#2d6358",
    "secondary_light":  "#d4ede9",

    "accent":           "#7b4f9e",
    "accent_hover":     "#5f3a7a",

    "pipeline":         "#c0550a",
    "pipeline_hover":   "#963f06",

    "background":       "#f0f5f7",
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

# ─── Palette nuit (mode actif) ───────────────────────────────────────────────
COLORS_DARK = {
    # Bleu acier principal — boutons, titres, focus
    "primary":          "#3ab8e8",
    "primary_hover":    "#2298cc",
    "primary_light":    "#0c2535",
    "primary_grad_top": "#48c8f8",
    "primary_grad_bot": "#2298cc",

    # Vert émeraude — Collec-Science
    "secondary":        "#2ecba8",
    "secondary_hover":  "#1caa88",
    "secondary_light":  "#0a2520",
    "secondary_grad_top": "#38dbb8",
    "secondary_grad_bot": "#1caa88",

    # Violet doux — COLISA logiciel
    "accent":           "#b888f0",
    "accent_hover":     "#9060d0",
    "accent_grad_top":  "#c898f8",
    "accent_grad_bot":  "#9060d0",

    # Orange chaud — Export complet
    "pipeline":         "#f0922e",
    "pipeline_hover":   "#d07010",
    "pipeline_grad_top": "#faa840",
    "pipeline_grad_bot": "#d07010",

    # Surfaces
    "background":       "#0d1b2a",
    "surface":          "#162636",
    "surface_alt":      "#1c3044",
    "border":           "#253d52",
    "border_dark":      "#1a3040",

    # Texte
    "text":             "#ddeef8",
    "text_secondary":   "#7aaabb",

    # Etats
    "success":          "#2ecba8",
    "warning":          "#f0a030",
    "error":            "#e86060",
    "info":             "#3ab8e8",

    # Tableau
    "table_header":     "#182e42",
    "table_row_alt":    "#142030",
    "table_hover":      "#1e3a54",
    "table_selected":   "#1e4a74",

    "import_btn":       "#3ab8e8",
    "import_btn_hover": "#2298cc",
}


def _get_arrow_path(dark_mode: bool) -> str:
    return _get_arrow_path_embedded(dark_mode)


def get_stylesheet(dark_mode: bool = False) -> str:
    c = COLORS_DARK if dark_mode else COLORS
    arrow_path = _get_arrow_path(dark_mode)

    return f"""
    /* ══════════════════════════════════════════════════════════════
       GLOBAL
    ══════════════════════════════════════════════════════════════ */
    QMainWindow {{
        background-color: {c['background']};
    }}

    QWidget {{
        color: {c['text']};
        font-family: 'Segoe UI', Arial, sans-serif;
        font-size: 10pt;
    }}

    /* ══════════════════════════════════════════════════════════════
       BANDEAU APPLICATION (contextCard + appTitle)
    ══════════════════════════════════════════════════════════════ */
    QFrame#contextCard {{
        background-color: {c['surface']};
        border: 1px solid {c['border']};
        border-left: 4px solid {c['primary']};
        border-radius: 10px;
    }}

    QLabel#appTitle {{
        font-size: 15pt;
        font-weight: 800;
        color: {c['primary']};
        letter-spacing: 1px;
        padding: 4px 10px;
    }}

    QLabel#appVersion {{
        font-size: 8pt;
        color: {c['text_secondary']};
        padding: 2px 6px;
    }}

    /* ══════════════════════════════════════════════════════════════
       GROUP BOX
    ══════════════════════════════════════════════════════════════ */
    QGroupBox {{
        background-color: {c['surface']};
        border: 1px solid {c['border']};
        border-radius: 12px;
        margin-top: 14px;
        padding: 10px 10px 10px 10px;
        font-weight: 700;
        font-size: 10pt;
    }}

    QGroupBox::title {{
        subcontrol-origin: margin;
        subcontrol-position: top left;
        padding: 3px 14px;
        background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
            stop:0 {c['primary']}, stop:1 {c['primary_hover']});
        color: white;
        border-radius: 7px;
        left: 12px;
        font-size: 10pt;
        font-weight: 700;
        letter-spacing: 0.3px;
    }}

    /* ══════════════════════════════════════════════════════════════
       BOUTONS — BASE
    ══════════════════════════════════════════════════════════════ */
    QPushButton {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 {c['primary_grad_top']}, stop:1 {c['primary_grad_bot']});
        color: white;
        border: none;
        border-radius: 8px;
        padding: 6px 16px;
        font-weight: 700;
        font-size: 10pt;
    }}

    QPushButton:hover {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 {c['primary']}, stop:1 {c['primary_hover']});
    }}

    QPushButton:pressed {{
        background-color: {c['primary_hover']};
        padding-top: 8px;
        padding-bottom: 4px;
    }}

    QPushButton:disabled {{
        background-color: {c['border']};
        color: {c['text_secondary']};
    }}

    /* ══ Lancer import — bleu acier ══ */
    QPushButton#btn_import {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 {c['primary_grad_top']}, stop:1 {c['primary_grad_bot']});
        font-size: 11pt;
        font-weight: 800;
        padding: 9px 22px;
        border-radius: 10px;
        border-bottom: 3px solid {c['primary_hover']};
    }}
    QPushButton#btn_import:hover {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 {c['primary']}, stop:1 {c['primary_hover']});
    }}

    /* ══ Collec-Science — vert émeraude ══ */
    QPushButton#btn_collec {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 {c['secondary_grad_top']}, stop:1 {c['secondary_grad_bot']});
        font-size: 11pt;
        font-weight: 800;
        padding: 9px 22px;
        border-radius: 10px;
        border-bottom: 3px solid {c['secondary_hover']};
    }}
    QPushButton#btn_collec:hover {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 {c['secondary']}, stop:1 {c['secondary_hover']});
    }}

    /* ══ COLISA logiciel — violet ══ */
    QPushButton#btn_colisa_logiciel {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 {c['accent_grad_top']}, stop:1 {c['accent_grad_bot']});
        font-size: 11pt;
        font-weight: 800;
        padding: 9px 22px;
        border-radius: 10px;
        border-bottom: 3px solid {c['accent_hover']};
    }}
    QPushButton#btn_colisa_logiciel:hover {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 {c['accent']}, stop:1 {c['accent_hover']});
    }}

    /* ══ Export complet — orange ══ */
    QPushButton#btn_pipeline {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 {c['pipeline_grad_top']}, stop:1 {c['pipeline_grad_bot']});
        font-size: 11pt;
        font-weight: 800;
        padding: 9px 22px;
        border-radius: 10px;
        border-bottom: 3px solid {c['pipeline_hover']};
    }}
    QPushButton#btn_pipeline:hover {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 {c['pipeline']}, stop:1 {c['pipeline_hover']});
    }}

    /* ══════════════════════════════════════════════════════════════
       LINE EDIT
    ══════════════════════════════════════════════════════════════ */
    QLineEdit {{
        background-color: {c['surface_alt']};
        border: 1px solid {c['border']};
        border-radius: 7px;
        padding: 5px 8px;
        font-size: 10pt;
        color: {c['text']};
    }}

    QLineEdit:focus {{
        border: 2px solid {c['primary']};
        background-color: {c['surface']};
    }}

    QLineEdit:disabled {{
        background-color: {c['background']};
        color: {c['text_secondary']};
    }}

    QTableWidget QLineEdit {{
        padding: 2px 4px;
        border: 1px solid {c['border']};
    }}

    /* ══════════════════════════════════════════════════════════════
       COMBO BOX
    ══════════════════════════════════════════════════════════════ */
    QComboBox {{
        background-color: {c['surface_alt']};
        border: 1px solid {c['border']};
        border-radius: 7px;
        min-height: 32px;
        padding: 0px 30px 0px 8px;
        font-size: 10pt;
        color: {c['text']};
    }}

    QComboBox QLineEdit {{
        background-color: transparent;
        border: none;
        min-height: 28px;
        padding: 0px;
        color: {c['text']};
        selection-background-color: {c['primary']};
        selection-color: white;
    }}

    QComboBox:hover {{
        border: 1px solid {c['primary']};
    }}

    QComboBox:focus {{
        border: 2px solid {c['primary']};
    }}

    QComboBox::drop-down {{
        border-left: 1px solid {c['border']};
        width: 26px;
        background-color: {c['surface_alt']};
        border-top-right-radius: 7px;
        border-bottom-right-radius: 7px;
    }}

    QComboBox::down-arrow {{
        image: url("{arrow_path}");
        width: 12px;
        height: 8px;
    }}

    QComboBox QAbstractItemView {{
        background-color: {c['surface']};
        border: 1px solid {c['border']};
        border-radius: 7px;
        selection-background-color: {c['primary']};
        selection-color: white;
        color: {c['text']};
        padding: 4px;
        outline: none;
    }}

    QComboBox QAbstractItemView::item {{
        padding: 7px 10px;
        color: {c['text']};
        border-radius: 4px;
    }}

    QComboBox QAbstractItemView::item:hover {{
        background-color: {c['table_hover']};
    }}

    QTableWidget QComboBox {{
        padding: 2px 4px;
        border: 1px solid {c['border']};
    }}

    /* ══════════════════════════════════════════════════════════════
       CHECK BOX
    ══════════════════════════════════════════════════════════════ */
    QCheckBox {{
        spacing: 8px;
        font-size: 10pt;
    }}

    QCheckBox::indicator {{
        width: 17px;
        height: 17px;
        border: 2px solid {c['border']};
        border-radius: 5px;
        background-color: {c['surface_alt']};
    }}

    QCheckBox::indicator:hover {{
        border: 2px solid {c['primary']};
    }}

    QCheckBox::indicator:checked {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 {c['primary_grad_top']}, stop:1 {c['primary_grad_bot']});
        border: 2px solid {c['primary']};
    }}

    /* ══════════════════════════════════════════════════════════════
       LABELS
    ══════════════════════════════════════════════════════════════ */
    QLabel {{
        color: {c['text']};
        font-size: 10pt;
        padding: 1px;
    }}

    QLabel[class="header"] {{
        font-size: 13pt;
        font-weight: bold;
        color: {c['primary']};
        padding: 4px 0px;
    }}

    QLabel#introHelp, QLabel#panelHelp {{
        color: {c['text_secondary']};
        font-size: 9pt;
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

    /* Status bar au bas ══ */
    QLabel#statusLabel {{
        font-size: 10pt;
        font-weight: 700;
        color: {c['primary']};
        background-color: {c['surface_alt']};
        border: 1px solid {c['border']};
        border-radius: 8px;
        padding: 6px 14px;
    }}

    QLabel[class="secondary"] {{
        color: {c['text_secondary']};
        font-size: 9pt;
    }}

    /* ══════════════════════════════════════════════════════════════
       TABLE VIEW
    ══════════════════════════════════════════════════════════════ */
    QTableView {{
        background-color: {c['surface']};
        alternate-background-color: {c['table_row_alt']};
        border: 1px solid {c['border']};
        border-radius: 10px;
        gridline-color: {c['border_dark']};
        selection-background-color: {c['table_selected']};
        selection-color: {c['text']};
        font-size: 9.5pt;
    }}

    QTableView::item {{
        padding: 4px 7px;
        border: none;
    }}

    QTableView::item:hover {{
        background-color: {c['table_hover']};
    }}

    QTableView::item:selected {{
        background-color: {c['table_selected']};
    }}

    QHeaderView::section {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 {c['surface_alt']}, stop:1 {c['table_header']});
        color: {c['text']};
        padding: 7px 8px;
        border: none;
        border-right: 1px solid {c['border_dark']};
        border-bottom: 2px solid {c['primary']};
        font-weight: 700;
        font-size: 9pt;
        letter-spacing: 0.2px;
    }}

    QHeaderView::section:first {{
        border-top-left-radius: 8px;
    }}

    QHeaderView::section:last {{
        border-top-right-radius: 8px;
        border-right: none;
    }}

    QHeaderView::section:vertical {{
        background-color: {c['table_header']};
        color: {c['text_secondary']};
        padding: 3px 6px;
        border: none;
        border-bottom: 1px solid {c['border_dark']};
        font-weight: 600;
        font-size: 9pt;
    }}

    /* ══════════════════════════════════════════════════════════════
       MENU BAR
    ══════════════════════════════════════════════════════════════ */
    QMenuBar {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 {c['surface_alt']}, stop:1 {c['surface']});
        color: {c['text']};
        padding: 4px;
        border-bottom: 2px solid {c['primary']};
        font-weight: 600;
        font-size: 10pt;
    }}

    QMenuBar::item {{
        background-color: transparent;
        padding: 6px 16px;
        border-radius: 6px;
        color: {c['text']};
    }}

    QMenuBar::item:selected {{
        background-color: {c['primary']};
        color: white;
    }}

    QMenu {{
        background-color: {c['surface']};
        border: 1px solid {c['border']};
        border-radius: 8px;
        padding: 6px;
    }}

    QMenu::item {{
        padding: 8px 28px 8px 14px;
        border-radius: 5px;
        color: {c['text']};
        font-size: 10pt;
    }}

    QMenu::item:selected {{
        background-color: {c['primary']};
        color: white;
    }}

    QMenu::separator {{
        height: 1px;
        background-color: {c['border']};
        margin: 4px 8px;
    }}

    /* ══════════════════════════════════════════════════════════════
       SCROLL BARS
    ══════════════════════════════════════════════════════════════ */
    QScrollBar:vertical {{
        background-color: {c['background']};
        width: 9px;
        border-radius: 4px;
        margin: 0px;
    }}

    QScrollBar::handle:vertical {{
        background-color: {c['border']};
        min-height: 32px;
        border-radius: 4px;
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
        height: 9px;
        border-radius: 4px;
        margin: 0px;
    }}

    QScrollBar::handle:horizontal {{
        background-color: {c['border']};
        min-width: 32px;
        border-radius: 4px;
        margin: 2px;
    }}

    QScrollBar::handle:horizontal:hover {{
        background-color: {c['primary']};
    }}

    QScrollBar::add-line:horizontal,
    QScrollBar::sub-line:horizontal {{
        width: 0px;
    }}

    /* ══════════════════════════════════════════════════════════════
       TOOLTIP
    ══════════════════════════════════════════════════════════════ */
    QToolTip {{
        background-color: {c['surface_alt']};
        color: {c['text']};
        border: 1px solid {c['primary']};
        border-radius: 6px;
        padding: 6px 10px;
        font-size: 9.5pt;
    }}

    /* ══════════════════════════════════════════════════════════════
       MESSAGE BOX / DIALOG
    ══════════════════════════════════════════════════════════════ */
    QMessageBox {{
        background-color: {c['surface']};
    }}

    QMessageBox QPushButton {{
        min-width: 90px;
        min-height: 30px;
    }}

    QDialog {{
        background-color: {c['background']};
        color: {c['text']};
    }}

    /* ══════════════════════════════════════════════════════════════
       TEXT EDIT
    ══════════════════════════════════════════════════════════════ */
    QTextEdit {{
        background-color: {c['surface']};
        color: {c['text']};
        border: 1px solid {c['border']};
        border-radius: 8px;
        padding: 8px;
        selection-background-color: {c['primary']};
        selection-color: white;
    }}

    QTextEdit:focus {{
        border: 2px solid {c['primary']};
    }}

    /* ══════════════════════════════════════════════════════════════
       SPLITTER / FRAME SEPARATORS
    ══════════════════════════════════════════════════════════════ */
    QFrame[frameShape="4"],
    QFrame[frameShape="5"] {{
        color: {c['border']};
    }}

    QFrame#vSep {{
        background-color: {c['border']};
        min-width: 1px;
        max-width: 1px;
        margin: 4px 0;
    }}

    /* ══════════════════════════════════════════════════════════════
       PROGRESS BAR
    ══════════════════════════════════════════════════════════════ */
    QProgressBar {{
        background-color: {c['surface_alt']};
        border: 1px solid {c['border']};
        border-radius: 7px;
        text-align: center;
        color: {c['text']};
        font-weight: 700;
    }}

    QProgressBar::chunk {{
        background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
            stop:0 {c['primary']}, stop:1 {c['primary_hover']});
        border-radius: 6px;
    }}

    /* ══════════════════════════════════════════════════════════════
       TAB WIDGET
    ══════════════════════════════════════════════════════════════ */
    QTabWidget::pane {{
        background-color: {c['surface']};
        border: 1px solid {c['border']};
        border-radius: 8px;
        top: -1px;
    }}

    QTabBar::tab {{
        background-color: {c['background']};
        color: {c['text_secondary']};
        padding: 8px 20px;
        border: 1px solid {c['border']};
        border-bottom: none;
        border-top-left-radius: 7px;
        border-top-right-radius: 7px;
        margin-right: 3px;
        font-weight: 600;
    }}

    QTabBar::tab:selected {{
        background-color: {c['surface']};
        color: {c['primary']};
        border-bottom: 2px solid {c['primary']};
        font-weight: 800;
    }}

    QTabBar::tab:hover:!selected {{
        background-color: {c['primary_light']};
        color: {c['text']};
    }}

    /* ══════════════════════════════════════════════════════════════
       STATUS BAR
    ══════════════════════════════════════════════════════════════ */
    QStatusBar {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 {c['surface_alt']}, stop:1 {c['surface']});
        color: {c['text']};
        border-top: 2px solid {c['primary']};
        padding: 4px;
        font-weight: 600;
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
