"""Generate the built-in COLISA workbook used by the application."""

from __future__ import annotations

import unicodedata
import datetime as dt
from pathlib import Path
from typing import Tuple

from config.constants import DEFAULT_TARGET_SHEET


TOTAL_COLUMNS = 41

# En-tetes codes en dur — structure identique aux fichiers COLISA de reference
HEADER_POSITIONS = {
    1:  "Code unite gestionnaire",
    2:  "Site Atelier",
    3:  "Numero du correspondant",
    4:  "Code type echantillon",
    5:  "Code echantillon",
    6:  "Code esp\u00e8ce",
    7:  "Sous-esp\u00e8ce ",
    8:  "Organisme pr\u00e9leveur",
    9:  "Nom de l'op\u00e9rateur",
    10: "Pays capture ",
    11: "Date capture",
    12: "Lac/riviere",
    13: "Lieu de capture / debarquement",
    14: "Cat\u00e9gorie p\u00eacheur ",
    15: "Type p\u00eache/engin ",
    16: "Maille (mm)",
    # 17 : colonne vide (r\u00e9serv\u00e9e)
    18: "CODE IDENTIFICATION",
    19: "Numero individu (numero de capture)",
    20: "Longueur totale (mm)",
    21: "Poids (g)",
    22: "Code stade",
    23: "Code maturit\u00e9 sexuelle",
    24: "Generer pour site colisa",
    25: "Code sexe",
    26: "Pr\u00e9sence de l'otolithe gauche (0 si non, 1 si oui)",
    27: "Pr\u00e9sence de l'otolithe droite (0 si non, 1 si oui)",
    28: "Nombre d' opercules en \u00e9tat",
    29: "Information stockage ",
    30: "Observation disponibilit\u00e9",
    31: "Autre \u00e9chantillon osseuses collect\u00e9e sur l'individu OUI/NON",
    32: "Age total",
    33: "Age rivi\u00e8re ",
    34: "Age lac",
    35: "Nombre de fraie",
    36: "Ecailles regen\u00e9r\u00e9es ? (0 si non, 1 si oui)",
    37: "Observations",
    38: "Ecailles brutes",
    39: "Mont\u00e9es",
    40: "Empreintes",
    41: "Otolithes",
}

# Donnees de la feuille "Type echantillon" (codes officiels COLISA)
TYPE_ECHANTILLON_DATA = [
    ("AN", "Abdomen de poisson", 291),
    ("BI", "Bile", 233),
    ("GN", "Bouche de poisson", 293),
    ("BN", "Branchie de poisson", 10),
    ("VN", "Colonne vert\u00e9brale de poisson", 306),
    ("HN", "Dos de Poisson", 294),
    ("EC", "Ecaille de poisson", 103),
    ("ES", "Estomac", None),
    ("FO", "Foie de poisson", 104),
    ("FN", "Fraction inconnue de poisson", 100),
    ("GO", "Gonade de poisson", 245),
    ("GR", "Graisse de poisson", 107),
    ("HM", "H\u00e9mocytes de poissons", 241),
    ("LE", "L\u00e8vre de poisson", 295),
    ("MN", "M\u00e2choire de poisson", 296),
    ("MU", "Muscle de poisson", 102),
    ("MP", "Muscle et peau de poisson", 272),
    ("MI", "Muscle, muscle+tissu adipeux face interne de peau", 164),
    ("QN", "Nageoire caudale de poisson", 299),
    ("NN", "Nageoire de poisson", 297),
    ("DN", "Nageoire dorsale de poisson", 292),
    ("PN", "Nageoire pectorale de poisson", 305),
    ("YN", "Oeil de poisson", 302),
    ("ON", "Opercules", 298),
    ("XN", "Orifice anal de poisson", 308),
    ("XU", "Orifice urog\u00e9nital de poisson", 301),
    ("OT", "Otolithes", None),
    ("KN", "P\u00e9doncule caudal de poisson", 304),
    ("CN", "Poisson entier", 101),
    ("WE", "Poisson etete et equeute", 155),
    ("WV", "Poisson sans visc\u00e8res ni gonades", 243),
    ("RE", "Rein de poisson", 105),
    ("NF", "Syst\u00e8me nerveux de poisson", 106),
    ("TN", "T\u00eate de poisson", 300),
    ("WN", "Tronc de poisson", 307),
]


def build_numero_identification_value(
    lac_riviere: object,
    code_type_echantillon: object,
    date_capture: object,
    numero_individu: object,
    type_peche: object = None,
    code_espece: object = None,
    code_echantillon: object = None,
) -> str:
    """Build the Numero d'identification value like Excel formula L/F/K/S/E."""
    lac_part = str(lac_riviere or "").strip()[:2]
    espece_source = code_espece if code_espece is not None else code_type_echantillon
    espece_part = str(espece_source or "").strip()[:1]
    num_part = str(numero_individu or "").strip()
    code_part = str(code_echantillon or "").strip()

    date_part = ""
    if isinstance(date_capture, dt.datetime):
        date_capture = date_capture.date()
    if isinstance(date_capture, dt.date):
        date_part = date_capture.strftime("%d%m%Y")

    if not (lac_part or espece_part or date_part or num_part or code_part):
        return ""

    num_segment = f"-{num_part}" if num_part else ""
    code_segment = f"-{code_part}" if code_part else ""
    return f"{lac_part}{espece_part}{date_part}{num_segment}{code_segment}".upper()


def build_numero_identification_formula(row_index: int) -> str:
    """Build the exact Excel formula for the CODE IDENTIFICATION column."""
    # Use English function names and commas for programmatic formula insertion.
    # Excel will localize function names for the user locale when the file is opened.
    return (
        f'=UPPER(CONCATENATE(LEFT(L{row_index},2),LEFT(F{row_index},1),'
        f'TEXT(K{row_index},"DDMMYYYY"),'
        f'IF(S{row_index}<>"",CONCATENATE("-",S{row_index}),""),"-",E{row_index}))'
    )


def build_code_echantillon_value(
    lac_riviere: object,
    code_type_echantillon: object,
    date_capture: object,
    age_total: object,
    numero_individu: object,
    type_peche: object = None,
    force_prefix: str = "T",
) -> str:
    """Build the Code echantillon : {PREFIXE}{NUMERO}.
    Le prefixe est force a 'T' par defaut (configurable via force_prefix).
    """
    num_part = str(numero_individu or "").strip()
    if not num_part:
        return ""
    prefix = force_prefix.strip().upper() if force_prefix and str(force_prefix).strip() else "T"
    return f"{prefix}{num_part}"


def create_internal_target_workbook(output_path: Path, openpyxl_module, template_path: Path | None = None) -> Path:
    """Create the built-in COLISA workbook used as import base."""
    if template_path and template_path.exists():
        return create_target_workbook_from_template(output_path, openpyxl_module, template_path)

    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, GradientFill

    workbook = openpyxl_module.Workbook()
    worksheet = workbook.active
    worksheet.title = DEFAULT_TARGET_SHEET

    # ── Colonnes format nombre et date ────────────────────────────────────────
    # Colonnes format nombre entier : Code unite (1), Num correspondant (3),
    # Ecailles brutes (38), Montees (39), Empreintes (40), Otolithes (41)
    _NUMBER_FORMAT_COLS = {1, 3, 38, 39, 40, 41}
    # Colonne date capture (11)
    _DATE_FORMAT_COLS = {11}

    # ── Style en-tete (identique aux fichiers COLISA de reference) ────────────
    header_font = Font(name="Times New Roman", bold=True, size=11)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    header_border = Border(
        bottom=Side(border_style="medium"),
    )

    # ── Ecriture des en-tetes et style ────────────────────────────────────────
    for col_index in range(1, TOTAL_COLUMNS + 1):
        header = HEADER_POSITIONS.get(col_index, "")
        cell = worksheet.cell(1, col_index)
        cell.value = header if header else None
        cell.font = header_font
        cell.alignment = header_align
        cell.border = header_border

    worksheet.row_dimensions[1].height = 73.5
    worksheet.freeze_panes = "A2"

    # ── Largeur colonnes + formats pre-appliques sur 2000 lignes ──────────────
    PRE_FORMAT_ROWS = 2000
    for col_index in range(1, TOTAL_COLUMNS + 1):
        header = HEADER_POSITIONS.get(col_index, "")
        width = max(14, min(len(header) + 4, 42)) if header else 8
        letter = _column_letter(col_index)
        worksheet.column_dimensions[letter].width = width

        if col_index in _NUMBER_FORMAT_COLS:
            fmt = "0"
        elif col_index in _DATE_FORMAT_COLS:
            fmt = "DD/MM/YYYY"
        else:
            continue

        for row_index in range(2, PRE_FORMAT_ROWS + 2):
            worksheet.cell(row_index, col_index).number_format = fmt

    # ── Feuil2 (vide) ──────────────────────────────────────────────────────────

    # ── Type echantillon ───────────────────────────────────────────────────────
    workbook.create_sheet("Feuil2")

    type_sheet = workbook.create_sheet("Type echantillon")
    type_sheet.cell(1, 1).value = "Code type echantillon"
    type_sheet.cell(1, 2).value = "Description"
    type_sheet.cell(1, 3).value = "code sandre"
    for row_index, (code, desc, sandre) in enumerate(TYPE_ECHANTILLON_DATA, start=2):
        type_sheet.cell(row_index, 1).value = code
        type_sheet.cell(row_index, 2).value = desc
        type_sheet.cell(row_index, 3).value = sandre
    type_sheet.column_dimensions["A"].width = 10
    type_sheet.column_dimensions["B"].width = 52
    type_sheet.column_dimensions["C"].width = 14

    # ── Especes ────────────────────────────────────────────────────────────────
    esp_sheet = workbook.create_sheet("Esp\u00e8ces")
    for col, hdr in enumerate(["Code esp\u00e8ce", "libell\u00e9", "code SANDRE", "code TAXREF"], start=1):
        esp_sheet.cell(1, col).value = hdr
    esp_sheet.column_dimensions["A"].width = 14
    esp_sheet.column_dimensions["B"].width = 30

    # ── Type peche ─────────────────────────────────────────────────────────────
    tp_sheet = workbook.create_sheet("Type p\u00eache ")
    tp_sheet.cell(1, 1).value = "Ligne "
    tp_sheet.column_dimensions["A"].width = 20

    # ── Stade ──────────────────────────────────────────────────────────────────
    stade_sheet = workbook.create_sheet("Stade")
    for col, hdr in enumerate(["CodeStade", "DescriptionCodeStade", "CodeSandre"], start=1):
        stade_sheet.cell(1, col).value = hdr
    stade_sheet.column_dimensions["A"].width = 12
    stade_sheet.column_dimensions["B"].width = 30

    # ── Sous espece ────────────────────────────────────────────────────────────
    sous_sheet = workbook.create_sheet("Sous esp\u00e8ce")
    sous_sheet.cell(1, 1).value = "Pal\u00e9e"
    sous_sheet.column_dimensions["A"].width = 20

    # ── Maturite sexuelle ──────────────────────────────────────────────────────
    mat_sheet = workbook.create_sheet("Maturit\u00e9 sexuelle")
    mat_sheet.cell(1, 1).value = "Code maturit\u00e9 sexuelle"
    mat_sheet.cell(1, 2).value = "Libell\u00e9"
    mat_sheet.column_dimensions["A"].width = 24
    mat_sheet.column_dimensions["B"].width = 30

    # ── Sexe ───────────────────────────────────────────────────────────────────
    sexe_sheet = workbook.create_sheet("Sexe")
    sexe_sheet.cell(1, 1).value = "Code sexe"
    sexe_sheet.cell(1, 2).value = "Libell\u00e9"
    sexe_sheet.column_dimensions["A"].width = 12
    sexe_sheet.column_dimensions["B"].width = 20

    # ── Type engin technique ───────────────────────────────────────────────────
    engin_sheet = workbook.create_sheet("Type engin technique ")
    engin_sheet.cell(1, 1).value = "Pics "
    engin_sheet.column_dimensions["A"].width = 20

    # ── type peche engins ──────────────────────────────────────────────────────
    pe_sheet = workbook.create_sheet("type p\u00eache engins")
    pe_sheet.cell(1, 1).value = "Type de p\u00eache /engins"
    pe_sheet.column_dimensions["A"].width = 24

    # ── Categorie de pecheurs ──────────────────────────────────────────────────
    cat_sheet = workbook.create_sheet("Categorie de pecheurs")
    cat_sheet.cell(1, 1).value = "Amateur "
    cat_sheet.column_dimensions["A"].width = 20

    # ── Sites Atelier ──────────────────────────────────────────────────────────
    site_sheet = workbook.create_sheet("Sites Atelier")
    site_sheet.cell(1, 1).value = "Nom du site atelier"
    site_sheet.column_dimensions["A"].width = 24

    # ── Observation disponibilite ──────────────────────────────────────────────
    obs_sheet = workbook.create_sheet("Observation disponibilit\u00e9")
    obs_sheet.cell(1, 1).value = "Code type echantillon"
    obs_sheet.cell(1, 2).value = "Description"
    obs_sheet.column_dimensions["A"].width = 24
    obs_sheet.column_dimensions["B"].width = 30

    # ── Sens migratoire ────────────────────────────────────────────────────────
    sens_sheet = workbook.create_sheet("Sens migratoire")
    sens_sheet.cell(1, 1).value = "CodeMigration"
    sens_sheet.cell(1, 2).value = "Description CodeMigration"
    sens_sheet.column_dimensions["A"].width = 16
    sens_sheet.column_dimensions["B"].width = 28

    # ── Code marque individuelle ───────────────────────────────────────────────
    marque_sheet = workbook.create_sheet("Code marque ind")
    marque_sheet.cell(1, 1).value = "Code marque individuelle"
    marque_sheet.cell(1, 2).value = "Libell\u00e9"
    marque_sheet.column_dimensions["A"].width = 26
    marque_sheet.column_dimensions["B"].width = 30

    # ── Correspondants ─────────────────────────────────────────────────────────
    corr_sheet = workbook.create_sheet("Correspondants")
    for col, hdr in enumerate(["Numero correspondant", "Nom", "Pr\u00e9nom", "Adresse", "T\u00e9l\u00e9phone", "Mail"], start=1):
        corr_sheet.cell(1, col).value = hdr
    corr_sheet.column_dimensions["A"].width = 22
    corr_sheet.column_dimensions["B"].width = 18
    corr_sheet.column_dimensions["C"].width = 18
    corr_sheet.column_dimensions["D"].width = 30
    corr_sheet.column_dimensions["E"].width = 16
    corr_sheet.column_dimensions["F"].width = 28

    _copy_reference_sheets_from_embedded_template(workbook, openpyxl_module)
    workbook.create_sheet("Feuil1 ")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    workbook.close()
    return output_path


def _copy_reference_sheets_from_embedded_template(workbook, openpyxl_module) -> None:
    """Fill reference sheets in the generated COLISA workbook from the embedded template."""
    try:
        from infrastructure.embedded_assets import get_colisa_logiciel_template_path

        template_path = get_colisa_logiciel_template_path()
        if not template_path.exists():
            return
        template_wb = openpyxl_module.load_workbook(template_path, read_only=False, data_only=True)
    except Exception:
        return

    sheet_name_map = {
        "Sites atelier": "Sites Atelier",
        "Types d'échantillon": "Type echantillon",
        "Types d'echantillon": "Type echantillon",
        "Espèces": "Espèces",
        "Especes": "Espèces",
        "Stades": "Stade",
        "Sens migratoires": "Sens migratoire",
        "Maturités sexuelles": "Maturité sexuelle",
        "Maturites sexuelles": "Maturité sexuelle",
        "Sexes": "Sexe",
        "Marquages individuels": "Code marque ind",
        "Correspondants": "Correspondants",
    }

    try:
        for src_name in template_wb.sheetnames:
            if normalize_sheet_name(src_name) == normalize_sheet_name("Echantillons"):
                continue

            dst_name = sheet_name_map.get(src_name, src_name)
            if dst_name in workbook.sheetnames:
                dst_ws = workbook[dst_name]
                _clear_sheet_values(dst_ws)
            else:
                dst_ws = workbook.create_sheet(dst_name)

            src_ws = template_wb[src_name]
            for row in src_ws.iter_rows():
                for src_cell in row:
                    dst_ws.cell(src_cell.row, src_cell.column).value = src_cell.value

            for col_letter, dimension in src_ws.column_dimensions.items():
                if dimension.width:
                    dst_ws.column_dimensions[col_letter].width = dimension.width
    finally:
        template_wb.close()


def _clear_sheet_values(worksheet) -> None:
    for row in worksheet.iter_rows():
        for cell in row:
            cell.value = None
            cell._comment = None
            cell.hyperlink = None


def _resolve_sheet_by_name(workbook, sheet_name: str):
    if sheet_name in workbook.sheetnames:
        return workbook[sheet_name]

    normalized_expected = normalize_sheet_name(sheet_name)
    for candidate in workbook.sheetnames:
        if normalize_sheet_name(candidate) == normalized_expected:
            return workbook[candidate]

    return workbook[workbook.sheetnames[0]]


def create_target_workbook_from_template(output_path: Path, openpyxl_module, template_path: Path) -> Path:
    """Clone the provided template workbook while keeping only the target sheet data rows."""
    workbook = openpyxl_module.load_workbook(template_path)
    try:
        target_sheet = _resolve_sheet_by_name(workbook, DEFAULT_TARGET_SHEET)
        _clear_worksheet_data_keep_header(target_sheet)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        workbook.save(output_path)
        return output_path
    finally:
        workbook.close()


def _clear_worksheet_data_keep_header(worksheet) -> None:
    """Remove row data while preserving the first header row and sheet layout."""
    max_row = worksheet.max_row
    max_col = worksheet.max_column
    if max_row <= 1:
        return

    for row_index in range(2, max_row + 1):
        for col_index in range(1, max_col + 1):
            cell = worksheet.cell(row_index, col_index)
            cell.value = None
            cell._comment = None
            cell.hyperlink = None


def normalize_sheet_name(value: object) -> str:
    return _normalize_header(value)


def validate_collect_science_source_workbook(workbook, sheet_name: str | None = None) -> Tuple[bool, str]:
    """Validate that an Excel file follows the COLISA structure needed by Collect-Science."""
    target_sheet = sheet_name or DEFAULT_TARGET_SHEET

    if workbook.sheetnames:
        worksheet = _resolve_sheet_by_name(workbook, target_sheet)
    else:
        worksheet = None

    if worksheet is None:
        return False, "Le fichier Excel ne contient aucune feuille exploitable."

    header_row = [worksheet.cell(1, col_index).value for col_index in range(1, min(worksheet.max_column + 1, 50))]
    normalized_headers = [_normalize_header(value) for value in header_row if _normalize_header(value)]

    # Recherche flexible : cherche si les mots-clés sont présents dans les en-têtes
    # Format : (mots-clés requis, mots-clés optionnels/alternatifs)
    required_keywords = {
        "Numero individu": (["individu"], ["numero"]),
        "Code espece": (["espece"], ["code"]),
        "Pays capture": (["pays"], ["capture"]),
        "Date capture": (["date"], ["capture"]),
        "Lac/riviere": (["lac", "riviere"], []),
        "Longueur totale": (["longueur"], ["totale", "mm"]),
        "Ecailles brutes": (["ecailles"], ["brutes"]),
        "Montees": (["montees"], []),
        "Empreintes": (["empreintes"], []),
        "Otolithes": (["otolithes"], ["otolithe"]),
        "Code echantillon": (["echantillon"], ["code"]),
    }

    def _matches_keywords(header: str, keywords_tuple: tuple) -> bool:
        """Check if required keywords are present, and at least one optional if provided."""
        required, optional = keywords_tuple
        # Tous les mots-clés requis doivent être présents
        has_required = all(keyword in header for keyword in required)
        if not has_required:
            return False
        # Si pas d'optionnels, c'est bon
        if not optional:
            return True
        # Sinon, au moins un optionnel doit être présent
        return any(keyword in header for keyword in optional)

    missing_headers = []
    for label, keywords_tuple in required_keywords.items():
        found = any(_matches_keywords(norm_header, keywords_tuple) for norm_header in normalized_headers)
        if not found:
            missing_headers.append(label)

    if missing_headers:
        missing_text = ", ".join(missing_headers[:4])
        if len(missing_headers) > 4:
            missing_text += ", ..."
        return (
            False,
            "Le fichier Excel choisi ne contient pas les colonnes COLISA attendues. "
            f"Colonnes manquantes: {missing_text}.",
        )

    return True, ""


def _column_letter(index: int) -> str:
    """Convert a 1-based column index to an Excel column letter."""
    letters = []
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def _normalize_header(value: object) -> str:
    """Normalize header to a simple searchable form."""
    if value is None:
        return ""
    import re
    text = str(value).strip().lower()
    # Normaliser les accents
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    # Enlever tout ce qui n'est pas alphanumérique ou espace
    text = re.sub(r'[^a-z0-9\s]', ' ', text)
    # Enlever les espaces multiples
    return " ".join(text.split())
