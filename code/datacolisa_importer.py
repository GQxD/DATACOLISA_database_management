#!/usr/bin/env python3
"""
Importateur DATACOLISA — CLI principal.

Sous-commandes disponibles :
  extract   Filtre la source par plage REF et génère un CSV de sélection.
  import    Importe vers la cible (Feuil1) selon le CSV de sélection.
  reimport  Liste les lignes marquées non importées ou à réimporter.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import sys
import unicodedata
from pathlib import Path
from typing import Any, List, Optional, Tuple

# Couche domaine
from domain.models import SourceRow, ImportConfig, ImportResult, ValidationResult
from domain.value_objects import RefCode, DateCapture
from domain.business_rules import TypePecheDeriver, PaysDeriver, ValidationRules

# Configuration
from config.mappings import SOURCE_POSITIONS, TYPE_SHEET_CANDIDATES
from config.constants import DEFAULT_SOURCE_SHEET, DEFAULT_TARGET_SHEET


# ---------------------------------------------------------------------------
# Dépendances optionnelles
# ---------------------------------------------------------------------------

REQUIRED_IMPORTS = {
    "openpyxl": "openpyxl",
    "xlrd":     "xlrd",
}


# ---------------------------------------------------------------------------
# Fonctions utilitaires
# ---------------------------------------------------------------------------

def fatal(msg: str) -> None:
    """Affiche une erreur sur stderr et quitte le processus."""
    print(f"ERREUR: {msg}", file=sys.stderr)
    sys.exit(1)


def ensure_deps() -> Tuple[Any, Any]:
    """Vérifie et charge openpyxl et xlrd ; quitte si l'un est absent."""
    missing = []
    loaded: dict = {}
    for mod_name in REQUIRED_IMPORTS:
        try:
            loaded[mod_name] = __import__(mod_name)
        except Exception:
            missing.append(mod_name)
    if missing:
        fatal(
            "Modules manquants : "
            + ", ".join(missing)
            + ". Installez-les avec : pip install -r requirements.txt"
        )
    return loaded["openpyxl"], loaded["xlrd"]


def normalize(s: Any) -> str:
    """Normalisation générale d'une valeur cellule en chaîne propre."""
    if s is None:
        return ""
    if isinstance(s, float):
        if s.is_integer():
            return str(int(s))
        return format(s, "f").rstrip("0").rstrip(".")
    return str(s).strip()


def normalize_header_name(s: Any) -> str:
    """Normalise un en-tête : minuscules, sans accents, sans ponctuation."""
    txt = normalize(s)
    if not txt:
        return ""
    txt = unicodedata.normalize("NFKD", txt)
    txt = "".join(ch for ch in txt if not unicodedata.combining(ch))
    txt = txt.lower()
    txt = re.sub(r"[^a-z0-9]+", " ", txt)
    return re.sub(r"\s+", " ", txt).strip()


# ---------------------------------------------------------------------------
# Codes de référence (REF)
# ---------------------------------------------------------------------------

def parse_ref_parts(code: str) -> Optional[Tuple[str, int]]:
    """Wrapper hérité pour RefCode.parse() — préférer RefCode dans le nouveau code."""
    ref = RefCode.parse(code)
    if ref:
        return (ref.prefix, ref.number)
    return None


def normalize_ref_code(code: Any) -> str:
    """Normalise les codes de référence, p. ex. 'XY 0682' → 'XY682'."""
    txt = normalize(code)
    if not txt:
        return ""
    ref = RefCode.parse(txt)
    if ref:
        return str(ref)
    return txt.upper().replace(" ", "")


def in_ref_range(code: str, start: str, end: str) -> bool:
    """Wrapper hérité pour RefCode.in_range() — préférer RefCode dans le nouveau code."""
    ref       = RefCode.parse(code)
    start_ref = RefCode.parse(start)
    end_ref   = RefCode.parse(end)
    if not ref or not start_ref or not end_ref:
        return False
    try:
        return ref.in_range(start_ref, end_ref)
    except ValueError:
        return False


# ---------------------------------------------------------------------------
# Lecture des fichiers source
# ---------------------------------------------------------------------------

def read_source_rows(
    xlrd: Any,
    source_file: Path,
    sheet_name: str,
) -> Tuple[List[List[Any]], int]:
    """Lit toutes les lignes d'un fichier .xls via xlrd."""
    wb = xlrd.open_workbook(str(source_file), formatting_info=False)
    if sheet_name not in wb.sheet_names():
        raise ValueError(f"Onglet source introuvable : {sheet_name}")
    ws   = wb.sheet_by_name(sheet_name)
    rows = [
        [ws.cell_value(r, c) if c < ws.ncols else None for c in range(ws.ncols)]
        for r in range(ws.nrows)
    ]
    return rows, int(getattr(wb, "datemode", 0))


def read_any_source_rows(
    source_file: Path,
    sheet_name: str,
) -> Tuple[List[List[Any]], int]:
    """Lit les lignes source depuis un fichier .xls ou .xlsx."""
    suffix = source_file.suffix.lower()
    if suffix == ".xls":
        _, xlrd = ensure_deps()
        return read_source_rows(xlrd, source_file, sheet_name)

    openpyxl, _ = ensure_deps()
    wb = openpyxl.load_workbook(source_file, read_only=True, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Onglet source introuvable : {sheet_name}")
        ws   = wb[sheet_name]
        rows = [list(row) for row in ws.iter_rows(values_only=True)]
        return rows, 0
    finally:
        wb.close()


def get_workbook_sheet_names(source_file: Path) -> List[str]:
    """Retourne la liste des onglets d'un fichier .xls ou .xlsx."""
    suffix = source_file.suffix.lower()
    if suffix == ".xls":
        _, xlrd = ensure_deps()
        wb = xlrd.open_workbook(str(source_file), formatting_info=False)
        return list(wb.sheet_names())

    openpyxl, _ = ensure_deps()
    wb = openpyxl.load_workbook(source_file, read_only=True, data_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# Accès aux cellules
# ---------------------------------------------------------------------------

def get_pos(row: List[Any], one_based_col: int) -> Any:
    """Retourne la valeur d'une colonne (numérotation 1-based)."""
    idx = one_based_col - 1
    if idx < 0 or idx >= len(row):
        return None
    return row[idx]


# ---------------------------------------------------------------------------
# Conversion des dates Excel
# ---------------------------------------------------------------------------

def _excel_date_to_date(value: Any, datemode: int = 0) -> Optional[dt.date]:
    """Convertit une valeur Excel (numérique ou texte) en objet date Python."""
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, (int, float)):
        try:
            import xlrd  # type: ignore
            tup = xlrd.xldate_as_tuple(float(value), datemode)
            return dt.date(tup[0], tup[1], tup[2])
        except Exception:
            pass
        try:
            return dt.date(1899, 12, 30) + dt.timedelta(days=int(float(value)))
        except Exception:
            return None
    if isinstance(value, str):
        s = value.strip()
        if not s:
            return None
        for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d"):
            try:
                return dt.datetime.strptime(s, fmt).date()
            except ValueError:
                continue
    return None


def format_capture_date(value: Any, datemode: int = 0) -> str:
    """Formate une date de capture au format DD/MM/YY (chaîne)."""
    d = _excel_date_to_date(value, datemode)
    if d:
        return d.strftime("%d/%m/%y")
    return normalize(value)


# ---------------------------------------------------------------------------
# Dérivations métier (wrappers hérités)
# ---------------------------------------------------------------------------

def derive_type_and_categorie_from_source(raw_engin: Any) -> Tuple[str, str]:
    """Wrapper hérité — utiliser TypePecheDeriver directement dans le nouveau code."""
    return TypePecheDeriver.derive_type_and_categorie(str(raw_engin or ""))


def derive_country_from_contexte(raw_contexte: Any) -> str:
    """Wrapper hérité — utiliser PaysDeriver directement dans le nouveau code."""
    return PaysDeriver.derive_country(str(raw_contexte or ""))


# ---------------------------------------------------------------------------
# Construction des lignes source
# ---------------------------------------------------------------------------

def find_candidate_rows(
    rows: List[List[Any]],
    datemode: int = 0,
) -> List[SourceRow]:
    """Construit les objets SourceRow depuis les positions fixes SOURCE_POSITIONS."""
    out: List[SourceRow] = []
    for i, row in enumerate(rows, start=1):
        ref = normalize_ref_code(get_pos(row, SOURCE_POSITIONS["num_individu_primary"]))
        if not ref:
            continue
        if not re.search(r"\d", ref):
            continue

        num_individu = ref or normalize_ref_code(
            get_pos(row, SOURCE_POSITIONS["num_individu_fallback"])
        )
        engin_raw    = get_pos(row, SOURCE_POSITIONS["engin_source"])
        contexte_raw = get_pos(row, SOURCE_POSITIONS["contexte"])
        type_peche, categorie = derive_type_and_categorie_from_source(engin_raw)
        pays_capture = derive_country_from_contexte(contexte_raw)

        poids_src = normalize(get_pos(row, SOURCE_POSITIONS["poids_g"]))
        sexe_src  = normalize(get_pos(row, SOURCE_POSITIONS["sexe"]))
        if poids_src == "_":
            poids_src = ""
        if sexe_src == "_":
            sexe_src = ""

        out.append(
            SourceRow(
                source_row_index=i,
                ref=ref,
                code_espece=get_pos(row, SOURCE_POSITIONS["code_espece"]),
                date_capture=format_capture_date(
                    get_pos(row, SOURCE_POSITIONS["date_capture"]), datemode
                ),
                lac_riviere=get_pos(row, SOURCE_POSITIONS["lac_riviere"]),
                num_individu=num_individu,
                longueur_mm=get_pos(row, SOURCE_POSITIONS["longueur_mm"]),
                poids_g=poids_src if poids_src else "N",
                maturite=get_pos(row, SOURCE_POSITIONS["maturite"]),
                sexe=sexe_src if sexe_src else "N",
                age_total=get_pos(row, SOURCE_POSITIONS["age_total"]),
                type_peche=type_peche,
                categorie=categorie,
                pecheur=normalize(get_pos(row, SOURCE_POSITIONS["pecheur"])),
                pays_capture=pays_capture,
                pecheur_source=normalize(engin_raw),
                observation_disponibilite="",
                ecailles_brutes="",
                montees="",
                empreintes="",
                otolithes="",
                sous_espece="",
                nom_operateur="",
                lieu_capture="",
                maille_mm="",
                code_stade="",
                presence_otolithe_gauche="",
                presence_otolithe_droite="",
                nb_opercules="",
                information_stockage="",
                age_riviere="",
                age_lac="",
                nb_fraie="",
                ecailles_regenerees="",
                observations="",
            )
        )
    return out


def find_candidate_rows_from_mapping(
    rows: List[List[Any]],
    datemode: int,
    mapping: dict[str, Any],
) -> List[SourceRow]:
    """Construit les objets SourceRow depuis un mapping de colonnes fourni par l'utilisateur."""
    header_row = max(1, int(mapping.get("header_row", 1) or 1))
    columns    = mapping.get("columns", {}) or {}
    start_idx  = header_row

    def get_mapped(row: List[Any], key: str) -> Any:
        idx = columns.get(key)
        if idx is None:
            return None
        try:
            idx = int(idx)
        except Exception:
            return None
        if idx < 0 or idx >= len(row):
            return None
        return row[idx]

    out: List[SourceRow] = []
    for row_idx, row in enumerate(rows[start_idx:], start=start_idx + 1):
        ref = normalize_ref_code(get_mapped(row, "num_individu"))
        if not ref or not re.search(r"\d", ref):
            continue

        engin_raw    = get_mapped(row, "engin_source")
        contexte_raw = get_mapped(row, "contexte")

        derived_type, derived_cat = derive_type_and_categorie_from_source(engin_raw)
        type_peche   = normalize(get_mapped(row, "type_peche"))  or derived_type
        categorie    = normalize(get_mapped(row, "categorie"))   or derived_cat
        pays_capture = (
            normalize(get_mapped(row, "pays_capture"))
            or derive_country_from_contexte(contexte_raw)
        )

        poids_src = normalize(get_mapped(row, "poids_g"))
        sexe_src  = normalize(get_mapped(row, "sexe"))
        if poids_src == "_":
            poids_src = ""
        if sexe_src == "_":
            sexe_src = ""

        out.append(
            SourceRow(
                source_row_index=row_idx,
                ref=ref,
                code_espece=get_mapped(row, "code_espece"),
                date_capture=format_capture_date(
                    get_mapped(row, "date_capture"), datemode
                ),
                lac_riviere=get_mapped(row, "lac_riviere"),
                num_individu=ref,
                longueur_mm=get_mapped(row, "longueur_mm"),
                poids_g=poids_src if poids_src else "N",
                maturite=get_mapped(row, "maturite"),
                sexe=sexe_src if sexe_src else "N",
                age_total=get_mapped(row, "age_total"),
                type_peche=type_peche,
                categorie=categorie,
                pecheur=normalize(get_mapped(row, "pecheur")),
                pays_capture=pays_capture,
                pecheur_source=normalize(engin_raw),
                observation_disponibilite=normalize(
                    get_mapped(row, "observation_disponibilite")
                ),
                ecailles_brutes=normalize(get_mapped(row, "ecailles_brutes")),
                montees=normalize(get_mapped(row, "montees")),
                empreintes=normalize(get_mapped(row, "empreintes")),
                otolithes=normalize(get_mapped(row, "otolithes")),
                sous_espece=normalize(get_mapped(row, "sous_espece")),
                nom_operateur=normalize(get_mapped(row, "nom_operateur")),
                lieu_capture=normalize(get_mapped(row, "lieu_capture")),
                maille_mm=normalize(get_mapped(row, "maille_mm")),
                code_stade=normalize(get_mapped(row, "code_stade")),
                presence_otolithe_gauche=normalize(get_mapped(row, "presence_otolithe_gauche")),
                presence_otolithe_droite=normalize(get_mapped(row, "presence_otolithe_droite")),
                nb_opercules=normalize(get_mapped(row, "nb_opercules")),
                information_stockage=normalize(get_mapped(row, "information_stockage")),
                age_riviere=normalize(get_mapped(row, "age_riviere")),
                age_lac=normalize(get_mapped(row, "age_lac")),
                nb_fraie=normalize(get_mapped(row, "nb_fraie")),
                ecailles_regenerees=normalize(get_mapped(row, "ecailles_regenerees")),
                observations=normalize(get_mapped(row, "observations")),
            )
        )
    return out


def validate_row(r: SourceRow, default_type_echantillon: str) -> List[str]:
    """Wrapper hérité — utiliser ValidationRules directement dans le nouveau code."""
    return ValidationRules.validate_source_row(r, default_type_echantillon)


# ---------------------------------------------------------------------------
# Résolution de la feuille cible
# ---------------------------------------------------------------------------

def resolve_target_sheet(wb: Any, requested_sheet: str) -> Any:
    """
    Retourne la feuille cible en tenant compte des variantes de nom.
    En cas d'ambiguïté, sélectionne la feuille la plus peuplée.
    """
    if requested_sheet in wb.sheetnames:
        ws = wb[requested_sheet]
        if ws.max_row > 1 or ws.max_column > 1:
            return ws

    req_norm   = normalize_header_name(requested_sheet)
    candidates = []
    for name in wb.sheetnames:
        if normalize_header_name(name) == req_norm:
            ws = wb[name]
            candidates.append((ws.max_row * ws.max_column, ws))
    if candidates:
        candidates.sort(key=lambda x: x[0], reverse=True)
        return candidates[0][1]

    fatal(f"Onglet cible introuvable : {requested_sheet}")


# ---------------------------------------------------------------------------
# Résolution de la feuille type échantillon
# ---------------------------------------------------------------------------

def _resolve_type_sheet(wb: Any) -> Any:
    """Recherche la feuille de types échantillon parmi les candidats connus."""
    names = list(getattr(wb, "sheetnames", []))
    for candidate in TYPE_SHEET_CANDIDATES:
        if candidate in names:
            return wb[candidate]

    normalized = {normalize(n).lower(): n for n in names}
    for candidate in TYPE_SHEET_CANDIDATES:
        key = normalize(candidate).lower()
        if key in normalized:
            return wb[normalized[key]]
    return None


# ---------------------------------------------------------------------------
# Gestion des types d'échantillon
# ---------------------------------------------------------------------------

def load_type_echantillon_options(workbook_path: Path) -> List[str]:
    """Charge la liste des types d'échantillon disponibles depuis le classeur cible."""
    if not workbook_path.exists():
        return []

    openpyxl, _ = ensure_deps()
    wb = openpyxl.load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        ws = _resolve_type_sheet(wb)
        if ws is None:
            return []

        out: List[str] = []
        seen: set = set()
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True):
            val = normalize(row[0] if row else "")
            if not val:
                continue
            if val.lower() in {
                "type echantillon",
                "type ?chantillon",
                "code type echantillon",
                "code type ?chantillon",
                "code echantillon",
                "code ?chantillon",
            }:
                continue
            if val not in seen:
                seen.add(val)
                out.append(val)
        return out
    finally:
        wb.close()


def append_type_echantillon_option(workbook_path: Path, new_type: str) -> bool:
    """Ajoute un nouveau type d'échantillon dans la feuille dédiée du classeur cible."""
    value = normalize(new_type)
    if not value or not workbook_path.exists():
        return False

    openpyxl, _ = ensure_deps()
    wb = openpyxl.load_workbook(workbook_path)
    try:
        ws = _resolve_type_sheet(wb)
        if ws is None:
            ws = wb.create_sheet(TYPE_SHEET_CANDIDATES[0])
            ws.cell(1, 1).value = "Code type échantillon"

        existing = {normalize(row[0] if row else "") for row in ws.iter_rows(values_only=True)}
        if value in existing:
            return True

        next_row = ws.max_row + 1
        if next_row == 1:
            next_row = 2
            ws.cell(1, 1).value = "Code type échantillon"
        ws.cell(next_row, 1).value = value
        wb.save(workbook_path)
        return True
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# Commandes CLI
# ---------------------------------------------------------------------------

def cmd_extract(args: argparse.Namespace) -> None:
    """
    Sous-commande 'extract' — délègue à ExtractionService.

    Filtre la source par plage REF et génère un CSV de sélection.
    """
    _, xlrd = ensure_deps()

    from infrastructure.excel_reader import ExcelReader
    from infrastructure.csv_repository import CSVRepository
    from application.extraction_service import ExtractionService

    excel_reader = ExcelReader(xlrd)
    csv_repo     = CSVRepository()
    service      = ExtractionService(excel_reader, csv_repo)

    start_ref = RefCode.parse(args.start_ref)
    end_ref   = RefCode.parse(args.end_ref)
    if not start_ref or not end_ref:
        fatal(f"Codes REF invalides : {args.start_ref}, {args.end_ref}")

    result = service.extract_range(
        source_path=Path(args.source),
        sheet_name=args.source_sheet,
        start_ref=start_ref,
        end_ref=end_ref,
        output_csv=Path(args.out_csv),
        default_type_echantillon=args.default_type_echantillon,
    )

    print(json.dumps(result.to_dict(), ensure_ascii=False, indent=2))


def cmd_import(args: argparse.Namespace) -> None:
    """
    Sous-commande 'import' — délègue à ImportService.

    Toute la logique complexe (validation, doublons, propagation de formules,
    copie de contexte) est encapsulée dans ImportService.
    """
    openpyxl, xlrd = ensure_deps()

    from infrastructure.excel_reader import ExcelReader
    from infrastructure.excel_writer import ExcelWriter
    from infrastructure.csv_repository import CSVRepository
    from infrastructure.history_repository import HistoryRepository
    from application.import_service import ImportService

    excel_reader = ExcelReader(xlrd)
    excel_writer = ExcelWriter(openpyxl)
    csv_repo     = CSVRepository()
    history_repo = HistoryRepository()
    service      = ImportService(excel_reader, excel_writer, csv_repo, history_repo)

    config = ImportConfig.from_cli_args(args)
    result = service.import_selection(config)

    print(json.dumps(result.to_summary(), ensure_ascii=False, indent=2))


def cmd_reimport(args: argparse.Namespace) -> None:
    """
    Sous-commande 'reimport' — liste les lignes à réimporter depuis l'historique.
    """
    hist = Path(args.history)
    if not hist.exists():
        fatal(f"Historique introuvable : {hist}")

    payload  = json.loads(hist.read_text(encoding="utf-8"))
    rows     = payload.get("rows", [])
    pending  = [r for r in rows if r.get("status") in ("a_reimporter", "non_importe_manuel")]

    if not pending:
        print("Aucune ligne à réimporter.")
        return

    selected_refs = set(args.refs)
    selected = [
        r for r in pending
        if not selected_refs or normalize(r.get("ref")) in selected_refs
    ]

    print(json.dumps(
        {"pending": len(pending), "selected": selected},
        ensure_ascii=False,
        indent=2,
    ))


# ---------------------------------------------------------------------------
# Construction du parseur CLI
# ---------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="datacolisa-importer",
        description="Import Excel → Excel pour DATACOLISA (MVP context master)",
    )
    sub = p.add_subparsers(dest="cmd", required=True)

    # --- extract ---
    e = sub.add_parser(
        "extract",
        help="Filtre la source par plage REF et génère un CSV de sélection",
    )
    e.add_argument("--source",                   required=True, help="Fichier source .xls")
    e.add_argument("--source-sheet",             default=DEFAULT_SOURCE_SHEET)
    e.add_argument("--start-ref",                required=True)
    e.add_argument("--end-ref",                  required=True)
    e.add_argument("--out-csv",                  default="selection_import.csv")
    e.add_argument("--default-type-echantillon", default="EC")
    e.set_defaults(func=cmd_extract)

    # --- import ---
    i = sub.add_parser(
        "import",
        help="Importe vers la cible Feuil1 selon le CSV de sélection",
    )
    i.add_argument("--selection-csv",   required=True)
    i.add_argument("--target",          required=True, help="Fichier cible .xlsx")
    i.add_argument("--target-sheet",    default=DEFAULT_TARGET_SHEET)
    i.add_argument("--out-target",      default="COLISA_imported.xlsx")
    i.add_argument("--history",         default="import_history.json")
    i.add_argument("--default-organisme", default="INRAE")
    i.add_argument("--default-country",   default="France")
    i.add_argument("--on-duplicate",
                   choices=["alert", "ignore", "replace"],
                   default="alert")
    i.set_defaults(func=cmd_import)

    # --- reimport ---
    r = sub.add_parser(
        "reimport",
        help="Liste les lignes marquées non importées ou à réimporter",
    )
    r.add_argument("--history", default="import_history.json")
    r.add_argument("--refs",    nargs="*", default=[])
    r.set_defaults(func=cmd_reimport)

    return p


# ---------------------------------------------------------------------------
# Point d'entrée
# ---------------------------------------------------------------------------

def main() -> None:
    parser = build_parser()
    args   = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
