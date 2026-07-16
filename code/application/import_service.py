"""Import service for orchestrating the import workflow."""

from __future__ import annotations

import datetime as dt
import logging
import re
import unicodedata
from typing import Any, Dict, List, Optional, Tuple

from config.mappings import TARGET_HEADERS
from domain.models import ImportConfig, ImportResult
from domain.value_objects import DateCapture, RefCode
from domain.business_rules import DataTransformations, TypePecheDeriver, ReferenceCodeGenerator
from infrastructure.excel_reader import ExcelReader
from infrastructure.excel_writer import ExcelWriter
from infrastructure.csv_repository import CSVRepository
from infrastructure.history_repository import HistoryRepository
from infrastructure.internal_target_workbook import (
    build_code_echantillon_value,
    build_numero_identification_formula,
)

logger = logging.getLogger(__name__)


def _to_int(value: Any) -> Any:
    """Convertit une valeur en entier si possible, sinon retourne None."""
    if value is None or str(value).strip() in ("", "N", "n"):
        return None
    try:
        return int(float(str(value).strip()))
    except (ValueError, TypeError):
        return None


class ImportService:
    """Orchestrates the import workflow from selection CSV to target Excel."""

    def __init__(
        self,
        excel_reader: ExcelReader,
        excel_writer: ExcelWriter,
        csv_repo: CSVRepository,
        history_repo: HistoryRepository
    ):
        self.excel_reader = excel_reader
        self.excel_writer = excel_writer
        self.csv_repo = csv_repo
        self.history_repo = history_repo

    def import_selection(self, config: ImportConfig) -> ImportResult:
        """
        Execute full import workflow.

        Raises:
            IOError: If files cannot be accessed
            ValueError: If data is invalid
        """
        source_label = config.selection_csv.name if config.selection_csv else "table"
        logger.info(f"Starting import: {source_label} -> {config.output_path.name}")

        # 1. Load selection rows
        if config.selection_rows:
            csv_rows = [dict(row) for row in config.selection_rows]
            logger.info(f"Loaded {len(csv_rows)} rows from in-memory table")
        else:
            csv_rows = self.csv_repo.load_selection(config.selection_csv)
            logger.info(f"Loaded {len(csv_rows)} rows from CSV")

        # 2. Load target workbook
        wb = self.excel_writer.load_workbook(config.target_path)
        ws = self._resolve_target_sheet(wb, config.target_sheet)
        logger.debug(f"Target sheet: {ws.title}")

        # 3. Find headers and build indices
        header_row, header_map = self._find_header_row_and_map(ws)
        existing_index = self._build_existing_index(ws, header_row, header_map)

        # 4. Initialize code sequence
        seq_state = ReferenceCodeGenerator.init_sequence_from_workbook(
            ws,
            header_row,
            header_map.get("code_echantillon")
        )
        # Stocker la config dans seq_state pour y acceder dans les sous-fonctions
        seq_state["config"] = config
        # Si le fichier est vide et qu'un numéro de départ est fourni, on l'applique
        if seq_state["num"] == 0 and config.start_numero > 0:
            seq_state["num"] = config.start_numero - 1

        # 5. Process rows
        result = ImportResult()
        history_rows = []
        run_rows: List[int] = []

        for csv_row in csv_rows:
            row_result = self._process_single_row(
                csv_row,
                ws,
                header_row,
                header_map,
                existing_index,
                seq_state,
                run_rows,
                config
            )

            # Collect results
            if row_result["status"] == "imported":
                result.imported.append(csv_row)
                run_rows.append(row_result["target_row"])
            elif row_result["status"] == "imported_replace":
                result.imported.append(csv_row)
                run_rows.append(row_result["target_row"])
            elif row_result["status"] == "imported_merge_ec_ot":
                result.imported.append(csv_row)
            elif row_result["status"] == "skipped_manual":
                result.skipped_manual.append(csv_row)
            elif row_result["status"] == "skipped_validation":
                result.skipped_validation.append({
                    "row": csv_row,
                    "errors": row_result.get("errors", [])
                })
            elif row_result["status"] == "duplicate":
                result.duplicates.append(csv_row)

            history_rows.append({
                "ref": csv_row.get("ref"),
                "status": row_result["status"],
                "reason": row_result.get("reason", "")
            })

        # 6. Save workbook
        self.excel_writer.save_workbook(wb, config.output_path)
        result.target_out = str(config.output_path)

        # 7. Save history
        self.history_repo.save_history(history_rows, config.history_path)
        result.history_path = str(config.history_path)

        return result

    def _process_single_row(
        self,
        csv_row: Dict[str, Any],
        ws,
        header_row: int,
        header_map: Dict[str, int],
        existing_index: Dict[Tuple[str, str], Dict[str, Any]],
        seq_state: Dict[str, Any],
        run_rows: List[int],
        config: ImportConfig
    ) -> Dict[str, Any]:
        """Process a single CSV row; return dict with 'status', 'target_row', 'reason', 'errors'."""
        # 1. Check if selected
        selected = self._normalize(csv_row.get("selected", "1")).lower()
        if selected not in ("1", "true", "yes"):
            return {
                "status": "skipped_manual",
                "reason": "Décoche utilisateur"
            }

        # 2. Validate required fields
        errors = self._validate_csv_row(csv_row)
        if errors:
            return {
                "status": "skipped_validation",
                "errors": errors,
                "reason": " | ".join(errors)
            }

        # 3. Check for duplicates
        key = (
            self._normalize(csv_row.get("num_individu")),
            self._normalize(csv_row.get("code_type_echantillon"))
        )

        if all(key) and key in existing_index:
            # Fusion EC/OT securisee : n'ecrit que les champs vides de la ligne existante
            if csv_row.get("_merge_ec_ot"):
                return self._handle_merge_ec_ot(csv_row, ws, header_map, existing_index, key)
            return self._handle_duplicate(
                csv_row,
                ws,
                header_row,
                header_map,
                existing_index,
                seq_state,
                run_rows,
                config,
                key
            )

        # 4. Write new row
        anchor_col = header_map.get("code_echantillon") or header_map["num_individu"]
        target_row = self._first_empty_row(ws, header_row, anchor_col)
        self._apply_target_row(ws, target_row, csv_row, header_map, config)

        # 5. Propagate formulas
        self._propagate_formulas_and_codes(
            ws,
            target_row,
            header_row,
            header_map,
            seq_state,
            csv_row
        )

        # 6. Copy context
        self._copy_context_if_needed(
            ws,
            target_row,
            header_row,
            header_map,
            run_rows,
            csv_row
        )

        # 7. Update existing index
        if all(key):
            existing_index[key] = {"row": target_row}

        return {
            "status": "imported",
            "target_row": target_row,
            "reason": "OK"
        }

    def _handle_merge_ec_ot(
        self,
        csv_row: Dict[str, Any],
        ws,
        header_map: Dict[str, int],
        existing_index: Dict,
        key: Tuple[str, str]
    ) -> Dict[str, Any]:
        """
        Fusion securisee EC/OT : ecrit uniquement les champs vides/zero de la ligne
        existante avec les valeurs apportees par csv_row.
        Utilise le meme ws que le reste de l'import → pas d'ouverture/fermeture separee.
        """
        CHAMPS = ["ecailles_brutes", "montees", "otolithes", "empreintes"]
        target_row = existing_index[key]["row"]
        modifie = False

        for champ in CHAMPS:
            col = header_map.get(champ)
            if not col:
                continue
            val_existante = ws.cell(target_row, col).value
            vide = val_existante in (None, "", 0)
            if vide:
                nouvelle_val = _to_int(csv_row.get(champ, ""))
                if nouvelle_val is not None and nouvelle_val != 0:
                    ws.cell(target_row, col).value = nouvelle_val
                    self.excel_writer.set_cell_format(ws, target_row, col, "0")
                    modifie = True

        if modifie:
            return {"status": "imported_merge_ec_ot", "target_row": target_row, "reason": "Fusion EC/OT"}
        return {"status": "duplicate", "reason": "Doublon EC/OT (rien a fusionner)"}

    def _handle_duplicate(
        self,
        csv_row: Dict[str, Any],
        ws,
        header_row: int,
        header_map: Dict[str, int],
        existing_index: Dict,
        seq_state: Dict[str, Any],
        run_rows: List[int],
        config: ImportConfig,
        key: Tuple[str, str]
    ) -> Dict[str, Any]:
        """Handle duplicate row based on policy."""
        action = config.on_duplicate

        if action in ("ignore", "alert"):
            return {
                "status": "duplicate",
                "reason": "Doublon exact (ignoré)" if action == "ignore" else "Doublon exact (alerte)"
            }

        if action == "replace":
            target_row = existing_index[key]["row"]
            self._apply_target_row(ws, target_row, csv_row, header_map, config)
            self._propagate_formulas_and_codes(ws, target_row, header_row, header_map, seq_state, csv_row)
            self._copy_context_if_needed(ws, target_row, header_row, header_map, run_rows, csv_row)

            return {
                "status": "imported_replace",
                "target_row": target_row,
                "reason": "Doublon remplacé"
            }

        return {"status": "duplicate", "reason": "Politique inconnue"}

    def _apply_target_row(
        self,
        ws,
        target_row: int,
        csv_row: Dict[str, Any],
        header_map: Dict[str, int],
        config: ImportConfig
    ) -> None:
        def set_if_header(key: str, value: Any) -> None:
            col = header_map.get(key)
            if col:
                ws.cell(target_row, col).value = value

        # Basic fields — code_unite et num_correspondant en format nombre
        for key, val in (
            ("code_unite_gestionnaire", config.default_code_unite_gestionnaire),
            ("numero_correspondant", config.default_numero_correspondant),
        ):
            col = header_map.get(key)
            if col:
                ws.cell(target_row, col).value = val
                if val not in (None, ""):
                    self.excel_writer.set_cell_format(ws, target_row, col, "0")
        set_if_header("site_atelier", config.default_site_atelier)
        set_if_header("code_type_echantillon", csv_row.get("code_type_echantillon", ""))
        set_if_header("code_espece", csv_row.get("code_espece", ""))
        set_if_header("organisme", config.default_organisme)

        # Country
        country = self._normalize(csv_row.get("pays_capture", "")) or config.default_country
        set_if_header("pays", country)

        # Date capture
        date_val = csv_row.get("date_capture", "")
        date_capture = DateCapture.from_excel(date_val)
        if date_capture:
            set_if_header("date_capture", date_capture.date)
            date_col = header_map.get("date_capture")
            if date_col:
                self.excel_writer.set_cell_format(ws, target_row, date_col, "DD/MM/YYYY")
        else:
            set_if_header("date_capture", date_val)

        # Lake
        lac_val = DataTransformations.normalize_lac(csv_row.get("lac_riviere"))
        set_if_header("lac_riviere", lac_val)
        lac_col = header_map.get("lac_riviere")
        if lac_col:
            self.excel_writer.set_cell_format(ws, target_row, lac_col, "@")

        # Type peche and categorie (re-derive from source)
        src_pecheur = self._normalize(csv_row.get("pecheur_source", ""))
        derived_type, derived_cat = TypePecheDeriver.derive_type_and_categorie(src_pecheur)

        type_peche_val = derived_type if derived_type else self._normalize(csv_row.get("type_peche", ""))
        categorie_val = derived_cat if derived_cat else self._normalize(csv_row.get("categorie", ""))

        set_if_header("categorie", categorie_val)
        set_if_header("type_peche", type_peche_val)

        # Autre oss (derived from otolithes)
        autre_oss = DataTransformations.determine_autre_oss(csv_row.get("otolithes"))
        set_if_header("autre_oss", autre_oss)

        # Ecailles / montees / empreintes / otolithes : valeur numerique + format entier
        for key in ("ecailles_brutes", "montees", "empreintes", "otolithes"):
            raw = csv_row.get(key, "")
            col = header_map.get(key)
            if col:
                num_val = _to_int(raw)
                ws.cell(target_row, col).value = num_val
                self.excel_writer.set_cell_format(ws, target_row, col, "0")

        obs_val = csv_row.get("observation_disponibilite", "")
        if not obs_val:
            obs_val = csv_row.get("observation_disponible", "")
        set_if_header("observation_disponibilite", obs_val)
        # Individual data
        set_if_header("num_individu", csv_row.get("num_individu", ""))
        # Longueur totale : valeur numerique
        longueur_col = header_map.get("longueur_mm")
        if longueur_col:
            lon_num = _to_int(csv_row.get("longueur_mm", ""))
            ws.cell(target_row, longueur_col).value = lon_num
            self.excel_writer.set_cell_format(ws, target_row, longueur_col, "0")

        # Normalize weight and sex
        poids_val, sexe_val = DataTransformations.normalize_poids_sexe(
            csv_row.get("poids_g"),
            csv_row.get("sexe")
        )
        set_if_header("poids_g", poids_val)
        set_if_header("sexe", sexe_val)
        set_if_header("maturite", csv_row.get("maturite", ""))
        set_if_header("age_total", csv_row.get("age_total", ""))
        set_if_header("sous_espece", csv_row.get("sous_espece", ""))
        set_if_header("nom_operateur", csv_row.get("nom_operateur", ""))
        set_if_header("lieu_capture", csv_row.get("lieu_capture", ""))
        set_if_header("maille_mm", csv_row.get("maille_mm", ""))
        set_if_header("code_stade", csv_row.get("code_stade", ""))
        set_if_header("presence_otolithe_gauche", csv_row.get("presence_otolithe_gauche", ""))
        set_if_header("presence_otolithe_droite", csv_row.get("presence_otolithe_droite", ""))
        set_if_header("nb_opercules", csv_row.get("nb_opercules", ""))
        set_if_header("information_stockage", csv_row.get("information_stockage", ""))
        set_if_header("age_riviere", csv_row.get("age_riviere", ""))
        set_if_header("age_lac", csv_row.get("age_lac", ""))
        set_if_header("nb_fraie", csv_row.get("nb_fraie", ""))
        set_if_header("ecailles_regenerees", csv_row.get("ecailles_regenerees", ""))
        set_if_header("observations", csv_row.get("observations", ""))

    def _propagate_formulas_and_codes(
        self,
        ws,
        target_row: int,
        header_row: int,
        header_map: Dict[str, int],
        seq_state: Dict[str, Any],
        csv_row: Dict[str, Any]
    ) -> None:
        """Propagate formulas and generate codes."""
        min_row = header_row + 1

        # Incrément séquentiel pour code_echantillon uniquement (T8956, T8957…)
        seq_state["num"] += 1
        num_seq = seq_state["num"]

        code_type_echantillon_val = self._normalize(csv_row.get("code_type_echantillon", ""))
        lac_val = self._normalize(csv_row.get("lac_riviere", ""))
        age_val = self._normalize(csv_row.get("age_total", ""))
        date_val = csv_row.get("date_capture", "")
        date_capture = DateCapture.from_excel(date_val)
        if isinstance(date_capture, dt.datetime):
            date_capture = date_capture.date()

        zero_pad_num = str(num_seq).zfill(int(seq_state.get("width", 5)))
        prefix = seq_state.get("prefix") or getattr(seq_state.get("config"), "code_echantillon_prefix", "T") or "T"

        # code_echantillon = prefixe + numéro séquentiel (T8956…)
        code_echantillon_col = header_map.get("code_echantillon")
        if code_echantillon_col:
            ws.cell(target_row, code_echantillon_col).value = build_code_echantillon_value(
                lac_riviere=lac_val,
                code_type_echantillon=code_type_echantillon_val,
                date_capture=date_capture,
                age_total=age_val,
                numero_individu=zero_pad_num,
                force_prefix=prefix,
            )

        # CODE IDENTIFICATION doit rester une formule Excel, pas une valeur Python.
        numero_identification_col = header_map.get("numero_identification")
        if numero_identification_col:
            ws.cell(target_row, numero_identification_col).value = build_numero_identification_formula(target_row)

        self.excel_writer.propagate_all_formulas(ws, target_row, min_row)

        # Keep code_echantillon formula fixed in the application logic.

    def _copy_context_if_needed(
        self,
        ws,
        target_row: int,
        header_row: int,
        header_map: Dict[str, int],
        run_rows: List[int],
        csv_row: Dict[str, Any]
    ) -> None:
        """Copy context fields from previous rows if needed."""
        parts = RefCode.parse(self._normalize(csv_row.get("num_individu", "")))
        expected_prefix = parts.prefix if parts else ""

        self.excel_writer.copy_context_fields(
            ws,
            target_row,
            header_map,
            source_rows=run_rows,
            min_row=header_row + 1,
            expected_type=self._normalize(csv_row.get("code_type_echantillon", "")),
            expected_ref_prefix=expected_prefix
        )

    # Helper methods

    @staticmethod
    def _normalize(value: Any) -> str:
        """Normalize value to string."""
        if value is None:
            return ""
        if isinstance(value, float):
            if value.is_integer():
                return str(int(value))
            return format(value, "f").rstrip("0").rstrip(".")
        return str(value).strip()

    @staticmethod
    def _validate_csv_row(row: Dict[str, Any]) -> List[str]:
        """Validate CSV row has required fields."""
        errors = []
        if not ImportService._normalize(row.get("date_capture")):
            errors.append("Date capture manquante")
        return errors

    def _resolve_target_sheet(self, wb, requested_sheet: str):
        """Resolve target sheet from workbook."""
        # D'abord essayer une correspondance exacte
        if requested_sheet in wb.sheetnames:
            ws = wb[requested_sheet]
            if ws.max_row > 1 or ws.max_column > 1:
                return ws

        # Chercher par correspondance normalisée du nom
        req_norm = self._normalize(requested_sheet).lower()
        candidates = []
        for name in wb.sheetnames:
            if self._normalize(name).lower() == req_norm:
                ws = wb[name]
                candidates.append((ws.max_row * ws.max_column, ws))

        if candidates:
            candidates.sort(key=lambda x: x[0], reverse=True)
            return candidates[0][1]

        # Si toujours pas trouvé, chercher une feuille qui contient les bonnes en-têtes
        # (colonnes "Code échantillon" et "Numero individu")
        def normalize_header(s: Any) -> str:
            txt = self._normalize(s)
            if not txt:
                return ""
            txt = unicodedata.normalize("NFKD", txt)
            txt = "".join(ch for ch in txt if not unicodedata.combining(ch))
            txt = txt.lower()
            txt = re.sub(r"[^a-z0-9]+", " ", txt)
            return re.sub(r"\s+", " ", txt).strip()
        
        required_headers = {"code echantillon", "numero individu"}
        
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            found_headers = set()
            
            # Chercher les en-têtes dans la première ligne uniquement
            for col_idx in range(1, min(ws.max_column + 1, 50)):
                cell_val = ws.cell(1, col_idx).value
                if cell_val:
                    normalized = normalize_header(cell_val)
                    if "code echantillon" in normalized:
                        found_headers.add("code echantillon")
                    elif "numero individu" in normalized or "numero de capture" in normalized:
                        found_headers.add("numero individu")
            
            if required_headers.issubset(found_headers):
                return ws

        raise ValueError(f"Onglet cible introuvable: {requested_sheet}")

    def _find_header_row_and_map(self, ws) -> Tuple[int, Dict[str, int]]:
        """Find header row and build column mapping."""

        def normalize_header(s: Any) -> str:
            txt = self._normalize(s)
            if not txt:
                return ""
            txt = unicodedata.normalize("NFKD", txt)
            txt = "".join(ch for ch in txt if not unicodedata.combining(ch))
            txt = txt.lower()
            txt = re.sub(r"[^a-z0-9]+", " ", txt)
            return re.sub(r"\s+", " ", txt).strip()

        alias_norm = {}
        for key, aliases in TARGET_HEADERS.items():
            alias_norm[key] = {normalize_header(a) for a in aliases if normalize_header(a)}

        logger.debug(f"Searching headers in sheet: {ws.title}, max_row={ws.max_row}, max_column={ws.max_column}")

        header_map = {}
        found_columns = {}  # Pour déboguer
        all_raw_headers = []  # Pour afficher tous les en-têtes trouvés
        
        for r in range(1, min(51, ws.max_row + 1)):
            values = [self._normalize(ws.cell(r, c).value) for c in range(1, min(ws.max_column + 1, 300))]
            
            # Debug: log les premières colonnes de chaque ligne
            if r == 1 and values:
                logger.debug(f"Line {r} raw values (first 15): {values[:15]}")
                all_raw_headers = values
            
            for idx, val in enumerate(values, start=1):
                val_norm = normalize_header(val)
                if not val_norm:
                    continue
                for key, aliases in alias_norm.items():
                    if val_norm in aliases:
                        header_map[key] = idx
                        found_columns[key] = val  # Pour déboguer
                        logger.debug(f"Found '{key}' at line {r}, column {idx}: '{val}' -> normalized: '{val_norm}'")
            
            if "num_individu" in header_map and "code_echantillon" in header_map:
                logger.debug(f"Headers found at line {r}: {header_map}")
                return r, header_map
            
            # Si on a trouvé au moins une colonne à la ligne 1, mais pas les deux,
            # continuer à chercher dans les prochaines lignes avec moins de rigueur
            if r == 1 and found_columns:
                # Réinitialiser pour continuer, mais ne pas effacer si on n'a rien trouvé
                if not ("num_individu" in found_columns or "code_echantillon" in found_columns):
                    header_map.clear()
                    found_columns.clear()
            else:
                header_map.clear()
                found_columns.clear()

        # Si on n'a rien trouvé après 50 lignes, essayer de recréer les en-têtes
        logger.error(f"Could not find headers after searching {min(51, ws.max_row + 1)} rows")
        logger.error(f"All raw headers from row 1: {all_raw_headers}")
        
        # Vérifier si c'est un fichier vide ou presque vide
        if ws.max_row <= 1 and (not all_raw_headers or all(not v for v in all_raw_headers)):
            logger.info("Detected empty or header-less file, attempting to add headers...")
            try:
                from infrastructure.internal_target_workbook import HEADER_POSITIONS
                
                # Ajouter les en-têtes à la ligne 1
                for col_index in range(1, len(HEADER_POSITIONS) + 1):
                    header = HEADER_POSITIONS.get(col_index, "")
                    ws.cell(1, col_index).value = header if header else None
                
                # Sauvegarder le workbook
                wb = ws.parent
                wb.save(ws.parent.filename)
                logger.info("Headers added successfully")
                
                # Relancer la recherche
                return self._find_header_row_and_map(ws)
            except Exception as e:
                logger.error(f"Failed to add headers: {e}")
        
        logger.error(f"Expected to find keys: 'num_individu' and 'code_echantillon'")
        logger.error(f"Looking for normalized headers like: {alias_norm.get('code_echantillon', {})} and {alias_norm.get('num_individu', {})}")
        raise ValueError(f"Impossible de trouver la ligne d'en-tête (colonnes trouvées: {list(found_columns.keys())})")

    def _build_existing_index(
        self,
        ws,
        header_row: int,
        header_map: Dict[str, int]
    ) -> Dict[Tuple[str, str], Dict[str, Any]]:
        """Build index of existing rows."""
        idx = {}
        num_col = header_map.get("num_individu")
        type_col = header_map.get("code_type_echantillon")
        if not num_col or not type_col:
            return idx

        for r in range(header_row + 1, ws.max_row + 1):
            num = self._normalize(ws.cell(r, num_col).value)
            typ = self._normalize(ws.cell(r, type_col).value)
            if not num:
                continue
            idx[(num, typ)] = {"row": r}

        return idx

    def _first_empty_row(self, ws, header_row: int, anchor_col: int) -> int:
        """Find first empty row after header."""
        r = header_row + 1
        while ws.cell(r, anchor_col).value not in (None, ""):
            r += 1
        return r
