"""Domain models for DATACOLISA application."""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional


@dataclass
class SourceRow:
    source_row_index: int
    ref: str
    code_espece: Optional[str]
    date_capture: Optional[str]
    lac_riviere: Optional[str]
    num_individu: Optional[str]
    longueur_mm: Optional[str]
    poids_g: Optional[str]
    maturite: Optional[str]
    sexe: Optional[str]
    age_total: Optional[str]
    type_peche: Optional[str]
    categorie: Optional[str]
    pecheur: Optional[str]
    pays_capture: Optional[str]
    pecheur_source: Optional[str]
    observation_disponibilite: Optional[str]
    ecailles_brutes: Optional[str] = ""
    montees: Optional[str] = ""
    empreintes: Optional[str] = ""
    otolithes: Optional[str] = ""
    sous_espece: Optional[str] = ""
    nom_operateur: Optional[str] = ""
    lieu_capture: Optional[str] = ""
    maille_mm: Optional[str] = ""
    code_stade: Optional[str] = ""
    presence_otolithe_gauche: Optional[str] = ""
    presence_otolithe_droite: Optional[str] = ""
    nb_opercules: Optional[str] = ""
    information_stockage: Optional[str] = ""
    age_riviere: Optional[str] = ""
    age_lac: Optional[str] = ""
    nb_fraie: Optional[str] = ""
    ecailles_regenerees: Optional[str] = ""
    observations: Optional[str] = ""

    def validate(self, default_type_echantillon: str = "") -> "ValidationResult":
        from domain.business_rules import ValidationRules
        errors = ValidationRules.validate_source_row(self, default_type_echantillon)
        return ValidationResult(is_valid=len(errors) == 0, errors=errors)

    def to_target_row(self, config: "ImportConfig", code_type_echantillon: str = "") -> "TargetRow":
        from domain.business_rules import TypePecheDeriver, DataTransformations
        type_peche, categorie = TypePecheDeriver.derive_type_and_categorie(self.pecheur_source or "")
        final_type_peche = self.type_peche or type_peche
        final_categorie = self.categorie or categorie
        poids_norm, sexe_norm = DataTransformations.normalize_poids_sexe(self.poids_g, self.sexe)
        lac_norm = DataTransformations.normalize_lac(self.lac_riviere)
        return TargetRow(
            code_type_echantillon=code_type_echantillon,
            code_espece=self.code_espece or "",
            organisme=config.default_organisme,
            pays=self.pays_capture or config.default_country,
            date_capture=self.date_capture,
            lac_riviere=lac_norm,
            categorie=final_categorie or "",
            type_peche=final_type_peche or "",
            num_individu=self.num_individu or "",
            longueur_mm=self.longueur_mm or "",
            poids_g=poids_norm,
            maturite=self.maturite or "",
            sexe=sexe_norm,
            age_total=self.age_total or "",
            ecailles_brutes=self.ecailles_brutes or "",
            montees=self.montees or "",
            empreintes=self.empreintes or "",
            otolithes=self.otolithes or "",
            observation_disponibilite=self.observation_disponibilite or "",
            sous_espece=self.sous_espece or "",
            nom_operateur=self.nom_operateur or "",
            lieu_capture=self.lieu_capture or "",
            maille_mm=self.maille_mm or "",
            code_stade=self.code_stade or "",
            presence_otolithe_gauche=self.presence_otolithe_gauche or "",
            presence_otolithe_droite=self.presence_otolithe_droite or "",
            nb_opercules=self.nb_opercules or "",
            information_stockage=self.information_stockage or "",
            age_riviere=self.age_riviere or "",
            age_lac=self.age_lac or "",
            nb_fraie=self.nb_fraie or "",
            ecailles_regenerees=self.ecailles_regenerees or "",
            observations=self.observations or "",
        )


@dataclass
class TargetRow:
    code_unite_gestionnaire: str = ""
    site_atelier: str = ""
    numero_correspondant: str = ""
    code_type_echantillon: str = ""
    code_echantillon: str = ""
    code_espece: str = ""
    organisme: str = ""
    pays: str = ""
    date_capture: Any = None
    lac_riviere: str = ""
    categorie: str = ""
    type_peche: str = ""
    num_individu: str = ""
    longueur_mm: str = ""
    poids_g: str = ""
    maturite: str = ""
    sexe: str = ""
    age_total: str = ""
    ecailles_brutes: str = ""
    montees: str = ""
    empreintes: str = ""
    otolithes: str = ""
    autre_oss: str = ""
    observation_disponibilite: str = ""
    sous_espece: str = ""
    nom_operateur: str = ""
    lieu_capture: str = ""
    maille_mm: str = ""
    code_stade: str = ""
    presence_otolithe_gauche: str = ""
    presence_otolithe_droite: str = ""
    nb_opercules: str = ""
    information_stockage: str = ""
    age_riviere: str = ""
    age_lac: str = ""
    nb_fraie: str = ""
    ecailles_regenerees: str = ""
    observations: str = ""


@dataclass
class ImportConfig:
    selection_csv: Path
    target_path: Path
    target_sheet: str
    output_path: Path
    history_path: Path
    default_organisme: str
    default_country: str
    on_duplicate: str
    # Nouveaux champs
    default_code_unite_gestionnaire: str = ""
    default_site_atelier: str = ""
    default_numero_correspondant: str = ""
    selection_rows: List[Dict[str, Any]] = field(default_factory=list)
    start_numero: int = 0
    code_echantillon_prefix: str = "T"

    @classmethod
    def from_cli_args(cls, args) -> "ImportConfig":
        return cls(
            selection_csv=Path(args.selection_csv),
            target_path=Path(args.target),
            target_sheet=args.target_sheet,
            output_path=Path(args.out_target),
            history_path=Path(args.history),
            default_organisme=args.default_organisme,
            default_country=args.default_country,
            on_duplicate=args.on_duplicate,
        )


@dataclass
class ValidationResult:
    is_valid: bool
    errors: List[str] = field(default_factory=list)

    def add_error(self, error: str):
        self.is_valid = False
        self.errors.append(error)


@dataclass
class ImportResult:
    imported: List[Dict[str, Any]] = field(default_factory=list)
    skipped_manual: List[Dict[str, Any]] = field(default_factory=list)
    skipped_validation: List[Dict[str, Any]] = field(default_factory=list)
    duplicates: List[Dict[str, Any]] = field(default_factory=list)
    target_out: str = ""
    history_path: str = ""

    @property
    def imported_count(self) -> int:
        return len(self.imported)

    def to_summary(self) -> Dict[str, Any]:
        """Generate summary with detailed lists for dialog display."""
        # Extract REF codes from each category
        imported_refs = [str(row.get("ref", "?")).strip() for row in self.imported]
        duplicate_refs = [str(row.get("ref", "?")).strip() for row in self.duplicates]
        skipped_manual_refs = [str(row.get("ref", "?")).strip() for row in self.skipped_manual]

        # Build validation details with errors
        # Each item is {"row": csv_row, "errors": [...list...]}
        skipped_validation_details = []
        for row in self.skipped_validation:
            csv_row_data = row.get("row", row)  # handle nested {"row": ..., "errors": ...} structure
            errors_raw = row.get("errors", [])
            if isinstance(errors_raw, list):
                error_list = [str(e).strip() for e in errors_raw if str(e).strip()]
            else:
                error_list = [e.strip() for e in str(errors_raw).split("|") if e.strip()]
            ref = str(csv_row_data.get("ref") or csv_row_data.get("num_individu", "?")).strip() or "?"
            skipped_validation_details.append({
                "ref": ref,
                "errors": error_list,
            })

        summary = {
            "imported": self.imported_count,
            "skipped_manual": len(self.skipped_manual),
            "skipped_validation": len(self.skipped_validation),
            "duplicates": len(self.duplicates),
            "target_out": self.target_out,
            "history": self.history_path,
            # Detailed lists for dialog
            "imported_refs": imported_refs,
            "duplicate_refs": duplicate_refs,
            "skipped_manual_refs": skipped_manual_refs,
            "skipped_validation_details": skipped_validation_details,
        }
        return summary


@dataclass
class ExtractionResult:
    rows: List[SourceRow] = field(default_factory=list)
    missing_codes: List[str] = field(default_factory=list)
    found_count: int = 0
    range_spec: str = ""
    extract_csv_path: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "range": self.range_spec,
            "found": self.found_count,
            "missing": self.missing_codes,
            "extract_csv": self.extract_csv_path,
        }
