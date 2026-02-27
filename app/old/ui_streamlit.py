#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import io
import json
from contextlib import redirect_stdout
from pathlib import Path
from typing import Any, Dict, List

import pandas as pd
import streamlit as st

import datacolisa_importer as core

TYPE_PECHE_OPTIONS = ["", "LIGNE", "FILET", "TRAINE", "SONDE"]
CATEGORIE_OPTIONS = ["", "PRO", "AMATEUR"]
OUI_NON_OPTIONS = ["", "OUI", "NON"]
NUMERIC_OPTIONS = [""] + [str(i) for i in range(1, 11)]


def build_rows_for_editor(source_file: Path, source_sheet: str, start_ref: str, end_ref: str, default_type: str) -> Dict[str, Any]:
    _, xlrd = core.ensure_deps()
    source_rows, datemode = core.read_source_rows(xlrd, source_file, source_sheet)
    candidates = core.find_candidate_rows(source_rows, datemode)
    filtered = [r for r in candidates if core.in_ref_range(r.ref, start_ref, end_ref)]
    filtered.sort(key=lambda r: core.parse_ref_parts(r.ref)[1] if core.parse_ref_parts(r.ref) else 0)

    found_codes = {core.normalize(r.ref).upper() for r in filtered}
    missing: List[str] = []
    p_start = core.parse_ref_parts(start_ref)
    p_end = core.parse_ref_parts(end_ref)
    if p_start and p_end and p_start[0] == p_end[0]:
        for n in range(p_start[1], p_end[1] + 1):
            code = f"{p_start[0]}{n}"
            if code not in found_codes:
                missing.append(code)

    rows: List[Dict[str, Any]] = []
    for r in filtered:
        errs = core.validate_row(r, default_type)
        rows.append(
            {
                "selected": False,
                "include": True,
                "status": "a_reimporter" if errs else "pret",
                "source_row": r.source_row_index,
                "ref": core.normalize(r.ref),
                "num_individu": core.normalize(r.num_individu),
                "date_capture": core.normalize(r.date_capture),
                "code_espece": core.normalize(r.code_espece),
                "lac_riviere": core.normalize(r.lac_riviere),
                "categorie": "",
                "type_peche": "TRAINE",
                "autre_oss": "",
                "ecailles_brutes": "",
                "montees": "",
                "otolithes": "",
                "longueur_mm": core.normalize(r.longueur_mm),
                "poids_g": core.normalize(r.poids_g),
                "maturite": core.normalize(r.maturite),
                "sexe": core.normalize(r.sexe),
                "age_total": core.normalize(r.age_total),
                "code_type_echantillon": default_type,
                "errors": " | ".join(errs),
            }
        )
    return {"rows": rows, "missing": missing}


def write_selection_csv(df: pd.DataFrame, out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    headers = [
        "include",
        "status",
        "source_row",
        "ref",
        "num_individu",
        "date_capture",
        "code_espece",
        "lac_riviere",
        "categorie",
        "type_peche",
        "autre_oss",
        "ecailles_brutes",
        "montees",
        "otolithes",
        "longueur_mm",
        "poids_g",
        "maturite",
        "sexe",
        "age_total",
        "code_type_echantillon",
        "errors",
    ]
    with out_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
        w.writeheader()
        for row in df.to_dict(orient="records"):
            payload = dict(row)
            payload["include"] = "1" if bool(payload.get("include")) else "0"
            w.writerow(payload)


def run_import(
    selection_csv: Path,
    target_file: Path,
    target_sheet: str,
    out_target: Path,
    history_path: Path,
    default_org: str,
    default_country: str,
    on_duplicate: str,
) -> Dict[str, Any]:
    args = argparse.Namespace(
        selection_csv=str(selection_csv),
        target=str(target_file),
        target_sheet=target_sheet,
        out_target=str(out_target),
        history=str(history_path),
        default_organisme=default_org,
        default_country=default_country,
        on_duplicate=on_duplicate,
    )
    buf = io.StringIO()
    with redirect_stdout(buf):
        core.cmd_import(args)
    raw = buf.getvalue().strip()
    return json.loads(raw) if raw else {"message": "Import termine"}


def _types_workbook_path(target_path: Path, out_target: Path) -> Path:
    return out_target if out_target.exists() else target_path


def _import_base_workbook(target_path: Path, out_target: Path) -> Path:
    return out_target if out_target.exists() else target_path


def _load_type_options(types_book: Path, fallback_default: str) -> List[str]:
    values = core.load_type_echantillon_options(types_book)
    if fallback_default and fallback_default not in values:
        values.append(fallback_default)
    return sorted(set(v for v in values if str(v).strip()))


def _ensure_editor_columns(df: pd.DataFrame) -> pd.DataFrame:
    defaults = {
        "selected": False,
        "include": True,
        "categorie": "",
        "type_peche": "TRAINE",
        "autre_oss": "",
        "ecailles_brutes": "",
        "montees": "",
        "otolithes": "",
    }
    for k, v in defaults.items():
        if k not in df.columns:
            df[k] = v
    return df


def main() -> None:
    st.set_page_config(page_title="DATACOLISA Import Studio", page_icon="🧪", layout="wide")
    st.title("DATACOLISA Import Studio")
    st.caption("Selection, validation et import Excel -> Excel")

    base_dir = Path(__file__).resolve().parent.parent
    default_source = base_dir / "PacFinalTL14novembrel2012.xls"
    default_target = base_dir / "COLISA 89463-.xlsx"
    default_out = base_dir / "COLISA_imported.xlsx"
    default_history = base_dir / "import_history.json"

    if "editor_df" not in st.session_state:
        st.session_state["editor_df"] = pd.DataFrame()
    if "missing_codes" not in st.session_state:
        st.session_state["missing_codes"] = []
    if "type_options" not in st.session_state:
        st.session_state["type_options"] = []
    if "last_used_type" not in st.session_state:
        st.session_state["last_used_type"] = "EC MONTEE"

    with st.sidebar:
        st.subheader("Parametres")
        source_path = Path(st.text_input("Source .xls", str(default_source)))
        source_sheet = st.text_input("Onglet source", core.DEFAULT_SOURCE_SHEET)
        target_path = Path(st.text_input("Cible .xlsx", str(default_target)))
        target_sheet = st.text_input("Onglet cible", core.DEFAULT_TARGET_SHEET)
        out_target = Path(st.text_input("Sortie .xlsx", str(default_out)))
        history_path = Path(st.text_input("Historique .json", str(default_history)))
        selection_csv = Path(st.text_input("Selection CSV", str(Path(__file__).resolve().parent / "selection_import.csv")))

        import_base = _import_base_workbook(target_path, out_target)
        st.caption(f"Base import cumulative: {import_base}")

        types_book = _types_workbook_path(target_path, out_target)
        loaded_options = _load_type_options(types_book, "EC MONTEE")
        if loaded_options:
            st.session_state["type_options"] = loaded_options

        st.divider()
        st.caption(f"Types depuis: {types_book}")
        type_options = st.session_state["type_options"] or ["EC MONTEE"]
        last_used = st.session_state.get("last_used_type", "EC MONTEE")
        default_index = type_options.index(last_used) if last_used in type_options else 0
        default_type = st.selectbox("Code type echantillon par defaut", options=type_options, index=default_index)
        st.session_state["last_used_type"] = default_type

        new_type_value = st.text_input("Nouveau type echantillon")
        if st.button("Ajouter ce type"):
            if not new_type_value.strip():
                st.warning("Saisir une valeur de type.")
            else:
                ok = core.append_type_echantillon_option(types_book, new_type_value)
                if ok:
                    st.session_state["type_options"] = _load_type_options(types_book, default_type)
                    st.session_state["last_used_type"] = new_type_value.strip()
                    st.success("Type ajoute et enregistre.")
                else:
                    st.error("Impossible d'ajouter le type (verifie le fichier cible/sortie).")

        default_org = st.text_input("Organisme", "INRAE")
        default_country = st.text_input("Pays", "France")
        duplicate_mode = st.selectbox("Doublons", ["alert", "ignore", "replace"], index=0)

    c1, c2, c3 = st.columns([1, 1, 2])
    start_ref = c1.text_input("Code debut", "CA961")
    end_ref = c2.text_input("Code fin", "CA989")
    load_clicked = c3.button("Charger la plage", width="stretch", type="primary")

    if load_clicked:
        try:
            payload = build_rows_for_editor(source_path, source_sheet, start_ref, end_ref, default_type)
            st.session_state["editor_df"] = pd.DataFrame(payload["rows"])
            st.session_state["missing_codes"] = payload["missing"]
        except Exception as exc:
            st.error(f"Chargement impossible: {exc}")

    df = _ensure_editor_columns(st.session_state["editor_df"])
    st.session_state["editor_df"] = df

    if df.empty:
        st.info("Clique sur 'Charger la plage' pour afficher les lignes.")
        return

    k1, k2, k3 = st.columns(3)
    k1.metric("Lignes trouvees", len(df))
    k2.metric("A reimporter (validation)", int((df["status"] == "a_reimporter").sum()))
    k3.metric("Codes manquants", len(st.session_state["missing_codes"]))

    if st.session_state["missing_codes"]:
        with st.expander("Voir les codes non trouves"):
            st.write(", ".join(st.session_state["missing_codes"]))

    st.subheader("Actions en masse")
    refs = sorted(df["ref"].dropna().astype(str).unique().tolist())
    selected_refs = st.multiselect("References ciblees", refs)

    b1, b2, b3, b4 = st.columns(4)
    if b1.button("Cocher include (refs)"):
        df.loc[df["ref"].isin(selected_refs), "include"] = True
    if b2.button("Decocher include (refs)"):
        df.loc[df["ref"].isin(selected_refs), "include"] = False
    bulk_options = st.session_state["type_options"] or [default_type]
    last_used_bulk = st.session_state.get("last_used_type", default_type)
    bulk_index = bulk_options.index(last_used_bulk) if last_used_bulk in bulk_options else 0
    bulk_type_refs = b3.selectbox("Type echantillon (refs)", options=bulk_options, index=bulk_index)
    if b4.button("Attribuer type (refs)"):
        df.loc[df["ref"].isin(selected_refs), "code_type_echantillon"] = bulk_type_refs
        st.session_state["last_used_type"] = bulk_type_refs

    c1, c2, c3, c4 = st.columns(4)
    if c1.button("Selectionner toutes lignes"):
        df["selected"] = True
    if c2.button("Vider selection lignes"):
        df["selected"] = False
    bulk_type_selected = c3.selectbox("Type echantillon (selection)", options=bulk_options, index=bulk_index)
    if c4.button("Attribuer type a la selection"):
        df.loc[df["selected"] == True, "code_type_echantillon"] = bulk_type_selected
        st.session_state["last_used_type"] = bulk_type_selected

    st.markdown("**Completer colonnes metier (lignes selectionnees)**")
    m1, m2, m3, m4, m5, m6, m7 = st.columns(7)
    bulk_categorie = m1.selectbox("Categorie", options=CATEGORIE_OPTIONS, index=0)
    bulk_type_peche = m2.selectbox("Type peche/engin", options=TYPE_PECHE_OPTIONS, index=2)
    bulk_autre_oss = m3.selectbox("Autre oss.", options=OUI_NON_OPTIONS, index=0)
    bulk_ecailles = m4.selectbox("Ecailles brutes", options=OUI_NON_OPTIONS, index=0)
    bulk_montees = m5.selectbox("Montees", options=NUMERIC_OPTIONS, index=0)
    bulk_otolithes = m6.selectbox("Otolithes", options=NUMERIC_OPTIONS, index=0)
    if m7.button("Appliquer metier"):
        mask = df["selected"] == True
        if bulk_categorie != "":
            df.loc[mask, "categorie"] = bulk_categorie
        if bulk_type_peche != "":
            df.loc[mask, "type_peche"] = bulk_type_peche
        if bulk_autre_oss != "":
            df.loc[mask, "autre_oss"] = bulk_autre_oss
        if bulk_ecailles != "":
            df.loc[mask, "ecailles_brutes"] = bulk_ecailles
        if bulk_montees != "":
            df.loc[mask, "montees"] = bulk_montees
        if bulk_otolithes != "":
            df.loc[mask, "otolithes"] = bulk_otolithes

    st.session_state["editor_df"] = df

    st.subheader("Selection et donnees")
    type_choices = list(st.session_state["type_options"] or [default_type])
    for current in st.session_state["editor_df"]["code_type_echantillon"].dropna().astype(str).unique().tolist():
        if current not in type_choices:
            type_choices.append(current)

    st.caption("REF verrouillee: utilise la colonne d'index a gauche pour garder le repere en scroll horizontal.")
    editor_input = st.session_state["editor_df"].set_index("ref", drop=True)

    edited_df = st.data_editor(
        editor_input,
        hide_index=False,
        width="stretch",
        column_config={
            "selected": st.column_config.CheckboxColumn("Selection"),
            "include": st.column_config.CheckboxColumn("Importer"),
            "code_type_echantillon": st.column_config.SelectboxColumn("Type echantillon", options=type_choices, required=True),
            "categorie": st.column_config.SelectboxColumn("Categorie pecheur", options=CATEGORIE_OPTIONS),
            "type_peche": st.column_config.SelectboxColumn("Type peche/engin", options=TYPE_PECHE_OPTIONS),
            "autre_oss": st.column_config.SelectboxColumn("Autre echantillon osseuses", options=OUI_NON_OPTIONS),
            "ecailles_brutes": st.column_config.SelectboxColumn("Ecailles brutes", options=OUI_NON_OPTIONS),
            "montees": st.column_config.SelectboxColumn("Montees", options=NUMERIC_OPTIONS),
            "otolithes": st.column_config.SelectboxColumn("Otolithes", options=NUMERIC_OPTIONS),
            "status": st.column_config.TextColumn("Statut", disabled=True),
            "source_row": st.column_config.NumberColumn("Ligne source", disabled=True),
            "errors": st.column_config.TextColumn("Alerte validation", disabled=True),
        },
        disabled=["status", "source_row", "errors"],
    )

    edited_df = edited_df.reset_index().rename(columns={"index": "ref"})
    st.session_state["editor_df"] = edited_df

    a1, a2 = st.columns(2)
    if a1.button("Enregistrer la selection CSV", width="stretch"):
        try:
            write_selection_csv(edited_df, selection_csv)
            st.success(f"Selection enregistree: {selection_csv}")
        except Exception as exc:
            st.error(f"Echec ecriture CSV: {exc}")

    if a2.button("Lancer l'import vers Excel cible", width="stretch", type="primary"):
        try:
            write_selection_csv(edited_df, selection_csv)
            summary = run_import(
                selection_csv=selection_csv,
                target_file=_import_base_workbook(target_path, out_target),
                target_sheet=target_sheet,
                out_target=out_target,
                history_path=history_path,
                default_org=default_org,
                default_country=default_country,
                on_duplicate=duplicate_mode,
            )
            st.success("Import termine")
            st.json(summary)
        except SystemExit as exc:
            st.error(f"Import arrete: {exc}")
        except Exception as exc:
            st.error(f"Import echoue: {exc}")

    if history_path.exists():
        with st.expander("Historique import"):
            try:
                payload = json.loads(history_path.read_text(encoding="utf-8"))
                st.json(payload)
            except Exception as exc:
                st.warning(f"Lecture historique impossible: {exc}")


if __name__ == "__main__":
    main()
