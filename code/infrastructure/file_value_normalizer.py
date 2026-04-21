from __future__ import annotations

import datetime as dt
import re
import unicodedata
from typing import Any


INTERNAL_NUMERIC_KEYS = {
    "source_row",
    "ecailles_brutes",
    "montees",
    "empreintes",
    "otolithes",
    "longueur_mm",
    "poids_g",
}

COLISA_NUMERIC_HEADERS = {
    "annee",
    "mois",
    "nombre",
    "longueur totale mm",
    "longueur totale estimee 0 pour non 1 pour oui",
    "longueur fourche mm",
    "longueur fourche estimee 0 pour non 1 pour oui",
    "poids g",
    "poux de mer 0 pour non 1 pour oui",
    "presence de l otolithe gauche 0 pour non 1 pour oui",
    "presence de l otolithe droit 0 pour non 1 pour oui",
}


def _normalize_label(value: str) -> str:
    text = str(value or "").strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def coerce_numeric_string(value: Any) -> Any:
    if value is None or value == "":
        return ""
    if isinstance(value, (bool, int, float, dt.date, dt.datetime)):
        return value

    text = str(value).strip()
    if not text:
        return ""

    if re.fullmatch(r"-?(0|[1-9]\d*)", text):
        return int(text)

    decimal_text = text.replace(",", ".")
    if re.fullmatch(r"-?(0|[1-9]\d*)\.\d+", decimal_text):
        return float(decimal_text)

    return value


def coerce_internal_value(key: str, value: Any) -> Any:
    if key in INTERNAL_NUMERIC_KEYS:
        return coerce_numeric_string(value)
    return value


def coerce_colisa_header_value(header_label: str, value: Any) -> Any:
    if _normalize_label(header_label) in COLISA_NUMERIC_HEADERS:
        return coerce_numeric_string(value)
    return value
