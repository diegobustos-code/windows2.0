# -*- coding: utf-8 -*-
"""
sede_colors.py
===============
Colores de relleno por sede para los reportes de asistencia: cada sede
tiene su propio color de fila, para que se distingan de un vistazo (igual
que en el listado de socios que se lleva a mano).

Los colores de las sedes conocidas son los mismos que usa el listado
oficial. Si aparece una sede que no está en la tabla, se le asigna uno de
una lista de colores de reserva (siempre el mismo para esa sede dentro de
un mismo archivo). "Sin sede" no lleva color.
"""

import re
import unicodedata
from typing import Dict, Iterable, Optional

from openpyxl.styles import PatternFill

# Clave = nombre de la sede en minúsculas y SIN tildes (la ñ cuenta como n).
_SEDE_PALETTE: Dict[str, str] = {
    "alameda": "DADADA",
    "alto hospicio": "B4C6E7",
    "antofagasta": "FBE4D5",
    "arica": "DADADA",
    "casa central": "C5C2C2",
    "chillan": "B4C6E7",
    "copiapo": "C5E0B3",
    "coquimbo": "BDD6EE",
    "curico": "F7CAAC",
    "gran avenida": "B4C6E7",
    "independencia": "FFE598",
    "iquique": "FBE4D5",
    "la florida": "C5E0B3",
    "la serena": "C5C2C2",
    "las condes": "FFF2CB",
    "los angeles": "C5E0B3",
    "maipu": "DADADA",
    "nunoa": "F7CAAC",
    "osorno": "99E6E6",
    "ovalle": "B7B7FF",
    "providencia": "EAD2F6",
    "puente alto": "9CE9F6",
    "puerto montt": "FBE4D5",
    "quilicura": "DEF880",
    "quilpue": "E2EFD9",
    "rancagua": "CEB288",
    "san bernardo": "A8CDF6",
    "san fernando": "F7CAAC",
    "talca": "B7B7FF",
    "temuco": "BDF852",
    "valdivia": "C8C8C8",
    "valparaiso": "A8CDF6",
    "vina del mar": "DEF880",
}

# Colores de reserva para sedes que no estén en la tabla de arriba.
_FALLBACK_COLORS = [
    "FFD966", "F4B084", "A9D08E", "9BC2E6", "D9B3FF", "FFB3B3", "B3E6CC", "E6E6B3",
]


def norm_sede(text: str) -> str:
    text = unicodedata.normalize("NFKD", (text or "").strip().lower())
    text = "".join(c for c in text if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", re.sub(r"[^a-z\s/]", " ", text)).strip()


def build_sede_fills(sedes: Iterable[str]) -> Dict[str, Optional[PatternFill]]:
    """Arma {sede_normalizada: relleno} para todas las sedes recibidas.
    Las sedes vacías o "sin sede" quedan en None (sin color)."""
    fills: Dict[str, Optional[PatternFill]] = {}
    extra = 0
    for sede in sedes:
        key = norm_sede(sede)
        if key in fills:
            continue
        if not key or key.startswith("sin sede"):
            fills[key] = None
            continue
        color = _SEDE_PALETTE.get(key)
        if color is None:
            color = _FALLBACK_COLORS[extra % len(_FALLBACK_COLORS)]
            extra += 1
        fills[key] = PatternFill(start_color="FF" + color, end_color="FF" + color, fill_type="solid")
    return fills


def fill_for(fills: Dict[str, Optional[PatternFill]], sede: str) -> Optional[PatternFill]:
    return fills.get(norm_sede(sede))
