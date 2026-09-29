# -*- coding: utf-8 -*-
"""
multi_month_processor.py
=========================
Lógica para el "Reporte Multi-Mes": permite cargar varios archivos CSV de
Zoom (uno por mes, o varios/todos los meses de un año) y cotejarlos TODOS
contra el MISMO listado oficial de socios, para producir una sola tabla
con una fila por socio y UNA COLUMNA POR MES (P/A).

Este módulo NO reinventa el cotejo: llama a
roster_matcher.match_zoom_to_roster(...) una vez POR CADA MES, así que el
comportamiento de emparejar nombres (fuzzy matching), detectar ambiguos y
detectar "no encontrados" es EXACTAMENTE el mismo, mes a mes, que usa el
resto de la aplicación. Lo único nuevo aquí es:
    1. Adivinar el mes a partir del nombre del archivo (para no tener que
       escribirlo a mano cada vez), y
    2. Combinar los resultados de varios meses en una sola tabla, con una
       columna de Asistencia por mes en vez de una sola.
"""

import re
import unicodedata
from typing import Dict, List, Optional, Tuple

from roster_matcher import match_zoom_to_roster

MONTH_NAMES_ES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "setiembre", "octubre",
    "noviembre", "diciembre",
]

_MONTH_DISPLAY = {
    "enero": "Enero", "febrero": "Febrero", "marzo": "Marzo", "abril": "Abril",
    "mayo": "Mayo", "junio": "Junio", "julio": "Julio", "agosto": "Agosto",
    "septiembre": "Septiembre", "setiembre": "Septiembre", "octubre": "Octubre",
    "noviembre": "Noviembre", "diciembre": "Diciembre",
}

# Para reconocer también abreviaciones típicas de nombre de archivo en
# inglés/abreviado (ej. "zoom_sep2026.csv", "asistencia-ago-26.csv").
_MONTH_ABBREV = {
    "ene": "enero", "jan": "enero",
    "feb": "febrero",
    "mar": "marzo",
    "abr": "abril", "apr": "abril",
    "may": "mayo",
    "jun": "junio",
    "jul": "julio",
    "ago": "agosto", "aug": "agosto",
    "sep": "septiembre", "sept": "septiembre",
    "oct": "octubre",
    "nov": "noviembre",
    "dic": "diciembre", "dec": "diciembre",
}


def _normalize(text: str) -> str:
    text = (text or "").strip().lower()
    text = unicodedata.normalize("NFKD", text)
    return "".join(c for c in text if not unicodedata.combining(c))


def detect_month_from_filename(filename: str) -> Optional[str]:
    """
    Intenta adivinar el mes (y año, si aparece) a partir del nombre del
    archivo, por ejemplo:
        "asistencia_Julio_2026.csv"        -> "Julio 2026"
        "Reunion Zoom - agosto 2026.csv"   -> "Agosto 2026"
        "zoom_sep2026_participantes.csv"   -> "Septiembre 2026"
        "asistencia-ago-26.csv"            -> "Agosto 2026" (año de 2 dígitos)

    Devuelve None si no se pudo detectar ningún mes en el nombre; en ese
    caso, el programa debe pedírselo al usuario manualmente.
    """
    normalized = _normalize(filename)
    normalized = re.sub(r"[^a-z0-9]+", " ", normalized)
    # Separa letras y números pegados (ej. "sep2026" -> "sep 2026",
    # "jul26" -> "jul 26") para poder detectar el mes en ambos casos.
    normalized = re.sub(r"(?<=[a-z])(?=[0-9])", " ", normalized)
    normalized = re.sub(r"(?<=[0-9])(?=[a-z])", " ", normalized)

    month_key = None
    for candidate in MONTH_NAMES_ES:
        if re.search(rf"\b{candidate}\b", normalized):
            month_key = candidate
            break
    if month_key is None:
        for abbrev, full in _MONTH_ABBREV.items():
            if re.search(rf"\b{abbrev}\b", normalized):
                month_key = full
                break
    if month_key is None:
        return None

    display = _MONTH_DISPLAY[month_key]

    year_match = re.search(r"\b(20\d{2})\b", normalized)
    if year_match:
        return f"{display} {year_match.group(1)}"

    # Año de 2 dígitos suelto (ej. "ago 26" -> se asume 20XX)
    short_year_match = re.search(rf"\b{month_key[:3]}\w*\s*(\d{{2}})\b", normalized)
    if short_year_match:
        return f"{display} 20{short_year_match.group(1)}"

    return display


def build_multi_month_matrix(
    months: List[Tuple[str, List[Dict[str, object]]]],
    roster_records: List[Dict[str, str]],
    threshold_minutes: float = 30.0,
) -> Tuple[
    List[Dict[str, object]],
    List[str],
    Dict[str, List[Dict[str, object]]],
    Dict[str, List[Dict[str, object]]],
]:
    """
    Coteja el listado oficial de socios contra VARIOS meses de asistencia
    de Zoom, uno a la vez (reutilizando match_zoom_to_roster sin cambios),
    y combina los resultados en una sola tabla: una fila por socio, con
    una columna de Asistencia (P/A) por cada mes.

    `months` es una lista de tuplas (etiqueta_del_mes, registros_zoom_del_mes),
    en el orden en que deben aparecer las columnas en el reporte final.

    Devuelve (filas, etiquetas_de_mes, no_encontrados_por_mes, ambiguos_por_mes):
        filas: lista de diccionarios con "Nombre", "Apellido", "Sede" y,
               además, una llave por cada etiqueta de mes con "P" o "A".
        etiquetas_de_mes: los meses, en el mismo orden recibido (define el
               orden de las columnas en el Excel).
        no_encontrados_por_mes / ambiguos_por_mes: igual que en el cotejo
               normal, pero agrupados por mes, para revisión manual.
    """
    combined: Dict[Tuple[str, str, str], Dict[str, object]] = {}
    order: List[Tuple[str, str, str]] = []
    no_encontrados_por_mes: Dict[str, List[Dict[str, object]]] = {}
    ambiguos_por_mes: Dict[str, List[Dict[str, object]]] = {}

    for month_label, zoom_records in months:
        resultado, no_encontrados, ambiguos = match_zoom_to_roster(
            zoom_records, roster_records, threshold_minutes
        )
        no_encontrados_por_mes[month_label] = no_encontrados
        ambiguos_por_mes[month_label] = ambiguos

        for row in resultado:
            key = (row["Nombre"], row["Apellido"], row["Sede"])
            if key not in combined:
                combined[key] = {
                    "Nombre": row["Nombre"],
                    "Apellido": row["Apellido"],
                    "Sede": row["Sede"],
                }
                order.append(key)
            combined[key][month_label] = row["Asistencia"]

    month_labels = [label for label, _ in months]

    filas: List[Dict[str, object]] = []
    for key in order:
        row = combined[key]
        # Resguardo: match_zoom_to_roster siempre genera una fila por cada
        # socio del listado oficial, así que en la práctica esto no debería
        # activarse nunca, pero se deja por seguridad ante datos atípicos.
        for label in month_labels:
            row.setdefault(label, "A")
        filas.append(row)

    return filas, month_labels, no_encontrados_por_mes, ambiguos_por_mes
