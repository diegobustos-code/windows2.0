# -*- coding: utf-8 -*-
"""
roster_matcher.py
==================
Módulo encargado de cotejar el listado OFICIAL de socios (un Excel con las
columnas Sede, Apellido Paterno, Apellido Materno, Nombres) contra los
participantes leídos desde el CSV de Zoom, para producir un listado de
asistencia basado en la lista oficial: una fila POR CADA SOCIO, marcado
como "P" (Presente) o "A" (Ausente) según si se le encontró y cuántos
minutos acumuló en la reunión de Zoom.

Por qué existe este módulo:
    El nombre que la gente escribe en Zoom es poco confiable (apodos,
    sedes pegadas al nombre, orden invertido, apellidos incompletos, etc.).
    En cambio, el listado oficial de socios es la fuente de verdad. Este
    módulo usa el listado oficial como base y busca, para cada socio, si
    hay algún participante de Zoom cuyo nombre coincida razonablemente.

Cómo funciona el cruce (en resumen):
    1. Se "tokeniza" (separa en palabras normalizadas, sin tildes ni
       mayúsculas) tanto el nombre de cada socio oficial como el nombre de
       cada participante de Zoom.
    2. Para que un socio sea candidato a coincidir con un participante de
       Zoom, al menos su apellido paterno o su apellido materno debe
       aparecer entre las palabras del nombre de Zoom.
    3. Entre los candidatos, se elige el que tiene más palabras en común
       (con más peso si el apellido paterno o materno coincide completo).
       Si hay empate exacto entre dos o más candidatos, ese participante
       de Zoom se marca como "ambiguo" y NO se asigna automáticamente
       (para evitar adjudicar minutos de asistencia a la persona
       equivocada); queda disponible para revisión manual.
    4. Si NINGÚN socio coincide por apellido (paso 2), se intenta una
       SEGUNDA FASE: cotejar solo por el nombre de pila, comparando
       contra el campo "Nombres" del socio. Esto cubre el caso de
       alguien que en Zoom escribió nomás su nombre, sin apellido. Si ese
       nombre es único en el listado (ningún otro socio comparte ese
       nombre), se asigna igual; si dos o más socios comparten ese mismo
       nombre, el participante queda "ambiguo" (nunca se asigna a
       ciegas). Esta segunda fase NUNCA se usa si ya hubo al menos un
       candidato por apellido en la primera fase.
    5. Si un participante de Zoom no coincide con ningún socio (ni por
       apellido ni por nombre), se marca como "no encontrado" (puede ser
       un invitado, un error de tipeo muy grande, o alguien que no está
       en el listado oficial).
    6. Si un mismo socio aparece más de una vez en el Zoom (por ejemplo,
       se desconectó y volvió a entrar), sus minutos se SUMAN.

Nada de esto es 100% infalible con datos tan heterogéneos como los
nombres escritos a mano en Zoom, por lo que el resultado incluye listas
separadas de "no encontrados" y "ambiguos" para que se puedan revisar y
corregir manualmente si hace falta.
"""

import re
import unicodedata
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook

from csv_processor import compute_attendance_flag


class RosterProcessingError(Exception):
    """Error controlado y con mensaje claro para mostrar al usuario final."""
    pass


ROSTER_COLUMNS = ["Sede", "Apellido Paterno", "Apellido Materno", "Nombres"]

# Formato alternativo aceptado para el listado de socios: el apellido
# viene en UNA SOLA columna (no separado en Paterno/Materno) y la columna
# de nombre puede tener cualquier título (incluso venir mal escrito o
# vacío) — se usa la que sobra una vez identificadas "Apellido" y "Sede".
# Ver _detect_simple_format() más abajo.
ROSTER_COLUMNS_SIMPLE = ["Sede", "Apellido"]

# Palabras que se ignoran al comparar nombres (no aportan para identificar
# a una persona: cargos, dispositivos, artículos, restos de "sede", etc.)
STOPWORDS = {
    "sede", "cede", "prof", "profesor", "profesora", "srta", "sra", "sr",
    "don", "dona", "doña", "usuario", "zoom", "invitado", "guest",
    "de", "del", "la", "el", "los", "las", "y",
    "iphone", "ipad", "galaxy", "samsung", "xiaomi", "huawei", "android",
    "moto", "macbook", "note", "ultra", "pro", "plus", "mini", "windows",
    "pc", "tablet", "of",
    "cpech",
}


def _normalize(text: str) -> str:
    text = (text or "").strip().lower()
    text = unicodedata.normalize("NFKD", text)
    return "".join(c for c in text if not unicodedata.combining(c))


def tokenize(text: str) -> List[str]:
    """Separa un texto en palabras normalizadas útiles para comparar nombres."""
    normalized = _normalize(text)
    normalized = re.sub(r"[^a-z\s]", " ", normalized)
    return [t for t in normalized.split() if len(t) >= 2 and t not in STOPWORDS]


def _find_exact_column(headers: List[str], wanted: str) -> Optional[int]:
    """Busca en `headers` la columna cuyo título normalizado sea EXACTAMENTE
    igual a `wanted` (ej. "Apellido" no matchea "Apellido Paterno"). Devuelve
    el índice de la columna, o None si no se encontró."""
    wanted_norm = _normalize(wanted)
    for idx, h in enumerate(headers):
        if _normalize(h) == wanted_norm:
            return idx
    return None


def _find_exact_column_variants(headers: List[str], variants: List[str]) -> Optional[int]:
    """Igual que _find_exact_column, pero probando varias variantes de
    título en orden de preferencia (ej. "Sede" antes que "Sedes"): se
    recorren TODOS los encabezados buscando la primera variante, y solo
    si no aparece en ninguno se pasa a la siguiente variante. Así, si el
    archivo tiene una columna exacta "Sede" Y además otra columna suelta
    llamada "Sedes" (como puede pasar con una columna vacía sobrante), se
    sigue prefiriendo siempre la coincidencia exacta "Sede"."""
    for wanted in variants:
        idx = _find_exact_column(headers, wanted)
        if idx is not None:
            return idx
    return None


def _detect_full_format(headers: List[str]) -> Optional[Dict[str, int]]:
    """Intenta detectar el formato ORIGINAL del listado de socios: columnas
    Sede, Apellido Paterno, Apellido Materno, Nombres (una columna por
    cada una, el apellido ya separado). Devuelve un diccionario
    {columna_canónica: índice} si TODAS están presentes, o None si falta
    alguna (en cuyo caso se prueba el formato simple)."""
    col_index: Dict[str, int] = {}
    for wanted in ROSTER_COLUMNS:
        variants = ["Sede", "Sedes"] if wanted == "Sede" else [wanted]
        idx = _find_exact_column_variants(headers, variants)
        if idx is None:
            return None
        col_index[wanted] = idx
    return col_index


def _detect_simple_format(headers: List[str]) -> Optional[Dict[str, int]]:
    """
    Intenta detectar el formato ALTERNATIVO del listado de socios: el
    apellido viene en UNA SOLA columna llamada exactamente "Apellido" (no
    separado en Paterno/Materno), más una columna "Sede", más una tercera
    columna con el nombre — cuyo título puede ser cualquier cosa ("Nombre",
    "Nombres", o incluso venir mal escrito), ya que se toma la primera
    columna con encabezado no vacío que no sea ni "Apellido" ni "Sede".

    Devuelve {"Sede": idx, "Apellido": idx, "Nombres": idx} si se pudo
    armar el trío completo, o None si falta "Apellido"/"Apellidos" o
    "Sede"/"Sedes" (en cuyo caso el archivo no corresponde a ninguno de
    los dos formatos aceptados) — si lo único que falta es la columna de
    nombre, igual devuelve None, para que el mensaje de error final sea
    claro.
    """
    sede_idx = _find_exact_column_variants(headers, ["Sede", "Sedes"])
    apellido_idx = _find_exact_column_variants(headers, ["Apellido", "Apellidos"])
    if sede_idx is None or apellido_idx is None:
        return None

    nombre_idx = None
    for idx, h in enumerate(headers):
        if idx in (sede_idx, apellido_idx):
            continue
        if h.strip():  # cualquier encabezado no vacío sirve como columna de nombre
            nombre_idx = idx
            break
    if nombre_idx is None:
        return None

    return {"Sede": sede_idx, "Apellido": apellido_idx, "Nombres": nombre_idx}


def load_roster_excel(filepath: str) -> List[Dict[str, str]]:
    """
    Carga el listado oficial de socios desde un archivo Excel (.xlsx).

    Acepta DOS formatos de columnas (el orden de las columnas no importa
    en ninguno de los dos, se detectan por el título del encabezado):

      1. Formato completo: Sede, Apellido Paterno, Apellido Materno,
         Nombres (el apellido ya viene separado en dos columnas).

      2. Formato simple: Sede, Apellido (el apellido completo en UNA sola
         columna) y una tercera columna con el nombre, con cualquier
         título. En este caso, el apellido completo se trata igual que
         el "Apellido Paterno" del formato completo (sin dividirlo), lo
         que no afecta el cotejo de nombres.

    Lanza RosterProcessingError con un mensaje claro si el archivo no
    corresponde a ninguno de los dos formatos, o si está vacío.
    """
    try:
        wb = load_workbook(filepath, data_only=True)
    except Exception as e:
        raise RosterProcessingError(
            f"No se pudo abrir el archivo Excel del listado de socios.\n"
            f"Detalle técnico: {e}"
        )

    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise RosterProcessingError("El listado de socios está vacío.")

    headers = [str(h).strip() if h is not None else "" for h in rows[0]]

    full_cols = _detect_full_format(headers)
    simple_cols = None if full_cols else _detect_simple_format(headers)

    if full_cols is not None:
        col_index = full_cols
        apellido_materno_idx = col_index["Apellido Materno"]
    elif simple_cols is not None:
        # Se reutiliza la llave "Apellido Paterno" para guardar el apellido
        # completo (sin dividir); "Apellido Materno" queda sin columna
        # (se completa como cadena vacía más abajo). El resto del programa
        # (cotejo, exportación) no necesita ningún cambio para esto, ya
        # que solo le importa el conjunto de palabras del apellido, no en
        # cuántas columnas venía dividido.
        col_index = {
            "Sede": simple_cols["Sede"],
            "Apellido Paterno": simple_cols["Apellido"],
            "Nombres": simple_cols["Nombres"],
        }
        apellido_materno_idx = None
    else:
        raise RosterProcessingError(
            "El listado de socios no tiene ninguno de los formatos de "
            "columnas aceptados.\n\n"
            "Encabezados encontrados en el archivo:\n"
            + ", ".join(h for h in headers if h)
            + "\n\nSe acepta cualquiera de estos dos formatos de columnas "
            "(\"Sede\"/\"Sedes\" y \"Apellido\"/\"Apellidos\" son equivalentes):\n"
            "  • Sede, Apellido Paterno, Apellido Materno, Nombres\n"
            "  • Sede, Apellido (el apellido completo en una sola columna), "
            "y una columna con el nombre (con cualquier título)"
        )

    roster: List[Dict[str, str]] = []
    for row in rows[1:]:
        if row is None or not any(row):
            continue

        def _cell(idx: Optional[int]) -> str:
            if idx is None or idx >= len(row) or row[idx] is None:
                return ""
            return str(row[idx]).strip()

        record = {
            "Sede": _cell(col_index["Sede"]),
            "Apellido Paterno": _cell(col_index["Apellido Paterno"]),
            "Apellido Materno": _cell(apellido_materno_idx),
            "Nombres": _cell(col_index["Nombres"]),
        }
        # Ignorar filas completamente vacías de nombre
        if not record["Apellido Paterno"] and not record["Apellido Materno"] and not record["Nombres"]:
            continue
        roster.append(record)

    if not roster:
        raise RosterProcessingError(
            "El listado de socios se leyó correctamente, pero no se "
            "encontraron filas de socios válidas."
        )

    return roster


def _score(entry: Dict[str, object], zoom_tokens: set) -> Optional[Tuple[int, float]]:
    """
    Calcula qué tan bien coincide un socio (`entry`) con las palabras del
    nombre de un participante de Zoom (`zoom_tokens`). Devuelve None si ni
    siquiera el apellido paterno ni el materno aparecen (candidato
    descartado de plano). Si hay coincidencia, devuelve una tupla
    (puntaje_bruto, similitud_jaccard) usada para ordenar candidatos.
    """
    paterno: set = entry["_paterno"]
    materno: set = entry["_materno"]
    all_tokens: set = entry["_all"]

    if not (paterno & zoom_tokens) and not (materno & zoom_tokens):
        return None

    overlap = all_tokens & zoom_tokens
    raw_score = len(overlap)
    if paterno and paterno.issubset(zoom_tokens):
        raw_score += 2
    if materno and materno.issubset(zoom_tokens):
        raw_score += 1

    union = all_tokens | zoom_tokens
    jaccard = len(overlap) / len(union) if union else 0.0
    return (raw_score, jaccard)


def _score_by_nombre(entry: Dict[str, object], zoom_tokens: set) -> Optional[Tuple[int, float]]:
    """
    Segunda fase de cotejo, usada SOLO cuando un participante de Zoom no
    tuvo ningún candidato por apellido (_score devolvió None para todos
    los socios) — el caso típico de alguien que en Zoom escribió nomás su
    nombre de pila, sin apellido. Aquí se compara contra el campo
    "Nombres" del socio en vez del apellido.

    Devuelve None si no hay ninguna palabra en común. Si la hay, se
    comporta igual que _score: se usa para desempatar entre varios
    candidatos débiles, y si al final queda más de uno empatado en primer
    lugar, match_zoom_to_roster lo marca como "ambiguo" en vez de
    asignarlo a ciegas — así que un nombre de pila repetido entre varios
    socios NUNCA se asigna solo, solo cuando es único en el listado.
    """
    nombres: set = entry["_nombres"]
    overlap = nombres & zoom_tokens
    if not overlap:
        return None
    raw_score = len(overlap)
    union = nombres | zoom_tokens
    jaccard = len(overlap) / len(union) if union else 0.0
    return (raw_score, jaccard)


# ---------------------------------------------------------------------------
# Herramientas del cotejo "intensivo" (se usan SOLO cuando el cotejo clásico
# no logra asignar a alguien: no encontrado o ambiguo). Nunca cambian una
# asignación que el cotejo clásico ya hacía de forma única.
# ---------------------------------------------------------------------------
def _edit_distance(a: str, b: str, limit: int) -> int:
    """Distancia de edición (Levenshtein) con corte temprano: si supera
    `limit` devuelve limit + 1."""
    if abs(len(a) - len(b)) > limit:
        return limit + 1
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        if min(cur) > limit:
            return limit + 1
        prev = cur
    return prev[-1]


def _tokens_similar(a: str, b: str) -> bool:
    """True si dos palabras son casi iguales (un error de tipeo: tilde
    mal puesta, letra cambiada, letra de más o de menos). Solo palabras de
    5 letras o más, para no confundir apellidos cortos distintos."""
    if a == b or len(a) < 5 or len(b) < 5:
        return False
    limit = 2 if min(len(a), len(b)) >= 10 else 1
    return _edit_distance(a, b, limit) <= limit


def _sede_key(text: str) -> str:
    key = re.sub(r"[^a-z\s]", " ", _normalize(text))
    key = re.sub(r"\bsede\b", " ", key)
    return re.sub(r"\s+", " ", key).strip()


def _same_sede(zoom_sede: str, roster_sede: str) -> bool:
    a, b = _sede_key(zoom_sede), _sede_key(roster_sede)
    if not a or not b:
        return False
    if a == b:
        return True
    return len(a) >= 4 and len(b) >= 4 and (a in b or b in a)


def _evidencia_nombre(entry: Dict[str, object], zoom_tokens: set) -> str:
    """"exacto" si alguna palabra del nombre de Zoom coincide con un nombre
    de pila del socio; "parecido" si solo coincide con una letra distinta
    (ej. gisselle/giselle, yngrid/ingrid); "" si no hay ninguna."""
    nombres = entry["_nombres"]
    if nombres & zoom_tokens:
        return "exacto"
    if any(_tokens_similar(zt, n) for zt in zoom_tokens for n in nombres):
        return "parecido"
    return ""


# Apodos muy comunes -> nombres formales. Se AGREGAN a lo que escribió la
# persona (no lo reemplazan), y solo ayudan cuando el apellido calza.
_APODOS = {
    "maite": {"maria", "teresa"},
    "pancho": {"francisco"},
    "nacho": {"ignacio"},
    "pepe": {"jose"},
    "lucho": {"luis"},
    "kike": {"enrique"},
    "memo": {"guillermo"},
}


def _split_joined_token(token: str, vocab: set) -> Optional[List[str]]:
    """Si `token` es una palabra larga que NO existe en el listado pero se
    puede armar pegando palabras que sí existen (ej. "mariateresacortes" ->
    maria + teresa + cortes), devuelve esas partes (la segmentación con
    menos partes). Si no, None."""
    n = len(token)
    if n < 8 or token in vocab:
        return None
    best: List[Optional[List[str]]] = [None] * (n + 1)
    best[0] = []
    for i in range(3, n + 1):
        for j in range(0, i - 2):
            if best[j] is None:
                continue
            part = token[j:i]
            if part in vocab:
                cand = best[j] + [part]
                if best[i] is None or len(cand) < len(best[i]):
                    best[i] = cand
    parts = best[n]
    if parts and len(parts) >= 2 and max(len(p) for p in parts) >= 4:
        return parts
    return None


def match_zoom_to_roster(
    zoom_records: List[Dict[str, object]],
    roster_records: List[Dict[str, str]],
    attendance_threshold: float = 30.0,
) -> Tuple[List[Dict[str, object]], List[Dict[str, object]], List[Dict[str, object]]]:
    """
    Igual que siempre: devuelve (resultado, no_encontrados, ambiguos).
    Internamente usa match_zoom_to_roster_detailed (que además informa qué
    asignaciones no fueron "exactas"); ver esa función para el detalle.
    """
    resultado, no_encontrados, ambiguos, _aprox = match_zoom_to_roster_detailed(
        zoom_records, roster_records, attendance_threshold
    )
    return resultado, no_encontrados, ambiguos


def match_zoom_to_roster_detailed(
    zoom_records: List[Dict[str, object]],
    roster_records: List[Dict[str, str]],
    attendance_threshold: float = 30.0,
) -> Tuple[
    List[Dict[str, object]], List[Dict[str, object]],
    List[Dict[str, object]], List[Dict[str, object]],
]:
    """
    Coteja los participantes de Zoom contra el listado oficial de socios.

    Devuelve (resultado, no_encontrados, ambiguos, aproximados):

      - resultado: una fila POR CADA SOCIO del listado oficial, EN EL MISMO
        ORDEN del listado, con las llaves "Nombre", "Apellido", "Sede",
        "Duración (minutos)" y "Asistencia" (P/A).

      - no_encontrados: participantes de Zoom que no coincidieron con
        ningún socio (para revisión manual).

      - ambiguos: participantes de Zoom que siguen coincidiendo con dos o
        más socios tras todos los desempates. Cada uno trae además la llave
        "Posibles" con los socios candidatos, para decidir a mano.

      - aproximados: asignaciones que el programa SÍ hizo pero que no fueron
        por apellido exacto (solo nombre único, apellido con otra
        ortografía, nombre escrito todo junto, desempate por sede...).
        Se entregan para que se puedan revisar; ya están contadas en
        "resultado".

    Orden de decisión para cada participante de Zoom:
      1. Por apellido (paterno o materno) exacto -> el clásico.
      2. Si nadie coincide por apellido: por nombre de pila si es único;
         si no, por apellido con ortografía parecida / iniciales + nombre.
      3. Si hay empate: se desempata con la SEDE que la persona escribió en
         Zoom, y luego con palabras de ortografía parecida.
      4. Si sigue sin haber candidatos: se prueba separando un nombre
         escrito todo junto ("mariateresacortes").
      5. Si todo falla -> "no encontrado" o "ambiguo" (nunca se adivina).
    """
    entries = []
    vocab: set = set()
    nombre_vocab: set = set()
    for r in roster_records:
        paterno_list = tokenize(r["Apellido Paterno"])
        materno_list = tokenize(r["Apellido Materno"])
        if not materno_list and len(paterno_list) == 2:
            # Listado con el apellido completo en UNA columna ("RIFFO VEGA"):
            # se separa en paterno + materno para que quien escribe solo un
            # apellido en Zoom no quede en desventaja frente a quien tiene el
            # mismo apellido repetido ("RIFFO RIFFO").
            paterno_list, materno_list = [paterno_list[0]], [paterno_list[1]]
        paterno = set(paterno_list)
        materno = set(materno_list)
        nombres = set(tokenize(r["Nombres"]))
        inic = ""
        if paterno_list and materno_list:
            inic = paterno_list[0][0] + materno_list[0][0]
        elif len(paterno_list) >= 2:  # formato simple: ambos apellidos en una columna
            inic = paterno_list[0][0] + paterno_list[1][0]
        entries.append({
            **r,
            "_paterno": paterno,
            "_materno": materno,
            "_nombres": nombres,
            "_all": paterno | materno | nombres,
            "_inic": inic,
            "_duracion_total": 0.0,
            "_matched_zoom_names": [],
        })
        vocab |= {t for t in (paterno | materno | nombres) if len(t) >= 3}
        nombre_vocab |= nombres

    def _label(i: int) -> str:
        e = entries[i]
        full = f"{e['Nombres']} {e['Apellido Paterno']} {e['Apellido Materno']}"
        full = re.sub(r"\s{2,}", " ", full).strip()
        return f"{full} ({e['Sede']})"

    def _desempatar(idxs: List[int], zoom_sede: str, zoom_tokens: set):
        """Intenta dejar UN solo socio entre los empatados. Devuelve
        (idx_o_None, texto_del_desempate, candidatos_que_quedan)."""
        if len(idxs) == 1:
            return idxs[0], "", idxs
        if zoom_sede:
            por_sede = [i for i in idxs if _same_sede(zoom_sede, entries[i]["Sede"])]
            if len(por_sede) == 1:
                return por_sede[0], "desempatado por sede", por_sede
            if len(por_sede) > 1:
                idxs = por_sede
        puntajes = [
            sum(1 for zt in zoom_tokens for t in entries[i]["_all"] if _tokens_similar(zt, t))
            for i in idxs
        ]
        mejor = max(puntajes)
        if mejor > 0:
            top = [i for i, p in zip(idxs, puntajes) if p == mejor]
            if len(top) == 1:
                return top[0], "desempatado por nombre parecido", top
            idxs = top
        return None, "", idxs

    def _fuzzy_apellido(zoom_tokens: set, zoom_sede: str):
        """Apellido con otra ortografía (o iniciales), SIEMPRE junto a un
        nombre de pila que coincida exacto. Devuelve (estado, idx, método,
        candidatos) o None si no hay ningún candidato."""
        fuzzy = []
        for idx, entry in enumerate(entries):
            apell = entry["_paterno"] | entry["_materno"]
            casi = [
                zt for zt in zoom_tokens
                if zt not in entry["_nombres"] and any(_tokens_similar(zt, a) for a in apell)
            ]
            iniciales = bool(entry["_inic"]) and entry["_inic"] in zoom_tokens
            if not casi and not iniciales:
                continue
            hits = len(entry["_nombres"] & zoom_tokens)
            if hits == 0:
                continue
            fuzzy.append((hits, idx, "iniciales del apellido" if (iniciales and not casi) else "apellido con ortografía parecida"))
        if not fuzzy:
            return None
        mejor = max(h for h, _i, _m in fuzzy)
        top_f = [(i, m) for h, i, m in fuzzy if h == mejor]
        idxs_f = [i for i, _m in top_f]
        idx, nota, resto = _desempatar(idxs_f, zoom_sede, zoom_tokens)
        if idx is not None:
            met = dict(top_f)[idx]
            return "ok", idx, met + (", " + nota if nota else ""), resto
        return "ambiguo", None, "", resto

    def _resolver(zoom_tokens: set, zoom_sede: str):
        """Devuelve (estado, idx, método, candidatos) con estado 'ok',
        'ambiguo' o 'no'."""
        # 1) Por apellido exacto (cotejo clásico)
        candidates = []
        for idx, entry in enumerate(entries):
            s = _score(entry, zoom_tokens)
            if s is not None:
                candidates.append((s, idx))
        if candidates:
            todos = [idx for _s, idx in candidates]
            score_of = {idx: s for s, idx in candidates}

            # Evidencia de que el NOMBRE de pila también coincide (exacto o
            # con una letra distinta). Entre quienes comparten el apellido,
            # se prefiere a quien además coincide en el nombre.
            evid = {i: _evidencia_nombre(entries[i], zoom_tokens) for i in todos}
            grupo = [i for i in todos if evid[i] == "exacto"] or \
                    [i for i in todos if evid[i] == "parecido"]
            if grupo:
                mejor = max(score_of[i] for i in grupo)
                top = [i for i in grupo if score_of[i] == mejor]
                base = "" if evid[top[0]] == "exacto" else "nombre con ortografía parecida"
                if len(top) == 1:
                    return "ok", top[0], base, top
                idx, nota, resto = _desempatar(top, zoom_sede, zoom_tokens)
                if idx is not None:
                    return "ok", idx, "apellido repetido, " + nota, resto
                return "ambiguo", None, "", resto

            # Nadie con ese apellido coincide en el nombre. Antes de
            # conformarse con "solo apellido", se mira si hay un socio con el
            # MISMO nombre de pila y el apellido escrito algo distinto
            # ("Isaac Valdebenito" -> "Isaac ... Valdevenito"). Solo se acepta
            # si queda uno claro; si no, se sigue con la lógica de siempre.
            fz = _fuzzy_apellido(zoom_tokens, zoom_sede)
            if fz is not None and fz[0] == "ok":
                return fz

            # Nadie coincide en el nombre: solo el apellido. Se asigna
            # únicamente si ese apellido lo lleva UN solo socio y la sede
            # (si la escribió) no lo contradice; si no, queda pendiente.
            if len(todos) == 1:
                idx = todos[0]
                choca = bool(zoom_sede) and not _same_sede(zoom_sede, entries[idx]["Sede"])
                # Si en Zoom escribió un nombre de pila que SÍ existe en el
                # listado pero no es el de este socio, es otra persona.
                otro_nombre = any(
                    t in nombre_vocab and t not in entries[idx]["_all"] for t in zoom_tokens
                )
                if not choca and not otro_nombre:
                    return "ok", idx, "solo apellido (único en el listado)", todos
            return "ambiguo", None, "", todos

        # 2) Nadie coincide por apellido: solo nombre de pila
        nombre_cands = []
        for idx, entry in enumerate(entries):
            s = _score_by_nombre(entry, zoom_tokens)
            if s is not None:
                nombre_cands.append((s, idx))
        nombre_top: List[int] = []
        if nombre_cands:
            nombre_cands.sort(key=lambda c: c[0], reverse=True)
            nscore = nombre_cands[0][0]
            nombre_top = [idx for s, idx in nombre_cands if s == nscore]
            if len(nombre_top) == 1:
                extra = sorted(
                    t for t in zoom_tokens
                    if len(t) >= 4 and t not in vocab and t not in entries[nombre_top[0]]["_nombres"]
                )
                nota = "solo nombre (único en el listado)"
                if extra:
                    nota += "; revisar: en Zoom escribió además «" + " ".join(extra) + "», que no está en el listado"
                return "ok", nombre_top[0], nota, nombre_top

        # 2a) Apellido con otra ortografía o iniciales, SIEMPRE junto a un
        #     nombre que coincida exacto
        fz = _fuzzy_apellido(zoom_tokens, zoom_sede)
        if fz is not None:
            return fz

        # 2b) Empate de nombre de pila: se intenta desempatar por sede
        if nombre_top:
            idx, nota, resto = _desempatar(nombre_top, zoom_sede, zoom_tokens)
            if idx is not None:
                return "ok", idx, "solo nombre, " + nota, resto
            return "ambiguo", None, "", resto

        return "no", None, "", []

    no_encontrados: List[Dict[str, object]] = []
    ambiguos: List[Dict[str, object]] = []
    aproximados: List[Dict[str, object]] = []

    for zr in zoom_records:
        zoom_tokens = set(tokenize(str(zr.get("Nombre", ""))))
        for apodo, formal in _APODOS.items():
            if apodo in zoom_tokens:
                zoom_tokens |= formal
        zoom_sede = str(zr.get("Sede", "") or "")

        estado, idx, metodo, cands = _resolver(zoom_tokens, zoom_sede)

        if estado != "ok" or metodo:
            # 3) ¿Nombre escrito todo junto? ("mariateresacortes",
            #    "Francisco MarinNaritelli"). Se prueba también cuando el
            #    primer intento quedó ambiguo (ej. solo calzaba el nombre
            #    de pila): si separado resuelve a una sola persona, se usa.
            estado_prev, cands_prev = estado, cands
            prev = (estado, idx, metodo, cands)
            reemplazos = set()
            for tok in zoom_tokens:
                partes = _split_joined_token(tok, vocab)
                if partes:
                    reemplazos.add((tok, tuple(partes)))
            if reemplazos:
                nuevos = set(zoom_tokens)
                for tok, partes in reemplazos:
                    nuevos.discard(tok)
                    nuevos.update(partes)
                estado, idx, metodo, cands = _resolver(nuevos, zoom_sede)
                if estado == "ok":
                    metodo = "nombre escrito todo junto" + (", " + metodo if metodo else "")
                elif estado_prev == "ok":
                    # La separación no mejoró nada: se queda la asignación original.
                    estado, idx, metodo, cands = prev
                elif estado_prev == "ambiguo":
                    # Sigue sin resolverse: se conserva el resultado ambiguo
                    # original (o el nuevo si acota más los candidatos).
                    estado = "ambiguo"
                    if not cands or (cands_prev and len(cands_prev) < len(cands)):
                        cands = cands_prev

        if estado == "no":
            no_encontrados.append(zr)
            continue
        if estado == "ambiguo":
            ambiguos.append({**zr, "Posibles": [_label(i) for i in cands]})
            continue

        entries[idx]["_duracion_total"] += float(zr.get("Duración (minutos)", 0) or 0)
        entries[idx]["_matched_zoom_names"].append(str(zr.get("Nombre", "")))
        if metodo:
            e = entries[idx]
            aproximados.append({
                "Nombre": str(zr.get("Nombre", "")),
                "Sede en Zoom": zoom_sede,
                "Socio asignado": _label(idx).rsplit(" (", 1)[0],
                "Sede del socio": e["Sede"],
                "Cómo se asignó": metodo[0].upper() + metodo[1:],
                "Duración (minutos)": float(zr.get("Duración (minutos)", 0) or 0),
            })

    resultado: List[Dict[str, object]] = []
    for entry in entries:
        duracion = round(entry["_duracion_total"], 2)
        apellido = f"{entry['Apellido Paterno']} {entry['Apellido Materno']}".strip()
        apellido = re.sub(r"\s{2,}", " ", apellido)
        resultado.append({
            "Nombre": entry["Nombres"].strip(),
            "Apellido": apellido,
            "Sede": entry["Sede"].strip(),
            "Duración (minutos)": duracion,
            "Asistencia": compute_attendance_flag(duracion, attendance_threshold),
        })

    return resultado, no_encontrados, ambiguos, aproximados
