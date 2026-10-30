# -*- coding: utf-8 -*-
"""
roster_sheet_filler.py
=======================
Lógica para "Rellenar listado existente": a diferencia del Reporte
Multi-Mes (que SIEMPRE genera un Excel nuevo desde cero), esta función
toma un listado de socios que el usuario YA viene llenando a mano mes a
mes —con columnas que se van agregando a medida que pasan los meses, y
con "J" (inasistencia justificada) escritas a mano en algunas celdas— y
lo actualiza IN-MEMORY con los resultados del cotejo de uno o varios
meses de Zoom, respetando todo lo que ya estaba en el archivo:

    - Si la celda de un socio para el mes que se está procesando ya dice
      "J", se deja TAL CUAL (nunca se sobrescribe una justificación).
    - Si el mes que se está cargando ya tiene una columna en la hoja
      (aunque el encabezado tenga texto extra, ej. "Octubre
      (extraordinaria)" o "Asistencia Agosto"), se usa esa columna.
    - Si el mes todavía no tiene columna, se agrega una columna nueva al
      final, con el nombre del mes como encabezado.
    - La columna "Sede" (y cualquier otra columna que no sea Nombre/
      Apellido/un mes reconocible, como "Licencias" o una columna vacía
      tipo "SEDES") se deja intacta, sin tocarla para nada.
    - Los participantes de Zoom que no se pudieron cotejar en cada mes
      (no encontrados / ambiguos) NO se pierden: quedan en dos pestañas
      nuevas del mismo archivo ("No Encontrados" y "Ambiguos"), igual que
      en el Reporte Multi-Mes, para revisión manual.

Reutiliza sin cambios la detección de columnas (roster_matcher) y el
cotejo (roster_matcher.match_zoom_to_roster_detailed) que usa el resto de la
aplicación — lo único nuevo aquí es ESCRIBIR el resultado dentro de la
hoja existente en vez de generar una tabla desde cero.
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.table import AutoFilter
from openpyxl.worksheet.worksheet import Worksheet

from copy import copy

from excel_exporter import (
    SHEET_AMBIGUOS, SHEET_APROXIMADOS, SHEET_NO_ENCONTRADOS,
    _write_aproximados_sheet, _write_review_sheet, _write_table_sheet, apply_status_style,
)
from multi_month_processor import extract_month_name_only
from roster_matcher import _detect_full_format, _detect_simple_format, match_zoom_to_roster_detailed


class RosterSheetFillError(Exception):
    """Error controlado y con mensaje claro para mostrar al usuario final."""
    pass


@dataclass
class RosterSheetInfo:
    """Todo lo que se necesita para rellenar un listado de socios ya
    existente: el libro abierto (para guardarlo al final), la hoja
    encontrada, dónde está cada columna fija, y un mapa de qué mes ya
    tiene columna y cuál sería la próxima columna libre para agregar uno
    nuevo."""
    workbook: Workbook
    sheet: Worksheet
    sheet_name: str
    row_numbers: List[int]
    roster_records: List[Dict[str, str]]
    # mes canónico (ej. "Octubre") -> (índice de columna, encabezado real tal
    # cual está escrito en el archivo, ej. "Octubre (extraordinaria)")
    month_columns: Dict[str, Tuple[int, str]] = field(default_factory=dict)
    next_free_col: int = 1
    # Columna (1-based) de la que se toma el relleno "original" de cada fila
    # (el color de su sede) para conservarlo en las celdas P.
    base_col: int = 1
    # Columnas (1-based) que NO se reformatean: las fijas (Nombre/Apellido/
    # Sede) y cualquier columna "Sede"/"Sedes" suelta.
    skip_cols: set = field(default_factory=set)


@dataclass
class FillResult:
    columns_created: List[str]
    columns_updated: List[str]
    no_encontrados_por_mes: Dict[str, List[Dict[str, object]]]
    ambiguos_por_mes: Dict[str, List[Dict[str, object]]]
    celdas_escritas: int
    celdas_respetadas_j: int
    aproximados_por_mes: Dict[str, List[Dict[str, object]]] = field(default_factory=dict)
    # Casos donde lo que ya figuraba en el listado (J, A o P) no coincide con
    # lo que dice el Zoom; se listan en la pestaña "Revisar diferencias".
    diferencias: List[Dict[str, object]] = field(default_factory=list)


SHEET_DIFERENCIAS = "Revisar diferencias"


def _repair_sheet_tables(ws: Worksheet, last_col: int, last_row: int) -> None:
    """
    Repara un problema heredado de algunos listados (típicamente los
    exportados desde Google Sheets a .xlsx): la hoja trae una Tabla de
    Excel (con sus filtros y su formato de franjas de color) cuyo rango
    NO coincide con el autoFiltro "suelto" a nivel de hoja, y a la Tabla
    le falta su propio autoFiltro interno. Ese desajuste es justamente lo
    que hace que, al abrir el archivo en Excel, aparezca el aviso de
    "hemos encontrado un problema con el contenido" y Excel termine
    quitando la función de Tabla al reparar el archivo.

    Esta función deja cada Tabla de la hoja con un rango consistente
    (expandido si hace falta para cubrir alguna columna de mes nueva que
    se haya agregado) y su propio autoFiltro interno sincronizado, y
    quita el autoFiltro suelto de la hoja para que no quede duplicado ni
    desalineado con la Tabla.
    """
    if not ws.tables:
        return

    for table in ws.tables.values():
        min_col, min_row, max_col, max_row = range_boundaries(table.ref)
        new_max_col = max(max_col, last_col)
        new_max_row = max(max_row, last_row)
        if new_max_col != max_col or new_max_row != max_row:
            start = f"{get_column_letter(min_col)}{min_row}"
            end = f"{get_column_letter(new_max_col)}{new_max_row}"
            table.ref = f"{start}:{end}"
        if table.autoFilter is None or table.autoFilter.ref != table.ref:
            table.autoFilter = AutoFilter(ref=table.ref)

    # El filtro ya queda definido dentro de cada Tabla; un autoFiltro de
    # hoja aparte (que además en estos archivos suele apuntar a un rango
    # viejo y más chico) sobra y es la causa del archivo "dañado".
    ws.auto_filter.ref = None


def load_roster_sheet_for_fill(filepath: str) -> RosterSheetInfo:
    """
    Abre un listado de socios que el usuario ya viene llenando a mano
    (columnas Nombre/Apellido/Sede + una columna por mes ya cotejado, más
    posibles columnas sueltas como "Licencias"), y prepara toda la
    información necesaria para poder rellenarlo.

    Si el archivo tiene varias hojas (algo común: aparte de la hoja con
    el listado puede haber otra hoja de resumen/estadísticas sin
    relación), se usa la PRIMERA hoja cuyas columnas calcen con alguno de
    los dos formatos de listado de socios que reconoce el resto de la
    aplicación (Sede + Apellido Paterno + Apellido Materno + Nombres, o
    Sede + Apellido + una columna de nombre) — las demás hojas se dejan
    intactas.

    Lanza RosterSheetFillError con un mensaje claro si ninguna hoja tiene
    un formato reconocible, o si el archivo está vacío.
    """
    try:
        wb_values = load_workbook(filepath, data_only=True)
    except Exception as e:
        raise RosterSheetFillError(
            f"No se pudo abrir el archivo Excel del listado de socios.\n"
            f"Detalle técnico: {e}"
        )

    chosen_sheet_name = None
    chosen_cols = None  # {"Sede": idx, "Apellido Paterno": idx, "Apellido Materno": idx|None, "Nombres": idx}

    for sheet_name in wb_values.sheetnames:
        ws = wb_values[sheet_name]
        first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if not first_row:
            continue
        headers = [str(h).strip() if h is not None else "" for h in first_row]

        full_cols = _detect_full_format(headers)
        if full_cols is not None:
            chosen_sheet_name = sheet_name
            chosen_cols = {
                "Sede": full_cols["Sede"],
                "Apellido Paterno": full_cols["Apellido Paterno"],
                "Apellido Materno": full_cols["Apellido Materno"],
                "Nombres": full_cols["Nombres"],
            }
            break

        simple_cols = _detect_simple_format(headers)
        if simple_cols is not None:
            chosen_sheet_name = sheet_name
            chosen_cols = {
                "Sede": simple_cols["Sede"],
                "Apellido Paterno": simple_cols["Apellido"],
                "Apellido Materno": None,
                "Nombres": simple_cols["Nombres"],
            }
            break

    if chosen_sheet_name is None:
        raise RosterSheetFillError(
            "No se encontró, en ninguna hoja del archivo, un formato de "
            "listado de socios reconocible.\n\n"
            "Se acepta cualquiera de estos dos formatos de columnas:\n"
            "  • Sede, Apellido Paterno, Apellido Materno, Nombres\n"
            "  • Sede, Apellido (el apellido completo en una sola columna), "
            "y una columna con el nombre (con cualquier título)"
        )

    ws_values = wb_values[chosen_sheet_name]
    first_row = next(ws_values.iter_rows(min_row=1, max_row=1, values_only=True), None)
    headers = [str(h).strip() if h is not None else "" for h in (first_row or [])]

    fixed_cols = {c for c in chosen_cols.values() if c is not None}  # índices 0-based

    # --- Detectar columnas de mes ya existentes, entre las columnas que
    # sobran (ni Sede, ni Apellido[s], ni Nombre) ---
    month_columns: Dict[str, Tuple[int, str]] = {}
    last_header_col_0based = -1
    for idx, h in enumerate(headers):
        if h:
            last_header_col_0based = idx
        if idx in fixed_cols or not h:
            continue
        mes = extract_month_name_only(h)
        if mes and mes not in month_columns:
            month_columns[mes] = (idx + 1, h)  # +1: openpyxl usa columnas 1-based

    next_free_col = last_header_col_0based + 2  # +1 para pasar a 1-based, +1 para la siguiente libre

    skip_cols = {c + 1 for c in fixed_cols}
    for idx, h in enumerate(headers):
        if h.strip().lower() in ("sede", "sedes"):
            skip_cols.add(idx + 1)
    base_col = (chosen_cols["Nombres"] if chosen_cols["Nombres"] is not None
                else chosen_cols["Apellido Paterno"]) + 1

    # --- Leer las filas de socios (Nombre/Apellido/Sede) y su número de
    # fila REAL en la hoja (necesario para escribir después en la celda
    # correcta de cada socio, incluso si hay filas vacías de por medio) ---
    row_numbers: List[int] = []
    roster_records: List[Dict[str, str]] = []
    for row_idx, row in enumerate(ws_values.iter_rows(min_row=2, values_only=True), start=2):
        if row is None or not any(row):
            continue

        def _cell(col_0based: Optional[int]) -> str:
            if col_0based is None or col_0based >= len(row) or row[col_0based] is None:
                return ""
            return str(row[col_0based]).strip()

        apellido_paterno = _cell(chosen_cols["Apellido Paterno"])
        apellido_materno = _cell(chosen_cols["Apellido Materno"])
        nombres = _cell(chosen_cols["Nombres"])
        if not apellido_paterno and not apellido_materno and not nombres:
            continue

        row_numbers.append(row_idx)
        roster_records.append({
            "Sede": _cell(chosen_cols["Sede"]),
            "Apellido Paterno": apellido_paterno,
            "Apellido Materno": apellido_materno,
            "Nombres": nombres,
        })

    if not roster_records:
        raise RosterSheetFillError(
            "El listado de socios se leyó correctamente, pero no se "
            "encontraron filas de socios válidas en la hoja "
            f"\"{chosen_sheet_name}\"."
        )

    # --- Se abre el archivo una SEGUNDA vez, ahora SIN data_only, para
    # escribir en él más abajo: si el archivo tuviera alguna fórmula en
    # otra hoja, abrir con data_only=True y volver a guardar ESE mismo
    # libro la reemplazaría por su valor fijo — abriendo aparte para
    # escritura se evita perder fórmulas que no tengan que ver con esto. ---
    try:
        wb_write = load_workbook(filepath, data_only=False)
    except Exception as e:
        raise RosterSheetFillError(
            f"No se pudo volver a abrir el archivo para escribir los "
            f"cambios.\nDetalle técnico: {e}"
        )
    ws_write = wb_write[chosen_sheet_name]

    return RosterSheetInfo(
        workbook=wb_write,
        sheet=ws_write,
        sheet_name=chosen_sheet_name,
        row_numbers=row_numbers,
        roster_records=roster_records,
        month_columns=month_columns,
        next_free_col=next_free_col,
        base_col=base_col,
        skip_cols=skip_cols,
    )


def fill_roster_sheet(
    sheet_info: RosterSheetInfo,
    months: List[Tuple[str, List[Dict[str, object]]]],
    threshold_minutes: float = 30.0,
) -> FillResult:
    """
    Rellena, en la hoja ya cargada por load_roster_sheet_for_fill, la
    columna de cada mes recibido en `months` (lista de tuplas
    (etiqueta_del_mes, registros_zoom_del_mes), en cualquier orden).

    Para cada mes:
        - Si ya existe una columna para ese mes (por nombre, ignorando
          texto extra como "(extraordinaria)"), se reutiliza esa columna
          TAL CUAL está (no se toca su encabezado).
        - Si no existe, se agrega una columna nueva al final de la hoja.
        - Para cada socio: si la celda de esa columna ya dice "J", se
          deja intacta; si no, se escribe "P" o "A" según el cotejo
          (con el mismo color semáforo que usa el resto de la app).

    Al final, agrega (o reemplaza, si ya existían de una corrida
    anterior) cuatro pestañas en el MISMO archivo, para revisión manual:
    "No Encontrados", "Ambiguos", "Cotejos aproximados" y "Revisar
    diferencias" (esta última lista a quienes el Zoom marca Presentes pero
    el listado tenía como J o A, y a quienes pasaron de P a A).

    Devuelve un FillResult con el resumen de lo que se hizo, para
    mostrárselo al usuario.
    """
    columns_created: List[str] = []
    columns_updated: List[str] = []
    no_encontrados_por_mes: Dict[str, List[Dict[str, object]]] = {}
    ambiguos_por_mes: Dict[str, List[Dict[str, object]]] = {}
    aproximados_por_mes: Dict[str, List[Dict[str, object]]] = {}
    diferencias: List[Dict[str, object]] = []
    celdas_escritas = 0
    celdas_respetadas_j = 0

    month_labels_order: List[str] = []

    for month_label, zoom_records in months:
        month_labels_order.append(month_label)
        bare_month = extract_month_name_only(month_label) or month_label

        existing = sheet_info.month_columns.get(bare_month)
        if existing is not None:
            col, header_text = existing
            columns_updated.append(header_text)
        else:
            col = sheet_info.next_free_col
            header_cell = sheet_info.sheet.cell(row=1, column=col, value=bare_month)
            header_cell.font = Font(bold=True)
            sheet_info.month_columns[bare_month] = (col, bare_month)
            sheet_info.next_free_col += 1
            columns_created.append(bare_month)

        resultado, no_encontrados, ambiguos, aproximados = match_zoom_to_roster_detailed(
            zoom_records, sheet_info.roster_records, threshold_minutes
        )
        no_encontrados_por_mes[month_label] = no_encontrados
        ambiguos_por_mes[month_label] = ambiguos
        aproximados_por_mes[month_label] = aproximados

        # resultado[i] corresponde exactamente a sheet_info.roster_records[i]
        # (match_zoom_to_roster conserva el orden), que a su vez corresponde
        # a sheet_info.row_numbers[i] — por eso se puede recorrer en paralelo.
        for idx, row_number in enumerate(sheet_info.row_numbers):
            cell = sheet_info.sheet.cell(row=row_number, column=col)
            current = str(cell.value).strip().upper() if cell.value is not None else ""
            nuevo = resultado[idx]["Asistencia"]  # "P" o "A"

            # Diferencias con lo que ya figuraba en el listado, para revisar:
            #   J y Zoom dice P (la J se respeta), A y Zoom dice P, P y Zoom dice A.
            if (current, nuevo) in (("J", "P"), ("A", "P"), ("P", "A")):
                reg = sheet_info.roster_records[idx]
                diferencias.append({
                    "Mes": month_label,
                    "Nombre": reg["Nombres"].strip(),
                    "Apellido": f"{reg['Apellido Paterno']} {reg['Apellido Materno']}".strip(),
                    "Sede": reg["Sede"].strip(),
                    "Figuraba en el listado": current,
                    "Según Zoom": nuevo,
                    "Minutos en Zoom": resultado[idx]["Duración (minutos)"],
                    "Qué hizo el programa": (
                        "Se respetó la J (no se modificó la celda)" if current == "J"
                        else f"Se cambió de {current} a {nuevo}"
                    ),
                })

            if current == "J":
                celdas_respetadas_j += 1
                continue  # Inasistencia justificada: nunca se sobrescribe.

            cell.value = nuevo  # el formato se aplica al final
            celdas_escritas += 1

    # --- Formato de asistencia: se aplica a TODAS las celdas P/A/J de los
    # socios (las recién escritas y las que ya estaban de meses anteriores,
    # aunque vinieran con el formato viejo). Las P conservan el relleno
    # original de su fila (el de la columna base); las A y J llevan el suyo.
    # Se omiten las columnas fijas y las columnas "Sede"/"Sedes". ---
    last_col = sheet_info.next_free_col - 1
    for row_number in sheet_info.row_numbers:
        base_fill = sheet_info.sheet.cell(row=row_number, column=sheet_info.base_col).fill
        for col in range(1, last_col + 1):
            if col in sheet_info.skip_cols:
                continue
            apply_status_style(sheet_info.sheet.cell(row=row_number, column=col), base_fill)

    # --- Deja cualquier Tabla de Excel presente en la hoja (franjas de
    # color, filtros) con un rango y un autoFiltro consistentes, cubriendo
    # también las columnas de mes nuevas que se hayan agregado recién.
    # Evita el aviso de "archivo dañado" que aparece en algunos listados
    # (típicamente los exportados desde Google Sheets). ---
    last_row_used = max(sheet_info.row_numbers) if sheet_info.row_numbers else 1
    _repair_sheet_tables(sheet_info.sheet, sheet_info.next_free_col - 1, last_row_used)

    # --- Pestañas de revisión: se eliminan primero si ya existían (de una
    # corrida anterior), para que la actualización sea idempotente y no
    # vaya acumulando pestañas duplicadas cada vez que se rellena. ---
    for title in (SHEET_NO_ENCONTRADOS, SHEET_AMBIGUOS, SHEET_APROXIMADOS, SHEET_DIFERENCIAS):
        if title in sheet_info.workbook.sheetnames:
            sheet_info.workbook.remove(sheet_info.workbook[title])

    _write_review_sheet(
        sheet_info.workbook, SHEET_NO_ENCONTRADOS, month_labels_order, no_encontrados_por_mes,
        empty_message="No hubo participantes sin encontrar en ningún mes.",
        table_name="TablaNoEncontradosRelleno",
    )
    _write_review_sheet(
        sheet_info.workbook, SHEET_AMBIGUOS, month_labels_order, ambiguos_por_mes,
        empty_message="No hubo participantes ambiguos en ningún mes.",
        table_name="TablaAmbiguosRelleno", con_posibles=True,
    )
    _write_aproximados_sheet(
        sheet_info.workbook, month_labels_order, aproximados_por_mes,
        table_name="TablaAproximadosRelleno",
    )
    columnas_dif = ["Mes", "Nombre", "Apellido", "Sede", "Figuraba en el listado",
                    "Según Zoom", "Minutos en Zoom", "Qué hizo el programa"]
    _write_table_sheet(
        sheet_info.workbook, SHEET_DIFERENCIAS, columnas_dif,
        [[d[c] for c in columnas_dif] for d in diferencias],
        "No hubo diferencias entre el listado y los reportes de Zoom.",
        "TablaDiferenciasRelleno",
    )

    return FillResult(
        columns_created=columns_created,
        columns_updated=columns_updated,
        no_encontrados_por_mes=no_encontrados_por_mes,
        ambiguos_por_mes=ambiguos_por_mes,
        celdas_escritas=celdas_escritas,
        celdas_respetadas_j=celdas_respetadas_j,
        aproximados_por_mes=aproximados_por_mes,
        diferencias=diferencias,
    )
