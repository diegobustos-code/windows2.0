# -*- coding: utf-8 -*-
"""
roster_fill_window.py
======================
Ventana secundaria "Rellenar Listado Existente": a diferencia del
Reporte Multi-Mes (que siempre arma un Excel NUEVO desde cero), esta
ventana sirve para el caso de un listado que el usuario ya viene
llenando a mano, mes a mes, agregando una columna nueva cada vez — con
"J" (inasistencia justificada) escritas en algunas celdas que hay que
respetar.

Se carga ese mismo archivo, se cargan uno o varios CSV de Zoom (uno por
mes, igual que en el Reporte Multi-Mes, con el mismo sistema de detectar
el mes por el nombre del archivo y pedir confirmación), y al generar se
guarda una copia ACTUALIZADA del listado: las columnas de mes que ya
existían se reutilizan (aunque su encabezado tenga texto extra, ej.
"Octubre (extraordinaria)"), las columnas de mes que todavía no existían
se agregan al final, las celdas con "J" nunca se sobrescriben, y el resto
del archivo (otras columnas, otras hojas) queda intacto.

Reutiliza sin cambios la misma lógica de lectura de CSV que usa el resto
de la aplicación (csv_processor.py) y la misma detección/confirmación de
mes que usa el Reporte Multi-Mes (multi_month_processor.py). Lo nuevo de
esta ventana es rellenar un archivo YA EXISTENTE en vez de generar uno
desde cero (roster_sheet_filler.py).
"""

import os

import tkinter as tk
from tkinter import ttk, filedialog, simpledialog, messagebox

from csv_processor import CSVProcessingError, load_zoom_csv
from multi_month_processor import detect_month_from_filename
from roster_sheet_filler import RosterSheetFillError, fill_roster_sheet, load_roster_sheet_for_fill
from ui_common import (
    COLOR_BG,
    COLOR_HEADER_BAR,
    COLOR_HEADER_TEXT,
    COLOR_BLUE,
    COLOR_GREEN,
    COLOR_TEAL,
    COLOR_ORANGE,
    COLOR_GRAY,
    make_flat_button,
    show_popup,
)


class RosterFillWindow(tk.Toplevel):
    """
    `main_app` es la ventana principal (ZoomAttendanceDesktopApp), usada
    solo para reutilizar su misma carpeta de exportación
    (Downloads / Asistencia Zoom / Cotejados).
    """

    def __init__(self, main_app):
        super().__init__(main_app)
        self.main_app = main_app
        self.title("Rellenar Listado Existente")
        self.geometry("880x640")
        self.minsize(740, 540)
        self.configure(bg=COLOR_BG)
        self.transient(main_app)

        self.sheet_info = None  # RosterSheetInfo, se llena al cargar el listado
        # Cada elemento: {"label": str, "records": list, "filename": str}
        self.month_entries = []

        self._build_ui()
        self._refresh_sheet_label()
        self._refresh_months_table()

        self.update_idletasks()
        x = main_app.winfo_rootx() + max((main_app.winfo_width() - self.winfo_width()) // 2, 0)
        y = main_app.winfo_rooty() + max((main_app.winfo_height() - self.winfo_height()) // 2, 0)
        self.geometry(f"+{x}+{y}")

    # ------------------------------------------------------------------
    def _build_ui(self):
        main = tk.Frame(self, bg=COLOR_BG, padx=14, pady=12)
        main.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            main, bg=COLOR_BG, fg="#2E3440", justify="left", anchor="w",
            font=("Segoe UI", 10), wraplength=840,
            text=(
                "Carga el listado de socios que ya vienes llenando a mano (con una "
                "columna por mes) y uno o varios CSV de Zoom (uno por mes). Las "
                "columnas de mes que ya existan se reutilizan tal cual, las que "
                "falten se agregan al final, y las celdas marcadas \"J\" (inasistencia "
                "justificada) NUNCA se sobrescriben. El resultado se guarda como un "
                "archivo NUEVO, sin tocar el original."
            ),
        ).pack(fill=tk.X, pady=(0, 10))

        # --- Listado a rellenar ---
        sheet_frame = tk.Frame(main, bg=COLOR_HEADER_BAR, padx=10, pady=8)
        sheet_frame.pack(fill=tk.X, pady=(0, 10))
        self.sheet_label_var = tk.StringVar()
        tk.Label(
            sheet_frame, textvariable=self.sheet_label_var, bg=COLOR_HEADER_BAR,
            fg=COLOR_HEADER_TEXT, font=("Segoe UI", 10, "bold"), anchor="w",
            justify="left", wraplength=620,
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)
        make_flat_button(
            sheet_frame, "Cargar listado a rellenar (Excel)", self._load_sheet, COLOR_BLUE
        ).pack(side=tk.RIGHT)

        # --- Botones para agregar/editar/quitar meses ---
        add_bar = tk.Frame(main, bg=COLOR_BG)
        add_bar.pack(fill=tk.X, pady=(0, 8))
        make_flat_button(
            add_bar, "Agregar archivo(s) CSV de Zoom", self._add_files, COLOR_GREEN
        ).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 4))
        make_flat_button(
            add_bar, "Editar mes seleccionado", self._edit_selected_month, COLOR_ORANGE
        ).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=4)
        make_flat_button(
            add_bar, "Quitar seleccionado", self._remove_selected, COLOR_GRAY
        ).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(4, 0))

        # --- Tabla de meses cargados ---
        table_frame = tk.Frame(main, bg=COLOR_BG)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        columns = ("orden", "mes", "archivo", "participantes")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
        headings = {
            "orden": ("#", 30),
            "mes": ("Mes (columna a rellenar)", 220),
            "archivo": ("Archivo", 340),
            "participantes": ("Participantes", 100),
        }
        for col, (label, width) in headings.items():
            self.tree.heading(col, text=label)
            self.tree.column(
                col, width=width,
                anchor=("center" if col in ("orden", "participantes") else "w"),
                stretch=(col == "archivo"),
            )
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind("<Double-1>", lambda e: self._edit_selected_month())

        tk.Label(
            main, bg=COLOR_BG, fg="#7A828B", font=("Segoe UI", 8), anchor="w",
            text=(
                "Para cada mes: si la columna ya existe en el listado (por nombre, aunque "
                "tenga texto extra como \"Octubre (extraordinaria)\"), se reutiliza; si no, "
                "se crea al final. Doble clic sobre una fila también permite editar el mes."
            ),
        ).pack(fill=tk.X, pady=(0, 10))

        # --- Umbral de "Presente" ---
        threshold_bar = tk.Frame(main, bg=COLOR_BG)
        threshold_bar.pack(fill=tk.X, pady=(0, 12))
        tk.Label(
            threshold_bar, text="Umbral de 'Presente' (minutos):", bg=COLOR_BG,
            fg="#2E3440", font=("Segoe UI", 10),
        ).pack(side=tk.LEFT)
        default_threshold = "30"
        try:
            default_threshold = self.main_app.threshold_var.get() or "30"
        except Exception:
            pass
        self.threshold_var = tk.StringVar(value=default_threshold)
        tk.Entry(
            threshold_bar, textvariable=self.threshold_var, width=8, font=("Segoe UI", 10)
        ).pack(side=tk.LEFT, padx=(8, 0))
        tk.Label(
            threshold_bar,
            text="(se aplica igual en todos los meses que se rellenen ahora)",
            bg=COLOR_BG, fg="#7A828B", font=("Segoe UI", 8),
        ).pack(side=tk.LEFT, padx=(8, 0))

        # --- Qué hacer con las J si el Zoom dice que sí estuvo ---
        self.cambiar_j_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            main, variable=self.cambiar_j_var, bg=COLOR_BG, activebackground=COLOR_BG,
            font=("Segoe UI", 10), anchor="w", justify="left",
            text="Cambiar J a P si el Zoom dice que estuvo presente (se anota en \"Revisar diferencias\")",
        ).pack(fill=tk.X, pady=(0, 12))

        # --- Generar ---
        make_flat_button(
            main, "Rellenar y guardar copia actualizada", self._generate, COLOR_TEAL
        ).pack(fill=tk.X)

    # ------------------------------------------------------------------
    def _refresh_sheet_label(self):
        if self.sheet_info is not None:
            meses_existentes = ", ".join(sorted(
                h for _col, h in self.sheet_info.month_columns.values()
            )) or "(ninguno todavía)"
            self.sheet_label_var.set(
                f"Listado cargado: {len(self.sheet_info.roster_records)} socios "
                f"(hoja \"{self.sheet_info.sheet_name}\"). Meses ya presentes: {meses_existentes}"
            )
        else:
            self.sheet_label_var.set("Listado a rellenar: (sin cargar todavía)")

    def _refresh_months_table(self):
        self.tree.delete(*self.tree.get_children())
        for i, entry in enumerate(self.month_entries, start=1):
            self.tree.insert(
                "", tk.END,
                values=(i, entry["label"], entry["filename"], len(entry["records"])),
            )

    # ------------------------------------------------------------------
    # Listado a rellenar
    # ------------------------------------------------------------------
    def _load_sheet(self):
        filepath = filedialog.askopenfilename(
            title="Abrir el listado de socios a rellenar (Excel)",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")],
        )
        if not filepath:
            return
        try:
            self.sheet_info = load_roster_sheet_for_fill(filepath)
        except RosterSheetFillError as e:
            show_popup(self, "Error al procesar el listado", str(e), success=False)
            return
        except Exception as e:
            show_popup(self, "Error inesperado", str(e), success=False)
            return
        self._refresh_sheet_label()

    # ------------------------------------------------------------------
    # Archivos de Zoom (uno o varios a la vez) — idéntico al Reporte
    # Multi-Mes: detecta el mes por el nombre del archivo y SIEMPRE pide
    # confirmación antes de darlo por bueno.
    # ------------------------------------------------------------------
    def _add_files(self):
        filepaths = filedialog.askopenfilenames(
            title="Abrir uno o varios CSV de Zoom (uno por mes)",
            filetypes=[("CSV o Excel", "*.csv *.xlsx *.xls"), ("Todos los archivos", "*.*")],
        )
        if not filepaths:
            return

        omitidos = 0
        for filepath in filepaths:
            try:
                records, _column_map = load_zoom_csv(filepath)
            except CSVProcessingError as e:
                show_popup(
                    self, f"Error al procesar {os.path.basename(filepath)}", str(e), success=False
                )
                continue
            except Exception as e:
                show_popup(
                    self, "Error inesperado", f"{os.path.basename(filepath)}: {e}", success=False
                )
                continue

            filename = os.path.basename(filepath)
            guessed = detect_month_from_filename(filename)

            if guessed:
                confirmado = messagebox.askyesno(
                    "Confirmar mes detectado",
                    f"Archivo:\n{filename}\n\n"
                    f"Se detectó que corresponde al mes: \"{guessed}\".\n\n"
                    f"¿Es correcto?",
                    parent=self,
                )
                if confirmado:
                    label = guessed
                else:
                    label = simpledialog.askstring(
                        "Mes de este archivo",
                        f"Ingresa el mes correcto para:\n{filename}\n(ej. \"Julio 2026\"):",
                        initialvalue=guessed,
                        parent=self,
                    )
                    if not label or not label.strip():
                        omitidos += 1
                        continue
                    label = label.strip()
            else:
                label = simpledialog.askstring(
                    "Mes de este archivo",
                    "No se pudo adivinar el mes a partir del nombre del "
                    f"archivo:\n{filename}\n\nIngresa el mes (ej. \"Julio 2026\"):",
                    parent=self,
                )
                if not label or not label.strip():
                    omitidos += 1
                    continue
                label = label.strip()

            label = self._unique_label(label)
            self.month_entries.append({"label": label, "records": records, "filename": filename})

        self._refresh_months_table()
        if omitidos:
            show_popup(
                self, "Archivo(s) omitido(s)",
                f"Se omitieron {omitidos} archivo(s) porque no se indicó un mes para ellos.",
            )

    def _unique_label(self, label: str) -> str:
        existing = {e["label"] for e in self.month_entries}
        if label not in existing:
            return label
        n = 2
        while f"{label} ({n})" in existing:
            n += 1
        return f"{label} ({n})"

    def _selected_index(self):
        sel = self.tree.selection()
        if not sel:
            return None
        return self.tree.index(sel[0])

    def _edit_selected_month(self):
        idx = self._selected_index()
        if idx is None:
            show_popup(self, "Nada seleccionado", "Selecciona primero una fila de la tabla.")
            return
        entry = self.month_entries[idx]
        new_label = simpledialog.askstring(
            "Editar mes", "Nombre del mes para esta columna:",
            initialvalue=entry["label"], parent=self,
        )
        if new_label and new_label.strip():
            new_label = new_label.strip()
            others = {e["label"] for i, e in enumerate(self.month_entries) if i != idx}
            if new_label in others:
                show_popup(
                    self, "Mes repetido",
                    "Ya hay otro archivo cargado con ese mismo mes.", success=False,
                )
                return
            entry["label"] = new_label
            self._refresh_months_table()

    def _remove_selected(self):
        idx = self._selected_index()
        if idx is None:
            show_popup(self, "Nada seleccionado", "Selecciona primero una fila de la tabla.")
            return
        del self.month_entries[idx]
        self._refresh_months_table()

    # ------------------------------------------------------------------
    # Rellenar y guardar
    # ------------------------------------------------------------------
    def _generate(self):
        if self.sheet_info is None:
            show_popup(
                self, "Falta el listado a rellenar",
                "Carga primero el listado de socios que quieres rellenar (Excel).",
            )
            return
        if not self.month_entries:
            show_popup(self, "Sin meses cargados", "Agrega al menos un archivo CSV de Zoom.")
            return

        try:
            threshold = float(str(self.threshold_var.get()).strip().replace(",", "."))
        except ValueError:
            threshold = 30.0

        months = [(e["label"], e["records"]) for e in self.month_entries]
        try:
            result = fill_roster_sheet(
                self.sheet_info, months, threshold,
                cambiar_j_si_presente=self.cambiar_j_var.get(),
            )
        except Exception as e:
            show_popup(self, "Error al rellenar el listado", str(e), success=False)
            return

        base_dir = self.main_app._export_dir("Cotejados")
        slug_first = self.main_app._slugify(months[0][0])
        slug_last = self.main_app._slugify(months[-1][0])
        filename = (
            f"listado_actualizado_{slug_first}.xlsx" if len(months) == 1
            else f"listado_actualizado_{slug_first}_a_{slug_last}.xlsx"
        )
        filepath = base_dir / filename

        try:
            self.sheet_info.workbook.save(str(filepath))
        except Exception as e:
            show_popup(self, "Error al guardar el archivo", str(e), success=False)
            return

        total_no_encontrados = sum(len(v) for v in result.no_encontrados_por_mes.values())
        total_ambiguos = sum(len(v) for v in result.ambiguos_por_mes.values())

        resumen = []
        if result.columns_created:
            resumen.append("Columnas nuevas creadas: " + ", ".join(result.columns_created))
        if result.columns_updated:
            resumen.append("Columnas existentes reutilizadas: " + ", ".join(result.columns_updated))
        resumen.append(
            f"{result.celdas_escritas} celdas rellenadas, "
            f"{result.celdas_respetadas_j} respetadas por tener \"J\"."
        )
        if result.celdas_j_cambiadas:
            resumen.append(
                f"{result.celdas_j_cambiadas} J cambiadas a P porque el Zoom las marca presentes."
            )
        if total_no_encontrados or total_ambiguos:
            resumen.append(
                f"Pendientes: {total_no_encontrados} no encontrado(s) y {total_ambiguos} "
                f"ambiguo(s), listados mes por mes en las pestañas \"No Encontrados\" y "
                f"\"Ambiguos\" del mismo archivo."
            )

        total_aprox = sum(len(v) for v in result.aproximados_por_mes.values())
        if total_aprox:
            resumen.append(
                f"{total_aprox} cotejo(s) asignados por aproximación (solo nombre único, "
                f"apellido con otra ortografía, desempate por sede...): pestaña "
                f"\"Cotejos aproximados\"."
            )
        if result.diferencias:
            resumen.append(
                f"{len(result.diferencias)} caso(s) donde el Zoom no coincide con lo que ya "
                f"figuraba en el listado (J / A / P): pestaña \"Revisar diferencias\"."
            )

        show_popup(
            self, "Listado actualizado correctamente",
            f"Se guardó una COPIA actualizada (el archivo original no se modificó) en:\n"
            f"Downloads / Asistencia Zoom / Cotejados\n\ncon el nombre:\n{filename}\n\n"
            + "\n".join(resumen),
            success=True, open_path=str(filepath), is_file=True,
        )
