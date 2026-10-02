# -*- coding: utf-8 -*-
"""
multi_month_window.py
======================
Ventana secundaria "Reporte Multi-Mes": permite cargar varios archivos
CSV de Zoom (uno por mes, o varios/todos los meses de un año a la vez) y
cotejarlos contra el MISMO listado oficial de socios, produciendo un solo
archivo Excel con una fila por socio y una columna de Asistencia (P/A)
por cada mes cargado.

Reutiliza sin cambios la misma lógica de lectura de CSV, cotejo y
exportación que usa el resto de la aplicación (csv_processor.py,
roster_matcher.py, excel_exporter.py). Lo único nuevo es la combinación
de varios meses en una sola tabla (multi_month_processor.py) y esta
ventana para armar la lista de meses antes de generar el reporte.
"""

import os

import tkinter as tk
from tkinter import ttk, filedialog, simpledialog, messagebox

from csv_processor import CSVProcessingError, load_zoom_csv
from roster_matcher import RosterProcessingError, load_roster_excel
from excel_exporter import export_multi_month_to_excel
from multi_month_processor import build_multi_month_matrix, detect_month_from_filename
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


class MultiMonthWindow(tk.Toplevel):
    """
    `main_app` es la ventana principal (ZoomAttendanceDesktopApp). Se usa
    para reutilizar su listado de socios si ya estaba cargado y su misma
    carpeta de exportación (Downloads / Asistencia Zoom / Cotejados).
    """

    def __init__(self, main_app):
        super().__init__(main_app)
        self.main_app = main_app
        self.title("Reporte Multi-Mes")
        self.geometry("860x620")
        self.minsize(720, 520)
        self.configure(bg=COLOR_BG)
        self.transient(main_app)

        # Cada elemento: {"label": str, "records": list, "filename": str}
        self.month_entries = []
        # Si la ventana principal ya tenía un listado de socios cargado,
        # se reutiliza como punto de partida (se puede cambiar igual).
        self.roster_records = list(getattr(main_app, "roster_records", []) or [])

        self._build_ui()
        self._refresh_roster_label()
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
            font=("Segoe UI", 10), wraplength=820,
            text=(
                "Carga uno o varios CSV de Zoom (un archivo por mes — puedes "
                "seleccionar varios a la vez, por ejemplo todos los meses de "
                "un año) y el listado oficial de socios. Al generar el "
                "reporte, cada mes se cotejará por separado contra el "
                "listado y se armará un solo Excel con una columna P/A por "
                "cada mes."
            ),
        ).pack(fill=tk.X, pady=(0, 10))

        # --- Listado de socios ---
        roster_frame = tk.Frame(main, bg=COLOR_HEADER_BAR, padx=10, pady=8)
        roster_frame.pack(fill=tk.X, pady=(0, 10))
        self.roster_label_var = tk.StringVar()
        tk.Label(
            roster_frame, textvariable=self.roster_label_var, bg=COLOR_HEADER_BAR,
            fg=COLOR_HEADER_TEXT, font=("Segoe UI", 10, "bold"), anchor="w",
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)
        make_flat_button(
            roster_frame, "Cargar listado de socios (Excel)", self._load_roster, COLOR_BLUE
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
            "mes": ("Mes (columna en el Excel)", 220),
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
                "El orden de la lista define el orden de las columnas en el Excel final. "
                "Doble clic sobre una fila también permite editar el mes."
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
            text="(se aplica igual en todos los meses del reporte)",
            bg=COLOR_BG, fg="#7A828B", font=("Segoe UI", 8),
        ).pack(side=tk.LEFT, padx=(8, 0))

        # --- Generar ---
        make_flat_button(
            main, "Generar reporte Excel (Multi-Mes)", self._generate, COLOR_TEAL
        ).pack(fill=tk.X)

    # ------------------------------------------------------------------
    def _refresh_roster_label(self):
        if self.roster_records:
            self.roster_label_var.set(f"Listado de socios: {len(self.roster_records)} socios cargados.")
        else:
            self.roster_label_var.set("Listado de socios: (sin cargar todavía)")

    def _refresh_months_table(self):
        self.tree.delete(*self.tree.get_children())
        for i, entry in enumerate(self.month_entries, start=1):
            self.tree.insert(
                "", tk.END,
                values=(i, entry["label"], entry["filename"], len(entry["records"])),
            )

    # ------------------------------------------------------------------
    # Listado de socios
    # ------------------------------------------------------------------
    def _load_roster(self):
        filepath = filedialog.askopenfilename(
            title="Abrir listado de socios (Excel)",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")],
        )
        if not filepath:
            return
        try:
            roster = load_roster_excel(filepath)
        except RosterProcessingError as e:
            show_popup(self, "Error al procesar el listado de socios", str(e), success=False)
            return
        except Exception as e:
            show_popup(self, "Error inesperado", str(e), success=False)
            return
        self.roster_records = roster
        self._refresh_roster_label()

    # ------------------------------------------------------------------
    # Archivos de Zoom (uno o varios a la vez)
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
                # Se detectó un mes automáticamente (por el nombre del mes
                # en el archivo, o por la fecha numérica que Zoom agrega
                # por defecto al exportarlo) — pero SIEMPRE se le pide
                # confirmación al usuario antes de darlo por bueno, ya que
                # reuniones extraordinarias u otros nombres de archivo
                # podrían llevar a una detección incorrecta.
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
                        continue  # el usuario canceló: se omite este archivo
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
                    continue  # el usuario canceló: se omite este archivo
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
        """Evita dos columnas con el mismo nombre de mes si el usuario
        carga el mismo mes dos veces por error (ej. dos CSV de julio)."""
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
    # Generar el reporte
    # ------------------------------------------------------------------
    def _generate(self):
        if not self.roster_records:
            show_popup(
                self, "Falta el listado de socios",
                "Carga primero el listado oficial de socios (Excel).",
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
        filas, month_labels, no_encontrados_por_mes, ambiguos_por_mes = build_multi_month_matrix(
            months, self.roster_records, threshold
        )

        base_dir = self.main_app._export_dir("Cotejados")
        slug_first = self.main_app._slugify(month_labels[0])
        slug_last = self.main_app._slugify(month_labels[-1])
        filename = (
            f"asistencia_multi_{slug_first}.xlsx" if len(month_labels) == 1
            else f"asistencia_multi_{slug_first}_a_{slug_last}.xlsx"
        )
        filepath = base_dir / filename

        try:
            export_multi_month_to_excel(
                filas, month_labels, str(filepath),
                no_encontrados_por_mes=no_encontrados_por_mes,
                ambiguos_por_mes=ambiguos_por_mes,
            )
        except Exception as e:
            show_popup(self, "Error al exportar", str(e), success=False)
            return

        total_no_encontrados = sum(len(v) for v in no_encontrados_por_mes.values())
        total_ambiguos = sum(len(v) for v in ambiguos_por_mes.values())
        pendientes_note = ""
        if total_no_encontrados or total_ambiguos:
            pendientes_note = (
                f"\n\nOjo: sumando todos los meses hubo {total_no_encontrados} "
                f"participante(s) no encontrado(s) en el listado y "
                f"{total_ambiguos} ambiguo(s). Quedaron listados, mes por mes, "
                f"en las pestañas \"No Encontrados\" y \"Ambiguos\" del mismo Excel."
            )

        show_popup(
            self, "Archivo guardado correctamente",
            f"Se guardó en:\nDownloads / Asistencia Zoom / Cotejados\n\n"
            f"con el nombre:\n{filename}{pendientes_note}",
            success=True, open_path=str(filepath), is_file=True,
        )
