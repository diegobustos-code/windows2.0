# -*- coding: utf-8 -*-
"""
desktop_app.py — Asistencia Zoom (versión Windows / escritorio)
=================================================================
Aplicación de escritorio para Windows con la MISMA lógica de
procesamiento que la versión Android (reutiliza sin ningún cambio los
archivos csv_processor.py, roster_matcher.py y excel_exporter.py), pero
con una interfaz hecha en Tkinter (viene incluido con Python, no hace
falta instalar nada aparte) y pensada para pantallas de PC: ventana
redimensionable, tabla de resultados más grande, y algunas comodidades
extra que en el celular no tenían tanto sentido (por ejemplo, poder
ver más columnas a la vez y una ventana que se puede maximizar).

Se ve y funciona igual que la versión Android:
    - Mismo logo y misma paleta de colores.
    - Pide el mes de la asistencia al abrir (obligatorio), igual que en
      el celular, con botón "Cambiar mes".
    - Abrir CSV de Zoom / listado de socios (Excel), cotejar, buscar,
      filtrar por duración mínima, ajustar el umbral de "Presente" y
      ordenar — misma lógica exacta que las otras dos versiones.
    - Exportar a CSV / Excel / Pendientes, guardando SIEMPRE los
      archivos organizados en:
          Downloads / Asistencia Zoom / Cotejados      (CSV y Excel)
          Downloads / Asistencia Zoom / Por revisar     (Pendientes)
      con un botón para abrir el archivo recién exportado (con la app
      que Windows tenga asociada, o el selector "Abrir con" si no hay
      ninguna) y otro para abrir directamente esa carpeta en el
      Explorador de Windows.
    - Pop-ups con ✓ verde / ✕ roja según el resultado, igual que en el
      celular.

Cómo compilarlo a un .exe: ver build-exe.yml (compila automáticamente
con GitHub Actions, igual que build-apk.yml compila el APK) o correr
manualmente PyInstaller — instrucciones en README-DESKTOP.md.
"""

import os
import re
import sys
import unicodedata
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog

from csv_processor import (
    CSVProcessingError,
    compute_attendance_flag,
    export_records_to_csv,
    is_disguised_excel_file,
    load_zoom_csv,
)
from excel_exporter import export_to_excel
from roster_matcher import (
    RosterProcessingError,
    load_roster_excel,
    match_zoom_to_roster,
)
from ui_common import (
    COLOR_BG,
    COLOR_HEADER_BAR,
    COLOR_HEADER_TEXT,
    COLOR_TABLE_HEADER_BG,
    COLOR_BLUE,
    COLOR_GREEN,
    COLOR_TEAL,
    COLOR_ORANGE,
    COLOR_GRAY,
    COLOR_ROW_P,
    COLOR_ROW_A,
    open_path_native,
    make_flat_button,
    show_popup,
)
from multi_month_window import MultiMonthWindow

APP_TITLE = "Asistencia Zoom"
DEFAULT_THRESHOLD = "30"

SORT_OPTIONS = [
    "Duración (mayor a menor)",
    "Duración (menor a mayor)",
    "Nombre (A-Z)",
    "Apellido (A-Z)",
    "Sede (A-Z)",
    "Asistencia",
]


def resource_path(relative_path: str) -> str:
    """Devuelve la ruta correcta a un recurso (ej. logo.png / logo.ico),
    tanto si el programa se corre directo con "python desktop_app.py"
    como si ya está empaquetado como .exe con PyInstaller. PyInstaller
    (en modo "un solo archivo") descomprime los archivos incluidos en
    una carpeta temporal distinta en cada ejecución, cuya ruta queda
    disponible en sys._MEIPASS — sin este truco, el programa no
    encontraría el logo dentro del .exe."""
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


class ZoomAttendanceDesktopApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1080x680")
        self.minsize(860, 560)
        self.configure(bg=COLOR_BG)
        try:
            self.iconbitmap(resource_path("logo.ico"))
        except Exception:
            pass  # Si por algún motivo el ícono no carga, la app igual funciona.

        # --- Estado interno (idéntico en espíritu al de la versión Android) ---
        self.zoom_records = []
        self.roster_records = []
        self.all_records = []
        self.filtered_records = []
        self.no_encontrados = []
        self.ambiguos = []
        self.mode = "zoom"
        self.zoom_loaded = False
        self.roster_loaded = False
        self.attendance_month = ""

        self._build_style()
        self._build_ui()

        # El mes se pide con un pequeño retraso para que la ventana
        # principal ya esté armada y centrada antes de mostrar el popup.
        self.after(250, self._ask_month_blocking)

    # ------------------------------------------------------------------
    # Estilo visual
    # ------------------------------------------------------------------
    def _build_style(self):
        style = ttk.Style(self)
        # "clam" es el único theme incluido con Tk que permite personalizar
        # colores de fondo en botones/Treeview de forma confiable en Windows.
        style.theme_use("clam")

        style.configure("Treeview", rowheight=26, font=("Segoe UI", 10),
                         background="white", fieldbackground="white")
        style.configure(
            "Treeview.Heading",
            background=COLOR_TABLE_HEADER_BG, foreground="white",
            font=("Segoe UI", 10, "bold"), relief="flat",
        )
        style.map("Treeview.Heading", background=[("active", COLOR_TABLE_HEADER_BG)])

        style.configure("TCombobox", font=("Segoe UI", 10))

    def _make_button(self, parent, text, command, color, disabled_getter=None):
        """Crea un botón con el mismo look plano y con esquinas de color
        sólido que usa la versión Android (AppButton), en vez de los
        botones grises por defecto de Windows."""
        btn = tk.Button(
            parent, text=text, command=command, bg=color, fg="white",
            activebackground=color, activeforeground="white",
            relief="flat", bd=0, font=("Segoe UI", 10, "bold"),
            padx=12, pady=9, cursor="hand2",
        )
        if disabled_getter is not None:
            btn._disabled_getter = disabled_getter
            btn._enabled_color = color
        return btn

    def _refresh_button_states(self):
        """Habilita/deshabilita botones según el estado actual (equivalente
        a los "disabled: ..." del KV en la versión Android)."""
        self.btn_cotejar.config(
            state=("normal" if (self.zoom_loaded and self.roster_loaded) else "disabled")
        )
        results_state = "normal" if self.filtered_records else "disabled"
        self.btn_export_csv.config(state=results_state)
        self.btn_export_excel.config(state=results_state)
        self.btn_pendientes.config(
            state=("normal" if (self.no_encontrados or self.ambiguos) else "disabled")
        )

    # ------------------------------------------------------------------
    # Construcción de la interfaz
    # ------------------------------------------------------------------
    def _build_ui(self):
        main = tk.Frame(self, bg=COLOR_BG, padx=12, pady=10)
        main.pack(fill=tk.BOTH, expand=True)

        # --- Barra superior: logo + mes + "Cambiar mes" ---
        header = tk.Frame(main, bg=COLOR_HEADER_BAR, padx=10, pady=6)
        header.pack(fill=tk.X, pady=(0, 8))

        self._logo_photo = None
        try:
            from PIL import Image, ImageTk
            img = Image.open(resource_path("logo.png")).convert("RGBA")
            img.thumbnail((36, 36))
            self._logo_photo = ImageTk.PhotoImage(img)
            tk.Label(header, image=self._logo_photo, bg=COLOR_HEADER_BAR).pack(side=tk.LEFT)
        except Exception:
            pass  # Si Pillow o el logo no están disponibles, se omite sin romper la app.

        self.month_label_var = tk.StringVar(value="Mes: (sin definir)")
        tk.Label(
            header, textvariable=self.month_label_var, bg=COLOR_HEADER_BAR,
            fg=COLOR_HEADER_TEXT, font=("Segoe UI", 11, "bold"), padx=10,
        ).pack(side=tk.LEFT)

        tk.Button(
            header, text="Cambiar mes", command=self.ask_attendance_month,
            bg="#4B5568", fg="white", activebackground="#4B5568",
            activeforeground="white", relief="flat", bd=0,
            font=("Segoe UI", 9, "bold"), padx=10, pady=6, cursor="hand2",
        ).pack(side=tk.RIGHT)

        tk.Button(
            header, text="Reporte Multi-Mes", command=self.open_multi_month_window,
            bg=COLOR_TEAL, fg="white", activebackground=COLOR_TEAL,
            activeforeground="white", relief="flat", bd=0,
            font=("Segoe UI", 9, "bold"), padx=10, pady=6, cursor="hand2",
        ).pack(side=tk.RIGHT, padx=(0, 8))

        # --- Botones principales ---
        toolbar = tk.Frame(main, bg=COLOR_BG)
        toolbar.pack(fill=tk.X, pady=(0, 8))
        btn_zoom = self._make_button(toolbar, "Abrir CSV Zoom", self.open_zoom_dialog, COLOR_BLUE)
        btn_roster = self._make_button(toolbar, "Abrir Socios (Excel)", self.open_roster_dialog, COLOR_BLUE)
        self.btn_cotejar = self._make_button(toolbar, "Cotejar", self.on_cotejar, COLOR_GREEN)
        self.btn_cotejar.config(state="disabled")
        for b in (btn_zoom, btn_roster, self.btn_cotejar):
            b.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=4)

        # --- Estado ---
        self.status_var = tk.StringVar(value="Ningún archivo cargado todavía.")
        tk.Label(
            main, textvariable=self.status_var, bg=COLOR_BG, fg="#2E3440",
            justify="left", anchor="w", font=("Segoe UI", 10), wraplength=1000,
        ).pack(fill=tk.X, pady=(0, 8))

        # --- Filtros: búsqueda, duración mínima, umbral ---
        filters = tk.Frame(main, bg=COLOR_BG)
        filters.pack(fill=tk.X, pady=(0, 6))
        self.search_var = tk.StringVar()
        self.min_duration_var = tk.StringVar()
        self.threshold_var = tk.StringVar(value=DEFAULT_THRESHOLD)

        search_entry = tk.Entry(filters, textvariable=self.search_var, font=("Segoe UI", 10))
        search_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 6))
        search_entry.insert(0, "")
        self._add_placeholder(search_entry, "Buscar nombre / apellido / sede")
        search_entry.bind("<Return>", lambda e: self.apply_filters())

        min_dur_entry = tk.Entry(filters, textvariable=self.min_duration_var, width=8, font=("Segoe UI", 10))
        min_dur_entry.pack(side=tk.LEFT, padx=(0, 6))
        self._add_placeholder(min_dur_entry, "Min. min")
        min_dur_entry.bind("<Return>", lambda e: self.apply_filters())

        threshold_entry = tk.Entry(filters, textvariable=self.threshold_var, width=8, font=("Segoe UI", 10))
        threshold_entry.pack(side=tk.LEFT)
        threshold_entry.bind("<Return>", lambda e: self.apply_filters())

        # --- Orden ---
        sortbar = tk.Frame(main, bg=COLOR_BG)
        sortbar.pack(fill=tk.X, pady=(0, 8))
        self.sort_var = tk.StringVar(value=SORT_OPTIONS[0])
        sort_combo = ttk.Combobox(
            sortbar, textvariable=self.sort_var, values=SORT_OPTIONS,
            state="readonly", font=("Segoe UI", 10),
        )
        sort_combo.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 6))
        sort_combo.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())
        apply_btn = self._make_button(sortbar, "Aplicar filtro", self.apply_filters, COLOR_BLUE)
        apply_btn.pack(side=tk.LEFT)

        # --- Tabla de resultados ---
        table_frame = tk.Frame(main, bg=COLOR_BG)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        columns = ("nombre", "apellido", "sede", "duracion", "asistencia")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
        headings = {
            "nombre": ("Nombre", 200),
            "apellido": ("Apellido", 200),
            "sede": ("Sede", 150),
            "duracion": ("Min.", 80),
            "asistencia": ("Asist.", 70),
        }
        for col, (label, width) in headings.items():
            anchor = "center" if col == "asistencia" else ("e" if col == "duracion" else "w")
            self.tree.heading(col, text=label)
            self.tree.column(col, width=width, anchor=anchor, stretch=(col in ("nombre", "apellido", "sede")))

        self.tree.tag_configure("P", background=COLOR_ROW_P)
        self.tree.tag_configure("A", background=COLOR_ROW_A)

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # --- Contador ---
        self.count_var = tk.StringVar(value="Total: 0 | Filtrados: 0")
        tk.Label(
            main, textvariable=self.count_var, bg=COLOR_BG, fg="#2E3440",
            font=("Segoe UI", 9),
        ).pack(fill=tk.X, pady=(0, 6))

        # --- Botones de exportación ---
        export_bar = tk.Frame(main, bg=COLOR_BG)
        export_bar.pack(fill=tk.X)
        self.btn_export_csv = self._make_button(export_bar, "Exportar CSV", self.export_csv, COLOR_TEAL)
        self.btn_export_excel = self._make_button(export_bar, "Exportar Excel", self.export_excel, COLOR_TEAL)
        self.btn_pendientes = self._make_button(export_bar, "Pendientes", self.export_pendientes, COLOR_ORANGE)
        btn_folder = self._make_button(export_bar, "Abrir carpeta", self.open_export_folder, "#4B5568")
        for b in (self.btn_export_csv, self.btn_export_excel, self.btn_pendientes, btn_folder):
            b.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=4)

        self.btn_export_csv.config(state="disabled")
        self.btn_export_excel.config(state="disabled")
        self.btn_pendientes.config(state="disabled")

    @staticmethod
    def _add_placeholder(entry: tk.Entry, placeholder: str):
        """Tkinter no trae "hint_text" incorporado como Kivy: se simula
        mostrando un texto gris que desaparece al hacer foco y reaparece
        si el campo queda vacío."""
        entry.insert(0, placeholder)
        entry.config(fg="#9AA0A6")

        def _on_focus_in(_e):
            if entry.get() == placeholder:
                entry.delete(0, tk.END)
                entry.config(fg="#1A1A1A")

        def _on_focus_out(_e):
            if not entry.get():
                entry.insert(0, placeholder)
                entry.config(fg="#9AA0A6")

        entry.bind("<FocusIn>", _on_focus_in)
        entry.bind("<FocusOut>", _on_focus_out)
        entry._placeholder = placeholder

    @staticmethod
    def _real_value(entry: tk.Entry) -> str:
        """Devuelve el texto real del campo, o "" si lo que hay escrito
        es solo el texto de ejemplo (placeholder)."""
        placeholder = getattr(entry, "_placeholder", None)
        value = entry.get()
        return "" if value == placeholder else value

    # ------------------------------------------------------------------
    # Mes de la asistencia (obligatorio, igual que en la versión Android)
    # ------------------------------------------------------------------
    def ask_attendance_month(self):
        self._open_month_popup(force=False)

    def _ask_month_blocking(self):
        self._open_month_popup(force=True)

    def _open_month_popup(self, force):
        win = tk.Toplevel(self)
        win.title("Mes de la asistencia (obligatorio)" if force else "Cambiar mes")
        win.configure(bg="white")
        win.resizable(False, False)
        win.transient(self)

        frame = tk.Frame(win, bg="white", padx=22, pady=18)
        frame.pack(fill=tk.BOTH, expand=True)

        intro = (
            "Antes de continuar, ingresa el MES de esta asistencia\n"
            "(ejemplo: \"Julio 2026\").\n\n"
            "Es obligatorio: se usa para nombrar los reportes exportados\n"
            "y para ordenarlos en sus carpetas."
        ) if force else "Ingresa el nuevo mes de la asistencia:"

        tk.Label(
            frame, text=intro, bg="white", fg="#262626", justify="left",
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(0, 12))

        month_var = tk.StringVar(value=self.attendance_month)
        entry = tk.Entry(frame, textvariable=month_var, font=("Segoe UI", 11), width=30)
        entry.pack(fill=tk.X, pady=(0, 6))
        entry.focus_set()

        error_var = tk.StringVar(value="")
        tk.Label(frame, textvariable=error_var, bg="white", fg="#C0392B", font=("Segoe UI", 9)).pack(anchor="w")

        def _confirm(*_a):
            value = month_var.get().strip()
            if not value:
                error_var.set("Debes ingresar el mes para poder continuar.")
                return
            self.attendance_month = value
            self.month_label_var.set(f"Mes: {value}")
            win.destroy()

        def _on_close():
            if force and not self.attendance_month:
                return  # Obligatorio: no se deja cerrar sin un mes válido.
            win.destroy()

        entry.bind("<Return>", _confirm)
        win.protocol("WM_DELETE_WINDOW", _on_close)

        tk.Button(
            frame, text="Continuar", command=_confirm, bg=COLOR_BLUE, fg="white",
            activebackground=COLOR_BLUE, activeforeground="white",
            relief="flat", bd=0, font=("Segoe UI", 10, "bold"), pady=8,
        ).pack(fill=tk.X, pady=(12, 0))

        win.update_idletasks()
        x = self.winfo_rootx() + max((self.winfo_width() - win.winfo_width()) // 2, 0)
        y = self.winfo_rooty() + max((self.winfo_height() - win.winfo_height()) // 2, 0)
        win.geometry(f"+{x}+{y}")
        win.grab_set()
        win.wait_window()

    def _month_ready(self) -> bool:
        if not self.attendance_month.strip():
            self.ask_attendance_month()
            return bool(self.attendance_month.strip())
        return True

    @staticmethod
    def _slugify(text: str) -> str:
        text = unicodedata.normalize("NFKD", text or "")
        text = "".join(c for c in text if not unicodedata.combining(c))
        text = re.sub(r"[^A-Za-z0-9]+", "_", text).strip("_")
        return text or "sin_mes"

    def _month_slug(self) -> str:
        return self._slugify(self.attendance_month)

    # ------------------------------------------------------------------
    # Selección y carga de archivos
    # ------------------------------------------------------------------
    def open_zoom_dialog(self):
        filepath = filedialog.askopenfilename(
            title="Abrir CSV de Zoom",
            filetypes=[("CSV o Excel", "*.csv *.xlsx *.xls"), ("Todos los archivos", "*.*")],
        )
        if filepath:
            self._load_zoom_file(filepath)

    def open_roster_dialog(self):
        filepath = filedialog.askopenfilename(
            title="Abrir listado de socios (Excel)",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")],
        )
        if filepath:
            self._load_roster_file(filepath)

    def _load_zoom_file(self, filepath):
        try:
            was_disguised = is_disguised_excel_file(filepath)
            records, _column_map = load_zoom_csv(filepath)
        except CSVProcessingError as e:
            show_popup(self, "Error al procesar el CSV", str(e), success=False)
            return
        except Exception as e:
            show_popup(self, "Error inesperado", str(e), success=False)
            return

        self.zoom_records = records
        self.zoom_loaded = True
        self.mode = "zoom"
        self.all_records = self.zoom_records
        self.no_encontrados = []
        self.ambiguos = []

        self._reset_filters()
        self.apply_filters()

        nombre_archivo = os.path.basename(filepath)
        aviso = " (era un Excel disfrazado de .csv, detectado automáticamente)" if was_disguised else ""
        self.status_var.set(f"Zoom: {nombre_archivo}{aviso} — {len(records)} participantes.")
        self._refresh_button_states()

    def _load_roster_file(self, filepath):
        try:
            roster = load_roster_excel(filepath)
        except RosterProcessingError as e:
            show_popup(self, "Error al procesar el listado de socios", str(e), success=False)
            return
        except Exception as e:
            show_popup(self, "Error inesperado", str(e), success=False)
            return

        self.roster_records = roster
        self.roster_loaded = True
        nombre_archivo = os.path.basename(filepath)
        siguiente = "Presiona Cotejar." if self.zoom_loaded else "Ahora abre también el CSV de Zoom."
        self.status_var.set(f"Socios: {nombre_archivo} — {len(roster)} socios. {siguiente}")
        self._refresh_button_states()

    # ------------------------------------------------------------------
    # Cotejo
    # ------------------------------------------------------------------
    def on_cotejar(self):
        if not self._month_ready():
            return
        if not self.zoom_records or not self.roster_records:
            show_popup(self, "Faltan archivos", "Debes cargar el CSV de Zoom y el listado de socios.")
            return

        threshold = self._safe_float(self.threshold_var.get(), 30.0)
        resultado, no_encontrados, ambiguos = match_zoom_to_roster(
            self.zoom_records, self.roster_records, threshold
        )
        self.all_records = resultado
        self.no_encontrados = no_encontrados
        self.ambiguos = ambiguos
        self.mode = "cotejo"

        self._reset_filters(keep_threshold=True)
        self.apply_filters()

        p = sum(1 for r in resultado if r["Asistencia"] == "P")
        a = sum(1 for r in resultado if r["Asistencia"] == "A")
        self.status_var.set(
            f"Cotejo listo: {p} 'P' y {a} 'A' de {len(resultado)} socios. "
            f"Sin coincidencia: {len(no_encontrados)} | Ambiguos: {len(ambiguos)}."
        )
        self._refresh_button_states()

    # ------------------------------------------------------------------
    # Filtros, orden y refresco de la tabla
    # ------------------------------------------------------------------
    def _reset_filters(self, keep_threshold=False):
        self.search_var.set("")
        self.min_duration_var.set("")
        if not keep_threshold:
            self.threshold_var.set(DEFAULT_THRESHOLD)

    @staticmethod
    def _safe_float(text, default):
        text = (text or "").strip().replace(",", ".")
        if not text:
            return default
        try:
            return float(text)
        except ValueError:
            return default

    def apply_filters(self):
        if not self.all_records:
            self._refresh_button_states()
            return

        threshold = self._safe_float(self.threshold_var.get(), 30.0)
        for record in self.all_records:
            record["Asistencia"] = compute_attendance_flag(
                record.get("Duración (minutos)", 0), threshold
            )

        search_text = self.search_var.get().strip().lower()
        min_duration = self._safe_float(self.min_duration_var.get(), 0.0)

        filtered = []
        for record in self.all_records:
            if record["Duración (minutos)"] < min_duration:
                continue
            if search_text:
                haystack = " ".join([
                    str(record.get("Nombre", "")),
                    str(record.get("Apellido", "")),
                    str(record.get("Sede", "")),
                ]).lower()
                if search_text not in haystack:
                    continue
            filtered.append(record)

        sort_choice = self.sort_var.get()
        if sort_choice == "Duración (mayor a menor)":
            filtered.sort(key=lambda r: r["Duración (minutos)"], reverse=True)
        elif sort_choice == "Duración (menor a mayor)":
            filtered.sort(key=lambda r: r["Duración (minutos)"])
        elif sort_choice == "Nombre (A-Z)":
            filtered.sort(key=lambda r: str(r.get("Nombre", "")).lower())
        elif sort_choice == "Apellido (A-Z)":
            filtered.sort(key=lambda r: str(r.get("Apellido", "")).lower())
        elif sort_choice == "Sede (A-Z)":
            filtered.sort(key=lambda r: str(r.get("Sede", "")).lower())
        elif sort_choice == "Asistencia":
            filtered.sort(key=lambda r: str(r.get("Asistencia", "")))

        self.filtered_records = filtered

        self.tree.delete(*self.tree.get_children())
        for r in filtered:
            tag = r.get("Asistencia", "")
            self.tree.insert(
                "", tk.END,
                values=(
                    r.get("Nombre", ""),
                    r.get("Apellido", ""),
                    r.get("Sede", ""),
                    r.get("Duración (minutos)", 0),
                    r.get("Asistencia", ""),
                ),
                tags=(tag,) if tag in ("P", "A") else (),
            )

        etiqueta = "socios" if self.mode == "cotejo" else "participantes"
        self.count_var.set(f"Total {etiqueta}: {len(self.all_records)} | Filtrados: {len(filtered)}")
        self._refresh_button_states()

    # ------------------------------------------------------------------
    # Ventana "Reporte Multi-Mes"
    # ------------------------------------------------------------------
    def open_multi_month_window(self):
        MultiMonthWindow(self)

    # ------------------------------------------------------------------
    # Carpeta de exportación: Downloads / Asistencia Zoom / <subcarpeta>
    # (misma estructura que usa la versión Android, cambiando "Download"
    # por la carpeta de Descargas real de Windows).
    # ------------------------------------------------------------------
    def _export_base_dir(self) -> Path:
        downloads = Path.home() / "Downloads"
        if not downloads.is_dir():
            downloads = Path.home()
        return downloads / APP_TITLE

    def _export_dir(self, subfolder: str) -> Path:
        target = self._export_base_dir() / subfolder
        target.mkdir(parents=True, exist_ok=True)
        return target

    def open_export_folder(self):
        base = self._export_base_dir()
        if not base.is_dir():
            show_popup(
                self, "Sin carpeta",
                "Todavía no se ha exportado nada en esta sesión, así que "
                "no existe ninguna carpeta para abrir. Exporta un archivo primero.",
            )
            return
        if not open_path_native(str(base), is_file=False):
            show_popup(self, "No se pudo abrir", f"La carpeta existe en:\n{base}", success=False)

    # ------------------------------------------------------------------
    # Exportación
    # ------------------------------------------------------------------
    def export_csv(self):
        if not self._month_ready():
            return
        if not self.filtered_records:
            show_popup(self, "Sin datos", "No hay resultados filtrados para exportar.")
            return
        folder = self._export_dir("Cotejados")
        filename = f"asistencia_{self._month_slug()}.csv"
        filepath = folder / filename
        try:
            export_records_to_csv(self.filtered_records, str(filepath))
        except Exception as e:
            show_popup(self, "Error al exportar", str(e), success=False)
            return
        show_popup(
            self, "Archivo guardado correctamente",
            f"Se guardó en:\nDownloads / {APP_TITLE} / Cotejados\n\ncon el nombre:\n{filename}",
            success=True, open_path=str(filepath), is_file=True,
        )

    def export_excel(self):
        if not self._month_ready():
            return
        if not self.filtered_records:
            show_popup(self, "Sin datos", "No hay resultados filtrados para exportar.")
            return
        folder = self._export_dir("Cotejados")
        filename = f"asistencia_{self._month_slug()}.xlsx"
        filepath = folder / filename
        try:
            export_to_excel(self.filtered_records, str(filepath))
        except Exception as e:
            show_popup(self, "Error al exportar", str(e), success=False)
            return
        show_popup(
            self, "Archivo guardado correctamente",
            f"Se guardó en:\nDownloads / {APP_TITLE} / Cotejados\n\ncon el nombre:\n{filename}",
            success=True, open_path=str(filepath), is_file=True,
        )

    def export_pendientes(self):
        if not self._month_ready():
            return
        if not self.no_encontrados and not self.ambiguos:
            show_popup(self, "Sin pendientes", "No hay registros pendientes de revisión.")
            return
        import csv as csv_module
        folder = self._export_dir("Por revisar")
        filename = f"pendientes_{self._month_slug()}.csv"
        filepath = folder / filename
        try:
            with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv_module.DictWriter(
                    f, fieldnames=["Categoría", "Nombre en Zoom", "Duración (minutos)"]
                )
                writer.writeheader()
                for r in self.no_encontrados:
                    writer.writerow({
                        "Categoría": "No encontrado en el listado oficial",
                        "Nombre en Zoom": r.get("Nombre", ""),
                        "Duración (minutos)": r.get("Duración (minutos)", 0),
                    })
                for r in self.ambiguos:
                    writer.writerow({
                        "Categoría": "Coincide con más de un socio (ambiguo)",
                        "Nombre en Zoom": r.get("Nombre", ""),
                        "Duración (minutos)": r.get("Duración (minutos)", 0),
                    })
        except Exception as e:
            show_popup(self, "Error al exportar", str(e), success=False)
            return
        show_popup(
            self, "Archivo guardado correctamente",
            f"Se guardó en:\nDownloads / {APP_TITLE} / Por revisar\n\ncon el nombre:\n{filename}",
            success=True, open_path=str(filepath), is_file=True,
        )


def main():
    app = ZoomAttendanceDesktopApp()
    app.mainloop()


if __name__ == "__main__":
    # Igual que main.py en la versión Android: si algo falla al iniciar
    # (por ejemplo, un módulo que no se empaquetó bien dentro del .exe),
    # se atrapa el error y se muestra en una ventana propia en vez de
    # que el programa se cierre en silencio sin dejar ningún rastro
    # visible — esto es más importante todavía en el .exe porque se
    # compila en modo "--windowed" (sin consola), así que sin esto un
    # error de arranque sería completamente invisible para el usuario.
    try:
        main()
    except Exception:
        error_text = traceback.format_exc()
        try:
            root = tk.Tk()
            root.withdraw()
            top = tk.Toplevel(root)
            top.title("Error al iniciar Asistencia Zoom")
            top.geometry("700x450")
            text = tk.Text(top, wrap="word", font=("Consolas", 9))
            text.insert(
                "1.0",
                "La aplicación encontró un error al iniciar. Copia este "
                "texto y compártelo para poder corregirlo:\n\n" + error_text,
            )
            text.configure(state="disabled")
            text.pack(fill=tk.BOTH, expand=True)
            top.protocol("WM_DELETE_WINDOW", root.destroy)
            root.mainloop()
        except Exception:
            print("ERROR FATAL AL INICIAR LA APP:")
            print(error_text)
