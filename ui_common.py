# -*- coding: utf-8 -*-
"""
ui_common.py
============
Piezas de interfaz (colores, botón plano, popup de resultado y apertura
de archivos/carpetas) COMPARTIDAS entre la ventana principal
(desktop_app.py) y la ventana nueva "Reporte Multi-Mes"
(multi_month_window.py).

Por qué existe este archivo:
    Antes, todo esto vivía dentro de desktop_app.py. Al agregar la
    ventana de Reporte Multi-Mes en su propio archivo, si esa ventana
    importara estas funciones directamente desde desktop_app.py se
    produciría una importación circular (desktop_app.py necesita
    importar la clase de la ventana nueva para poder abrirla, y esa
    ventana necesitaría a su vez importar cosas desde desktop_app.py).
    Moviendo estas piezas comunes a un tercer archivo neutral, ambos
    (desktop_app.py y multi_month_window.py) pueden importarlas sin
    depender uno del otro.
"""

import os
import subprocess

import tkinter as tk

# ---------------------------------------------------------------------------
# Paleta de colores — los mismos tonos que usa la versión Android, para
# que todas las ventanas se sientan como "la misma aplicación".
# ---------------------------------------------------------------------------
COLOR_BG = "#F0F2F7"
COLOR_HEADER_BAR = "#DEE9FC"
COLOR_HEADER_TEXT = "#233857"
COLOR_TABLE_HEADER_BG = "#233857"
COLOR_BLUE = "#3373BF"
COLOR_GREEN = "#29A06B"
COLOR_TEAL = "#1F8C99"
COLOR_ORANGE = "#D98C26"
COLOR_GRAY = "#7A828B"
COLOR_ROW_P = "#D7F2DC"
COLOR_ROW_A = "#FBDCDC"


def open_path_native(path: str, is_file: bool) -> bool:
    """Abre un archivo con la app que Windows tenga asociada (Excel,
    LibreOffice, etc.) o una carpeta en el Explorador. Si el archivo no
    tiene ninguna app asociada, muestra el mismo selector nativo
    "Abrir con..." que usa el propio Explorador de Windows al hacer
    clic derecho → Abrir con, en vez de simplemente fallar. Devuelve
    True si se pudo abrir (o al menos mostrar el selector), False si
    ni siquiera eso funcionó."""
    try:
        if not is_file:
            os.startfile(path)  # type: ignore[attr-defined]
            return True
        try:
            os.startfile(path)  # type: ignore[attr-defined]
            return True
        except OSError:
            # No hay ninguna app asociada a esta extensión: se muestra
            # el selector nativo de Windows para elegir con cuál abrirlo
            # (el mismo mecanismo que usa el Explorador internamente).
            subprocess.Popen(["rundll32.exe", "shell32.dll,OpenAs_RunDLL", path])
            return True
    except Exception:
        return False


def make_flat_button(parent, text, command, color):
    """Crea un botón con el mismo look plano y con esquinas de color
    sólido que usa el resto de la aplicación, en vez de los botones
    grises por defecto de Windows."""
    return tk.Button(
        parent, text=text, command=command, bg=color, fg="white",
        activebackground=color, activeforeground="white",
        relief="flat", bd=0, font=("Segoe UI", 10, "bold"),
        padx=12, pady=9, cursor="hand2",
    )


# ---------------------------------------------------------------------------
# Pop-up de mensaje con ícono ✓ / ✕, igual que en la versión Android.
# ---------------------------------------------------------------------------
def show_popup(parent, title, message, success=None, open_path=None, is_file=False):
    win = tk.Toplevel(parent)
    win.title(title)
    win.configure(bg="white")
    win.resizable(False, False)
    win.transient(parent)

    container = tk.Frame(win, bg="white", padx=22, pady=18)
    container.pack(fill=tk.BOTH, expand=True)

    if success is not None:
        icon_color = "#1F8C4B" if success else "#D13B3B"
        icon_text = "\u2713" if success else "\u2715"  # ✓ / ✕
        tk.Label(
            container, text=icon_text, fg=icon_color, bg="white",
            font=("Segoe UI", 32, "bold"),
        ).pack(pady=(0, 10))

    tk.Label(
        container, text=message, bg="white", fg="#262626",
        justify="left", anchor="w", wraplength=420, font=("Segoe UI", 10),
    ).pack(fill=tk.X, pady=(0, 16))

    btns = tk.Frame(container, bg="white")
    btns.pack(fill=tk.X)

    if open_path:
        def _open(*_a):
            ok = open_path_native(open_path, is_file)
            if not ok:
                show_popup(
                    parent, "No se pudo abrir",
                    f"No se pudo abrir automáticamente. Puedes buscarlo aquí:\n{open_path}",
                    success=False,
                )

        tk.Button(
            btns, text=("Abrir archivo" if is_file else "Abrir carpeta"),
            command=_open, bg=COLOR_TEAL, fg="white", relief="flat",
            activebackground=COLOR_TEAL, activeforeground="white",
            font=("Segoe UI", 10, "bold"), padx=10, pady=8, bd=0,
        ).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 8))

    tk.Button(
        btns, text="Cerrar", command=win.destroy, bg=COLOR_GRAY, fg="white",
        activebackground=COLOR_GRAY, activeforeground="white",
        relief="flat", font=("Segoe UI", 10, "bold"), padx=10, pady=8, bd=0,
    ).pack(side=tk.LEFT, expand=True, fill=tk.X)

    win.update_idletasks()
    px = parent.winfo_rootx() + max((parent.winfo_width() - win.winfo_width()) // 2, 0)
    py = parent.winfo_rooty() + max((parent.winfo_height() - win.winfo_height()) // 2, 0)
    win.geometry(f"+{px}+{py}")
    win.grab_set()
