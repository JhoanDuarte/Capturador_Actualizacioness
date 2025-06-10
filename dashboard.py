import os
import ssl
import sys
import csv

# — Configuración de entorno para GTK/Cairo —
# Ajusta PATH para cargar las DLL de GTK sin privilegios de administrador
os.environ['PATH'] = (
    r"C:\Users\pysnepsdbs08\gtk3-runtime\bin"
    + os.pathsep + os.environ.get('PATH', '')
)
# Inserta ruta a site-packages para cargar la instalación correcta de Pandas y CairoSVG
sys.path.insert(0,
    r"C:\Users\pysnepsdbs08\AppData\Local\Programs\Python\Python313\Lib\site-packages"
)

# — Librerías estándar —
import datetime
import re
import subprocess
import tkinter as tk
from io import BytesIO
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtWidgets import QGraphicsDropShadowEffect, QGraphicsBlurEffect, QGraphicsScene, QGraphicsPixmapItem
from PyQt5.QtGui import QPainter, QPainterPath, QImage, QRegion, QColor
from PyQt5.QtCore import QRectF, QRect, QPoint
from tkinter import filedialog, messagebox, ttk

# — Terceros —
import bcrypt
import cairosvg
import customtkinter as ctk  # sólo esto para CustomTkinter
from PIL import Image
import pandas as pd
import requests
import math
from tkcalendar import DateEntry
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageBreak, Paragraph
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
import traceback
from PIL import Image, ImageDraw, ImageTk
from threading import Thread

# — Módulos propios —
from db_connection import conectar_sql_server

def safe_destroy(win):
    # cancela todos los after del intérprete
    try:
        for aid in win.tk.call('after', 'info'):
            try:
                win.tk.call('after', 'cancel', aid)
            except Exception:
                pass
    except Exception:
        pass
    win.destroy()


ICON_CACHE = {}

def load_icon_from_url(url, size):
    """Descarga un ícono SVG y lo convierte a CTkImage usando caché."""
    key = (url, size)
    if key in ICON_CACHE:
        return ICON_CACHE[key]
    try:
        resp = requests.get(url, timeout=10)
        resp.raise_for_status()
        png_bytes = cairosvg.svg2png(bytestring=resp.content,
                                     output_width=size[0],
                                     output_height=size[1])
        img = Image.open(BytesIO(png_bytes))
    except Exception:
        img = Image.new("RGBA", size, (0, 0, 0, 0))
    icon = ctk.CTkImage(light_image=img, dark_image=img, size=size)
    ICON_CACHE[key] = icon
    return icon


def prefetch_icons(urls, size=(18, 18)):
    """Pre-carga íconos en un hilo para acelerar la apertura de ventanas."""
    def _worker():
        for u in urls:
            key = (u, size)
            if key not in ICON_CACHE:
                try:
                    load_icon_from_url(u, size)
                except Exception:
                    ICON_CACHE[key] = ctk.CTkImage(
                        light_image=Image.new("RGBA", size, (0, 0, 0, 0)),
                        dark_image=Image.new("RGBA", size, (0, 0, 0, 0)),
                        size=size,
                    )
    Thread(target=_worker, daemon=True).start()

# Lista de íconos usados con frecuencia en la aplicación
COMMON_ICONS = [
    "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/save.svg",
    "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/plus-circle.svg",
    "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/trash-alt.svg",
    "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/sign-out-alt.svg",
    "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/user-circle.svg",
    "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/calendar.svg",
    "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/id-card.svg",
    "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/hashtag.svg",
    "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/stethoscope.svg",
    "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/file-invoice.svg",
    "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/tools.svg",
    "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/list-ol.svg",
    "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/dollar-sign.svg",
    "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/coins.svg",
    "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/align-left.svg",
]

# Prefetch en un hilo para que las ventanas abran más rápido
prefetch_icons(COMMON_ICONS)



def apply_ctk_theme_from_settings():
    """Configura la apariencia de CustomTkinter según el tema guardado."""
    settings = QtCore.QSettings("Procesos Y Servicios", "CapturadorDeDatos")
    theme = settings.value("theme", "dark")
    if theme == "light":
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
    else:
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")
    return theme


class AutocompleteEntry(ctk.CTkEntry):
    def __init__(self, parent, values, textvariable=None, **kwargs):
        # Preparar StringVar (propio o el que pase el usuario)
        
        if textvariable is None:
            self.var = tk.StringVar()
        else:
            self.var = textvariable

        # Evitar duplicar textvariable
        kwargs.pop('textvariable', None)
        super().__init__(parent, **kwargs)
        self.configure(textvariable=self.var)

        self._values = values
        self._listbox_window = None
        self._listbox = None

        # Bindings
        self.var.trace_add('write', lambda *args: self._show_matches())
        self.bind('<Down>', self._on_down)
        self.bind('<Escape>', lambda e: self._hide_listbox())
        # Evitar cerrar desplegable al interactuar
        self.bind('<FocusOut>', lambda e: None)

    def _show_matches(self):
        txt = self.var.get().strip().lower()
        if not txt:
            return self._hide_listbox()

        matches = [v for v in self._values if v.lower().startswith(txt)]
        if not matches:
            return self._hide_listbox()

        if not self._listbox_window:
            self._listbox_window = tk.Toplevel(self)
            self._listbox_window.overrideredirect(True)
            lb = tk.Listbox(self._listbox_window)
            lb.pack(expand=True, fill='both')
            lb.bind('<<ListboxSelect>>', self._on_listbox_select)
            lb.bind('<Return>',          self._on_listbox_select)
            lb.bind('<Up>',              self._on_listbox_nav)
            lb.bind('<Down>',            self._on_listbox_nav)
            self._listbox = lb
        else:
            self._listbox.delete(0, tk.END)

        for m in matches:
            self._listbox.insert(tk.END, m)

        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        w = self.winfo_width()
        h = min(100, len(matches) * 20)
        self._listbox_window.geometry(f"{w}x{h}+{x}+{y}")

    def _on_listbox_select(self, event):
        if not self._listbox:
            return
        sel = self._listbox.get(self._listbox.curselection())
        self.var.set(sel)
        self.icursor(tk.END)
        self._hide_listbox()

    def _on_listbox_nav(self, event):
        if not self._listbox:
            return "break"
        idx = self._listbox.curselection()
        if not idx:
            self._listbox.selection_set(0)
            idx = (0,)
        i = idx[0]
        if event.keysym == 'Up' and i > 0:
            self._listbox.selection_clear(0, tk.END)
            self._listbox.selection_set(i - 1)
        elif event.keysym == 'Down' and i < self._listbox.size() - 1:
            self._listbox.selection_clear(0, tk.END)
            self._listbox.selection_set(i + 1)
        return "break"

    def _on_down(self, event):
        self._show_matches()
        if self._listbox:
            self._listbox.focus_set()
            self._listbox.selection_set(0)
            self._listbox.activate(0)
        return "break"

    def _hide_listbox(self):
        if self._listbox_window:
            self._listbox_window.destroy()
            self._listbox_window = None
            self._listbox = None

class CodeAutocompleteEntry(AutocompleteEntry):
    """
    Un AutocompleteEntry que muestra en la lista "CÓDIGO – NOMBRE"
    pero al seleccionar sólo deja el CÓDIGO en el StringVar.
    
    Parámetros:
      parent         – widget padre
      code_to_name   – dict mapping código (str) a nombre (str)
      textvariable   – StringVar opcional donde se volcará sólo el código
    """
    def __init__(self, parent, code_to_name, textvariable=None, **kwargs):
        # Construimos el mapa inverso "CÓDIGO – NOMBRE" -> "CÓDIGO"
        self._display_to_code = {
            f"{code} - {name}": code
            for code, name in code_to_name.items()
        }
        # Llamamos al AutocompleteEntry con la lista de valores a mostrar
        super().__init__(
            parent,
            values=list(self._display_to_code.keys()),
            textvariable=textvariable,
            **kwargs
        )

    def _on_listbox_select(self, event):
        """Cuando el usuario selecciona un ítem (doble clic o Enter)."""
        if not self._listbox:
            return
        try:
            sel = self._listbox.get(self._listbox.curselection())
        except Exception:
            return
        # Traducimos al código puro y lo ponemos en el StringVar
        code = self._display_to_code.get(sel, sel)
        self.var.set(code)
        self.icursor('end')
        self._hide_listbox()

    def _show_matches(self, *args):
        """
        Después de mostrar el listado, volvemos a enlazar Enter
        al nuevo _on_listbox_select.
        """
        super()._show_matches(*args)
        if self._listbox:
            # Asegurarnos de que Enter ejecute nuestra selección personalizada
            self._listbox.unbind('<Return>')
            self._listbox.bind('<Return>', self._on_listbox_select)
            
class FullMatchAutocompleteEntry(AutocompleteEntry):
    def _show_matches(self):
        txt = self.var.get().strip().lower()
        if not txt:
            return self._hide_listbox()

        # Aquí cambiamos startswith por in
        matches = [v for v in self._values if txt in v.lower()]
        if not matches:
            return self._hide_listbox()

        # — el resto igual —
        if not self._listbox_window:
            self._listbox_window = tk.Toplevel(self)
            self._listbox_window.overrideredirect(True)
            lb = tk.Listbox(self._listbox_window)
            lb.pack(expand=True, fill='both')
            lb.bind('<<ListboxSelect>>', self._on_listbox_select)
            lb.bind('<Return>',          self._on_listbox_select)
            lb.bind('<Up>',              self._on_listbox_nav)
            lb.bind('<Down>',            self._on_listbox_nav)
            self._listbox = lb
        else:
            self._listbox.delete(0, tk.END)

        for m in matches:
            self._listbox.insert(tk.END, m)

        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        w = self.winfo_width()
        h = min(100, len(matches) * 20)
        self._listbox_window.geometry(f"{w}x{h}+{x}+{y}")
        
class ScrollableFrame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        # 1) Canvas + scrollbar
        self.canvas = tk.Canvas(self, highlightthickness=0, bg=self.cget("fg_color"))
        self.scrollbar = ctk.CTkScrollbar(self, orientation="vertical", command=self.canvas.yview)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        # 2) Frame interior (donde pondrás tus campos)
        self.scrollable_frame = ctk.CTkFrame(self.canvas, fg_color="transparent")
        self._window = self.canvas.create_window((0,0),
                                                 window=self.scrollable_frame,
                                                 anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        # 3) Actualizar scrollregion al cambiar tamaño
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(
            scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(
            self._window, width=e.width))

    def scroll_to_widget(self, widget, margin=10):
            c = self.canvas
            f = self.scrollable_frame

            # 1) Forzar recálculo de tamaños
            c.update_idletasks()
            f.update_idletasks()

            total_h = f.winfo_height()
            view_h  = c.winfo_height()
            top_frac, top_pix = c.yview()[0], c.yview()[0] * total_h

            # 2) Encontrar el "contenedor" directo en scrollable_frame
            anc = widget
            while anc.master and anc.master != f:
                anc = anc.master
            cont = anc  # este frame engloba etiqueta + entrada

            # 3) Posición relativa y alto del contenedor
            widget_y = cont.winfo_rooty() - f.winfo_rooty()
            widget_h = cont.winfo_height()

            # 4) Si está por encima → subir hasta dejarlo margin px debajo del top
            if widget_y < top_pix + margin:
                new_top = max(widget_y - margin, 0)
                c.yview_moveto(new_top / float(total_h - view_h))

            # 5) Si está por debajo → bajar hasta dejarlo margin px arriba del bottom
            elif widget_y + widget_h > top_pix + view_h - margin:
                new_top = min(widget_y + widget_h + margin - view_h,
                            total_h - view_h)
                c.yview_moveto(new_top / float(total_h - view_h))
        
def iniciar_tipificacion(parent_root, conn, current_user_id):
    
    settings = QtCore.QSettings("Procesos Y Servicios", "CapturadorDeDatos")
    tema_actual = settings.value("theme", "dark")  # “dark” o “light”
    
    if tema_actual == "dark":
        color_container  = "#1e1e1e"  # fondo general “dark”
        color_card       = "#2b2b2b"  # fondo del “card” en dark
        fg_text_color    = "white"    # todo texto va en blanco
        entry_fg_color   = "#424242"  # gris oscuro en vez de blanco
        entry_text_color = "white"    # texto blanco dentro de la entrada
        placeholder_color= "#555555"  # gris medio para placeholder
    else:
        # tema “light-blue”
        color_container   = "#F8F8F8"  # gris clarito casi blanco
        color_card        = "#EAEAEA"  # un pelín más oscuro que container
        fg_text_color    = "black"    # todo texto en negro
        entry_fg_color   = "white"    # campos claros en tema light
        entry_text_color = "black"    # texto negro dentro de CTkEntry
        placeholder_color= "#AAAAAA"  # gris para placeholder
        
    entry_radicado_var = tk.StringVar()
    entry_nit_var      = tk.StringVar()
    entry_factura_var  = tk.StringVar()

    
    # 1) Obtener el último paquete cargado con TIPO_PAQUETE = 'DIGITACION'
    cur = conn.cursor()
    cur.execute("""
        SELECT MAX(NUM_PAQUETE)
        FROM ASIGNACION_TIPIFICACION
        WHERE UPPER(LTRIM(RTRIM(TIPO_PAQUETE))) = %s AND TIPO_PAQUETE = 'DIGITACION'
    """, ("DIGITACION",))
    pkg = cur.fetchone()[0] or 0
    tipo_paquete = 'DIGITACION';
    cur.close()

    # 2) Variables de asignación
    radicado = None
    nit = None
    factura = None
    
    cur2 = conn.cursor()
    cur2.execute("""
        SELECT campo
        FROM PAQUETE_CAMPOS
        WHERE NUM_PAQUETE = %s
          AND tipo_paquete = %s
    """, (pkg, tipo_paquete))
    campos_paquete = {row[0] for row in cur2.fetchall()}
    cur2.close()

    # 3) Ventana principal
    win = ctk.CTkToplevel(parent_root)
    win.title(f"Capturador De Datos · Paquete {pkg}")
    

    def copy_radicado():
        win.clipboard_clear()
        win.clipboard_append(entry_radicado_var.get())

    def load_assignment():
        nonlocal radicado, nit, factura

        cur = conn.cursor()
        cur.execute("""
            SELECT TOP 1 RADICADO, NIT, FACTURA
              FROM ASIGNACION_TIPIFICACION
             WHERE STATUS_ID    = 1
               AND NUM_PAQUETE  = %s
               AND TIPO_PAQUETE = 'DIGITACION'
             ORDER BY NEWID()
        """, (pkg,))
        row = cur.fetchone()
        cur.close()
        if not row:
            return False

        radicado, nit, factura = row
        # marcamos como “en curso”
        cur2 = conn.cursor()
        cur2.execute(
            "UPDATE ASIGNACION_TIPIFICACION SET STATUS_ID = 2 WHERE RADICADO = %s",
            (radicado,)
        )
        conn.commit()
        cur2.close()

        # actualizamos UI y copiamos al portapapeles
        entry_radicado_var.set(str(radicado))
        entry_nit_var.set(str(nit))
        entry_factura_var.set(str(factura))
        copy_radicado()
        return True

    # ————— Llamada inicial —————
    if not load_assignment():
        messagebox.showinfo("Sin asignaciones", "No hay asignaciones pendientes.")
        return

    # Obtener la resolución de la pantalla
    screen_width = win.winfo_screenwidth()
    screen_height = win.winfo_screenheight()

    # Calcular el alto de la ventana para que no cubra la barra de tareas
    taskbar_height = 40  # Estimación del alto de la barra de tareas de Windows (puede variar)
    window_height = screen_height - taskbar_height  # Resta el alto de la barra de tareas

    # Establecer la geometría de la ventana para que ocupe toda la pantalla, pero sin la barra de tareas
    win.geometry(f"{screen_width}x{window_height}")

    # Calcular la posición para centrar la ventana
    center_x = (screen_width // 2) - (screen_width // 2)
    center_y = (window_height // 2) - (window_height // 2)

    # Establecer la nueva geometría centrada
    win.geometry(f"{screen_width}x{window_height}+{center_x}+{center_y}")

    win.grab_set()

    container = ctk.CTkFrame(win, fg_color=color_container)
    container.grid(row=0, column=0, sticky="nsew")
    win.grid_rowconfigure(0, weight=1)
    win.grid_columnconfigure(0, weight=1)

    card = ctk.CTkFrame(container, fg_color=color_card)
    card.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
    container.grid_rowconfigure(0, weight=1)
    container.grid_columnconfigure(0, weight=1)

    # Avatar y título
    avatar = load_icon_from_url(
        "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/user-circle.svg",
        size=(80, 80)
    )
    ctk.CTkLabel(card, image=avatar, text="").pack(pady=(20, 5))
    ctk.CTkLabel(
        card,
        text=f"📦 Paquete #{pkg}",
        font=ctk.CTkFont(size=26, weight='bold'),
        text_color=fg_text_color
    ).pack(pady=(0, 15))


   # 4) Lectura de Radicado, NIT, Factura
    read_frame = ctk.CTkFrame(card, fg_color='transparent')
    read_frame.pack(fill='x', padx=30)
    read_frame.grid_columnconfigure(1, weight=1)

    # Labels fijos
    ctk.CTkLabel(read_frame, text="Radicado:", anchor='w', text_color=fg_text_color).grid(row=0, column=0, pady=5, sticky='w')
    ctk.CTkEntry(
        read_frame,
        textvariable=entry_radicado_var,   # <-- aquí
        fg_color=entry_fg_color,            # fondo de la entrada según tema
        text_color=entry_text_color,        # texto de la entrada según tema
        placeholder_text_color=placeholder_color,
        state='readonly',
        width=300
    ).grid(row=0, column=1, pady=5, sticky='ew', padx=(10,0))

    ctk.CTkLabel(read_frame, text="NIT:", anchor='w', text_color=fg_text_color).grid(row=1, column=0, pady=5, sticky='w')
    ctk.CTkEntry(
        read_frame,
        textvariable=entry_nit_var,        # <-- y aquí
        state='readonly',
        width=300,
        fg_color=entry_fg_color,            # fondo de la entrada según tema
        text_color=entry_text_color,        # texto de la entrada según tema
        placeholder_text_color=placeholder_color
    ).grid(row=1, column=1, pady=5, sticky='ew', padx=(10,0))

    ctk.CTkLabel(read_frame, text="Factura:", anchor='w', text_color=fg_text_color).grid(row=2, column=0, pady=5, sticky='w')
    ctk.CTkEntry(
        read_frame,
        textvariable=entry_factura_var,    # <-- y aquí
        state='readonly',
        width=300,
        fg_color=entry_fg_color,            # fondo de la entrada según tema
        text_color=entry_text_color,        # texto de la entrada según tema
        placeholder_text_color=placeholder_color
    ).grid(row=2, column=1, pady=5, sticky='ew', padx=(10,0))


    # 5) Scrollable y grid de 3 columnas
    scroll = ScrollableFrame(card, fg_color=color_card)
    scroll.pack(fill='both', expand=True, padx=20, pady=(10,0))
    card.pack_propagate(False)
    card.grid_rowconfigure(1, weight=1)
    card.grid_columnconfigure(0, weight=1)
    for col in range(3):
        scroll.scrollable_frame.grid_columnconfigure(col, weight=1, uniform="col")

    # 6) Variables de posición y contenedores
    fixed_row = 0
    fixed_col = 0
    field_vars = {}
    widgets = {}
    detail_vars = []
    service_frames = []
    

    def place_fixed_field(frame):
        nonlocal fixed_row, fixed_col
        frame.grid(row=fixed_row, column=fixed_col, padx=10, pady=8, sticky='nsew')
        fixed_col += 1
        if fixed_col == 3:
            fixed_col = 0
            fixed_row += 1

    def on_close():
        # Si la ventana se cierra sin guardar, cambiamos el estado de la asignación a 1
        if radicado is not None:
            cur = conn.cursor()
            cur.execute(
                """
                UPDATE ASIGNACION_TIPIFICACION
                SET STATUS_ID = 1
                WHERE RADICADO = %s
                """,
                (radicado,),
            )
            conn.commit()
            cur.close()
        win.destroy()  # Cierra la ventana después de actualizar el estado

    # Configurar el evento de cierre de la ventana
    win.protocol("WM_DELETE_WINDOW", on_close)

    # 7) Funciones auxiliares de validación y selecciones
    def select_all(event):
        w = event.widget
        try:
            w.select_range(0, 'end')
            w.icursor('end')
        except: pass

    def clear_selection_on_key(event):
        w = event.widget
        if event.keysym in ('BackSpace', 'Delete'):
            if w.selection_present(): w.delete(0, 'end'); return
        ch = event.char
        if len(ch)==1 and ch.isprintable() and w.selection_present():
            w.delete(0, 'end')

    def bind_select_all(widget):
        for child in widget.winfo_children():
            if isinstance(child, ctk.CTkEntry):
                child.bind("<Double-Button-1>", select_all)
                child.bind("<FocusIn>", select_all)
                child.bind("<Key>", clear_selection_on_key)
            bind_select_all(child)

    def mark_required(w, var):
        def chk(e=None):
            if not var.get().strip():
                w.configure(border_color='red', border_width=2)
            else:
                w.configure(border_color='#2b2b2b', border_width=1)
        w.bind('<FocusOut>', chk)

    def make_field(label_text, icon_url=None):
        frame = ctk.CTkFrame(scroll.scrollable_frame, fg_color='transparent')
        if icon_url:
            ico = load_icon_from_url(icon_url, size=(20,20))
            ctk.CTkLabel(frame, image=ico, text="").pack(side='left', padx=(0,5))
        ctk.CTkLabel(
            frame,
            text=label_text,
            anchor='w',
            text_color=fg_text_color,   # <–– cada Label hereda el color dinámico

            font=("Arial", 12, "bold")
        ).pack(fill='x')
        return frame


    # Función para seleccionar todo el texto en el campo
    def select_all(event):
        w = event.widget
        w.select_range(0, 'end')
        return 'break'

    # Función para borrar todo al presionar cualquier tecla si hay selección
    def clear_selection_on_key(event):
        w = event.widget
        if event.keysym in ('BackSpace', 'Delete'):
            if w.selection_present(): 
                w.delete(0, 'end')
            return
        ch = event.char
        if len(ch) == 1 and ch.isprintable() and w.selection_present():
            w.delete(0, 'end')

    # Función para formatear el campo de fecha mientras el usuario escribe
    def format_fecha(event):
        txt = var_fecha.get()
        # Si es borrado o navegación, no formatear aquí
        if event.keysym in ('BackSpace', 'Delete', 'Left', 'Right', 'Home', 'End'):
            return

        # Quitamos cualquier slash existente y limitamos a 8 dígitos (DDMMYYYY)
        digits = txt.replace('/', '')[:8]

        # Reconstruimos con slashes: DD / MM / AAAA
        parts = []
        if len(digits) >= 2:
            parts.append(digits[:2])
            if len(digits) >= 4:
                parts.append(digits[2:4])
                parts.append(digits[4:])
            else:
                parts.append(digits[2:])
        else:
            parts.append(digits)

        new_text = '/'.join(parts)
        var_fecha.set(new_text)
        entry_fecha.icursor(len(new_text))  # colocamos el cursor al final

    def val_fecha(e=None):
        txt = var_fecha.get().strip()
        try:
            d = datetime.datetime.strptime(txt, '%d/%m/%Y').date()
            if d > datetime.date.today():
                raise ValueError("Fecha futura")
            entry_fecha.configure(border_color='#2b2b2b', border_width=1)
            lbl_err_fecha.configure(text='')
            return True
        except Exception:
            entry_fecha.configure(border_color='red', border_width=2)
            lbl_err_fecha.configure(text='Fecha inválida')
            return False

    def apply_focus_style(widget, scroll):
        # 1) detectamos el widget interno que realmente recibe el foco
        #    para CTkEntry es widget._entry
        real = getattr(widget, '_entry', widget)

        def on_focus_in(e):
            widget.configure(border_color='#3399FF', border_width=2)
            scroll.scroll_to_widget(widget)

        def on_focus_out(e):
            widget.configure(border_color='#2b2b2b', border_width=1)

        # 2) enlazamos los eventos al widget interno
        real.bind('<FocusIn>', on_focus_in)
        real.bind('<FocusOut>', on_focus_out)




    # ————— Bloque de creación del campo de fecha —————

    if 'FECHA_SERVICIO' in campos_paquete:
        frm = make_field(
            'Fecha Servicio:',
            'https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/calendar.svg'
        )
        var_fecha = tk.StringVar()

        entry_fecha = ctk.CTkEntry(
            frm,
            textvariable=var_fecha,
            placeholder_text='DD/MM/AAAA',
            width=300,
            fg_color=entry_fg_color,
            text_color=entry_text_color,
            placeholder_text_color=placeholder_color,
            validate='key',
            validatecommand=(win.register(lambda s: bool(re.match(r"^[0-9/]$", s))), '%S')
        )
        entry_fecha.pack(fill='x', pady=(5, 0))
        apply_focus_style(entry_fecha,scroll)

        # Selección completa en doble-click o focus
        entry_fecha.bind("<Double-Button-1>", select_all)
        entry_fecha.bind("<FocusIn>", select_all)

        # Borra todo al presionar BackSpace o Delete
        entry_fecha.bind("<Key>", clear_selection_on_key)

        # Formateo dinámico al escribir
        entry_fecha.bind("<KeyRelease>", format_fecha)

        lbl_err_fecha = ctk.CTkLabel(frm, text='', text_color='red')
        lbl_err_fecha.pack(fill='x')

        field_vars['FECHA_SERVICIO'] = var_fecha
        widgets['FECHA_SERVICIO']   = entry_fecha

        # Validación al perder foco
        entry_fecha.bind('<FocusOut>', val_fecha)

        # Posicionar en el layout
        place_fixed_field(frm)

    if 'TIPO_DOC_ID' in campos_paquete:
        frm = make_field('Tipo Doc:',
                        'https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/id-card.svg')
        # Carga opciones
        cur_td = conn.cursor()
        cur_td.execute("SELECT NAME FROM TIPO_DOC")
        opts_td = [r[0] for r in cur_td.fetchall()]
        cur_td.close()

        var_tipo = tk.StringVar()
        # Trace: solo A–Z y uppercase
        var_tipo.trace_add('write', lambda *_: var_tipo.set(
            ''.join(ch for ch in var_tipo.get().upper() if 'A' <= ch <= 'Z')
        ))

        entry_tipo = AutocompleteEntry(frm, opts_td, width=300,fg_color=entry_fg_color,
            text_color=entry_text_color,
            placeholder_text_color=placeholder_color, textvariable=var_tipo)
        entry_tipo.pack(fill='x', pady=(5,0))
        apply_focus_style(entry_tipo,scroll)

        # Forzar mayúsculas en KeyRelease
        def to_upper_on_key(event, var=var_tipo):
            var.set(var.get().upper())
        entry_tipo.bind('<KeyRelease>', to_upper_on_key)

        # Etiqueta de error
        lbl_err_td = ctk.CTkLabel(frm, text='', text_color='red')
        lbl_err_td.pack(fill='x', pady=(2,0))

        # Validación al perder foco
        def val_tipo(e=None):
            nombre = var_tipo.get().strip().upper()

            # 1) Obligatorio
            if not nombre:
                entry_tipo.configure(border_color='red', border_width=2)
                lbl_err_td.configure(text='Tipo de documento obligatorio')
                return False

            # 2) Verificar existencia en la base
            cur_chk = conn.cursor()
            cur_chk.execute(
                "SELECT COUNT(*) FROM TIPO_DOC WHERE UPPER(NAME) = %s",
                (nombre,)
            )
            existe = cur_chk.fetchone()[0] > 0
            cur_chk.close()

            if not existe:
                entry_tipo.configure(border_color='red', border_width=2)
                lbl_err_td.configure(text='Tipo de documento no existe')
                return False

            # 3) Todo OK
            entry_tipo.configure(border_color='#2b2b2b', border_width=1)
            lbl_err_td.configure(text='')
            return True

        entry_tipo.bind('<FocusOut>', val_tipo)


        field_vars['TIPO_DOC_ID'] = var_tipo
        widgets['TIPO_DOC_ID']    = entry_tipo
        place_fixed_field(frm)


    if 'NUM_DOC' in campos_paquete:
        frm = make_field('Num Doc:',
                        'https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/hashtag.svg')
        var_num = tk.StringVar()
        entry_num = ctk.CTkEntry(
            frm, textvariable=var_num,
            placeholder_text='Solo dígitos', width=300,
            fg_color=entry_fg_color,
            text_color=entry_text_color,
            placeholder_text_color=placeholder_color,
            validate='key', validatecommand=(win.register(lambda s: s.isdigit()), '%S')
        )
        entry_num.pack(fill='x', pady=(5,0))
        apply_focus_style(entry_num,scroll)

        # Etiqueta de error
        lbl_err_num = ctk.CTkLabel(frm, text='', text_color='red')
        lbl_err_num.pack(fill='x', pady=(2,0))

        # Validación al perder foco
        def val_num(e=None):
            if not var_num.get().strip():
                entry_num.configure(border_color='red', border_width=2)
                lbl_err_num.configure(text='Número de documento obligatorio')
                return False
            else:
                entry_num.configure(border_color='#2b2b2b', border_width=1)
                lbl_err_num.configure(text='')
                return True
        entry_num.bind('<FocusOut>', val_num)

        field_vars['NUM_DOC'] = var_num
        widgets['NUM_DOC']    = entry_num
        place_fixed_field(frm)


    if 'DIAGNOSTICO' in campos_paquete:
        frm = make_field('Diagnóstico:',
                        'https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/stethoscope.svg')

        # Carga mapa CIE10
        cur_dx = conn.cursor()
        cur_dx.execute("SELECT CODIGO, NOMBRE FROM TBL_CIE10")
        dx_map = {cod: nombre for cod, nombre in cur_dx.fetchall()}
        cur_dx.close()

        opciones = [f"{cod} - {nombre}" for cod, nombre in dx_map.items()]

        var_diag = tk.StringVar()
        # Trace: filtrar y uppercase
        var_diag.trace_add('write', lambda *_: var_diag.set(
            ''.join(ch for ch in var_diag.get().upper() if ch.isalnum() or ch in (' ', '-', '_'))
        ))

        entry_diag = FullMatchAutocompleteEntry(
            frm,
            values=opciones,
            width=300,
            fg_color=entry_fg_color,
            text_color=entry_text_color,
            placeholder_text_color=placeholder_color,
            textvariable=var_diag
        )
        entry_diag.pack(fill='x', pady=(5,0))
        apply_focus_style(entry_diag,scroll)

        # Etiqueta de error
        lbl_err_diag = ctk.CTkLabel(frm, text='', text_color='red')
        lbl_err_diag.pack(fill='x', pady=(2,0))

        # Extraer código al seleccionar
        def on_select(event=None):
            text = var_diag.get()
            if " - " in text:
                cod, _ = text.split(" - ", 1)  # Extrae solo el código
                var_diag.set(cod)  # Establece solo el código en el campo de entrada

        # Asegurarse de que on_select siempre se dispare cuando se seleccione un ítem del desplegable
        entry_diag.bind('<<ListboxSelect>>', on_select)  # Al seleccionar un ítem del desplegable

        # También aseguramos que se ejecute al hacer "Enter" después de la selección
        entry_diag.bind('<Return>', on_select)  # Al presionar Enter (si se usa para seleccionar)

        # Asegurarse de que el valor se actualice también cuando el campo pierde el foco (cuando el usuario hace click fuera o tabula)
        def on_focus_out(event):
            on_select(event)

        entry_diag.bind('<FocusOut>', on_focus_out)  # Actualiza al cambiar de campo (FocusOut)

        # Validación al perder foco (primero on_select, luego val)
        def val_diag(e=None):
            on_select()  # extrae el código a var_diag

            codigo = var_diag.get().strip().upper()

            # 1) Obligatorio
            if not codigo:
                entry_diag.configure(border_color='red', border_width=2)
                lbl_err_diag.configure(text='Diagnóstico obligatorio')
                return False

            # 2) Verificar que el código esté en dx_map
            if codigo not in dx_map:
                entry_diag.configure(border_color='red', border_width=2)
                lbl_err_diag.configure(text='Código de diagnóstico no existe')
                return False

            # 3) Todo OK
            entry_diag.configure(border_color='#2b2b2b', border_width=1)
            lbl_err_diag.configure(text='')
            return True

        entry_diag.bind('<FocusOut>', val_diag)

        field_vars['DIAGNOSTICO'] = var_diag
        widgets['DIAGNOSTICO']    = entry_diag
        place_fixed_field(frm)


    # 9) Campos dinámicos
    DETAIL_ICONS = {
        'AUTORIZACION':    'https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/file-invoice.svg',
        'CODIGO_SERVICIO': 'https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/tools.svg',
        'CANTIDAD':        'https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/list-ol.svg',
        'VLR_UNITARIO':    'https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/dollar-sign.svg',
        'COPAGO':          'https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/coins.svg',
        'OBSERVACION':     'https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/align-left.svg',
    }
    
    dynamic_row = fixed_row + 1  # agrega estas variables ANTES de la función
    dynamic_col = 0

    def reflow_service_blocks():
        """Recoloca todos los bloques de servicio en orden."""
        nonlocal dynamic_row, dynamic_col
        row = fixed_row + 1
        col = 0
        for frames in service_frames:
            for frm in frames:
                frm.grid_configure(row=row, column=col)
                col += 1
                if col == 3:
                    col = 0
                    row += 1
            if col != 0:
                row += 1
                col = 0
        dynamic_row = row
        dynamic_col = col

    def add_service_block():
        nonlocal dynamic_row, dynamic_col
        dv = {}
        current_frames = []
        
        skip_obs = len(detail_vars) >= 1

        # Start a new row if the last block did not end at column 0.
        # This prevents overlaps when the number of fields isn't a
        # multiple of three.
        if dynamic_col != 0:
            dynamic_row += 1
            dynamic_col = 0


        for campo, icon_url in DETAIL_ICONS.items():
            if campo not in campos_paquete:
                continue
            
            if campo == 'OBSERVACION' and skip_obs:
                continue

            # 1) Crear y posicionar el frame del campo
            frm = make_field(campo.replace('_', ' ') + ':', icon_url)
            frm.grid(row=dynamic_row, column=dynamic_col, padx=10, pady=8, sticky='nsew')
            current_frames.append(frm)

            # 2) Variable y etiqueta de error
            default = '0' if campo == 'COPAGO' else ''
            var = tk.StringVar(master=frm, value=default)
            lbl_err = ctk.CTkLabel(frm, text='', text_color='red')
            
            # 3) Crear el widget según el tipo de campo
            if campo == 'AUTORIZACION':
                def only_digits_len(P):
                    return P == "" or (P.isdigit() and len(P) <= 9)
                vcmd_auth = (win.register(only_digits_len), '%P')

                w = ctk.CTkEntry(
                    frm, textvariable=var, width=300,
                    fg_color=entry_fg_color,
                    text_color=entry_text_color,
                    placeholder_text_color=placeholder_color,
                    placeholder_text='Solo 9 dígitos', validate='key',
                    validatecommand=vcmd_auth
                )
                w.pack(fill='x', pady=(5, 0))
                apply_focus_style(w,scroll)
                lbl_err.pack(fill='x', pady=(2, 8))

                def val_autorizacion(e=None, var=var, w=w, lbl=lbl_err):
                    txt = var.get().strip()
                    # Permitimos campo vacío
                    if not txt:
                        w.configure(border_color='#2b2b2b', border_width=1)
                        lbl.configure(text='')
                        return True
                    # Si no está vacío, debe ser exactamente 9 dígitos
                    if len(txt) != 9:
                        w.configure(border_color='red', border_width=2)
                        lbl.configure(text='Debe tener 9 dígitos')
                        return False
                    # Todo ok
                    w.configure(border_color='#2b2b2b', border_width=1)
                    lbl.configure(text='')
                    return True


                w.bind('<FocusOut>', val_autorizacion)
                dv['VALIDAR_AUTORIZACION'] = val_autorizacion

            elif campo == 'CODIGO_SERVICIO':
                def only_alphanum(P):
                    return P == "" or P.isalnum()
                vcmd_cs = (win.register(only_alphanum), '%P')

                def to_upper_and_filter(*args, var=var):
                    txt = var.get()
                    filtered = ''.join(ch for ch in txt if ch.isalnum()).upper()
                    if txt != filtered:
                        var.set(filtered)
                var.trace_add('write', to_upper_and_filter)

                w = ctk.CTkEntry(
                    frm, textvariable=var,
                    placeholder_text='CÓDIGO DE SERVICIO', width=300,
                    fg_color=entry_fg_color,
                    text_color=entry_text_color,
                    placeholder_text_color=placeholder_color,
                    validate='key', validatecommand=vcmd_cs
                )
                w.pack(fill='x', pady=(5, 0))
                apply_focus_style(w,scroll)
                w.bind('<KeyRelease>', lambda e, v=var: v.set(v.get().upper()))

                lbl_err_codigo = ctk.CTkLabel(frm, text='', text_color='red')
                lbl_err_codigo.pack(fill='x', pady=(2, 8))

                def val_codigo_servicio(e=None, var=var, w=w, lbl=lbl_err_codigo):
                    txt = var.get().strip()
                    if not txt:
                        w.configure(border_color='red', border_width=2)
                        lbl.configure(text='Código de servicio obligatorio')
                        return False
                    w.configure(border_color='#2b2b2b', border_width=1)
                    lbl.configure(text='')
                    return True

                w.bind('<FocusOut>', val_codigo_servicio)
                dv['VALIDAR_CODIGO_SERVICIO'] = val_codigo_servicio

            elif campo in ('CANTIDAD', 'VLR_UNITARIO', 'COPAGO'):
                if campo == 'CANTIDAD':
                    def only_digits_len3(P):
                        return P == "" or (P.isdigit() and len(P) <= 3)
                    vcmd_num = (win.register(only_digits_len3), '%P')
                    placeholder = '0-999'
                else:
                    def only_digits(P):
                        return P == "" or P.isdigit()
                    vcmd_num = (win.register(only_digits), '%P')
                    placeholder = default

                w = ctk.CTkEntry(
                    frm, textvariable=var, width=300,
                    fg_color=entry_fg_color,
                    text_color=entry_text_color,
                    placeholder_text_color=placeholder_color,
                    placeholder_text=placeholder,
                    validate='key', validatecommand=vcmd_num
                )
                w.pack(fill='x', pady=(5, 0))
                apply_focus_style(w,scroll)
                lbl_err.pack(fill='x', pady=(2, 8))

                def make_val_general(var, w, lbl, campo):
                    def validator(e=None):
                        txt = var.get().strip()
                        if not txt:
                            w.configure(border_color='red', border_width=2)
                            lbl.configure(text=f'{campo.replace("_", " ").title()} obligatorio')
                            return False
                        w.configure(border_color='#2b2b2b', border_width=1)
                        lbl.configure(text='')
                        return True
                    return validator

                val_func = make_val_general(var, w, lbl_err, campo)
                w.bind('<FocusOut>', val_func)
                dv[f'VALIDAR_{campo}'] = val_func

            else:
                w = ctk.CTkEntry(
                    frm, textvariable=var, width=300,
                    fg_color=entry_fg_color,
                    text_color=entry_text_color,
                    placeholder_text_color=placeholder_color,
                    placeholder_text=default
                )
                w.pack(fill='x', pady=(5, 0))
                apply_focus_style(w,scroll)

            # 4) Guardar referencia del campo
            dv[campo] = {'var': var, 'widget': w}

            # 5) Avanzar posición en el grid
            dynamic_col += 1
            if dynamic_col == 3:
                dynamic_col = 0
                dynamic_row += 1

        # 6) Añadir el set de variables y frames a las listas
        detail_vars.append(dv)
        service_frames.append(current_frames)

        if len(service_frames) > 1:
            btn_del.configure(state='normal')

        reflow_service_blocks()


    if any(c in campos_paquete for c in DETAIL_ICONS):
        add_service_block()
    
        def remove_service_block():
            nonlocal dynamic_row, dynamic_col

            # Solo eliminar si hay más de 1 (el inicial no se toca)
            if len(service_frames) <= 1:
                return

            # Eliminar el último bloque agregado dinámicamente
            last_frames = service_frames.pop()
            for f in last_frames:
                f.destroy()

            # Eliminar su conjunto de variables
            detail_vars.pop()

            # Recalcular posición del grid (esto es opcional visualmente)
            cnt = len(last_frames)
            dynamic_col -= cnt
            while dynamic_col < 0:
                dynamic_row -= 1
                dynamic_col += 3

            # Si ya no hay más servicios adicionales, desactivar botón
            if len(service_frames) <= 1:
                btn_del.configure(state='disabled')

            reflow_service_blocks()

    # -----------------------------
    # Validar y guardar en BD
    # -----------------------------
    def validate_and_save(final):
            # 1) ¿Alguna observación completada?
            any_obs = any(
                (lambda v: v.get().strip())(dv.get('OBSERVACION', {}).get('var'))
                if dv.get('OBSERVACION', {}).get('var') else False
                for dv in detail_vars
            )
            ok = True

            if not any_obs:
                # 2) Chequeo obligatorio de fecha si existe
                if 'FECHA_SERVICIO' in field_vars:
                    ok &= val_fecha()

                # 3) Campos fijos obligatorios
                for k, var in field_vars.items():
                    w = widgets[k]
                    if not var.get().strip():
                        w.configure(border_color='red', border_width=2)
                        ok = False
                    else:
                        w.configure(border_color='#2b2b2b', border_width=1)

                # 4) Validaciones específicas
                if 'TIPO_DOC_ID' in field_vars and not val_tipo():
                    ok = False
                if 'DIAGNOSTICO'  in field_vars and not val_diag():
                    ok = False

                # 5) Detalle: si no hay observación en ese bloque, validar sus campos
                for dv in detail_vars:
                    # lectura segura de OBSERVACION
                    var_obs = dv.get('OBSERVACION', {}).get('var')
                    obs = var_obs.get().strip() if var_obs else ""
                    if obs:
                        # si hay observación, salto validación de este bloque
                        continue

                    # 5b) Validar el resto de campos dinámicos
                    for campo, info in dv.items():
                        # saltar validadores y OBSERVACION
                        if not isinstance(info, dict) or campo == 'OBSERVACION':
                            continue
                        w   = info['widget']
                        val = info['var'].get().strip()

                        # AUTORIZACION puede ir vacía
                        if campo == 'AUTORIZACION':
                            # si hay texto, validamos su longitud
                            if val:
                                if len(val) != 9:
                                    w.configure(border_color='red', border_width=2)
                                    ok = False
                                else:
                                    w.configure(border_color='#2b2b2b', border_width=1)
                            else:
                                # vacío permitido
                                w.configure(border_color='#2b2b2b', border_width=1)
                            continue

                        # para los demás, son obligatorios
                        if not val:
                            w.configure(border_color='red', border_width=2)
                            ok = False
                        else:
                            w.configure(border_color='#2b2b2b', border_width=1)

            return ok

    def reset_form():
        """Limpia todos los campos del formulario y deja solo el bloque inicial"""
        nonlocal dynamic_row, dynamic_col

        for var in field_vars.values():
            var.set("")

        for dv in detail_vars:
            for info in dv.values():
                if isinstance(info, dict):
                    info['var'].set("")

        while len(service_frames) > 1:
            last_frames = service_frames.pop()
            for f in last_frames:
                f.destroy()
            detail_vars.pop()

        dynamic_row = fixed_row + 1
        dynamic_col = 0
        btn_del.configure(state='disabled')

    def do_save(final=False):
        if not validate_and_save(final):
            return

        cur2 = conn.cursor()
        asig_id = int(radicado)

        # --- 1) Preparar datos tipificación ---
        num_doc_i = int(var_num.get().strip()) if 'NUM_DOC' in field_vars and var_num.get().strip() else None
        fecha_obj = (datetime.datetime.strptime(var_fecha.get().strip(), "%d/%m/%Y").date()
                    if 'FECHA_SERVICIO' in field_vars and var_fecha.get().strip() else None)
        # TipoDoc
        if 'TIPO_DOC_ID' in field_vars and var_tipo.get().strip():
            nombre = var_tipo.get().strip().upper()
            cur2.execute(
                "SELECT ID FROM TIPO_DOC WHERE UPPER(NAME) = %s",
                (nombre,)   # ¡ojo: la coma para que sea tupla de un solo elemento!
            )
            row = cur2.fetchone()
            tipo_doc_id = row[0] if row else None
        else:
            tipo_doc_id = None

        # Diagnóstico: si está vacío, usamos None para que SQL reciba NULL
        if 'DIAGNOSTICO' in field_vars:
            raw = field_vars['DIAGNOSTICO'].get().strip().upper()
            diag_code = raw or None
        else:
            diag_code = None

        # --- 2) Insertar cabecera TIPIFICACION con USER_ID ---
        cur2.execute("""
            INSERT INTO TIPIFICACION
            (ASIGNACION_ID, FECHA_SERVICIO, TIPO_DOC_ID, NUM_DOC, DIAGNOSTICO, USER_ID)
            OUTPUT INSERTED.ID
            VALUES (%s, %s, %s, %s, %s, %s)
        """, (asig_id, fecha_obj, tipo_doc_id, num_doc_i, diag_code, current_user_id,))
        tip_id = cur2.fetchone()[0]

        # --- 3) Insertar detalles y detectar si hay observaciones ---
        tiene_obs = False
        for dv in detail_vars:
            # Leer cada campo
            auth_var = dv.get('AUTORIZACION', {}).get('var')
            auth = auth_var.get().strip() if auth_var else None
            auth = int(auth) if auth else None

           # CODIGO_SERVICIO
            cs_var = dv.get('CODIGO_SERVICIO', {}).get('var')
            cs = cs_var.get().strip().upper() if cs_var and cs_var.get().strip() else None

            # CANTIDAD
            qty_var = dv.get('CANTIDAD', {}).get('var')
            qty_raw = qty_var.get().strip() if qty_var else None
            qty = int(qty_raw) if qty_raw else None

            # VLR_UNITARIO
            valor_var = dv.get('VLR_UNITARIO', {}).get('var')
            valor_raw = valor_var.get().strip() if valor_var else None
            valor = float(valor_raw) if valor_raw else None

            # COPAGO
            copago_var = dv.get('COPAGO', {}).get('var')
            copago_raw = copago_var.get().strip() if copago_var else None
            copago = float(copago_raw) if copago_raw else None

            # OBSERVACION
            obs_var = dv.get('OBSERVACION', {}).get('var')
            obs = obs_var.get().strip() if obs_var and obs_var.get().strip() else None
            if obs:
                tiene_obs = True

            # Si todo es None, saltamos
            if all(v is None for v in (auth, cs, qty, valor, copago, obs)):
                continue

            # (Opcional) validaciones de cs...
            cur2.execute("""
                INSERT INTO TIPIFICACION_DETALLES
                (TIPIFICACION_ID, AUTORIZACION, CODIGO_SERVICIO, CANTIDAD, VLR_UNITARIO, COPAGO, OBSERVACION)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
            """, (tip_id, auth, cs, qty, valor, copago, obs,))

        conn.commit()

        # --- 4) Actualizar estado ASIGNACION_TIPIFICACION ---
        nuevo_status = 4 if tiene_obs else 3
        cur2.execute(
            "UPDATE ASIGNACION_TIPIFICACION SET STATUS_ID = %s WHERE RADICADO = %s",
            (nuevo_status, asig_id,)
        )
        conn.commit()
        cur2.close()

        # --- 5) Cerrar ventana y continuar/volver ---
        if final:
            # Desvincular eventos de fecha(s) existentes
            for campo in ('FECHA_SERVICIO', 'FECHA_SERVICIO_FINAL'):
                if campo in widgets:
                    widgets[campo].unbind("<KeyRelease>")
                    widgets[campo].unbind("<FocusOut>")
            
            # 2) Y programamos la destrucción en el idle loop,
            #    así los callbacks que ya estén en cola pueden finalizar sin error.
            win.after_idle(win.destroy)
            return
        else:
            reset_form()
            # En lugar de win.destroy + reiniciar toda la función,
            # simplemente recargamos la siguiente asignación
            if not load_assignment():
                if 'FECHA_SERVICIO' in widgets:
                    widgets['FECHA_SERVICIO'].focus_set()
                messagebox.showinfo("Sin asignaciones", "No hay más asignaciones pendientes.")
                win.destroy()



    bind_select_all(card)
    # -----------------------------
    # Botonera
    # -----------------------------
    footer = ctk.CTkFrame(card, fg_color='transparent')
    footer.pack(side='bottom', fill='x', padx=30, pady=10)

    save_img = load_icon_from_url(
        "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/save.svg", size=(18,18)
    )
    btn_save = ctk.CTkButton(
        footer,
        text="Guardar y siguiente",
        image=save_img,
        compound="left",
        fg_color="#28a745",
        hover_color="#218838",
        command=lambda: do_save(final=False)
    )
    btn_save.pack(side='left', expand=True, fill='x', padx=5)
    apply_focus_style(btn_save,scroll)

    add_img = load_icon_from_url(
        "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/plus-circle.svg", size=(18,18)
    )
    btn_add = ctk.CTkButton(
        footer,
        text="Agregar servicio",
        image=add_img,
        compound="left",
        fg_color="#17a2b8",
        hover_color="#138496",
        command=add_service_block
    )
    btn_add.pack(side='left', expand=True, fill='x', padx=5)
    apply_focus_style(btn_add,scroll)
    
    del_img = load_icon_from_url(
        "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/trash-alt.svg",
        size=(18,18)
    )
    btn_del = ctk.CTkButton(
        footer,
        text="Eliminar servicio",
        image=del_img,
        compound="left",
        fg_color="#dc3545",
        hover_color="#c82333",
        command=remove_service_block,
        state='disabled'
    )
    btn_del.pack(side='left', expand=True, fill='x', padx=5)
    apply_focus_style(btn_del,scroll)

    exit_img = load_icon_from_url(
        "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/sign-out-alt.svg", size=(18,18)
    )
    btn_exit = ctk.CTkButton(
        footer,
        text="Salir y Guardar",
        image=exit_img,
        compound="left",
        fg_color="#dc3545",
        hover_color="#c82333",
        command=lambda: do_save(final=True)
    )
    btn_exit.pack(side='left', expand=True, fill='x', padx=5)
    apply_focus_style(btn_exit,scroll)

    # Bind Enter para cada botón
    for b in (btn_save, btn_add, btn_del, btn_exit):
        b.bind("<Return>", lambda e, btn=b: btn.invoke())

def modificar_radicado(parent_root, conn, user_id):
    parent_root.withdraw()
    win = ctk.CTkToplevel(parent_root)
    win.title("Actualizar Radicado")
    win.geometry("1100x650")
    win.grab_set()

    # Colores según tema guardado
    settings = QtCore.QSettings("Procesos Y Servicios", "CapturadorDeDatos")
    tema_actual = settings.value("theme", "dark")
    if tema_actual == "dark":
        color_container  = "#1e1e1e"
        color_card       = "#2b2b2b"
        fg_text_color    = "white"
        entry_fg_color   = "#424242"  # gris oscuro
        entry_text_color = "white"
        placeholder_color= "#BBBBBB"
    else:
        color_container  = "#F8F8F8"
        color_card       = "#EAEAEA"
        fg_text_color    = "black"

        entry_fg_color   = "white"    # campos claros
        entry_text_color = "black"
        placeholder_color= "#666666"

    container = ctk.CTkFrame(win, fg_color=color_container)
    container.pack(fill="both", expand=True)
    card = ctk.CTkFrame(container, fg_color=color_card)
    card.pack(fill="both", expand=True, padx=10, pady=10)

    # Variables
    entry_radicado_var = tk.StringVar()
    entry_nit_var      = tk.StringVar()
    entry_factura_var  = tk.StringVar()

    var_fecha  = tk.StringVar()
    var_fecha_final = tk.StringVar()
    var_tipo   = tk.StringVar()
    var_num    = tk.StringVar()
    var_diag   = tk.StringVar()
    var_auth   = tk.StringVar()
    var_cs     = tk.StringVar()
    var_qty    = tk.StringVar()
    var_valor  = tk.StringVar()
    var_copago = tk.StringVar()
    var_obs    = tk.StringVar()

    field_vars = {}
    detail_vars = []

    # Columnas numéricas para conversión
    numeric_main = {"TIPO_DOC_ID", "NUM_DOC"}
    numeric_det  = {"AUTORIZACION", "CANTIDAD", "VLR_UNITARIO", "COPAGO"}

    # Formatea fecha DD/MM/AAAA
    def _format_fecha_factory(var, widget):
        def fmt(e):
            txt = re.sub(r"[^0-9]", "", var.get())[:8]
            parts = [txt[:2]]
            if len(txt) > 2: parts.append(txt[2:4])
            if len(txt) > 4: parts.append(txt[4:])
            var.set("/".join(parts))
            widget.icursor(len(var.get()))
        return fmt

    # Busca los datos y precarga el formulario
    def _buscar():
        rad_str = entry_radicado_var.get().strip()
        if not rad_str:
            messagebox.showwarning("Advertencia", "Ingresa un radicado.")
            return
        try:
            rad = int(rad_str)
        except ValueError:
            messagebox.showerror("Error", "El radicado debe ser un número.")
            return

        cur = conn.cursor()
        # 1) Paquete y permisos (también traemos TIPO_PAQUETE)
        cur.execute("""
            SELECT at.NUM_PAQUETE, at.TIPO_PAQUETE
              FROM ASIGNACION_TIPIFICACION at
              JOIN TIPIFICACION t ON t.ASIGNACION_ID = at.RADICADO
             WHERE at.RADICADO = %s
               AND t.USER_ID  = %s
               AND at.STATUS_ID IN (3,4)
        """, (rad, user_id))
        row = cur.fetchone()
        if not row:
            messagebox.showerror("Error", "No autorizado o no existe.")
            cur.close()
            return
        num_pkg, tipo_pkg = row
        is_calidad = (tipo_pkg or "").upper() == "CALIDAD"

        # 2) Campos definidos para este paquete
        cur.execute("SELECT campo FROM PAQUETE_CAMPOS WHERE num_paquete=%s", (num_pkg,))
        campos = {r[0] for r in cur.fetchall()}

        # 3) Cabecera de tipificación (incluye fecha_final)
        cur.execute("""
            SELECT t.ID,
                   at.RADICADO, at.NIT, at.FACTURA,
                   t.FECHA_SERVICIO, t.FECHA_SERVICIO_FINAL,
                   td.NAME         AS TIPO_DOC_NAME,
                   dx.CODIGO + ' - ' + dx.NOMBRE AS DIAG_NAME,
                   t.NUM_DOC
              FROM TIPIFICACION t
        JOIN ASIGNACION_TIPIFICACION at ON at.RADICADO = t.ASIGNACION_ID
        LEFT JOIN TIPO_DOC td ON td.ID = t.TIPO_DOC_ID
        LEFT JOIN TBL_CIE10 dx ON dx.CODIGO = t.DIAGNOSTICO
             WHERE t.ASIGNACION_ID = %s
               AND t.USER_ID       = %s
        """, (rad, user_id))
        cab = cur.fetchone() or [None]*9
        tipif_id = cab[0]

        # 4) DETALLES: **todos** los servicios
        cur.execute("""
            SELECT ID, AUTORIZACION, CODIGO_SERVICIO, CANTIDAD,
                   VLR_UNITARIO, COPAGO, OBSERVACION
              FROM TIPIFICACION_DETALLES
             WHERE TIPIFICACION_ID = %s
        """, (tipif_id,))
        detalles = cur.fetchall()
        cur.close()

        # 5) Cargo variables principales
        entry_radicado_var.set(str(cab[1] or rad))
        entry_nit_var.set(str(cab[2] or ""))
        entry_factura_var.set(str(cab[3] or ""))
        var_fecha.set(
            cab[4].strftime("%d/%m/%Y") if isinstance(cab[4], datetime.date) else ""
        )
        if is_calidad:
            var_fecha_final.set(
                cab[5].strftime("%d/%m/%Y") if isinstance(cab[5], datetime.date) else ""
            )
        var_tipo.set(str(cab[6] or ""))
        init_diag = str(cab[7] or "")
        var_diag.set(init_diag.split(" - ")[0] if " - " in init_diag else init_diag)
        var_num.set(str(cab[8] or ""))

        # 6) Pinto los bloques de detalle y los campos dinámicos
        _render_campos(campos, detalles, is_calidad)


    def _render_campos(campos_paquete, detalles_list, is_calidad):
        # limpio scroll y vars
        for w in scroll.winfo_children():
            w.destroy()
        field_vars.clear()
        detail_vars.clear()

        # Opciones para autocompletes
        cur = conn.cursor()
        cur.execute("SELECT NAME FROM TIPO_DOC")
        tipo_doc_opts = [r[0] for r in cur.fetchall()]
        cur.execute("SELECT CODIGO, NOMBRE FROM TBL_CIE10")
        full_diag_opts = [f"{c} - {n}" for c,n in cur.fetchall()]
        cur.close()

        # DEFINICIÓN de los campos fijos
        MAIN_DEFS = [
            ("Fecha Servicio",       var_fecha,        "FECHA_SERVICIO",       "date",         None),
        ]
        if is_calidad:
            MAIN_DEFS.append(
                ("Fecha Servicio Final", var_fecha_final,  "FECHA_SERVICIO_FINAL", "date", None)
            )
        MAIN_DEFS += [
            ("Tipo Documento",       var_tipo,         "TIPO_DOC_ID",          "autocomplete", tipo_doc_opts),
            ("Número Documento",     var_num,          "NUM_DOC",              "int",          None),
            ("Diagnóstico",          var_diag,         "DIAGNOSTICO",          "fullmatch",    full_diag_opts),
        ]

        # DEFINICIÓN de los campos de detalle (uno por servicio)
        DETAIL_DEFS = [
            ("Autorización",    "AUTORIZACION",    "int9"),
            ("Código Servicio", "CODIGO_SERVICIO", "alnum"),
            ("Cantidad",        "CANTIDAD",        "int3"),
            ("Valor Unitario",  "VLR_UNITARIO",    "int"),
            ("Copago",          "COPAGO",          "int"),
            ("Observación",     "OBSERVACION",     "text"),
        ]

        row = 0
        # 1) Pinto los campos fijos
        for label, var, key, ctype, opts in MAIN_DEFS:
            if key not in campos_paquete:
                continue
                
            ctk.CTkLabel(
                scroll,
                text=label + ":",
                font=ctk.CTkFont(weight="bold")
            ).grid(row=row, column=0, sticky="w", padx=5, pady=5)

            if ctype == "autocomplete":
                w = AutocompleteEntry(scroll, opts, textvariable=var, width=200,
                                     fg_color=entry_fg_color, text_color=entry_text_color,
                                     placeholder_text_color=placeholder_color)
                w.grid(row=row, column=1, sticky="ew", padx=5)
                var.trace_add("write", lambda *a, v=var: v.set(v.get().upper()))
            elif ctype == "fullmatch":
                w = FullMatchAutocompleteEntry(scroll, full_diag_opts, textvariable=var, width=200,
                                             fg_color=entry_fg_color, text_color=entry_text_color,
                                             placeholder_text_color=placeholder_color)
                w.grid(row=row, column=1, sticky="ew", padx=5)
                var.trace_add("write", lambda *a, v=var: v.set("".join(
                    ch for ch in v.get().upper() if ch.isalnum() or ch in (" ", "-")
                )))
                var.trace_add("write", lambda *a, v=var: v.set(
                    v.get().split(" - ")[0]
                ) if v.get() in full_diag_opts else None)
            elif ctype == "date":
                w = ctk.CTkEntry(scroll, textvariable=var, width=200, placeholder_text="DD/MM/AAAA",
                                 fg_color=entry_fg_color, text_color=entry_text_color,
                                 placeholder_text_color=placeholder_color)
                w.grid(row=row, column=1, sticky="ew", padx=5)
                w.bind("<KeyRelease>", _format_fecha_factory(var, w))
            else:
                w = ctk.CTkEntry(scroll, textvariable=var, width=200,
                                 fg_color=entry_fg_color, text_color=entry_text_color,
                                 placeholder_text_color=placeholder_color)
                w.grid(row=row, column=1, sticky="ew", padx=5)
            field_vars[key] = var
            row += 1

        # 2) Pinto un bloque por cada detalle existente
        for idx, detalle in enumerate(detalles_list, start=1):
            detalle_id, *campos_detalle = detalle
            ctk.CTkLabel(scroll, text=f"--- Servicio #{idx} ---",
                         font=("Arial", 12, "bold"), text_color=fg_text_color)\
                .grid(row=row, column=0, columnspan=2, pady=(15,5), sticky="w")
            row += 1

            dv = {'ID': detalle_id}
            for j, (label, key, ctype) in enumerate(DETAIL_DEFS):
                if key not in campos_paquete:
                    continue
                value = campos_detalle[j]
                var = tk.StringVar(value=str(value) if value is not None else "")

                ctk.CTkLabel(scroll, text=label + ":", font=ctk.CTkFont(weight="bold"))\
                    .grid(row=row, column=0, sticky="w", padx=5, pady=5)
                ent = ctk.CTkEntry(
                    scroll,
                    textvariable=var,
                    width=200,
                    fg_color=entry_fg_color,
                    text_color=entry_text_color,
                    placeholder_text_color=placeholder_color,
                    border_color=fg_text_color,
                    border_width=1,
                )

                ent.grid(row=row, column=1, sticky="ew", padx=5, pady=5)
                dv[key] = var
                row += 1

            detail_vars.append(dv)

    def _guardar():
        cur = conn.cursor()

        # 1) Actualizo cada bloque de detalle por su ID
        for dv in detail_vars:
            det_id = dv.pop('ID')
            set_parts, params = [], []
            for k, var in dv.items():
                set_parts.append(f"{k} = %s")
                val = var.get().strip()
                if k in numeric_det:
                    try:    params.append(int(val))
                    except: params.append(None)
                else:
                    params.append(val or None)
            params.append(det_id)
            cur.execute(
                f"UPDATE TIPIFICACION_DETALLES SET {', '.join(set_parts)} WHERE ID = %s",
                params
            )

        # 2) Recolecto valores principales
        updates = {k: v.get().strip() for k, v in field_vars.items()}
        # 3) Resuelvo IDs y validaciones
        tipo_name = updates.get('TIPO_DOC_ID')
        if tipo_name:
            cur.execute("SELECT ID FROM TIPO_DOC WHERE NAME = %s", (tipo_name,))
            r = cur.fetchone()
            updates['TIPO_DOC_ID'] = r[0] if r else None

        diag_code = updates.get('DIAGNOSTICO')
        if diag_code:
            cur.execute("SELECT CODIGO FROM TBL_CIE10 WHERE CODIGO = %s", (diag_code,))
            updates['DIAGNOSTICO'] = diag_code if cur.fetchone() else None

        # 4) Actualizo TIPIFICACION
        main_cols = {"FECHA_SERVICIO", "FECHA_SERVICIO_FINAL", "TIPO_DOC_ID", "NUM_DOC", "DIAGNOSTICO"}
        main_updates = {k: updates[k] for k in updates if k in main_cols}
        if main_updates:
            set_parts, params = [], []
            for k, val in main_updates.items():
                if k in ("FECHA_SERVICIO", "FECHA_SERVICIO_FINAL"):
                    params.append(datetime.datetime.strptime(val, "%d/%m/%Y").date())
                elif k in numeric_main:
                    try:    params.append(int(val))
                    except: params.append(None)
                else:
                    params.append(val or None)
                set_parts.append(f"{k} = %s")
            # modificado_por
            set_parts.append("modificado_por = %s")
            params.append(user_id)
            # filtro
            rad = int(entry_radicado_var.get())
            params += [rad, user_id]
            cur.execute(
                f"UPDATE TIPIFICACION SET {', '.join(set_parts)} "
                "WHERE ASIGNACION_ID = %s AND USER_ID = %s",
                params
            )

        conn.commit()
        cur.close()
        messagebox.showinfo("Éxito", "Radicado actualizado correctamente.")
        win.destroy()
        parent_root.deiconify()

    # Construye la UI
    frm_search = ctk.CTkFrame(card, fg_color="transparent")
    frm_search.pack(fill="x", padx=20, pady=(20,10))
    ctk.CTkLabel(frm_search, text="Buscar Radicado:", anchor="w",
                 font=ctk.CTkFont(weight="bold")).pack(fill="x")
    ctk.CTkEntry(frm_search, textvariable=entry_radicado_var).pack(side="left", fill="x", expand=True, pady=5)
    ctk.CTkButton(frm_search, text="Buscar", command=_buscar).pack(side="right", padx=(10,0))
    frm_info = ctk.CTkFrame(card, fg_color="transparent")
    frm_info.pack(fill="x", padx=20, pady=(0,10))
    for label_text, var in [("Radicado:", entry_radicado_var),("NIT:", entry_nit_var),("Factura:", entry_factura_var)]:
        cell = ctk.CTkFrame(frm_info); cell.pack(side="left", expand=True, fill="x", padx=5)
        ctk.CTkLabel(cell, text=label_text, anchor="w",
                     font=ctk.CTkFont(weight="bold")).pack(fill="x")
        ctk.CTkEntry(cell, textvariable=var, state="readonly").pack(fill="x")
    scroll = ctk.CTkScrollableFrame(win, fg_color=color_card)
    scroll.pack(fill="both", expand=True, padx=20, pady=(0,10))
    for i in range(3): scroll.grid_columnconfigure(i, weight=1, uniform="col")

    # Botón final para guardar todo
    ctk.CTkButton(win, text="Actualizar todo", command=_guardar).pack(pady=10)

def iniciar_calidad(parent_root, conn, current_user_id):
    settings = QtCore.QSettings("Procesos Y Servicios", "CapturadorDeDatos")
    tema_actual = settings.value("theme", "dark")  # “dark” o “light”
    
    if tema_actual == "dark":
        color_container  = "#1e1e1e"  # fondo general “dark”
        color_card       = "#2b2b2b"  # fondo del “card” en dark
        fg_text_color    = "white"    # todo texto va en blanco
        entry_fg_color   = "#424242"  # gris oscuro armonioso
        entry_text_color = "white"    # texto blanco en entradas
        placeholder_color= "#555555"  # gris medio
    else:
        # tema “light-blue”
        color_container   = "#F8F8F8"  # gris clarito casi blanco
        color_card        = "#EAEAEA"  # un pelín más oscuro que container
        fg_text_color    = "black"    # todo texto en negro
        entry_fg_color   = "white"    # campos blancos en tema light
        entry_text_color = "black"    # texto negro
        placeholder_color= "#AAAAAA"  # gris para placeholder
        
    entry_radicado_var = tk.StringVar()
    entry_nit_var      = tk.StringVar()
    entry_factura_var  = tk.StringVar()
    # 1) Carga paquete y campos
    cur = conn.cursor()
    cur.execute("""
        SELECT MAX(NUM_PAQUETE)
        FROM ASIGNACION_TIPIFICACION
        WHERE UPPER(LTRIM(RTRIM(TIPO_PAQUETE))) = %s AND TIPO_PAQUETE = 'CALIDAD'
    """, ("CALIDAD",))
    pkg = cur.fetchone()[0] or 0
    tipo_paquete = 'CALIDAD'
    cur.close()
    
    cur2 = conn.cursor()
    cur2.execute("""
        SELECT campo
        FROM PAQUETE_CAMPOS
        WHERE NUM_PAQUETE = %s
          AND tipo_paquete = %s
    """, (pkg, tipo_paquete))
    campos_paquete = {row[0] for row in cur2.fetchall()}
    cur2.close()

    # DECLARAMOS las variables de la asignación para todo el ámbito
    radicado = None
    nit = None
    factura = None

    # 3) Ventana principal
    win = ctk.CTkToplevel(parent_root)
    win.title(f"Capturador De Datos · Paquete {pkg}")
    
    def copy_radicado():
        win.clipboard_clear()
        win.clipboard_append(entry_radicado_var.get())
        
    def load_assignment():
        nonlocal radicado, nit, factura

        cur = conn.cursor()
        cur.execute("""
            SELECT TOP 1 RADICADO, NIT, FACTURA
            FROM ASIGNACION_TIPIFICACION
            WHERE STATUS_ID = 1
              AND NUM_PAQUETE = %s
              AND TIPO_PAQUETE = 'CALIDAD'
            ORDER BY NEWID()
        """, (pkg,))
        row = cur.fetchone()
        if not row:
            cur.close()
            return False
        radicado, nit, factura = row
        cur.execute(
            "UPDATE ASIGNACION_TIPIFICACION SET STATUS_ID = 2 WHERE RADICADO = %s",
            (radicado,)
        )
        conn.commit()
        cur.close()

        entry_radicado_var.set(str(radicado))
        entry_nit_var.set(str(nit))
        entry_factura_var.set(str(factura))
        copy_radicado()
        return True

    # 4) Llamada inicial
    if not load_assignment():
        messagebox.showinfo("Sin asignaciones", "No hay asignaciones pendientes.")
        return

    # Obtener la resolución de la pantalla
    screen_width = win.winfo_screenwidth()
    screen_height = win.winfo_screenheight()

    # Calcular el alto de la ventana para que no cubra la barra de tareas
    taskbar_height = 40  # Estimación del alto de la barra de tareas de Windows (puede variar)
    window_height = screen_height - taskbar_height  # Resta el alto de la barra de tareas

    # Establecer la geometría de la ventana para que ocupe toda la pantalla, pero sin la barra de tareas
    win.geometry(f"{screen_width}x{window_height}")

    # Calcular la posición para centrar la ventana
    center_x = (screen_width // 2) - (screen_width // 2)
    center_y = (window_height // 2) - (window_height // 2)

    # Establecer la nueva geometría centrada
    win.geometry(f"{screen_width}x{window_height}+{center_x}+{center_y}")

    win.grab_set()

    container = ctk.CTkFrame(win, fg_color=color_container)
    container.grid(row=0, column=0, sticky="nsew")
    win.grid_rowconfigure(0, weight=1)
    win.grid_columnconfigure(0, weight=1)

    card = ctk.CTkFrame(container, fg_color=color_card)
    card.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
    container.grid_rowconfigure(0, weight=1)
    container.grid_columnconfigure(0, weight=1)

    # Avatar y título
    avatar = load_icon_from_url(
        "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/user-circle.svg",
        size=(80, 80)
    )
    ctk.CTkLabel(card, image=avatar, text="").pack(pady=(20, 5))
    ctk.CTkLabel(
        card,
        text=f"📦 Paquete #{pkg}",
        font=ctk.CTkFont(size=26, weight='bold'),
        text_color=fg_text_color
    ).pack(pady=(0, 15))


   # 4) Lectura de Radicado, NIT, Factura
    read_frame = ctk.CTkFrame(card, fg_color='transparent')
    read_frame.pack(fill='x', padx=30)
    read_frame.grid_columnconfigure(1, weight=1)

    # Labels fijos
    ctk.CTkLabel(read_frame, text="Radicado:", anchor='w', text_color=fg_text_color).grid(row=0, column=0, pady=5, sticky='w')
    ctk.CTkEntry(
        read_frame,
        textvariable=entry_radicado_var,   # <-- aquí
        state='readonly',
        width=300,
        fg_color=entry_fg_color,            # fondo de la entrada según tema
        text_color=entry_text_color,        # texto de la entrada según tema
        placeholder_text_color=placeholder_color
    ).grid(row=0, column=1, pady=5, sticky='ew', padx=(10,0))

    ctk.CTkLabel(read_frame, text="NIT:", anchor='w',text_color=fg_text_color).grid(row=1, column=0, pady=5, sticky='w')
    ctk.CTkEntry(
        read_frame,
        textvariable=entry_nit_var,        # <-- y aquí
        state='readonly',
        width=300,
        fg_color=entry_fg_color,            # fondo de la entrada según tema
        text_color=entry_text_color,        # texto de la entrada según tema
        placeholder_text_color=placeholder_color
    ).grid(row=1, column=1, pady=5, sticky='ew', padx=(10,0))

    ctk.CTkLabel(read_frame, text="Factura:", anchor='w',text_color=fg_text_color).grid(row=2, column=0, pady=5, sticky='w')
    ctk.CTkEntry(
        read_frame,
        textvariable=entry_factura_var,    # <-- y aquí
        state='readonly',
        width=300,
        fg_color=entry_fg_color,            # fondo de la entrada según tema
        text_color=entry_text_color,        # texto de la entrada según tema
        placeholder_text_color=placeholder_color
    ).grid(row=2, column=1, pady=5, sticky='ew', padx=(10,0))


    # 5) Scrollable y grid de 3 columnas
    scroll = ScrollableFrame(card, fg_color=color_card)
    scroll.pack(fill='both', expand=True, padx=20, pady=(10,0))
    card.pack_propagate(False)
    card.grid_rowconfigure(1, weight=1)
    card.grid_columnconfigure(0, weight=1)
    for col in range(3):
        scroll.scrollable_frame.grid_columnconfigure(col, weight=1, uniform="col")

    # 6) Variables de posición y contenedores
    fixed_row = 0
    fixed_col = 0
    field_vars = {}
    widgets = {}
    detail_vars = []
    service_frames = []

    def place_fixed_field(frame):
        nonlocal fixed_row, fixed_col
        frame.grid(row=fixed_row, column=fixed_col, padx=10, pady=8, sticky='nsew')
        fixed_col += 1
        if fixed_col == 3:
            fixed_col = 0
            fixed_row += 1

    def on_close():
        # Si la ventana se cierra sin guardar, cambiamos el estado de la asignación a 1
        if radicado is not None:
            cur = conn.cursor()
            cur.execute(
                """
                UPDATE ASIGNACION_TIPIFICACION
                SET STATUS_ID = 1
                WHERE RADICADO = %s
                """,
                (radicado,),
            )
            conn.commit()
            cur.close()
        win.destroy()  # Cierra la ventana después de actualizar el estado

    # Configurar el evento de cierre de la ventana
    win.protocol("WM_DELETE_WINDOW", on_close)

    # 7) Funciones auxiliares de validación y selecciones
    def select_all(event):
        w = event.widget
        try:
            w.select_range(0, 'end')
            w.icursor('end')
        except: pass

    def clear_selection_on_key(event):
        w = event.widget
        if event.keysym in ('BackSpace', 'Delete'):
            if w.selection_present(): w.delete(0, 'end'); return
        ch = event.char
        if len(ch)==1 and ch.isprintable() and w.selection_present():
            w.delete(0, 'end')

    def bind_select_all(widget):
        for child in widget.winfo_children():
            if isinstance(child, ctk.CTkEntry):
                child.bind("<Double-Button-1>", select_all)
                child.bind("<FocusIn>", select_all)
                child.bind("<Key>", clear_selection_on_key)
            bind_select_all(child)

    def mark_required(w, var):
        def chk(e=None):
            if not var.get().strip():
                w.configure(border_color='red', border_width=2)
            else:
                w.configure(border_color='#2b2b2b', border_width=1)
        w.bind('<FocusOut>', chk)

    def make_field(label_text, icon_url=None):
        frame = ctk.CTkFrame(scroll.scrollable_frame, fg_color='transparent')
        if icon_url:
            ico = load_icon_from_url(icon_url, size=(20,20))
            ctk.CTkLabel(frame, image=ico, text="").pack(side='left', padx=(0,5))
        ctk.CTkLabel(
            frame,
            text=label_text,
            anchor='w',
            text_color=fg_text_color,
            font=("Arial", 12, "bold")  # etiquetas resaltadas
        ).pack(fill='x')
        return frame


    # Función para seleccionar todo el texto en el campo
    def select_all(event):
        w = event.widget
        w.select_range(0, 'end')
        return 'break'

    # Función para borrar todo al presionar cualquier tecla si hay selección
    def clear_selection_on_key(event):
        w = event.widget
        if event.keysym in ('BackSpace', 'Delete'):
            if w.selection_present(): 
                w.delete(0, 'end')
            return
        ch = event.char
        if len(ch) == 1 and ch.isprintable() and w.selection_present():
            w.delete(0, 'end')

    # Función para formatear el campo de fecha mientras el usuario escribe
    def format_fecha(event):
        txt = var_fecha.get()
        # Si es borrado o navegación, no formatear aquí
        if event.keysym in ('BackSpace', 'Delete', 'Left', 'Right', 'Home', 'End'):
            return

        # Quitamos cualquier slash existente y limitamos a 8 dígitos (DDMMYYYY)
        digits = txt.replace('/', '')[:8]

        # Reconstruimos con slashes: DD / MM / AAAA
        parts = []
        if len(digits) >= 2:
            parts.append(digits[:2])
            if len(digits) >= 4:
                parts.append(digits[2:4])
                parts.append(digits[4:])
            else:
                parts.append(digits[2:])
        else:
            parts.append(digits)

        new_text = '/'.join(parts)
        var_fecha.set(new_text)
        entry_fecha.icursor(len(new_text))  # colocamos el cursor al final

    def val_fecha(e=None):
        txt = var_fecha.get().strip()
        try:
            d = datetime.datetime.strptime(txt, '%d/%m/%Y').date()
            if d > datetime.date.today():
                raise ValueError("Fecha futura")
            entry_fecha.configure(border_color='#2b2b2b', border_width=1)
            lbl_err_fecha.configure(text='')
            return True
        except Exception:
            entry_fecha.configure(border_color='red', border_width=2)
            lbl_err_fecha.configure(text='Fecha inválida')
            return False

    def apply_focus_style(widget, scroll):
        # 1) detectamos el widget interno que realmente recibe el foco
        #    para CTkEntry es widget._entry
        real = getattr(widget, '_entry', widget)

        def on_focus_in(e):
            widget.configure(border_color='#3399FF', border_width=2)
            scroll.scroll_to_widget(widget)

        def on_focus_out(e):
            widget.configure(border_color='#2b2b2b', border_width=1)

        # 2) enlazamos los eventos al widget interno
        real.bind('<FocusIn>', on_focus_in)
        real.bind('<FocusOut>', on_focus_out)

    # ————— Bloque de creación del campo de fecha —————

    if 'FECHA_SERVICIO' in campos_paquete:
        frm = make_field(
            'Fecha Servicio:',
            'https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/calendar.svg'
        )
        var_fecha = tk.StringVar()

        entry_fecha = ctk.CTkEntry(
            frm,
            textvariable=var_fecha,
            placeholder_text='DD/MM/AAAA',
            width=300,
            fg_color=entry_fg_color,
            text_color=entry_text_color,
            placeholder_text_color=placeholder_color,
            validate='key',
            validatecommand=(win.register(lambda s: bool(re.match(r"^[0-9/]$", s))), '%S')
        )
        entry_fecha.pack(fill='x', pady=(5, 0))
        apply_focus_style(entry_fecha,scroll)

        # Selección completa en doble-click o focus
        entry_fecha.bind("<Double-Button-1>", select_all)
        entry_fecha.bind("<FocusIn>", select_all)

        # Borra todo al presionar BackSpace o Delete
        entry_fecha.bind("<Key>", clear_selection_on_key)

        # Formateo dinámico al escribir
        entry_fecha.bind("<KeyRelease>", format_fecha)

        lbl_err_fecha = ctk.CTkLabel(frm, text='', text_color='red')
        lbl_err_fecha.pack(fill='x')

        field_vars['FECHA_SERVICIO'] = var_fecha
        widgets['FECHA_SERVICIO']   = entry_fecha

        # Validación al perder foco
        entry_fecha.bind('<FocusOut>', val_fecha)

        # Posicionar en el layout
        place_fixed_field(frm)

    # ————— Bloque de creación del campo de fecha final —————
# ————— Bloque de creación del campo de fecha final —————

    if 'FECHA_SERVICIO_FINAL' in campos_paquete:
        frm_final = make_field(
            'Fecha Servicio Final:',
            'https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/calendar.svg'
        )
        var_fecha_final = tk.StringVar()

        entry_fecha_final = ctk.CTkEntry(
            frm_final,
            textvariable=var_fecha_final,
            placeholder_text='DD/MM/AAAA',
            width=300,
            fg_color=entry_fg_color,
            text_color=entry_text_color,
            placeholder_text_color=placeholder_color,
            validate='key',
            validatecommand=(win.register(lambda s: bool(re.match(r"^[0-9/]$", s))), '%S')
        )
        entry_fecha_final.pack(fill='x', pady=(5, 0))
        apply_focus_style(entry_fecha_final,scroll)

        # Función de formateo para fecha final
        def format_fecha_final(event):
            txt = var_fecha_final.get()
            if event.keysym in ('BackSpace','Delete','Left','Right','Home','End'):
                return
            digits = txt.replace('/', '')[:8]
            parts = []
            if len(digits) >= 2:
                parts.append(digits[:2])
                if len(digits) >= 4:
                    parts.append(digits[2:4])
                    parts.append(digits[4:])
                else:
                    parts.append(digits[2:])
            else:
                parts.append(digits)
            new_text = '/'.join(parts)
            var_fecha_final.set(new_text)
            entry_fecha_final.icursor(len(new_text))

        # Función de validación para fecha final
        def val_fecha_final(e=None):
            txt = var_fecha_final.get().strip()
            try:
                d = datetime.datetime.strptime(txt, '%d/%m/%Y').date()
                if d > datetime.date.today():
                    raise ValueError("Fecha futura")
                entry_fecha_final.configure(border_color='#2b2b2b', border_width=1)
                lbl_err_fecha_final.configure(text='')
                return True
            except Exception:
                entry_fecha_final.configure(border_color='red', border_width=2)
                lbl_err_fecha_final.configure(text='Fecha inválida')
                return False

        lbl_err_fecha_final = ctk.CTkLabel(frm_final, text='', text_color='red')
        lbl_err_fecha_final.pack(fill='x')

        # Bindings idénticos a los de fecha_servicio
        entry_fecha_final.bind("<Double-Button-1>", select_all)
        entry_fecha_final.bind("<FocusIn>", select_all)
        entry_fecha_final.bind("<Key>", clear_selection_on_key)
        entry_fecha_final.bind("<KeyRelease>", format_fecha_final)
        entry_fecha_final.bind("<FocusOut>", val_fecha_final)

        field_vars['FECHA_SERVICIO_FINAL'] = var_fecha_final
        widgets['FECHA_SERVICIO_FINAL']   = entry_fecha_final

        place_fixed_field(frm_final)


    if 'TIPO_DOC_ID' in campos_paquete:
        frm = make_field('Tipo Doc:',
                        'https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/id-card.svg')
        # Carga opciones
        cur_td = conn.cursor()
        cur_td.execute("SELECT NAME FROM TIPO_DOC")
        opts_td = [r[0] for r in cur_td.fetchall()]
        cur_td.close()

        var_tipo = tk.StringVar()
        # Trace: solo A–Z y uppercase
        var_tipo.trace_add('write', lambda *_: var_tipo.set(
            ''.join(ch for ch in var_tipo.get().upper() if 'A' <= ch <= 'Z')
        ))

        entry_tipo = AutocompleteEntry(frm, opts_td, width=300, textvariable=var_tipo)
        entry_tipo.pack(fill='x', pady=(5,0))
        apply_focus_style(entry_tipo,scroll)

        # Forzar mayúsculas en KeyRelease
        def to_upper_on_key(event, var=var_tipo):
            var.set(var.get().upper())
        entry_tipo.bind('<KeyRelease>', to_upper_on_key)

        # Etiqueta de error
        lbl_err_td = ctk.CTkLabel(frm, text='', text_color='red')
        lbl_err_td.pack(fill='x', pady=(2,0))

        # Validación al perder foco
        def val_tipo(e=None):
            nombre = var_tipo.get().strip().upper()

            # 1) Obligatorio
            if not nombre:
                entry_tipo.configure(border_color='red', border_width=2)
                lbl_err_td.configure(text='Tipo de documento obligatorio')
                return False

            # 2) Verificar existencia en la base
            cur_chk = conn.cursor()
            cur_chk.execute(
                "SELECT COUNT(*) FROM TIPO_DOC WHERE UPPER(NAME) = %s",
                (nombre,)
            )
            existe = cur_chk.fetchone()[0] > 0
            cur_chk.close()

            if not existe:
                entry_tipo.configure(border_color='red', border_width=2)
                lbl_err_td.configure(text='Tipo de documento no existe')
                return False

            # 3) Todo OK
            entry_tipo.configure(border_color='#2b2b2b', border_width=1)
            lbl_err_td.configure(text='')
            return True
        entry_tipo.bind('<FocusOut>', val_tipo)

        field_vars['TIPO_DOC_ID'] = var_tipo
        widgets['TIPO_DOC_ID']    = entry_tipo
        place_fixed_field(frm)


    if 'NUM_DOC' in campos_paquete:
        frm = make_field('Num Doc:',
                        'https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/hashtag.svg')
        var_num = tk.StringVar()
        entry_num = ctk.CTkEntry(
            frm, textvariable=var_num,
            placeholder_text='Solo dígitos', width=300,
            fg_color=entry_fg_color,
            text_color=entry_text_color,
            placeholder_text_color=placeholder_color,
            validate='key', validatecommand=(win.register(lambda s: s.isdigit()), '%S')
        )
        entry_num.pack(fill='x', pady=(5,0))
        apply_focus_style(entry_num,scroll)

        # Etiqueta de error
        lbl_err_num = ctk.CTkLabel(frm, text='', text_color='red')
        lbl_err_num.pack(fill='x', pady=(2,0))

        # Validación al perder foco
        def val_num(e=None):
            if not var_num.get().strip():
                entry_num.configure(border_color='red', border_width=2)
                lbl_err_num.configure(text='Número de documento obligatorio')
                return False
            else:
                entry_num.configure(border_color='#2b2b2b', border_width=1)
                lbl_err_num.configure(text='')
                return True
        entry_num.bind('<FocusOut>', val_num)

        field_vars['NUM_DOC'] = var_num
        widgets['NUM_DOC']    = entry_num
        place_fixed_field(frm)


    if 'DIAGNOSTICO' in campos_paquete:
        frm = make_field('Diagnóstico:',
                        'https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/stethoscope.svg')

        # Carga mapa CIE10
        cur_dx = conn.cursor()
        cur_dx.execute("SELECT CODIGO, NOMBRE FROM TBL_CIE10")
        dx_map = {cod: nombre for cod, nombre in cur_dx.fetchall()}
        cur_dx.close()

        opciones = [f"{cod} - {nombre}" for cod, nombre in dx_map.items()]

        var_diag = tk.StringVar()
        # Trace: filtrar y uppercase
        var_diag.trace_add('write', lambda *_: var_diag.set(
            ''.join(ch for ch in var_diag.get().upper() if ch.isalnum() or ch in (' ', '-', '_'))
        ))

        entry_diag = FullMatchAutocompleteEntry(
            frm,
            values=opciones,
            width=300,
            textvariable=var_diag
        )
        entry_diag.pack(fill='x', pady=(5,0))
        apply_focus_style(entry_diag,scroll)

        # Etiqueta de error
        lbl_err_diag = ctk.CTkLabel(frm, text='', text_color='red')
        lbl_err_diag.pack(fill='x', pady=(2,0))

        # Extraer código al seleccionar
        def on_select(event=None):
            text = var_diag.get()
            if " - " in text:
                cod, _ = text.split(" - ", 1)  # Extrae solo el código
                var_diag.set(cod)  # Establece solo el código en el campo de entrada

        # Asegurarse de que on_select siempre se dispare cuando se seleccione un ítem del desplegable
        entry_diag.bind('<<ListboxSelect>>', on_select)  # Al seleccionar un ítem del desplegable

        # También aseguramos que se ejecute al hacer "Enter" después de la selección
        entry_diag.bind('<Return>', on_select)  # Al presionar Enter (si se usa para seleccionar)

        # Asegurarse de que el valor se actualice también cuando el campo pierde el foco (cuando el usuario hace click fuera o tabula)
        def on_focus_out(event):
            on_select(event)

        entry_diag.bind('<FocusOut>', on_focus_out)  # Actualiza al cambiar de campo (FocusOut)

        # Validación al perder foco (primero on_select, luego val)
        def val_diag(e=None):
            on_select()  # extrae el código a var_diag

            codigo = var_diag.get().strip().upper()

            # 1) Obligatorio
            if not codigo:
                entry_diag.configure(border_color='red', border_width=2)
                lbl_err_diag.configure(text='Diagnóstico obligatorio')
                return False

            # 2) Verificar que el código esté en dx_map
            if codigo not in dx_map:
                entry_diag.configure(border_color='red', border_width=2)
                lbl_err_diag.configure(text='Código de diagnóstico no existe')
                return False

            # 3) Todo OK
            entry_diag.configure(border_color='#2b2b2b', border_width=1)
            lbl_err_diag.configure(text='')
            return True

        entry_diag.bind('<FocusOut>', val_diag)

        field_vars['DIAGNOSTICO'] = var_diag
        widgets['DIAGNOSTICO']    = entry_diag
        place_fixed_field(frm)


    # 9) Campos dinámicos
    DETAIL_ICONS = {
        'AUTORIZACION':    'https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/file-invoice.svg',
        'CODIGO_SERVICIO': 'https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/tools.svg',
        'CANTIDAD':        'https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/list-ol.svg',
        'VLR_UNITARIO':    'https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/dollar-sign.svg',
        'COPAGO':          'https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/coins.svg',
        'OBSERVACION':     'https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/align-left.svg',
    }
    
    dynamic_row = fixed_row + 1  # agrega estas variables ANTES de la función
    dynamic_col = 0
    def reflow_service_blocks():
        """Recoloca todos los bloques de servicio en orden."""
        nonlocal dynamic_row, dynamic_col
        row = fixed_row + 1
        col = 0
        for frames in service_frames:
            for frm in frames:
                frm.grid_configure(row=row, column=col)
                col += 1
                if col == 3:
                    col = 0
                    row += 1
            if col != 0:
                row += 1
                col = 0
        dynamic_row = row
        dynamic_col = col

    
    def add_service_block():
        nonlocal dynamic_row, dynamic_col
        dv = {}
        current_frames = []
        
        skip_obs = len(detail_vars) >= 1
        # Start a new row if the last block did not end at column 0
        if dynamic_col != 0:
            dynamic_row += 1
            dynamic_col = 0

        for campo, icon_url in DETAIL_ICONS.items():
            if campo not in campos_paquete:
                continue
            
            if campo == 'OBSERVACION' and skip_obs:
                continue

            # 1) Crear y posicionar el frame del campo
            frm = make_field(campo.replace('_', ' ') + ':', icon_url)
            frm.grid(row=dynamic_row, column=dynamic_col, padx=10, pady=8, sticky='nsew')
            current_frames.append(frm)

            # 2) Variable y etiqueta de error
            default = '0' if campo == 'COPAGO' else ''
            var = tk.StringVar(master=frm, value=default)
            lbl_err = ctk.CTkLabel(frm, text='', text_color='red')
            
            # 3) Crear el widget según el tipo de campo
            if campo == 'AUTORIZACION':
                def only_digits_len(P):
                    return P == "" or (P.isdigit() and len(P) <= 9)
                vcmd_auth = (win.register(only_digits_len), '%P')

                w = ctk.CTkEntry(
                    frm, textvariable=var, width=300,
                    fg_color=entry_fg_color,
                    text_color=entry_text_color,
                    placeholder_text_color=placeholder_color,
                    placeholder_text='Solo 9 dígitos', validate='key',
                    validatecommand=vcmd_auth
                )
                w.pack(fill='x', pady=(5, 0))
                apply_focus_style(w,scroll)
                lbl_err.pack(fill='x', pady=(2, 8))

                def val_autorizacion(e=None, var=var, w=w, lbl=lbl_err):
                    txt = var.get().strip()
                    if len(txt) != 9:
                        w.configure(border_color='red', border_width=2)
                        lbl.configure(text='Debe tener 9 dígitos')
                        return False
                    w.configure(border_color='#2b2b2b', border_width=1)
                    lbl.configure(text='')
                    return True

                w.bind('<FocusOut>', val_autorizacion)
                dv['VALIDAR_AUTORIZACION'] = val_autorizacion

            elif campo == 'CODIGO_SERVICIO':
                def only_alphanum(P):
                    return P == "" or P.isalnum()
                vcmd_cs = (win.register(only_alphanum), '%P')

                def to_upper_and_filter(*args, var=var):
                    txt = var.get()
                    filtered = ''.join(ch for ch in txt if ch.isalnum()).upper()
                    if txt != filtered:
                        var.set(filtered)
                var.trace_add('write', to_upper_and_filter)

                w = ctk.CTkEntry(
                    frm, textvariable=var,
                    placeholder_text='CÓDIGO DE SERVICIO', width=300,
                    fg_color=entry_fg_color,
                    text_color=entry_text_color,
                    placeholder_text_color=placeholder_color,
                    validate='key', validatecommand=vcmd_cs
                )
                w.pack(fill='x', pady=(5, 0))
                apply_focus_style(w,scroll)
                w.bind('<KeyRelease>', lambda e, v=var: v.set(v.get().upper()))

                lbl_err_codigo = ctk.CTkLabel(frm, text='', text_color='red')
                lbl_err_codigo.pack(fill='x', pady=(2, 8))

                def val_codigo_servicio(e=None, var=var, w=w, lbl=lbl_err_codigo):
                    txt = var.get().strip()
                    if not txt:
                        w.configure(border_color='red', border_width=2)
                        lbl.configure(text='Código de servicio obligatorio')
                        return False
                    w.configure(border_color='#2b2b2b', border_width=1)
                    lbl.configure(text='')
                    return True

                w.bind('<FocusOut>', val_codigo_servicio)
                dv['VALIDAR_CODIGO_SERVICIO'] = val_codigo_servicio

            elif campo in ('CANTIDAD', 'VLR_UNITARIO', 'COPAGO'):
                if campo == 'CANTIDAD':
                    def only_digits_len3(P):
                        return P == "" or (P.isdigit() and len(P) <= 3)
                    vcmd_num = (win.register(only_digits_len3), '%P')
                    placeholder = '0-999'
                else:
                    def only_digits(P):
                        return P == "" or P.isdigit()
                    vcmd_num = (win.register(only_digits), '%P')
                    placeholder = default

                w = ctk.CTkEntry(
                    frm, textvariable=var, width=300,
                    fg_color=entry_fg_color,
                    text_color=entry_text_color,
                    placeholder_text_color=placeholder_color,
                    placeholder_text=placeholder,
                    validate='key', validatecommand=vcmd_num
                )
                w.pack(fill='x', pady=(5, 0))
                apply_focus_style(w,scroll)
                lbl_err.pack(fill='x', pady=(2, 8))

                def make_val_general(var, w, lbl, campo):
                    def validator(e=None):
                        txt = var.get().strip()
                        if not txt:
                            w.configure(border_color='red', border_width=2)
                            lbl.configure(text=f'{campo.replace("_", " ").title()} obligatorio')
                            return False
                        w.configure(border_color='#2b2b2b', border_width=1)
                        lbl.configure(text='')
                        return True
                    return validator

                val_func = make_val_general(var, w, lbl_err, campo)
                w.bind('<FocusOut>', val_func)
                dv[f'VALIDAR_{campo}'] = val_func

            else:
                w = ctk.CTkEntry(
                    frm, textvariable=var, width=300,
                    fg_color=entry_fg_color,
                    text_color=entry_text_color,
                    placeholder_text_color=placeholder_color,
                    placeholder_text=default
                )
                w.pack(fill='x', pady=(5, 0))
                apply_focus_style(w,scroll)

            # 4) Guardar referencia del campo
            dv[campo] = {'var': var, 'widget': w}

            # 5) Avanzar posición en el grid
            dynamic_col += 1
            if dynamic_col == 3:
                dynamic_col = 0
                dynamic_row += 1

        # 6) Añadir el set de variables y frames a las listas
        detail_vars.append(dv)
        service_frames.append(current_frames)

        if len(service_frames) > 1:
            btn_del.configure(state='normal')

        reflow_service_blocks()

    if any(c in campos_paquete for c in DETAIL_ICONS):
        add_service_block()
    
        def remove_service_block():
            
            nonlocal dynamic_row, dynamic_col

            if len(service_frames) <= 1:
                return  # nada que eliminar

            # Sacamos el último bloque de frames y lo destruimos
            last_frames = service_frames.pop()
            for f in last_frames:
                f.destroy()

            # También eliminamos sus datos de detail_vars
            detail_vars.pop()

            # Ajustamos dynamic_row/col para volver a esa posición
            cnt = len(last_frames)
            dynamic_col -= cnt
            # si dinamyc_col queda <0, retrocedemos fila
            while dynamic_col < 0:
                dynamic_row -= 1
                dynamic_col += 3

            if len(service_frames) <= 1:
                btn_del.configure(state='disabled')

            reflow_service_blocks()

    # -----------------------------
    # Validar y guardar en BD
    # -----------------------------
    def validate_and_save(final):
        # 1) ¿Alguna observación completada?
        any_obs = any(
            (lambda v: v.get().strip())(dv.get('OBSERVACION', {}).get('var'))
            if dv.get('OBSERVACION', {}).get('var') else False
            for dv in detail_vars
        )
        ok = True

        if not any_obs:
            # Chequeo obligatorio de fecha si existe
            if 'FECHA_SERVICIO' in field_vars:
                ok &= val_fecha()

            # Campos fijos obligatorios
            for k, v in field_vars.items():
                w = widgets[k]
                if not v.get().strip():
                    w.configure(border_color='red', border_width=2)
                    ok = False
                else:
                    w.configure(border_color='#2b2b2b', border_width=1)
                    
            if 'TIPO_DOC_ID' in field_vars:
                # val_tipo() retorna False si no existe en la tabla
                if not val_tipo():
                    ok = False
            if 'DIAGNOSTICO' in field_vars:
                if not val_diag():
                    ok = False

            # Detalle: si no hay observación en ese bloque, validar sus campos
            for dv in detail_vars:
                
                var_obs = dv.get('OBSERVACION', {}).get('var')
                obs = var_obs.get().strip() if var_obs else ""
                if obs:
                    # Si éste bloque tiene observación, salta validación de sus campos
                    continue


                # Resto de campos del detalle: solo los que almacenan dict {'var','widget'}
                for campo, info in dv.items():
                    # saltar validadores y OBSERVACION
                    if not isinstance(info, dict) or campo == 'OBSERVACION':
                         continue
                    w   = info['widget']
                    val = info['var'].get().strip()

                    # AUTORIZACION puede ir vacía
                    if campo == 'AUTORIZACION':
                        # si hay texto, validamos su longitud
                        if val:
                            if len(val) != 9:
                                w.configure(border_color='red', border_width=2)
                                ok = False
                            else:
                                w.configure(border_color='#2b2b2b', border_width=1)
                        else:
                            # vacío permitido
                            w.configure(border_color='#2b2b2b', border_width=1)
                        continue

                    # para los demás, son obligatorios
                    if not val:
                        w.configure(border_color='red', border_width=2)
                        ok = False
                    else:
                        w.configure(border_color='#2b2b2b', border_width=1)

        return ok

    def reset_form():
        """Limpia el formulario para la siguiente asignación"""
        nonlocal dynamic_row, dynamic_col

        for var in field_vars.values():
            var.set("")

        for dv in detail_vars:
            for info in dv.values():
                if isinstance(info, dict):
                    info['var'].set("")

        while len(service_frames) > 1:
            last_frames = service_frames.pop()
            for f in last_frames:
                f.destroy()
            detail_vars.pop()

        dynamic_row = fixed_row + 1
        dynamic_col = 0
        btn_del.configure(state='disabled')

    def do_save(final=False):
        if not validate_and_save(final):
            return

        cur2 = conn.cursor()
        asig_id = int(radicado)

        # --- 1) Preparar datos tipificación ---
        num_doc_i = int(var_num.get().strip()) if 'NUM_DOC' in field_vars and var_num.get().strip() else None
        fecha_obj = (datetime.datetime.strptime(var_fecha.get().strip(), "%d/%m/%Y").date()
                    if 'FECHA_SERVICIO' in field_vars and var_fecha.get().strip() else None)
        fecha_final_obj = (datetime.datetime.strptime(var_fecha_final.get().strip(), "%d/%m/%Y").date()
                    if 'FECHA_SERVICIO_FINAL' in field_vars and var_fecha_final.get().strip() else None)
        # TipoDoc
        if 'TIPO_DOC_ID' in field_vars and var_tipo.get().strip():
            nombre = var_tipo.get().strip().upper()
            cur2.execute(
                "SELECT ID FROM TIPO_DOC WHERE UPPER(NAME) = %s",
                (nombre,)   # ¡ojo: la coma para que sea tupla de un solo elemento!
            )
            row = cur2.fetchone()
            tipo_doc_id = row[0] if row else None
        else:
            tipo_doc_id = None

        # Diagnóstico: si está vacío, usamos None para que SQL reciba NULL
        if 'DIAGNOSTICO' in field_vars:
            raw = field_vars['DIAGNOSTICO'].get().strip().upper()
            diag_code = raw or None
        else:
            diag_code = None

        # --- 2) Insertar cabecera TIPIFICACION con USER_ID ---
        cur2.execute("""
            INSERT INTO TIPIFICACION
            (ASIGNACION_ID, FECHA_SERVICIO, FECHA_SERVICIO_FINAL, TIPO_DOC_ID, NUM_DOC, DIAGNOSTICO, USER_ID)
            OUTPUT INSERTED.ID
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        """, (asig_id, fecha_obj, fecha_final_obj, tipo_doc_id, num_doc_i, diag_code, current_user_id,))
        tip_id = cur2.fetchone()[0]

        # --- 3) Insertar detalles y detectar si hay observaciones ---
        tiene_obs = False
        for dv in detail_vars:
            # Leer cada campo
            auth_var = dv.get('AUTORIZACION', {}).get('var')
            auth = auth_var.get().strip() if auth_var else None
            auth = int(auth) if auth else None

           # CODIGO_SERVICIO
            cs_var = dv.get('CODIGO_SERVICIO', {}).get('var')
            cs = cs_var.get().strip().upper() if cs_var and cs_var.get().strip() else None

            # CANTIDAD
            qty_var = dv.get('CANTIDAD', {}).get('var')
            qty_raw = qty_var.get().strip() if qty_var else None
            qty = int(qty_raw) if qty_raw else None

            # VLR_UNITARIO
            valor_var = dv.get('VLR_UNITARIO', {}).get('var')
            valor_raw = valor_var.get().strip() if valor_var else None
            valor = float(valor_raw) if valor_raw else None

            # COPAGO
            copago_var = dv.get('COPAGO', {}).get('var')
            copago_raw = copago_var.get().strip() if copago_var else None
            copago = float(copago_raw) if copago_raw else None

            # OBSERVACION
            obs_var = dv.get('OBSERVACION', {}).get('var')
            obs = obs_var.get().strip() if obs_var and obs_var.get().strip() else None
            if obs:
                tiene_obs = True

            # Si todo es None, saltamos
            if all(v is None for v in (auth, cs, qty, valor, copago, obs)):
                continue

            # (Opcional) validaciones de cs...
            cur2.execute("""
                INSERT INTO TIPIFICACION_DETALLES
                (TIPIFICACION_ID, AUTORIZACION, CODIGO_SERVICIO, CANTIDAD, VLR_UNITARIO, COPAGO, OBSERVACION)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
            """, (tip_id, auth, cs, qty, valor, copago, obs,))

        conn.commit()

        # --- 4) Actualizar estado ASIGNACION_TIPIFICACION ---
        nuevo_status = 4 if tiene_obs else 3
        cur2.execute(
            "UPDATE ASIGNACION_TIPIFICACION SET STATUS_ID = %s WHERE RADICADO = %s",
            (nuevo_status, asig_id,)
        )
        conn.commit()
        cur2.close()

        # --- 5) Cerrar ventana y continuar/volver ---
        if final:
            # Desvincular eventos de fecha(s) existentes
            for campo in ('FECHA_SERVICIO', 'FECHA_SERVICIO_FINAL'):
                if campo in widgets:
                    widgets[campo].unbind("<KeyRelease>")
                    widgets[campo].unbind("<FocusOut>")
            
            # 2) Y programamos la destrucción en el idle loop,
            #    así los callbacks que ya estén en cola pueden finalizar sin error.
            win.after_idle(win.destroy)
            return

        else:
            reset_form()
            # En lugar de win.destroy + reiniciar toda la función,
            # simplemente recargamos la siguiente asignación
            if not load_assignment():
                if 'FECHA_SERVICIO' in widgets:
                    widgets['FECHA_SERVICIO'].focus_set()
                messagebox.showinfo("Sin asignaciones", "No hay más asignaciones pendientes.")
                win.destroy()


    bind_select_all(card)
    
    def remove_service_block():
        # se redefinirá más abajo si hay bloques dinámicos
        pass
    # -----------------------------
    # Botonera
    # -----------------------------
    footer = ctk.CTkFrame(card, fg_color='transparent')
    footer.pack(side='bottom', fill='x', padx=30, pady=10)

    save_img = load_icon_from_url(
        "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/save.svg", size=(18,18)
    )
    btn_save = ctk.CTkButton(
        footer,
        text="Guardar y siguiente",
        image=save_img,
        compound="left",
        fg_color="#28a745",
        hover_color="#218838",
        command=lambda: do_save(final=False)
    )
    btn_save.pack(side='left', expand=True, fill='x', padx=5)
    apply_focus_style(btn_save,scroll)

    add_img = load_icon_from_url(
        "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/plus-circle.svg", size=(18,18)
    )
    btn_add = ctk.CTkButton(
        footer,
        text="Agregar servicio",
        image=add_img,
        compound="left",
        fg_color="#17a2b8",
        hover_color="#138496",
        command=add_service_block
    )
    btn_add.pack(side='left', expand=True, fill='x', padx=5)
    apply_focus_style(btn_add,scroll)
    
    del_img = load_icon_from_url(
        "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/trash-alt.svg",
        size=(18,18)
    )
    btn_del = ctk.CTkButton(
        footer,
        text="Eliminar servicio",
        image=del_img,
        compound="left",
        fg_color="#dc3545",
        hover_color="#c82333",
        command=remove_service_block,
        state='disabled'
    )
    btn_del.pack(side='left', expand=True, fill='x', padx=5)
    apply_focus_style(btn_del,scroll)

    exit_img = load_icon_from_url(
        "https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free/svgs/solid/sign-out-alt.svg", size=(18,18)
    )
    btn_exit = ctk.CTkButton(
        footer,
        text="Salir y Guardar",
        image=exit_img,
        compound="left",
        fg_color="#dc3545",
        hover_color="#c82333",
        command=lambda: do_save(final=True)
    )
    btn_exit.pack(side='left', expand=True, fill='x', padx=5)
    apply_focus_style(btn_exit,scroll)

    # Bind Enter para cada botón
    for b in (btn_save, btn_add, btn_del, btn_exit):
        b.bind("<Return>", lambda e, btn=b: btn.invoke())
        
def ver_progreso(root, conn):
    # — Auxiliar para parsear fechas de texto —
    def parse_fecha(s):
        try:
            return datetime.datetime.strptime(s, "%d/%m/%Y").date()
        except:
            return None

    # — Construye WHERE y parámetros según filtros de UI —
    def construir_filtros():
        try:
            pkg = int(pkg_var.get())
        except ValueError:
            messagebox.showwarning(
                "Selección inválida", "Selecciona un paquete válido."
            )
            return None, None, None, None

        filtros = ["a.NUM_PAQUETE = %s"]
        params = [pkg]

        tipo = var_tipo_paquete.get().strip()
        if tipo:
            filtros.append("a.TIPO_PAQUETE = %s")
            params.append(tipo)

        if var_fecha_desde.get().strip():
            d1 = fecha_desde.get_date()
            filtros.append("CONVERT(date, t.fecha_creacion, 103) >= %s")
            params.append(d1)
        if var_fecha_hasta.get().strip():
            d2 = fecha_hasta.get_date()
            filtros.append("CONVERT(date, t.fecha_creacion, 103) <= %s")
            params.append(d2)

        sel_est = [e for e, v in estado_vars.items() if v.get()]
        if 0 < len(sel_est) < len(estados):
            ph = ", ".join("%s" for _ in sel_est)
            filtros.append(f"s.NAME IN ({ph})")
            params.extend(sel_est)

        sel_usr = [u for u, v in user_vars.items() if v.get()]
        if 0 < len(sel_usr) < len(usuarios):
            ph = ", ".join("%s" for _ in sel_usr)
            filtros.append(f"(u.FIRST_NAME + ' ' + u.LAST_NAME) IN ({ph})")
            params.extend(sel_usr)

        raw = rad_text.get("0.0", "end").strip()
        if raw:
            # cada línea un radicado
            try:
                lista = [int(r.strip()) for r in raw.splitlines() if r.strip()]
            except ValueError:
                messagebox.showwarning("Radicados inválidos", "Revisa la lista de radicados.")
                return None, None
            if lista:
                ph = ", ".join("%s" for _ in lista)
                filtros.append(f"a.RADICADO IN ({ph})")
                params.extend(lista)

        where_clause = " AND ".join(filtros)
        return where_clause, tuple(params), pkg, tipo
    

    # — Filtrar/mostrar solo checks de estado coincidentes —
    def _filtrar_est(event=None):
        term = buscar_est.get().lower()
        for est in estados:
            cb = estado_checks[est]
            cb.pack_forget()
            if term in est.lower():
                cb.pack(anchor="w", pady=2)

    # — Marcar/desmarcar todos los estados —
    def _marcar_est(val):
        for var in estado_vars.values():
            var.set(val)

    # — Seleccionar solo los estados actualmente visibles —
    def _solo_visibles_est():
        for est, cb in estado_checks.items():
            estado_vars[est].set(cb.winfo_ismapped())
        actualizar_tabs()

    # — Filtrar/mostrar solo checks de usuario coincidentes —
    def _filtrar_usr(event=None):
        term = buscar_usr.get().lower()
        for usr in usuarios:
            cb = user_checks[usr]
            cb.pack_forget()
            if term in usr.lower():
                cb.pack(anchor="w", pady=2)

    # — Marcar/desmarcar todos los usuarios —
    def _marcar_usr(val):
        for var in user_vars.values():
            var.set(val)

    # — Seleccionar solo los usuarios actualmente visibles —
    def _solo_visibles_usr():
        for usr, cb in user_checks.items():
            user_vars[usr].set(cb.winfo_ismapped())
        actualizar_tabs()

    # — Carga datos en las dos pestañas según filtros —
    import datetime  # asegúrate de tener esto al inicio de tu módulo

    def actualizar_tabs():
        # 0) Reconstruye filtros y params
        where, params, pkg_sel, tipo_sel = construir_filtros()
        if where is None:
            return

        # — Pestaña "Por Estado" —
        frame1 = tabs.tab("Por Estado")
        for w in frame1.winfo_children():
            w.destroy()

        cur = conn.cursor()
        sql1 = (
            "SELECT UPPER(s.NAME) AS ESTADO, COUNT(*) AS CNT "
            "FROM ASIGNACION_TIPIFICACION a "
            "JOIN STATUS s ON a.STATUS_ID = s.ID "
            "JOIN TIPIFICACION t ON t.ASIGNACION_ID = a.RADICADO "
            "JOIN USERS u ON t.USER_ID = u.ID "
            f"WHERE {where} GROUP BY s.NAME ORDER BY s.NAME"
        )
        cur.execute(sql1, params)
        rows1 = cur.fetchall()
        cur.close()

        # Dibuja cada estado y su conteo
        for i, (est, cnt) in enumerate(rows1):
            ctk.CTkLabel(frame1, text=est).grid(row=i, column=0, sticky="w", padx=5, pady=2)
            ctk.CTkLabel(frame1, text=cnt).grid(row=i, column=1, sticky="e", padx=5, pady=2)

        # --- Totales generales ---
        cur2 = conn.cursor()
        sql_tot_asig = (
            "SELECT COUNT(*) "
            "FROM ASIGNACION_TIPIFICACION a "
            "LEFT JOIN TIPIFICACION t ON t.ASIGNACION_ID = a.RADICADO "
            "LEFT JOIN USERS u        ON t.USER_ID        = u.ID "
            "LEFT JOIN STATUS s       ON a.STATUS_ID      = s.ID "
            f"WHERE {where} AND a.STATUS_ID <> 1"
        )
        cur2.execute(sql_tot_asig, params)
        total_asignados = cur2.fetchone()[0] or 0
        cur2.close()

        cur3 = conn.cursor()
        sql_tot_all = (
            "SELECT COUNT(*) "
            "FROM ASIGNACION_TIPIFICACION a "
            "LEFT JOIN TIPIFICACION t ON t.ASIGNACION_ID = a.RADICADO "
            "LEFT JOIN USERS u        ON t.USER_ID        = u.ID "
            "LEFT JOIN STATUS s       ON a.STATUS_ID      = s.ID "
            f"WHERE {where}"
        )
        cur3.execute(sql_tot_all, params)
        total_global = cur3.fetchone()[0] or 0
        cur3.close()

        fila_final = len(rows1)
        ctk.CTkLabel(frame1, text="TOTAL ASIGNADOS", font=("Arial", 12, "bold")) \
            .grid(row=fila_final, column=0, sticky="w", padx=5, pady=4)
        ctk.CTkLabel(frame1, text=str(total_asignados), font=("Arial", 12, "bold")) \
            .grid(row=fila_final, column=1, sticky="e", padx=5, pady=4)

        ctk.CTkLabel(frame1, text="TOTAL GENERAL", font=("Arial", 12, "bold")) \
            .grid(row=fila_final + 1, column=0, sticky="w", padx=5, pady=4)
        ctk.CTkLabel(frame1, text=str(total_global), font=("Arial", 12, "bold")) \
            .grid(row=fila_final + 1, column=1, sticky="e", padx=5, pady=4)


        # — Cálculo del intervalo promedio general entre tipificaciones —
        cur_int = conn.cursor()
        sql_int = f"""
            SELECT AVG(dif) FROM (
                SELECT DATEDIFF(SECOND,
                    LAG(t.fecha_creacion) OVER (ORDER BY t.fecha_creacion),
                    t.fecha_creacion
                ) AS dif
                FROM ASIGNACION_TIPIFICACION a
                JOIN TIPIFICACION t ON t.ASIGNACION_ID = a.RADICADO
                JOIN USERS u ON t.USER_ID = u.ID
                JOIN STATUS s ON a.STATUS_ID = s.ID
                WHERE {where}
                    -- excluir los que son exactamente medianoche
                    AND t.fecha_creacion <> CAST(t.fecha_creacion AS date)
            ) sub
            WHERE dif IS NOT NULL
        """
        cur_int.execute(sql_int, params)
        avg_int_sec = cur_int.fetchone()[0] or 0
        cur_int.close()

        avg_int_td = datetime.timedelta(seconds=int(avg_int_sec))
        ctk.CTkLabel(frame1,
            text=f"Intervalo promedio general: {avg_int_td}",
            font=("Arial", 10, "italic")
        ).grid(row=fila_final + 2, column=0, columnspan=2, sticky="w", padx=5, pady=4)


        # — Pestaña "Por Usuario" —
        frame2 = tabs.tab("Por Usuario")
        for w in frame2.winfo_children():
            w.destroy()

        # 1) Conteos por usuario
        cur3 = conn.cursor()
        sql2 = (
            "SELECT u.ID, "
            "       u.FIRST_NAME + ' ' + u.LAST_NAME AS USUARIO, "
            "       SUM(CASE WHEN a.STATUS_ID=2 THEN 1 ELSE 0 END) AS PENDIENTES, "
            "       SUM(CASE WHEN a.STATUS_ID=3 THEN 1 ELSE 0 END) AS PROCESADOS, "
            "       SUM(CASE WHEN a.STATUS_ID=4 THEN 1 ELSE 0 END) AS CON_OBS "
            "FROM ASIGNACION_TIPIFICACION a "
            "JOIN TIPIFICACION t ON t.ASIGNACION_ID = a.RADICADO "
            "JOIN USERS u ON t.USER_ID = u.ID "
            "JOIN STATUS s ON a.STATUS_ID = s.ID "
            f"WHERE {where} GROUP BY u.ID, u.FIRST_NAME, u.LAST_NAME ORDER BY USUARIO"
        )
        cur3.execute(sql2, params)
        rows2 = cur3.fetchall()
        cur3.close()

        # 2) Intervalo promedio entre tipificaciones por usuario
        cur_int_u = conn.cursor()
        sql_int_user = f"""
           SELECT user_id,
               AVG(dif) AS AVG_SEC
           FROM (
               SELECT 
                   u.ID      AS user_id,
                   DATEDIFF(SECOND,
                       LAG(t.fecha_creacion) OVER (
                           PARTITION BY u.ID 
                           ORDER BY t.fecha_creacion
                       ),
                       t.fecha_creacion
                   ) AS dif
               FROM ASIGNACION_TIPIFICACION a
               JOIN TIPIFICACION t ON t.ASIGNACION_ID = a.RADICADO
               JOIN USERS u       ON t.USER_ID        = u.ID
               JOIN STATUS s      ON a.STATUS_ID      = s.ID
               WHERE {where}
                 -- excluir los que son exactamente medianoche
                 AND t.fecha_creacion <> CAST(t.fecha_creacion AS date)
           ) sub2
           WHERE dif IS NOT NULL
           GROUP BY user_id
       """
        cur_int_u.execute(sql_int_user, params)
        rows_int_user = cur_int_u.fetchall()
        cur_int_u.close()
        # convertir a dict {user_id: avg_sec}
        avg_by_user = {uid: sec for uid, sec in rows_int_user}

        # Construir lista final de filas
        processed = []
        for id_, usuario, pendientes, procesados, con_obs in rows2:
            hechos = procesados + con_obs
            avg_sec_user = avg_by_user.get(id_, 0)
            td_user = datetime.timedelta(seconds=int(avg_sec_user))
            processed.append((id_, usuario, pendientes, procesados, con_obs, hechos, str(td_user)))

        headers = ["ID", "USUARIO", "PENDIENTES", "PROCESADOS", "CON_OBS", "TOTAL", "INTERVALO"]
        rows_per_page = 10
        total_general = sum(r[5] for r in processed)
        total_pages = max(1, math.ceil(len(processed) / rows_per_page))

        table_frame = ctk.CTkFrame(frame2, fg_color="transparent")
        table_frame.pack(fill="both", expand=True)

        nav_frame = ctk.CTkFrame(frame2, fg_color="transparent")
        nav_frame.pack(pady=4)

        page_var = tk.IntVar(value=0)

        def render_page():
            for w in table_frame.winfo_children():
                w.destroy()
            p = page_var.get()
            start = p * rows_per_page
            subset = processed[start : start + rows_per_page]
            for j, h in enumerate(headers):
                ctk.CTkLabel(table_frame, text=h, font=("Arial", 12, "bold"))\
                    .grid(row=0, column=j, padx=5, pady=4, sticky="w")
            for i, row in enumerate(subset, start=1):
                for j, val in enumerate(row):
                    ctk.CTkLabel(table_frame, text=str(val))\
                        .grid(row=i, column=j, padx=5, pady=2, sticky="w")

        def go_prev():
            if page_var.get() > 0:
                page_var.set(page_var.get() - 1)
                render_page()
                lbl_page.configure(text=f"{page_var.get()+1}/{total_pages}")

        def go_next():
            if page_var.get() < total_pages - 1:
                page_var.set(page_var.get() + 1)
                render_page()
                lbl_page.configure(text=f"{page_var.get()+1}/{total_pages}")

        btn_prev = ctk.CTkButton(nav_frame, text="<", width=30, command=go_prev)
        btn_prev.pack(side="left", padx=5)
        lbl_page = ctk.CTkLabel(nav_frame, text=f"1/{total_pages}")
        lbl_page.pack(side="left", padx=5)
        btn_next = ctk.CTkButton(nav_frame, text=">", width=30, command=go_next)
        btn_next.pack(side="left", padx=5)

        render_page()

        ctk.CTkLabel(frame2, text="TOTAL GENERAL", font=("Arial", 12, "bold"))\
            .pack(pady=(10,0))
        ctk.CTkLabel(frame2, text=str(total_general), font=("Arial", 12, "bold"))\
            .pack()



        
    def exportar_excel(path, headers, rows):
        # 1) Crear DataFrame
        df = pd.DataFrame(rows, columns=headers)

        # 2) Escribir con XlsxWriter
        with pd.ExcelWriter(path, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Datos')
            workbook  = writer.book
            worksheet = writer.sheets['Datos']

            # 3) Definir formatos
            header_fmt = workbook.add_format({
                'bold': True,
                'border': 2,             # borde grueso
                'align': 'center',       # centrado horizontal
                'valign': 'vcenter',     # centrado vertical
                'text_wrap': True        # Ajusta el texto
            })
            data_fmt = workbook.add_format({
                'border': 1,             # borde fino
                'align': 'center',
                'valign': 'vcenter',
                'text_wrap': True        # Ajusta el texto
            })

            # 4) Ajustar anchos y aplicar formatos
            for col_num, column in enumerate(df.columns):
                # ancho automático basado en contenido
                max_len = max(
                    df[column].astype(str).map(len).max(),
                    len(column)
                ) + 2
                # set_column aplica data_fmt a todas las celdas de la columna
                worksheet.set_column(col_num, col_num, max_len, data_fmt)
                # reescribimos el encabezado con header_fmt
                worksheet.write(0, col_num, column, header_fmt)

            # 5) (Opcional) Filtros automáticos
            worksheet.autofilter(0, 0, len(df), len(df.columns)-1)

            
    def exportar():
        where, params, _, _ = construir_filtros()
        if where is None:
            return

        # Detectamos tipo de paquete y campos configurados
        tipo = var_tipo_paquete.get().strip().upper()
        cur_f = conn.cursor()
        cur_f.execute(
            "SELECT campo FROM PAQUETE_CAMPOS WHERE NUM_PAQUETE=%s AND tipo_paquete=%s",
            (int(pkg_var.get()), tipo)
        )
        campos_set = {r[0] for r in cur_f.fetchall()}
        cur_f.close()

        # 1) Definimos dinámicamente las columnas a SELECT
        select_cols = ["a.RADICADO"]
        if "FECHA_SERVICIO" in campos_set:
            select_cols.append("t.FECHA_SERVICIO AS FECHA_SERVICIO")
        if "FECHA_SERVICIO_FINAL" in campos_set:
            select_cols.append("t.FECHA_SERVICIO_FINAL AS FECHA_SERVICIO_FINAL")
        if "AUTORIZACION" in campos_set:
            select_cols.append("d.AUTORIZACION")
        if "CODIGO_SERVICIO" in campos_set:
            select_cols.append("d.CODIGO_SERVICIO")
        if "CANTIDAD" in campos_set:
            select_cols.append("d.CANTIDAD")
        if "VLR_UNITARIO" in campos_set:
            select_cols.append("d.VLR_UNITARIO")
        if "DIAGNOSTICO" in campos_set:
            select_cols.append("t.DIAGNOSTICO")
        if "TIPO_DOC_ID" in campos_set:
            select_cols.append("COALESCE(td.NAME, '') AS TipoDocumento")
        if "NUM_DOC" in campos_set:
            select_cols.append("t.NUM_DOC AS NumeroDocumento")
        if "COPAGO" in campos_set:
            select_cols.append("d.COPAGO AS CM_COPAGO")
        if "OBSERVACION" in campos_set:
            select_cols.append("d.OBSERVACION")
        select_cols.append("COALESCE(s.NAME, '') AS ESTADO")
        # Campos comunes al final
        select_cols.append("t.fecha_creacion AS FechaDigitacion")
        select_cols.append(
            "CONCAT(u.FIRST_NAME, ' ', u.LAST_NAME, ' - ', CAST(u.NUM_DOC AS varchar(20))) AS Funcionario"
        )

        sql_export = (
            "SELECT " + ", ".join(select_cols) + " "
            "FROM ASIGNACION_TIPIFICACION a "
            "JOIN TIPIFICACION t               ON t.ASIGNACION_ID          = a.RADICADO "
            "LEFT JOIN TIPO_DOC td             ON t.TIPO_DOC_ID            = td.ID "
            "LEFT JOIN USERS u                 ON t.USER_ID                = u.ID "
            "LEFT JOIN TIPIFICACION_DETALLES d ON d.TIPIFICACION_ID        = t.ID "
            "LEFT JOIN STATUS s                ON s.ID                     = a.STATUS_ID "
            f"WHERE {where}"
        )

        # 2) Elegir ruta/extensión
        path = filedialog.asksaveasfilename(
            title="Guardar archivo",
            initialfile="reporte",
            defaultextension=".csv",
            filetypes=[
                ("CSV (texto)", "*.csv"),
                ("TXT (texto)", "*.txt"),
                ("Excel con estilo", "*.xlsx"),
                ("PDF", "*.pdf"),
            ],
        )
        if not path:
            return
        ext = path.rsplit('.', 1)[-1].lower()

        # 3) Ejecutar consulta
        cur = conn.cursor()
        cur.execute(sql_export, params)
        rows = cur.fetchall()
        headers = [col[0] for col in cur.description]
        rename_map = {
            "FechaDigitacion": "FECHA DE DIGITACION",
            "Funcionario": "FUNCIONARIO",
        }
        headers = [rename_map.get(h, h) for h in headers]
        cur.close()

        # 4) Exportar según extensión
        if ext in ("csv", "txt"):
            sep = ',' if ext == "csv" else '\t'
            with open(path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f, delimiter=sep)
                writer.writerow(headers)
                writer.writerows(rows)
            messagebox.showinfo("Exportar", f"{len(rows)} filas exportadas en {ext.upper()}:\n{path}")

        elif ext == "xlsx":
            exportar_excel(path, headers, rows)
            messagebox.showinfo("Exportar", f"{len(rows)} filas exportadas en Excel:\n{path}")

        elif ext == "pdf":
            exportar_pdf(path, headers, rows)
            messagebox.showinfo("Exportar", f"{len(rows)} filas exportadas en PDF:\n{path}")




    def exportar_pdf(path, headers, rows):
        # Documento horizontal, márgenes reducidos
        doc = SimpleDocTemplate(
            path,
            pagesize=landscape(A4),
            leftMargin=10,
            rightMargin=10,
            topMargin=10,
            bottomMargin=10
        )
        wrap_style = ParagraphStyle(
            name="wrap_style",
            fontSize=7,
            leading=9,    # altura de línea para wrapping
        )

        # reconstruye la tabla usando Paragraph en cada celda
        data = [
            [Paragraph(str(col), wrap_style) for col in headers]
        ] + [
            [Paragraph(str(cell), wrap_style) for cell in row]
            for row in rows
        ]

        # Estilo de tabla común
        table_style = TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#D3D3D3')),
            ('TEXTCOLOR',  (0,0), (-1,0), colors.black),
            ('ALIGN',      (0,0), (-1,-1), 'LEFT'),
            ('VALIGN',     (0,0), (-1,-1), 'TOP'),
            ('GRID',       (0,0), (-1,-1), 0.5, colors.grey),
            ('FONTSIZE',   (0,0), (-1,-1), 7),  # Fuente pequeña para que quepa
            ('BOTTOMPADDING', (0,0), (-1,0), 6),
            ('TOPPADDING',    (0,0), (-1,0), 6),
        ])

        # ¿Cuántas columnas caben en una página?
        total_width = doc.width
        min_col_width = 40  # puntos mínimos por columna
        max_cols_per_page = max(1, int(total_width // min_col_width))

        elements = []
        if len(headers) > max_cols_per_page:
            # Partir en “páginas” de columnas
            for start in range(0, len(headers), max_cols_per_page):
                end = start + max_cols_per_page
                chunk_headers = headers[start:end]
                chunk_rows = [row[start:end] for row in rows]
                chunk_data = [chunk_headers] + chunk_rows

                colW = total_width / len(chunk_headers)
                tbl = Table(chunk_data, colWidths=[colW]*len(chunk_headers), repeatRows=1)
                tbl.setStyle(table_style)
                elements.append(tbl)
                elements.append(PageBreak())
        else:
            # Una sola tabla
            colW = total_width / len(headers)
            tbl = Table(data, colWidths=[colW]*len(headers), repeatRows=1)
            tbl.setStyle(table_style)
            elements.append(tbl)

        doc.build(elements)

    # — Creación de ventana principal y layout de filtros/pestañas —
    win = ctk.CTkToplevel(root)
    win.title("Ver Progreso de Paquetes")
    win.minsize(1080, 500)
    win.grab_set()
    win.protocol("WM_DELETE_WINDOW", lambda: safe_destroy(win))

    # -- Layout general: barra lateral de filtros + contenido central --
    sidebar = ctk.CTkScrollableFrame(win, width=320)
    sidebar.pack(side="left", fill="y", padx=(20, 10), pady=20)
    for i in range(3):
        sidebar.grid_columnconfigure(i, weight=1)

    content = ctk.CTkFrame(win, fg_color="transparent")
    content.pack(side="left", fill="both", expand=True, padx=(0, 20), pady=20)

    # Paquete
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT NUM_PAQUETE FROM ASIGNACION_TIPIFICACION ORDER BY NUM_PAQUETE")
    paquetes = [str(r[0]) for r in cur.fetchall()] or ["0"]
    cur.close()
    pkg_var = tk.StringVar(value=paquetes[0])
    ctk.CTkLabel(sidebar, text="Paquete:").grid(row=0, column=0, sticky="w")
    ctk.CTkOptionMenu(sidebar, values=paquetes, variable=pkg_var, width=120).grid(row=0, column=1, padx=(0,10), sticky="w")
    pkg_var.trace_add("write", lambda *a: actualizar_tabs())

    # Tipo de paquete
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT TIPO_PAQUETE FROM ASIGNACION_TIPIFICACION ORDER BY TIPO_PAQUETE")
    tipos_paquete = [r[0] or "" for r in cur.fetchall()]
    cur.close()
    var_tipo_paquete = tk.StringVar(value=tipos_paquete[0])
    ctk.CTkLabel(sidebar, text="Tipo Paquete:").grid(row=1, column=0, sticky="w", pady=(10,0))
    ctk.CTkOptionMenu(sidebar, values=tipos_paquete, variable=var_tipo_paquete, width=120).grid(row=1, column=1, padx=(0,10), sticky="w")
    var_tipo_paquete.trace_add("write", lambda *a: actualizar_tabs())

    # Fechas
    var_fecha_desde = tk.StringVar()
    
    var_fecha_hasta = tk.StringVar()
    
    def limpiar_fechas():
        # 1) Borra las variables para que construir_filtros ignore las fechas
        var_fecha_desde.set("")
        var_fecha_hasta.set("")
        # 2) Borra el texto que queda en los DateEntry (usa su método delete)
        fecha_desde.delete(0, "end")
        fecha_hasta.delete(0, "end")
    
    ctk.CTkLabel(sidebar, text="Desde:").grid(row=2, column=0, sticky="w", pady=(10,0))
    fecha_desde = DateEntry(sidebar, width=12, locale='es_CO', date_pattern='dd/MM/yyyy', textvariable=var_fecha_desde)
    fecha_desde.grid(row=2, column=1, padx=(0,10), sticky="w")
    fecha_desde.delete(0, 'end')
    var_fecha_desde.trace_add("write", lambda *a: actualizar_tabs())

    ctk.CTkLabel(sidebar, text="Hasta:").grid(row=3, column=0, sticky="w")
    fecha_hasta = DateEntry(sidebar, width=12, locale='es_CO', date_pattern='dd/MM/yyyy', textvariable=var_fecha_hasta)
    fecha_hasta.grid(row=3, column=1, padx=(0,10), sticky="w")
    fecha_hasta.delete(0, 'end')
    var_fecha_hasta.trace_add("write", lambda *a: actualizar_tabs())

    ctk.CTkButton(sidebar, text="Limpiar fechas", command=limpiar_fechas, width=100).grid(row=4, column=0, columnspan=2, pady=(10,0))
    ctk.CTkButton(sidebar, text="Exportar", command=exportar, width=100).grid(row=5, column=0, columnspan=2, pady=(10,0))
    # Filtro de estados
    cur = conn.cursor()
    cur.execute("SELECT NAME FROM STATUS ORDER BY NAME")
    estados = [r[0] for r in cur.fetchall()]
    cur.close()
    ctk.CTkLabel(sidebar, text="Estado:").grid(row=6, column=0, sticky="w", pady=(10,5))
    buscar_est = ctk.CTkEntry(sidebar, width=180, placeholder_text="Buscar estado...")
    buscar_est.grid(row=6, column=1, columnspan=2, sticky="w", padx=(0,10), pady=(10,5))
    buscar_est.bind("<KeyRelease>", _filtrar_est)
    ctk.CTkButton(sidebar, text="Todo", command=lambda: _marcar_est(True), width=60).grid(row=7, column=0, sticky="w")
    ctk.CTkButton(sidebar, text="Solo visibles", command=_solo_visibles_est, width=80).grid(row=7, column=1)
    ctk.CTkButton(sidebar, text="Ninguno", command=lambda: _marcar_est(False), width=60).grid(row=7, column=2, sticky="e")
    estado_frame = ctk.CTkScrollableFrame(sidebar, width=280, height=120)
    estado_frame.grid(row=8, column=0, columnspan=3, sticky="w")
    estado_vars = {}
    estado_checks = {}
    for est in estados:
        var = tk.BooleanVar(value=True)
        cb = ctk.CTkCheckBox(estado_frame, text=est, variable=var, command=actualizar_tabs)
        cb.pack(anchor="w", pady=2)
        estado_vars[est] = var
        estado_checks[est] = cb

    def _solo_visibles_est():
        for var in estado_vars.values():
            var.set(False)
        for est in estados:
            cb = estado_checks[est]
            if cb.winfo_ismapped():
                estado_vars[est].set(True)
        actualizar_tabs()


    # Filtro de usuarios
    cur = conn.cursor()
    cur.execute("SELECT FIRST_NAME + ' ' + LAST_NAME FROM USERS ORDER BY FIRST_NAME")
    usuarios = [r[0] for r in cur.fetchall()]
    cur.close()
    ctk.CTkLabel(sidebar, text="Usuario:").grid(row=9, column=0, sticky="w", pady=(10,5))
    buscar_usr = ctk.CTkEntry(sidebar, width=180, placeholder_text="Buscar usuario...")
    buscar_usr.grid(row=9, column=1, columnspan=2, sticky="w", padx=(0,10), pady=(10,5))
    buscar_usr.bind("<KeyRelease>", _filtrar_usr)
    ctk.CTkButton(sidebar, text="Todo", command=lambda: _marcar_usr(True), width=60).grid(row=10, column=0, sticky="w")
    ctk.CTkButton(sidebar, text="Solo visibles", command=_solo_visibles_usr, width=80).grid(row=10, column=1)
    ctk.CTkButton(sidebar, text="Ninguno", command=lambda: _marcar_usr(False), width=60).grid(row=10, column=2, sticky="e")

    user_frame = ctk.CTkScrollableFrame(sidebar, width=280, height=120)
    user_frame.grid(row=11, column=0, columnspan=3, sticky="w")
    user_vars = {}
    user_checks = {}
    for usr in usuarios:
        var = tk.BooleanVar(value=True)
        cb = ctk.CTkCheckBox(user_frame, text=usr, variable=var, command=actualizar_tabs)
        cb.pack(anchor="w", pady=2)
        user_vars[usr] = var
        user_checks[usr] = cb
        

    ctk.CTkLabel(sidebar, text="Radicados (uno por línea):").grid(row=12, column=0, sticky="nw", pady=(10,0))
    rad_text = ctk.CTkTextbox(sidebar, width=240, height=100)
    rad_text.grid(row=12, column=1, columnspan=2, sticky="w", pady=(10,0))

    rad_text.bind("<KeyRelease>", lambda e: actualizar_tabs())

    # Pestañas de resultados
    tabs = ctk.CTkTabview(content, width=760, height=440)
    tabs.pack(padx=20, pady=(10,20), fill="both", expand=True)
    tabs.add("Por Estado")
    tabs.add("Por Usuario")
    win._tabview = tabs

    # Carga inicial al abrir ventana
    actualizar_tabs()


def actualizar_tabs(win, conn, num_paquete,where, params):
    tabs = win._tabview
    # -- Pestaña “Por Estado” --
    frame1 = tabs.tab("Por Estado")
    for w in frame1.winfo_children(): w.destroy()

    cur = conn.cursor()
    cur.execute("""
        SELECT UPPER(s.NAME) AS ESTADO, COUNT(*) AS CNT
          FROM ASIGNACION_TIPIFICACION at
          JOIN STATUS s ON at.STATUS_ID = s.ID
         WHERE at.NUM_PAQUETE = %s
         GROUP BY s.NAME
         ORDER BY s.NAME
    """, (num_paquete,))
    datos = cur.fetchall(); cur.close()
    
    for i, (est, cnt) in enumerate(datos):
        ctk.CTkLabel(frame1, text=est).grid(row=i, column=0, sticky="w", padx=5, pady=2)
        ctk.CTkLabel(frame1, text=cnt).grid(row=i, column=1, sticky="e", padx=5, pady=2)

    # —> Cálculo total de HECHOS (STATUS_ID 3 + 4)
    cur2 = conn.cursor()
    sql_hechos = (
        "SELECT SUM(CASE WHEN a.STATUS_ID IN (3,4) THEN 1 ELSE 0 END) "
        "FROM ASIGNACION_TIPIFICACION a "
        "JOIN TIPIFICACION t ON t.ASIGNACION_ID = a.RADICADO "
        "JOIN USERS u ON t.USER_ID = u.ID "
        "JOIN STATUS s ON a.STATUS_ID = s.ID "
        f"WHERE {where}"
    )
    cur2.execute(sql_hechos, params)
    total_hechos = cur2.fetchone()[0] or 0
    cur2.close()

    # Pintamos la fila "HECHOS" al final
    fila_final = len(datos)
    ctk.CTkLabel(frame1, text="HECHOS",   font=("Arial", 12, "bold"))\
        .grid(row=fila_final, column=0, sticky="w", padx=5, pady=4)
    ctk.CTkLabel(frame1, text=str(total_hechos), font=("Arial", 12, "bold"))\
        .grid(row=fila_final, column=1, sticky="e", padx=5, pady=4)


    # -- Pestaña “Por Usuario” --
    frame2 = tabs.tab("Por Usuario")
    for w in frame2.winfo_children(): w.destroy()

    cur = conn.cursor()
    cur.execute("""
        SELECT u.ID,
               u.FIRST_NAME + ' ' + u.LAST_NAME AS USUARIO,
               SUM(CASE WHEN at.STATUS_ID=2 THEN 1 ELSE 0 END) AS PENDIENTES,
               SUM(CASE WHEN at.STATUS_ID=3 THEN 1 ELSE 0 END) AS PROCESADOS,
               SUM(CASE WHEN at.STATUS_ID=4 THEN 1 ELSE 0 END) AS CON_OBS
          FROM TIPIFICACION t
          JOIN USERS u     ON t.USER_ID = u.ID
          JOIN ASIGNACION_TIPIFICACION at
            ON at.RADICADO = t.ASIGNACION_ID
         WHERE at.NUM_PAQUETE = %s
         GROUP BY u.ID, u.FIRST_NAME, u.LAST_NAME
         ORDER BY USUARIO
    """, (num_paquete,))
    usuarios = cur.fetchall(); cur.close()

    # Encabezados
    cols = ["ID", "USUARIO", "PENDIENTES", "PROCESADOS", "CON_OBS"]
    for j, h in enumerate(cols):
        ctk.CTkLabel(frame2, text=h, anchor="w")\
            .grid(row=0, column=j, padx=5, pady=4)

    # Filas
    for i, fila in enumerate(usuarios, start=1):
        for j, val in enumerate(fila):
            ctk.CTkLabel(frame2, text=str(val), anchor="w")\
                .grid(row=i, column=j, padx=5, pady=2)

# Subclase de AutocompleteEntry que fuerza mayúsculas y mantiene el desplegable
class UppercaseAutocompleteEntry(AutocompleteEntry):
    def __init__(self, parent, values, textvariable=None, **kwargs):
        super().__init__(parent, values, textvariable=textvariable, **kwargs)
        # quita cualquier traza previa
        for trace in self.var.trace_info():
            if trace[0] == 'write':
                self.var.trace_remove('write', trace[1])
        # añade nueva traza
        self.var.trace_add('write', self._on_var_write)

    def _on_var_write(self, *args):
        txt = self.var.get()
        up = txt.upper()
        if up != txt:
            pos = self.index(tk.INSERT)
            self.var.set(up)
            self.icursor(pos)
        # luego lanza el autocomplete normal
        self._show_matches()


def modificar_estado_usuario(root, conn):
    # 1) Crear ventana
    win = ctk.CTkToplevel(root)
    win.title("Modificar Estado de Usuario")
    win.geometry("500x400")
    win.grab_set()
    win.protocol("WM_DELETE_WINDOW", lambda w=win: safe_destroy(w))

    frm = ctk.CTkFrame(win, fg_color="transparent")
    frm.pack(padx=20, pady=20, fill="x")

    # 2) Cargo tipo_doc ID y NAME
    cur = conn.cursor()
    cur.execute("SELECT ID, NAME FROM TIPO_DOC")
    tipo_rows = cur.fetchall()  # [(1,'CC'),(2,'TI'),...]
    cur.close()
    tipo_map   = {name.upper(): tid for tid, name in tipo_rows}
    tipo_names = [name for _, name in tipo_rows]

    # Campo Autocomplete de TipoDoc (uppercase inside class)
    ctk.CTkLabel(frm, text="Tipo Doc:").grid(row=0, column=0, sticky="w", pady=5)
    entry_tipo = UppercaseAutocompleteEntry(
        frm,
        tipo_names,
        width=250
    )
    entry_tipo.grid(row=0, column=1, pady=5)

    # Num Doc
    ctk.CTkLabel(frm, text="Num Doc:").grid(row=1, column=0, sticky="w", pady=5)
    num_var = tk.StringVar()
    entry_num = ctk.CTkEntry(frm, textvariable=num_var, width=250)
    entry_num.grid(row=1, column=1, pady=5)

    # Botón Buscar
    ctk.CTkButton(frm, text="Buscar", width=100,
                  command=lambda: buscar_usuario()).grid(
        row=2, column=0, columnspan=2, pady=(10,20))

    # Área de resultados
    result_frame = ctk.CTkFrame(win)
    result_frame.pack(padx=20, pady=(0,20), fill="both", expand=True)

    def buscar_usuario():
        # limpio previos
        for w in result_frame.winfo_children():
            w.destroy()

        # obtengo tipo_id
        nombre = entry_tipo.get().strip().upper()
        tipo_id = tipo_map.get(nombre)
        if tipo_id is None:
            messagebox.showwarning("Error", "Tipo Doc no válido.")
            return

        # num doc
        try:
            nd = int(num_var.get().strip())
        except ValueError:
            messagebox.showwarning("Error", "Num Doc debe ser número.")
            return

        # consulta usuario
        cur2 = conn.cursor()
        cur2.execute("""
            SELECT u.ID, u.FIRST_NAME, u.LAST_NAME, u.STATUS_ID, s.NAME
              FROM USERS u
              JOIN STATUS s ON u.STATUS_ID = s.ID
             WHERE u.TYPE_DOC_ID=%s AND u.NUM_DOC=%s AND u.STATUS_ID IN (5,6)
        """, (tipo_id, nd))
        row = cur2.fetchone()
        cur2.close()

        if not row:
            messagebox.showinfo("No encontrado",
                                "No hay usuario en estado 5 o 6 con esos datos.")
            return

        user_id, fn, ln, st_id, st_name = row
        ctk.CTkLabel(result_frame,
                     text=f"Usuario: {fn} {ln}  (ID {user_id})",
                     anchor="w").pack(fill="x", pady=(0,5))
        ctk.CTkLabel(result_frame,
                     text=f"Estado actual: {st_name.upper()}",
                     anchor="w").pack(fill="x", pady=(0,10))

        # cargo estados 5 y 6
        cur3 = conn.cursor()
        cur3.execute("SELECT ID, NAME FROM STATUS WHERE ID IN (5,6)")
        estados = cur3.fetchall()  # [(5,"PENDIENTE"),(6,"RECHAZADO")]
        cur3.close()
        est_map   = {name.upper(): id_ for (id_, name) in estados}
        est_names = [name.upper() for (_, name) in estados]

        # Selector en lugar de Autocomplete
        ctk.CTkLabel(result_frame, text="Nuevo estado:").pack(anchor="w", pady=(5,0))
        estado_var = tk.StringVar(value=st_name.upper())
        opt = ctk.CTkOptionMenu(
            result_frame,
            values=est_names,
            variable=estado_var,
            width=250
        )
        opt.pack(pady=(0,10))

        def actualizar():
            sel = estado_var.get().strip().upper()
            new_id = est_map.get(sel)
            if new_id is None:
                messagebox.showwarning("Error", "Estado no válido.")
                return
            cur4 = conn.cursor()
            cur4.execute("UPDATE USERS SET STATUS_ID=%s WHERE ID=%s", (new_id, user_id))
            conn.commit()
            cur4.close()
            messagebox.showinfo("Listo", f"Estado cambiado a {sel}.")
            safe_destroy(win)

        ctk.CTkButton(result_frame, text="Actualizar", command=actualizar, width=120)\
            .pack(pady=(10,0))

    entry_tipo.focus()

def exportar_paquete(root, conn):

    # 1) Crear ventana
    win = ctk.CTkToplevel(root)
    win.title("Exportar Paquete")
    win.geometry("500x450")
    win.grab_set()
    win.protocol("WM_DELETE_WINDOW", lambda w=win: safe_destroy(w))

    frm = ctk.CTkFrame(win, fg_color="transparent")
    frm.pack(padx=20, pady=20, fill="both", expand=True)

    # 2) Obtener lista de paquetes
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT NUM_PAQUETE FROM ASIGNACION_TIPIFICACION ORDER BY NUM_PAQUETE")
    paquetes = [r[0] for r in cur.fetchall()]
    cur.close()
    if not paquetes:
        messagebox.showinfo("Exportar", "No hay paquetes para exportar.")
        win.destroy()
        return

    # 3) Selector de paquete
    ctk.CTkLabel(frm, text="Paquete:").grid(row=0, column=0, sticky="w", pady=(0,5))
    pkg_var = tk.StringVar(value=str(paquetes[0]))
    ctk.CTkOptionMenu(
        frm,
        values=[str(p) for p in paquetes],
        variable=pkg_var,
        width=120
    ).grid(row=0, column=1, pady=(0,5), sticky="w")

    # 4) Selector de formato
    ctk.CTkLabel(frm, text="Formato:").grid(row=1, column=0, sticky="w", pady=(0,10))
    fmt_var = tk.StringVar(value="CSV")
    ctk.CTkOptionMenu(
        frm,
        values=["CSV", "TXT"],
        variable=fmt_var,
        width=120
    ).grid(row=1, column=1, pady=(0,10), sticky="w")

    # 5) Área de texto para radicados
    ctk.CTkLabel(frm, text="Radicados (uno por línea, opcional):") \
        .grid(row=2, column=0, columnspan=2, sticky="w")
    text_frame = tk.Frame(frm)
    text_frame.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=(0,10))
    frm.grid_rowconfigure(3, weight=1)
    frm.grid_columnconfigure(1, weight=1)

    scrollbar = tk.Scrollbar(text_frame)
    scrollbar.pack(side="right", fill="y")
    txt = tk.Text(text_frame, height=8, yscrollcommand=scrollbar.set)
    txt.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=txt.yview)

    # 6) Botón Exportar
    def _export():
        pkg = int(pkg_var.get())
        fmt = fmt_var.get()
        ext = ".txt" if fmt == "TXT" else ".csv"
        sep = ";" if fmt == "TXT" else ","

        # Ajustar diálogo según formato
        filetypes = [("Text Files", "*.txt")] if fmt == "TXT" else [("CSV Files", "*.csv")]
        path = filedialog.asksaveasfilename(
            filetypes=filetypes,
            defaultextension=ext,
            initialfile=f"reporte{ext}",
            title=f"Guardar como {fmt}"
        )
        if not path:
            return

        # 7) Leo radicados del textarea
        lines = txt.get("1.0", "end").splitlines()
        radicados = sorted({int(L) for L in lines if L.strip().isdigit()})


        # 8) Determino tipo y campos configurados
        cur_t = conn.cursor()
        cur_t.execute(
            "SELECT TOP 1 TIPO_PAQUETE FROM ASIGNACION_TIPIFICACION WHERE NUM_PAQUETE=%s",
            (pkg,)
        )
        tipo_pkg = (cur_t.fetchone() or [""])[0]
        cur_t.close()

        cur_c = conn.cursor()
        cur_c.execute(
            "SELECT campo FROM PAQUETE_CAMPOS WHERE NUM_PAQUETE=%s AND tipo_paquete=%s",
            (pkg, tipo_pkg)
        )
        campos_set = {r[0] for r in cur_c.fetchall()}
        cur_c.close()

        # 9) Construyo SELECT dinámico
        select_cols = ["a.RADICADO"]
        if "FECHA_SERVICIO" in campos_set:
            select_cols.append("CONVERT(varchar(10), t.FECHA_SERVICIO, 103) AS FECHA_SERVICIO")
        if "FECHA_SERVICIO_FINAL" in campos_set:
            select_cols.append("CONVERT(varchar(10), t.FECHA_SERVICIO_FINAL, 103) AS FECHA_SERVICIO_FINAL")
        if "AUTORIZACION" in campos_set:
            select_cols.append("d.AUTORIZACION")
        if "CODIGO_SERVICIO" in campos_set:
            select_cols.append("d.CODIGO_SERVICIO AS COD_SERVICIO")
        if "CANTIDAD" in campos_set:
            select_cols.append("d.CANTIDAD")
        if "VLR_UNITARIO" in campos_set:
            select_cols.append("d.VLR_UNITARIO")
        if "DIAGNOSTICO" in campos_set:
            select_cols.append("t.DIAGNOSTICO")
        if "TIPO_DOC_ID" in campos_set:
            select_cols.append("td.NAME AS TipoDocumento")
        if "NUM_DOC" in campos_set:
            select_cols.append("t.NUM_DOC AS NumeroDocumento")
        if "COPAGO" in campos_set:
            select_cols.append("d.COPAGO AS CM_COPAGO")
        if "OBSERVACION" in campos_set:
            select_cols.append("d.OBSERVACION")
        select_cols.append("t.fecha_creacion AS FechaDigitacion")
        select_cols.append("u2.NUM_DOC AS Funcionario")

        base_sql = (
            "SELECT " + ", ".join(select_cols) + " "
            "FROM ASIGNACION_TIPIFICACION a "
            "JOIN TIPIFICACION t             ON t.ASIGNACION_ID        = a.RADICADO "
            "LEFT JOIN TIPIFICACION_DETALLES d ON d.TIPIFICACION_ID      = t.ID "
            "LEFT JOIN USERS u2               ON u2.ID                 = t.USER_ID "
            "LEFT JOIN TIPO_DOC td            ON td.ID                 = t.TIPO_DOC_ID "
        )

        # 9) Elijo la cláusula WHERE según si hay radicados
        if radicados:
            ph = ",".join("%s" for _ in radicados)
            where_clause = f"WHERE a.RADICADO IN ({ph})"
            params = radicados
        else:
            where_clause = "WHERE a.NUM_PAQUETE = %s"
            params = [pkg]

        sql = f"{base_sql} {where_clause} ORDER BY a.RADICADO, t.FECHA_SERVICIO"

        # 10) Ejecutar y escribir archivo
        cur2 = conn.cursor()
        cur2.execute(sql, params)
        rows = cur2.fetchall()
        headers = [col[0] for col in cur2.description]
        rename_map = {
            "FechaDigitacion": "FECHA DE DIGITACION",
            "Funcionario": "FUNCIONARIO",
        }
        headers = [rename_map.get(h, h) for h in headers]
        cur2.close()

        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, delimiter=sep)
            writer.writerow(headers)
            writer.writerows(rows)

        messagebox.showinfo(
            "Exportar",
            f"Exportación completada:\n{path}\n"
            f"{len(rows)} registros."
        )
        win.destroy()

    ctk.CTkButton(frm, text="Exportar", command=_export, width=200) \
        .grid(row=4, column=0, columnspan=2, pady=(10,0))

def actualizar_usuario(root, conn, user_id):
    """
    Abre una ventana para actualizar nombre, apellido, correo y contraseña
    del usuario identificado por user_id. La contraseña se encripta con bcrypt
    sólo si el usuario ingresa una nueva.
    """
    # 1) Crear ventana
    win = ctk.CTkToplevel(root)
    win.title("Actualizar Mis Datos")
    win.geometry("400x300")
    win.grab_set()
    
    letters_vcmd = win.register(
    lambda P: bool(re.fullmatch(r"[A-Za-zÑñ ]*", P))
    )

    # 2) Variables de texto
    first_name_var = tk.StringVar()
    last_name_var  = tk.StringVar()
    email_var      = tk.StringVar()
    password_var   = tk.StringVar()

    # 3) Precargar datos del usuario
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT FIRST_NAME, LAST_NAME, CORREO FROM USERS WHERE ID = %s",
            (user_id,)
        )
        row = cur.fetchone()
        cur.close()

        if not row:
            messagebox.showerror("Error", "Usuario no encontrado.")
            win.destroy()
            return

        first_name_var.set(row[0])
        last_name_var.set(row[1])
        email_var.set(row[2])

    except Exception as e:
        messagebox.showerror("Error", f"Error al cargar datos:\n{e}")
        win.destroy()
        return

    # 4) Construir formulario
    frm = ctk.CTkFrame(win, fg_color="transparent")
    frm.pack(padx=20, pady=20, fill="both", expand=True)

    # Campos: etiqueta + entrada
    fields = [
        ("Nombres:",     first_name_var, False),
        ("Apellidos:",   last_name_var,  False),
        ("Correo:",      email_var,      False),
        ("Nueva contraseña:", password_var, True),
    ]
    for i, (label_text, var, is_password) in enumerate(fields):
        ctk.CTkLabel(frm, text=label_text).grid(row=i, column=0, sticky="w", pady=5)
        entry = ctk.CTkEntry(
            frm,
            textvariable=var,
            width=250,
            show="*" if is_password else None,
            validate="key" if not is_password else "none",
            validatecommand=(letters_vcmd, "%P") if not is_password else None
        )
        entry.grid(row=i, column=1, pady=5)
        if i == 0:
            entry.focus()

        # Forzar mayúsculas en nombres/apellidos
        if not is_password:
            def on_write(*args, v=var):
                txt = v.get()
                if txt != txt.upper():
                    v.set(txt.upper())
            var.trace_add("write", on_write)

    # 5) Función para validar y guardar cambios
    def guardar_usuario():
        # Validar campos obligatorios (excepto contraseña)
        if not (first_name_var.get().strip() and last_name_var.get().strip() and email_var.get().strip()):
            messagebox.showwarning("Campos Vacíos", "Completa nombre, apellido y correo.")
            return

        pwd_text = password_var.get().strip()
        pwd_hash = None
        if pwd_text:
            # Encriptar nueva contraseña
            try:
                pwd_hash = bcrypt.hashpw(pwd_text.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo encriptar la contraseña:\n{e}")
                return

        try:
            cur = conn.cursor()
            if pwd_hash:
                # Actualizar incluyendo contraseña
                cur.execute("""
                    UPDATE USERS
                       SET FIRST_NAME = %s,
                           LAST_NAME  = %s,
                           CORREO     = %s,
                           PASSWORD   = %s
                     WHERE ID = %s
                """, (
                    first_name_var.get().strip(),
                    last_name_var.get().strip(),
                    email_var.get().strip(),
                    pwd_hash,
                    user_id
                ))
            else:
                # Actualizar sin tocar la contraseña
                cur.execute("""
                    UPDATE USERS
                       SET FIRST_NAME = %s,
                           LAST_NAME  = %s,
                           CORREO     = %s
                     WHERE ID = %s
                """, (
                    first_name_var.get().strip(),
                    last_name_var.get().strip(),
                    email_var.get().strip(),
                    user_id
                ))
            conn.commit()
            cur.close()

            messagebox.showinfo("Éxito", "Usuario actualizado correctamente.")
            win.destroy()

        except Exception as e:
            messagebox.showerror("Error", f"Error al actualizar:\n{e}")

    # 6) Botón Guardar
    btn = ctk.CTkButton(win, text="Guardar", command=guardar_usuario, width=200)
    btn.pack(pady=10)

    win.mainloop()

if "--crear-usuario" in sys.argv:
    import tkinter as tk
    from tkinter import messagebox
    import customtkinter as ctk
    import re, bcrypt
    # aquí pones tu clase o función crear_usuario idéntica
    class AppTk:
        def __init__(self, conn):
            apply_ctk_theme_from_settings()
            self.conn = conn
            self._tk_root = tk.Tk()
            self._tk_root.withdraw()
        def crear_usuario(self):
        # —–– Asegurarse de tener un root de Tkinter para el register() —––
                if not hasattr(self, '_tk_root'):
                    self._tk_root = tk.Tk()
                    self._tk_root.withdraw()

                # Validaciones
                def only_letters_and_spaces(P):
                    # P es el texto completo tras el cambio
                    return all(c.isalpha() or c == ' ' for c in P)

                def only_digits(P):
                    return P == "" or P.isdigit()

                def validate_email(email):
                    regex = r'^[a-zA-Z0-9_.+\-]+@[a-zA-Z0-9\-]+\.[a-zA-Z0-9\-.]+$'
                    return re.match(regex, email) is not None

                # Aquí usamos el register() del Tk oculto
                vcmd_letters = (self._tk_root.register(only_letters_and_spaces), '%P')
                vcmd_digits  = (self._tk_root.register(only_digits), '%P')

                # Crear ventana secundaria
                top = ctk.CTkToplevel(self._tk_root)
                top.title("Crear Usuario")
                top.geometry("500x600")
                top.resizable(False, False)

                # Recuperar catálogos
                cur = self.conn.cursor()
                cur.execute("SELECT ID, NAME FROM TIPO_DOC")
                tipos = cur.fetchall()
                cur.execute("SELECT ID, NAME FROM STATUS WHERE ID IN (%s, %s)", (5, 6))
                statuses = cur.fetchall()
                cur.execute("SELECT ID, NAME FROM ROL")
                roles = cur.fetchall()
                cur.close()

                tipo_map   = {name: tid for tid, name in tipos}
                status_map = {name: sid for sid, name in statuses}

                # Variables
                fn_var   = tk.StringVar()
                ln_var   = tk.StringVar()
                doc_var  = tk.StringVar()
                pwd_var  = tk.StringVar()
                email_var = tk.StringVar()  # Variable para el correo
                tipo_var = tk.StringVar(value=tipos[0][1] if tipos else "")
                stat_var = tk.StringVar(value=statuses[0][1] if statuses else "")

                # Frame
                top.grid_rowconfigure(0, weight=1)
                top.grid_columnconfigure(0, weight=1)
                frm = ctk.CTkFrame(top, corner_radius=8)
                frm.grid(row=0, column=0, padx=20, pady=20)

                # Definición de campos
                labels = [
                    ("Nombres:", fn_var, "entry_upper", vcmd_letters),
                    ("Apellidos:", ln_var, "entry_upper", vcmd_letters),
                    ("Tipo Doc:", tipo_var, "combo", [n for _, n in tipos]),
                    ("N° Documento:", doc_var, "entry_digit", vcmd_digits),
                    ("Correo:", email_var, "entry_email", None),  # Nuevo campo correo
                    ("Contraseña:", pwd_var, "entry_pass", None),
                    ("Status:", stat_var, "combo", [n for _, n in statuses])
                ]

                for i, (text, var, kind, cmd) in enumerate(labels):
                    ctk.CTkLabel(frm, text=text).grid(row=i, column=0, sticky="w", pady=(10, 0))
                    if kind == "entry_upper":
                        widget = ctk.CTkEntry(
                            frm,
                            textvariable=var,
                            width=300,
                            validate="key",
                            validatecommand=vcmd_letters   # <-- usa el validador actualizado
                        )
                        # Forzar mayúsculas tras cada tecla
                        widget.bind("<KeyRelease>", lambda e, v=var: v.set(v.get().upper()))

                    elif kind == "entry_digit":
                        widget = ctk.CTkEntry(
                            frm,
                            textvariable=var,
                            width=300,
                            validate="key",
                            validatecommand=cmd
                        )

                    elif kind == "entry_pass":
                        widget = ctk.CTkEntry(
                            frm,
                            textvariable=var,
                            show="*",
                            width=300
                        )

                    elif kind == "entry_email":
                        widget = ctk.CTkEntry(
                            frm,
                            textvariable=var,
                            width=300
                        )

                    else:  # combo
                        widget = ctk.CTkComboBox(
                            frm,
                            values=cmd,
                            variable=var,
                            width=300
                        )

                    widget.grid(row=i, column=1, padx=(10, 0), pady=(10, 0))
                    if i == 0:
                        widget.focus()

                # Checkboxes de roles
                ctk.CTkLabel(frm, text="Roles:").grid(row=len(labels), column=0, sticky="nw", pady=(10, 0))
                rol_vars = {}
                chk_frame = ctk.CTkFrame(frm)
                chk_frame.grid(row=len(labels), column=1, sticky="w", pady=(10, 0))
                for j, (rid, rname) in enumerate(roles):
                    var_chk = tk.BooleanVar()
                    rol_vars[rid] = var_chk
                    ctk.CTkCheckBox(chk_frame, text=rname, variable=var_chk).grid(row=j, column=0, sticky="w", pady=2)

                # Función para guardar usuario
                def guardar_usuario(event=None):
                    # Validar campos
                    if not all([fn_var.get(), ln_var.get(), doc_var.get(), pwd_var.get(), email_var.get()]):
                        messagebox.showwarning("Faltan datos", "Completa todos los campos.")
                        return

                    # Validar correo
                    email = email_var.get().strip()
                    if not validate_email(email):
                        messagebox.showwarning("Correo inválido", "Por favor ingresa un correo válido.")
                        return

                    # Validar longitud de la contraseña
                    if len(pwd_var.get().strip()) < 6:
                        messagebox.showwarning("Contraseña inválida", "La contraseña debe tener al menos 6 caracteres.")
                        return

                    # Validar si al menos un rol ha sido seleccionado
                    if not any(v.get() for v in rol_vars.values()):
                        messagebox.showwarning("Faltan roles", "Debe seleccionar al menos un rol para el usuario.")
                        return

                    try:
                        first_name = fn_var.get().strip()
                        last_name  = ln_var.get().strip()
                        num_doc    = int(doc_var.get().strip())
                        email_address = email  # Correo ingresado por el usuario
                        type_id    = tipo_map[tipo_var.get()]

                        # Verificar duplicado
                        cursor = self.conn.cursor()
                        cursor.execute(
                            "SELECT COUNT(*) FROM USERS WHERE TYPE_DOC_ID = %s AND NUM_DOC = %s",
                            (type_id, num_doc)
                        )
                        if cursor.fetchone()[0] > 0:
                            cursor.close()
                            messagebox.showerror("Duplicado", "Ya existe un usuario con ese tipo y número de documento.")
                            return

                        # Cifrar contraseña con bcrypt
                        raw_pwd   = pwd_var.get().strip()
                        pwd_bytes = raw_pwd.encode('utf-8')
                        salt      = bcrypt.gensalt(rounds=12)
                        pwd_hash  = bcrypt.hashpw(pwd_bytes, salt).decode('utf-8')

                        status_id = status_map[stat_var.get()]
                        selected  = [rid for rid, v in rol_vars.items() if v.get()]

                        # Insertar usuario
                        cursor.execute(
                            """
                            INSERT INTO USERS 
                            (FIRST_NAME, LAST_NAME, TYPE_DOC_ID, NUM_DOC, PASSWORD, CORREO, STATUS_ID)
                            OUTPUT INSERTED.ID
                            VALUES (%s, %s, %s, %s, %s, %s, %s)
                            """,
                            (first_name, last_name, type_id, num_doc, pwd_hash, email_address, status_id)
                        )
                        new_id = cursor.fetchone()[0]

                        # Insertar roles
                        for rid in selected:
                            cursor.execute(
                                "INSERT INTO USER_ROLES (USER_ID, ROL_ID) VALUES (%s, %s)",
                                (new_id, rid)
                            )

                        self.conn.commit()
                        cursor.close()
                        messagebox.showinfo("Éxito", f"Usuario creado con ID {new_id}")
                        top.destroy()

                    except Exception as e:
                        messagebox.showerror("Error", str(e))

                # Botón Guardar
                btn = ctk.CTkButton(frm, text="Guardar Usuario", command=guardar_usuario, width=200)
                btn.grid(row=len(labels)+1, column=0, columnspan=2, pady=20)
                top.bind("<Return>", guardar_usuario)
        def run(self):
            self.crear_usuario()
            self._tk_root.mainloop()

    conn = conectar_sql_server("DB_DATABASE")
    if conn is None:
        sys.exit("No se pudo conectar a la BD.")
    AppTk(conn).run()   # ahora existe .run()
    sys.exit(0)
    
    
if "--iniciar-tipificacion" in sys.argv:
        # Configura tema según la preferencia guardada
        apply_ctk_theme_from_settings()
        # 1) crea el root oculto
        root = tk.Tk()
        root.withdraw()

        # 2) conecta a la BD
        conn = conectar_sql_server("DB_DATABASE")
        if conn is None:
            sys.exit("No se pudo conectar a la BD.")

        # 3) extrae el user_id que viene justo después del flag
        idx = sys.argv.index("--iniciar-tipificacion")
        try:
            user_id = int(sys.argv[idx + 1])
        except Exception:
            sys.exit("Falta user_id tras --iniciar-tipificacion")

        # 4) llama a tu función
        iniciar_tipificacion(root, conn, user_id)

        # 5) arranca el loop de Tk para CustomTkinter
        root.mainloop()
        sys.exit(0)
        
if "--iniciar-calidad" in sys.argv:
        # Configura tema según la preferencia guardada
        apply_ctk_theme_from_settings()
        # 1) crea el root oculto
        root = tk.Tk()
        root.withdraw()

        # 2) conecta a la BD
        conn = conectar_sql_server("DB_DATABASE")
        if conn is None:
            sys.exit("No se pudo conectar a la BD.")

        # 3) extrae el user_id que viene justo después del flag
        idx = sys.argv.index("--iniciar-calidad")
        try:
            user_id = int(sys.argv[idx + 1])
        except Exception:
            sys.exit("Falta user_id tras --iniciar-calidad")

        # 4) llama a tu función
        iniciar_calidad(root, conn, user_id)

        # 5) arranca el loop de Tk para CustomTkinter
        root.mainloop()
        sys.exit(0)
        
if "--ver-progreso" in sys.argv:
        # Configura tema según la preferencia guardada
        apply_ctk_theme_from_settings()
        # 1) crea el root oculto
        root = tk.Tk()
        root.withdraw()

        # 2) conecta a la BD
        conn = conectar_sql_server("DB_DATABASE")
        if conn is None:
            sys.exit("No se pudo conectar a la BD.")

        # 3) extrae el user_id que viene justo después del flag
        idx = sys.argv.index("--ver-progreso")

        # 4) llama a tu función
        ver_progreso(root, conn)

        # 5) arranca el loop de Tk para CustomTkinter
        root.mainloop()
        sys.exit(0)
        
if "--exportar-paquete" in sys.argv:
        # Configura tema según la preferencia guardada
        apply_ctk_theme_from_settings()
        # 1) crea el root oculto
        root = tk.Tk()
        root.withdraw()

        # 2) conecta a la BD
        conn = conectar_sql_server("DB_DATABASE")
        if conn is None:
            sys.exit("No se pudo conectar a la BD.")

        # 3) extrae el user_id que viene justo después del flag
        idx = sys.argv.index("--exportar-paquete")
            
        # 4) llama a tu función
        exportar_paquete(root, conn)

        # 5) arranca el loop de Tk para CustomTkinter
        root.mainloop()
        sys.exit(0)

if "--actualizar-datos" in sys.argv:
        # Configura tema según la preferencia guardada
        apply_ctk_theme_from_settings()
        # 1) crea el root oculto
        root = tk.Tk()
        root.withdraw()

        # 2) conecta a la BD
        conn = conectar_sql_server("DB_DATABASE")
        if conn is None:
            sys.exit("No se pudo conectar a la BD.")

        # 3) extrae el user_id que viene justo después del flag
        idx = sys.argv.index("--actualizar-datos")
        try:
            user_id = int(sys.argv[idx + 1])
        except Exception:
            sys.exit("Falta user_id tras --actualizar-dato")
            
        # 4) llama a tu función
        actualizar_usuario(root, conn, user_id)

        # 5) arranca el loop de Tk para CustomTkinter
        root.mainloop()
        sys.exit(0)
        
if "--modificar-radicado" in sys.argv:
    # Configura el tema según la preferencia guardada
    apply_ctk_theme_from_settings()

    # Crea el root (no lo ocultes)
    root = tk.Tk()

    # Conectar a la base de datos
    conn = conectar_sql_server("DB_DATABASE")
    if conn is None:
        sys.exit("No se pudo conectar a la BD.")
    print("Conexión exitosa")

    # Extrae el user_id que viene justo después del flag
    idx = sys.argv.index("--modificar-radicado")
    try:
        user_id = int(sys.argv[idx + 1])
    except Exception:
        sys.exit("Falta user_id tras --modificar-radicado")
    
    # Llama a tu función para modificar radicado
    modificar_radicado(root, conn, user_id)

    # Inicia el loop de Tkinter
    root.mainloop()
    sys.exit(0)

        
if "--desactivar-usuario" in sys.argv:
        # Configura tema según la preferencia guardada
        apply_ctk_theme_from_settings()
        # 1) crea el root oculto
        root = tk.Tk()
        root.withdraw()

        # 2) conecta a la BD
        conn = conectar_sql_server("DB_DATABASE")
        if conn is None:
            sys.exit("No se pudo conectar a la BD.")

        # 3) extrae el user_id que viene justo después del flag
        idx = sys.argv.index("--desactivar-usuario")
            
        # 4) llama a tu función
        modificar_estado_usuario(root, conn)

        # 5) arranca el loop de Tk para CustomTkinter
        root.mainloop()
        sys.exit(0)


def resource_path(rel_path):
    """
    Devuelve la ruta absoluta a `rel_path`, 
    usando _MEIPASS cuando PyInstaller congela la app.
    """
    base = getattr(sys, '_MEIPASS', os.path.dirname(__file__))
    return os.path.join(base, rel_path)

def safe_destroy(win):
    """
    Cancela todos los callbacks `after` pendientes de Tkinter 
    y destruye la ventana sin errores de comando Tcl.
    """
    try:
        for aid in win.tk.call('after', 'info'):
            try:
                win.after_cancel(aid)
            except Exception:
                pass
    except Exception:
        pass
    # desconectar triggers de CTk
    for child in win.winfo_children():
        try:
            child.destroy()
        except:
            pass
    try:
        win.destroy()
    except:
        pass
    

def make_semitransparent_image(w, h, radius=20, alpha=150, rgb_color=(0,0,0)):
    """
    Crea una imagen RGBA de fondo semitransparente redondeado
    usando `rgb_color` para el canal RGB y `alpha` para la opacidad.
    """
    # Partimos de un lienzo completamente transparente
    img = Image.new("RGBA", (w, h), (0,0,0,0))
    mask = Image.new("L", (w, h), 0)                   # máscara de niveles de gris
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle((0, 0, w, h), radius=radius, fill=alpha)

    # Creamos un recuadro RGBA (rgb_color + alpha)
    rgba_color = (rgb_color[0], rgb_color[1], rgb_color[2], alpha)
    rect = Image.new("RGBA", (w, h), rgba_color)

    # Pegamos el rectángulo usando la máscara
    img.paste(rect, (0,0), mask)
    return img


# -----------------------------------------------------------------------------
# styled_window: crea un Toplevel con fondo y panel redondeado/transparente
# -----------------------------------------------------------------------------
def styled_window(parent, title, bg_file, width, height, use_light_panel=False):
    """
    Si use_light_panel=True, el panel se pintará en blanco semitransparente.
    Si es False, en negro semitransparente.
    """
    win = ctk.CTkToplevel(parent)
    win.title(title)
    win.geometry(f"{width}x{height}")
    win.resizable(False, False)

    # … (carga del fondo igual a como lo tenías) …
    canvas = tk.Canvas(win, width=width, height=height, highlightthickness=0)
    canvas.place(x=0, y=0)
    bg_path = resource_path(bg_file)
    if os.path.exists(bg_path):
        bg = Image.open(bg_path).resize((width, height), Image.LANCZOS)
        tk_bg = ImageTk.PhotoImage(bg)
        canvas.create_image(0, 0, anchor="nw", image=tk_bg)
        canvas.bg_img = tk_bg

    # 3) panel semitransparente redondeado
    panel_w, panel_h = int(width * 0.8), int(height * 0.8)
    panel_x = (width - panel_w) // 2
    panel_y = (height - panel_h) // 2

    # Elegimos blanco o negro según use_light_panel:
    if use_light_panel:
        rgb = (255, 255, 255)
    else:
        rgb = (0, 0, 0)

    semi = make_semitransparent_image(panel_w, panel_h, radius=20, alpha=150, rgb_color=rgb)
    tk_semi = ImageTk.PhotoImage(semi)
    canvas.create_image(panel_x, panel_y, anchor="nw", image=tk_semi)
    canvas.semi_img = tk_semi

    # 4) frame transparente para tus widgets
    content = ctk.CTkFrame(win, fg_color="transparent", width=panel_w, height=panel_h)
    content.place(x=panel_x, y=panel_y)
    content.pack_propagate(False)

    return win, content


class BlurFrame(QtWidgets.QFrame):
    def __init__(self, bg_blur_pixmap: QtGui.QPixmap,
                 corner_radius: int = 20,
                 overlay_alpha: int = 155,
                 overlay_rgb: tuple = (0,0,0),
                 *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.bg_blur = bg_blur_pixmap
        self.corner_radius = corner_radius
        # overlay_rgb es (r,g,b) de 0–255
        self.overlay_color = QColor(overlay_rgb[0], overlay_rgb[1], overlay_rgb[2], overlay_alpha)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        self.setAutoFillBackground(False)

        # Máscara redondeada inicial
        path = QPainterPath()
        path.addRoundedRect(QRectF(self.rect()), self.corner_radius, self.corner_radius)
        self.setMask(QRegion(path.toFillPolygon().toPolygon()))

    def set_overlay_color(self, rgb_color: tuple, alpha: int = 155):
        """
        Permite cambiar dinámicamente el color de la capa semitransparente.
        """
        self.overlay_color = QColor(rgb_color[0], rgb_color[1], rgb_color[2], alpha)
        self.update() 

    def resizeEvent(self, event: QtGui.QResizeEvent):
        """
        Cada vez que el frame cambia de tamaño, actualizamos la máscara para mantener
        las esquinas redondeadas en la nueva geometría.
        """
        path = QPainterPath()
        path.addRoundedRect(QRectF(self.rect()), self.corner_radius, self.corner_radius)
        self.setMask(QRegion(path.toFillPolygon().toPolygon()))
        super().resizeEvent(event)

    def paintEvent(self, event: QtGui.QPaintEvent):
        """
        Sobrescribimos paintEvent para:
        1. Recortar en bg_blur la región que corresponde a la posición del frame.
        2. Dibujar ese sub-pixmap en las coordenadas locales del frame.
        3. Pintar una capa semitransparente encima (overlay).
        """
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # 1) Crear ruta redondeada y establecerla como clip
        path = QPainterPath()
        path.addRoundedRect(QRectF(self.rect()), self.corner_radius, self.corner_radius)
        painter.setClipPath(path)

        # 2) Calcular el offset (posición) del frame respecto a la ventana principal
        top_left_global = self.mapTo(self.window(), QPoint(0, 0))
        x0, y0 = top_left_global.x(), top_left_global.y()
        w, h = self.width(), self.height()

        # 3) Asegurarnos de no pasarnos de los límites del pixmap completo
        full_pixmap = self.bg_blur
        x0_clamped = max(0, min(x0, full_pixmap.width()))
        y0_clamped = max(0, min(y0, full_pixmap.height()))
        w_clamped = min(w, full_pixmap.width() - x0_clamped)
        h_clamped = min(h, full_pixmap.height() - y0_clamped)

        if w_clamped > 0 and h_clamped > 0:
            # 4) Extraer el sub-pixmap (región borrosa) correspondiente al frame
            cropped = full_pixmap.copy(x0_clamped, y0_clamped, w_clamped, h_clamped)
            # 5) Dibujar ese sub-pixmap en (0, 0) de este frame
            painter.drawPixmap(0, 0, w_clamped, h_clamped, cropped)

        # 6) Dibujar overlay semitransparente (negro con alpha)
        painter.fillPath(path, self.overlay_color)
        painter.end()

        # 7) Llamar al paintEvent de la clase base para que luego se pinten hijos, bordes, etc.
        super().paintEvent(event)
        
class DashboardWindow(QtWidgets.QMainWindow):
    def __init__(self, user_id, first_name, last_name, *, theme: str = "dark", parent=None):
        super().__init__(parent)
        self.theme = theme
        settings = QtCore.QSettings("Procesos Y Servicios", "CapturadorDeDatos")
        self.theme = settings.value("theme", self.theme)   # lee "dark" o "light"
        self.icon_moon = QtGui.QIcon(resource_path("moon.png"))
        self.icon_sun  = QtGui.QIcon(resource_path("sun.png"))
        import tkinter as tk
        self.user_id = user_id
        self.first_name = first_name
        self.last_name = last_name
 # ——— HABILITAR DIBUJO DEL FONDO ———        
        self.setAutoFillBackground(True)

        # ——— CARGAR TU IMAGEN DE FONDO ———
        central = QtWidgets.QWidget()
        central.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        central.setAutoFillBackground(False)
        self.setCentralWidget(central)
        central.setAutoFillBackground(True)

        # ——— CARGAR IMAGEN DE FONDO SEGÚN EL TEMA ———
        # Asumimos que self.theme == "dark" ó "light"
        bg_file = "FondoDashboardDark.png" if self.theme == "dark" else "FondoDashboardWhite.png"
        bg_path = resource_path(bg_file)

        if os.path.exists(bg_path):
            # 1) Escalar el pixmap al tamaño actual de la ventana
            pix = QtGui.QPixmap(bg_path).scaled(
                self.size(),
                QtCore.Qt.KeepAspectRatioByExpanding,
                QtCore.Qt.SmoothTransformation
            )

            # Aplicar pix como fondo de la ventana (solo este fragmento, sin asignar bg_blur_pix aquí)
            palette = QtGui.QPalette()
            palette.setBrush(QtGui.QPalette.Window, QtGui.QBrush(pix))
            self.setPalette(palette)
            central.setPalette(palette)
        else:
            # …
            self.setAutoFillBackground(False)
            central.setPalette(QtGui.QPalette())

        # … luego, más abajo, cuando generas el blur:
        pix = QtGui.QPixmap(bg_path).scaled(
            self.size(),
            QtCore.Qt.KeepAspectRatioByExpanding,
            QtCore.Qt.SmoothTransformation
        )
        scene = QGraphicsScene()
        item = QGraphicsPixmapItem(pix)
        blur = QGraphicsBlurEffect()
        blur.setBlurRadius(15)
        item.setGraphicsEffect(blur)
        scene.addItem(item)

        img = QImage(pix.size(), QImage.Format_ARGB32_Premultiplied)
        img.fill(QtCore.Qt.transparent)
        painter = QPainter(img)
        scene.render(painter)
        painter.end()

        # → Aquí sí puedes asignar correctamente:
        self.bg_blur_pix = QtGui.QPixmap.fromImage(img)
            
        # ——— AHORA PUEDES SEGUIR CON TU CÓDIGO NORMAL ——–
        self.setWindowFlag(QtCore.Qt.WindowMaximizeButtonHint, False)
        self.center_on_screen()

        # Después, inicializa CTk y el resto según el tema guardado
        apply_ctk_theme_from_settings()

        # Conexión a BD
        self.conn = conectar_sql_server("DB_DATABASE")
        if not self.conn:
            QtWidgets.QMessageBox.critical(
                None, "Error", "No se pudo conectar a la base de datos."
            )
            sys.exit(1)
            
        # Configuración de la ventana
        self.setWindowTitle("Dashboard · Capturador De Datos")
        self.resize(900, 900)
        self.center_on_screen()
            
        pix = QtGui.QPixmap(bg_path).scaled(
            self.size(),
            QtCore.Qt.KeepAspectRatioByExpanding,
            QtCore.Qt.SmoothTransformation
        )
        
        # Aplicar el fondo normal
        palette = QtGui.QPalette()
        palette.setBrush(QtGui.QPalette.Window, QtGui.QBrush(pix))
        self.setPalette(palette)

        # Crear la versión borrosa del mismo pixmap
        scene = QGraphicsScene()
        item = QGraphicsPixmapItem(pix)
        blur = QGraphicsBlurEffect()
        blur.setBlurRadius(15)
        item.setGraphicsEffect(blur)
        scene.addItem(item)

        img = QImage(pix.size(), QImage.Format_ARGB32_Premultiplied)
        img.fill(QtCore.Qt.transparent)
        painter = QPainter(img)
        scene.render(painter)
        painter.end()

        self.bg_blur_pix = QtGui.QPixmap.fromImage(img)

        # ——————————————————————————————
        # 2) Panel central semitransparente
        # ——————————————————————————————
        central = QtWidgets.QWidget()
        central.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        central.setAutoFillBackground(False)
        self.setCentralWidget(central)
        v_layout = QtWidgets.QVBoxLayout(central)
        v_layout.setContentsMargins(0, 0, 0, 0)
        v_layout.setSpacing(0)

        self.panel = BlurFrame(self.bg_blur_pix, corner_radius=20)
        self.panel.setObjectName("dashboardPanel")
        # Quitas el border-radius de la stylesheet o lo dejas solo como “seguro”:
        self.panel.setStyleSheet("""
            QFrame#dashboardPanel { border: none; }
        """)
        # Aplicas tu shadow y layout como antes
        shadow = QGraphicsDropShadowEffect(self.panel)
        shadow.setBlurRadius(20)
        shadow.setOffset(0, 4)
        shadow.setColor(QtGui.QColor(0, 0, 0, 160))
        self.panel.setGraphicsEffect(shadow)

        # Blur (frosted glass)
        blur = QGraphicsBlurEffect(self.panel)
        blur.setBlurRadius(15)
        p_layout = QtWidgets.QVBoxLayout(self.panel)
        p_layout.setContentsMargins(50, 50, 50, 50)   # más padding interno
        p_layout.setSpacing(16)
        
        p_layout.addSpacing(40)

        v_layout.addStretch()
        v_layout.addWidget(self.panel, alignment=QtCore.Qt.AlignCenter)
        v_layout.addStretch()

        # ——————————————————————————————
        # 3) Encabezado con saludo
        # ——————————————————————————————
        self.lbl_saludo = QtWidgets.QLabel(f"Bienvenido, {first_name} {last_name}")
        self.lbl_saludo.setAlignment(QtCore.Qt.AlignCenter)
        # Color inicial se ajustará luego en apply_theme
        self.lbl_saludo.setStyleSheet("""
            color: #FFFFFF;              /* placeholder, será reemplazado */
            font-size: 28px;
            font-weight: 600;
            background: transparent;
        """)
        p_layout.addWidget(self.lbl_saludo)

        # ——————————————————————————————
        # 4) Selector de rol
        # ——————————————————————————————
        class PopupOnClickFilter(QtCore.QObject):
            def __init__(self, combo):
                super().__init__(combo)
                self.combo = combo

            def eventFilter(self, obj, ev):
                if ev.type() == QtCore.QEvent.MouseButtonPress:
                    self.combo.showPopup()
                    return True
                return False

        # …

        self.cmb_role = QtWidgets.QComboBox()
        self.cmb_role.setEditable(True)
        self.cmb_role.setFixedWidth(220)  # espacio suficiente para el texto
        le = self.cmb_role.lineEdit()
        le.setAlignment(QtCore.Qt.AlignCenter)
        le.setReadOnly(True)
        # Márgenes extra para evitar que el texto se corte
        le.setTextMargins(10, 0, 10, 0)

        # Instalamos el filtro
        f = PopupOnClickFilter(self.cmb_role)
        le.installEventFilter(f)
        
        p_layout.addWidget(self.cmb_role, alignment=QtCore.Qt.AlignCenter)


        
        p_layout.addSpacing(40)

        # Cargar roles desde BD
        self._load_roles()

        # ——————————————————————————————
        # 5) Botones de acciones según rol
        # ——————————————————————————————
        self.buttons_layout = QtWidgets.QVBoxLayout()
        self.buttons_layout.setSpacing(20)
        p_layout.addLayout(self.buttons_layout)
        p_layout.addSpacing(40)
        self.cmb_role.currentIndexChanged.connect(self._refresh_buttons)
        self._refresh_buttons()  # inicializa botones para el rol por defecto

        # ——————————————————————————————
        # 6) Logout
        # ——————————————————————————————
        
        btn_logout = QtWidgets.QPushButton("Cerrar Sesión")
        btn_logout.setFixedSize(120, 35)
        btn_logout.setStyleSheet("""
            QPushButton {
                background-color: #FF6F61;
                color: white;
                border-radius: 20px;
                font-size: 16px;
                padding: 8px 16px;
                min-width: 120px;
                min-height: 35px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #E14D50;
            }
        """)
        btn_logout.clicked.connect(self.on_logout)
        hbox_logout = QtWidgets.QHBoxLayout()
        hbox_logout.addStretch()
        hbox_logout.addWidget(btn_logout)
        hbox_logout.addStretch()
        p_layout.addLayout(hbox_logout)
        
        self.theme_btn = QtWidgets.QPushButton(self)
        self.theme_btn.setToolTip("Cambiar tema día/noche")
        self.theme_btn.setCursor(QtCore.Qt.PointingHandCursor)
        self.theme_btn.setFixedSize(40, 40)
        self.theme_btn.setIconSize(QtCore.QSize(40, 40))
        self.theme_btn.setFlat(True)
        self.theme_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: none;
            }
        """)

        # Posición y estado inicial
        self.theme_btn.move(self.width() - 60, self.height() - 60)
        self.theme_btn.setIcon(self.icon_moon if self.theme == "dark" else self.icon_sun)
        self.theme_btn.clicked.connect(self.toggle_theme)
        self.theme_btn.raise_()
        
        self.apply_theme(self.theme)

        # Para reposicionar al redimensionar
        self.resizeEvent = self._on_resize
        
    def apply_theme(self, theme: str):
        """
        Aplica el stylesheet global según 'dark' o 'light', actualiza el icono
        del botón y cambia la imagen de fondo a FondoDashboardDark.png o FondoDashboardWhite.png.
        """
        from PyQt5 import QtGui, QtCore
        from PyQt5.QtWidgets import QGraphicsScene, QGraphicsPixmapItem, QGraphicsBlurEffect
        from PyQt5.QtGui import QPainter, QImage, QPalette, QBrush
        
        bg_file = "FondoDashboardDark.png" if theme == "dark" else "FondoDashboardWhite.png"
        app = QtWidgets.QApplication.instance()

        # — Stylesheet para tema oscuro —
        dark_ss = """
        QWidget { background-color: #2b2b2b; color: #f0f0f0; }
        QFrame#dashboardPanel { background-color: rgba(0,0,0,140); }
        QPushButton { 
            background-color: #2B7DBF; color: #fff; border-radius:20px; 
            padding:8px 16px; font-weight:bold; 
        }
        QPushButton:hover { background-color: #3291DE; }
        QComboBox { background-color: rgba(255,255,255,200); color: #000; }
        """

        # — Stylesheet para tema claro —
        light_ss = """
        QWidget { background-color: #f4f4f9; color: #000; }
        QFrame#dashboardPanel { background-color: rgba(255,255,255,200); border:1px solid rgba(0,0,0,25%); }
        QPushButton { 
            background-color: #1E90FF; color: #fff; border-radius:20px; 
            padding:8px 16px; font-weight:bold; 
        }
        QPushButton:hover { background-color: #006FDE; }
        QComboBox { background-color: rgba(0,0,0,10%); color: #000; }
        """
        if hasattr(self, "cmb_role"):
          
            if theme == "light":

                self.cmb_role.setStyleSheet("""
                    QComboBox {
                        background-color: rgba(0, 0, 0, 120);
                        color: #FFFFFF;
                        border-radius: 10px;
                        padding: 4px 20px;
                        font-size: 14px;
                        font-weight: bold;
                    }
                    QComboBox::drop-down { border: none; }
                    QComboBox QAbstractItemView {
                        background-color: #2b2b2b;
                        color: #FFFFFF;
                        selection-background-color: #454545;
                    }
                """)

            else:

                self.cmb_role.setStyleSheet("""
                    QComboBox {
                        background-color: rgba(255, 255, 255, 180);
                        color: #000000;
                        border-radius: 10px;
                        padding: 4px 20px;
                        font-size: 14px;
                        font-weight: bold;
                    }
                    QComboBox::drop-down { border: none; }
                    QComboBox QAbstractItemView {
                        background-color: #E0E0E0;
                        color: #000000;
                        selection-background-color: #CCCCCC;
                    }
                """)
        if hasattr(self, "lbl_saludo"):
            if theme == "light":
                # Texto negro sobre fondo claro
                self.lbl_saludo.setStyleSheet("""
                    color: #000000;
                    font-size: 28px;
                    font-weight: 600;
                    background: transparent;
                """)
            else:
                # Texto blanco sobre fondo oscuro
                self.lbl_saludo.setStyleSheet("""
                    color: #FFFFFF;
                    font-size: 28px;
                    font-weight: 600;
                    background: transparent;
                """)
                
        bg_path = resource_path(bg_file)
        if os.path.exists(bg_path):
            pix = QtGui.QPixmap(bg_path).scaled(
                self.size(),
                QtCore.Qt.KeepAspectRatioByExpanding,
                QtCore.Qt.SmoothTransformation
            )

            # 2) Actualizar la paleta de la ventana
            palette = QtGui.QPalette()
            palette.setBrush(QtGui.QPalette.Window, QtGui.QBrush(pix))
            self.setPalette(palette)

            # 3) Volver a generar el pixmap borroso
            scene = QGraphicsScene()
            item  = QGraphicsPixmapItem(pix)
            blur = QGraphicsBlurEffect()
            blur.setBlurRadius(15)
            item.setGraphicsEffect(blur)
            scene.addItem(item)

            img = QImage(pix.size(), QImage.Format_ARGB32_Premultiplied)
            img.fill(QtCore.Qt.transparent)
            painter = QPainter(img)
            scene.render(painter)
            painter.end()

            # 4) Asignar el nuevo blur al dashboard y al panel
            self.bg_blur_pix = QtGui.QPixmap.fromImage(img)
            if hasattr(self, 'panel'):
                self.panel.bg_blur = self.bg_blur_pix
                self.panel.update()
        # Elige stylesheet e icono según el tema
        if theme == "dark":
            app.setStyleSheet(dark_ss)
            self.theme_btn.setIcon(self.icon_moon)
            bg_file = "FondoDashboardDark.png"
        else:
            app.setStyleSheet(light_ss)
            self.theme_btn.setIcon(self.icon_sun)
            bg_file = "FondoDashboardWhite.png"
            
        if theme == "dark":
            self.panel.set_overlay_color((0,0,0), alpha=155)
        else:
            # En tema claro, usamos un panel blanco semitransparente
            self.panel.set_overlay_color((255,255,255), alpha=155)

        # — Ahora cargamos el PNG correspondiente y lo escalamos al tamaño actual —
        from PyQt5 import QtGui, QtCore

        bg_path = resource_path(bg_file)
        if os.path.exists(bg_path):
            pix = QtGui.QPixmap(bg_path).scaled(
                self.size(),
                QtCore.Qt.KeepAspectRatioByExpanding,
                QtCore.Qt.SmoothTransformation
            )
            palette = QtGui.QPalette()
            palette.setBrush(QtGui.QPalette.Window, QtGui.QBrush(pix))
            self.setPalette(palette)

            # También ponemos un estilo de CSS “por si acaso” (opcional):
            bg_url = bg_path.replace("\\", "/")
            self.setStyleSheet(f"""
                QMainWindow {{
                    background-image: url("{bg_url}");
                    background-position: center;
                    background-repeat: no-repeat;
                }}
            """)
        else:
            # Si no existe el archivo, al menos quitamos cualquier fondo anterior
            # para que no se quede colgado.
            self.setAutoFillBackground(False)
            self.setStyleSheet("")

    def toggle_theme(self):
        """Invierte el tema, lo reaplica y lo guarda en QSettings."""
        self.theme = "light" if self.theme == "dark" else "dark"
        self.apply_theme(self.theme)
        settings = QtCore.QSettings("Procesos Y Servicios", "CapturadorDeDatos")
        settings.setValue("theme", self.theme)


    def _on_resize(self, event):
        """Reposiciona el botón de tema al cambiar tamaño."""
        self.theme_btn.move(self.width() - 60, self.height() - 60)
        return super().resizeEvent(event)

    def center_on_screen(self):
        """Centra la ventana en la pantalla."""
        screen = QtWidgets.QApplication.primaryScreen().availableGeometry()
        size = self.geometry()
        x = (screen.width() - size.width()) // 2
        y = (screen.height() - size.height()) // 2
        self.move(x, y)

    def _load_roles(self):
        """Carga los roles del usuario y llena el combo."""
        cur = self.conn.cursor()
        cur.execute("""
            SELECT R.ID, R.NAME
              FROM USER_ROLES UR
              JOIN ROL R ON UR.ROL_ID = R.ID
             WHERE UR.USER_ID = %s
        """, (self.user_id,))
        rows = cur.fetchall()
        cur.close()

        # Map name->id
        self.role_map = {name: rid for rid, name in rows}
        self.cmb_role.clear()
        self.cmb_role.addItems(list(self.role_map.keys()))

    def _refresh_buttons(self):
        """Recrea los botones de acción cuando cambia de rol."""
        # Limpia layout
        while self.buttons_layout.count():
            w = self.buttons_layout.takeAt(0).widget()
            if w:
                w.deleteLater()

        name = self.cmb_role.currentText()
        rid = self.role_map.get(name)
        # Definición de botones por rol
        btns_by_role = {
            1: [
                ("Capturar Datos",                          self.on_iniciar_digitacion),
                ("Validar Calidad",                         self.on_iniciar_calidad),
                ("Cargar Paquete",                          self.on_cargar_paquete),
                ("Crear Usuario",                           self.on_crear_usuario),
                ("Actualizar Datos",                        self.on_actualizar_datos),
                ("Modificar Datos Capturados",              self.on_modificar_radicado),
                ("Ver Progreso",                            self.on_ver_progreso),
                ("Desactivar Usuario",                      self.on_modificar_estado_usuario),
                ("Exportar Datos",                          self.on_exportar_paquete),
            ],
            2: [
                ("Capturar Datos",                          self.on_iniciar_digitacion),
                ("Actualizar Datos",                        self.on_actualizar_datos),
                ("Modificar Datos Capturados",              self.on_modificar_radicado),
                ("Ver Progreso",                            self.on_ver_progreso),
                ("Exportar Datos",                          self.on_exportar_paquete),
            ],
            3: [
                ("Validar Calidad",                         self.on_iniciar_calidad),
                ("Actualizar Datos",                        self.on_actualizar_datos),
                ("Modificar Datos Capturados",              self.on_modificar_radicado),
                ("Ver Progreso",                            self.on_ver_progreso),
                ("Exportar Datos",                          self.on_exportar_paquete),
            ],
        }

        for text, slot in btns_by_role.get(rid, []):
            btn = QtWidgets.QPushButton(text)
            btn.setFixedHeight(40)
            btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: #2B7DBF;
                    color: white;
                    border-radius: 20px;
                    font-size: 16px;
                    padding: 8px 16px;
                    font-weight: bold;
                }}
                QPushButton:hover {{
                    background-color: #3291DE;
                }}
            """)
            btn.clicked.connect(slot)
            self.buttons_layout.addWidget(btn)

    def on_logout(self):
        """Cierra el dashboard y relanza el login."""
        settings = QtCore.QSettings("Procesos Y Servicios", "CapturadorDeDatos")
        settings.setValue("theme", self.theme)
        self.close()
        script = resource_path("login_app.py")
        subprocess.Popen([sys.executable, script], cwd=os.path.dirname(script))

    # ——————————————————————————————
    # Métodos de acción (stubs; implementa tu lógica aquí)
    # ——————————————————————————————
    def cargar_paquete(self):

        # -----------------------------------------
        # Asegura un root de Tkinter para los Toplevel
        # -----------------------------------------
        theme = apply_ctk_theme_from_settings()
        if not hasattr(self, '_tk_root'):
            self._tk_root = tk.Tk()
            self._tk_root.withdraw()

        # 0) Selección de Tipo de Paquete
        def elegir_tipo():
            win = ctk.CTkToplevel(self._tk_root)
            win.title("Seleccione Tipo de Paquete")
            # Evitamos colores con transparencia (Tk no los soporta)
            bg = "#2f2f2f" if theme == "dark" else "#f0f0f0"
            fg = "white" if theme == "dark" else "black"
            win.configure(fg_color=bg)
            tipo_var = tk.StringVar(master=self._tk_root, value="DIGITACION")
            aceptado = tk.BooleanVar(master=self._tk_root, value=False)
            ctk.CTkLabel(win, text="Tipo de Paquete:", text_color=fg,
                        fg_color=bg, font=("Arial",14,"bold")).pack(pady=10)
            opt_bg = "#2b2b2b" if theme == "light" else "#e0e0e0"
            opt_fg = "#FFFFFF" if theme == "light" else "#000000"

            ctk.CTkOptionMenu(
                win,
                values=["DIGITACION","CALIDAD"],
                variable=tipo_var,

                fg_color=opt_bg,
                text_color=opt_fg,
                button_color=opt_bg,
                button_hover_color=opt_bg,
                font=("Arial",14,"bold")
            ).pack(pady=5)
            def ok():
                aceptado.set(True)
                win.destroy()
            ctk.CTkButton(win, text="Aceptar", command=ok).pack(pady=10)
            win.grab_set()
            self._tk_root.wait_window(win)
            return tipo_var.get().upper() if aceptado.get() else None

        tipo = elegir_tipo()
        if not tipo:
            return

        # 1) Selección de archivo
        path = filedialog.askopenfilename(
            title="Selecciona archivo de paquete",
            filetypes=[("Excel/CSV", "*.xls* *.csv"), ("Todos", "*")]
        )
        if not path:
            return

        # 2) Leer DataFrame
        try:
            df = (pd.read_excel(path) if path.lower().endswith(('xls','xlsx'))
                else pd.read_csv(path))
        except Exception:
            return messagebox.showerror("Error lectura", "No se pudo leer el archivo.")

        # Normalize columns names
        df.columns = [c.strip().upper().replace(" ","_") for c in df.columns]
        required = {"RADICADO","NIT","RAZON_SOCIAL","FACTURA","VALOR_FACTURA",
                    "FECHA_RADICACION","ESTADO_FACTURA","IMAGEN","RADICADO_IMAGEN",
                    "LINEA","ID_ASIGNACION","ESTADO_PYS","OBSERVACION_PYS",
                    "LINEA_PYS","RANGOS","DEF"}
        faltan = required - set(df.columns)
        if faltan:
            return messagebox.showerror("Columnas faltantes", f"Faltan: {faltan}")
        if df.empty:
            return messagebox.showinfo("Sin datos","El archivo está vacío.")

        # 3) Utilidades de conversión
        def safe_int(x, default=0):
            try:
                return int(x)
            except Exception:
                return default
        def safe_str(x, length):
            return str(x).strip()[:length] if pd.notna(x) else None
        def safe_date(x):
            if pd.isna(x):
                return None
            if isinstance(x, (pd.Timestamp, datetime.date)):
                return x.date() if hasattr(x,'date') else x
            try:
                return pd.to_datetime(x).date()
            except Exception:
                return None

        # 4) Detectar duplicados
        rad = [safe_int(v) for v in df['RADICADO'] if pd.notna(v)]
        if rad:
            cur = self.conn.cursor()
            exist = set()
            MAX_PARAMS = 2000
            for i in range(0, len(rad), MAX_PARAMS):
                chunk = rad[i:i + MAX_PARAMS]
                placeholders = ','.join('%s' for _ in chunk)
                q = (
                    "SELECT RADICADO FROM ASIGNACION_TIPIFICACION "
                    f"WHERE RADICADO IN ({placeholders})"
                )
                cur.execute(q, chunk)
                exist.update(r[0] for r in cur.fetchall())
            cur.close()
            if exist:
                return messagebox.showerror("Duplicados", "Existen: %s" % exist)

        # 5) Calcular NUM_PAQUETE
        cur = self.conn.cursor()
        cur.execute(
            "SELECT ISNULL(MAX(NUM_PAQUETE),0) FROM ASIGNACION_TIPIFICACION WHERE TIPO_PAQUETE=%s",
            (tipo,)
        )
        ultimo = cur.fetchone()[0]
        cur.close()
        num_paquete = ultimo + 1

        # 6) Construir filas de inserción
        rows = []
        for _,r in df.iterrows():
            rows.append((
                safe_int(r['RADICADO']),
                safe_int(r['NIT']),
                safe_str(r['RAZON_SOCIAL'],255),
                safe_str(r['FACTURA'],30),
                safe_int(r['VALOR_FACTURA'],0),
                None,                  # FECHA_FACTURA
                safe_date(r['FECHA_RADICACION']),
                None,                  # TIPO_DOC_ID
                safe_int(r.get('NUM_DOC'), None),
                safe_str(r['ESTADO_FACTURA'],255),
                safe_str(r['IMAGEN'],2),
                safe_str(r['RADICADO_IMAGEN'],255),
                safe_str(r['LINEA'],255),
                safe_str(r['ID_ASIGNACION'],255),
                safe_str(r['ESTADO_PYS'],255),
                safe_str(r['OBSERVACION_PYS'],255),
                safe_str(r['LINEA_PYS'],255),
                safe_str(r['RANGOS'],255),
                safe_str(r['DEF'],255),
                tipo,
                1,
                num_paquete
            ))

        # 7) Insertar en BD con barra de progreso
        sql_insert = '''
        INSERT INTO ASIGNACION_TIPIFICACION
        (RADICADO,NIT,RAZON_SOCIAL,FACTURA,VALOR_FACTURA,
        FECHA_FACTURA,FECHA_RADICACION,TIPO_DOC_ID,
        NUM_DOC,ESTADO_FACTURA,IMAGEN,RADICADO_IMAGEN,
        LINEA,ID_ASIGNACION,ESTADO_PYS,OBSERVACION_PYS,
        LINEA_PYS,RANGOS,DEF,TIPO_PAQUETE,STATUS_ID,NUM_PAQUETE)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)'''

        progress_win = ctk.CTkToplevel(self._tk_root)
        progress_win.title("Cargando paquete")
        pb = ctk.CTkProgressBar(progress_win, width=300)
        pb.pack(padx=20, pady=20)
        pb.set(0)
        progress_win.update_idletasks()

        cur = self.conn.cursor()
        try:
            try:
                cur.fast_executemany = True
            except Exception:
                pass
            cur.execute("SET IDENTITY_INSERT ASIGNACION_TIPIFICACION ON;")
            batch_size = 500
            total = len(rows)
            for i in range(0, total, batch_size):
                batch = rows[i:i + batch_size]
                cur.executemany(sql_insert, batch)
                pb.set((i + len(batch)) / total)
                progress_win.update_idletasks()
            cur.execute("SET IDENTITY_INSERT ASIGNACION_TIPIFICACION OFF;")
            self.conn.commit()
        except Exception as e:
            self.conn.rollback()
            cur.close()
            progress_win.destroy()
            return messagebox.showerror("Error inserción", str(e))
        cur.close()
        progress_win.destroy()

        # 8) Confirmar conteo
        cur2 = self.conn.cursor()
        cur2.execute(
            "SELECT COUNT(*) FROM ASIGNACION_TIPIFICACION WHERE NUM_PAQUETE=%s AND TIPO_PAQUETE=%s",
            (num_paquete, tipo)
        )
        cnt = cur2.fetchone()[0]
        cur2.close()
        messagebox.showinfo("Carga completa",
                            f"Paquete {num_paquete} ({tipo}) insertado. Filas: {cnt}")

        # 9) Selección de campos
        sel2 = ctk.CTkToplevel(self._tk_root)
        sel2.title(f"Paquete {num_paquete}: Campos")
        campos = [
            "FECHA_SERVICIO","FECHA_FINAL","TIPO_DOC_ID","NUM_DOC",
            "DIAGNOSTICO","AUTORIZACION","CODIGO_SERVICIO",
            "CANTIDAD","VLR_UNITARIO","COPAGO","OBSERVACION"
        ]
        vars_chk = {}
        for c in campos:
            vars_chk[c] = tk.BooleanVar(master=self._tk_root, value=True)
            ctk.CTkCheckBox(sel2, text=c, variable=vars_chk[c]).pack(anchor="w", padx=20)
        def guardar_campos():
            data = [(num_paquete,c,tipo) for c,v in vars_chk.items() if v.get()]
            c2 = self.conn.cursor()
            c2.executemany(
                "INSERT INTO PAQUETE_CAMPOS (NUM_PAQUETE,campo,tipo_paquete) VALUES (%s,%s,%s)",
                data
            )
            self.conn.commit()
            c2.close()
            sel2.destroy()
            messagebox.showinfo("Campos guardados", f"{len(data)} campos registrados.")
        ctk.CTkButton(sel2, text="Guardar", command=guardar_campos).pack(pady=10)
        sel2.grab_set()
        self._tk_root.wait_window(sel2)



        
    def on_cargar_paquete(self):    
        self.cargar_paquete()
        pass
        
    def on_crear_usuario(self):
        import subprocess
        script = os.path.abspath(sys.argv[0])
        subprocess.Popen([sys.executable, script, "--crear-usuario"])

    def on_iniciar_digitacion(self):
        import subprocess, os, sys
        script = os.path.abspath(sys.argv[0])
        subprocess.Popen([
            sys.executable,
            script,
            "--iniciar-tipificacion",
            str(self.user_id)          # <–– Aquí debe ir el user_id
        ])
        pass

    def on_iniciar_calidad(self):
        import subprocess, os, sys
        script = os.path.abspath(sys.argv[0])
        subprocess.Popen([
            sys.executable,
            script,
            "--iniciar-calidad",
            str(self.user_id)          # <–– Aquí debe ir el user_id
        ])
        pass

    def on_ver_progreso(self):
        import subprocess, os, sys
        script = os.path.abspath(sys.argv[0])
        subprocess.Popen([
            sys.executable,
            script,
            "--ver-progreso",
            str(self)          # <–– Aquí debe ir el user_id
        ])
        pass

    def on_exportar_paquete(self):
        import subprocess, os, sys
        script = os.path.abspath(sys.argv[0])
        subprocess.Popen([
            sys.executable,
            script,
            "--exportar-paquete",
            str(self.user_id)          # <–– Aquí debe ir el user_id
        ])
        pass

    def on_actualizar_datos(self):
        import subprocess, os, sys
        script = os.path.abspath(sys.argv[0])
        subprocess.Popen([
            sys.executable,
            script,
            "--actualizar-datos",
            str(self.user_id)          # <–– Aquí debe ir el user_id
        ])
        pass

    def on_modificar_estado_usuario(self):
        import subprocess, os, sys
        script = os.path.abspath(sys.argv[0])
        subprocess.Popen([
            sys.executable,
            script,
            "--desactivar-usuario",
            str(self.user_id)          # <–– Aquí debe ir el user_id
        ])
        pass
    
    def on_modificar_radicado(self):
        import subprocess, os, sys
        script = os.path.abspath(sys.argv[0])
        subprocess.Popen([
            sys.executable,
            script,
            "--modificar-radicado",
            str(self.user_id)          # <–– Aquí debe ir el user_id
        ])
        pass


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    if len(sys.argv) != 4:
        print("Uso: python dashboard.py <user_id> <first_name> <last_name>")
        sys.exit(1)
    _, uid, fn, ln = sys.argv
    window = DashboardWindow(int(uid), fn, ln)
    window.show()
    sys.exit(app.exec_())
