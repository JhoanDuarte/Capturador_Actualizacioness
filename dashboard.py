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
from tkinter import filedialog, messagebox, ttk

# — Terceros —
import bcrypt
import cairosvg
import customtkinter as ctk  # sólo esto para CustomTkinter
from PIL import Image
import pandas as pd
import requests
from tkcalendar import DateEntry
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageBreak, Paragraph
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

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



# Tu tema y apariencia
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

def load_icon_from_url(url, size):
    resp = requests.get(url)
    resp.raise_for_status()
    # convierte SVG bytes a PNG bytes
    png_bytes = cairosvg.svg2png(bytestring=resp.content,
                                 output_width=size[0],
                                 output_height=size[1])
    img = Image.open(BytesIO(png_bytes))
    return ctk.CTkImage(light_image=img, dark_image=img, size=size)


def cargar_paquete(root, conn):
    import traceback
    import tkinter as tk
    from tkinter import filedialog, messagebox
    import customtkinter as ctk
    
    # 0) Selección de Tipo de Paquete
    sel = ctk.CTkToplevel(root)
    sel.title("Seleccione Tipo de Paquete")
    tipo_paquete_var = tk.StringVar(value="DIGITACION")
    ctk.CTkLabel(sel, text="Tipo de Paquete:").pack(pady=10)
    ctk.CTkOptionMenu(sel,
        values=["DIGITACION", "CALIDAD"],
        variable=tipo_paquete_var
    ).pack(pady=5)
    ctk.CTkButton(sel, text="Aceptar", command=sel.destroy).pack(pady=10)
    sel.grab_set()
    root.wait_window(sel)
    tipo_paquete = tipo_paquete_var.get()

    # 1) Definir encabezados que esperamos
    expected_headers = {
        "RADICADO", "NIT", "RAZON_SOCIAL", "FACTURA",
        "VALOR_FACTURA", "FECHA RADICACION",
        "ESTADO_FACTURA", "IMAGEN",
        "RADICADO_IMAGEN", "LINEA", "ID ASIGNACION",
        "ESTADO PYS", "OBSERVACION PYS", "LINEA PYS",
        "RANGOS", "Def"
    }

    # 2) Seleccionar archivo
    path = filedialog.askopenfilename(
        parent=root,
        title="Selecciona el archivo de paquete",
        filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("Todos", "*.*")]
    )
    if not path:
        return

    # 3) Leer con pandas
    try:
        if path.lower().endswith(('.xls', '.xlsx')):
            df = pd.read_excel(path)
        else:
            df = pd.read_csv(path)
    except Exception:
        messagebox.showerror("Error lectura", "No se pudo leer el archivo. Verifica formato.")
        return

    # 4) Validar encabezados
    actual_headers = set(df.columns)
    missing = expected_headers - actual_headers
    extra   = actual_headers - expected_headers
    if missing or extra:
        msg = []
        if missing:
            msg.append("Faltan columnas:\n  • " + "\n  • ".join(sorted(missing)))
        if extra:
            msg.append("Columnas inesperadas:\n  • " + "\n  • ".join(sorted(extra)))
        messagebox.showerror(
            "Encabezados incorrectos",
            "El archivo no tiene la estructura esperada.\n\n"
            "Esperados:\n  • " + "\n  • ".join(sorted(expected_headers)) +
            "\n\nEncontrados:\n  • " + "\n  • ".join(sorted(actual_headers)) +
            "\n\n" + "\n\n".join(msg)
        )
        return

    total = len(df)
    if total == 0:
        messagebox.showinfo("Sin datos", "El archivo está vacío.")
        return

    # 5) Detectar RADICADOS ya existentes
    try:
        rad_list = df["RADICADO"].dropna().astype(int).unique().tolist()
    except Exception:
        messagebox.showerror(
            "Error en RADICADO",
            "No se pudieron convertir los valores de RADICADO a enteros.\n"
            "Verifica que esa columna contenga sólo números."
        )
        return

    if rad_list:
        in_clause = ",".join(str(r) for r in rad_list)
        sql = f"SELECT RADICADO FROM ASIGNACION_TIPIFICACION WHERE RADICADO IN ({in_clause})"
        cur = conn.cursor()
        try:
            cur.execute(sql)
            existentes = sorted(r[0] for r in cur.fetchall())
        except Exception as e:
            messagebox.showerror(
                "Error al verificar duplicados",
                f"No se pudo comprobar duplicados:\n\n{e}"
            )
            cur.close()
            return
        cur.close()

        if existentes:
            messagebox.showerror(
                "Radicados duplicados",
                "Los siguientes radicados ya existen:\n\n" +
                "\n".join(f"• {r}" for r in existentes)
            )
            return

    # 6) Calcular NUM_PAQUETE → tomo el mayor de TODOS los registros
    cur = conn.cursor()
    cur.execute("SELECT ISNULL(MAX(NUM_PAQUETE), 0) FROM ASIGNACION_TIPIFICACION")
    ultimo = cur.fetchone()[0]   # si no hay ninguno, devuelve 0
    NUM_PAQUETE = ultimo + 1

    # DEBUG: imprimo en consola para verificar
    print(f"[DEBUG] Nuevo NUM_PAQUETE → {NUM_PAQUETE}")


    # 7) Activar IDENTITY_INSERT
    cur.execute("SET IDENTITY_INSERT ASIGNACION_TIPIFICACION ON;")

    # 8) Crear barra de progreso
    progress_window = ctk.CTkToplevel(root)
    progress_window.title("Progreso de Carga")
    progress_window.geometry("400x150")
    progress_label = ctk.CTkLabel(progress_window, text="Cargando registros...")
    progress_label.pack(pady=10)
    
    progress_bar = ctk.CTkProgressBar(progress_window, orientation="horizontal", mode="determinate")
    progress_bar.pack(padx=20, pady=20, fill="x")

    # Inicializar la barra de progreso en 0
    progress_bar.set(0)
    
    # 9) Preparamos la lista de parámetros
    params_list = []
    for idx, row in df.iterrows():
        try:
            # (mismos sanitizados que antes…)
            radicado       = int(row["RADICADO"])
            nit            = int(row["NIT"])
            razon          = str(row["RAZON_SOCIAL"])
            factura        = str(row["FACTURA"])
            valor_factura  = int(row["VALOR_FACTURA"])
            fecha_factura = None
            if "FECHA FACTURA" in df.columns:
                v = row["FECHA FACTURA"]
                fecha_factura = None if pd.isna(v) else str(v)
            num_doc = None
            tipo_doc_id = None
            if "TIPO DOC" in df.columns:
                v = row["TIPO DOC"]
                tipo_doc_id = None if pd.isna(v) else str(v)
            num_doc = None
            if "NUM DOC" in df.columns:
                v = row["NUM DOC"]
                num_doc = None if pd.isna(v) else int(v)
            fecha_rad      = row["FECHA RADICACION"]
            estado_factura = str(row.get("ESTADO_FACTURA","")).strip() or None
            imagen         = str(row.get("IMAGEN","")).strip() or None

            def s(col):
                v = row.get(col)
                return None if pd.isna(v) else str(v)
            rad_img   = s("RADICADO_IMAGEN")
            linea     = s("LINEA")
            id_asig   = s("ID ASIGNACION")
            est_pys   = s("ESTADO PYS")
            obs_pys   = s("OBSERVACION PYS")
            linea_pys = s("LINEA PYS")
            rangos    = s("RANGOS")
            def_col   = s("Def")

            params_list.append((
                radicado, nit, razon, factura, valor_factura,
                fecha_factura, fecha_rad, tipo_doc_id, num_doc,
                estado_factura, imagen,
                rad_img, linea, id_asig, est_pys,
                obs_pys, linea_pys, rangos, def_col,tipo_paquete,
                NUM_PAQUETE,
            ))
        except Exception:
            print(f"Error preparando fila {idx}:")
            traceback.print_exc()

        # Actualizar la barra de progreso
        progress_bar.set(idx + 1)

    # 10) Ejecutamos todos de golpe
    cur.executemany(
        """
        INSERT INTO ASIGNACION_TIPIFICACION
          (RADICADO, NIT, RAZON_SOCIAL, FACTURA, VALOR_FACTURA,
           FECHA_FACTURA, FECHA_RADICACION, TIPO_DOC_ID,
           NUM_DOC, ESTADO_FACTURA, IMAGEN, RADICADO_IMAGEN,
           LINEA, ID_ASIGNACION, ESTADO_PYS, OBSERVACION_PYS,
           LINEA_PYS, RANGOS, DEF, TIPO_PAQUETE, STATUS_ID, NUM_PAQUETE)
        VALUES
      (%s, %s, %s, %s, %s, %s, %s, %s,
       %s, %s, %s, %s, %s, %s, %s, %s,
       %s, %s, %s, %s, 1, %s)
        """,
        params_list
    )
    inserted = len(params_list)

    # 11) Desactivar IDENTITY_INSERT y commit
    cur.execute("SET IDENTITY_INSERT ASIGNACION_TIPIFICACION OFF;")
    conn.commit()
    cur.close()

    # Cerrar la ventana de progreso
    progress_window.destroy()

    messagebox.showinfo(
        "Carga completa",
        f"Total filas: {total}\nInsertadas: {inserted}\nPaquete: {NUM_PAQUETE}"
    )

    # 12) Selección de campos
    sel = ctk.CTkToplevel(root)
    sel.title(f"Paquete {NUM_PAQUETE}: Selecciona campos")
    campos = [
        "FECHA_SERVICIO", "FECHA_FINAL", "TIPO_DOC_ID", "NUM_DOC", "DIAGNOSTICO",
        "AUTORIZACION", "CODIGO_SERVICIO", "CANTIDAD", "VLR_UNITARIO",
        "COPAGO", "OBSERVACION"
    ]
    vars_chk = {}
    for campo in campos:
        vars_chk[campo] = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(sel, text=campo, variable=vars_chk[campo]).pack(
            anchor="w", padx=20, pady=2
        )

    def guardar_campos(paquete=NUM_PAQUETE):
        cur2 = conn.cursor()
        for campo, var in vars_chk.items():
            if var.get():
                cur2.execute(
                    "INSERT INTO PAQUETE_CAMPOS (NUM_PAQUETE, campo) VALUES (%s, %s)",
                    (paquete, campo)
                )
        conn.commit()
        cur2.close()
        sel.destroy()
        messagebox.showinfo("Guardado", f"Campos del paquete {NUM_PAQUETE} guardados.")

    ctk.CTkButton(sel, text="Guardar", command=guardar_campos).pack(pady=10)

def crear_usuario(root, conn):
    import re
    import bcrypt
    import tkinter as tk
    from tkinter import messagebox
    import customtkinter as ctk

    # Validaciones
    def only_upper_letters(input_str):
        # Acepta solo letras mayúsculas y la letra "ñ"
        if all(c.isalpha() or c == 'ñ' or c == 'Ñ' for c in input_str):
            return True
        else:
            return False

    def only_digits(P):
        return P == "" or P.isdigit()

    def validate_email(email):
        # Validación de correo utilizando una expresión regular simple
        regex = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
        return re.match(regex, email) is not None

    vcmd_letters = (root.register(only_upper_letters), '%P')
    vcmd_digits  = (root.register(only_digits), '%P')

    # Crear ventana secundaria
    top = ctk.CTkToplevel(root)
    top.title("Crear Usuario")
    top.geometry("500x600")
    top.resizable(False, False)

    # Recuperar catálogos
    cur = conn.cursor()
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
                validatecommand=cmd
            )
            # Convierte a mayúsculas al soltar cada tecla
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
            cursor = conn.cursor()
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

            conn.commit()
            cursor.close()
            messagebox.showinfo("Éxito", f"Usuario creado con ID {new_id}")
            top.destroy()

        except Exception as e:
            messagebox.showerror("Error", str(e))

    # Botón Guardar
    btn = ctk.CTkButton(frm, text="Guardar Usuario", command=guardar_usuario, width=200)
    btn.grid(row=len(labels)+1, column=0, columnspan=2, pady=20)
    top.bind("<Return>", guardar_usuario)


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

# -----------------------------
# Función iniciar_tipificacion
# -----------------------------
def iniciar_tipificacion(parent_root, conn, current_user_id):
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
    cur.close()

    
    

    # 2) Asignación aleatoria
    cur = conn.cursor()
    cur.execute("""
        SELECT TOP 1 RADICADO, NIT, FACTURA
        FROM ASIGNACION_TIPIFICACION
        WHERE STATUS_ID = 1
        AND NUM_PAQUETE = %s
        AND TIPO_PAQUETE = 'DIGITACION'
        ORDER BY NEWID()
    """, (pkg,))
    row = cur.fetchone()
    if not row:
        messagebox.showinfo("Sin asignaciones", "No hay asignaciones pendientes.")
        cur.close()
        return
    radicado, nit, factura = row
    entry_radicado_var.set(str(radicado))
    entry_nit_var.set(str(nit))
    entry_factura_var.set(str(factura))
    cur.execute("UPDATE ASIGNACION_TIPIFICACION SET STATUS_ID = 2 WHERE RADICADO = %s", (radicado,))
    conn.commit()
    cur.close()
    
    cur2 = conn.cursor()
    cur2.execute("SELECT campo FROM PAQUETE_CAMPOS WHERE NUM_PAQUETE = %s", (pkg,))
    campos_paquete = {r[0] for r in cur2.fetchall()}
    cur2.close()

    # 3) Ventana principal
    win = ctk.CTkToplevel(parent_root)
    win.title(f"Capturador De Datos · Paquete {pkg}")

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

    container = ctk.CTkFrame(win, fg_color="#1e1e1e")
    container.grid(row=0, column=0, sticky="nsew")
    win.grid_rowconfigure(0, weight=1)
    win.grid_columnconfigure(0, weight=1)

    card = ctk.CTkFrame(container, fg_color="#2b2b2b")
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
        text_color='white'
    ).pack(pady=(0, 15))


   # 4) Lectura de Radicado, NIT, Factura
    read_frame = ctk.CTkFrame(card, fg_color='transparent')
    read_frame.pack(fill='x', padx=30)
    read_frame.grid_columnconfigure(1, weight=1)

    # Labels fijos
    ctk.CTkLabel(read_frame, text="Radicado:", anchor='w').grid(row=0, column=0, pady=5, sticky='w')
    ctk.CTkEntry(
        read_frame,
        textvariable=entry_radicado_var,   # <-- aquí
        state='readonly',
        width=300
    ).grid(row=0, column=1, pady=5, sticky='ew', padx=(10,0))

    ctk.CTkLabel(read_frame, text="NIT:", anchor='w').grid(row=1, column=0, pady=5, sticky='w')
    ctk.CTkEntry(
        read_frame,
        textvariable=entry_nit_var,        # <-- y aquí
        state='readonly',
        width=300
    ).grid(row=1, column=1, pady=5, sticky='ew', padx=(10,0))

    ctk.CTkLabel(read_frame, text="Factura:", anchor='w').grid(row=2, column=0, pady=5, sticky='w')
    ctk.CTkEntry(
        read_frame,
        textvariable=entry_factura_var,    # <-- y aquí
        state='readonly',
        width=300
    ).grid(row=2, column=1, pady=5, sticky='ew', padx=(10,0))


    # 5) Scrollable y grid de 3 columnas
    scroll = ctk.CTkScrollableFrame(card, fg_color='#2b2b2b')
    scroll.pack(fill='both', expand=True, padx=20, pady=(10,0))
    card.pack_propagate(False) 
    card.grid_rowconfigure(1, weight=1)
    card.grid_columnconfigure(0, weight=1)
    for col in range(3):
        scroll.grid_columnconfigure(col, weight=1, uniform="col")

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
        cur = conn.cursor()
        cur.execute("""
            UPDATE ASIGNACION_TIPIFICACION 
            SET STATUS_ID = 1 
            WHERE RADICADO = %s
        """, (radicado,))
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
        frame = ctk.CTkFrame(scroll, fg_color='transparent')
        if icon_url:
            ico = load_icon_from_url(icon_url, size=(20,20))
            ctk.CTkLabel(frame, image=ico, text='').pack(side='left', padx=(0,5))
        ctk.CTkLabel(frame, text=label_text, anchor='w').pack(fill='x')
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
            validate='key',
            validatecommand=(win.register(lambda s: bool(re.match(r"^[0-9/]$", s))), '%S')
        )
        entry_fecha.pack(fill='x', pady=(5, 0))

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

        entry_tipo = AutocompleteEntry(frm, opts_td, width=300, textvariable=var_tipo)
        entry_tipo.pack(fill='x', pady=(5,0))

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
            validate='key', validatecommand=(win.register(lambda s: s.isdigit()), '%S')
        )
        entry_num.pack(fill='x', pady=(5,0))

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
    
    def add_service_block():
        nonlocal dynamic_row, dynamic_col
        dv = {}
        current_frames = []
        
        skip_obs = len(detail_vars) >= 1

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
                    placeholder_text='Solo 9 dígitos', validate='key',
                    validatecommand=vcmd_auth
                )
                w.pack(fill='x', pady=(5, 0))
                lbl_err.pack(fill='x', pady=(2, 8))

                def val_autorizacion(e=None, var=var, w=w, lbl=lbl_err):
                    txt = var.get().strip()
                    if not txt:
                        w.configure(border_color='red', border_width=2)
                        lbl.configure(text='Autorización obligatoria')
                        return False
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

                def to_upper_and_filter(*args):
                    txt = var.get()
                    filtered = ''.join(ch for ch in txt if ch.isalnum()).upper()
                    if txt != filtered:
                        var.set(filtered)
                var.trace_add('write', to_upper_and_filter)

                w = ctk.CTkEntry(
                    frm, textvariable=var,
                    placeholder_text='CÓDIGO DE SERVICIO', width=300,
                    validate='key', validatecommand=vcmd_cs
                )
                w.pack(fill='x', pady=(5, 0))
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
                    placeholder_text=placeholder,
                    validate='key', validatecommand=vcmd_num
                )
                w.pack(fill='x', pady=(5, 0))
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
                    placeholder_text=default
                )
                w.pack(fill='x', pady=(5, 0))

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

    # -----------------------------
    # Validar y guardar en BD
    # -----------------------------
    def validate_and_save(final):
    # ¿Alguna observación completada?
        any_obs = any(
            dv.get('OBSERVACION', {}).get('var').get().strip()
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

            # Detalle: si no hay observación en ese bloque, validar sus campos
            for dv in detail_vars:
                obs = dv.get('OBSERVACION', {}).get('var').get().strip()
                if obs:
                    # Si éste bloque tiene observación, salta validación de sus campos
                    continue

                # Validar AUTORIZACION si existe
                if 'VALIDAR_AUTORIZACION' in dv and callable(dv['VALIDAR_AUTORIZACION']):
                    if not dv['VALIDAR_AUTORIZACION']():
                        ok = False

                # Resto de campos del detalle: solo los que almacenan dict {'var','widget'}
                for campo, info in dv.items():
                    if not isinstance(info, dict):
                        continue  # saltar funciones u otras entradas
                    # ya omitimos VALIDAR_AUTORIZACION y OBSERVACION
                    if campo == 'OBSERVACION':
                        continue

                    w = info['widget']
                    if not info['var'].get().strip():
                        w.configure(border_color='red', border_width=2)
                        ok = False
                    else:
                        w.configure(border_color='#2b2b2b', border_width=1)

        return ok

    def load_assignment():
        """Carga aleatoriamente un radicado pendiente y actualiza los widgets."""
        nonlocal radicado, nit, factura

        # 1) Obtener nueva asignación
        cur = conn.cursor()
        cur.execute("""
            SELECT TOP 1 RADICADO, NIT, FACTURA
            FROM ASIGNACION_TIPIFICACION
            WHERE STATUS_ID = 1
            AND NUM_PAQUETE = %s
            AND TIPO_PAQUETE = 'DIGITACION'
            ORDER BY NEWID()
        """, (pkg,))
        row = cur.fetchone()
        if not row:
            messagebox.showinfo("Sin asignaciones", "No hay asignaciones pendientes.")
            cur.close()
            return False  # Indica que no hay más asignaciones
        radicado, nit, factura = row
        cur.execute(
            "UPDATE ASIGNACION_TIPIFICACION SET STATUS_ID = 2 WHERE RADICADO = %s",
            (radicado,)
        )
        conn.commit()
        cur.close()

        # 2) Actualizar campos de la GUI
        entry_radicado_var.set(str(radicado))
        entry_nit_var.set(str(nit))
        entry_factura_var.set(str(factura))

        # 3) Limpiar campos de tipificación previos
        for var in field_vars.values():
            var.set('')
        for dv in detail_vars:
            for info in dv.values():
                if isinstance(info, dict):
                    info['var'].set('')

        return True

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
            auth   = dv.get('AUTORIZACION', {}).get('var').get().strip() or None
            auth   = int(auth) if auth else None

            cs     = dv.get('CODIGO_SERVICIO', {}).get('var').get().strip().upper() or None
            qty    = dv.get('CANTIDAD', {}).get('var').get().strip() or None
            qty    = int(qty) if qty else None

            valor  = dv.get('VLR_UNITARIO', {}).get('var').get().strip() or None
            valor  = float(valor) if valor else None

            copago = dv.get('COPAGO', {}).get('var').get().strip() or None
            copago = float(copago) if copago else None

            obs    = dv.get('OBSERVACION', {}).get('var').get().strip() or None
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
            entry_fecha.unbind("<KeyRelease>")
            entry_fecha.unbind("<FocusOut>")

            # 2) Y programamos la destrucción en el idle loop,
            #    así los callbacks que ya estén en cola pueden finalizar sin error.
            win.after_idle(win.destroy)
            return
        if parent_root:
            parent_root.deiconify()
        else:
            # En lugar de win.destroy + reiniciar toda la función,
            # simplemente recargamos la siguiente asignación
            if not load_assignment():
                # Si no hay más asignaciones, cerramos
                win.destroy()
                if parent_root:
                    parent_root.deiconify()


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

    # Bind Enter para cada botón
    for b in (btn_save, btn_add, btn_del, btn_exit):
        b.bind("<Return>", lambda e, btn=b: btn.invoke())

def iniciar_calidad(parent_root, conn, current_user_id):
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
    cur.close()


    # 2) Asignación aleatoria
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
        messagebox.showinfo("Sin asignaciones", "No hay asignaciones pendientes.")
        cur.close()
        return
    radicado, nit, factura = row
    entry_radicado_var.set(str(radicado))
    entry_nit_var.set(str(nit))
    entry_factura_var.set(str(factura))
    cur.execute("UPDATE ASIGNACION_TIPIFICACION SET STATUS_ID = 2 WHERE RADICADO = %s", (radicado,))
    conn.commit()
    cur.close()
    
    cur2 = conn.cursor()
    cur2.execute("SELECT campo FROM PAQUETE_CAMPOS WHERE NUM_PAQUETE = %s", (pkg,))
    campos_paquete = {r[0] for r in cur2.fetchall()}
    cur2.close()

    # 3) Ventana principal
    win = ctk.CTkToplevel(parent_root)
    win.title(f"Capturador De Datos · Paquete {pkg}")

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

    container = ctk.CTkFrame(win, fg_color="#1e1e1e")
    container.grid(row=0, column=0, sticky="nsew")
    win.grid_rowconfigure(0, weight=1)
    win.grid_columnconfigure(0, weight=1)

    card = ctk.CTkFrame(container, fg_color="#2b2b2b")
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
        text_color='white'
    ).pack(pady=(0, 15))


   # 4) Lectura de Radicado, NIT, Factura
    read_frame = ctk.CTkFrame(card, fg_color='transparent')
    read_frame.pack(fill='x', padx=30)
    read_frame.grid_columnconfigure(1, weight=1)

    # Labels fijos
    ctk.CTkLabel(read_frame, text="Radicado:", anchor='w').grid(row=0, column=0, pady=5, sticky='w')
    ctk.CTkEntry(
        read_frame,
        textvariable=entry_radicado_var,   # <-- aquí
        state='readonly',
        width=300
    ).grid(row=0, column=1, pady=5, sticky='ew', padx=(10,0))

    ctk.CTkLabel(read_frame, text="NIT:", anchor='w').grid(row=1, column=0, pady=5, sticky='w')
    ctk.CTkEntry(
        read_frame,
        textvariable=entry_nit_var,        # <-- y aquí
        state='readonly',
        width=300
    ).grid(row=1, column=1, pady=5, sticky='ew', padx=(10,0))

    ctk.CTkLabel(read_frame, text="Factura:", anchor='w').grid(row=2, column=0, pady=5, sticky='w')
    ctk.CTkEntry(
        read_frame,
        textvariable=entry_factura_var,    # <-- y aquí
        state='readonly',
        width=300
    ).grid(row=2, column=1, pady=5, sticky='ew', padx=(10,0))


    # 5) Scrollable y grid de 3 columnas
    scroll = ctk.CTkScrollableFrame(card, fg_color='#2b2b2b')
    scroll.pack(fill='both', expand=True, padx=20, pady=(10,0))
    card.pack_propagate(False) 
    card.grid_rowconfigure(1, weight=1)
    card.grid_columnconfigure(0, weight=1)
    for col in range(3):
        scroll.grid_columnconfigure(col, weight=1, uniform="col")

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
        cur = conn.cursor()
        cur.execute("""
            UPDATE ASIGNACION_TIPIFICACION 
            SET STATUS_ID = 1 
            WHERE RADICADO = %s
        """, (radicado,))
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
        frame = ctk.CTkFrame(scroll, fg_color='transparent')
        if icon_url:
            ico = load_icon_from_url(icon_url, size=(20,20))
            ctk.CTkLabel(frame, image=ico, text='').pack(side='left', padx=(0,5))
        ctk.CTkLabel(frame, text=label_text, anchor='w').pack(fill='x')
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
            validate='key',
            validatecommand=(win.register(lambda s: bool(re.match(r"^[0-9/]$", s))), '%S')
        )
        entry_fecha.pack(fill='x', pady=(5, 0))

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
            validate='key',
            validatecommand=(win.register(lambda s: bool(re.match(r"^[0-9/]$", s))), '%S')
        )
        entry_fecha_final.pack(fill='x', pady=(5, 0))

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

        # Forzar mayúsculas en KeyRelease
        def to_upper_on_key(event, var=var_tipo):
            var.set(var.get().upper())
        entry_tipo.bind('<KeyRelease>', to_upper_on_key)

        # Etiqueta de error
        lbl_err_td = ctk.CTkLabel(frm, text='', text_color='red')
        lbl_err_td.pack(fill='x', pady=(2,0))

        # Validación al perder foco
        def val_tipo(e=None):
            if not var_tipo.get().strip():
                entry_tipo.configure(border_color='red', border_width=2)
                lbl_err_td.configure(text='Tipo de documento obligatorio')
                return False
            else:
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
            validate='key', validatecommand=(win.register(lambda s: s.isdigit()), '%S')
        )
        entry_num.pack(fill='x', pady=(5,0))

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
            on_select()
            if not var_diag.get().strip():
                entry_diag.configure(border_color='red', border_width=2)
                lbl_err_diag.configure(text='Diagnóstico obligatorio')
                return False
            else:
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
    
    def add_service_block():
        nonlocal dynamic_row, dynamic_col
        dv = {}
        current_frames = []
        
        skip_obs = len(detail_vars) >= 1

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
                    placeholder_text='Solo 9 dígitos', validate='key',
                    validatecommand=vcmd_auth
                )
                w.pack(fill='x', pady=(5, 0))
                lbl_err.pack(fill='x', pady=(2, 8))

                def val_autorizacion(e=None, var=var, w=w, lbl=lbl_err):
                    txt = var.get().strip()
                    if not txt:
                        w.configure(border_color='red', border_width=2)
                        lbl.configure(text='Autorización obligatoria')
                        return False
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

                def to_upper_and_filter(*args):
                    txt = var.get()
                    filtered = ''.join(ch for ch in txt if ch.isalnum()).upper()
                    if txt != filtered:
                        var.set(filtered)
                var.trace_add('write', to_upper_and_filter)

                w = ctk.CTkEntry(
                    frm, textvariable=var,
                    placeholder_text='CÓDIGO DE SERVICIO', width=300,
                    validate='key', validatecommand=vcmd_cs
                )
                w.pack(fill='x', pady=(5, 0))
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
                    placeholder_text=placeholder,
                    validate='key', validatecommand=vcmd_num
                )
                w.pack(fill='x', pady=(5, 0))
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
                    placeholder_text=default
                )
                w.pack(fill='x', pady=(5, 0))

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

    # -----------------------------
    # Validar y guardar en BD
    # -----------------------------
    def validate_and_save(final):
    # ¿Alguna observación completada?
        any_obs = any(
            dv.get('OBSERVACION', {}).get('var').get().strip()
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

            # Detalle: si no hay observación en ese bloque, validar sus campos
            for dv in detail_vars:
                obs = dv.get('OBSERVACION', {}).get('var').get().strip()
                if obs:
                    # Si éste bloque tiene observación, salta validación de sus campos
                    continue

                # Validar AUTORIZACION si existe
                if 'VALIDAR_AUTORIZACION' in dv and callable(dv['VALIDAR_AUTORIZACION']):
                    if not dv['VALIDAR_AUTORIZACION']():
                        ok = False

                # Resto de campos del detalle: solo los que almacenan dict {'var','widget'}
                for campo, info in dv.items():
                    if not isinstance(info, dict):
                        continue  # saltar funciones u otras entradas
                    # ya omitimos VALIDAR_AUTORIZACION y OBSERVACION
                    if campo == 'OBSERVACION':
                        continue

                    w = info['widget']
                    if not info['var'].get().strip():
                        w.configure(border_color='red', border_width=2)
                        ok = False
                    else:
                        w.configure(border_color='#2b2b2b', border_width=1)

        return ok

    def load_assignment():
        """Carga aleatoriamente un radicado pendiente y actualiza los widgets."""
        nonlocal radicado, nit, factura

        # 1) Obtener nueva asignación
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
            messagebox.showinfo("Sin asignaciones", "No hay asignaciones pendientes.")
            cur.close()
            return False  # Indica que no hay más asignaciones
        radicado, nit, factura = row
        cur.execute(
            "UPDATE ASIGNACION_TIPIFICACION SET STATUS_ID = 2 WHERE RADICADO = %s",
            (radicado,)
        )
        conn.commit()
        cur.close()

        # 2) Actualizar campos de la GUI
        entry_radicado_var.set(str(radicado))
        entry_nit_var.set(str(nit))
        entry_factura_var.set(str(factura))

        # 3) Limpiar campos de tipificación previos
        for var in field_vars.values():
            var.set('')
        for dv in detail_vars:
            for info in dv.values():
                if isinstance(info, dict):
                    info['var'].set('')

        return True

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
            auth   = dv.get('AUTORIZACION', {}).get('var').get().strip() or None
            auth   = int(auth) if auth else None

            cs     = dv.get('CODIGO_SERVICIO', {}).get('var').get().strip().upper() or None
            qty    = dv.get('CANTIDAD', {}).get('var').get().strip() or None
            qty    = int(qty) if qty else None

            valor  = dv.get('VLR_UNITARIO', {}).get('var').get().strip() or None
            valor  = float(valor) if valor else None

            copago = dv.get('COPAGO', {}).get('var').get().strip() or None
            copago = float(copago) if copago else None

            obs    = dv.get('OBSERVACION', {}).get('var').get().strip() or None
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
            entry_fecha.unbind("<KeyRelease>")
            entry_fecha.unbind("<FocusOut>")

            # 2) Y programamos la destrucción en el idle loop,
            #    así los callbacks que ya estén en cola pueden finalizar sin error.
            win.after_idle(win.destroy)
            return
        if parent_root:
            parent_root.deiconify()
        else:
            # En lugar de win.destroy + reiniciar toda la función,
            # simplemente recargamos la siguiente asignación
            if not load_assignment():
                # Si no hay más asignaciones, cerramos
                win.destroy()
                if parent_root:
                    parent_root.deiconify()


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
            messagebox.showwarning("Selección inválida", "Selecciona un paquete válido.")
            return None, None

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

        where_clause = " AND ".join(filtros)
        return where_clause, tuple(params)

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

    # — Carga datos en las dos pestañas según filtros —
    def actualizar_tabs():
        where, params = construir_filtros()
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
        for i, (est, cnt) in enumerate(cur.fetchall()):
            ctk.CTkLabel(frame1, text=est).grid(row=i, column=0, sticky="w", padx=5, pady=2)
            ctk.CTkLabel(frame1, text=cnt).grid(row=i, column=1, sticky="e", padx=5, pady=2)
        cur.close()

        # — Pestaña "Por Usuario" —
        frame2 = tabs.tab("Por Usuario")
        for w in frame2.winfo_children():
            w.destroy()
        cur = conn.cursor()
        sql2 = (
            "SELECT u.ID, u.FIRST_NAME + ' ' + u.LAST_NAME AS USUARIO, "
            "SUM(CASE WHEN a.STATUS_ID=2 THEN 1 ELSE 0 END) AS PENDIENTES, "
            "SUM(CASE WHEN a.STATUS_ID=3 THEN 1 ELSE 0 END) AS PROCESADOS, "
            "SUM(CASE WHEN a.STATUS_ID=4 THEN 1 ELSE 0 END) AS CON_OBS "
            "FROM ASIGNACION_TIPIFICACION a "
            "JOIN TIPIFICACION t ON t.ASIGNACION_ID = a.RADICADO "
            "JOIN USERS u ON t.USER_ID = u.ID "
            "JOIN STATUS s ON a.STATUS_ID = s.ID "
            f"WHERE {where} GROUP BY u.ID, u.FIRST_NAME, u.LAST_NAME ORDER BY USUARIO"
        )
        cur.execute(sql2, params)
        headers = ["ID", "USUARIO", "PENDIENTES", "PROCESADOS", "CON_OBS"]
        for j, h in enumerate(headers):
            ctk.CTkLabel(frame2, text=h).grid(row=0, column=j, padx=5, pady=4, sticky="w")
        for i, row in enumerate(cur.fetchall(), start=1):
            for j, val in enumerate(row):
                ctk.CTkLabel(frame2, text=val).grid(row=i, column=j, padx=5, pady=2, sticky="w")
        cur.close()
        
    def exportar_excel(path, headers, rows):
        import pandas as pd
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
            })
            data_fmt = workbook.add_format({
                'border': 1,             # borde fino
                'align': 'center',
                'valign': 'vcenter',
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
        where, params = construir_filtros()
        if where is None:
            return

        # — 1) SQL que incluye solo las columnas requeridas —
        sql_export = (
            "SELECT "
            "  a.RADICADO, "
            "  t.FECHA_SERVICIO, "
            "  d.AUTORIZACION, "
            "  d.CODIGO_SERVICIO, "
            "  d.CANTIDAD, "
            "  d.VLR_UNITARIO, "
            "  t.DIAGNOSTICO, "
            "  t.fecha_creacion    AS CreatedOn, "
            " CONCAT(u.FIRST_NAME, ' ', u.LAST_NAME, ' - ', CAST(u.NUM_DOC AS varchar(20))) AS ModifiedBy, "
            "  td.NAME           AS TipoDocumento, "
            "  t.NUM_DOC           AS NumeroDocumento, "
            "  d.COPAGO            AS CM_COPAGO, "
            "  d.OBSERVACION, "
            "  s.NAME              AS ESTADO "
            "FROM ASIGNACION_TIPIFICACION a "
            "JOIN TIPIFICACION t          ON t.ASIGNACION_ID     = a.RADICADO "
            "JOIN TIPO_DOC td          ON t.TIPO_DOC_ID     = td.ID "
            "JOIN USERS u          ON t.USER_ID     = u.ID "
            "JOIN TIPIFICACION_DETALLES d ON d.TIPIFICACION_ID   = t.ID "
            "JOIN STATUS s                ON s.ID                = a.STATUS_ID "
            f"WHERE {where}"
        )

        # — 2) Diálogo para elegir ruta y extensión —
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

        # — 3) Ejecutar consulta una sola vez —
        cur = conn.cursor()
        cur.execute(sql_export, params)
        rows = cur.fetchall()
        headers = [col[0] for col in cur.description]
        cur.close()

        # — 4) Rutas de exportación —
        if ext in ("csv", "txt"):
            sep = ',' if ext=="csv" else '\t'
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
        styles = getSampleStyleSheet()
        data = [headers] + [list(r) for r in rows]

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
    win.geometry("1150x700")
    win.grab_set()
    win.protocol("WM_DELETE_WINDOW", lambda: safe_destroy(win))

    topfrm = ctk.CTkFrame(win, fg_color="transparent")
    topfrm.pack(fill="x", padx=20, pady=(20, 0))

    # Paquete
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT NUM_PAQUETE FROM ASIGNACION_TIPIFICACION ORDER BY NUM_PAQUETE")
    paquetes = [str(r[0]) for r in cur.fetchall()] or ["0"]
    cur.close()
    pkg_var = tk.StringVar(value=paquetes[0])
    ctk.CTkLabel(topfrm, text="Paquete:").grid(row=0, column=0, sticky="w")
    ctk.CTkOptionMenu(topfrm, values=paquetes, variable=pkg_var, width=80).grid(row=0, column=1, padx=(0,20), sticky="w")

    # Tipo de paquete
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT TIPO_PAQUETE FROM ASIGNACION_TIPIFICACION ORDER BY TIPO_PAQUETE")
    tipos_paquete = [r[0] or "" for r in cur.fetchall()]
    cur.close()
    var_tipo_paquete = tk.StringVar(value=tipos_paquete[0])
    ctk.CTkLabel(topfrm, text="Tipo Paquete:").grid(row=0, column=2, sticky="w")
    ctk.CTkOptionMenu(topfrm, values=tipos_paquete, variable=var_tipo_paquete, width=80).grid(row=0, column=3, padx=(0,20), sticky="w")

    # Fechas
    var_fecha_desde = tk.StringVar()
    var_fecha_hasta = tk.StringVar()
    ctk.CTkLabel(topfrm, text="Desde:").grid(row=0, column=4, sticky="w")
    fecha_desde = DateEntry(topfrm, width=12, locale='es_CO', date_pattern='dd/MM/yyyy', textvariable=var_fecha_desde)
    fecha_desde.grid(row=0, column=5, padx=(0,20), sticky="w")
    fecha_desde.delete(0, 'end')
    ctk.CTkLabel(topfrm, text="Hasta:").grid(row=0, column=6, sticky="w")
    fecha_hasta = DateEntry(topfrm, width=12, locale='es_CO', date_pattern='dd/MM/yyyy', textvariable=var_fecha_hasta)
    fecha_hasta.grid(row=0, column=7, padx=(0,20), sticky="w")
    fecha_hasta.delete(0, 'end')

    ctk.CTkButton(topfrm, text="Limpiar fechas", command=lambda: (var_fecha_desde.set(""), var_fecha_hasta.set("")), width=100).grid(row=0, column=8, padx=(0,20))
    ctk.CTkButton(topfrm, text="Aplicar filtros", command=actualizar_tabs, width=120).grid(row=0, column=9, padx=(0,20))
    ctk.CTkButton(topfrm,
        text="Exportar",
        command=exportar,
        width=100
    ).grid(row=0, column=10)
    # Filtro de estados
    cur = conn.cursor()
    cur.execute("SELECT NAME FROM STATUS ORDER BY NAME")
    estados = [r[0] for r in cur.fetchall()]
    cur.close()
    ctk.CTkLabel(topfrm, text="Estado:").grid(row=1, column=0, sticky="w", pady=(20,5))
    buscar_est = ctk.CTkEntry(topfrm, width=200, placeholder_text="Buscar estado...")
    buscar_est.grid(row=1, column=1, columnspan=3, sticky="w", padx=(0,20), pady=(20,5))
    buscar_est.bind("<KeyRelease>", _filtrar_est)
    ctk.CTkButton(topfrm, text="Todo", command=lambda: _marcar_est(True), width=60).grid(row=1, column=4, pady=(20,5), sticky="w")
    ctk.CTkButton(topfrm, text="Ninguno", command=lambda: _marcar_est(False), width=60).grid(row=1, column=5, pady=(20,5), sticky="w")
    estado_frame = ctk.CTkScrollableFrame(topfrm, width=250, height=120)
    estado_frame.grid(row=2, column=0, columnspan=6, sticky="w")
    estado_vars = {}
    estado_checks = {}
    for est in estados:
        var = tk.BooleanVar(value=True)
        cb = ctk.CTkCheckBox(estado_frame, text=est, variable=var)
        cb.pack(anchor="w", pady=2)
        estado_vars[est] = var
        estado_checks[est] = cb

    # Filtro de usuarios
    cur = conn.cursor()
    cur.execute("SELECT FIRST_NAME + ' ' + LAST_NAME FROM USERS ORDER BY FIRST_NAME")
    usuarios = [r[0] for r in cur.fetchall()]
    cur.close()
    ctk.CTkLabel(topfrm, text="Usuario:").grid(row=1, column=6, sticky="w", pady=(20,5), padx=(20,0))
    buscar_usr = ctk.CTkEntry(topfrm, width=200, placeholder_text="Buscar usuario...")
    buscar_usr.grid(row=1, column=7, columnspan=2, sticky="w", pady=(20,5))
    buscar_usr.bind("<KeyRelease>", _filtrar_usr)
    ctk.CTkButton(topfrm, text="Todo", command=lambda: _marcar_usr(True), width=60).grid(row=1, column=9, pady=(20,5), sticky="w", padx=(5,0))
    ctk.CTkButton(topfrm, text="Ninguno", command=lambda: _marcar_usr(False), width=60).grid(row=1, column=10, pady=(20,5), sticky="w", padx=(5,0))
    user_frame = ctk.CTkScrollableFrame(topfrm, width=250, height=120)
    user_frame.grid(row=2, column=6, columnspan=5, sticky="w")
    user_vars = {}
    user_checks = {}
    for usr in usuarios:
        var = tk.BooleanVar(value=True)
        cb = ctk.CTkCheckBox(user_frame, text=usr, variable=var)
        cb.pack(anchor="w", pady=2)
        user_vars[usr] = var
        user_checks[usr] = cb

    # Pestañas de resultados
    tabs = ctk.CTkTabview(win, width=760, height=440)
    tabs.pack(padx=20, pady=(10,20), fill="both", expand=True)
    tabs.add("Por Estado")
    tabs.add("Por Usuario")
    win._tabview = tabs

    # Carga inicial al abrir ventana
    actualizar_tabs()


def actualizar_tabs(win, conn, num_paquete):
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

    # Encabezados
    ctk.CTkLabel(frame1, text="ESTADO", anchor="w").grid(row=0, column=0, padx=5, pady=4)
    ctk.CTkLabel(frame1, text="CANTIDAD", anchor="e").grid(row=0, column=1, padx=5, pady=4)

    # Filas
    for i, (estado, cnt) in enumerate(datos, start=1):
        ctk.CTkLabel(frame1, text=estado, anchor="w")\
            .grid(row=i, column=0, sticky="w", padx=5, pady=2)
        ctk.CTkLabel(frame1, text=str(cnt), anchor="e")\
            .grid(row=i, column=1, sticky="e", padx=5, pady=2)


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
    win.geometry("500x350")
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
    import csv
    from tkinter import filedialog, messagebox, Text, Scrollbar

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
    ctk.CTkOptionMenu(frm, values=[str(p) for p in paquetes], variable=pkg_var, width=120)\
        .grid(row=0, column=1, pady=(0,5), sticky="w")

    # 4) Selector de formato
    ctk.CTkLabel(frm, text="Formato:").grid(row=1, column=0, sticky="w", pady=(0,10))
    fmt_var = tk.StringVar(value="CSV")
    ctk.CTkOptionMenu(frm, values=["CSV", "TXT"], variable=fmt_var, width=120)\
        .grid(row=1, column=1, pady=(0,10), sticky="w")

    # 5) Área de texto para radicados
    ctk.CTkLabel(frm, text="Radicados (uno por línea, opcional):").grid(row=2, column=0, columnspan=2, sticky="w")
    text_frame = tk.Frame(frm)
    text_frame.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=(0,10))
    # Ajustar expansión
    frm.grid_rowconfigure(3, weight=1)
    frm.grid_columnconfigure(1, weight=1)

    scrollbar = Scrollbar(text_frame)
    scrollbar.pack(side="right", fill="y")
    txt = Text(text_frame, height=8, yscrollcommand=scrollbar.set)
    txt.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=txt.yview)

    # 6) Botón Exportar
    def _export():
        pkg = int(pkg_var.get())
        fmt = fmt_var.get()
        ext = ".txt" if fmt == "TXT" else ".csv"
        sep = ";" if fmt == "TXT" else ","
        path = filedialog.asksaveasfilename(
            filetypes=[("CSV Files", "*.csv")],
            defaultextension=".csv",
            initialfile="reporte.csv",   
            title="Guardar como CSV" 
        )
        if not path:
            return

        # 7) Leo radicados del textarea
        lines = [line.strip() for line in txt.get("1.0", "end").splitlines()]
        radicados = []
        for L in lines:
            if L.isdigit():
                radicados.append(int(L))
        radicados = sorted(set(radicados))

        # 8) Construyo la SQL base
        base_sql = """
            SELECT
              a.RADICADO                                   AS RADICADO,
              CONVERT(varchar(10), t.FECHA_SERVICIO, 103)  AS FECHA_SERVICIO,
              d.AUTORIZACION                               AS AUTORIZACION,
              d.CODIGO_SERVICIO                            AS COD_SERVICIO,
              CONVERT(int, d.CANTIDAD)                     AS CANTIDAD,
              CONVERT(int, d.VLR_UNITARIO)                 AS VLR_UNITARIO,
              t.DIAGNOSTICO                                AS DIAGNOSTICO,
               CONVERT(varchar(10), t.CreatedOn, 103)      AS CreatedOn,
              u2.NUM_DOC                                   AS ModifiedBy,
              td.NAME                                      AS TipoDocumento,
              a.NUM_DOC                                    AS NumeroDocumento,
              CONVERT(int, d.COPAGO)                       AS CM_COPAGO
            FROM ASIGNACION_TIPIFICACION a
            JOIN TIPIFICACION t  ON t.ASIGNACION_ID = a.RADICADO
            JOIN TIPIFICACION_DETALLES d ON d.TIPIFICACION_ID = t.ID
            JOIN USERS u2       ON u2.ID = t.USER_ID
            JOIN TIPO_DOC td    ON td.ID = t.TIPO_DOC_ID
            WHERE a.NUM_PAQUETE = %s
        """

        params = [pkg]

        # 9) Si hay radicados en textarea, agrego filtro
        if radicados:
            placeholders = ",".join("%s" for _ in radicados)
            base_sql += f" AND a.RADICADO IN ({placeholders})"
            params.extend(radicados)

        base_sql += " ORDER BY a.RADICADO, t.FECHA_SERVICIO"

        # 10) Ejecuto y exporto
        cur2 = conn.cursor()
        cur2.execute(base_sql, params)
        rows = cur2.fetchall()
        headers = [col[0] for col in cur2.description]
        cur2.close()

        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, delimiter=sep)
            writer.writerow(headers)
            writer.writerows(rows)

        messagebox.showinfo("Exportar",
            f"Paquete {pkg} exportado en:\n{path}\n"
            f"{len(rows)} registros."
        )
        win.destroy()

    ctk.CTkButton(frm, text="Exportar", command=_export, width=200)\
        .grid(row=4, column=0, columnspan=2, pady=(10,0))

    win.mainloop()

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
            show="*" if is_password else None
        )
        entry.grid(row=i, column=1, pady=5)
        if i == 0:
            entry.focus()

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

def open_dashboard(user_id, first_name, last_name, parent):
    conn = conectar_sql_server('DB_DATABASE')
    if conn is None:
        messagebox.showerror("Error", "No se pudo conectar a la base de datos.")
        return

    # Obtengo roles del usuario
    cursor = conn.cursor()
    cursor.execute(
        "SELECT R.ID, R.NAME "
        "FROM USER_ROLES UR "
        "JOIN ROL R ON UR.ROL_ID = R.ID "
        "WHERE UR.USER_ID = %s",
        (user_id,)
    )
    roles = cursor.fetchall()  # lista de (id, name)
    cursor.close()

    # Mapeo nombre->id para usar en el OptionMenu
    role_map = {name: rid for (rid, name) in roles}
    role_names = list(role_map.keys())

    # Creo el Toplevel
    root = ctk.CTkToplevel(parent)
    root.title("Dashboard - Capturador De Datos")
    root.geometry("500x500")
    root.resizable(False, False)

    # Bienvenida
    ctk.CTkLabel(
        root,
        text=f"Bienvenido, {first_name} {last_name}",
        font=ctk.CTkFont(size=20, weight="bold")
    ).pack(pady=(20, 10))

    # Selector de rol
    role_var = tk.StringVar(value=role_names[0] if role_names else "")
    option = ctk.CTkOptionMenu(
        root,
        values=role_names,
        variable=role_var
    )
    option.pack(pady=(0, 10))

    # Marco para los botones
    btn_frame = ctk.CTkFrame(root)
    btn_frame.pack(padx=20, pady=20, fill="both", expand=True)

    # Función para cerrar sesión
    def on_logout():
        safe_destroy(root)
        login_script = os.path.join(os.path.dirname(__file__), "login_app.py")
        subprocess.Popen(
            [sys.executable, login_script],
            cwd=os.path.dirname(__file__)
        )

    # Botones por rol
    def start_tipificacion_and_close():
        safe_destroy(root)
        iniciar_tipificacion(None, conn, user_id)

    def start_quality_and_close():
        safe_destroy(root)
        iniciar_calidad(None, conn, user_id)
    
    buttons_by_role = {
        1: [
            ("Cargar Paquete",     lambda: cargar_paquete(root, conn)),
            ("Crear Usuario",      lambda: crear_usuario(root, conn)),
            ("Actualizar Datos",   lambda: actualizar_usuario(root, conn, user_id)),
            ("Ver Progreso",       lambda: ver_progreso(root, conn)),
            ("Activar/Desactivar Usuario", lambda: modificar_estado_usuario(root, conn)),
            ("Exportar Capturación De Datos", lambda: exportar_paquete(root, conn)),
        ],
        2: [
            ("Iniciar Capturación De Datos", start_tipificacion_and_close),
            ("Actualizar Datos",   lambda: actualizar_usuario(root, conn, user_id)),
            ("Ver Progreso",       lambda: ver_progreso(root, conn)),
            ("Exportar Capturación De Datos", lambda: exportar_paquete(root, conn)),
        ],
        3: [
            ("Iniciar Validación Calidad Datos", start_quality_and_close)
        ]
    }

    def show_role_buttons(selected):
        for w in btn_frame.winfo_children():
            w.destroy()
        rid = role_map[selected]
        for text, cmd in buttons_by_role.get(rid, []):
            ctk.CTkButton(btn_frame, text=text, command=cmd, width=200)\
                .pack(pady=5, anchor="center")

    option.configure(command=show_role_buttons)
    if role_names:
        show_role_buttons(role_var.get())

    def on_logout():
        # destruye el dashboard actual
        safe_destroy(root)
        # vuelve a lanzar el script de login en un proceso separado
        login_script = os.path.join(os.path.dirname(__file__), "login_app.py")
        subprocess.Popen(
            [sys.executable, login_script],
            cwd=os.path.dirname(__file__)
        )


    ctk.CTkButton(
        root,
        text="Logout",
        command=on_logout,
        width=120
    ).place(relx=0.5, rely=0.9, anchor="center")

if __name__ == "__main__":
    import sys
    import customtkinter as ctk

    if len(sys.argv) != 4:
        print("Uso: python dashboard.py <user_id> <first_name> <last_name>")
        sys.exit(1)

    _, uid, fn, ln = sys.argv

    root = ctk.CTk()
    root.withdraw()   # ocultas la raíz

    # ahora open_dashboard devuelve la ventana hija con su propio protocol
    dashboard = open_dashboard(int(uid), fn, ln, parent=root)

    root.mainloop()
