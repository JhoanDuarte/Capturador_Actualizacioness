from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import os
import sys
import subprocess
from threading import Thread
import time
from tkinter import messagebox
import customtkinter as ctk
import bcrypt
import smtplib
import random
import string
import requests
import tkinter as tk
from tkinter import ttk
from tqdm import tqdm
from PyQt5.QtWidgets import QMessageBox,QGraphicsDropShadowEffect,QGraphicsBlurEffect
from PyQt5.QtCore import QRegularExpression
from PyQt5.QtGui  import QRegularExpressionValidator
from PyQt5.QtGui import QIntValidator


# — Librerías estándar —
# (subprocess ya importado arriba si lo necesitas para llamadas externas)

# — Librerías de PyQt5 —
from PyQt5.QtGui import QIntValidator
from PyQt5 import QtCore, QtGui, QtWidgets

# — Módulos propios —
from db_connection import conectar_sql_server
from dashboard import DashboardWindow

from version import __version__ as local_version

APP_NAME = "Dashboard_Capturador_Datos"
UPDATE_JSON_URL = "https://raw.githubusercontent.com/JhoanDuarte/Capturador_Actualizacioness/main/latest.json"

try:
    from version import __version__ as local_version
except ImportError:
    local_version = "1.3.3"  # Si no hay versión, se forzará la actualización

import os
import sys
import requests
import tkinter as tk
from tkinter import messagebox

def get_target_zip_path(version):
    # Misma lógica de antes
    app_dir = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else __file__)
    target_dir = os.path.abspath(os.path.join(app_dir, "..", ".."))
    zip_name = f"{APP_NAME}_{version}.zip"
    return os.path.join(target_dir, zip_name)

def show_update_required_window(zip_path, version):
    # Crea una ventana modal que no deja avanzar al login
    window = tk.Tk()
    window.title("Actualización requerida")
    window.geometry("520x180")
    window.resizable(False, False)

    zip_name = os.path.basename(zip_path)
    msg = (
        f"Se ha descargado la actualización requerida.\n\n"
        f"Versión: {version}\n"
        f"Archivo: {zip_name}\n"
        f"Ubicación: {os.path.dirname(zip_path)}"
    )
    tk.Label(window, text=msg, font=("Arial", 10), justify="left", wraplength=500).pack(pady=10)

    def abrir_directorio():
        # Abre el explorador y luego cierra la app
        os.startfile(os.path.dirname(zip_path))
        window.destroy()

    tk.Button(
        window,
        text="Abrir carpeta de actualización",
        command=abrir_directorio
    ).pack(pady=10)

    # Si cierra con la “X”, también sale toda la app
    window.protocol("WM_DELETE_WINDOW", window.destroy)
    window.mainloop()

    # Al destruir la ventana, matamos el proceso para que no siga al login
    os._exit(0)

def check_for_update_and_exit_if_needed():
    try:
        # Descargamos el JSON de versiones
        r = requests.get(UPDATE_JSON_URL, timeout=10)
        r.raise_for_status()
        data = r.json()
        remote_version = data.get("version")
        zip_url = data.get("url")

        # Si cambió la versión, forzamos descarga y bloqueo
        if local_version != remote_version:
            zip_path = get_target_zip_path(remote_version)

            # Sólo descargar si no existe ya
            if not os.path.exists(zip_path):
                r2 = requests.get(zip_url, timeout=30)
                r2.raise_for_status()
                os.makedirs(os.path.dirname(zip_path), exist_ok=True)
                with open(zip_path, "wb") as f:
                    f.write(r2.content)

            # Muestra ventana y, al cerrarla o pulsar el botón, sale todo
            show_update_required_window(zip_path, remote_version)

        # Si la versión es la misma, simplemente retorna y deja continuar al login
    except Exception as e:
        # En caso de error al verificar actualización, avisa y no deja avanzar
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Error", f"No se pudo verificar la actualización:\n{e}")
        os._exit(0)

def run_dashboard_from_args():
    # Si estamos en el exe congelado y el primer arg es dashboard.py
    if getattr(sys, 'frozen', False) \
       and len(sys.argv) >= 2 \
       and os.path.basename(sys.argv[1]).lower() == "dashboard.py":
        # argv[2], [3], [4] son user_id, first_name, last_name
        try:
            uid   = int(sys.argv[2])
            fn    = sys.argv[3]
            ln    = sys.argv[4]
        except Exception:
            print("Uso incorrecto: login_app.exe dashboard.py <id> <nombre> <apellido>")
            sys.exit(1)
        # Conectamos
        conn = conectar_sql_server('DB_DATABASE')
        if conn is None:
            raise RuntimeError("No se pudo conectar a la BD.")
        # Creamos root y abrimos el dashboard
        def run_dashboard_from_args():
            if getattr(sys, 'frozen', False) \
            and len(sys.argv) >= 2 and os.path.basename(sys.argv[1]).lower() == "dashboard.py":
                try:
                    uid = int(sys.argv[2])
                    fn  = sys.argv[3]
                    ln  = sys.argv[4]
                except Exception:
                    print("Uso incorrecto: login_app.exe dashboard.py <id> <nombre> <apellido>")
                    sys.exit(1)

                # Arrancamos Qt en vez de CTk
                app = QtWidgets.QApplication(sys.argv)
                window = DashboardWindow(uid, fn, ln)   # sin parent=…
                window.show()
                sys.exit(app.exec_())


# Ejecutamos la detección *antes* de definir nada más
run_dashboard_from_args()


# Conexión a BD
conn = conectar_sql_server('DB_DATABASE')
if conn is None:
    raise RuntimeError("No se pudo conectar a la base de datos.")

def resource_path(rel):
    # si estamos “frozen” (onefile), todo está en sys._MEIPASS
    base = getattr(sys, '_MEIPASS', os.path.dirname(__file__))
    return os.path.join(base, rel)

def authenticate_user_by_doc(num_doc: str, password: str):
    cursor = conn.cursor()
    cursor.execute(
        "SELECT ID, FIRST_NAME, LAST_NAME, PASSWORD, STATUS_ID FROM USERS WHERE NUM_DOC = %s",
        (num_doc,),
    )
    row = cursor.fetchone()
    cursor.close()
    if not row:
        return None

    user_id, first_name, last_name, stored_hash, status_id = row

    # bcrypt espera bytes
    pwd_bytes       = password.strip().encode("utf-8")
    stored_hash_b   = stored_hash.encode("utf-8")

    # Verificar con bcrypt
    if bcrypt.checkpw(pwd_bytes, stored_hash_b):
        return (user_id, first_name, last_name, status_id)
    else:
        return None

class LoginWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        settings = QtCore.QSettings("Procesos Y Servicios", "CapturadorDeDatos")
        stored = settings.value("theme", "dark")
        self.is_dark = (stored == "dark")



        # y fija el icono del botón
        self.setWindowTitle("Login - Capturador De Datos")
        screen = QtWidgets.QApplication.primaryScreen().availableGeometry()
        win_w = int(max(480, min(1200, screen.width() * 0.6)))
        win_h = int(max(500, min(1000, screen.height() * 0.85)))
        self.resize(win_w, win_h)
        self.center_on_screen()
        self.setWindowFlag(QtCore.Qt.WindowMaximizeButtonHint, False)
        self.setFixedSize(win_w, win_h)

        panel_w = int(win_w * 0.57)
        panel_h = int(win_h * 0.62)
        x = (win_w - panel_w) // 2
        y = (win_h - panel_h) // 2
        self.panel_rect = QtCore.QRect(x, y, panel_w, panel_h)

        # Cargamos la imagen de fondo completa
        bg_path = os.path.join(os.path.dirname(__file__), "Fondo.png")
        full_pix = None
        if os.path.exists(bg_path):
            pix = QtGui.QPixmap(bg_path)
            if not pix.isNull():
                full_pix = pix.scaled(
                    self.size(),
                    QtCore.Qt.IgnoreAspectRatio,
                    QtCore.Qt.SmoothTransformation
                )

        # — Carga de los fondos oscuro y claro —
        self.bg_dark = QtGui.QPixmap(resource_path("FondoLoginDark.png"))
        if not self.bg_dark.isNull():
            self.bg_dark = self.bg_dark.scaled(
                self.size(),
                QtCore.Qt.IgnoreAspectRatio,
                QtCore.Qt.SmoothTransformation
            )

        self.bg_light = QtGui.QPixmap(resource_path("FondoLoginWhite.png"))
        if not self.bg_light.isNull():
            self.bg_light = self.bg_light.scaled(
                self.size(),
                QtCore.Qt.IgnoreAspectRatio,
                QtCore.Qt.SmoothTransformation
            )

        # — Widget de fondo completo (siempre creado) —
        self.bg_label = QtWidgets.QLabel(self)
        # Si cargaste full_pix úsalo; si no, usa el fondo según el tema
        if full_pix:
            self.bg_label.setPixmap(full_pix)
        else:
            self.bg_label.setPixmap(self.bg_dark if self.is_dark else self.bg_light)

        # Ocupa toda la ventana
        self.bg_label.setGeometry(self.rect())
        # Y lo enviamos al fondo
        self.bg_label.lower()

        # Zona difuminada
        self.blurred_bg = QtWidgets.QLabel(self)
        blur_fx = QGraphicsBlurEffect(self.blurred_bg)
        blur_fx.setBlurRadius(15)
        self.blurred_bg.setGraphicsEffect(blur_fx)

        # Copiamos del fondo correspondiente inicialmente
        base_pix = self.bg_dark if self.is_dark else self.bg_light
        self.blurred_bg.setPixmap(base_pix.copy(self.panel_rect))
        self.blurred_bg.setGeometry(self.panel_rect)

        # — Panel central semitransparente —
        self.panel = QtWidgets.QFrame(self)
        self.panel.setObjectName("centralPanel")
        self.panel.setGeometry(self.panel_rect)
        self.panel.setFocusPolicy(QtCore.Qt.NoFocus)

        shadow = QGraphicsDropShadowEffect(self.panel)
        shadow.setBlurRadius(30)
        shadow.setOffset(0, 5)
        shadow.setColor(QtGui.QColor(0,0,0,80))
        self.panel.setGraphicsEffect(shadow)


        # Aseguramos el orden de apilamiento
        if full_pix:
            self.bg_label.lower()
            self.blurred_bg.stackUnder(self.panel)
        self.panel.raise_()

        # — Contenido interno del panel —
        vbox = QtWidgets.QVBoxLayout(self.panel)
        vbox.setContentsMargins(40, 30, 40, 30)
        vbox.setSpacing(30)

       
        logo_path_dark  = resource_path("LogoImg_light.png")
        logo_path_light = resource_path("LogoImg_dark.png")
        self.doc_icon_dark  = QtGui.QPixmap(resource_path("doc_white.png")).scaled(40, 40, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
        self.doc_icon_light = QtGui.QPixmap(resource_path("doc_black.png")).scaled(40, 40, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
        self.lock_icon_dark  = QtGui.QPixmap(resource_path("lock_white.png")).scaled(40, 40, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
        self.lock_icon_light = QtGui.QPixmap(resource_path("lock_black.png")).scaled(40, 40, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)

        # Creamos el QLabel que contendrá el logo
        self.logo_label = QtWidgets.QLabel(self.panel)
        self.logo_label.setStyleSheet("background-color: transparent;")
        self.logo_label.setAttribute(QtCore.Qt.WA_TranslucentBackground)

        # Precargamos ambos pixmaps
        orig_dark = QtGui.QPixmap(logo_path_dark)
        self.pix_dark = orig_dark.scaled(
            150, 150,
            QtCore.Qt.KeepAspectRatio,
            QtCore.Qt.SmoothTransformation
        )
        self.pix_light = QtGui.QPixmap(logo_path_light).scaled(
            150, 150,
            QtCore.Qt.KeepAspectRatio,
            QtCore.Qt.SmoothTransformation
        )

        # Arrancamos mostrando el logo para modo oscuro
        self.logo_label.setPixmap(self.pix_dark)
        # fijas tamaño del contenedor de logo a 150×150
        self.logo_label.setFixedSize(150, 150)
        # centras el pixmap dentro de ese contenedor
        self.logo_label.setAlignment(QtCore.Qt.AlignCenter)
        # permites que si el pixmap es más pequeño, aparezca centrado sin recortes
        self.logo_label.setScaledContents(False)
        vbox.addWidget(self.logo_label, alignment=QtCore.Qt.AlignCenter)

        # Título
        lbl_text = QtWidgets.QLabel("Iniciar Sesión", alignment=QtCore.Qt.AlignCenter)
        lbl_text.setStyleSheet("""
            font-size: 25px;
            font-weight: bold;
            background: transparent;
        """)
        vbox.addWidget(lbl_text)

        # Documento
        hdoc = QtWidgets.QHBoxLayout()
        self.lbl_doc = QtWidgets.QLabel(self.panel)
        self.lbl_doc.setPixmap(self.doc_icon_dark)
        self.lbl_doc.setFixedWidth(40)
        self.lbl_doc.setStyleSheet("font-size: 30px; background: transparent;")
        hdoc.addWidget(self.lbl_doc)

        lbl_doctxt = QtWidgets.QLabel("Documento:")
        lbl_doctxt.setStyleSheet("font-size: 16px; font-weight: bold; background: transparent;")
        hdoc.addWidget(lbl_doctxt)

        self.edit_doc = QtWidgets.QLineEdit()
        self.edit_doc.setPlaceholderText("12345678")
        regex = QtCore.QRegularExpression(r"\d*")
        validator = QtGui.QRegularExpressionValidator(regex, self.edit_doc)
        self.edit_doc.setValidator(validator)
        hdoc.addWidget(self.edit_doc, stretch=1)

        vbox.addLayout(hdoc)


        # Contraseña
        hpwd = QtWidgets.QHBoxLayout()
        self.lbl_pwd = QtWidgets.QLabel(self.panel)
        self.lbl_pwd.setPixmap(self.lock_icon_dark)
        self.lbl_pwd.setFixedWidth(40)
        self.lbl_pwd.setStyleSheet("font-size: 30px; background: transparent;")
        hpwd.addWidget(self.lbl_pwd)

        lbl_pwdtxt = QtWidgets.QLabel("Contraseña:")
        lbl_pwdtxt.setStyleSheet("font-size: 16px; font-weight: bold; background: transparent;")
        hpwd.addWidget(lbl_pwdtxt)

        self.edit_pwd = QtWidgets.QLineEdit()
        self.edit_pwd.setEchoMode(QtWidgets.QLineEdit.Password)
        self.edit_pwd.setPlaceholderText("••••••••")
        hpwd.addWidget(self.edit_pwd, stretch=1)

        vbox.addLayout(hpwd)


        # Botón Iniciar sesión
        btn = QtWidgets.QPushButton("Iniciar sesión", self.panel)
        btn.setFixedSize(200, 50)
        btn.setStyleSheet("""
            QPushButton {
                background-color: #007BFF;
                border-radius: 25px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #339CFF;
            }
        """)
        btn.clicked.connect(self.on_login)
        vbox.addWidget(btn, alignment=QtCore.Qt.AlignCenter)

        # Olvidaste tu contraseña
        btn_forgot = QtWidgets.QPushButton("¿Olvidaste tu contraseña?", self.panel)
        btn_forgot.setObjectName("forgotBtn")
        btn_forgot.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                font-size: 18px;
                font-weight: bold;
            }
            QPushButton:hover {
                color: #339CFF;
            }
        """)
        btn_forgot.clicked.connect(self.on_forgot_password)
        vbox.addWidget(btn_forgot, alignment=QtCore.Qt.AlignCenter)
        
        # — Estado inicial de tema —
                # — Botón circulito para cambiar tema —
        self.theme_btn = QtWidgets.QPushButton(self)
        self.theme_btn.setAttribute(QtCore.Qt.WA_TranslucentBackground, True)
        self.theme_btn.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
            }
            QPushButton:hover {
                background-color: rgba(255,255,255,30);  /* ligero resalte */
            }
        """)
        self.theme_btn.setToolTip("Cambiar tema día/noche")
        self.theme_btn.setCursor(QtCore.Qt.PointingHandCursor)
        self.theme_btn.setFixedSize(50, 50)
        self.theme_btn.setFlat(True)

        # Cargamos y escalamos los pixmaps de luna y sol
        moon_pix = QtGui.QPixmap(resource_path("moon.png")) \
            .scaled(50, 50, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
        sun_pix  = QtGui.QPixmap(resource_path("sun.png")) \
            .scaled(50, 50, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)

        self.moon_icon = QtGui.QIcon(moon_pix)
        self.sun_icon  = QtGui.QIcon(sun_pix)

        # Glow effect (único “borde” visible)
        self.glow = QGraphicsDropShadowEffect(self.theme_btn)
        self.glow.setBlurRadius(20)
        self.glow.setOffset(0)
        self.theme_btn.setGraphicsEffect(self.glow)

        # Posición abajo-derecha
        self.theme_btn.move(self.width() - 60, self.height() - 60)
        self.theme_btn.clicked.connect(self.toggle_theme)
        self.theme_btn.raise_()
        self.theme_btn.setIcon(self.moon_icon if self.is_dark else self.sun_icon)
        self.theme_btn.setIconSize(QtCore.QSize(50, 50))
        self.glow.setColor(
            QtGui.QColor(200,200,255,180) if self.is_dark
            else QtGui.QColor(255,200,50,180)
        )

        self.apply_theme(self.is_dark)


    def center_on_screen(self):
        screen = QtWidgets.QApplication.primaryScreen().availableGeometry()
        size = self.geometry()
        x = (screen.width() - size.width()) // 2
        y = (screen.height() - size.height()) // 2
        self.move(x, y)

    def on_login(self):
        num_doc = self.edit_doc.text().strip()
        pwd     = self.edit_pwd.text().strip()

        if not num_doc or not pwd:
            QtWidgets.QMessageBox.warning(self, "Datos faltantes", "Debe ingresar documento y contraseña.")
            return

        auth = authenticate_user_by_doc(num_doc, pwd)
        if auth is None:
            QtWidgets.QMessageBox.critical(self, "Error", "Documento o contraseña incorrectos.")
            return

        user_id, first_name, last_name, status_id = auth
        if status_id != 5:
            QtWidgets.QMessageBox.warning(self, "Usuario inactivo", "Tu cuenta no está activa. Contacta al administrador.")
            return

        self.hide()  # oculto el login
        theme = "dark" if self.is_dark else "light"

        self.dashboard = DashboardWindow(
            user_id,
            first_name,
            last_name,
            theme=theme
        )
        self.dashboard.show()

    def on_forgot_password(self):
        # Aquí llamamos a la ventana de recuperación de contraseña
        self.close()
        self.recover_window = RecuperarContrasenaWindow(login_window=self)
        self.recover_window.show()

    def apply_theme(self, dark: bool):
        app = QtWidgets.QApplication.instance()
        if dark:
            style = """
                QWidget { background-color: #2b2b2b; color: #f0f0f0; }
                QFrame#centralPanel { background-color: rgba(0,0,0,125); border-radius: 20px; }
                QLineEdit { background-color: rgba(0,0,0,200); color: #f0f0f0; font-size: 15px; border-radius: 15px; padding: 8px 12px; }
                QPushButton { background-color: #007BFF; color: #ffffff; border-radius: 25px; font-size: 16px; }
                QPushButton:hover { background-color: #339CFF; }
                QPushButton#forgotBtn { background-color: transparent; color: #ffffff; font-size: 18px; font-weight: bold; }
                QPushButton#forgotBtn:hover { color: #339CFF; }
            """
            self.bg_label.setPixmap(self.bg_dark)
            self.blurred_bg.setPixmap(self.bg_dark.copy(self.panel_rect))
            self.lbl_doc.setPixmap(self.doc_icon_dark)
            self.lbl_pwd.setPixmap(self.lock_icon_dark)
            self.logo_label.setPixmap(self.pix_dark)
            self.theme_btn.setIcon(self.moon_icon)
            self.glow.setColor(QtGui.QColor(200,200,255,180))
        else:
            style = """
                QWidget { background-color: #f4f4f9; color: #000000; }
                QFrame#centralPanel { background-color: rgba(255,255,255, 200); border-radius: 20px; border: 1px solid rgba(0,0,0,0.1); }
                QLineEdit { background-color: #ffffff; color: #000000; border: 1px solid rgba(0,0,0,0.15); border-radius: 15px; font-size: 15px; padding: 8px 12px; }
                QPushButton { background-color: #1E90FF; color: #ffffff; border-radius: 25px; }
                QPushButton:hover { background-color: #006FDE; }
                QPushButton#forgotBtn { background-color: transparent; color: #000000; font-size: 18px; font-weight: bold; }
                QPushButton#forgotBtn:hover { color: #339CFF; }
            """
            self.bg_label.setPixmap(self.bg_light)
            self.blurred_bg.setPixmap(self.bg_light.copy(self.panel_rect))
            self.lbl_doc.setPixmap(self.doc_icon_light)
            self.lbl_pwd.setPixmap(self.lock_icon_light)
            self.logo_label.setPixmap(self.pix_light)
            self.theme_btn.setIcon(self.sun_icon)
            self.glow.setColor(QtGui.QColor(255,200,50,180))
        app.setStyleSheet(style)
        
    def toggle_theme(self):
        settings = QtCore.QSettings("Procesos Y Servicios", "CapturadorDeDatos")
        self.is_dark = not self.is_dark
        self.apply_theme(self.is_dark)
        settings.setValue("theme", "dark" if self.is_dark else "light")

# Función para enviar código por correo
def enviar_codigo_por_email(email_destino):
    # Generar un código aleatorio de 6 dígitos
    codigo = ''.join(random.choices(string.digits, k=6))

    # Configura el servidor SMTP (usando Gmail como ejemplo)
    try:
        # Conexión al servidor SMTP de Gmail
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()

        # Tu dirección de correo y contraseña (asegúrate de usar una contraseña de aplicación si usas Gmail)
        servidor_email = 'jhoanduartedg@gmail.com'
        contrasena = 'abbb kces zryg isxc'  # Usa una contraseña de aplicación aquí

        # Crear el mensaje
        mensaje = f"""
        <html>
            <body style="font-family: Arial, sans-serif; color: #333; background-color: #f4f4f9; padding: 20px;">
                <div style="text-align: center; padding: 20px;">
                    <img src="https://raw.githubusercontent.com/JhoanDuarte/Capturador_Actualizacioness/main/LogoImg.png" alt="Logo" style="width: 150px; margin-bottom: 20px;">
                    <h2 style="color: #4CAF50;">Recuperación de Contraseña</h2>
                    <p style="font-size: 16px;">¡Hola! Para completar el proceso de recuperación de contraseña, por favor ingresa el siguiente código:</p>
                    <h3 style="font-size: 28px; color: #4CAF50; font-weight: bold;">{codigo}</h3>
                    <p style="font-size: 16px; color: #777;">Este código es válido solo por 15 minutos.</p>
                    <p style="font-size: 14px; color: #777;">Si no solicitaste este cambio, ignora este correo.</p>
                    <hr style="border: 0; border-top: 1px solid #ddd; margin: 20px 0;">
                    <footer style="font-size: 12px; color: #aaa;">
                        <p>&copy; 2025 - Dashboard Capturador Datos. Todos los derechos reservados. PYS - Jhoan David Duarte Guayazan - Numero de contacto soporte: 3135517480 </p>
                    </footer>
                </div>
            </body>
        </html>
        """

        # Crear el mensaje como MIMEText para asegurar codificación en UTF-8
        msg = MIMEMultipart()
        msg['From'] = servidor_email
        msg['To'] = email_destino
        msg['Subject'] = "Recuperación de Contraseña"
        
        # Adjuntar el cuerpo HTML
        msg.attach(MIMEText(mensaje, 'html', 'utf-8'))

        # Enviar el mensaje
        server.login(servidor_email, contrasena)
        server.sendmail(servidor_email, email_destino, msg.as_string())
        server.quit()

        return codigo  # Retorna el código para su validación
    except Exception as e:
        error_box = QMessageBox()
        error_box.setIcon(QMessageBox.Critical)
        error_box.setWindowTitle("Error al Enviar Correo")
        error_box.setText(f"Hubo un problema al enviar el correo:\n{str(e)}")
        error_box.exec_()
        return None

class RecuperarContrasenaWindow(QtWidgets.QWidget):
    def __init__(self, login_window):
        super().__init__()
        self.login_window = login_window

        # — Configuración básica de la ventana —
        self.setWindowTitle("Recuperar Contraseña")
        screen = QtWidgets.QApplication.primaryScreen().availableGeometry()
        win_w = int(max(420, min(800, screen.width() * 0.5)))
        win_h = int(max(320, min(700, screen.height() * 0.6)))
        self.resize(win_w, win_h)
        self.setWindowFlag(QtCore.Qt.WindowMaximizeButtonHint, False)
        self.setFixedSize(win_w, win_h)

        panel_width  = int(win_w * 0.8)
        panel_height = int(win_h * 0.6)
        x = (win_w - panel_width) // 2
        y = (win_h - panel_height) // 2
        self.panel_rect = QtCore.QRect(x, y, panel_width, panel_height)
        
        # — Fondos escalados con suavizado —
        self.bg_dark = QtGui.QPixmap(resource_path("FondoLoginDark.png")) \
            .scaled(self.size(), QtCore.Qt.IgnoreAspectRatio, QtCore.Qt.SmoothTransformation)
        self.bg_light = QtGui.QPixmap(resource_path("FondoLoginWhite.png")) \
            .scaled(self.size(), QtCore.Qt.IgnoreAspectRatio, QtCore.Qt.SmoothTransformation)

        # — Label de fondo según tema —
        self.bg_label = QtWidgets.QLabel(self)
        bg_pix = self.bg_dark if self.login_window.is_dark else self.bg_light
        self.bg_label.setPixmap(bg_pix)
        self.bg_label.setGeometry(self.rect())
        self.bg_label.setScaledContents(True)

        # — Iconos preparados (blanco para oscuro, negro para claro) —
        self.doc_icon_white = QtGui.QPixmap(resource_path("doc_white.png")) \
            .scaled(28, 28, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
        self.doc_icon_black = QtGui.QPixmap(resource_path("doc_black.png")) \
            .scaled(28, 28, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)

        # — Zona difuminada detrás del panel —
        self.blurred_bg = QtWidgets.QLabel(self)
        blur_fx = QGraphicsBlurEffect(self.blurred_bg)
        blur_fx.setBlurRadius(15)
        self.blurred_bg.setGraphicsEffect(blur_fx)
        self.blurred_bg.setPixmap(bg_pix.copy(self.panel_rect))
        self.blurred_bg.setGeometry(self.panel_rect)

        # — Panel semitransparente y con sombra —
        self.panel = QtWidgets.QFrame(self)
        self.panel.setObjectName("centralPanel")
        self.panel.setGeometry(self.panel_rect)
        shadow = QGraphicsDropShadowEffect(self.panel)
        shadow.setBlurRadius(30)
        shadow.setOffset(0, 5)
        shadow.setColor(QtGui.QColor(0, 0, 0, 80))
        self.panel.setGraphicsEffect(shadow)

        # Aseguramos orden de apilado
        self.bg_label.lower()
        self.blurred_bg.stackUnder(self.panel)
        self.panel.raise_()

        # — Layout interno del panel —
        vbox = QtWidgets.QVBoxLayout(self.panel)
        vbox.setContentsMargins(40, 30, 40, 30)
        vbox.setSpacing(20)

        # — Título —
        lbl_title = QtWidgets.QLabel("Recuperar Contraseña", self.panel)
        lbl_title.setAlignment(QtCore.Qt.AlignCenter)
        lbl_title.setStyleSheet(f"""
            QLabel {{
                color: {'#f0f0f0' if self.login_window.is_dark else '#000000'};
                font-size: 24px;
                font-weight: bold;
                background-color: transparent;
            }}
        """)
        vbox.addWidget(lbl_title)
        vbox.addStretch(1)

        # — Campo Documento con icono —
        hdoc = QtWidgets.QHBoxLayout()
        lbl_doc = QtWidgets.QLabel(self.panel)
        # Fondo transparente en el label
        lbl_doc.setAttribute(QtCore.Qt.WA_TranslucentBackground)
        # Icono según tema
        icon = self.doc_icon_white if self.login_window.is_dark else self.doc_icon_black
        lbl_doc.setPixmap(icon)
        lbl_doc.setFixedSize(28, 28)
        hdoc.addWidget(lbl_doc)

        lbl_doctxt = QtWidgets.QLabel("Documento:", self.panel)
        lbl_doctxt.setStyleSheet(f"""
            QLabel {{
                color: {'#f0f0f0' if self.login_window.is_dark else '#000000'};
                font-size: 16px;
                font-weight: bold;
                background-color: transparent;
            }}
        """)
        hdoc.addWidget(lbl_doctxt)

        self.edit_doc = QtWidgets.QLineEdit(self.panel)
        self.edit_doc.setPlaceholderText("12345678")
        regex = QtCore.QRegularExpression(r"\d*")
        validator = QtGui.QRegularExpressionValidator(regex, self.edit_doc)
        self.edit_doc.setValidator(validator)
        self.edit_doc.setStyleSheet(f"""
            QLineEdit {{
                background-color: {'rgba(0,0,0,200)' if self.login_window.is_dark else '#ffffff'};
                color: {'#f0f0f0' if self.login_window.is_dark else '#000000'};
                border: 1px solid rgba(0,0,0,0.15);
                border-radius: 15px;
                padding: 8px 12px;
                font-size: 15px;
            }}
        """)
        hdoc.addWidget(self.edit_doc, stretch=1)
        vbox.addLayout(hdoc)
        vbox.addStretch(1)

        # — Botón Enviar código —
        btn_send = QtWidgets.QPushButton("Enviar código", self.panel)
        btn_send.setFixedSize(200, 50)
        btn_send.setStyleSheet("""
            QPushButton {
                background-color: #007BFF;
                color: #ffffff;
                border-radius: 25px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #339CFF;
            }
        """)
        btn_send.clicked.connect(self.enviar_codigo)
        vbox.addWidget(btn_send, alignment=QtCore.Qt.AlignCenter)
        vbox.addStretch(2)

        # — Centrar el panel dentro de la ventana —
        x = (self.width() - self.panel_rect.width()) // 2
        y = (self.height() - self.panel_rect.height()) // 2
        self.panel_rect.moveTo(x, y)
        self.panel.setGeometry(self.panel_rect)
        self.blurred_bg.setGeometry(self.panel_rect)

        
    def closeEvent(self, event):
        # cuando cierras esta ventana, reaparece el login
        if event.spontaneous():
            self.login_window.show()
        event.accept()

    def enviar_codigo(self):
        num_doc = self.edit_doc.text().strip()
        if not num_doc:
            QtWidgets.QMessageBox.warning(self, "Datos faltantes", "Debe ingresar su documento.")
            return

        cursor = conn.cursor()
        cursor.execute("SELECT CORREO FROM USERS WHERE NUM_DOC = %s", (num_doc,))
        row = cursor.fetchone()
        cursor.close()

        if not row:
            QtWidgets.QMessageBox.warning(self, "Usuario no encontrado", "No se encontró un usuario con ese documento.")
            return

        codigo = enviar_codigo_por_email(row[0])
        if codigo:
            self.codigo_recibido = codigo
            QtWidgets.QMessageBox.information(self, "Código Enviado", "Te hemos enviado un código por correo electrónico.")
            self.close()
            self.mostrar_ventana_codigo()
        else:
            QtWidgets.QMessageBox.critical(self, "Error", "Hubo un problema al enviar el correo.")

    def mostrar_ventana_codigo(self):
        win = QtWidgets.QWidget()
        win.login_window = self.login_window
        win.setWindowTitle("Verificar Código")
        screen = QtWidgets.QApplication.primaryScreen().availableGeometry()
        win_w = int(max(420, min(800, screen.width() * 0.5)))
        win_h = int(max(320, min(700, screen.height() * 0.6)))
        win.resize(win_w, win_h)
        win.setWindowFlag(QtCore.Qt.WindowMaximizeButtonHint, False)
        win.setFixedSize(win_w, win_h)

        # — Cargamos y escalamos ambos fondos —
        bg_dark  = QtGui.QPixmap(resource_path("FondoLoginDark.png")).scaled(
            win.size(), QtCore.Qt.IgnoreAspectRatio, QtCore.Qt.SmoothTransformation
        )
        bg_light = QtGui.QPixmap(resource_path("FondoLoginWhite.png")).scaled(
            win.size(), QtCore.Qt.IgnoreAspectRatio, QtCore.Qt.SmoothTransformation
        )

        # — QLabel de fondo según tema —
        win.bg_label = QtWidgets.QLabel(win)
        if self.login_window.is_dark:
            win.bg_label.setPixmap(bg_dark)
        else:
            win.bg_label.setPixmap(bg_light)
        win.bg_label.setGeometry(win.rect())

        # Definimos geometría del panel
        panel_w = int(win_w * 0.8)
        panel_h = int(win_h * 0.6)
        panel_rect = QtCore.QRect(
            (win_w - panel_w) // 2,
            (win_h - panel_h) // 2,
            panel_w,
            panel_h,
        )

        # — Blur detrás del panel —
        blurred_bg = QtWidgets.QLabel(win)
        base_pix = bg_dark if self.login_window.is_dark else bg_light
        crop = base_pix.copy(panel_rect)
        blurred_bg.setPixmap(crop)
        blurred_bg.setGeometry(panel_rect)
        blur_fx = QGraphicsBlurEffect(blurred_bg)
        blur_fx.setBlurRadius(15)
        blurred_bg.setGraphicsEffect(blur_fx)

        # — Panel semitransparente —
        panel = QtWidgets.QFrame(win)
        panel.setObjectName("centralPanel")
        panel.setGeometry(panel_rect)
        panel.setStyleSheet(f"""
            QFrame#centralPanel {{
                background-color: rgba({ '0,0,0,125' if self.login_window.is_dark else '255,255,255,200' });
                border-radius: 20px;
                border: 1px solid rgba(0,0,0,0.1);
            }}
        """)
        shadow = QGraphicsDropShadowEffect(panel)
        shadow.setBlurRadius(20)
        shadow.setOffset(0, 4)
        shadow.setColor(QtGui.QColor(0, 0, 0, 160))
        panel.setGraphicsEffect(shadow)
        win.bg_label.lower()
        blurred_bg.stackUnder(panel)
        panel.raise_()

        # — Layout interno —
        layout = QtWidgets.QVBoxLayout(panel)
        layout.setContentsMargins(40, 30, 40, 30)
        layout.setSpacing(20)

        # Carga de iconos según tema
        icon_color = "white" if self.login_window.is_dark else "black"
        doc_icon = QtGui.QPixmap(resource_path(f"doc_{icon_color}.png")).scaled(28,28,QtCore.Qt.KeepAspectRatio,QtCore.Qt.SmoothTransformation)

        # Etiqueta de instrucción
        lbl = QtWidgets.QLabel("Ingresa el código de recuperación enviado a tu correo:", panel)
        lbl.setAlignment(QtCore.Qt.AlignCenter)
        lbl.setWordWrap(True)
        lbl.setStyleSheet(f"""
            QLabel {{
                color: {'#f0f0f0' if self.login_window.is_dark else '#000000'};
                font-size: 16px;
                font-weight: bold;
                background-color: transparent;
            }}
        """)
        layout.addWidget(lbl)

        # Campo de código
        self.edit_codigo = QtWidgets.QLineEdit(panel)
        self.edit_codigo.setPlaceholderText("Código")
        self.edit_codigo.setStyleSheet(f"""
            QLineEdit {{
                background-color: {'rgba(0,0,0,200)' if self.login_window.is_dark else '#ffffff'};
                color: {'#f0f0f0' if self.login_window.is_dark else '#000000'};
                border: 1px solid rgba(0,0,0,0.15);
                border-radius: 15px;
                padding: 8px 12px;
                font-size: 15px;
            }}
        """)
        layout.addWidget(self.edit_codigo)

        # Botón verificar
        btn = QtWidgets.QPushButton("Verificar Código", panel)
        btn.setFixedSize(200, 50)
        btn.setStyleSheet("""
            QPushButton {
                background-color: #007BFF;
                color: #ffffff;
                border-radius: 25px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #339CFF;
            }
        """)
        btn.clicked.connect(self.verificar_codigo)
        layout.addWidget(btn, alignment=QtCore.Qt.AlignCenter)

        win.show()
        win.closeEvent = lambda e: (self.login_window.show(), e.accept()) if e.spontaneous() else e.accept()
        self.codigo_window = win


    def mostrar_ventana_cambio_contrasena(self):
        win = QtWidgets.QWidget()
        win.login_window = self.login_window
        win.setWindowTitle("Cambiar Contraseña")
        screen = QtWidgets.QApplication.primaryScreen().availableGeometry()
        win_w = int(max(380, min(700, screen.width() * 0.4)))
        win_h = int(max(280, min(600, screen.height() * 0.5)))
        win.resize(win_w, win_h)
        win.setWindowFlag(QtCore.Qt.WindowMaximizeButtonHint, False)
        win.setFixedSize(win_w, win_h)

        # — Fondos según tema —
        bg_dark  = QtGui.QPixmap(resource_path("FondoLoginDark.png")).scaled(
            win.size(), QtCore.Qt.IgnoreAspectRatio, QtCore.Qt.SmoothTransformation
        )
        bg_light = QtGui.QPixmap(resource_path("FondoLoginWhite.png")).scaled(
            win.size(), QtCore.Qt.IgnoreAspectRatio, QtCore.Qt.SmoothTransformation
        )
        win.bg_label = QtWidgets.QLabel(win)
        win.bg_label.setPixmap(bg_dark if self.login_window.is_dark else bg_light)
        win.bg_label.setGeometry(win.rect())

        # Panel
        panel_w = int(win_w * 0.75)
        panel_h = int(win_h * 0.8)
        panel_rect = QtCore.QRect(
            (win_w - panel_w) // 2,
            (win_h - panel_h) // 2,
            panel_w,
            panel_h,
        )
        blurred_bg = QtWidgets.QLabel(win)
        base_pix = bg_dark if self.login_window.is_dark else bg_light
        crop = base_pix.copy(panel_rect)
        blurred_bg.setPixmap(crop)
        blurred_bg.setGeometry(panel_rect)
        blur_fx = QGraphicsBlurEffect(blurred_bg)
        blur_fx.setBlurRadius(15)
        blurred_bg.setGraphicsEffect(blur_fx)

        panel = QtWidgets.QFrame(win)
        panel.setObjectName("centralPanel")
        panel.setGeometry(panel_rect)
        panel.setStyleSheet(f"""
            QFrame#centralPanel {{
                background-color: rgba({ '0,0,0,125' if self.login_window.is_dark else '255,255,255,200' });
                border-radius: 20px;
                border: 1px solid rgba(0,0,0,0.1);
            }}
        """)
        shadow = QGraphicsDropShadowEffect(panel)
        shadow.setBlurRadius(20)
        shadow.setOffset(0, 4)
        shadow.setColor(QtGui.QColor(0, 0, 0, 160))
        panel.setGraphicsEffect(shadow)
        win.bg_label.lower()
        blurred_bg.stackUnder(panel)
        panel.raise_()

        # Layout y widgets
        layout = QtWidgets.QVBoxLayout(panel)
        layout.setContentsMargins(40, 30, 40, 30)
        layout.setSpacing(20)

        lbl = QtWidgets.QLabel("Ingresa tu nueva contraseña:", panel)
        lbl.setAlignment(QtCore.Qt.AlignCenter)
        lbl.setStyleSheet(f"""
            QLabel {{
                color: {'#f0f0f0' if self.login_window.is_dark else '#000000'};
                font-size: 16px;
                font-weight: bold;
                background-color: transparent;
            }}
        """)
        layout.addWidget(lbl)

        self.edit_new_pwd = QtWidgets.QLineEdit(panel)
        self.edit_new_pwd.setEchoMode(QtWidgets.QLineEdit.Password)
        self.edit_new_pwd.setPlaceholderText("Nueva contraseña")
        self.edit_new_pwd.setStyleSheet(f"""
            QLineEdit {{
                background-color: {'rgba(0,0,0,200)' if self.login_window.is_dark else '#ffffff'};
                color: {'#f0f0f0' if self.login_window.is_dark else '#000000'};
                border: 1px solid rgba(0,0,0,0.15);
                border-radius: 15px;
                padding: 8px 12px;
                font-size: 15px;
            }}
        """)
        layout.addWidget(self.edit_new_pwd)

        btn = QtWidgets.QPushButton("Cambiar Contraseña", panel)
        btn.setFixedSize(200, 50)
        btn.setStyleSheet("""
            QPushButton {
                background-color: #007BFF;
                color: #ffffff;
                border-radius: 25px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #339CFF;
            }
        """)
        btn.clicked.connect(self.cambiar_contrasena)
        layout.addWidget(btn, alignment=QtCore.Qt.AlignCenter)

        win.show()
        win.closeEvent = lambda e: (self.login_window.show(), e.accept()) if e.spontaneous() else e.accept()
        self.cambio_window = win



    def verificar_codigo(self):
        if self.edit_codigo.text().strip() == getattr(self, "codigo_recibido", ""):
            self.codigo_window.close()
            self.mostrar_ventana_cambio_contrasena()
        else:
            QtWidgets.QMessageBox.warning(self, "Código incorrecto", "El código ingresado es incorrecto.")


    def cambiar_contrasena(self):
        new_pwd = self.edit_new_pwd.text().strip()
        if not new_pwd:
            QtWidgets.QMessageBox.warning(self, "Contraseña vacía", "La contraseña no puede estar vacía.")
            return

        # … aquí guardas la nueva contraseña en la BD …

        QtWidgets.QMessageBox.information(self, "Contraseña cambiada", "Tu contraseña ha sido cambiada exitosamente.")
        # Cierra la ventana de cambio
        self.cambio_window.close()
        # Vuelve a mostrar el login
        self.login_window.show()
def main():
    # Inicializa tu ventana de Tkinter (o de Qt, según cómo lo tengas montado)
    root = tk.Tk()
    # Aquí podrías hacer cualquier setup adicional que ya tenías,
    # por ejemplo:
    dashboard = DashboardWindow(root)
    # … configuración de layout, título, iconos, etc. …g
    root.mainloop()

if __name__ == "__main__":
    check_for_update_and_exit_if_needed()  # 👈 Esto va PRIMERO. Si no está actualizado, se sale.
    app = QtWidgets.QApplication(sys.argv)
    w = LoginWindow()
    w.show()
    sys.exit(app.exec_())
