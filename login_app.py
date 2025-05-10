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
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtCore import QRegularExpression
from PyQt5.QtGui  import QRegularExpressionValidator


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
    local_version = "1.2.1"  # Si no hay versión, se forzará la actualización

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
        self.setWindowTitle("Login - Capturador De Datos") 
        self.resize(700, 800)
        self.center_on_screen()

        # Layout principal
        main_layout = QtWidgets.QVBoxLayout(self)
        main_layout.addStretch()

        # Fondo
        bg_path = os.path.join(os.path.dirname(__file__), "Fondo.png")
        if os.path.exists(bg_path):
            palette = QtGui.QPalette()
            pix = QtGui.QPixmap(bg_path).scaled(self.size(), QtCore.Qt.IgnoreAspectRatio)
            palette.setBrush(QtGui.QPalette.Window, QtGui.QBrush(pix))
            self.setPalette(palette)

        # Panel central semitransparente
        self.panel = QtWidgets.QFrame()
        self.panel.setStyleSheet(""" 
            QFrame { 
                background-color: rgba(0, 0, 0, 150); 
                border-radius: 20px; 
            } 
        """)

        vbox = QtWidgets.QVBoxLayout(self.panel)
        vbox.setContentsMargins(40, 30, 40, 30)
        vbox.setSpacing(30)

        # Logo (o texto si no existe)
        logo_path = resource_path("LogoImg.png")
        if os.path.exists(logo_path):
            lbl_logo = QtWidgets.QLabel(self.panel)
            lbl_logo.setAttribute(QtCore.Qt.WA_TranslucentBackground)
            pixmap = QtGui.QPixmap(logo_path).scaled(140, 140,
                QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
            lbl_logo.setPixmap(pixmap)
            lbl_logo.setAlignment(QtCore.Qt.AlignCenter)
            vbox.addWidget(lbl_logo)
        else:
            lbl_logo = QtWidgets.QLabel(self.panel)
            lbl_logo.setText("")  # imagen vacía
            lbl_logo.setAlignment(QtCore.Qt.AlignCenter)
            vbox.addWidget(lbl_logo)

        # Texto debajo del logo
        lbl_text = QtWidgets.QLabel("Iniciar Sesión", alignment=QtCore.Qt.AlignCenter) #
        lbl_text.setStyleSheet("""
            color: white;
            font-size: 28px;  /* Tamaño aumentado */
            font-weight: bold;  /* Negrita */
            background: transparent;
        """)
        vbox.addWidget(lbl_text)

        # Documento
        hdoc = QtWidgets.QHBoxLayout()
        lbl_doc = QtWidgets.QLabel("📄")
        lbl_doc.setFixedWidth(28)
        lbl_doc.setStyleSheet("font-size: 22px; background: transparent;")
        hdoc.addWidget(lbl_doc)

        lbl_doctxt = QtWidgets.QLabel("Documento:")
        lbl_doctxt.setStyleSheet("color: white; background: transparent; font-size: 16px; font-weight: bold;")
        hdoc.addWidget(lbl_doctxt)

        self.edit_doc = QtWidgets.QLineEdit()
        self.edit_doc.setPlaceholderText("12345678")
        self.edit_doc.setStyleSheet("""
            QLineEdit { 
                background-color: rgba(0,0,0,200); 
                color: white; 
                border-radius: 15px; 
                padding: 8px 12px; 
                font-size: 14px; 
            } 
        """)
        regex = QRegularExpression(r"\d*")
        validator = QRegularExpressionValidator(regex, self.edit_doc)
        self.edit_doc.setValidator(validator)
        # Solo permitir dígitos, hasta 12 caracteres
        hdoc.addWidget(self.edit_doc, stretch=1)
        vbox.addLayout(hdoc)

        # Contraseña
        hpwd = QtWidgets.QHBoxLayout()
        lbl_pwd = QtWidgets.QLabel("🔒")
        lbl_pwd.setFixedWidth(28)
        lbl_pwd.setStyleSheet("font-size: 22px; background: transparent;")
        hpwd.addWidget(lbl_pwd)

        lbl_pwdtxt = QtWidgets.QLabel("Contraseña:")
        lbl_pwdtxt.setStyleSheet("color: white; background: transparent; font-size: 16px; font-weight: bold;")
        hpwd.addWidget(lbl_pwdtxt)

        self.edit_pwd = QtWidgets.QLineEdit()
        self.edit_pwd.setEchoMode(QtWidgets.QLineEdit.Password)
        self.edit_pwd.setPlaceholderText("••••••••")
        self.edit_pwd.setStyleSheet("""
            QLineEdit { 
                background-color: rgba(0,0,0,200); 
                color: white; 
                border-radius: 15px; 
                padding: 8px 12px; 
                font-size: 14px; 
            } 
        """)
        hpwd.addWidget(self.edit_pwd, stretch=1)
        vbox.addLayout(hpwd)

        # Botón redondeado y azul
        btn = QtWidgets.QPushButton("Iniciar sesión")
        btn.setFixedSize(200, 50)
        btn.setStyleSheet("""
            QPushButton { 
                background-color: #007BFF; 
                color: white; 
                border-radius: 25px; 
                font-size: 16px; 
            } 
            QPushButton:hover { 
                background-color: #339CFF; 
            } 
        """)
        btn.clicked.connect(self.on_login)
        vbox.addWidget(btn, alignment=QtCore.Qt.AlignCenter)

        # Botón "Olvidaste tu contraseña?"
        btn_forgot_password = QtWidgets.QPushButton("¿Olvidaste tu contraseña?")
        btn_forgot_password.setStyleSheet("""
            QPushButton { 
                background-color: transparent; 
                color: white; 
                font-size: 18px;  /* Tamaño aumentado */
                font-weight: bold;  /* Negrita */
            } 
            QPushButton:hover { 
                color: #339CFF; 
            } 
        """)
        btn_forgot_password.clicked.connect(self.on_forgot_password)
        vbox.addWidget(btn_forgot_password, alignment=QtCore.Qt.AlignCenter)

        # Centrar el panel en la ventana
        h_container = QtWidgets.QHBoxLayout()
        h_container.addStretch()
        h_container.addWidget(self.panel)
        h_container.addStretch()
        main_layout.addLayout(h_container)
        main_layout.addStretch()

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
        self.dashboard = DashboardWindow(user_id, first_name, last_name)
        self.dashboard.show()

    def on_forgot_password(self):
        # Aquí llamamos a la ventana de recuperación de contraseña
        self.close()
        self.recover_window = RecuperarContrasenaWindow(login_window=self)
        self.recover_window.show()

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
        self.setWindowTitle("Recuperar Contraseña")
        self.resize(400, 300)
        # — Fondo (misma imagen que en Login) —
        bg_path = resource_path("Fondo.png")
        if os.path.exists(bg_path):
            palette = QtGui.QPalette()
            pix = QtGui.QPixmap(bg_path).scaled(self.size(), QtCore.Qt.IgnoreAspectRatio)
            palette.setBrush(QtGui.QPalette.Window, QtGui.QBrush(pix))
            self.setPalette(palette)

        # — Panel semitransparente y redondeado —
        self.panel = QtWidgets.QFrame(self)
        self.panel.setStyleSheet("""
            QFrame {
                background-color: rgba(0, 0, 0, 150);
                border-radius: 20px;
            }
        """)

        # — Layout interno del panel —
        vbox = QtWidgets.QVBoxLayout(self.panel)
        vbox.setContentsMargins(40, 30, 40, 30)
        vbox.setSpacing(20)

        # Título
        lbl_title = QtWidgets.QLabel("Recuperar Contraseña", self.panel)
        lbl_title.setAlignment(QtCore.Qt.AlignCenter)
        lbl_title.setStyleSheet("color: white; font-size: 24px; font-weight: bold; background: transparent;")
        vbox.addWidget(lbl_title)

        vbox.addStretch(1)

        # Documento
        hdoc = QtWidgets.QHBoxLayout()
        lbl_doc = QtWidgets.QLabel("📄", self.panel)
        lbl_doc.setFixedWidth(28)
        lbl_doc.setStyleSheet("font-size: 22px; background: transparent;")
        hdoc.addWidget(lbl_doc)

        lbl_doctxt = QtWidgets.QLabel("Documento:", self.panel)
        lbl_doctxt.setStyleSheet("color: white; font-size: 16px; font-weight: bold; background: transparent;")
        hdoc.addWidget(lbl_doctxt)

        self.edit_doc = QtWidgets.QLineEdit(self.panel)
        self.edit_doc.setPlaceholderText("12345678")
        self.edit_doc.setStyleSheet("""
            QLineEdit {
                background-color: rgba(0,0,0,200);
                color: white;
                border-radius: 15px;
                padding: 8px 12px;
                font-size: 14px;
            }
            QLineEdit:focus {
                border: 1px solid #339CFF;
            }
        """)
        # Solo dígitos
        regex = QRegularExpression(r"\d*")
        validator = QRegularExpressionValidator(regex, self.edit_doc)
        self.edit_doc.setValidator(validator)
        hdoc.addWidget(self.edit_doc, stretch=1)
        vbox.addLayout(hdoc)

        vbox.addStretch(1)

        # Botón Enviar código
        btn_send = QtWidgets.QPushButton("Enviar código", self.panel)
        btn_send.setFixedSize(200, 50)
        btn_send.setStyleSheet("""
            QPushButton {
                background-color: #007BFF;
                color: white;
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

        # — Centrar panel en la ventana —
        main_layout = QtWidgets.QVBoxLayout(self)
        main_layout.addStretch()
        h_center = QtWidgets.QHBoxLayout()
        h_center.addStretch()
        h_center.addWidget(self.panel)
        h_center.addStretch()
        main_layout.addLayout(h_center)
        main_layout.addStretch()
        
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
        win.resize(400, 300)
        # Mismo fondo
        bg_path = resource_path("Fondo.png")
        if os.path.exists(bg_path):
            pal = QtGui.QPalette()
            pix = QtGui.QPixmap(bg_path).scaled(win.size(), QtCore.Qt.IgnoreAspectRatio)
            pal.setBrush(QtGui.QPalette.Window, QtGui.QBrush(pix))
            win.setPalette(pal)

        panel = QtWidgets.QFrame(win)
        panel.setStyleSheet("""
            QFrame {
                background-color: rgba(0,0,0,150);
                border-radius: 20px;
            }
        """)
        layout = QtWidgets.QVBoxLayout(panel)
        layout.setContentsMargins(40, 30, 40, 30)
        layout.setSpacing(20)

        lbl = QtWidgets.QLabel("Ingresa el código de recuperación enviado a tu correo:", panel)
        lbl.setAlignment(QtCore.Qt.AlignCenter)
        lbl.setStyleSheet("color: white; font-size: 16px; font-weight: bold; background: transparent;")
        layout.addWidget(lbl)

        self.edit_codigo = QtWidgets.QLineEdit(panel)
        self.edit_codigo.setPlaceholderText("Código")
        self.edit_codigo.setStyleSheet("""
            QLineEdit {
                background-color: rgba(0,0,0,200);
                color: white;
                border-radius: 15px;
                padding: 8px 12px;
                font-size: 14px;
            }
            QLineEdit:focus {
                border: 1px solid #339CFF;
            }
        """)
        layout.addWidget(self.edit_codigo)

        btn = QtWidgets.QPushButton("Verificar Código", panel)
        btn.setFixedSize(200, 50)
        btn.setStyleSheet("""
            QPushButton {
                background-color: #007BFF;
                color: white;
                border-radius: 25px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #339CFF;
            }
        """)
        btn.clicked.connect(self.verificar_codigo)
        layout.addWidget(btn, alignment=QtCore.Qt.AlignCenter)

        # Centrar
        main = QtWidgets.QVBoxLayout(win)
        main.addStretch()
        hc = QtWidgets.QHBoxLayout()
        hc.addStretch()
        hc.addWidget(panel)
        hc.addStretch()
        main.addLayout(hc)
        main.addStretch()

        def _on_close(event):
                    # solo si el cierre es manual por el usuario
            if event.spontaneous():
                win.login_window.show()
            event.accept()
        win.closeEvent = _on_close

                # 5) Guardamos y mostramos
        self.codigo_window = win
        win.show()

    def verificar_codigo(self):
        if self.edit_codigo.text().strip() == getattr(self, "codigo_recibido", ""):
            self.codigo_window.close()
            self.mostrar_ventana_cambio_contrasena()
        else:
            QtWidgets.QMessageBox.warning(self, "Código incorrecto", "El código ingresado es incorrecto.")

    def mostrar_ventana_cambio_contrasena(self):
        win = QtWidgets.QWidget()
        win.login_window = self.login_window
        win.setWindowTitle("Cambiar Contraseña")
        win.resize(400, 300)
        # Mismo fondo
        bg_path = resource_path("Fondo.png")
        if os.path.exists(bg_path):
            pal = QtGui.QPalette()
            pix = QtGui.QPixmap(bg_path).scaled(win.size(), QtCore.Qt.IgnoreAspectRatio)
            pal.setBrush(QtGui.QPalette.Window, QtGui.QBrush(pix))
            win.setPalette(pal)

        panel = QtWidgets.QFrame(win)
        panel.setStyleSheet("""
            QFrame {
                background-color: rgba(0,0,0,150);
                border-radius: 20px;
            }
        """)
        layout = QtWidgets.QVBoxLayout(panel)
        layout.setContentsMargins(40, 30, 40, 30)
        layout.setSpacing(20)

        lbl = QtWidgets.QLabel("Ingresa tu nueva contraseña:", panel)
        lbl.setAlignment(QtCore.Qt.AlignCenter)
        lbl.setStyleSheet("color: white; font-size: 16px; font-weight: bold; background: transparent;")
        layout.addWidget(lbl)

        self.edit_new_pwd = QtWidgets.QLineEdit(panel)
        self.edit_new_pwd.setEchoMode(QtWidgets.QLineEdit.Password)
        self.edit_new_pwd.setPlaceholderText("Nueva contraseña")
        self.edit_new_pwd.setStyleSheet("""
            QLineEdit {
                background-color: rgba(0,0,0,200);
                color: white;
                border-radius: 15px;
                padding: 8px 12px;
                font-size: 14px;
            }
            QLineEdit:focus {
                border: 1px solid #339CFF;
            }
        """)
        layout.addWidget(self.edit_new_pwd)

        btn = QtWidgets.QPushButton("Cambiar Contraseña", panel)
        btn.setFixedSize(200, 50)
        btn.setStyleSheet("""
            QPushButton {
                background-color: #007BFF;
                color: white;
                border-radius: 25px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #339CFF;
            }
        """)
        btn.clicked.connect(self.cambiar_contrasena)
        layout.addWidget(btn, alignment=QtCore.Qt.AlignCenter)

        # Centrar
        main = QtWidgets.QVBoxLayout(win)
        main.addStretch()
        hc = QtWidgets.QHBoxLayout()
        hc.addStretch()
        hc.addWidget(panel)
        hc.addStretch()
        main.addLayout(hc)
        main.addStretch()
        
        def _on_close(event):
            if event.spontaneous():
                win.login_window.show()
            event.accept()
        win.closeEvent = _on_close

        self.cambio_window = win
        win.show()

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


if __name__ == "__main__":
    check_for_update_and_exit_if_needed()  # 👈 Esto va PRIMERO. Si no está actualizado, se sale.
    app = QtWidgets.QApplication(sys.argv)
    w = LoginWindow()
    w.show()
    sys.exit(app.exec_())