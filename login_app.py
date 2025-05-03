import ctypes
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import io
import json
import logging
import os
import shutil
import sys
import subprocess
from threading import Thread
import time
from tkinter import messagebox
import customtkinter as ctk
import bcrypt
import smtplib
import random
import zipfile
import string
import requests
import tkinter as tk
from tkinter import ttk
from tqdm import tqdm
import psutil

# — Librerías estándar —
# (subprocess ya importado arriba si lo necesitas para llamadas externas)

# — Librerías de PyQt5 —
from PyQt5.QtGui import QIntValidator
from PyQt5 import QtCore, QtGui, QtWidgets

# — Módulos propios —
from db_connection import conectar_sql_server
from dashboard import open_dashboard

from version import __version__ as local_version

APP_NAME = "Dashboard_Capturador_Datos"
UPDATE_JSON_URL = "https://raw.githubusercontent.com/JhoanDuarte/Capturador_Actualizacioness/main/latest.json"

try:
    from version import __version__ as local_version
except ImportError:
    local_version = "1.0.4"  # Si no hay versión, se forzará la actualización

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
        root = ctk.CTk()
        open_dashboard(uid, fn, ln, parent=root)
        root.mainloop()
        sys.exit(0)

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
        logo_path = os.path.join(os.path.dirname(__file__), "LogoImg.png")
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
                text-decoration: underline;
            } 
            QPushButton:hover { 
                color: #FF7043; 
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

        # Lanza dashboard.py en un proceso separado
        dashboard_script = resource_path("dashboard.py")
        subprocess.Popen([
            sys.executable,
            dashboard_script,
            str(user_id),
            first_name,
            last_name
        ], cwd=os.path.dirname(dashboard_script))
        self.close()

    def on_forgot_password(self):
        # Aquí llamamos a la ventana de recuperación de contraseña
        self.recover_window = RecuperarContrasenaWindow()
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
                    <img src="cid:logo_img" alt="Logo" style="width: 150px; margin-bottom: 20px;">
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

        # Adjuntar la imagen del logo
        with open('LogoImg.png', 'rb') as img_file:
            img = MIMEImage(img_file.read())
            img.add_header('Content-ID', '<logo_img>')
            msg.attach(img)

        # Enviar el mensaje
        server.login(servidor_email, contrasena)
        server.sendmail(servidor_email, email_destino, msg.as_string())
        server.quit()

        return codigo  # Retorna el código para su validación
    except Exception as e:
        # Mostrar mensaje de error en la ventana emergente
        messagebox.critical(None, "Error al Enviar Correo", f"Hubo un problema al enviar el correo: {str(e)}")
        return None

class RecuperarContrasenaWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Recuperar Contraseña")
        self.resize(400, 300)

        # Fondo más oscuro
        self.setStyleSheet("""
            background: #121212;  /* Fondo oscuro */
            color: white;
            font-family: 'Arial', sans-serif;
        """)

        main_layout = QtWidgets.QVBoxLayout(self)
        
        # Título centrado y con un estilo más moderno
        lbl_title = QtWidgets.QLabel("Recuperar Contraseña")
        lbl_title.setStyleSheet("""
            color: white;
            font-size: 24px;
            font-weight: bold;
            text-align: center;
            margin-bottom: 20px;
        """)
        lbl_title.setAlignment(QtCore.Qt.AlignCenter)
        main_layout.addWidget(lbl_title)

        # Espacio para separar el título y el campo
        main_layout.addStretch(1)

        # Ingreso del número de documento
        hdoc = QtWidgets.QHBoxLayout()
        lbl_doc = QtWidgets.QLabel("📄")
        lbl_doc.setFixedWidth(28)
        lbl_doc.setStyleSheet("font-size: 22px; background: transparent;")
        hdoc.addWidget(lbl_doc)

        lbl_doctxt = QtWidgets.QLabel("Documento:")
        lbl_doctxt.setStyleSheet("""
            color: white;
            background: transparent;
            font-size: 16px;
            font-weight: bold;
        """)
        hdoc.addWidget(lbl_doctxt)

        self.edit_doc = QtWidgets.QLineEdit()
        self.edit_doc.setPlaceholderText("12345678")
        self.edit_doc.setStyleSheet("""
            QLineEdit {
                background-color: rgba(0, 0, 0, 150);
                color: white;
                border-radius: 12px;
                padding: 10px;
                font-size: 16px;
                border: 2px solid #333;
            }
            QLineEdit:focus {
                border-color: #4f90e2;
            }
        """)
        hdoc.addWidget(self.edit_doc, stretch=1)
        main_layout.addLayout(hdoc)

        # Espacio intermedio
        main_layout.addStretch(1)

        # Botón de enviar código con estilo moderno
        btn_send_code = QtWidgets.QPushButton("Enviar código")
        btn_send_code.setStyleSheet("""
            QPushButton {
                color: white;
                background-color: #339CFF;
                border-radius: 25px;
                font-size: 16px;
                border: none;
                padding: 15px 30px;
            }
            QPushButton:hover {
                background-color: #007BFF;
            }
            QPushButton:pressed {
                background-color: #005bb5;
            }
        """)
        btn_send_code.clicked.connect(self.enviar_codigo)
        main_layout.addWidget(btn_send_code, alignment=QtCore.Qt.AlignCenter)

        # Espacio final para separar el botón del borde inferior
        main_layout.addStretch(2)

    def enviar_codigo(self):
        num_doc = self.edit_doc.text().strip()

        if not num_doc:
            QtWidgets.QMessageBox.warning(self, "Datos faltantes", "Debe ingresar su documento.")
            return

        # Buscar el correo del usuario en la base de datos
        cursor = conn.cursor()
        cursor.execute("SELECT CORREO FROM USERS WHERE NUM_DOC = %s", (num_doc,))
        row = cursor.fetchone()
        cursor.close()

        if not row:
            QtWidgets.QMessageBox.warning(self, "Usuario no encontrado", "No se encontró un usuario con ese documento.")
            return

        email_destino = row[0]

        # Enviar el código
        codigo = enviar_codigo_por_email(email_destino)
        if codigo:
            self.codigo_recibido = codigo
            QtWidgets.QMessageBox.information(self, "Código Enviado", "Te hemos enviado un código por correo electrónico.")
            self.mostrar_ventana_codigo()
        else:
            QtWidgets.QMessageBox.critical(self, "Error", "Hubo un problema al enviar el correo.")

    def mostrar_ventana_codigo(self):
        # Crear una nueva ventana para ingresar el código
        self.codigo_window = QtWidgets.QWidget()
        self.codigo_window.setWindowTitle("Verificar Código")
        self.codigo_window.resize(400, 300)
        self.codigo_window.setStyleSheet("background: #121212; color: white;")

        layout = QtWidgets.QVBoxLayout(self.codigo_window)

        lbl_text = QtWidgets.QLabel("Ingresa el código de recuperación enviado a tu correo:")
        lbl_text.setStyleSheet("""
            color: white;
            font-size: 16px;
            font-weight: normal;
            margin-bottom: 20px;
        """)
        layout.addWidget(lbl_text)

        self.edit_codigo = QtWidgets.QLineEdit()
        self.edit_codigo.setPlaceholderText("Código")
        self.edit_codigo.setStyleSheet("""
            QLineEdit {
                background-color: rgba(0, 0, 0, 150);
                color: white;
                border-radius: 12px;
                padding: 10px;
                font-size: 16px;
                border: 2px solid #333;
            }
            QLineEdit:focus {
                border-color: #4f90e2;
            }
        """)
        layout.addWidget(self.edit_codigo)

        btn_verify = QtWidgets.QPushButton("Verificar Código")
        btn_verify.setStyleSheet("""
            QPushButton {
                background-color: #339CFF;
                color: white;
                border-radius: 25px;
                font-size: 16px;
                border: none;
                padding: 15px 30px;
            }
            QPushButton:hover {
                background-color: #339CFF;
            }
            QPushButton:pressed {
                background-color: #005bb5;
            }
        """)
        btn_verify.clicked.connect(self.verificar_codigo)
        layout.addWidget(btn_verify)

        self.codigo_window.show()

    def verificar_codigo(self):
        # Verificar que el código ingresado sea correcto
        codigo_ingresado = self.edit_codigo.text().strip()

        if codigo_ingresado == self.codigo_recibido:
            self.mostrar_ventana_cambio_contrasena()
        else:
            QtWidgets.QMessageBox.warning(self, "Código incorrecto", "El código ingresado es incorrecto.")

    def mostrar_ventana_cambio_contrasena(self):
        # Ventana para cambiar la contraseña
        self.cambio_window = QtWidgets.QWidget()
        self.cambio_window.setWindowTitle("Cambiar Contraseña")
        self.cambio_window.resize(400, 300)
        self.cambio_window.setStyleSheet("background: #121212; color: white;")

        layout = QtWidgets.QVBoxLayout(self.cambio_window)

        lbl_text = QtWidgets.QLabel("Ingresa tu nueva contraseña:")
        lbl_text.setStyleSheet("""
            color: white;
            font-size: 16px;
            font-weight: normal;
            margin-bottom: 20px;
        """)
        layout.addWidget(lbl_text)

        self.edit_new_pwd = QtWidgets.QLineEdit()
        self.edit_new_pwd.setEchoMode(QtWidgets.QLineEdit.Password)
        self.edit_new_pwd.setPlaceholderText("Nueva contraseña")
        self.edit_new_pwd.setStyleSheet("""
            QLineEdit {
                background-color: rgba(0, 0, 0, 150);
                color: white;
                border-radius: 12px;
                padding: 10px;
                font-size: 16px;
                border: 2px solid #333;
            }
            QLineEdit:focus {
                border-color: #4f90e2;
            }
        """)
        layout.addWidget(self.edit_new_pwd)

        btn_change_pwd = QtWidgets.QPushButton("Cambiar Contraseña")
        btn_change_pwd.setStyleSheet("""
            QPushButton {
                background-color: #339CFF;
                color: white;
                border-radius: 25px;
                font-size: 16px;
                border: none;
                padding: 15px 30px;
            }
            QPushButton:hover {
                background-color: #339CFF;
            }
            QPushButton:pressed {
                background-color: #005bb5;
            }
        """)
        btn_change_pwd.clicked.connect(self.cambiar_contrasena)
        layout.addWidget(btn_change_pwd)

        self.cambio_window.show()

    def cambiar_contrasena(self):
        # Cambiar la contraseña en la base de datos
        new_pwd = self.edit_new_pwd.text().strip()

        if not new_pwd:
            QtWidgets.QMessageBox.warning(self, "Contraseña vacía", "La contraseña no puede estar vacía.")
            return

        # Encriptar la nueva contraseña
        hashed_pwd = bcrypt.hashpw(new_pwd.encode('utf-8'), bcrypt.gensalt())

        # Aquí deberías guardar la contraseña en la base de datos (simulación)
        print(f"Contraseña cambiada: {hashed_pwd.decode('utf-8')}")
        QtWidgets.QMessageBox.information(self, "Contraseña cambiada", "Tu contraseña ha sido cambiada exitosamente.")
        self.cambio_window.destroy()
        self.destroy()


if __name__ == "__main__":
    check_for_update_and_exit_if_needed()  # 👈 Esto va PRIMERO. Si no está actualizado, se sale.
    app = QtWidgets.QApplication(sys.argv)
    w = LoginWindow()
    w.show()
    sys.exit(app.exec_())