import os, sys
import pytds
from dotenv import load_dotenv

def obtener_ruta_recurso(nombre_archivo):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, nombre_archivo)
    return nombre_archivo

# Carga tu .env
load_dotenv(dotenv_path=obtener_ruta_recurso('.env'))

def conectar_sql_server(env_database_var):
    # Separa servidor y puerto si vienen juntos
    raw = os.getenv('DB_SERVER', '')
    if ',' in raw:
        server, port_str = raw.split(',', 1)
        try:
            port = int(port_str)
        except ValueError:
            port = 1433
    else:
        server = raw
        port   = 1433

    database = os.getenv(env_database_var)
    user     = os.getenv('DB_USERNAME')
    password = os.getenv('DB_PASSWORD')

    if not all([server, database, user, password]):
        print("Faltan variables de entorno para la conexión.")
        return None

    try:
        conn = pytds.connect(
            server=server,
            port=port,
            database=database,
            user=user,
            password=password,
            autocommit=True
        )
        return conn
    except Exception as e:
        print("Error al conectar con pytds:", e)
        return None
