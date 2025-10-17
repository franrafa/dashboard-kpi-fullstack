import pandas as pd
from sqlalchemy import create_engine
import os # Importar os para leer variables de entorno

# --- CONFIGURACIÓN CON VARIABLES DE ENTORNO ---
HOST = os.environ.get("HOST")
USUARIO = os.environ.get("USUARIO")
CONTRASENA = os.environ.get("CONTRASENA")
PUERTO = os.environ.get("PUERTO")
BASE_DE_DATOS = os.environ.get("BASE_DE_DATOS")

# --- CONFIGURACIÓN DEL PROYECTO ---
RUTA_ARCHIVO = "FullStack_Consolidado.xlsx"
HOJA_DATOS = "Consolidado FullStack"
NOMBRE_TABLA = "consolidado_fullstack"

# Validar que todas las variables de entorno se cargaron
if not all([HOST, USUARIO, CONTRASENA, PUERTO, BASE_DE_DATOS]):
    print("ERROR: Faltan una o más variables de entorno (HOST, USUARIO, CONTRASENA, PUERTO, BASE_DE_DATOS).")
    exit(1)

print("Iniciando migración de datos a Railway...")

try:
    # --- 1. LEER DATOS DEL EXCEL ---
    print(f"Leyendo el archivo Excel: {RUTA_ARCHIVO}")
    df = pd.read_excel(RUTA_ARCHIVO, sheet_name=HOJA_DATOS)
    df.columns = [
        str(col).replace(' ', '_').replace('á', 'a').replace('é', 'e').replace('í', 'i')
           .replace('ó', 'o').replace('ú', 'u').replace('ñ', 'n').upper()
        for col in df.columns
    ]
    print(f"Se han leído {len(df)} filas del Excel.")

    # --- 2. CONECTARSE A RAILWAY ---
    cadena_conexion = f"mysql+pymysql://{USUARIO}:{CONTRASENA}@{HOST}:{PUERTO}/{BASE_DE_DATOS}"
    engine = create_engine(cadena_conexion)

    # --- 3. INSERTAR DATOS ---
    print(f"Conectando a Railway y cargando datos en la tabla '{NOMBRE_TABLA}'...")
    df.to_sql(
        name=NOMBRE_TABLA,
        con=engine,
        if_exists='replace',
        index=False,
        chunksize=1000
    )
    print(f"¡Migración a Railway completada! Se han insertado {len(df)} filas.")

except Exception as e:
    print(f"--- OCURRIÓ UN ERROR DURANTE LA MIGRACIÓN ---")
    print(f"Error: {e}")
    exit(1)