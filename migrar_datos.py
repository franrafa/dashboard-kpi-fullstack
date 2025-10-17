import pandas as pd
from sqlalchemy import create_engine
import os # Importar os para leer variables de entorno

# --- CONFIGURACIÓN CON VARIABLES DE ENTORNO ---
cadena_conexion = os.environ.get("DATABASE_URL")

# --- CONFIGURACIÓN DEL PROYECTO ---
RUTA_ARCHIVO = "FullStack_Consolidado.xlsx"
HOJA_DATOS = "Consolidado FullStack"
NOMBRE_TABLA = "consolidado_fullstack"

# Validar que todas las variables de entorno se cargaron
if not cadena_conexion:
    print("ERROR: Faltan la variable de entorno DATABASE_URL.")
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