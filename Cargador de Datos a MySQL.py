import pandas as pd
from sqlalchemy import create_engine, text, inspect
from sqlalchemy.exc import SQLAlchemyError

# --- CONFIGURACIÓN DE LA BASE DE DATOS MYSQL ---
# Datos que proporcionaste
USUARIO = "root"
CONTRASENA = "fran1080"
HOST = "127.0.0.1"
PUERTO = "3306"
BASE_DE_DATOS = "Dashboard_KPI"
NOMBRE_TABLA = "CONSOLIDADO_FULLSTACK"

# --- CONFIGURACIÓN DEL ARCHIVO EXCEL ---
RUTA_ARCHIVO = r"C:\Users\Haintech\Desktop\Consolidado_Ordenes_PowerQuery\FullStack_Consolidado.xlsx"
HOJA_DATOS = "Consolidado FullStack"

print("Iniciando el proceso de carga de datos a MySQL...")

try:
    # --- 1. LEER DATOS DEL ARCHIVO EXCEL ---
    print(f"Leyendo el archivo Excel desde: {RUTA_ARCHIVO}")
    df = pd.read_excel(RUTA_ARCHIVO, sheet_name=HOJA_DATOS)
    print(f"Se han leído {len(df)} filas del archivo Excel.")

    # --- Limpieza de nombres de columna para compatibilidad con SQL ---
    df.columns = [col.replace(' ', '_').replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u').replace('ñ', 'n').upper() for col in df.columns]
    print("Nombres de columnas normalizados para la base de datos.")

    # --- 2. CONECTARSE A LA BASE DE DATOS ---
    print(f"Conectando a la base de datos MySQL en {HOST}...")
    cadena_conexion = f"mysql+pymysql://{USUARIO}:{CONTRASENA}@{HOST}:{PUERTO}/{BASE_DE_DATOS}"
    engine = create_engine(cadena_conexion)
    
    with engine.connect() as connection:
        print("¡Conexión a MySQL exitosa!")
        
        # --- 3. INSERTAR LOS DATOS EN MYSQL ---
        print(f"Creando/Reemplazando la tabla '{NOMBRE_TABLA}' e insertando los datos...")
        df.to_sql(
            name=NOMBRE_TABLA.lower(), # MySQL prefiere nombres de tabla en minúsculas
            con=engine,
            if_exists='replace',
            index=False,
            chunksize=1000,
        )
        print(f"¡Proceso completado! Se han insertado {len(df)} filas en la tabla '{NOMBRE_TABLA}'.")

except SQLAlchemyError as e:
    print("\n--- ERROR DE BASE DE DATOS ---")
    print(f"No se pudo conectar o escribir en la base de datos MySQL: {e}")
    print("\nPosibles soluciones:")
    print("1. Verifica que el servicio de MySQL (ej. desde XAMPP) esté corriendo.")
    print("2. Asegúrate de que las credenciales (USUARIO, CONTRASENA, HOST, etc.) sean correctas.")
    print("3. Comprueba que la base de datos 'Dashboard_KPI' haya sido creada.")

except FileNotFoundError:
    print(f"\n--- ERROR DE ARCHIVO ---")
    print(f"No se pudo encontrar el archivo Excel en la ruta: {RUTA_ARCHIVO}")

except Exception as e:
    print(f"\n--- OCURRIÓ UN ERROR INESPERADO ---")
    print(f"Error: {e}")
