import pandas as pd
from sqlalchemy import create_engine
import gspread
from google.oauth2.service_account import Credentials
import os

# --- 1. CONFIGURACIÓN ---
DB_URL = os.environ.get("DATABASE_URL")
GCP_SA_KEY = os.environ.get("GCP_SA_KEY")
SHEET_ID = os.environ.get("SHEET_ID")

HOJA_DATOS = "Consolidado FullStack"
NOMBRE_TABLA = "consolidado_fullstack"
COLUMNA_FECHA = "FECHA"

# Validar que todas las variables de entorno se cargaron
if not all([DB_URL, GCP_SA_KEY, SHEET_ID]):
    print("ERROR: Faltan una o más variables de entorno (DATABASE_URL, GCP_SA_KEY, SHEET_ID).")
    exit(1)

print("Iniciando migración de datos desde Google Sheets a Aiven...")

try:
    # --- 2. CONECTARSE A GOOGLE SHEETS ---
    print("Autenticando con Google...")
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    
    with open("google_credentials.json", "w") as f:
        f.write(GCP_SA_KEY)
        
    creds = Credentials.from_service_account_file("google_credentials.json", scopes=scopes)
    client = gspread.authorize(creds)
    
    print(f"Leyendo Google Sheet ID: {SHEET_ID}")
    spreadsheet = client.open_by_key(SHEET_ID)
    worksheet = spreadsheet.worksheet(HOJA_DATOS)
    
    values = worksheet.get_all_values()
    os.remove("google_credentials.json")
    
    if not values:
        print("ERROR: La hoja de cálculo está vacía.")
        exit(1)
        
    df = pd.DataFrame(values[1:], columns=values[0])
    print(f"Se encontraron {len(values) - 1} filas de datos (más 1 fila de encabezado).")

    # --- 3. LIMPIAR Y FILTRAR LOS DATOS ---
    print("Limpiando y filtrando los datos...")
    
    df.columns = [
        str(col).replace(' ', '_').replace('á', 'a').replace('é', 'e').replace('í', 'i')
           .replace('ó', 'o').replace('ú', 'u').replace('ñ', 'n').upper()
        for col in df.columns
    ]

    df[COLUMNA_FECHA] = pd.to_datetime(df[COLUMNA_FECHA], dayfirst=True, errors='coerce')
    df.dropna(subset=[COLUMNA_FECHA], inplace=True)
    
    # df = df[df[COLUMNA_FECHA].dt.month >= 8] # Filtro de mes deshabilitado
    
    print(f"Se han limpiado los datos. Se cargarán {len(df)} filas.")

    # --- 4. CONECTARSE A LA BASE DE DATOS (CON TIMEOUT y SSL) ---
    engine = create_engine(
        DB_URL,
        connect_args={
            'connect_timeout': 60,
            'ssl_mode': 'REQUIRED'
        }
    )

    # --- 5. INSERTAR DATOS ---
    print(f"Conectando a la base de datos y cargando datos en la tabla '{NOMBRE_TABLA}'...")
    df.to_sql(
        name=NOMBRE_TABLA,
        con=engine,
        if_exists='replace',
        index=True,         # <-- CAMBIO 1: De False a True
        index_label='id'    # <-- CAMBIO 2: Añadir esta línea
    )
    print(f"¡Migración a la base de datos completada! Se han insertado {len(df)} filas.")

except Exception as e:
    print(f"--- OCURRIÓ UN ERROR DURANTE LA MIGRACIÓN ---")
    print(f"Error: {e}")
    if os.path.exists("google_credentials.json"):
        os.remove("google_credentials.json")
    exit(1)