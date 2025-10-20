import pandas as pd
from sqlalchemy import create_engine
import gspread
from google.oauth2.service_account import Credentials
import os

# --- 1. CONFIGURACIÓN ---
# Lee las credenciales de los Secretos de GitHub
DB_URL = os.environ.get("DATABASE_URL")
GCP_SA_KEY = os.environ.get("GCP_SA_KEY")
SHEET_ID = os.environ.get("SHEET_ID")

HOJA_DATOS = "Consolidado FullStack" # Nombre de la pestaña en tu Google Sheet
NOMBRE_TABLA = "consolidado_fullstack"

# Validar que todas las variables de entorno se cargaron
if not all([DB_URL, GCP_SA_KEY, SHEET_ID]):
    print("ERROR: Faltan una o más variables de entorno (DATABASE_URL, GCP_SA_KEY, SHEET_ID).")
    exit(1)

print("Iniciando migración de datos desde Google Sheets a Railway...")

try:
    # --- 2. CONECTARSE A GOOGLE SHEETS ---
    print("Autenticando con Google...")
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    
    # Escribir el contenido del secreto JSON a un archivo temporal
    with open("google_credentials.json", "w") as f:
        f.write(GCP_SA_KEY)
        
    creds = Credentials.from_service_account_file("google_credentials.json", scopes=scopes)
    client = gspread.authorize(creds)
    
    print(f"Leyendo Google Sheet ID: {SHEET_ID}")
    spreadsheet = client.open_by_key(SHEET_ID)
    worksheet = spreadsheet.worksheet(HOJA_DATOS)
    
    records = worksheet.get_all_records()
    df = pd.DataFrame(records)
    
    # Limpiar el archivo de credenciales temporal
    os.remove("google_credentials.json")

    df.columns = [
        str(col).replace(' ', '_').replace('á', 'a').replace('é', 'e').replace('í', 'i')
           .replace('ó', 'o').replace('ú', 'u').replace('ñ', 'n').upper()
        for col in df.columns
    ]
    print(f"Se han leído {len(df)} filas del Excel.")

    # --- 3. CONECTARSE A RAILWAY ---
    engine = create_engine(DB_URL)

    # --- 4. INSERTAR DATOS ---
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
    # Limpiar el archivo de credenciales si falla
    if os.path.exists("google_credentials.json"):
        os.remove("google_credentials.json")
    exit(1)