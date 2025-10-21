# 📊 Dashboard KPI Fullstack

Sistema automatizado para visualizar KPIs actualizados desde una hoja de Google Sheets.  
Los datos se sincronizan cada **15 minutos** y se muestran en un **dashboard interactivo** alojado en la nube.

✅ **Tu única tarea**: actualizar los datos en Google Sheets.  
🔁 **El resto es 100% automático**.

-------------------------------------------------------------------------------------------------------------------------------------------------

## 🔄 Arquitectura del Sistema

El sistema se compone de **dos flujos independientes pero integrados**:  
1. **Pipeline de datos automático** (cada 15 minutos).  
2. **Dashboard interactivo** (cuando un usuario visita la URL).


```plaintext
                                 ┌──────────────────┐     ┌──────────────────┐
                                 │                  │     │                  │
                                 │   TÚ (ACTUALIZAS │     │   USUARIO FINAL  │
                                 │   DATOS EN       │     │   (VISITA URL)   │
                                 │   GOOGLE SHEETS) │     │                  │
                                 └─────────┬────────┘     └─────────┬────────┘
                                           │                        │
                                           │                        │
                                           ▼                        ▼
                                 ┌─────────┴────────┐     ┌─────────┴────────┐
                                 │                  │     │                  │
                                 │   GOOGLE SHEETS  │◄────┤    RENDER        │
                                 │   (Fuente de     │     │    (App Dash)    │
                                 │    Datos Maestra)│     │                  │
                                 └─────────┬────────┘     └─────────▲────────┘
                                           │                        │
                                           │ (API vía SA)           │ (Consulta DB)
                                           ▼                        │
                                 ┌─────────┴────────┐               │
                                 │                  │               │
                                 │   GITHUB ACTIONS │───────────────┘
                                 │   (ETL Automático│
                                 │    cada 15 min)  │
                                 └─────────┬────────┘
                                           │
                                           │ (Escribe en DB)
                                           ▼
                                 ┌─────────┴────────────────────┐
                                 │                              │
                                 │   RAILWAY                    │
                                 │   (MySQL)                    │
                                 │   • Tabla:                   │
                                 │     consolidado_fullstack    │
                                 └──────────────────────────────┘
```


> 🔑 **Leyenda**:  
> - **Flechas**: dirección del flujo de datos.  
> - **GitHub Actions**: proceso programado (no interactivo).  
> - **Render**: servicio que aloja la app Dash.  
> - **Railway**: base de datos MySQL gestionada.

-------------------------------------------------------------------------------------------------------------------------------------------------


### 🤖 Flujo 1: Pipeline de Datos (ETL Automático)

Este flujo se ejecuta **sin intervención humana** cada 15 minutos:

1. **Disparo**: GitHub Actions se activa mediante un cron job (`*/15 * * * *`).
2. **Lectura**: Usa las credenciales seguras (`GCP_SA_KEY` y `SHEET_ID`) para leer datos desde Google Sheets vía API.
3. **Carga**: El script [`migrar_datos.py`](./migrar_datos.py) conecta a la base de datos en Railway (usando `DATABASE_URL`) y **reemplaza completamente** la tabla `consolidado_fullstack`.
4. **Resultado**: La base de datos siempre refleja el estado más reciente de tu hoja maestra.

✅ **Tu única tarea manual**: actualizar los datos en Google Sheets.

-------------------------------------------------------------------------------------------------------------------------------------------------

### 🖥️ Flujo 2: Dashboard Web (Interacción del Usuario)

Cuando un usuario accede a [https://dashboard-kpi-fullstack.onrender.com](https://dashboard-kpi-fullstack.onrender.com):

1. **Autenticación**: Render solicita credenciales básicas (`dash_auth`).
2. **Carga de datos**: La app Dash (`dashboard_kpi_DB.py`) se conecta a Railway usando la variable de entorno `DATABASE_URL`.
3. **Visualización**: Los datos se leen de la tabla `consolidado_fullstack` y se renderizan como gráficos interactivos en el navegador.

💡 **Ventaja clave**: el dashboard **nunca consulta Google Sheets directamente**, lo que mejora rendimiento, seguridad y confiabilidad.

-------------------------------------------------------------------------------------------------------------------------------------------------

### ✅ Beneficios del Diseño

- **Automatización total** tras la actualización manual en Sheets.
- **Separación de capas**: fuente de datos, procesamiento, almacenamiento y presentación.
- **Seguridad**: credenciales sensibles nunca están en el código (usadas como secretos en GitHub/Render).
- **Escalable**: fácil de adaptar a otras fuentes de datos o servicios de hosting.

-------------------------------------------------------------------------------------------------------------------------------------------------

### 🛠️ Tecnologías Utilizadas

| Capa               | Tecnología                |
|--------------------|---------------------------|
| Fuente de datos    | Google Sheets             |
| Orquestación ETL   | GitHub Actions            |
| Base de datos      | Railway (MySQL)           |
| Aplicación web     | Plotly Dash en Render     |
| Autenticación      | `dash_auth` (HTTP Basic)  |

-------------------------------------------------------------------------------------------------------------------------------------------------

### 📁 Estructura del Proyecto
├── migrar_datos.py # Script ETL: Sheets → Railway
├── dashboard_kpi_DB.py # App Dash: Railway → Render
├── requirements.txt # Dependencias de Python
└── README.md # Este archivo


-------------------------------------------------------------------------------------------------------------------------------------------------

### 🔐 Variables de Entorno y Secretos

| Entorno       | Variable         | Descripción                                  |
|---------------|------------------|----------------------------------------------|
| GitHub Actions| `GCP_SA_KEY`     | Credenciales de Service Account de Google    |
|               | `SHEET_ID`       | ID de la hoja de Google Sheets               |
|               | `DATABASE_URL`   | URL de conexión a la base de datos en Railway|
| Render        | `DATABASE_URL`   | Misma URL, usada por la app Dash             |
|               | `DASH_USERNAME`  | Usuario para autenticación básica            |
|               | `DASH_PASSWORD`  | Contraseña para autenticación básica         |


-------------------------------------------------------------------------------------------------------------------------------------------------

> 📌 **Nota**: Este sistema está diseñado para **simplicidad, mantenibilidad y bajo costo**. Ideal para KPIs operativos actualizados periódicamente.

-------------------------------------------------------------------------------------------------------------------------------------------------

