# Dashboard KPI FullStack Interactivo

Este es un dashboard interactivo de KPIs (Key Performance Indicators) construido con Dash y Plotly. La aplicación visualiza datos de gestión de órdenes, permitiendo a los usuarios filtrar y analizar el rendimiento por torre, ejecutivo, mes y semana.

El sistema está completamente automatizado y alojado en la nube.

**Ver el Dashboard en Vivo:** [https://dashboard-kpi-fullstack.onrender.com](https://dashboard-kpi-fullstack.onrender.com)

---

## 🚀 Características Principales

* **Autenticación de Usuarios:** Acceso protegido por nombre de usuario y contraseña.
* **Tarjetas de KPIs Dinámicas:** Métricas clave (Gestiones Totales, Tasa de Resolutividad, Gestión FTE Día) que se actualizan con los filtros.
* **Filtros Interactivos:** Filtra datos por Mes, Quincena/Semana, Torre y Ejecutivo.
* **Tablas Detalladas:** Vistas de resumen por Mes, Día, Torre y Status.
* **Gráficos Interactivos:** Gráficos de pastel y barras para visualizar la distribución del trabajo.
* **Ranking de Ejecutivos Clave:** Una pestaña dedicada para comparar el rendimiento de los ejecutivos (Resolutividad y Volumen).
* **Descargas Múltiples:**
    * Descarga de los datos detallados (Consolidado) con filtros aplicados.
    * Descarga de los rankings KPI en un archivo Excel con múltiples hojas.

---

## ⚙️ Arquitectura y Flujo de Datos Automatizado

Este proyecto utiliza una arquitectura de CI/CD (Integración Continua/Despliegue Continuo) para la automatización completa del flujo de datos, desde la fuente hasta la visualización.

1.  **Fuente de Datos (Google Sheets):** Los datos maestros se actualizan y mantienen en una **Hoja de Cálculo de Google** privada.
2.  **ETL Automatizado (GitHub Actions):** Un "robot" (flujo de trabajo de GitHub Actions) se despierta **cada 15 minutos**.
3.  **Extracción y Carga:** El robot ejecuta el script `migrar_datos.py`, se conecta de forma segura a la API de Google Sheets, extrae los datos, los limpia y los carga en la base de datos de producción.
4.  **Base de Datos (Railway):** Una base de datos **MySQL** alojada en Railway sirve como el almacén de datos (Data Warehouse) para la aplicación.
5.  **Aplicación Web (Render):** La aplicación Dash (`dashboard_kpi_DB.py`) está alojada en Render. **No lee archivos locales**.
6.  **Visualización:** Cuando un usuario carga el dashboard, la aplicación en Render consulta la base de datos de Railway en tiempo real para mostrar los datos más frescos.
7.  **Auto-Actualización:** El dashboard también incluye un `dcc.Interval` que vuelve a consultar la base de datos cada minuto, asegurando que los nuevos datos cargados por el robot de GitHub se reflejen sin necesidad de recargar la página.

### Tu Nuevo Flujo de Trabajo
* **Para Actualizar DATOS:** Simplemente edita la **Hoja de Cálculo de Google**. El sistema se actualizará solo.
* **Para Actualizar CÓDIGO:** Sube los cambios de los archivos `.py` a GitHub.

---

## 🛠️ Stack Tecnológico

* **Python**
* **Framework Web:** Dash, Plotly
* **Análisis de Datos:** Pandas
* **Base de Datos:** MySQL (alojada en Railway)
* **Fuente de Datos:** Google Sheets API
* **Hosting de Aplicación:** Render
* **Automatización (CI/CD):** GitHub Actions
* **Servidor de Producción:** Gunicorn


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


------------------------------------------------------------------------------------------------------------------------------------------------------

> 📌 **Nota**: Este sistema está diseñado para **simplicidad, mantenibilidad y bajo costo**. Ideal para KPIs operativos actualizados periódicamente.

------------------------------------------------------------------------------------------------------------------------------------------------------

