import pandas as pd
import plotly.express as px
import dash
import dash_auth
import dash_bootstrap_components as dbc
from dash import html, dcc, dash_table, Input, Output, callback, State
from dash.exceptions import PreventUpdate
import locale
from datetime import datetime
import io
import os
from sqlalchemy import create_engine
from sqlalchemy.exc import SQLAlchemyError
import traceback

# --- 1. CONFIGURACIÓN GENERAL ---
USUARIO = os.environ.get("root")
CONTRASENA = os.environ.get("nhdHlnglnSsIekEJJDAbykayPacNIBhB")
HOST = os.environ.get("shortline.proxy.rlwy.net")
PUERTO = os.environ.get("54379")
BASE_DE_DATOS = os.environ.get("railway")
NOMBRE_TABLA = "consolidado_fullstack"
RUTA_ARCHIVO = "FullStack_Consolidado.xlsx"
HOJA_DATOS = "Consolidado FullStack"
try:
    # Intenta la configuración ideal para sistemas Linux
    locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
except locale.Error:
    try:
        # Si falla, intenta la configuración para Windows
        locale.setlocale(locale.LC_TIME, 'Spanish')
    except locale.Error:
        # Si todo falla, imprime una advertencia y continúa con el locale por defecto (inglés)
        print("Advertencia: No se pudo establecer el locale a español. Los meses se mostrarán en inglés.")
        pass # La aplicación NO se detendrá
COLUMNA_FECHA = "FECHA"
COLUMNA_ANALISTA = "EJECUTIVO"
COLUMNA_ORDEN = "NUMERO_DE_PEDIDO"
COLUMNA_STATUS = "STATUS_REAL"
COLUMNA_TORRE = "TORRE"
VALID_USERNAME_PASSWORD_PAIRS = {'haintech': 'dashboard2025'}


# --- 2. FUNCIÓN CENTRAL DE CARGA Y PROCESAMIENTO DE DATOS ---
def actualizar_y_cargar_datos_desde_excel():
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Iniciando proceso de actualización...")
    df_excel = pd.read_excel(RUTA_ARCHIVO, sheet_name=HOJA_DATOS)
    df_excel.columns = [
        col.replace(' ', '_').replace('á', 'a').replace('é', 'e').replace('í', 'i')
           .replace('ó', 'o').replace('ú', 'u').replace('ñ', 'n').upper()
        for col in df_excel.columns
    ]
    cadena_conexion = f"mysql+pymysql://{USUARIO}:{CONTRASENA}@{HOST}:{PUERTO}/{BASE_DE_DATOS}"
    engine = create_engine(cadena_conexion)
    with engine.connect() as connection:
        df_excel.to_sql(name=NOMBRE_TABLA, con=connection, if_exists='replace', index=False, chunksize=1000)
    with engine.connect() as connection:
        df_dashboard = pd.read_sql_table(NOMBRE_TABLA, connection)
    df_dashboard[COLUMNA_FECHA] = pd.to_datetime(df_dashboard[COLUMNA_FECHA], errors='coerce')
    df_dashboard.dropna(subset=[COLUMNA_FECHA, COLUMNA_ANALISTA, COLUMNA_TORRE, COLUMNA_STATUS], inplace=True)
    df_dashboard = df_dashboard[df_dashboard[COLUMNA_FECHA].dt.month >= 8]
    df_dashboard.sort_values(by=COLUMNA_FECHA, inplace=True)
    df_dashboard['Mes'] = df_dashboard[COLUMNA_FECHA].dt.strftime('%B').str.capitalize()
    df_dashboard['Year'] = df_dashboard[COLUMNA_FECHA].dt.isocalendar().year
    df_dashboard['Semana_Num'] = df_dashboard[COLUMNA_FECHA].dt.isocalendar().week
    df_dashboard['WeekStartDate'] = pd.to_datetime(df_dashboard['Year'].astype(str) + df_dashboard['Semana_Num'].astype(str) + '1', format='%G%V%u')
    df_dashboard['WeekEndDate'] = df_dashboard['WeekStartDate'] + pd.to_timedelta('6 days')
    df_dashboard['WeekLabel'] = "Semana " + df_dashboard['Semana_Num'].astype(str) + " (" + df_dashboard['WeekStartDate'].dt.strftime('%d %b') + " - " + df_dashboard['WeekEndDate'].dt.strftime('%d %b') + ")"
    print("Procesamiento de datos para el dashboard completado.")
    return df_dashboard


# --- 3. INICIALIZACIÓN DE LA APLICACIÓN DASH ---
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.LUX], suppress_callback_exceptions=True)
server = app.server
server.secret_key = os.urandom(24)
auth = dash_auth.BasicAuth(app, VALID_USERNAME_PASSWORD_PAIRS)

# --- Carga inicial de datos ---
try:
    df_principal = actualizar_y_cargar_datos_desde_excel()
    meses_disponibles = sorted(df_principal['Mes'].unique(), key=lambda m: pd.to_datetime(f'01-{m}-2025', format='%d-%B-%Y').month)
    week_map = df_principal[['Semana_Num', 'WeekLabel']].drop_duplicates().sort_values('Semana_Num')
    semanas_disponibles_options = week_map.apply(lambda row: {'label': row['WeekLabel'], 'value': row['Semana_Num']}, axis=1).tolist()
    ejecutivos_disponibles = sorted(df_principal[COLUMNA_ANALISTA].unique())
    torres_disponibles = sorted(df_principal[COLUMNA_TORRE].unique())
    datos_cargados_correctamente = True
    initial_load_time_str = f"Datos cargados desde Excel y DB a las: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
except (SQLAlchemyError, FileNotFoundError, Exception) as e:
    error_mensaje = f"Ocurrió un error crítico durante la carga inicial de datos: {e}"
    datos_cargados_correctamente = False
    df_principal = pd.DataFrame()
    initial_load_time_str = "Error al cargar datos."
    traceback.print_exc()


# --- 4. DISEÑO DE LA APLICACIÓN WEB (LAYOUT) ---
if datos_cargados_correctamente:
    app.layout = dbc.Container([
        dcc.Store(id='store-main-data', data=df_principal.to_json(date_format='iso', orient='split')),
        dcc.Interval(id='interval-component', interval=60 * 1000, n_intervals=0),
        dcc.Download(id="download-excel"),
        dcc.Store(id='store-resumen-conteo-data'),
        dcc.Store(id='store-resumen-porcentaje-data'),
        dcc.Store(id='store-download-raw-data'),
        
        dbc.Row(dbc.Col(html.H1("Dashboard Consolidado FullStack", className="text-center text-info my-4"))),
        dbc.Card(dbc.CardBody([
             dbc.Row([
                dbc.Col(dcc.Dropdown(id='filtro-mes', options=meses_disponibles, placeholder="Seleccionar Mes(es)", multi=True), md=3),
                dbc.Col([
                    html.Label("Filtrar por:", style={'fontWeight': 'bold'}),
                    dcc.RadioItems(id='modo-filtro-tiempo', options=[{'label': ' Quincena', 'value': 'quincena'}, {'label': ' Semana', 'value': 'semana'}], value='quincena', inline=True, labelStyle={'margin-right': '10px'}),
                    html.Div(id='contenedor-filtro-quincena', children=[dcc.Dropdown(id='filtro-quincena', options=[{'label': '1ra Quincena', 'value': 1}, {'label': '2da Quincena', 'value': 2}], placeholder="Seleccionar Quincena", className="mt-1")]),
                    html.Div(id='contenedor-filtro-semana', children=[dcc.Dropdown(id='filtro-semana', options=semanas_disponibles_options, placeholder="Seleccionar Semana(s)", multi=True, className="mt-1")], style={'display': 'none'})
                ], md=3),
                dbc.Col(dcc.Dropdown(id='filtro-torre', options=torres_disponibles, placeholder="Seleccionar Torre(s)", multi=True), md=3),
                dbc.Col(dcc.Dropdown(id='filtro-ejecutivo', options=ejecutivos_disponibles, placeholder="Seleccionar Ejecutivo(s)", multi=True), md=3),
            ]),
            dbc.Row(dbc.Col(dbc.Button("Limpiar Filtros", id="btn-limpiar", color="dark", outline=True, className="w-100 mt-3"), width=12))
        ]), className="mb-4 shadow"),
        
        dbc.Tabs([
            dbc.Tab(label="Resumen Mensual", children=[dbc.Row(id='tarjetas-kpi-mensual', className="my-4"), dbc.Row([dbc.Col([html.H4("Resumen Mensual por Torre y Ejecutivo", className="border-bottom pb-2 mb-3"), dash_table.DataTable(id='tabla-resumen-mensual', style_header={'backgroundColor': '#E0E6F8', 'fontWeight': 'bold'}, style_cell={'textAlign': 'center'}, style_data_conditional=[{'if': {'filter_query': '{Tipo} = "Torre"'}, 'backgroundColor': '#E0E6F8', 'fontWeight': 'bold'},{'if': {'column_id': 'Etiquetas de Fila'}, 'textAlign': 'left', 'fontWeight': 'bold'},{'if': {'column_id': 'Total General'}, 'fontWeight': 'bold'}])], width=12)], className="mb-4")]),
            dbc.Tab(label="Detalle Diario", children=[
                dbc.Row(id='tarjetas-kpi-diario', className="my-4"), 
                dbc.Row([dbc.Col([html.H4("Resumen Diario por Torre", className="border-bottom pb-2 my-3"), dash_table.DataTable(
                    id='tabla-resumen-torre', 
                    style_table={'overflowX': 'auto'}, 
                    style_header={'backgroundColor': '#e8f5e9'}, 
                    style_cell={'textAlign': 'center', 'minWidth': '120px'}, 
                    style_cell_conditional=[
                        {'if': {'column_id': COLUMNA_TORRE}, 'textAlign': 'left', 'fontWeight': 'bold', 'minWidth': '180px'}, 
                        {'if': {'column_id': 'Total General'}, 'fontWeight': 'bold', 'backgroundColor': '#e8f5e9'}
                    ],
                    style_data_conditional=[{
                        'if': {'filter_query': f'{{{COLUMNA_TORRE}}} = "Total General"'},
                        'backgroundColor': '#e8f5e9',
                        'fontWeight': 'bold'
                    }]
                )], width=12)], className="mb-4"), 
                dbc.Row([dbc.Col([html.H4("Resumen Diario por Status", className="border-bottom pb-2 mb-3"), dash_table.DataTable(
                    id='tabla-resumen-status', 
                    style_table={'overflowX': 'auto'}, 
                    style_header={'backgroundColor': '#fff3e0'}, 
                    style_cell={'textAlign': 'center', 'minWidth': '120px'}, 
                    style_cell_conditional=[
                        {'if': {'column_id': COLUMNA_STATUS}, 'textAlign': 'left', 'fontWeight': 'bold', 'minWidth': '180px'}, 
                        {'if': {'column_id': 'Total General'}, 'fontWeight': 'bold', 'backgroundColor': '#fff3e0'}
                    ],
                    style_data_conditional=[{
                        'if': {'filter_query': f'{{{COLUMNA_STATUS}}} = "Total General"'},
                        'backgroundColor': '#fff3e0',
                        'fontWeight': 'bold'
                    }]
                )], width=12)], className="mb-4"), 
                dbc.Row([dbc.Col([html.H4("Resumen Diario por Ejecutivo (Cantidad)", className="border-bottom pb-2 mb-3"), dash_table.DataTable(
                    id='tabla-resumen-ejecutivo-conteo', 
                    style_table={'overflowX': 'auto'}, 
                    style_header={'backgroundColor': '#f2e3fd'}, 
                    style_cell={'textAlign': 'center', 'minWidth': '120px'}, 
                    style_cell_conditional=[
                        {'if': {'column_id': COLUMNA_ANALISTA}, 'textAlign': 'left', 'fontWeight': 'bold', 'minWidth': '180px'}, 
                        {'if': {'column_id': 'Total General'}, 'fontWeight': 'bold', 'backgroundColor': '#f2e3fd'}
                    ],
                    style_data_conditional=[{
                        'if': {'filter_query': f'{{{COLUMNA_ANALISTA}}} = "Total General"'},
                        'backgroundColor': '#f2e3fd',
                        'fontWeight': 'bold'
                    }]
                )], width=12)], className="mb-4"), 
                dbc.Row([dbc.Col([html.H4("Porcentaje de Resolutividad Diario por Ejecutivo", className="border-bottom pb-2 mb-3"), dash_table.DataTable(id='tabla-resumen-ejecutivo-porcentaje', style_table={'overflowX': 'auto'}, style_header={'backgroundColor': '#e3f2fd'}, style_cell={'textAlign': 'center', 'minWidth': '120px'}, style_cell_conditional=[{'if': {'column_id': COLUMNA_ANALISTA}, 'textAlign': 'left', 'fontWeight': 'bold', 'minWidth': '180px'}, {'if': {'column_id': 'Total General'}, 'fontWeight': 'bold', 'backgroundColor': '#e3f2fd'}])], width=12)], className="mb-4")
            ]),
            dbc.Tab(label="Gráficos", children=[dbc.Row(id='tarjetas-kpi-graficos', className="my-4"), dbc.Row([dbc.Col(dcc.Graph(id='grafico-torta-torre'), md=6), dbc.Col(dcc.Graph(id='grafico-barras-resolutividad'), md=6)], className="my-4"), dbc.Row([dbc.Col(dcc.Graph(id='grafico-volumen-ejecutivo'), md=6), dbc.Col(dcc.Graph(id='grafico-composicion-status'), md=6)], className="my-4")]),
            dbc.Tab(label="Descargar", children=[
                dbc.Row([
                    dbc.Col([
                        html.H4("Panel de Descarga", className="mt-4"),
                        html.P("Usa los filtros principales del dashboard y el selector de fechas de abajo para definir los datos a descargar."),
                        dcc.DatePickerRange(
                            id='download-date-picker',
                            min_date_allowed=df_principal[COLUMNA_FECHA].min().date(),
                            max_date_allowed=df_principal[COLUMNA_FECHA].max().date(),
                            start_date=df_principal[COLUMNA_FECHA].min().date(),
                            end_date=df_principal[COLUMNA_FECHA].max().date(),
                            display_format='DD/MM/YYYY'
                        ),
                        dbc.Button("Generar Archivo para Descarga", id="btn-generate-download", color="primary", className="mt-3"),
                        html.Div(id="download-preview-container", className="mt-4"),
                        dbc.Button(
                            "Descargar Archivo Completo (3 Hojas) como XLSX", 
                            id="btn-download-all", 
                            color="success", 
                            className="mt-3", 
                            disabled=True 
                        )
                    ], className="text-center", md={'size': 8, 'offset': 2})
                ], className="my-4")
            ])
        ], className="mt-4"),
        html.Div(id='last-updated-text', children=[initial_load_time_str], style={'textAlign': 'right', 'color': 'grey', 'marginTop': '20px'})
    ], fluid=True)
else:
    app.layout = dbc.Container([
        dbc.Alert(error_mensaje, color="danger", className="mt-4")
    ])

# --- 5. LÓGICA DE INTERACTIVIDAD (CALLBACKS) ---

@callback(
    Output('store-main-data', 'data'),
    Output('last-updated-text', 'children'),
    Input('interval-component', 'n_intervals'),
    prevent_initial_call=True
)
def auto_update_data(n):
    try:
        new_df = actualizar_y_cargar_datos_desde_excel()
        new_data_json = new_df.to_json(date_format='iso', orient='split')
        update_time_str = f"Datos actualizados desde Excel y DB: {datetime.now().strftime('%H:%M:%S')}"
        return new_data_json, update_time_str
    except Exception as e:
        print(f"Error durante la actualización automática de datos: {e}")
        traceback.print_exc()
        raise PreventUpdate

@callback(Output('contenedor-filtro-quincena', 'style'), Output('contenedor-filtro-semana', 'style'), Input('modo-filtro-tiempo', 'value'))
def controlar_visibilidad_filtros(modo):
    if modo == 'quincena': return {'display': 'block'}, {'display': 'none'}
    else: return {'display': 'none'}, {'display': 'block'}

def crear_tabla_conteo_diario(df, index_col, date_range=None):
    if df.empty: return pd.DataFrame(), [], []
    df['Fecha_Dia'] = df[COLUMNA_FECHA].dt.date
    total_general_col = df.groupby(index_col)[COLUMNA_ORDEN].count().to_frame('Total General')
    pivot_dia = pd.pivot_table(df, values=COLUMNA_ORDEN, index=index_col, columns='Fecha_Dia', aggfunc='count', fill_value=0)
    if date_range is not None:
        pivot_dia.columns = pd.to_datetime(pivot_dia.columns)
        pivot_dia = pivot_dia.reindex(columns=date_range, fill_value=0)
    resumen_df = total_general_col.join(pivot_dia).fillna(0).astype(int)
    resumen_df.sort_values(by='Total General', ascending=False, inplace=True)
    resumen_df.reset_index(inplace=True)
    if not resumen_df.empty:
        total_row = {index_col: 'Total General'}
        numeric_cols = resumen_df.select_dtypes(include='number').columns
        total_row.update(resumen_df[numeric_cols].sum().to_dict())
        total_row_df = pd.DataFrame([total_row])
        resumen_df = pd.concat([resumen_df, total_row_df], ignore_index=True)
    resumen_df.columns = [col.strftime('%d-%m-%Y') if hasattr(col, 'strftime') else col for col in resumen_df.columns]
    dia_cols = sorted([c for c in resumen_df.columns if c not in [index_col, 'Total General']], key=lambda x: pd.to_datetime(x, format='%d-%m-%Y'))
    column_order = [index_col] + dia_cols + ['Total General']
    resumen_df = resumen_df[column_order]
    return resumen_df, resumen_df.to_dict('records'), [{'name': c, 'id': c} for c in column_order]

def crear_tabla_porcentaje_corregido(df, index_col, date_range=None):
    if df.empty: return pd.DataFrame(), [], []
    df['Fecha_Dia'] = df[COLUMNA_FECHA].dt.date
    pivot_total = pd.pivot_table(df, values=COLUMNA_ORDEN, index=index_col, columns='Fecha_Dia', aggfunc='count', fill_value=0)
    pivot_corregido = pd.pivot_table(df[df[COLUMNA_STATUS] == 'Corregido'], values=COLUMNA_ORDEN, index=index_col, columns='Fecha_Dia', aggfunc='count', fill_value=0)
    if date_range is not None:
        pivot_total.columns = pd.to_datetime(pivot_total.columns)
        pivot_corregido.columns = pd.to_datetime(pivot_corregido.columns)
        pivot_total = pivot_total.reindex(columns=date_range, fill_value=0)
        pivot_corregido = pivot_corregido.reindex(columns=date_range, fill_value=0)
    pivot_porcentaje = (pivot_corregido / pivot_total).fillna(0)
    total_general_counts = df.groupby(index_col)[COLUMNA_ORDEN].count()
    resumen_df = pivot_porcentaje
    resumen_df['Total General'] = total_general_counts
    resumen_df.fillna(0, inplace=True)
    resumen_df.sort_values(by='Total General', ascending=False, inplace=True)
    for col in [c for c in resumen_df.columns if c != 'Total General']:
        resumen_df[col] = resumen_df[col].apply(lambda x: f"{x:.0%}")
    resumen_df['Total General'] = resumen_df['Total General'].astype(int)
    resumen_df.reset_index(inplace=True)
    resumen_df.columns = [col.strftime('%d-%m-%Y') if hasattr(col, 'strftime') else col for col in resumen_df.columns]
    dia_cols = sorted([c for c in resumen_df.columns if c not in [index_col, 'Total General']], key=lambda x: pd.to_datetime(x, format='%d-%m-%Y'))
    column_order = [index_col] + dia_cols + ['Total General']
    resumen_df = resumen_df[column_order]
    return resumen_df, resumen_df.to_dict('records'), [{'name': c, 'id': c} for c in column_order]

@callback(
    Output('tabla-resumen-mensual', 'data'), Output('tabla-resumen-mensual', 'columns'),
    Output('tabla-resumen-torre', 'data'), Output('tabla-resumen-torre', 'columns'),
    Output('tabla-resumen-status', 'data'), Output('tabla-resumen-status', 'columns'),
    Output('tabla-resumen-ejecutivo-conteo', 'data'), Output('tabla-resumen-ejecutivo-conteo', 'columns'),
    Output('tabla-resumen-ejecutivo-porcentaje', 'data'), Output('tabla-resumen-ejecutivo-porcentaje', 'columns'),
    Output('tarjetas-kpi-mensual', 'children'),
    Output('tarjetas-kpi-diario', 'children'),
    Output('tarjetas-kpi-graficos', 'children'),
    Output('grafico-torta-torre', 'figure'),
    Output('grafico-barras-resolutividad', 'figure'),
    Output('grafico-volumen-ejecutivo', 'figure'),
    Output('grafico-composicion-status', 'figure'),
    Input('store-main-data', 'data'),
    Input('filtro-mes', 'value'), Input('filtro-quincena', 'value'), Input('filtro-semana', 'value'),
    Input('filtro-torre', 'value'), Input('filtro-ejecutivo', 'value'),
    State('modo-filtro-tiempo', 'value')
)
def actualizar_dashboard_completo(json_data, meses, quincena, semanas, torres, ejecutivos, modo_tiempo):
    if not json_data:
        raise PreventUpdate
        
    df_principal = pd.read_json(io.StringIO(json_data), orient='split')
    df_principal[COLUMNA_FECHA] = pd.to_datetime(df_principal[COLUMNA_FECHA])
    
    dff = df_principal.copy()
    
    if meses: dff = dff[dff['Mes'].isin(meses)]
    if modo_tiempo == 'quincena' and quincena:
        dff = dff[dff[COLUMNA_FECHA].dt.day <= 15 if quincena == 1 else dff[COLUMNA_FECHA].dt.day > 15]
    elif modo_tiempo == 'semana' and semanas:
        dff = dff[dff['Semana_Num'].isin(semanas)]
    if torres: dff = dff[dff[COLUMNA_TORRE].isin(torres)]
    if ejecutivos: dff = dff[dff[COLUMNA_ANALISTA].isin(ejecutivos)]

    if dff.empty:
        empty_df_dict = [{'Nota': 'No hay datos para los filtros seleccionados'}]
        empty_cols = [{'name': 'Nota', 'id': 'Nota'}]
        no_data_msg = [dbc.Col(dbc.Alert("No hay datos para mostrar con los filtros seleccionados.", color="warning"))]
        return (empty_df_dict, empty_cols, empty_df_dict, empty_cols, empty_df_dict, empty_cols,
                empty_df_dict, empty_cols, empty_df_dict, empty_cols, no_data_msg, no_data_msg, no_data_msg,
                {}, {}, {}, {})
    
    all_months_ordered_local = sorted(df_principal['Mes'].unique(), key=lambda m: pd.to_datetime(f'01-{m}-2025', format='%d-%B-%Y').month)
    
    pivot_mensual = pd.pivot_table(dff, values=COLUMNA_ORDEN, index=[COLUMNA_TORRE, COLUMNA_ANALISTA], columns='Mes', aggfunc='count', fill_value=0)
    pivot_mensual['Total General'] = pivot_mensual.sum(axis=1)
    active_months = dff['Mes'].unique()
    month_order_map = {month: i for i, month in enumerate(all_months_ordered_local)}
    sorted_active_months = sorted(active_months, key=lambda m: month_order_map.get(m, 99))
    if 'Total General' in pivot_mensual.columns: pivot_mensual = pivot_mensual[sorted_active_months + ['Total General']]
    records = []
    torre_totals = dff.groupby(COLUMNA_TORRE)[COLUMNA_ORDEN].count().sort_values(ascending=False)
    for torre in torre_totals.index:
        df_torre_pivot = pivot_mensual.loc[torre]
        torre_sum = df_torre_pivot.sum()
        torre_row = {'Etiquetas de Fila': torre, 'Tipo': 'Torre'}; torre_row.update(torre_sum); records.append(torre_row)
        if isinstance(df_torre_pivot, pd.Series):
            ejec_row = {'Etiquetas de Fila': f'     {df_torre_pivot.name}', 'Tipo': 'Ejecutivo'}; ejec_row.update(df_torre_pivot); records.append(ejec_row)
        else:
            for ejecutivo_name, data in df_torre_pivot.iterrows():
                ejec_row = {'Etiquetas de Fila': f'     {ejecutivo_name}', 'Tipo': 'Ejecutivo'}; ejec_row.update(data); records.append(ejec_row)
    df_mensual_final = pd.DataFrame(records)
    cols_mensual = [{'name': c, 'id': c} for c in df_mensual_final.columns if c != 'Tipo']
    data_mensual = df_mensual_final.to_dict('records')

    date_range_for_tables = None
    if modo_tiempo == 'semana' and semanas:
        min_date = df_principal[df_principal['Semana_Num'].isin(semanas)]['WeekStartDate'].min()
        max_date = df_principal[df_principal['Semana_Num'].isin(semanas)]['WeekEndDate'].max()
        date_range_for_tables = pd.date_range(start=min_date, end=max_date)

    _, data_torre, cols_torre = crear_tabla_conteo_diario(dff, COLUMNA_TORRE, date_range_for_tables)
    _, data_status, cols_status = crear_tabla_conteo_diario(dff, COLUMNA_STATUS, date_range_for_tables)
    _, data_ejecutivo_conteo, cols_ejecutivo_conteo = crear_tabla_conteo_diario(dff, COLUMNA_ANALISTA, date_range_for_tables)
    _, data_ejecutivo_porcentaje, cols_ejecutivo_porcentaje = crear_tabla_porcentaje_corregido(dff, COLUMNA_ANALISTA, date_range_for_tables)

    dias_trabajados = dff[COLUMNA_FECHA].dt.normalize().nunique()
    gestion_totales = dff[COLUMNA_ORDEN].count()
    total_ejecutivos = dff[COLUMNA_ANALISTA].nunique()
    if dias_trabajados > 0 and total_ejecutivos > 0:
        if 'Capacidad' in dff[COLUMNA_STATUS].unique():
            total_capacidad = dff[dff[COLUMNA_STATUS] == 'Capacidad'][COLUMNA_ORDEN].count()
            gestiones_atendidas_raw = (gestion_totales - total_capacidad) / gestion_totales if gestion_totales > 0 else 0
            gestion_fte_dia = int(((gestion_totales - total_capacidad) / dias_trabajados) / total_ejecutivos)
        else:
            gestiones_atendidas_raw = 1 if gestion_totales > 0 else 0
            gestion_fte_dia = int((gestion_totales / dias_trabajados) / total_ejecutivos)
    else:
        gestiones_atendidas_raw = 0
        gestion_fte_dia = 0
    gestiones_atendidas = f"{gestiones_atendidas_raw:.2%}"
    total_corregido = dff[dff[COLUMNA_STATUS] == 'Corregido'][COLUMNA_ORDEN].count()
    tasa_resolutividad_raw = (total_corregido / gestion_totales) if gestion_totales > 0 else 0
    tasa_resolutividad = f"{tasa_resolutividad_raw:.2%}"
    def crear_tarjeta_kpi(titulo, valor, color_valor="primary"):
        return dbc.Col(dbc.Card(dbc.CardBody([html.H6(titulo, className="card-title text-muted"), html.H3(valor, className=f"card-text text-{color_valor}") ]), className="shadow-sm text-center"))
    tarjetas = [crear_tarjeta_kpi("Gestiones Totales", f"{gestion_totales}"), crear_tarjeta_kpi("Total Ejecutivos", f"{total_ejecutivos}"), crear_tarjeta_kpi("Gestiones Atendidas", gestiones_atendidas, "success"), crear_tarjeta_kpi("Tasa de Resolutividad", tasa_resolutividad, "info"), crear_tarjeta_kpi("Gestión FTE Día", f"{gestion_fte_dia}", "dark")]
    
    df_torre_chart = dff.groupby(COLUMNA_TORRE)[COLUMNA_ORDEN].count().reset_index()
    fig_torta_torre = px.pie(df_torre_chart, names=COLUMNA_TORRE, values=COLUMNA_ORDEN, title='Distribución de Gestiones por Torre', hole=.4, template='plotly_white')
    fig_torta_torre.update_traces(textposition='inside', textinfo='percent+label', hoverinfo='label+percent+value', marker=dict(line=dict(color='#000000', width=1)))
    fig_torta_torre.update_layout(showlegend=False, title_x=0.5)
    df_ejec_total = dff.groupby(COLUMNA_ANALISTA)[COLUMNA_ORDEN].count()
    df_ejec_corr = dff[dff[COLUMNA_STATUS]=='Corregido'].groupby(COLUMNA_ANALISTA)[COLUMNA_ORDEN].count()
    df_resolutividad = ((df_ejec_corr / df_ejec_total).fillna(0) * 100).reset_index(name='Tasa de Resolutividad').sort_values('Tasa de Resolutividad', ascending=False)
    fig_bar_resolutividad = px.bar(df_resolutividad, x='Tasa de Resolutividad', y=COLUMNA_ANALISTA, title='Tasa de Resolutividad por Ejecutivo', text_auto='.0f', orientation='h', template='plotly_white')
    fig_bar_resolutividad.update_traces(texttemplate='%{x:.0f}%', textposition='outside', marker_color='#28a745')
    fig_bar_resolutividad.update_layout(yaxis={'categoryorder':'total ascending'}, xaxis_title='Porcentaje (%)', yaxis_title=None, title_x=0.5)
    df_volumen_ejec = dff.groupby(COLUMNA_ANALISTA)[COLUMNA_ORDEN].count().reset_index(name='Cantidad')
    fig_volumen_ejec = px.pie(df_volumen_ejec, names=COLUMNA_ANALISTA, values='Cantidad', title='Distribución de Gestiones por Ejecutivo', hole=.4, template='plotly_white')
    fig_volumen_ejec.update_traces(textposition='inside', textinfo='percent+label', hoverinfo='label+percent+value', marker=dict(line=dict(color='#000000', width=1)))
    fig_volumen_ejec.update_layout(showlegend=False, title_x=0.5)
    df_status_exec_chart = dff.groupby([COLUMNA_ANALISTA, COLUMNA_STATUS])[COLUMNA_ORDEN].count().reset_index(name='Cantidad')
    total_volume_order = dff.groupby(COLUMNA_ANALISTA)[COLUMNA_ORDEN].count().sort_values(ascending=False).index
    fig_composicion_status = px.bar(df_status_exec_chart, x=COLUMNA_ANALISTA, y='Cantidad', color=COLUMNA_STATUS, title='Composición de Status por Ejecutivo (Cantidad)', template='plotly_white', text_auto=True)
    fig_composicion_status.update_layout(barmode='stack', xaxis_title=None, yaxis_title='Cantidad de Gestiones', title_x=0.5, xaxis={'categoryorder':'array', 'categoryarray': total_volume_order})
    
    return (
        data_mensual, cols_mensual, 
        data_torre, cols_torre, 
        data_status, cols_status, 
        data_ejecutivo_conteo, cols_ejecutivo_conteo, 
        data_ejecutivo_porcentaje, cols_ejecutivo_porcentaje, 
        tarjetas, tarjetas, tarjetas, 
        fig_torta_torre, fig_bar_resolutividad, fig_volumen_ejec, fig_composicion_status
    )

@callback(
    Output('filtro-mes', 'value'), Output('filtro-quincena', 'value'), Output('filtro-semana', 'value'),
    Output('filtro-torre', 'value'), Output('filtro-ejecutivo', 'value'), Output('modo-filtro-tiempo', 'value'),
    Input('btn-limpiar', 'n_clicks'),
    prevent_initial_call=True
)
def limpiar_filtros(n_clicks):
    return [], None, [], [], [], 'quincena'

@callback(
    Output('download-preview-container', 'children'),
    Output('store-download-raw-data', 'data'),
    Output('store-resumen-conteo-data', 'data'),
    Output('store-resumen-porcentaje-data', 'data'),
    Output('btn-download-all', 'disabled'),
    Input('btn-generate-download', 'n_clicks'),
    State('filtro-mes', 'value'),
    State('filtro-quincena', 'value'),
    State('filtro-semana', 'value'),
    State('filtro-torre', 'value'),
    State('filtro-ejecutivo', 'value'),
    State('modo-filtro-tiempo', 'value'),
    State('download-date-picker', 'start_date'),
    State('download-date-picker', 'end_date'),
    State('store-main-data', 'data'),
    prevent_initial_call=True
)
def generate_download_file(n_clicks, meses, quincena, semanas, torres, ejecutivos, modo_tiempo, start_date, end_date, json_data):
    if not start_date or not end_date or not json_data:
        return dbc.Alert("Fechas de inicio/fin no seleccionadas.", color="warning"), None, None, None, True
    df = pd.read_json(io.StringIO(json_data), orient='split')
    df[COLUMNA_FECHA] = pd.to_datetime(df[COLUMNA_FECHA])
    dff = df.copy()
    if meses: dff = dff[dff['Mes'].isin(meses)]
    if modo_tiempo == 'quincena' and quincena:
        dff = dff[dff[COLUMNA_FECHA].dt.day <= 15 if quincena == 1 else dff[COLUMNA_FECHA].dt.day > 15]
    elif modo_tiempo == 'semana' and semanas:
        dff = dff[dff['Semana_Num'].isin(semanas)]
    if torres: dff = dff[dff[COLUMNA_TORRE].isin(torres)]
    if ejecutivos: dff = dff[dff[COLUMNA_ANALISTA].isin(ejecutivos)]
    start_date_dt = pd.to_datetime(start_date)
    end_date_dt = pd.to_datetime(end_date)
    dff_download = dff[(dff[COLUMNA_FECHA] >= start_date_dt) & (dff[COLUMNA_FECHA] <= end_date_dt)]
    if dff_download.empty:
        return dbc.Alert("No hay datos para los filtros y rango de fechas seleccionados.", color="info"), None, None, None, True
    df_conteo, _, _ = crear_tabla_conteo_diario(dff_download, COLUMNA_ANALISTA)
    df_porcentaje, _, _ = crear_tabla_porcentaje_corregido(dff_download, COLUMNA_ANALISTA)
    preview_table = dash_table.DataTable(
        data=dff_download.head(10).to_dict('records'),
        columns=[{'name': i, 'id': i} for i in dff_download.columns if i not in ['Year', 'Semana_Num', 'WeekStartDate', 'WeekEndDate', 'WeekLabel']],
        page_size=10,
        style_table={'overflowX': 'auto'},
    )
    preview_content = [html.H5(f"Vista previa de los datos detallados (primeras 10 de {len(dff_download)} filas):"), preview_table]
    return (preview_content, dff_download.to_json(date_format='iso', orient='split'),
            df_conteo.to_json(orient='split'), df_porcentaje.to_json(orient='split'), False)

@callback(
    Output("download-excel", "data"),
    Input("btn-download-all", "n_clicks"),
    State('store-download-raw-data', 'data'),
    State('store-resumen-conteo-data', 'data'),
    State('store-resumen-porcentaje-data', 'data'),
    prevent_initial_call=True,
)
def download_all_in_one_excel(n_clicks, json_raw, json_conteo, json_porcentaje):
    if not n_clicks or not json_raw or not json_conteo or not json_porcentaje:
        raise PreventUpdate
    df_raw = pd.read_json(io.StringIO(json_raw), orient='split')
    df_conteo = pd.read_json(io.StringIO(json_conteo), orient='split')
    df_porcentaje = pd.read_json(io.StringIO(json_porcentaje), orient='split')
    if COLUMNA_FECHA in df_raw.columns:
        df_raw[COLUMNA_FECHA] = pd.to_datetime(df_raw[COLUMNA_FECHA]).dt.date
    cols_to_drop = ['Year', 'Semana_Num', 'WeekStartDate', 'WeekEndDate', 'WeekLabel']
    df_raw = df_raw.drop(columns=[col for col in cols_to_drop if col in df_raw.columns])
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_raw.to_excel(writer, sheet_name='Datos Detallados', index=False)
        df_conteo.to_excel(writer, sheet_name='Resumen Cantidad', index=False)
        df_porcentaje.to_excel(writer, sheet_name='Resumen Resolutividad', index=False)
    output.seek(0)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"reporte_completo_{timestamp}.xlsx"
    return dcc.send_bytes(output.read(), filename)


# --- 6. INICIAR EL SERVIDOR ---
if __name__ == '__main__':
    app.run(debug=True)