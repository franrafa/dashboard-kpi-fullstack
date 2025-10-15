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

# --- CONFIGURACIÓN ---
try:
    locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
except locale.Error:
    locale.setlocale(locale.LC_TIME, 'Spanish')

RUTA_ARCHIVO = r"C:\Users\Haintech\Desktop\Consolidado_Ordenes_PowerQuery\FullStack_Consolidado.xlsx"
HOJA_DATOS = "Consolidado FullStack"
COLUMNA_FECHA = "Fecha"
COLUMNA_ANALISTA = "Ejecutivo"
COLUMNA_ORDEN = "Número de pedido"
COLUMNA_STATUS = "Status Real"
COLUMNA_TORRE = "Torre"

EJECUTIVOS_FILTRADOS = [
    "Miguel Mantilla",
    "Miguel Aravena",
    "Nilsson Diaz",
    "Francisco Narvaez",
    "Carlos Quezada",
    "Gia Marin",
    "Marcos Coyan"
]

VALID_USERNAME_PASSWORD_PAIRS = {
    'haintech': 'dashboard2025'
}

# --- INICIO DEL CÓDIGO ---
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.LUX], suppress_callback_exceptions=True)

auth = dash_auth.BasicAuth(
    app,
    VALID_USERNAME_PASSWORD_PAIRS
)

# --- FUNCIÓN DE CARGA DE DATOS ---
def load_data():
    df = pd.read_excel(RUTA_ARCHIVO, sheet_name=HOJA_DATOS)
    df[COLUMNA_FECHA] = pd.to_datetime(df[COLUMNA_FECHA], errors='coerce')

    df.dropna(subset=[COLUMNA_FECHA, COLUMNA_ANALISTA, COLUMNA_TORRE, COLUMNA_STATUS], inplace=True)
    df[COLUMNA_ANALISTA] = df[COLUMNA_ANALISTA].astype(str).str.strip()
    df[COLUMNA_TORRE] = df[COLUMNA_TORRE].astype(str).str.strip()
    df[COLUMNA_STATUS] = df[COLUMNA_STATUS].astype(str).str.strip()
    df = df[df[COLUMNA_ANALISTA] != ""]
    df = df[df[COLUMNA_TORRE] != ""]

    df = df[df[COLUMNA_ANALISTA].isin(EJECUTIVOS_FILTRADOS)]
    df = df[df[COLUMNA_FECHA].dt.month >= 8]
    df.sort_values(by=COLUMNA_FECHA, inplace=True)

    df['Mes'] = df[COLUMNA_FECHA].dt.strftime('%B').str.capitalize()
    df['Year'] = df[COLUMNA_FECHA].dt.isocalendar().year
    df['Semana_Num'] = df[COLUMNA_FECHA].dt.isocalendar().week
    df['WeekStartDate'] = pd.to_datetime(df['Year'].astype(str) + df['Semana_Num'].astype(str) + '1', format='%G%V%u')
    df['WeekEndDate'] = df['WeekStartDate'] + pd.to_timedelta('6 days')
    df['WeekLabel'] = "Semana " + df['Semana_Num'].astype(str) + " (" + df['WeekStartDate'].dt.strftime('%d %b') + " - " + df['WeekEndDate'].dt.strftime('%d %b') + ")"
    return df

# --- 1. LECTURA INICIAL DE DATOS ---
try:
    df_principal = load_data()
    last_modified_time = os.path.getmtime(RUTA_ARCHIVO)
    
    meses_disponibles = sorted(df_principal['Mes'].unique(), key=lambda m: pd.to_datetime(f'01-{m}-2025', format='%d-%B-%Y').month)
    week_map = df_principal[['Semana_Num', 'WeekLabel']].drop_duplicates().sort_values('Semana_Num')
    semanas_disponibles_options = week_map.apply(lambda row: {'label': row['WeekLabel'], 'value': row['Semana_Num']}, axis=1).tolist()
    ejecutivos_disponibles = sorted(df_principal[COLUMNA_ANALISTA].unique())
    torres_disponibles = sorted(df_principal[COLUMNA_TORRE].unique())

    datos_cargados_correctamente = True
except Exception as e:
    error_mensaje = f"Ocurrió un error al cargar o procesar el archivo Excel: {e}"
    datos_cargados_correctamente = False
    df_principal = pd.DataFrame()
    last_modified_time = 0

# --- 2. DISEÑO DE LA APLICACIÓN WEB (LAYOUT) ---
if datos_cargados_correctamente:
    app.layout = dbc.Container([
        dcc.Store(id='store-main-data', data=df_principal.to_json(date_format='iso', orient='split')),
        dcc.Store(id='store-last-modified', data=last_modified_time),
        dcc.Interval(id='interval-component', interval=30 * 1000, n_intervals=0),
        dcc.Download(id="download-excel"),
        dcc.Store(id='store-download-data'),
        
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
            dbc.Tab(label="Detalle Diario", children=[dbc.Row(id='tarjetas-kpi-diario', className="my-4"), dbc.Row([dbc.Col([html.H4("Resumen Diario por Torre", className="border-bottom pb-2 my-3"), dash_table.DataTable(id='tabla-resumen-torre', style_table={'overflowX': 'auto'}, style_header={'backgroundColor': '#e8f5e9'}, style_cell={'textAlign': 'center', 'minWidth': '120px'}, style_cell_conditional=[{'if': {'column_id': COLUMNA_TORRE}, 'textAlign': 'left', 'fontWeight': 'bold', 'minWidth': '180px'}, {'if': {'column_id': 'Total General'}, 'fontWeight': 'bold', 'backgroundColor': '#e8f5e9'}])], width=12)], className="mb-4"), dbc.Row([dbc.Col([html.H4("Resumen Diario por Status", className="border-bottom pb-2 mb-3"), dash_table.DataTable(id='tabla-resumen-status', style_table={'overflowX': 'auto'}, style_header={'backgroundColor': '#fff3e0'}, style_cell={'textAlign': 'center', 'minWidth': '120px'}, style_cell_conditional=[{'if': {'column_id': COLUMNA_STATUS}, 'textAlign': 'left', 'fontWeight': 'bold', 'minWidth': '180px'}, {'if': {'column_id': 'Total General'}, 'fontWeight': 'bold', 'backgroundColor': '#fff3e0'}])], width=12)], className="mb-4"), dbc.Row([dbc.Col([html.H4("Resumen Diario por Ejecutivo (Cantidad)", className="border-bottom pb-2 mb-3"), dash_table.DataTable(id='tabla-resumen-ejecutivo-conteo', style_table={'overflowX': 'auto'}, style_header={'backgroundColor': '#f2e3fd'}, style_cell={'textAlign': 'center', 'minWidth': '120px'}, style_cell_conditional=[{'if': {'column_id': COLUMNA_ANALISTA}, 'textAlign': 'left', 'fontWeight': 'bold', 'minWidth': '180px'}, {'if': {'column_id': 'Total General'}, 'fontWeight': 'bold', 'backgroundColor': '#f2e3fd'}])], width=12)], className="mb-4"), dbc.Row([dbc.Col([html.H4("Porcentaje de Resolutividad Diario por Ejecutivo", className="border-bottom pb-2 mb-3"), dash_table.DataTable(id='tabla-resumen-ejecutivo-porcentaje', style_table={'overflowX': 'auto'}, style_header={'backgroundColor': '#e3f2fd'}, style_cell={'textAlign': 'center', 'minWidth': '120px'}, style_cell_conditional=[{'if': {'column_id': COLUMNA_ANALISTA}, 'textAlign': 'left', 'fontWeight': 'bold', 'minWidth': '180px'}, {'if': {'column_id': 'Total General'}, 'fontWeight': 'bold', 'backgroundColor': '#e3f2fd'}])], width=12)], className="mb-4")]),
            dbc.Tab(label="Gráficos", children=[dbc.Row(id='tarjetas-kpi-graficos', className="my-4"), dbc.Row([dbc.Col(dcc.Graph(id='grafico-torta-torre'), md=6), dbc.Col(dcc.Graph(id='grafico-barras-resolutividad'), md=6)], className="my-4"), dbc.Row([dbc.Col(dcc.Graph(id='grafico-volumen-ejecutivo'), md=6), dbc.Col(dcc.Graph(id='grafico-composicion-status'), md=6)], className="my-4")]),
            dbc.Tab(label="Descargar", children=[
                dbc.Row([
                    dbc.Col([
                        html.H4("Seleccionar Rango de Fechas para Descarga", className="mt-4"),
                        dcc.DatePickerRange(
                            id='download-date-picker',
                            min_date_allowed=df_principal[COLUMNA_FECHA].min().date(),
                            max_date_allowed=df_principal[COLUMNA_FECHA].max().date(),
                            start_date=df_principal[COLUMNA_FECHA].min().date(),
                            end_date=df_principal[COLUMNA_FECHA].max().date(),
                            display_format='DD/MM/YYYY'
                        ),
                        dbc.Button("Generar Vista Previa", id="btn-preview", color="primary", className="mt-3"),
                        html.Div(id="download-preview-container", className="mt-4")
                    ], className="text-center", md={'size': 8, 'offset': 2})
                ], className="my-4")
            ])
        ], className="mt-4"),
        html.Div(id='last-updated-text', style={'textAlign': 'right', 'color': 'grey', 'marginTop': '20px'})
    ], fluid=True)
else:
    app.layout = dbc.Container([dbc.Alert(error_mensaje, color="danger", className="mt-4")])

# --- 3. LÓGICA DE INTERACTIVIDAD (CALLBACKS) ---
@callback(Output('store-main-data', 'data'), Output('store-last-modified', 'data'), Output('last-updated-text', 'children'), Input('interval-component', 'n_intervals'), State('store-last-modified', 'data'))
def auto_update_data(n, stored_last_modified):
    try:
        current_last_modified = os.path.getmtime(RUTA_ARCHIVO)
        if current_last_modified > stored_last_modified:
            new_df = load_data()
            new_data_json = new_df.to_json(date_format='iso', orient='split')
            update_time_str = f"Última actualización: {datetime.now().strftime('%H:%M:%S')}"
            return new_data_json, current_last_modified, update_time_str
    except Exception as e:
        print(f"Error al actualizar datos: {e}")
    raise PreventUpdate

@callback(Output('contenedor-filtro-quincena', 'style'), Output('contenedor-filtro-semana', 'style'), Input('modo-filtro-tiempo', 'value'))
def controlar_visibilidad_filtros(modo):
    if modo == 'quincena': return {'display': 'block'}, {'display': 'none'}
    else: return {'display': 'none'}, {'display': 'block'}

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
    Input('filtro-torre', 'value'), Input('filtro-ejecutivo', 'value'), Input('btn-limpiar', 'n_clicks'),
    State('modo-filtro-tiempo', 'value')
)
def actualizar_dashboard_completo(json_data, meses, quincena, semanas, torres, ejecutivos, n_clicks, modo_tiempo):
    if not json_data:
        raise PreventUpdate

    df_principal = pd.read_json(io.StringIO(json_data), orient='split')
    df_principal[COLUMNA_FECHA] = pd.to_datetime(df_principal[COLUMNA_FECHA])

    dff = df_principal.copy()
    
    ctx = dash.callback_context
    if ctx.triggered and 'btn-limpiar' in ctx.triggered[0]['prop_id']:
        meses, quincena, semanas, torres, ejecutivos = None, None, None, None, None
    else:
        if meses: dff = dff[dff['Mes'].isin(meses)]
        if modo_tiempo == 'quincena' and quincena:
            dff = dff[dff[COLUMNA_FECHA].dt.day <= 15 if quincena == 1 else dff[COLUMNA_FECHA].dt.day > 15]
        elif modo_tiempo == 'semana' and semanas:
            dff = dff[dff['Semana_Num'].isin(semanas)]
        if torres: dff = dff[dff[COLUMNA_TORRE].isin(torres)]
        if ejecutivos: dff = dff[dff[COLUMNA_ANALISTA].isin(ejecutivos)]

    if dff.empty:
        empty_df, empty_cols, tarjetas, empty_fig = pd.DataFrame(), [], [], {}
        return empty_df.to_dict('records'), empty_cols, empty_df.to_dict('records'), empty_cols, empty_df.to_dict('records'), empty_cols, empty_df.to_dict('records'), empty_cols, empty_df.to_dict('records'), empty_cols, tarjetas, tarjetas, tarjetas, empty_fig, empty_fig, empty_fig, empty_fig
    
    # --- Cálculos y lógica ---
    all_months_ordered_local = sorted(df_principal['Mes'].unique(), key=lambda m: pd.to_datetime(f'01-{m}-2025', format='%d-%B-%Y').month)
    
    # --- Cálculos para la tabla mensual ---
    pivot_mensual = pd.pivot_table(dff, values=COLUMNA_ORDEN, index=[COLUMNA_TORRE, COLUMNA_ANALISTA], columns='Mes', aggfunc='count', fill_value=0)
    pivot_mensual['Total General'] = pivot_mensual.sum(axis=1)
    active_months = dff['Mes'].unique()
    month_order_map = {month: i for i, month in enumerate(all_months_ordered_local)}
    sorted_active_months = sorted(active_months, key=lambda m: month_order_map.get(m, 99))
    if 'Total General' in pivot_mensual.columns: pivot_mensual = pivot_mensual[sorted_active_months + ['Total General']]
    records = []
    torre_totals = dff.groupby(COLUMNA_TORRE)[COLUMNA_ORDEN].count().sort_values(ascending=False)
    for torre in torre_totals.index:
        df_torre = pivot_mensual.loc[torre]
        torre_sum = df_torre.sum()
        torre_row = {'Etiquetas de Fila': torre, 'Tipo': 'Torre'}; torre_row.update(torre_sum); records.append(torre_row)
        if isinstance(df_torre, pd.Series):
            ejec_row = {'Etiquetas de Fila': f'     {df_torre.name}', 'Tipo': 'Ejecutivo'}; ejec_row.update(df_torre); records.append(ejec_row)
        else:
            for ejecutivo, data in df_torre.iterrows():
                ejec_row = {'Etiquetas de Fila': f'     {ejecutivo}', 'Tipo': 'Ejecutivo'}; ejec_row.update(data); records.append(ejec_row)
    df_mensual_final = pd.DataFrame(records)
    cols_mensual = [{'name': c, 'id': c} for c in df_mensual_final.columns if c != 'Tipo']
    data_mensual = df_mensual_final.to_dict('records')

    date_range_for_tables = None
    if modo_tiempo == 'semana' and semanas:
        min_date = df_principal[df_principal['Semana_Num'].isin(semanas)]['WeekStartDate'].min()
        max_date = df_principal[df_principal['Semana_Num'].isin(semanas)]['WeekEndDate'].max()
        date_range_for_tables = pd.date_range(start=min_date, end=max_date)

    def crear_tabla_conteo_diario(df, index_col, date_range=None):
        if df.empty: return [], []
        df['Fecha_Dia'] = df[COLUMNA_FECHA].dt.date
        total_general = df.groupby(index_col)[COLUMNA_ORDEN].count().to_frame('Total General')
        pivot_dia = pd.pivot_table(df, values=COLUMNA_ORDEN, index=index_col, columns='Fecha_Dia', aggfunc='count', fill_value=0)
        
        if date_range is not None:
            pivot_dia.columns = pd.to_datetime(pivot_dia.columns)
            pivot_dia = pivot_dia.reindex(columns=date_range, fill_value=0)
            
        resumen_df = total_general.join(pivot_dia).fillna(0).astype(int); resumen_df.sort_values(by='Total General', ascending=False, inplace=True); resumen_df.reset_index(inplace=True)
        resumen_df.columns = [col.strftime('%d-%m-%Y') if hasattr(col, 'strftime') else col for col in resumen_df.columns]
        dia_cols = sorted([c for c in resumen_df.columns if c not in [index_col, 'Total General']], key=lambda x: pd.to_datetime(x, format='%d-%m-%Y'))
        column_order = [index_col] + dia_cols + ['Total General']
        resumen_df = resumen_df[column_order]
        return resumen_df.to_dict('records'), [{'name': c, 'id': c} for c in column_order]

    def crear_tabla_porcentaje_corregido(df, index_col, date_range=None):
        if df.empty: return [], []
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
        resumen_df = pivot_porcentaje; resumen_df['Total General'] = total_general_counts; resumen_df.fillna(0, inplace=True); resumen_df.sort_values(by='Total General', ascending=False, inplace=True)
        for col in [c for c in resumen_df.columns if c != 'Total General']: resumen_df[col] = resumen_df[col].apply(lambda x: f"{x:.0%}")
        resumen_df['Total General'] = resumen_df['Total General'].astype(int); resumen_df.reset_index(inplace=True)
        resumen_df.columns = [col.strftime('%d-%m-%Y') if hasattr(col, 'strftime') else col for col in resumen_df.columns]
        dia_cols = sorted([c for c in resumen_df.columns if c not in [index_col, 'Total General']], key=lambda x: pd.to_datetime(x, format='%d-%m-%Y'))
        column_order = [index_col] + dia_cols + ['Total General']
        resumen_df = resumen_df[column_order]
        return resumen_df.to_dict('records'), [{'name': c, 'id': c} for c in resumen_df.columns]

    data_torre, cols_torre = crear_tabla_conteo_diario(dff, COLUMNA_TORRE, date_range_for_tables)
    data_status, cols_status = crear_tabla_conteo_diario(dff, COLUMNA_STATUS, date_range_for_tables)
    data_ejecutivo_conteo, cols_ejecutivo_conteo = crear_tabla_conteo_diario(dff, COLUMNA_ANALISTA, date_range_for_tables)
    data_ejecutivo_porcentaje, cols_ejecutivo_porcentaje = crear_tabla_porcentaje_corregido(dff, COLUMNA_ANALISTA, date_range_for_tables)

    gestion_totales = dff[COLUMNA_ORDEN].count()
    total_ejecutivos = dff[COLUMNA_ANALISTA].nunique()
    if 'Capacidad' in dff[COLUMNA_STATUS].unique():
        total_capacidad = dff[dff[COLUMNA_STATUS] == 'Capacidad'][COLUMNA_ORDEN].count()
        gestiones_atendidas = (gestion_totales - total_capacidad) / gestion_totales if gestion_totales > 0 else 0
        gestion_fte_dia = int(((gestion_totales - total_capacidad) / 6) / total_ejecutivos if total_ejecutivos > 0 else 0)
    else:
        gestiones_atendidas = 1 if gestion_totales > 0 else 0
        gestion_fte_dia = int((gestion_totales / 6) / total_ejecutivos if total_ejecutivos > 0 else 0)
    total_corregido = dff[dff[COLUMNA_STATUS] == 'Corregido'][COLUMNA_ORDEN].count()
    tasa_resolutividad = (total_corregido / gestion_totales) if gestion_totales > 0 else 0
    
    def crear_tarjeta_kpi(titulo, valor, color_valor="primary"):
        return dbc.Col(dbc.Card(dbc.CardBody([html.H6(titulo, className="card-title text-muted"), html.H3(valor, className=f"card-text text-{color_valor}") ]), className="shadow-sm text-center"))

    tarjetas = [crear_tarjeta_kpi("Gestion Totales", f"{gestion_totales}"), crear_tarjeta_kpi("Total Ejecutivos", f"{total_ejecutivos}"), crear_tarjeta_kpi("Gestiones atendidas", f"{gestiones_atendidas:.0%}", "success"), crear_tarjeta_kpi("Tasa de resolutividad", f"{tasa_resolutividad:.0%}", "info"), crear_tarjeta_kpi("Gestion FTE dia", f"{gestion_fte_dia}", "dark")]
    
    df_torre_chart = dff.groupby(COLUMNA_TORRE)[COLUMNA_ORDEN].count().reset_index()
    fig_torta_torre = px.pie(df_torre_chart, names=COLUMNA_TORRE, values=COLUMNA_ORDEN, title='Distribución de Gestiones por Torre', hole=.4, template='plotly_white')
    fig_torta_torre.update_traces(textposition='inside', textinfo='percent+label', hoverinfo='label+percent+value', marker=dict(line=dict(color='#000000', width=1)))
    fig_torta_torre.update_layout(showlegend=False, title_x=0.5)

    df_ejec_total = dff.groupby(COLUMNA_ANALISTA)[COLUMNA_ORDEN].count()
    df_ejec_corr = dff[dff[COLUMNA_STATUS]=='Corregido'].groupby(COLUMNA_ANALISTA)[COLUMNA_ORDEN].count()
    df_resolutividad = ((df_ejec_corr / df_ejec_total).fillna(0) * 100).reset_index(name='Tasa de Resolutividad').sort_values('Tasa de Resolutividad', ascending=False)
    fig_bar_resolutividad = px.bar(df_resolutividad, x='Tasa de Resolutividad', y=COLUMNA_ANALISTA, title='Tasa de Resolutividad por Ejecutivo', text_auto='.0f', orientation='h', template='plotly_white')
    fig_bar_resolutividad.update_traces(texttemplate='%{x}%', textposition='outside', marker_color='#28a745')
    fig_bar_resolutividad.update_layout(yaxis={'categoryorder':'total ascending'}, xaxis_title='Porcentaje (%)', yaxis_title=None, title_x=0.5)
    
    df_volumen_ejec = dff.groupby(COLUMNA_ANALISTA)[COLUMNA_ORDEN].count().reset_index(name='Cantidad')
    fig_volumen_ejec = px.pie(df_volumen_ejec, names=COLUMNA_ANALISTA, values='Cantidad', title='Distribución de Gestiones por Ejecutivo', hole=.4, template='plotly_white')
    fig_volumen_ejec.update_traces(textposition='inside', textinfo='percent+label', hoverinfo='label+percent+value', marker=dict(line=dict(color='#000000', width=1)))
    fig_volumen_ejec.update_layout(showlegend=False, title_x=0.5)

    df_status_exec_chart = dff.groupby([COLUMNA_ANALISTA, COLUMNA_STATUS])[COLUMNA_ORDEN].count().reset_index(name='Cantidad')
    total_volume_order = dff.groupby(COLUMNA_ANALISTA)[COLUMNA_ORDEN].count().sort_values(ascending=False).index
    fig_composicion_status = px.bar(df_status_exec_chart, x=COLUMNA_ANALISTA, y='Cantidad', color=COLUMNA_STATUS, title='Composición de Status por Ejecutivo (Cantidad)', template='plotly_white', text_auto=True)
    fig_composicion_status.update_layout(barmode='stack', xaxis_title=None, yaxis_title='Cantidad de Gestiones', title_x=0.5, xaxis={'categoryorder':'array', 'categoryarray': total_volume_order})
    
    return data_mensual, cols_mensual, data_torre, cols_torre, data_status, cols_status, data_ejecutivo_conteo, cols_ejecutivo_conteo, data_ejecutivo_porcentaje, cols_ejecutivo_porcentaje, tarjetas, tarjetas, tarjetas, fig_torta_torre, fig_bar_resolutividad, fig_volumen_ejec, fig_composicion_status

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
    Output('store-download-data', 'data'),
    Input('btn-preview', 'n_clicks'),
    State('download-date-picker', 'start_date'),
    State('download-date-picker', 'end_date'),
    prevent_initial_call=True
)
def update_download_preview(n_clicks, start_date, end_date):
    if not start_date or not end_date:
        return None, None

    dff_download = df_principal[(df_principal[COLUMNA_FECHA] >= start_date) & (df_principal[COLUMNA_FECHA] <= end_date)]
    
    preview_table = dash_table.DataTable(
        data=dff_download.head(10).to_dict('records'),
        columns=[{'name': i, 'id': i} for i in dff_download.columns if i not in ['Year', 'Semana_Num', 'WeekStartDate', 'WeekEndDate', 'WeekLabel']],
        page_size=10,
        style_table={'overflowX': 'auto'},
    )
    
    download_button = dbc.Button("Descargar Selección como XLSX", id="btn-download", color="success", className="mt-3")
    
    return [html.H5(f"Mostrando primeras 10 de {len(dff_download)} filas:"), preview_table, download_button], dff_download.to_json(date_format='iso', orient='split')

@callback(
    Output("download-excel", "data"),
    Input("btn-download", "n_clicks"),
    State('store-download-data', 'data'),
    prevent_initial_call=True,
)
def download_as_excel(n_clicks, json_data):
    if not json_data:
        return None
    
    dff_download = pd.read_json(io.StringIO(json_data), orient='split')
    # Limpiar columnas auxiliares antes de descargar
    cols_to_drop = ['Year', 'Semana_Num', 'WeekStartDate', 'WeekEndDate', 'WeekLabel']
    dff_download = dff_download.drop(columns=[col for col in cols_to_drop if col in dff_download.columns])
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"consolidado_filtrado_{timestamp}.xlsx"
    
    return dcc.send_data_frame(dff_download.to_excel, filename, sheet_name="Datos", index=False)

# --- 4. INICIAR EL SERVIDOR ---
if __name__ == '__main__':
    app.run(debug=True)

