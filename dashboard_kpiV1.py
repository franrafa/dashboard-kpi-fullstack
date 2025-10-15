import pandas as pd
import plotly.express as px
import dash
import dash_auth
import dash_bootstrap_components as dbc
from dash import html, dcc, dash_table, Input, Output, callback, State
import locale

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
    "Carlos",
    "Gia Marin",
    "Marcos"
]

VALID_USERNAME_PASSWORD_PAIRS = {
    'haintech': 'dashboard2025'
}

# --- INICIO DEL CÓDIGO ---
# <<< MEJORA ESTÉTICA: Se cambia el tema a LUX para un look más moderno >>>
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.LUX])

auth = dash_auth.BasicAuth(
    app,
    VALID_USERNAME_PASSWORD_PAIRS
)

# --- 1. LECTURA Y LIMPIEZA DE DATOS ---
try:
    df_principal = pd.read_excel(RUTA_ARCHIVO, sheet_name=HOJA_DATOS)
    df_principal[COLUMNA_FECHA] = pd.to_datetime(df_principal[COLUMNA_FECHA], errors='coerce')

    df_principal.dropna(subset=[COLUMNA_FECHA, COLUMNA_ANALISTA, COLUMNA_TORRE, COLUMNA_STATUS], inplace=True)
    df_principal[COLUMNA_ANALISTA] = df_principal[COLUMNA_ANALISTA].astype(str).str.strip()
    df_principal[COLUMNA_TORRE] = df_principal[COLUMNA_TORRE].astype(str).str.strip()
    df_principal[COLUMNA_STATUS] = df_principal[COLUMNA_STATUS].astype(str).str.strip()
    df_principal = df_principal[df_principal[COLUMNA_ANALISTA] != ""]
    df_principal = df_principal[df_principal[COLUMNA_TORRE] != ""]

    df_principal = df_principal[df_principal[COLUMNA_ANALISTA].isin(EJECUTIVOS_FILTRADOS)]

    df_principal = df_principal[df_principal[COLUMNA_FECHA].dt.month >= 8]

    df_principal.sort_values(by=COLUMNA_FECHA, inplace=True)

    df_principal['Mes'] = df_principal[COLUMNA_FECHA].dt.strftime('%B').str.capitalize()
    df_principal['Semana_Num'] = df_principal[COLUMNA_FECHA].dt.isocalendar().week
    
    meses_disponibles = df_principal['Mes'].unique()
    semanas_disponibles = sorted(df_principal['Semana_Num'].unique())
    ejecutivos_disponibles = sorted(df_principal[COLUMNA_ANALISTA].unique())
    torres_disponibles = sorted(df_principal[COLUMNA_TORRE].unique())

    datos_cargados_correctamente = True
except Exception as e:
    error_mensaje = f"Ocurrió un error al cargar o procesar el archivo Excel: {e}"
    datos_cargados_correctamente = False
    df_principal = pd.DataFrame()

# --- 2. DISEÑO DE LA APLICACIÓN WEB (LAYOUT) ---
if datos_cargados_correctamente:
    app.layout = dbc.Container([
        dbc.Row(dbc.Col(html.H1("Dashboard Consolidado FullStack", className="text-center text-primary my-4"))),
        # <<< MEJORA ESTÉTICA: Se añade sombra al panel de filtros >>>
        dbc.Card(dbc.CardBody([
            dbc.Row([
                dbc.Col(dcc.Dropdown(id='filtro-mes', options=meses_disponibles, placeholder="Seleccionar Mes(es)", multi=True), md=3),
                dbc.Col([
                    html.Label("Filtrar por:", style={'fontWeight': 'bold'}),
                    dcc.RadioItems(
                        id='modo-filtro-tiempo',
                        options=[{'label': ' Quincena', 'value': 'quincena'}, {'label': ' Semana', 'value': 'semana'}],
                        value='quincena', inline=True, labelStyle={'margin-right': '10px'}
                    ),
                    html.Div(id='contenedor-filtro-quincena', children=[
                        dcc.Dropdown(id='filtro-quincena', options=[{'label': '1ra Quincena', 'value': 1}, {'label': '2da Quincena', 'value': 2}], placeholder="Seleccionar Quincena", className="mt-1")
                    ]),
                    html.Div(id='contenedor-filtro-semana', children=[
                        dcc.Dropdown(id='filtro-semana', options=[{'label': f"Semana {s}", 'value': s} for s in semanas_disponibles], placeholder="Seleccionar Semana(s)", multi=True, className="mt-1")
                    ], style={'display': 'none'})
                ], md=3),
                dbc.Col(dcc.Dropdown(id='filtro-torre', options=torres_disponibles, placeholder="Seleccionar Torre(s)", multi=True), md=3),
                dbc.Col(dcc.Dropdown(id='filtro-ejecutivo', options=ejecutivos_disponibles, placeholder="Seleccionar Ejecutivo(s)", multi=True), md=3),
            ]),
            dbc.Row(dbc.Col(dbc.Button("Limpiar Filtros", id="btn-limpiar", color="dark", outline=True, className="w-100 mt-3"), width=12))
        ]), className="mb-4 shadow"),

        dbc.Row(id='tarjetas-kpi', className="mb-4"),
        
        # <<< MEJORA ESTÉTICA: Títulos con línea inferior y tablas con estilo mejorado >>>
        dbc.Row([
            dbc.Col([html.H4("Resumen Diario por Torre", className="border-bottom pb-2 mb-3"), dash_table.DataTable(id='tabla-resumen-torre', style_header={'backgroundColor': '#e8f5e9'}, style_cell={'textAlign': 'center'}, style_cell_conditional=[{'if': {'column_id': COLUMNA_TORRE}, 'textAlign': 'left'}, {'if': {'column_id': 'Total General'},'fontWeight': 'bold', 'backgroundColor': '#e8f5e9'}])], width=12),
        ], className="mb-4"),

        dbc.Row([
            dbc.Col([html.H4("Resumen Diario por Status", className="border-bottom pb-2 mb-3"), dash_table.DataTable(id='tabla-resumen-status', style_header={'backgroundColor': '#fff3e0'}, style_cell={'textAlign': 'center'}, style_cell_conditional=[{'if': {'column_id': COLUMNA_STATUS}, 'textAlign': 'left'}, {'if': {'column_id': 'Total General'},'fontWeight': 'bold', 'backgroundColor': '#fff3e0'}])], width=12),
        ], className="mb-4"),
        
        dbc.Row([
            dbc.Col([html.H4("Resumen Diario por Ejecutivo", className="border-bottom pb-2 mb-3"), dash_table.DataTable(id='tabla-resumen-ejecutivo', style_header={'backgroundColor': '#e3f2fd'}, style_cell={'textAlign': 'center'}, style_cell_conditional=[{'if': {'column_id': COLUMNA_ANALISTA}, 'textAlign': 'left'}, {'if': {'column_id': 'Total General'},'fontWeight': 'bold', 'backgroundColor': '#e3f2fd'}])], width=12),
        ], className="mb-4")
    ], fluid=True)
else:
    app.layout = dbc.Container([dbc.Alert(error_mensaje, color="danger", className="mt-4")])

# --- 3. LÓGICA DE INTERACTIVIDAD (CALLBACKS) ---
@callback(
    Output('contenedor-filtro-quincena', 'style'),
    Output('contenedor-filtro-semana', 'style'),
    Input('modo-filtro-tiempo', 'value')
)
def controlar_visibilidad_filtros(modo):
    if modo == 'quincena': return {'display': 'block'}, {'display': 'none'}
    else: return {'display': 'none'}, {'display': 'block'}

@callback(
    Output('tabla-resumen-torre', 'data'),
    Output('tabla-resumen-torre', 'columns'),
    Output('tabla-resumen-status', 'data'),
    Output('tabla-resumen-status', 'columns'),
    Output('tabla-resumen-ejecutivo', 'data'),
    Output('tabla-resumen-ejecutivo', 'columns'),
    Output('tarjetas-kpi', 'children'),
    Input('filtro-mes', 'value'),
    Input('filtro-quincena', 'value'),
    Input('filtro-semana', 'value'),
    Input('filtro-torre', 'value'),
    Input('filtro-ejecutivo', 'value'),
    Input('btn-limpiar', 'n_clicks'),
    State('modo-filtro-tiempo', 'value')
)
def actualizar_dashboard_completo(meses, quincena, semanas, torres, ejecutivos, n_clicks, modo_tiempo):
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
        empty_df, empty_cols, tarjetas = pd.DataFrame(), [], []
        return empty_df.to_dict('records'), empty_cols, empty_df.to_dict('records'), empty_cols, empty_df.to_dict('records'), empty_cols, tarjetas
    
    def crear_tabla_dinamica(df, index_col):
        if df.empty: return pd.DataFrame(), []
        df['Fecha_Dia'] = df[COLUMNA_FECHA].dt.date
        total_general = df.groupby(index_col)[COLUMNA_ORDEN].count().to_frame('Total General')
        pivot_dia = pd.pivot_table(df, values=COLUMNA_ORDEN, index=index_col, columns='Fecha_Dia', aggfunc='count', fill_value=0)
        pivot_dia.columns = [col.strftime('%d-%m-%Y') for col in pivot_dia.columns]
        resumen_df = total_general.join(pivot_dia).fillna(0).astype(int)
        resumen_df.sort_values(by='Total General', ascending=False, inplace=True)
        resumen_df.reset_index(inplace=True)
        dia_cols = sorted([c for c in resumen_df.columns if isinstance(c, str) and '-' in c], key=lambda x: pd.to_datetime(x, format='%d-%m-%Y'))
        column_order = [index_col] + dia_cols + ['Total General']
        resumen_df = resumen_df[column_order]
        return resumen_df.to_dict('records'), [{'name': c, 'id': c} for c in resumen_df.columns]

    data_torre, cols_torre = crear_tabla_dinamica(dff, COLUMNA_TORRE)
    data_status, cols_status = crear_tabla_dinamica(dff, COLUMNA_STATUS)
    data_ejecutivo, cols_ejecutivo = crear_tabla_dinamica(dff, COLUMNA_ANALISTA)

    # --- Lógica para las Tarjetas KPI ---
    gestion_totales = dff[COLUMNA_ORDEN].count()
    total_ejecutivos = dff[COLUMNA_ANALISTA].nunique()
    gestion_fte_dia = int((gestion_totales / 6) / total_ejecutivos if total_ejecutivos > 0 else 0)
    total_corregido = dff[dff[COLUMNA_STATUS] == 'Corregido'][COLUMNA_ORDEN].count()
    tasa_resolutividad = (total_corregido / gestion_totales) if gestion_totales > 0 else 0
    gestiones_atendidas = 1 if gestion_totales > 0 else 0

    # <<< MEJORA ESTÉTICA: Función para crear tarjetas de KPI con mejor estilo >>>
    def crear_tarjeta_kpi(titulo, valor, color_valor="primary"):
        return dbc.Col(dbc.Card(dbc.CardBody([
            html.H6(titulo, className="card-title text-muted"),
            html.H3(valor, className=f"card-text text-{color_valor}")
        ]), className="shadow-sm text-center"))

    tarjetas = [
        crear_tarjeta_kpi("Gestion Totales", f"{gestion_totales}"),
        crear_tarjeta_kpi("Total Ejecutivos", f"{total_ejecutivos}"),
        crear_tarjeta_kpi("Gestiones atendidas", f"{gestiones_atendidas:.0%}", "success"),
        crear_tarjeta_kpi("Tasa de resolutividad", f"{tasa_resolutividad:.0%}", "info"),
        crear_tarjeta_kpi("Gestion FTE dia", f"{gestion_fte_dia}", "dark"),
    ]
    
    return data_torre, cols_torre, data_status, cols_status, data_ejecutivo, cols_ejecutivo, tarjetas

@callback(
    Output('filtro-mes', 'value'),
    Output('filtro-quincena', 'value'),
    Output('filtro-semana', 'value'),
    Output('filtro-torre', 'value'),
    Output('filtro-ejecutivo', 'value'),
    Output('modo-filtro-tiempo', 'value'),
    Input('btn-limpiar', 'n_clicks'),
    prevent_initial_call=True
)
def limpiar_filtros(n_clicks):
    return [], None, [], [], [], 'quincena'

# --- 4. INICIAR EL SERVIDOR ---
if __name__ == '__main__':
    app.run(debug=True)

