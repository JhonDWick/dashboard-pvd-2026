import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as gg
import os
import io

# -----------------------------------------------------------------------------
# CONFIGURACIÓN DE LA PÁGINA Y ESTILOS VISUALES INSTITUCIONALES
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="PVD - Control y Monitoreo de Mantenimiento Vial - Año Fiscal 2026",
    page_icon="🛣️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS personalizados inspirados en el diseño corporativo
st.markdown("""
<style>
    /* Estructura general */
    .main {
        background-color: #F8F9FA;
    }
    
    /* Header Principal */
    .header-container {
        background: linear-gradient(135deg, #2E1A47 0%, #4A2E80 100%);
        padding: 24px 30px;
        border-radius: 12px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }
    .header-container h1 {
        margin: 0;
        font-size: 28px;
        font-weight: 700;
        color: #FFFFFF;
    }
    .header-container p {
        margin: 5px 0 0 0;
        font-size: 14px;
        color: #D1C4E9;
    }

    /* Cards KPI */
    .kpi-card {
        background-color: #FFFFFF;
        border-radius: 10px;
        padding: 18px;
        border-left: 5px solid #4A2E80;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        text-align: left;
        margin-bottom: 10px;
    }
    .kpi-title {
        font-size: 12px;
        text-transform: uppercase;
        color: #6C757D;
        font-weight: 600;
        letter-spacing: 0.5px;
    }
    .kpi-value {
        font-size: 26px;
        font-weight: 800;
        color: #2E1A47;
        margin: 4px 0;
    }
    .kpi-subtext {
        font-size: 11px;
        color: #28A745;
        font-weight: 600;
    }

    /* Alertas */
    .alert-card-warning {
        background-color: #FFF3CD;
        border: 1px solid #FFEBAA;
        border-left: 5px solid #FFC107;
        border-radius: 8px;
        padding: 14px;
        color: #856404;
        margin-bottom: 15px;
    }
    .alert-card-danger {
        background-color: #F8D7DA;
        border: 1px solid #F5C6CB;
        border-left: 5px solid #DC3545;
        border-radius: 8px;
        padding: 14px;
        color: #721C24;
        margin-bottom: 15px;
    }
    .alert-card-success {
        background-color: #D4EDDA;
        border: 1px solid #C3E6CB;
        border-left: 5px solid #28A745;
        border-radius: 8px;
        padding: 14px;
        color: #155724;
        margin-bottom: 15px;
    }

    /* Tabs estilizadas */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #E9ECEF;
        border-radius: 6px 6px 0 0;
        padding: 10px 20px;
        color: #495057;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background-color: #4A2E80 !important;
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# RUTAS Y CARGA DE DATOS (EXCEL LOCAL + LOGO + SUBIDA ALTERNATIVA)
# -----------------------------------------------------------------------------
RUTA_PREDETERMINADA = r"C:\Users\ING. YONATHAN JESUS\Desktop\BASE DE DATOS\3. CODIGOS\seguimiento_rutinario_2026.xlsx"
LOGO_PATH = r"C:\Users\ING. YONATHAN JESUS\Desktop\BASE DE DATOS\3. CODIGOS\PVD.png"

# Centrar logo en la barra lateral usando columnas
c1, c2, c3 = st.sidebar.columns([1, 2, 1])
with c2:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, use_container_width=True)
    elif os.path.exists("PVD.png"):
        st.image("PVD.png", use_container_width=True)

# Texto de autoría resaltado (más grande, oscuro y centrado)
st.sidebar.markdown(
    """
    <p style='text-align: center; font-size: 13px; color: #1E1E1E; font-weight: 700; margin-top: 5px; margin-bottom: 25px;'>
        Elaborado por:<br>
        <span style='color: #2E1A47; font-size: 15px;'>Ing. Yonathan J. Monteza Garrido</span>
    </p>
    """,
    unsafe_allow_html=True
)

st.sidebar.title("🛠️ Configuración de Datos")

modo_carga = st.sidebar.radio("Origen de datos:", ["Ruta Predeterminada en Servidor/PC", "Subir Archivo Excel Manualmente"])

@st.cache_data(ttl=60)
def cargar_datos_excel(ruta_o_file):
    try:
        df = pd.read_excel(ruta_o_file, sheet_name="Dashboard")
        df.columns = df.columns.str.strip()
        
        for col in ['Avance Físico (%)', 'Avance Finan. (%)']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                if df[col].max() > 1.0:
                    df[col] = df[col] / 100.0
                    
        columnas_texto = ['Departamento', 'Provincia', 'Distrito', 'Tipo Mantenimiento', 
                          'Tipo de Red Vial', 'Convenio', 'FITSA Aprobado', 'Estado Actual']
        for col in columnas_texto:
            if col in df.columns:
                df[col] = df[col].fillna("Sin Información").astype(str).str.strip()
                
        return df
    except Exception as e:
        st.error(f"Error al leer el archivo Excel: {e}")
        return None

df_raw = None

if modo_carga == "Ruta Predeterminada en Servidor/PC":
    if os.path.exists(RUTA_PREDETERMINADA):
        df_raw = cargar_datos_excel(RUTA_PREDETERMINADA)
    else:
        st.sidebar.warning(f"No se encontró el archivo en la ruta local:\n`{RUTA_PREDETERMINADA}`")
        st.sidebar.info("Carga el archivo mediante la opción manual abajo.")
        archivo_subido = st.sidebar.file_uploader("Subir archivo .xlsx", type=["xlsx"])
        if archivo_subido:
            df_raw = cargar_datos_excel(archivo_subido)
else:
    archivo_subido = st.sidebar.file_uploader("Subir archivo .xlsx", type=["xlsx"])
    if archivo_subido:
        df_raw = cargar_datos_excel(archivo_subido)

if df_raw is None or df_raw.empty:
    st.info("👈 Por favor selecciona o sube un archivo `seguimiento_rutinario_2026.xlsx` válido en el panel izquierdo.")
    st.stop()

# -----------------------------------------------------------------------------
# FILTROS DINÁMICOS MULTI-REGIÓN
# -----------------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.header("🎯 Filtros de Control")

# 1. Departamento / Región
deptos_disponibles = sorted(df_raw['Departamento'].unique().tolist())
deptos_defecto = [d for d in ["Huancavelica", "Tacna", "Moquegua"] if d in deptos_disponibles]
if not deptos_defecto:
    deptos_defecto = deptos_disponibles

sel_deptos = st.sidebar.multiselect(
    "Departamento / Región:", 
    options=deptos_disponibles, 
    default=deptos_defecto,
    placeholder="Seleccione departamento..."
)

df_filtered = df_raw[df_raw['Departamento'].isin(sel_deptos)] if sel_deptos else df_raw.copy()

# 2. Provincia
prov_disponibles = sorted(df_filtered['Provincia'].unique().tolist())
sel_prov = st.sidebar.multiselect(
    "Provincia:", 
    options=prov_disponibles,
    placeholder="Seleccione provincia..."
)
if sel_prov:
    df_filtered = df_filtered[df_filtered['Provincia'].isin(sel_prov)]

# 3. Distrito
dist_disponibles = sorted(df_filtered['Distrito'].unique().tolist())
sel_dist = st.sidebar.multiselect(
    "Distrito:", 
    options=dist_disponibles,
    placeholder="Seleccione distrito..."
)
if sel_dist:
    df_filtered = df_filtered[df_filtered['Distrito'].isin(sel_dist)]

# 4. Tipo de Mantenimiento
mant_disponibles = sorted(df_filtered['Tipo Mantenimiento'].unique().tolist())
sel_mant = st.sidebar.multiselect(
    "Tipo Mantenimiento:", 
    options=mant_disponibles,
    placeholder="Seleccione tipo..."
)
if sel_mant:
    df_filtered = df_filtered[df_filtered['Tipo Mantenimiento'].isin(sel_mant)]

# 5. Red Vial
red_disponibles = sorted(df_filtered['Tipo de Red Vial'].unique().tolist())
sel_red = st.sidebar.multiselect(
    "Red Vial:", 
    options=red_disponibles,
    placeholder="Seleccione red vial..."
)
if sel_red:
    df_filtered = df_filtered[df_filtered['Tipo de Red Vial'].isin(sel_red)]

# 6. Estado FITSA
fitsa_disponibles = sorted(df_filtered['FITSA Aprobado'].unique().tolist())
sel_fitsa = st.sidebar.multiselect(
    "FITSA Aprobado:", 
    options=fitsa_disponibles,
    placeholder="Seleccione estado FITSA..."
)
if sel_fitsa:
    df_filtered = df_filtered[df_filtered['FITSA Aprobado'].isin(sel_fitsa)]

# 7. Estado Actual
estado_disponibles = sorted(df_filtered['Estado Actual'].unique().tolist())
sel_estado = st.sidebar.multiselect(
    "Estado Actual del Servicio:", 
    options=estado_disponibles,
    placeholder="Seleccione estado..."
)
if sel_estado:
    df_filtered = df_filtered[df_filtered['Estado Actual'].isin(sel_estado)]

# -----------------------------------------------------------------------------
# CABECERA Y RESUMEN
# -----------------------------------------------------------------------------
st.markdown("""
<div class="header-container">
    <h1>🛣️ Control de Mantenimiento Vial - Año Fiscal 2026</h1>
    <p>Provías Descentralizado | Monitoreo Operativo, Presupuestal y Ambiental</p>
</div>
""", unsafe_allow_html=True)

# CALCULOS KPI
total_tramos = len(df_filtered)
avg_fisico = df_filtered['Avance Físico (%)'].mean() if total_tramos > 0 else 0
avg_finan = df_filtered['Avance Finan. (%)'].mean() if total_tramos > 0 else 0

fitsa_si = len(df_filtered[df_filtered['FITSA Aprobado'].str.upper() == 'SI'])
fitsa_pendiente = len(df_filtered[df_filtered['FITSA Aprobado'].str.upper() == 'PENDIENTE'])
fitsa_na = len(df_filtered[df_filtered['FITSA Aprobado'].str.upper().isin(['N/A', 'NO APLICA'])])

en_ejecucion = len(df_filtered[df_filtered['Estado Actual'].str.lower() == 'en ejecución'])

# RENDERING KPIS
col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">Total Intervenciones</div>
        <div class="kpi-value">{total_tramos}</div>
        <div class="kpi-subtext">📍 Tramos Monitoreados</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">En Ejecución</div>
        <div class="kpi-value">{en_ejecucion}</div>
        <div class="kpi-subtext" style="color: #17A2B8;">🟢 Servicios Activos</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">Avance Físico Prom.</div>
        <div class="kpi-value">{avg_fisico*100:.2f}%</div>
        <div class="kpi-subtext" style="color: #6F42C1;">📊 Promedio General</div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">Avance Finan. Prom.</div>
        <div class="kpi-value">{avg_finan*100:.2f}%</div>
        <div class="kpi-subtext" style="color: #28A745;">💰 Ejercicio Presupuestal</div>
    </div>
    """, unsafe_allow_html=True)

with col5:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-title">Trámite FITSA</div>
        <div class="kpi-value">{fitsa_si} / {fitsa_si + fitsa_pendiente}</div>
        <div class="kpi-subtext" style="color: {'#DC3545' if fitsa_pendiente > 0 else '#28A745'};">
            ⚠️ {fitsa_pendiente} Pendientes
        </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# MONITOREO DE ALERTAS SOCIOAMBIENTALES (FITSA)
# -----------------------------------------------------------------------------
if fitsa_pendiente > 0:
    st.markdown(f"""
    <div class="alert-card-danger">
        <strong>⚠️ Atención - Alerta de Instrumento Ambiental (FITSA):</strong> 
        Se han detectado <b>{fitsa_pendiente} trámites/servicios</b> con estado <b>PENDIENTE</b> en su Ficha Técnica Socioambiental (FITSA). 
        Recordar que el FITSA aprobado es requisito indispensable para dar inicio a las actividades de mantenimiento periódico.
    </div>
    """, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# PESTAÑAS PRINCIPALES DE ANÁLISIS
# -----------------------------------------------------------------------------
tab_resumen, tab_avances, tab_fitsa, tab_tabla = st.tabs([
    "📈 Resumen General", 
    "📊 Comparativo Físico vs Financiero", 
    "🌿 Control FITSA & Red Vial", 
    "📋 Matriz Detallada de Datos"
])

# -----------------------------------------------------------------------------
# TAB 1: RESUMEN GENERAL
# -----------------------------------------------------------------------------
with tab_resumen:
    c_graph1, c_graph2 = st.columns(2)
    
    with c_graph1:
        st.subheader("📌 Distribución por Estado Actual del Servicio")
        df_estado = df_filtered['Estado Actual'].value_counts().reset_index()
        df_estado.columns = ['Estado Actual', 'Cantidad']
        
        fig_estado = px.pie(
            df_estado, 
            names='Estado Actual', 
            values='Cantidad',
            hole=0.45,
            color_discrete_sequence=px.colors.qualitative.Prism
        )
        fig_estado.update_traces(textinfo='percent+label+value')
        fig_estado.update_layout(margin=dict(t=20, b=20, l=20, r=20), height=350)
        st.plotly_chart(fig_estado, use_container_width=True)

    with c_graph2:
        st.subheader("🏛️ Servicios por Departamento / Región")
        df_dep = df_filtered.groupby(['Departamento', 'Tipo de Red Vial']).size().reset_index(name='Cantidad')
        
        fig_dep = px.bar(
            df_dep, 
            x='Departamento', 
            y='Cantidad', 
            color='Tipo de Red Vial',
            barmode='stack',
            color_discrete_map={'Vecinal': '#4A2E80', 'Departamental': '#00A896', 'Sin Información': '#A8A8A8'}
        )
        fig_dep.update_layout(margin=dict(t=20, b=20, l=20, r=20), height=350)
        st.plotly_chart(fig_dep, use_container_width=True)

    st.markdown("---")
    
    st.subheader("📍 Resumen de Intervenciones por Provincia")
    df_prov_summary = df_filtered.groupby('Provincia').agg(
        Total_Tramos=('Código de Vía', 'count'),
        Prom_Fisico=('Avance Físico (%)', 'mean'),
        Prom_Financiero=('Avance Finan. (%)', 'mean')
    ).reset_index()
    
    df_prov_summary['Prom_Fisico'] = (df_prov_summary['Prom_Fisico'] * 100).round(2).astype(str) + '%'
    df_prov_summary['Prom_Financiero'] = (df_prov_summary['Prom_Financiero'] * 100).round(2).astype(str) + '%'
    
    st.dataframe(df_prov_summary, use_container_width=True, hide_index=True)

# -----------------------------------------------------------------------------
# TAB 2: COMPARATIVO FÍSICO VS FINANCIERO (DISEÑO SUAVE Y DESCANSO VISUAL)
# -----------------------------------------------------------------------------
with tab_avances:
    st.subheader("📊 Comparación de Avance Físico (%) vs. Avance Financiero (%) por Tramo")
    
    df_ejec = df_filtered[df_filtered['Avance Físico (%)'] > 0].copy()
    
    if not df_ejec.empty:
        df_ejec['Avance Físico %'] = (df_ejec['Avance Físico (%)'] * 100).round(2)
        df_ejec['Avance Financiero %'] = (df_ejec['Avance Finan. (%)'] * 100).round(2)
        
        # --- BUSCADOR Y FILTROS DE TRAMOS ---
        col_search1, col_search2 = st.columns([2, 1])
        
        with col_search1:
            tramos_lista = sorted(df_ejec['Código de Vía'].unique().tolist())
            tramos_seleccionados = st.multiselect(
                "🔍 Buscar o filtrar por código(s) de vía específico(s):",
                options=tramos_lista,
                placeholder="Escriba o seleccione código de vía (ej: MO-524, R1801290)..."
            )
            
        with col_search2:
            max_tramos = len(df_ejec)
            if not tramos_seleccionados:
                top_n = st.slider(
                    "Tramos a mostrar (de mayor avance a menor):", 
                    min_value=3, 
                    max_value=min(30, max_tramos), 
                    value=min(8, max_tramos)
                )
        
        # Aplicar filtro de búsqueda o corte de slider
        if tramos_seleccionados:
            df_chart = df_ejec[df_ejec['Código de Vía'].isin(tramos_seleccionados)].copy()
        else:
            df_chart = df_ejec.sort_values(by='Avance Físico %', ascending=False).head(top_n).copy()
        
        # --- CONSTRUCCIÓN DEL GRÁFICO CON ALTO CONFORT VISUAL ---
        fig_comp = gg.Figure()
        
        # Avance Físico Real - Púrpura Institucional PVD
        fig_comp.add_trace(gg.Bar(
            y=df_chart['Código de Vía'],
            x=df_chart['Avance Físico %'],
            name='Avance Físico Real (%)',
            orientation='h',
            marker=dict(color='#4A2E80', cornerradius=4), # Púrpura elegante con bordes redondeados
            text=df_chart['Avance Físico %'].astype(str) + '%',
            textposition='outside',
            textfont=dict(color='#2E1A47', size=11, family='sans-serif')
        ))
        
        # Avance Financiero - Verde Menta Suave Desaturado (Agradable y diferenciado)
        fig_comp.add_trace(gg.Bar(
            y=df_chart['Código de Vía'],
            x=df_chart['Avance Financiero %'],
            name='Avance Financiero (%)',
            orientation='h',
            marker=dict(color='#58D68D', cornerradius=4), # Verde claro desaturado descansado
            text=df_chart['Avance Financiero %'].astype(str) + '%',
            textposition='outside',
            textfont=dict(color='#1E8449', size=11, family='sans-serif')
        ))
        
        fig_comp.update_layout(
            barmode='group',
            bargap=0.25,      # Separación entre grupos de tramos
            bargroupgap=0.1,  # Separación entre la barra física y financiera
            xaxis_title="Porcentaje (%)",
            yaxis_title="Código de Vía",
            height=max(400, len(df_chart) * 55), # Espaciado vertical amplio por tramo
            margin=dict(t=30, b=20, l=120, r=50),
            legend=dict(
                orientation="h", 
                yanchor="bottom", 
                y=1.02, 
                xanchor="right", 
                x=1,
                font=dict(size=12, color="#2C3E50")
            ),
            yaxis=dict(autorange="reversed"),
            xaxis=dict(range=[0, 115], gridcolor='#F0F0F0', zeroline=False), # Rango para dar espacio a las etiquetas externas
            plot_bgcolor='#FFFFFF'
        )
        st.plotly_chart(fig_comp, use_container_width=True)
    else:
        st.info("No se registraron tramos con avance mayor al 0% en la selección actual.")

# -----------------------------------------------------------------------------
# TAB 3: CONTROL FITSA (SOLO MANTENIMIENTO PERIÓDICO) & RED VIAL
# -----------------------------------------------------------------------------
with tab_fitsa:
    col_f1, col_f2 = st.columns(2)
    
    with col_f1:
        st.subheader("🌿 Estado de FITSA (Solo Mantenimientos Periódicos)")
        
        # Filtrar el DataFrame excluyendo Mantenimiento Rutinario
        df_periodico = df_filtered[
            df_filtered['Tipo Mantenimiento'].str.lower().str.contains('periódico|periodico', na=False)
        ]
        
        if not df_periodico.empty:
            df_fitsa_chart = df_periodico['FITSA Aprobado'].value_counts().reset_index()
            df_fitsa_chart.columns = ['FITSA Aprobado', 'Cantidad']
            
            fig_fitsa = px.pie(
                df_fitsa_chart, 
                names='FITSA Aprobado', 
                values='Cantidad',
                color='FITSA Aprobado',
                hole=0.35, # Formato dona elegante
                color_discrete_map={
                    'Si': '#28A745', 
                    'SI': '#28A745', 
                    'Pendiente': '#DC3545', 
                    'PENDIENTE': '#DC3545', 
                    'N/A': '#95A5A6', 
                    'No Aplica': '#95A5A6'
                }
            )
            fig_fitsa.update_traces(
                textinfo='value+percent', 
                textfont_size=12
            )
            fig_fitsa.update_layout(
                height=350, 
                margin=dict(t=20, b=20, l=20, r=20),
                legend=dict(orientation="h", yanchor="bottom", y=-0.15, xanchor="center", x=0.5)
            )
            st.plotly_chart(fig_fitsa, use_container_width=True)
        else:
            st.info("No se encontraron registros de Mantenimiento Periódico en la selección actual.")

    with col_f2:
        st.subheader("KM Distribución por Tipo de Red Vial")
        df_red_chart = df_filtered['Tipo de Red Vial'].value_counts().reset_index()
        df_red_chart.columns = ['Tipo de Red Vial', 'Cantidad']
        
        fig_red = px.bar(
            df_red_chart,
            x='Tipo de Red Vial',
            y='Cantidad',
            color='Tipo de Red Vial',
            text='Cantidad',
            color_discrete_map={'Vecinal': '#4A2E80', 'Departamental': '#17A2B8'}
        )
        fig_red.update_traces(textposition='outside')
        fig_red.update_layout(
            height=350, 
            margin=dict(t=20, b=20, l=20, r=20),
            yaxis=dict(range=[0, df_red_chart['Cantidad'].max() * 1.15])
        )
        st.plotly_chart(fig_red, use_container_width=True)

# -----------------------------------------------------------------------------
# TAB 4: MATRIZ DETALLADA DE DATOS Y EXPORTACIÓN
# -----------------------------------------------------------------------------
with tab_tabla:
    st.subheader("📋 Matriz Consolidada de Mantenimientos Viales")
    
    df_display = df_filtered.copy()
    
    df_display['Avance Físico (%)'] = (df_display['Avance Físico (%)'] * 100).round(2).astype(str) + '%'
    df_display['Avance Finan. (%)'] = (df_display['Avance Finan. (%)'] * 100).round(2).astype(str) + '%'
    
    st.dataframe(df_display, use_container_width=True, hide_index=True)
    
    st.markdown("### 💾 Exportar Reporte Filtrado")
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_filtered.to_excel(writer, sheet_name='Reporte_PVD', index=False)
        
    st.download_button(
        label="📥 Descargar Reporte en Excel (.xlsx)",
        data=buffer.getvalue(),
        file_name="Reporte_Mantenimiento_Vial_PVD.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# -----------------------------------------------------------------------------
# PIE DE PÁGINA INSTITUCIONAL
# -----------------------------------------------------------------------------
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #6C757D; font-size: 12px;'>"
    "Sistema de Monitoreo de Mantenimientos Viales Rutinarios y Periódicos | Provías Descentralizado © 2026"
    "</div>", 
    unsafe_allow_html=True
)