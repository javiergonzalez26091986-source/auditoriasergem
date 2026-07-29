import streamlit as st
import pandas as pd
import requests

# -----------------------------------------------------------------------------
# 1. CONFIGURACIÓN Y ESTILOS AVANZADOS (INCLUYE CORRECCIÓN DE BOTONES)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Auditoría SGSI - SERGEM", 
    page_icon="🛡️", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

st.markdown("""
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        /* FORZAR TEMA GRIS PROFESIONAL Y TEXTOS OSCUROS */
        .stApp, [data-testid="stAppViewContainer"] { background-color: #e9ecef !important; }
        [data-testid="stHeader"] { background-color: transparent !important; }
        .stApp p, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6, .stApp label, .stApp li {
            color: #2c3e50 !important;
        }
        
        /* CORRECCIÓN A FONDO DE LOS BOTONES: Inmunes al Dark Mode del navegador */
        div.stButton > button {
            background-color: #ffffff !important;
            color: #002b5e !important;
            border: 2px solid #002b5e !important;
            font-weight: 600 !important;
            border-radius: 6px !important;
            transition: 0.3s ease;
        }
        div.stButton > button p { color: #002b5e !important; margin: 0 !important; }
        
        /* Efecto al pasar el ratón (Hover) */
        div.stButton > button:hover {
            background-color: #002b5e !important;
            border-color: #002b5e !important;
        }
        div.stButton > button:hover p { color: #ffffff !important; }

        /* OCULTAR ELEMENTOS NATIVOS Y FIJAR SIDEBAR */
        #MainMenu {visibility: hidden;} header {visibility: hidden;} footer {visibility: hidden;}
        [data-testid="collapsedControl"] {display: none !important;}
        section[data-testid="stSidebar"] { width: 300px !important; min-width: 300px !important; background-color: #dbe2e8 !important; }
        
        /* ESTILOS DE COMPONENTES GENERALES */
        .block-container {padding-top: 0rem; padding-bottom: 0rem; max-width: 95%;}
        .navbar-custom {background-color: #002b5e; padding: 15px; margin-bottom: 20px;}
        .navbar-brand {color: #ffffff !important; font-weight: bold; font-size: 1.6rem;}
        .card-custom {border: none; border-radius: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.08); background: #ffffff; padding: 25px; margin-bottom: 20px;}
        .card-header-custom {border-bottom: 3px solid #002b5e; font-weight: 800; color: #002b5e; padding-bottom: 10px; margin-bottom: 20px; font-size: 1.4rem;}
        .pdf-frame {border: 1px dashed #cccccc; border-radius: 8px;}
    </style>
    <nav class="navbar navbar-expand-lg navbar-custom">
      <div class="container-fluid"><span class="navbar-brand">🛡️ SERGEM Mensajería S.A.S. - Portal de Auditoría SGSI 2026</span></div>
    </nav>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 2. CONEXIÓN API EN VIVO (GOOGLE APPS SCRIPT)
# -----------------------------------------------------------------------------
URL_API_DRIVE = "https://script.google.com/macros/s/AKfycbwUfREwwrhpFdQoTnFdW5KMGUlMBBHaZ9vtR-RtVgeT4OPxvXgh1Ak1_VktrvXyPGW9UA/exec"

@st.cache_data(ttl=120)
def obtener_archivos_drive():
    try:
        res = requests.get(URL_API_DRIVE)
        if res.status_code == 200:
            return pd.DataFrame(res.json())
    except:
        pass
    return pd.DataFrame()

def mostrar_visor_archivo(file_id, nombre_archivo):
    ext = nombre_archivo.split('.')[-1].lower()
    if ext in ['pdf', 'xlsx', 'xls', 'csv']:
        url = f"https://drive.google.com/file/d/{file_id}/preview" if ext == 'pdf' else f"https://docs.google.com/spreadsheets/d/{file_id}/preview"
        st.markdown(f'<iframe class="pdf-frame" src="{url}" width="100%" height="800"></iframe>', unsafe_allow_html=True)
    else:
        url = f"https://drive.google.com/file/d/{file_id}/preview"
        st.markdown(f'<iframe class="pdf-frame" src="{url}" width="100%" height="800"></iframe>', unsafe_allow_html=True)

df_archivos = obtener_archivos_drive()

if 'visor_id' not in st.session_state: st.session_state.visor_id = None
if 'visor_nombre' not in st.session_state: st.session_state.visor_nombre = None

# -----------------------------------------------------------------------------
# 3. BARRA LATERAL FIJA
# -----------------------------------------------------------------------------
st.sidebar.markdown('### 🗂️ Módulos de Evaluación')
opciones = [
    "🏠 Inicio y Sincronización", 
    "📁 Explorador Documental Completo",
    "📊 Novedades Auditoría Pasada"
]
seleccion = st.sidebar.radio("Seleccione la vista:", opciones)

# -----------------------------------------------------------------------------
# 4. LÓGICA DE VISTAS (INICIO Y EXPLORADOR)
# -----------------------------------------------------------------------------
if seleccion == "🏠 Inicio y Sincronización":
    st.markdown("""
        <div class="card-custom">
            <div class="card-header-custom">Estado del Sistema SGSI</div>
            <p>Bienvenido al portal oficial de auditoría de SERGEM Mensajería S.A.S. 
            El sistema se encuentra sincronizado con el repositorio documental en tiempo real.</p>
        </div>
    """, unsafe_allow_html=True)
    
    if st.button("🔄 Forzar Sincronización con Drive"):
        st.cache_data.clear()
        st.success("✅ Datos sincronizados correctamente.")
        
    st.info(f"Total de archivos y carpetas detectados en la nube: **{len(df_archivos)}**")

elif seleccion == "📁 Explorador Documental Completo":
    st.markdown("""
        <div class="card-custom">
            <div class="card-header-custom">Repositorio Documental (Auditoría Kreston)</div>
            <p>Seleccione una carpeta en el panel izquierdo y haga clic sobre cualquier documento para visualizarlo de inmediato.</p>
        </div>
    """, unsafe_allow_html=True)
    
    if not df_archivos.empty:
        col_explorer, col_viewer = st.columns([1, 2])
        with col_explorer:
            st.markdown("##### 📂 Estructuras Disponibles")
            rutas = sorted(df_archivos['ruta'].unique())
            for ruta in rutas:
                archivos_en_ruta = df_archivos[(df_archivos['ruta'] == ruta) & (df_archivos['tipo'] == 'Archivo')]
                if not archivos_en_ruta.empty:
                    with st.expander(f"📁 {ruta}"):
                        for _, row in archivos_en_ruta.iterrows():
                            if st.button(f"📄 {row['nombre']}", key=row['id'], use_container_width=True):
                                st.session_state.visor_id = row['id']
                                st.session_state.visor_nombre = row['nombre']
                                
        with col_viewer:
            if st.session_state.visor_id:
                st.markdown(f"**Documento Seleccionado:** `{st.session_state.visor_nombre}`")
                mostrar_visor_archivo(st.session_state.visor_id, st.session_state.visor_nombre)
            else:
                st.info("👈 Seleccione un documento en el panel izquierdo para previsualizarlo en pantalla.")
    else:
        st.error("No se encontraron archivos en la sincronización.")

# -----------------------------------------------------------------------------
# 5. MÓDULO CORREGIDO: NOVEDADES AUDITORÍA PASADA LEYENDO DESDE DRIVE
# -----------------------------------------------------------------------------
elif seleccion == "📊 Novedades Auditoría Pasada":
    st.markdown("""
        <div class="card-custom">
            <div class="card-header-custom">Hallazgos y Novedades Auditoría 2025</div>
            <p>Resumen interactivo de las observaciones pasadas, indicando claramente qué aspectos fueron subsanados y cuáles siguen pendientes.</p>
        </div>
    """, unsafe_allow_html=True)

    @st.cache_data(ttl=300)
    def cargar_matriz_observaciones(df_archivos_nube):
        if df_archivos_nube.empty:
            return pd.DataFrame()
            
        # 1. Buscamos el archivo de matriz en la lista que trajo la API de Drive
        match = df_archivos_nube[df_archivos_nube['nombre'].str.contains("Matriz de Observaciones", case=False, na=False)]
        
        if match.empty:
            return pd.DataFrame() # No se encontró el archivo
            
        # 2. Obtenemos su ID y creamos un link de descarga directa
        file_id = match.iloc[0]['id']
        url_descarga = f"https://drive.google.com/uc?export=download&id={file_id}"
        
        try:
            # 3. Leemos directo desde la URL. header=15 equivale a la fila 16 de tu Excel
            df = pd.read_excel(url_descarga, sheet_name='AÑO', header=15)
            df = df.iloc[1:].copy() # Saltamos la fila donde dice "DD MM AA"
            
            # 4. Extraemos y renombramos solo las columnas necesarias
            columnas = ['NOMBRE DE INFORME O AUDITORIA', 'COMPONENTE', 'OBSERVACIÓN', 'ESTADO', 'TIPO', 'COMO FUE SUBSANADO (ACTIVIDAD REALIZADA)']
            df_clean = df[columnas].copy()
            
            df_clean = df_clean.rename(columns={
                'NOMBRE DE INFORME O AUDITORIA': 'Informe',
                'COMPONENTE': 'Componente',
                'OBSERVACIÓN': 'Observación',
                'ESTADO': 'Estado',
                'TIPO': 'Tipo',
                'COMO FUE SUBSANADO (ACTIVIDAD REALIZADA)': 'Subsanación (Actividad)'
            })
            
            df_clean = df_clean.dropna(subset=['Observación'])
            df_clean['Estado'] = df_clean['Estado'].fillna('SIN ESTADO')
            return df_clean
        except Exception as e:
            st.error(f"Error interno al procesar el Excel: {e}")
            return pd.DataFrame()

    df_nov = cargar_matriz_observaciones(df_archivos)

    if not df_nov.empty:
        conteo_estados = df_nov['Estado'].value_counts()
        
        st.markdown("### 📈 Resumen General de Hallazgos")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Observaciones", len(df_nov))
        col2.metric("✅ Subsanadas", conteo_estados.get('SUBSANADA', 0))
        col3.metric("⚠️ NO Subsanadas", conteo_estados.get('NO SUBSANADA', 0))
        col4.metric("🔒 Cerradas", conteo_estados.get('CERRADO', 0))

        st.divider()
        st.markdown("### 🔍 Detalle Interactivo de Observaciones")
        
        filtro = st.selectbox("Filtrar estado de la novedad:", ["Todos los Estados"] + list(df_nov['Estado'].unique()))
        df_mostrar = df_nov if filtro == "Todos los Estados" else df_nov[df_nov['Estado'] == filtro]

        for _, row in df_mostrar.iterrows():
            emoji = "✅" if row['Estado'] == "SUBSANADA" else "⚠️" if row['Estado'] == "NO SUBSANADA" else "🔒"
            with st.expander(f"{emoji} {row['Componente']} - Estado: {row['Estado']}"):
                st.markdown("**📌 Observación Original:**")
                st.info(row['Observación'])
                st.markdown("**🛠️ Actividad Realizada / Cómo fue subsanado:**")
                
                if pd.notna(row['Subsanación (Actividad)']) and str(row['Subsanación (Actividad)']).strip() != "":
                    st.success(row['Subsanación (Actividad)'])
                else:
                    st.warning("Aún no hay actividad de subsanación registrada en el archivo.")
    else:
        st.error("No se pudo cargar el archivo Excel. Verifica que el archivo de la Matriz esté alojado en tu Google Drive junto a los demás documentos y que el script de Apps Script lo esté detectando.")
