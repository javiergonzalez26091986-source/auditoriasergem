import streamlit as st
import pandas as pd
import requests

# -----------------------------------------------------------------------------
# 1. CONFIGURACIÓN Y ESTILOS AVANZADOS (MENÚ LATERAL FIJO)
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
        /* Ocultar elementos nativos innecesarios */
        #MainMenu {visibility: hidden;} 
        header {visibility: hidden;} 
        footer {visibility: hidden;}
        
        /* FIJAR EL MENÚ LATERAL Y OCULTAR EL BOTÓN DE COLAPSO (FLECHITAS) */
        [data-testid="collapsedControl"] {display: none !important;}
        section[data-testid="stSidebar"] {
            width: 300px !important;
            min-width: 300px !important;
            background-color: #f0f2f6;
        }
        
        /* Estilos generales del portal */
        .block-container {padding-top: 0rem; padding-bottom: 0rem; max-width: 95%;}
        body {background-color: #f8f9fa;}
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
    """Renderiza de forma inteligente según el formato del archivo"""
    ext = nombre_archivo.split('.')[-1].lower()
    
    if ext == 'pdf':
        url = f"https://drive.google.com/file/d/{file_id}/preview"
        st.markdown(f'<iframe class="pdf-frame" src="{url}" width="100%" height="800"></iframe>', unsafe_allow_html=True)
    elif ext in ['xlsx', 'xls', 'csv']:
        url = f"https://docs.google.com/spreadsheets/d/{file_id}/preview"
        st.markdown(f'<iframe class="pdf-frame" src="{url}" width="100%" height="800"></iframe>', unsafe_allow_html=True)
    else:
        # Para imágenes u otros formatos
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
    "📁 Explorador Documental Completo"
]
seleccion = st.sidebar.radio("Seleccione la vista:", opciones)

# -----------------------------------------------------------------------------
# 4. LÓGICA DE VISTAS
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
