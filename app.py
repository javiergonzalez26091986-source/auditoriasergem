import streamlit as st
import pandas as pd
import requests

# -----------------------------------------------------------------------------
# 1. CONFIGURACIÓN Y ESTILOS AVANZADOS (MENÚ LATERAL FIJO Y FORZAR TEMA)
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
        /* =========================================================
           FORZAR TEMA GRIS PROFESIONAL (Ignora Dark Mode del navegador)
           ========================================================= */
        .stApp, [data-testid="stAppViewContainer"] {
            background-color: #e9ecef !important; /* Gris profesional claro */
        }
        [data-testid="stHeader"] {
            background-color: transparent !important;
        }
        /* Forzar texto oscuro para contrastar con el fondo gris */
        .stApp p, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6, .stApp label {
            color: #2c3e50 !important;
        }
        
        /* Ocultar elementos nativos innecesarios */
        #MainMenu {visibility: hidden;} 
        header {visibility: hidden;} 
        footer {visibility: hidden;}
        
        /* FIJAR EL MENÚ LATERAL Y OCULTAR EL BOTÓN DE COLAPSO (FLECHITAS) */
        [data-testid="collapsedControl"] {display: none !important;}
        section[data-testid="stSidebar"] {
            width: 300px !important;
            min-width: 300px !important;
            background-color: #dbe2e8 !important; /* Gris un poco más oscuro para contraste */
        }
        
        /* Estilos generales del portal */
        .block-container {padding-top: 0rem; padding-bottom: 0rem; max-width: 95%;}
        .navbar-custom {background-color: #002b5e; padding: 15px; margin-bottom: 20px;}
        .navbar-brand {color: #ffffff !important; font-weight: bold; font-size: 1.6rem;}
        .card-custom {border: none; border-radius: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.08); background: #ffffff; padding: 25px; margin-bottom: 20px;}
        .card-header-custom {border-bottom: 3px solid #002b5e; font-weight: 800; color: #002b5e; padding-bottom: 10px; margin-bottom: 20px; font-size: 1.4rem;}
        .pdf-frame {border: 1px dashed #cccccc; border-radius: 8px;}
        
        /* Ajuste para que las tarjetas blancas no pierdan su fondo con el código !important de arriba */
        .card-custom p, .card-custom div { color: #333333 !important; }
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
    "📁 Explorador Documental Completo",
    "📊 Novedades Auditoría Pasada"  # <--- NUEVO MÓDULO AGREGADO AQUÍ
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

# -----------------------------------------------------------------------------
# 5. NUEVO MÓDULO: NOVEDADES AUDITORÍA PASADA
# -----------------------------------------------------------------------------
elif seleccion == "📊 Novedades Auditoría Pasada":
    st.markdown("""
        <div class="card-custom">
            <div class="card-header-custom">Hallazgos y Novedades Auditoría 2025</div>
            <p>Resumen interactivo de las observaciones pasadas, indicando claramente qué aspectos fueron subsanados y cuáles siguen pendientes de solución.</p>
        </div>
    """, unsafe_allow_html=True)

    @st.cache_data
    def cargar_matriz_observaciones():
        nombre_archivo = "RM-4278-25-Matriz de Observaciones Aud. Control Interno Legal-SERGEM MENSAJERIA S.A.S (1) 2025.xlsx"
        try:
            # Lee la matriz saltando hasta la fila de encabezados reales de la plantilla
            df = pd.read_excel(nombre_archivo, sheet_name='AÑO', header=14)
            df = df.iloc[1:].copy() # Saltar fila temporal (DD MM AA)
            
            # Mapeo y selección de las columnas clave
            col_map = {
                'Unnamed: 1': 'Informe',
                'Unnamed: 6': 'Componente',
                'Unnamed: 7': 'Observación',
                'Unnamed: 12': 'Estado',
                'Unnamed: 13': 'Tipo',
                'Unnamed: 18': 'Subsanación (Actividad)'
            }
            df = df.rename(columns=col_map)
            df_clean = df[['Informe', 'Componente', 'Observación', 'Estado', 'Tipo', 'Subsanación (Actividad)']].dropna(subset=['Observación'])
            df_clean['Estado'] = df_clean['Estado'].fillna('SIN ESTADO')
            return df_clean
        except Exception as e:
            return pd.DataFrame()

    df_nov = cargar_matriz_observaciones()

    if not df_nov.empty:
        # Extraer conteos para los indicadores de resumen
        conteo_estados = df_nov['Estado'].value_counts()
        total_hallazgos = len(df_nov)
        subsanadas = conteo_estados.get('SUBSANADA', 0)
        no_subsanadas = conteo_estados.get('NO SUBSANADA', 0)
        cerradas = conteo_estados.get('CERRADO', 0)

        # 1. Mostrar Tablero Resumido (Indicadores)
        st.markdown("### 📈 Resumen General de Hallazgos")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Observaciones", total_hallazgos)
        col2.metric("✅ Subsanadas", subsanadas)
        col3.metric("⚠️ NO Subsanadas", no_subsanadas)
        col4.metric("🔒 Cerradas", cerradas)

        st.divider()

        # 2. Panel Interactivo de Búsqueda y Filtrado
        st.markdown("### 🔍 Detalle Interactivo de Observaciones")
        
        filtro = st.selectbox("Filtrar estado de la novedad:", ["Todos los Estados"] + list(df_nov['Estado'].unique()))
        
        if filtro != "Todos los Estados":
            df_mostrar = df_nov[df_nov['Estado'] == filtro]
        else:
            df_mostrar = df_nov

        # 3. Mostrar las novedades en acordeones para no saturar la pantalla
        for _, row in df_mostrar.iterrows():
            # Asignar un emoji según el estado
            emoji = "✅" if row['Estado'] == "SUBSANADA" else "⚠️" if row['Estado'] == "NO SUBSANADA" else "🔒"
            
            with st.expander(f"{emoji} {row['Componente']} - Estado: {row['Estado']}"):
                st.markdown(f"**📌 Observación Original:**")
                st.info(row['Observación'])
                st.markdown(f"**🛠️ Actividad Realizada / Cómo fue subsanado:**")
                
                texto_subsanacion = row['Subsanación (Actividad)']
                if pd.notna(texto_subsanacion) and str(texto_subsanacion).strip() != "":
                    st.success(texto_subsanacion)
                else:
                    st.warning("Aún no hay actividad de subsanación registrada en el archivo.")
                    
        # 4. Mensaje de Preparación para Dashboard (A futuro)
        st.markdown("""
            <br><hr>
            <div style="background-color: #dbe2e8; padding: 15px; border-left: 5px solid #002b5e; border-radius: 5px;">
                <small><i>💡 <b>Módulo Dashboard Ready:</b> Los datos ya están limpios y agrupados en memoria. Cuando inicies el desarrollo del módulo de Dashboard, podrás pasar directamente estos DataFrames a librerías como Plotly o Altair para generar gráficos de barras y tortas mostrando la evolución de los hallazgos.</i></small>
            </div>
        """, unsafe_allow_html=True)
        
    else:
        st.error("No se pudo cargar el archivo Excel, o no contiene observaciones válidas. Verifica que el archivo `RM-4278-25-Matriz de Observaciones Aud. Control Interno Legal-SERGEM MENSAJERIA S.A.S (1) 2025.xlsx` exista en la misma carpeta que este script de Python.")
