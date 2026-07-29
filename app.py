import streamlit as st
import pandas as pd
import requests

# -----------------------------------------------------------------------------
# 1. CONFIGURACIÓN DE LA PÁGINA (STREAMLIT)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Auditoría SGSI - SERGEM",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -----------------------------------------------------------------------------
# 2. INYECCIÓN DE HTML5, CSS3 Y BOOTSTRAP 5
# -----------------------------------------------------------------------------
st.markdown("""
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        /* Ocultar UI nativa de Streamlit */
        #MainMenu {visibility: hidden;}
        header {visibility: hidden;}
        footer {visibility: hidden;}
        
        /* Ajustar contenedor principal */
        .block-container {padding-top: 0rem; padding-bottom: 0rem; max-width: 95%;}
        body {background-color: #f8f9fa;}
        
        /* Navbar Corporativa */
        .navbar-custom {background-color: #002b5e; padding: 15px;}
        .navbar-brand {color: #ffffff !important; font-weight: bold; font-size: 1.6rem; margin-left: 15px;}
        
        /* Tarjetas de Módulos (Cards) */
        .card-custom {
            border: none; 
            border-radius: 10px; 
            box-shadow: 0 4px 10px rgba(0,0,0,0.08); 
            background: #ffffff; 
            margin-bottom: 2rem; 
            padding: 25px;
        }
        .card-header-custom {
            background-color: #ffffff; 
            border-bottom: 3px solid #002b5e; 
            font-weight: 800; 
            color: #002b5e; 
            padding-bottom: 10px; 
            margin-bottom: 20px; 
            font-size: 1.4rem;
        }
        
        /* iframe del PDF */
        .pdf-frame {border: 1px dashed #cccccc; border-radius: 8px;}
    </style>
    
    <!-- Navbar HTML -->
    <nav class="navbar navbar-expand-lg navbar-custom">
      <div class="container-fluid">
        <a class="navbar-brand" href="#">🛡️ SERGEM Mensajería S.A.S. - Portal de Auditoría SGSI 2026</a>
      </div>
    </nav>
    <br>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 3. CONEXIÓN API EN VIVO (GOOGLE APPS SCRIPT -> DRIVE)
# -----------------------------------------------------------------------------
# URL de tu despliegue de Apps Script
URL_API_DRIVE = "https://script.google.com/macros/s/AKfycbwUfREwwrhpFdQoTnFdW5KMGUlMBBHaZ9vtR-RtVgeT4OPxvXgh1Ak1_VktrvXyPGW9UA/exec"

@st.cache_data(ttl=120) # Se refresca automáticamente cada 2 minutos
def obtener_archivos_drive():
    try:
        response = requests.get(URL_API_DRIVE)
        if response.status_code == 200:
            datos = response.json()
            return pd.DataFrame(datos)
        else:
            st.error("Error al conectar con la API de Google Drive.")
    except Exception as e:
        st.error(f"Fallo de conexión: {e}")
    return pd.DataFrame()

df_archivos = obtener_archivos_drive()

def buscar_id_archivo(palabra_clave):
    """Busca un archivo en el DataFrame devuelto por la API usando una palabra clave."""
    if not df_archivos.empty:
        resultado = df_archivos[df_archivos['nombre'].str.contains(palabra_clave, case=False, na=False)]
        if not resultado.empty:
            return resultado.iloc[0]['id']
    return None

def mostrar_pdf_drive(file_id):
    """Renderiza el PDF embebido directamente desde Drive."""
    url = f"https://drive.google.com/file/d/{file_id}/preview"
    iframe = f'<iframe class="pdf-frame" src="{url}" width="100%" height="700"></iframe>'
    st.markdown(iframe, unsafe_allow_html=True)

@st.cache_data(ttl=120)
def cargar_excel_observaciones(file_id):
    """Descarga y lee la matriz de Excel. Se configuran las filas (skiprows) según tu formato."""
    url = f'https://drive.google.com/uc?id={file_id}&export=download'
    # El EDA demostró que los encabezados reales de la matriz de Kreston empiezan en la fila 16 (índice 15)
    df = pd.read_excel(url, engine='openpyxl', skiprows=15)
    return df

# -----------------------------------------------------------------------------
# 4. BARRA LATERAL (NAVEGACIÓN DE MÓDULOS KRESTON)
# -----------------------------------------------------------------------------
st.sidebar.markdown('### 🗂️ Módulos de Evaluación')

opciones = [
    "🏠 Inicio y Sincronización",
    "1️⃣ Políticas de Seguridad",
    "2️⃣ Gestión de Activos de TI",
    "3️⃣ Seguridad en RRHH",
    "4️⃣ Continuidad y Respaldos",
    "5️⃣ Cumplimiento Legal",
    "📊 6. Cierre Hallazgos 2025"
]
seleccion = st.sidebar.radio("Seleccione el componente:", opciones)

# -----------------------------------------------------------------------------
# 5. LÓGICA DE VISTAS (PÁGINAS)
# -----------------------------------------------------------------------------
if seleccion == "🏠 Inicio y Sincronización":
    st.markdown("""
        <div class="card-custom">
            <div class="card-header-custom">Estado del Sistema SGSI</div>
            <p style="font-size: 1.1rem;">Bienvenido al portal oficial de auditoría de SERGEM Mensajería S.A.S. 
            El sistema está conectado en tiempo real al repositorio documental seguro.</p>
        </div>
    """, unsafe_allow_html=True)
    
    col_btn, col_info = st.columns([1, 3])
    with col_btn:
        if st.button("🔄 Forzar Sincronización"):
            st.cache_data.clear()
            st.success("✅ Datos sincronizados con Drive.")
    
    with col_info:
        st.info(f"Archivos y carpetas detectados en la nube: **{len(df_archivos)}**")

    if not df_archivos.empty:
        with st.expander("Ver índice de archivos detectados (Logs)"):
            st.dataframe(df_archivos[['nombre', 'tipo', 'ruta']], use_container_width=True)

elif seleccion == "1️⃣ Políticas de Seguridad":
    st.markdown('<div class="card-custom"><div class="card-header-custom">Módulo: Políticas de Seguridad</div></div>', unsafe_allow_html=True)
    
    # Busca la política en tiempo real. 
    id_pdf = buscar_id_archivo("Politica") 
    if id_pdf:
        mostrar_pdf_drive(id_pdf)
    else:
        st.warning("⚠️ No se detectó un archivo con la palabra 'Politica' en el Google Drive. Súbalo para visualizarlo automáticamente.")

elif seleccion == "📊 6. Cierre Hallazgos 2025":
    st.markdown("""
        <div class="card-custom">
            <div class="card-header-custom">Matriz de Observaciones e Indicadores (RM-4278)</div>
            <p>Control automatizado de hallazgos de la vigencia anterior.</p>
        </div>
    """, unsafe_allow_html=True)
    
    id_excel = buscar_id_archivo("Matriz de Observaciones")
    if id_excel:
        try:
            df_obs = cargar_excel_observaciones(id_excel)
            
            # Limpieza básica para evitar filas vacías del Excel
            df_obs = df_obs.dropna(subset=['OBSERVACIÓN', 'ESTADO'])
            
            # 1. Tabla de datos
            st.dataframe(
                df_obs[['COMPONENTE', 'OBSERVACIÓN', 'ESTADO', 'COMO FUE SUBSANADO (ACTIVIDAD REALIZADA)']], 
                use_container_width=True, 
                hide_index=True
            )
            
            # 2. Gráfico Dinámico con Columnas
            st.markdown('<br><h4 style="color:#002b5e;">Estado de Cierre de Observaciones</h4>', unsafe_allow_html=True)
            col1, col2 = st.columns([1, 2])
            
            with col1:
                conteo_estados = df_obs['ESTADO'].value_counts()
                st.dataframe(conteo_estados, use_container_width=True)
            with col2:
                st.bar_chart(conteo_estados, color="#28a745")
                
        except Exception as e:
            st.error(f"Error procesando los datos del Excel: {e}")
    else:
        st.warning("⚠️ No se encontró la Matriz de Observaciones en el Drive.")

else:
    # Plantilla por defecto para los demás módulos mientras los construyes
    titulo_modulo = seleccion[4:] # Quita el emoji y el número
    st.markdown(f"""
        <div class="card-custom">
            <div class="card-header-custom">{titulo_modulo}</div>
            <p>Sincronización en curso con Google Drive. Suba los documentos correspondientes a este módulo.</p>
        </div>
    """, unsafe_allow_html=True)
