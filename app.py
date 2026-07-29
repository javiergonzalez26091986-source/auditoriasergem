import streamlit as st
import pandas as pd
import requests

# 1. CONFIGURACIÓN Y ESTILOS
st.set_page_config(page_title="Auditoría SGSI - SERGEM", page_icon="🛡️", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        #MainMenu {visibility: hidden;} header {visibility: hidden;} footer {visibility: hidden;}
        /* El botón de colapsar menú YA NO ESTÁ OCULTO para que no se te pierda la barra lateral */
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

# NUEVA FUNCIÓN INTELIGENTE PARA LEER EL EXCEL SIN IMPORTAR EN QUÉ FILA ESTÉN LOS TÍTULOS
@st.cache_data(ttl=120)
def cargar_excel_observaciones(file_id):
    url = f'https://drive.google.com/uc?id={file_id}&export=download'
    # Cargamos el archivo sin asignar encabezados iniciales
    df_raw = pd.read_excel(url, engine='openpyxl', header=None)
    
    header_idx = 0
    # Buscamos dinámicamente la fila que contenga la palabra "ESTADO"
    for idx, row in df_raw.iterrows():
        # Convertimos la fila a texto para buscar
        row_str = " ".join([str(val).upper().strip() for val in row.values if pd.notna(val)])
        if "ESTADO" in row_str and "OBSERVACIÓN" in row_str:
            header_idx = idx
            break
            
    # Ahora sí cargamos el archivo saltando hasta la fila exacta del encabezado
    df = pd.read_excel(url, engine='openpyxl', skiprows=header_idx)
    # Limpiamos los nombres de las columnas para evitar el KeyError por espacios ocultos (ej: 'ESTADO ')
    df.columns = [str(col).strip().upper() for col in df.columns]
    
    return df

def mostrar_pdf_drive(file_id):
    url = f"https://drive.google.com/file/d/{file_id}/preview"
    st.markdown(f'<iframe class="pdf-frame" src="{url}" width="100%" height="800"></iframe>', unsafe_allow_html=True)

df_archivos = obtener_archivos_drive()

if 'visor_id' not in st.session_state: st.session_state.visor_id = None
if 'visor_nombre' not in st.session_state: st.session_state.visor_nombre = None

# BARRA LATERAL
st.sidebar.markdown('### 🗂️ Módulos de Evaluación')
opciones = ["🏠 Inicio y Sincronización", "📁 Explorador Documental Completo", "📊 Cierre Hallazgos 2025 (Matriz)"]
seleccion = st.sidebar.radio("Seleccione la vista:", opciones)

# VISTAS
if seleccion == "🏠 Inicio y Sincronización":
    st.markdown('<div class="card-custom"><div class="card-header-custom">Estado del Sistema SGSI</div><p>Bienvenido al portal oficial de auditoría.</p></div>', unsafe_allow_html=True)
    if st.button("🔄 Forzar Sincronización con Drive"):
        st.cache_data.clear()
        st.success("✅ Datos sincronizados correctamente.")

elif seleccion == "📁 Explorador Documental Completo":
    st.markdown('<div class="card-custom"><div class="card-header-custom">Repositorio Documental (Auditoría Kreston)</div></div>', unsafe_allow_html=True)
    if not df_archivos.empty:
        col_explorer, col_viewer = st.columns([1, 2])
        with col_explorer:
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
                mostrar_pdf_drive(st.session_state.visor_id)
            else:
                st.info("👈 Seleccione un documento para visualizarlo.")

elif seleccion == "📊 Cierre Hallazgos 2025 (Matriz)":
    st.markdown('<div class="card-custom"><div class="card-header-custom">Matriz de Observaciones (RM-4278-25)</div></div>', unsafe_allow_html=True)
    if not df_archivos.empty:
        resultado = df_archivos[df_archivos['nombre'].str.contains("Matriz de Observaciones", case=False, na=False)]
        if not resultado.empty:
            try:
                df_obs = cargar_excel_observaciones(resultado.iloc[0]['id'])
                # Filtramos las filas que sí tienen datos reales
                df_obs = df_obs.dropna(subset=['OBSERVACIÓN', 'ESTADO'])
                
                st.dataframe(df_obs[['COMPONENTE', 'OBSERVACIÓN', 'ESTADO', 'COMO FUE SUBSANADO (ACTIVIDAD REALIZADA)']], use_container_width=True, hide_index=True)
                
                st.markdown('<br><h4>Estado de Cierre</h4>', unsafe_allow_html=True)
                col1, col2 = st.columns([1, 2])
                conteo_estados = df_obs['ESTADO'].value_counts()
                with col1: st.dataframe(conteo_estados, use_container_width=True)
                with col2: st.bar_chart(conteo_estados, color="#002b5e")
            except Exception as e:
                st.error(f"Error procesando la Matriz: {e}")
