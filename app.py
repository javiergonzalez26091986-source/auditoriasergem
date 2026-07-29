import streamlit as st
import pandas as pd
import requests
import base64
import os

# -----------------------------------------------------------------------------
# 1. CONFIGURACIÓN Y ESTILOS AVANZADOS (CON ICONO Y LOGO PERSONALIZADO)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Auditoría SGSI - SERGEM", 
    page_icon="sergemLogo.ico", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

def obtener_logo_base64(ruta_imagen):
    if os.path.exists(ruta_imagen):
        with open(ruta_imagen, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode()
        return f'<img src="data:image/png;base64,{encoded_string}" style="height: 45px; margin-right: 15px; background-color: #ffffff; padding: 4px; border-radius: 6px;">'
    return "🛡️ " 

html_logo = obtener_logo_base64("sergemLogo.png")

st.markdown(f"""
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        .stApp, [data-testid="stAppViewContainer"] {{ background-color: #e9ecef !important; }}
        [data-testid="stHeader"] {{ background-color: transparent !important; }}
        .stApp p, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6, .stApp label, .stApp li {{
            color: #2c3e50 !important;
        }}
        
        div.stButton > button {{
            background-color: #ffffff !important;
            color: #002b5e !important;
            border: 2px solid #002b5e !important;
            font-weight: 600 !important;
            border-radius: 6px !important;
            transition: 0.3s ease;
        }}
        div.stButton > button p {{ color: #002b5e !important; margin: 0 !important; }}
        div.stButton > button:hover {{ background-color: #002b5e !important; border-color: #002b5e !important; }}
        div.stButton > button:hover p {{ color: #ffffff !important; }}

        #MainMenu {{visibility: hidden;}} header {{visibility: hidden;}} footer {{visibility: hidden;}}
        [data-testid="collapsedControl"] {{display: none !important;}}
        section[data-testid="stSidebar"] {{ width: 300px !important; min-width: 300px !important; background-color: #dbe2e8 !important; }}
        
        .block-container {{padding-top: 0rem; padding-bottom: 0rem; max-width: 95%;}}
        .navbar-custom {{background-color: #002b5e; padding: 15px; margin-bottom: 20px; align-items: center; display: flex;}}
        .navbar-brand {{color: #ffffff !important; font-weight: bold; font-size: 1.6rem; display: flex; align-items: center;}}
        .card-custom {{border: none; border-radius: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.08); background: #ffffff; padding: 25px; margin-bottom: 20px;}}
        .card-header-custom {{border-bottom: 3px solid #002b5e; font-weight: 800; color: #002b5e; padding-bottom: 10px; margin-bottom: 20px; font-size: 1.4rem;}}
        .pdf-frame {{border: 1px dashed #cccccc; border-radius: 8px;}}
    </style>
    <nav class="navbar navbar-expand-lg navbar-custom">
      <div class="container-fluid">
        <span class="navbar-brand">{html_logo} SERGEM Mensajería S.A.S. - Portal de Auditoría SGSI</span>
      </div>
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
    "📊 Novedades Auditoría Pasada",
    "🛠️ Preparador de Auditoría Automático"
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
# 5. MÓDULO: NOVEDADES AUDITORÍA PASADA 
# -----------------------------------------------------------------------------
elif seleccion == "📊 Novedades Auditoría Pasada":
    st.markdown("""
        <div class="card-custom">
            <div class="card-header-custom">Hallazgos y Novedades Auditoría Pasada</div>
            <p>Resumen interactivo de las observaciones pasadas. Expanda cada componente para ver el detalle de la observación y la actividad de subsanación.</p>
        </div>
    """, unsafe_allow_html=True)

    @st.cache_data(ttl=300)
    def cargar_matriz_observaciones(df_archivos_nube):
        if df_archivos_nube.empty:
            return pd.DataFrame(), None
            
        match = df_archivos_nube[df_archivos_nube['nombre'].str.contains("RM-4278-25-Matriz", case=False, na=False)]
        
        if match.empty:
            st.error("No se encontró el archivo RM-4278-25-Matriz en la nube.")
            return pd.DataFrame(), None 
            
        file_id = match.iloc[0]['id']
        url_descarga = f"https://drive.google.com/uc?export=download&id={file_id}"
        
        try:
            df = pd.read_excel(url_descarga, sheet_name='AÑO', header=15)
            df = df.iloc[1:].copy()
            
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
            
            return df_clean, file_id 
        except Exception as e:
            st.error(f"Error al intentar leer las pestañas del archivo: {e}")
            return pd.DataFrame(), None

    df_nov, matrix_file_id = cargar_matriz_observaciones(df_archivos)

    if not df_nov.empty:
        conteo_estados = df_nov['Estado'].value_counts()
        
        st.markdown("### 📈 Resumen General de Hallazgos")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Observaciones", len(df_nov))
        col2.metric("✅ Subsanadas", conteo_estados.get('SUBSANADA', 0))
        col3.metric("⚠️ NO Subsanadas", conteo_estados.get('NO SUBSANADA', 0))
        col4.metric("🔒 Cerradas", conteo_estados.get('CERRADO', 0))

        st.divider()

        if matrix_file_id:
            with st.expander("📄 Clic aquí para verificar el archivo matriz original (Excel de Auditoría)"):
                st.info("Vista en vivo del documento fuente alojado en Google Drive. Si modificas el archivo allí, los cambios se reflejarán aquí tras la sincronización.")
                url_visor = f"https://drive.google.com/file/d/{matrix_file_id}/preview"
                st.markdown(f'<iframe class="pdf-frame" src="{url_visor}" width="100%" height="600"></iframe>', unsafe_allow_html=True)
                
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

# -----------------------------------------------------------------------------
# 6. MÓDULO INTELIGENTE: PREPARACIÓN DE AUDITORÍA
# -----------------------------------------------------------------------------
elif seleccion == "🛠️ Preparador de Auditoría Automático":
    st.markdown("""
        <div class="card-custom">
            <div class="card-header-custom">Preparación Automática para Auditoría (ISO 27001)</div>
            <p>El sistema escanea el inventario del repositorio documental buscando los requisitos del formato <b>RM-4901-26</b> de Kreston. 
            Identifica qué archivos ya poseemos y permite generar copias actualizadas con fecha del año en curso en la carpeta 'Auditoría actual'.</p>
        </div>
    """, unsafe_allow_html=True)

    if not df_archivos.empty:
        # Diccionario de requisitos del PDF de Kreston (Palabras clave para buscar en el Drive)
        requisitos = {
            "Políticas de Seguridad de la Información": ["POLITICA", "SEGURIDAD", "INFORMACION"],
            "Matriz de Riesgos de TI": ["MATRIZ", "RIESGO"],
            "Plan de Continuidad del Negocio": ["CONTINUIDAD", "NEGOCIO"],
            "Copias de Seguridad (Estado)": ["COPIA", "SEGURIDAD", "REPORTE"],
            "Procedimientos de Seguridad": ["PROCEDIMIENTO", "SEGURIDAD"],
            "Inventario de TI y Licenciamiento": ["INVENTARIO", "LICENCIA"],
            "Plan de contingencia (Ataque/Daño)": ["PLAN", "CONTINGENCIA"],
            "Borrados Seguros y Altas/Bajas": ["PROCEDIMIENTO", "BORRADO", "ALTAS"],
            "Matriz de Roles y Responsabilidades": ["ROLES", "RESPONSABILIDAD"]
        }

        archivos_encontrados = []
        ids_para_copiar = []

        st.markdown("### 📋 Análisis de Requisitos Documentales")
        
        for req, keywords in requisitos.items():
            mask = df_archivos['nombre'].str.upper().str.contains('|'.join(keywords))
            coincidencias = df_archivos[mask]

            if not coincidencias.empty:
                candidato = coincidencias.iloc[0]
                archivos_encontrados.append({
                    "Requisito": req, 
                    "Estado": "✅ Encontrado", 
                    "Archivo Base": candidato['nombre']
                })
                ids_para_copiar.append(candidato['id'])
            else:
                archivos_encontrados.append({
                    "Requisito": req, 
                    "Estado": "❌ Faltante / No detectado", 
                    "Archivo Base": "Requiere carga manual"
                })

        df_analisis = pd.DataFrame(archivos_encontrados)
        st.dataframe(df_analisis, use_container_width=True, hide_index=True)

        st.divider()
        st.markdown("### 🚀 Acción de Automatización")
        st.info(f"Se encontraron **{len(ids_para_copiar)}** documentos base en el sistema que cumplen con el check-list de auditoría.")
        
        if st.button("▶️ Generar Copias Actualizadas en 'Auditoría Actual'", type="primary"):
            with st.spinner("Conectando con Google Drive y generando copias... esto puede tomar unos segundos."):
                payload = {
                    "action": "copiar_archivos",
                    "fileIds": ids_para_copiar
                }
                try:
                    res_post = requests.post(URL_API_DRIVE, json=payload)
                    respuesta = res_post.json()
                    
                    if respuesta.get("status") == "success":
                        st.success(f"✅ ¡Éxito! Se han copiado y actualizado los nombres de {len(respuesta.get('copiados', []))} archivos en la carpeta Auditoría actual.")
                        with st.expander("Ver detalle de archivos creados"):
                            for f in respuesta.get("copiados", []):
                                st.write(f"- {f}")
                        # Forzamos limpiar cache para que el explorador recargue los nuevos archivos
                        st.cache_data.clear()
                    else:
                        st.error(f"Error en Google Drive: {respuesta.get('message')}")
                except Exception as e:
                    st.error(f"Error de comunicación con la API: {e}")
