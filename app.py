import streamlit as st
import pandas as pd
import requests
import base64
import os
import io

# -----------------------------------------------------------------------------
# 1. CONFIGURACIÓN Y ESTILOS AVANZADOS
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
            background-color: #ffffff !important; color: #002b5e !important;
            border: 2px solid #002b5e !important; font-weight: 600 !important;
            border-radius: 6px !important; transition: 0.3s ease;
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
# 2. CONEXIÓN API EN VIVO
# -----------------------------------------------------------------------------
URL_API_DRIVE = "https://script.google.com/macros/s/AKfycbzg7ezgkf0lU94fjXKRBGxlK5khR0pCaOgCLko6SEwUWYp55_IwYf3Syp1ownlT8D2ahQ/exec"

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
# 3. GENERADOR DE PLANTILLAS
# -----------------------------------------------------------------------------
def generar_plantilla_csv(tipo_requisito):
    if "capacitaciones" in tipo_requisito.lower():
        df = pd.DataFrame(columns=["FECHA (2026)", "NOMBRE DEL COLABORADOR", "CARGO", "TEMA DE CAPACITACION", "FIRMA"])
    elif "personal retirado" in tipo_requisito.lower():
        df = pd.DataFrame(columns=["FECHA DE RETIRO (2026)", "NOMBRE COMPLETO", "CÉDULA", "PAZ Y SALVO TI (SI/NO)", "ACCESOS REVOCADOS (SI/NO)"])
    elif "inventario" in tipo_requisito.lower():
        df = pd.DataFrame(columns=["TIPO DE ACTIVO", "MARCA/MODELO", "SERIAL", "RESPONSABLE ASIGNADO", "SISTEMA OPERATIVO", "ESTADO"])
    else:
        df = pd.DataFrame(columns=["FECHA (2026)", "DESCRIPCIÓN", "RESPONSABLE", "OBSERVACIONES"])
    
    return df.to_csv(index=False).encode('utf-8-sig')

# -----------------------------------------------------------------------------
# 4. BARRA LATERAL FIJA
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
# 5. LÓGICA DE VISTAS (INICIO Y EXPLORADOR)
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
# 6. MÓDULO: NOVEDADES AUDITORÍA PASADA 
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
        mask = df_archivos_nube['nombre'].str.contains("Matriz de Observaciones", case=False, na=False)
        archivos_matriz = df_archivos_nube[mask & (df_archivos_nube['tipo'] == 'Archivo')]
        if archivos_matriz.empty:
            st.error("No se encontró ningún archivo de matriz de observaciones en la nube.")
            return pd.DataFrame(), None 
        candidato = archivos_matriz.iloc[0]
        file_id = candidato['id']
        url_descarga = f"https://drive.google.com/uc?export=download&id={file_id}"
        
        try:
            df = pd.read_excel(url_descarga, sheet_name=0, header=15)
            df.columns = df.columns.astype(str).str.replace('\n', ' ').str.strip()
            
            col_informe = next((col for col in df.columns if 'INFORME' in col.upper()), None)
            col_componente = next((col for col in df.columns if 'COMPONENTE' in col.upper()), None)
            col_observacion = next((col for col in df.columns if 'OBSERVACIÓN' in col.upper() or 'OBSERVACION' in col.upper()), None)
            col_subsanacion = next((col for col in df.columns if 'COMO FUE SUBSANADO' in col.upper() or 'ACTIVIDAD' in col.upper()), None)
            col_estado = next((col for col in df.columns if 'ESTADO' in col.upper()), None)
            
            datos_limpios = []
            if col_observacion:
                for idx, row in df.iterrows():
                    obs = str(row.get(col_observacion)).strip()
                    if pd.isna(row.get(col_observacion)) or obs.upper() in ['', 'NAN', 'OBSERVACIÓN', 'OBSERVADA']:
                        continue
                    estado = 'SIN ESTADO'
                    if col_estado and pd.notna(row.get(col_estado)):
                        estado = str(row[col_estado]).upper()
                    elif 'SUBSANADO' in df.columns:
                        idx_subs = df.columns.get_loc('SUBSANADO')
                        val_si = str(row.iloc[idx_subs]).strip().upper()
                        val_no = str(row.iloc[idx_subs + 1]).strip().upper() if (idx_subs + 1) < len(df.columns) else ''
                        if val_si == 'X':
                            estado = 'SUBSANADA'
                        elif val_no == 'X':
                            estado = 'NO SUBSANADA'

                    actividad = 'Sin actividad registrada'
                    if col_subsanacion and pd.notna(row.get(col_subsanacion)):
                        val_act = str(row[col_subsanacion]).strip()
                        if val_act.upper() not in ['', 'NAN']:
                            actividad = val_act

                    datos_limpios.append({
                        'Informe': str(row.get(col_informe, 'N/A')),
                        'Componente': str(row.get(col_componente, 'N/A')),
                        'Observación': obs,
                        'Estado': estado,
                        'Subsanación (Actividad)': actividad
                    })
            df_clean = pd.DataFrame(datos_limpios)
            return df_clean, file_id 
        except Exception as e:
            st.error(f"Error al procesar el archivo: {e}")
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
# 7. MÓDULO INTELIGENTE: PREPARACIÓN DE AUDITORÍA
# -----------------------------------------------------------------------------
elif seleccion == "🛠️ Preparador de Auditoría Automático":
    st.markdown("""
        <div class="card-custom">
            <div class="card-header-custom">Preparación Automática para Auditoría (ISO 27001/27002)</div>
            <p>El sistema escanea el inventario del repositorio documental buscando los requisitos exactos del formato <b>RM-4901-26</b>. 
            Identifica qué archivos ya poseemos y permite generar copias actualizadas en 'Auditoría actual'.</p>
        </div>
    """, unsafe_allow_html=True)

    if not df_archivos.empty:
        df_archivos_base = df_archivos[(~df_archivos['nombre'].str.contains("Actualizado", case=False, na=False)) & (df_archivos['tipo'] == 'Archivo')]

        # Diccionario ajustado EXACTAMENTE al documento de solicitud RM-4901-26
        requisitos = {
            "Políticas de la seguridad de la información": ["POLITICA", "SEGURIDAD", "INFORMACION"],
            "Políticas de protección de datos (Habeas Data)": ["HABEAS", "DATOS"],
            "Procedimientos, planillas y/o documentos (capacitaciones)": ["CAPACITACION", "PLANILLA"],
            "Procedimiento disciplinario": ["DISCIPLINARIO"],
            "Inventario de TI": ["INVENTARIO", "TI"],
            "Plan de actualización de los recursos tecnológicos": ["ACTUALIZACION", "RECURSOS"],
            "Procedimientos de seguridad": ["PROCEDIMIENTO", "SEGURIDAD"],
            "Hoja de vida de los equipos de cómputo y servidores": ["HOJA DE VIDA", "EQUIPO", "SERVIDOR"],
            "Políticas de control de acceso": ["CONTROL", "ACCESO"],
            "Base de datos, personal retirado 2026": ["RETIRADO", "2026"],
            "Contratos y cláusulas de confidencialidad": ["CONFIDENCIALIDAD", "CLAUSULA"],
            "Plan de respuesta a emergencias (Pérdida de info.)": ["EMERGENCIA", "PERDIDA"],
            "Políticas de contraseñas": ["CONTRASEÑA", "CLAVE"],
            "Políticas de uso de dispositivos móviles": ["MOVIL", "DISPOSITIVO"],
            "Políticas, procedimientos de incidentes": ["INCIDENTE", "RESPONSABILIDAD"],
            "Procedimiento de notificación de incidentes": ["NOTIFICACION", "INCIDENTE"],
            "Inventario de Licenciamiento": ["LICENCIA", "INVENTARIO"],
            "Documentos soporte de adquisición de licencias": ["SOPORTE", "ADQUISICION", "LICENCIA"],
            "Certificación software legal (Representante Legal)": ["CERTIFICACION", "LEGAL", "REPRESENTANTE"],
            "Acuerdos de servicio (Proveedores/Terceros)": ["ACUERDO", "SERVICIO", "PROVEEDOR"],
            "Copias de seguridad vigentes y estado": ["COPIA", "SEGURIDAD", "BACKUP"],
            "Prueba de restauración": ["RESTAURACION", "PRUEBA"],
            "Plan de continuidad del negocio": ["CONTINUIDAD", "NEGOCIO"],
            "Matriz de riesgos de TI": ["MATRIZ", "RIESGO"],
            "Informe de pruebas de vulnerabilidad (Ethical Hacking)": ["VULNERABILIDAD", "HACKING", "ETHICAL"],
            "Documentos de gestión de seguridad en contratos": ["CONTRATO", "SEGURIDAD", "PRESTADOR"],
            "SGSI (Sistema de gestión de seguridad)": ["SGSI"],
            "Plan de acción, preventivo y correctivo": ["PLAN", "ACCION", "PREVENTIVO", "CORRECTIVO"]
        }

        archivos_encontrados = []
        archivos_validos = [] 
        ids_procesados = set() 
        lista_faltantes = [] # Para almacenar los que requieren plantilla

        st.markdown("### 📋 Análisis de Requisitos Documentales (Kreston)")
        st.info("💡 **Guía de Acción:** Los documentos con estado '✅ Encontrado' serán actualizados automáticamente. Si un documento marca '❌ Faltante', podrás descargar una plantilla base para llenarla, firmarla y subirla al Drive.")

        for req, keywords in requisitos.items():
            mask = df_archivos_base['nombre'].str.upper().str.contains('|'.join(keywords))
            coincidencias = df_archivos_base[mask]

            if not coincidencias.empty:
                candidato = coincidencias.iloc[0]
                archivos_encontrados.append({
                    "Requisito": req, 
                    "Estado": "✅ Encontrado", 
                    "Archivo Base": candidato['nombre']
                })
                if candidato['id'] not in ids_procesados:
                    archivos_validos.append({"nombre": candidato['nombre'], "id": candidato['id']})
                    ids_procesados.add(candidato['id'])
            else:
                archivos_encontrados.append({
                    "Requisito": req, 
                    "Estado": "❌ Faltante", 
                    "Archivo Base": "Requiere carga (Descargar plantilla abajo)"
                })
                lista_faltantes.append(req)

        df_analisis = pd.DataFrame(archivos_encontrados)
        
        filtro_req = st.radio(
            "🔍 Filtrar estado de los documentos:", 
            ["Mostrar Todos", "❌ Solo Faltantes", "✅ Solo Encontrados"], 
            horizontal=True
        )
        
        if filtro_req == "❌ Solo Faltantes":
            df_mostrar = df_analisis[df_analisis['Estado'] == "❌ Faltante"]
        elif filtro_req == "✅ Solo Encontrados":
            df_mostrar = df_analisis[df_analisis['Estado'] == "✅ Encontrado"]
        else:
            df_mostrar = df_analisis
            
        st.dataframe(df_mostrar, use_container_width=True, hide_index=True)

        # ---------------------------------------------------------
        # SECCIÓN NUEVA: DESCARGA DE PLANTILLAS PARA FALTANTES
        # ---------------------------------------------------------
        if lista_faltantes:
            st.markdown("### 📑 Generador de Plantillas Base (2026)")
            st.warning("Selecciona un documento faltante de la lista para descargar una plantilla estructurada. Solo debes llenarla con la información de este año, hacerla firmar y subirla a Drive.")
            
            req_seleccionado = st.selectbox("Seleccione el requisito faltante:", lista_faltantes)
            
            if req_seleccionado:
                csv_data = generar_plantilla_csv(req_seleccionado)
                nombre_archivo_sugerido = f"Plantilla_{req_seleccionado.replace(' ', '_').replace('/', '-')}_2026.csv"
                
                st.download_button(
                    label=f"⬇️ Descargar Plantilla: {req_seleccionado}",
                    data=csv_data,
                    file_name=nombre_archivo_sugerido,
                    mime='text/csv',
                    type="secondary"
                )

        st.divider()
        st.markdown("### 🚀 Acción de Automatización")
        st.info(f"Se encontraron **{len(archivos_validos)}** documentos base en el sistema que cumplen con el check-list de la auditoría.")
        
        if st.button("▶️ Generar Copias Actualizadas en 'Auditoría Actual'", type="primary"):
            st.markdown("#### Progreso de la copia:")
            barra_progreso = st.progress(0)
            texto_estado = st.empty()
            resultados_finales = []
            
            for i, doc in enumerate(archivos_validos):
                texto_estado.write(f"⏳ Evaluando y copiando: {doc['nombre']}...")
                payload = {"action": "copiar_archivos", "fileIds": [doc['id']]}
                try:
                    res_post = requests.post(URL_API_DRIVE, json=payload)
                    respuesta = res_post.json()
                    if respuesta.get("status") == "success":
                        resultados_finales.extend(respuesta.get("copiados", []))
                    else:
                        resultados_finales.append(f"❌ Omitido (Bloqueo severo del dueño o conexión): {doc['nombre']}")
                except Exception as e:
                    resultados_finales.append(f"❌ Omitido (Archivo inaccesible o restringido): {doc['nombre']}")
                    
                barra_progreso.progress((i + 1) / len(archivos_validos))
            
            texto_estado.empty()
            st.success("✅ ¡Proceso finalizado! A continuación el detalle del estado de cada documento:")
            with st.expander("Ver detalle de operaciones", expanded=True):
                for f in resultados_finales:
                    st.write(f"- {f}")
            st.cache_data.clear()
