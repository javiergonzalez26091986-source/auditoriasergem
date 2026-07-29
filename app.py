import streamlit as st
import pandas as pd
import requests
import base64
import os
import io
import plotly.express as px
import openpyxl
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

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
# 2. CONEXIONES API Y HERRAMIENTAS DE EDICIÓN
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

def actualizar_fecha_inventario_excel(file_id):
    url_descarga = f"https://drive.google.com/uc?export=download&id={file_id}"
    try:
        r = requests.get(url_descarga)
        if r.status_code == 200:
            wb = openpyxl.load_workbook(io.BytesIO(r.content))
            ws = wb.active
            for row in ws.iter_rows(min_row=1, max_row=10, min_col=1, max_col=20):
                for cell in row:
                    if cell.value and isinstance(cell.value, str) and '2025' in cell.value:
                        cell.value = cell.value.replace('2025', '2026')
            output = io.BytesIO()
            wb.save(output)
            return output.getvalue()
    except Exception as e:
        pass
    return None

def generar_documento_word(requisito):
    doc = Document()
    
    # Márgenes
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # NUEVO ENCABEZADO 2024: 2 filas x 5 columnas
    table = doc.add_table(rows=2, cols=5)
    table.style = 'Table Grid'
    
    # Ajuste aproximado de columnas
    widths = [Inches(1.2), Inches(1.5), Inches(1.0), Inches(1.0), Inches(1.2)]
    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width

    # Logo Izquierdo (Fusión Vertical)
    cell_logo_L = table.cell(0, 0)
    cell_logo_L.merge(table.cell(1, 0))
    p_logo_L = cell_logo_L.paragraphs[0]
    p_logo_L.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Logo Derecho (Fusión Vertical)
    cell_logo_R = table.cell(0, 4)
    cell_logo_R.merge(table.cell(1, 4))
    p_logo_R = cell_logo_R.paragraphs[0]
    p_logo_R.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Incrustar imágenes a ambos lados
    for p in [p_logo_L, p_logo_R]:
        try:
            if os.path.exists("sergemLogo.png"):
                p.add_run().add_picture("sergemLogo.png", width=Inches(0.9))
            else:
                p.add_run("LOGO").bold = True
        except:
            p.add_run("LOGO").bold = True

    # Título Central (Fusión Horizontal en fila 0, abarca col 1, 2 y 3)
    cell_title = table.cell(0, 1)
    cell_title.merge(table.cell(0, 3))
    p_title = cell_title.paragraphs[0]
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run(requisito.upper())
    run_title.bold = True
    run_title.font.size = Pt(11)

    # Metadatos en fila 1 (Código, Versión, Fecha)
    p_cod = table.cell(1, 1).paragraphs[0]
    p_cod.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cod.add_run("Código: PO-07-014").bold = True

    p_ver = table.cell(1, 2).paragraphs[0]
    p_ver.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ver.add_run("Versión No.1").bold = True

    p_fec = table.cell(1, 3).paragraphs[0]
    p_fec.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_fec.add_run("29/07/2026").bold = True

    doc.add_paragraph() # Espacio separador

    # FUNCIÓN INTERNA PARA CREAR SECCIONES TIPO "ACTA" (Cuadros)
    def crear_seccion_cuadro(titulo, contenido):
        t = doc.add_table(rows=2, cols=1)
        t.style = 'Table Grid'
        
        p_tit = t.cell(0, 0).paragraphs[0]
        p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p_tit.add_run(titulo)
        run.bold = True
        
        t.cell(1, 0).text = contenido
        doc.add_paragraph() 

    # LÓGICA DE CONTENIDO SEGÚN REQUISITO
    if "capacitaci" in requisito.lower() or "planillas" in requisito.lower():
        crear_seccion_cuadro("AGENDA DE LA REUNIÓN", "Se programa personal administrativo a nivel nacional: Cali, Barraquilla, Bogotá, Cartagena, Ibagué, Santa Marta para validación de: " + requisito)
        crear_seccion_cuadro("DESARROLLO DE LA REUNIÓN", "- Socialización de políticas y controles de Seguridad de la Información correspondientes al periodo 2026.")
        crear_seccion_cuadro("COMPROMISOS", "Dar la información necesaria a los diferentes grupos de interés, así como establecer los lineamientos que garanticen la protección de los datos y activos a través de los procedimientos de SERGEM.")
    else:
        crear_seccion_cuadro("OBJETIVO DEL DOCUMENTO", f"Establecer los lineamientos requeridos para dar cumplimiento normativo a: {requisito}.")
        crear_seccion_cuadro("ALCANCE", "Aplica para todos los empleados, contratistas y proveedores de SERGEM Mensajería S.A.S. a nivel nacional.")
        
        texto_politica = "- El personal debe cumplir estrictamente con los controles.\n- El incumplimiento generará medidas disciplinarias."
        if "copias" in requisito.lower() or "restauración" in requisito.lower():
            texto_politica = "- SERGEM delega la gestión de copias de seguridad en la nube al proveedor SOLINUX.\n- Se realizarán pruebas de restauración periódicas avaladas por el proveedor."
        elif "contraseñas" in requisito.lower():
            texto_politica = "- Las contraseñas deben ser alfanuméricas con una longitud mínima de 8 caracteres.\n- Se exige cambio obligatorio cada 90 días."
            
        crear_seccion_cuadro("REGLAS GENERALES / DESARROLLO", texto_politica)
        crear_seccion_cuadro("COMPROMISOS", "Garantizar la actualización constante y el resguardo de la información según la norma ISO 27001.")

    # FIRMA
    p_firma = doc.add_paragraph("\n\nFIRMA RESPONSABLE / APROBADOR: ___________________________________")
    p_firma.bold = True
    
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


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
            <p>Resumen visual e interactivo de las observaciones pasadas. Expanda cada componente para ver el detalle.</p>
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
        col_metric, col_chart = st.columns([1, 2])
        
        with col_metric:
            st.metric("Total Observaciones", len(df_nov))
            st.metric("✅ Subsanadas", conteo_estados.get('SUBSANADA', 0))
            st.metric("⚠️ NO Subsanadas", conteo_estados.get('NO SUBSANADA', 0))
            
        with col_chart:
            fig = px.pie(
                values=conteo_estados.values, 
                names=conteo_estados.index, 
                hole=0.4, 
                color=conteo_estados.index,
                color_discrete_map={'SUBSANADA':'#2ecc71', 'NO SUBSANADA':'#e74c3c', 'SIN ESTADO':'#f1c40f'}
            )
            fig.update_layout(margin=dict(t=0, b=0, l=0, r=0), height=300)
            st.plotly_chart(fig, use_container_width=True)

        st.divider()

        if matrix_file_id:
            with st.expander("📄 Clic aquí para verificar el archivo matriz original (Excel de Auditoría)"):
                st.info("Vista en vivo del documento fuente alojado en Google Drive.")
                url_visor = f"https://drive.google.com/file/d/{matrix_file_id}/preview"
                st.markdown(f'<iframe class="pdf-frame" src="{url_visor}" width="100%" height="600"></iframe>', unsafe_allow_html=True)
            
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

# -----------------------------------------------------------------------------
# 6. MÓDULO INTELIGENTE: PREPARACIÓN DE AUDITORÍA Y GENERADOR DE WORD
# -----------------------------------------------------------------------------
elif seleccion == "🛠️ Preparador de Auditoría Automático":
    st.markdown("""
        <div class="card-custom">
            <div class="card-header-custom">Preparación Automática para Auditoría (ISO 27001/27002)</div>
            <p>El sistema escanea el inventario del repositorio documental buscando los requisitos exactos del formato <b>RM-4901-26</b>.</p>
        </div>
    """, unsafe_allow_html=True)

    if not df_archivos.empty:
        df_archivos_base = df_archivos[(~df_archivos['nombre'].str.contains("Actualizado", case=False, na=False)) & (df_archivos['tipo'] == 'Archivo')]

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
        inventario_id = None 
        lista_faltantes = []

        st.markdown("### 📋 Análisis de Requisitos Documentales (Kreston)")
        
        for req, keywords in requisitos.items():
            mask = df_archivos_base['nombre'].str.upper().str.contains('|'.join(keywords))
            coincidencias = df_archivos_base[mask]

            if not coincidencias.empty:
                candidato = coincidencias.iloc[0]
                estado = "✅ Encontrado"
                
                if "INVENTARIO" in req.upper() and "TI" in req.upper() and candidato['nombre'].endswith(('.xls', '.xlsx')):
                    estado = "⚙️ Encontrado (Editable)"
                    inventario_id = candidato['id']

                archivos_encontrados.append({
                    "Requisito": req, 
                    "Estado": estado, 
                    "Archivo Base": candidato['nombre']
                })
                if candidato['id'] not in ids_procesados:
                    archivos_validos.append({"nombre": candidato['nombre'], "id": candidato['id']})
                    ids_procesados.add(candidato['id'])
            else:
                archivos_encontrados.append({
                    "Requisito": req, 
                    "Estado": "❌ Faltante", 
                    "Archivo Base": "Generar desde módulo inferior"
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
            df_mostrar = df_analisis[df_analisis['Estado'].str.contains("Encontrado")]
        else:
            df_mostrar = df_analisis
            
        st.dataframe(df_mostrar, use_container_width=True, hide_index=True)

        st.divider()

        # ---------------------------------------------------------
        # ZONA DE ACCIÓN: GENERADOR DE WORD Y EMPAQUE
        # ---------------------------------------------------------
        col_qms, col_auto = st.columns(2)
        
        with col_qms:
            st.markdown("### 📝 Motor Generador de Documentos QMS")
            st.info("Para los documentos faltantes, autogenera el formato oficial idéntico al acta de calidad de SERGEM (Versión 2024), listo para firmar.")
            
            if lista_faltantes:
                req_selec = st.selectbox("Seleccione el documento a construir:", lista_faltantes)
                
                if st.button(f"🪄 Crear Word Oficial: {req_selec}"):
                    archivo_word = generar_documento_word(req_selec)
                    nombre_descarga = f"{req_selec.replace('/', '_').replace(' ', '_')}_SERGEM_2026.docx"
                    
                    st.download_button(
                        label="⬇️ Descargar Documento Listo para Firmar",
                        data=archivo_word,
                        file_name=nombre_descarga,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="secondary"
                    )
            else:
                st.success("✅ ¡Todos los documentos están listos!")

        with col_auto:
            st.markdown("### 🚀 Acción de Automatización y Empaque")
            
            if inventario_id:
                st.warning("Se detectó el 'Inventario de TI' en formato Excel. Puedes descargar el archivo con la fecha actualizada a 2026.")
                if st.button("🪄 Descargar Inventario Actualizado (2026)"):
                    with st.spinner("Modificando celdas del Excel en segundo plano..."):
                        excel_modificado = actualizar_fecha_inventario_excel(inventario_id)
                        if excel_modificado:
                            st.download_button(
                                label="⬇️ Guardar Excel Actualizado",
                                data=excel_modificado,
                                file_name="Inventario_de_computadores_Actualizado_2026.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        else:
                            st.error("Hubo un error al intentar modificar el Excel. Revisa los permisos.")

            st.info(f"Se empaquetarán **{len(archivos_validos)}** documentos listos para la carpeta de auditoría.")
            
            if st.button("▶️ Generar Copias Oficiales en Drive", type="primary"):
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
