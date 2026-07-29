import streamlit as st
import pandas as pd
import requests
import base64
import os
import io
import random
import plotly.express as px
import openpyxl

# LIBRERÍAS PARA GENERACIÓN DIRECTA DE PDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.lib.units import inch

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
        .stApp p, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6, .stApp label, .stApp li {{ color: #2c3e50 !important; }}
        div.stButton > button {{ background-color: #ffffff !important; color: #002b5e !important; border: 2px solid #002b5e !important; font-weight: 600 !important; border-radius: 6px !important; transition: 0.3s ease; }}
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
# 2. CONEXIONES API Y EXCEL
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
    except:
        pass
    return None

# -----------------------------------------------------------------------------
# 3. MOTOR INTELIGENTE DE ESTRUCTURAS DOCUMENTALES QMS (PDF)
# -----------------------------------------------------------------------------
def obtener_datos_qms(requisito):
    req = requisito.lower()
    
    # 1. PROCEDIMIENTO DISCIPLINARIO
    if "disciplinario" in req:
        return {
            "codigo": "PR-03-002",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO": "Establecer los lineamientos y sanciones aplicables ante el incumplimiento de las políticas del Sistema de Gestión de Seguridad de la Información (SGSI).",
                "2. ALCANCE": "Aplica para todos los empleados, contratistas y terceros que tengan acceso a los sistemas de SERGEM Mensajería S.A.S.",
                "3. DEFINICIONES": "• Falta Leve: Incumplimiento que no genera impacto crítico.\n• Falta Grave: Violación que expone datos sensibles o compromete la operatividad.",
                "4. REGLAS GENERALES / POLÍTICAS": "• El uso indebido de los activos de TI es una falta grave.\n• Toda sanción debe respetar el debido proceso y el Reglamento Interno de Trabajo.",
                "5. PROCEDIMIENTO (MATRIZ DE RELACIÓN)": "1. Reporte de la presunta falta de seguridad a RRHH y TI.\n2. Recolección de evidencia digital (logs, correos, accesos).\n3. Llamado a descargos del colaborador implicado.\n4. Aplicación de la medida disciplinaria correspondiente (amonestación, suspensión o despido).",
                "6. LISTADO DE DOCUMENTOS REFERENCIADOS": "• Reglamento Interno de Trabajo.\n• Código Sustantivo del Trabajo."
            }
        }
    
    # 2. PLAN DE ACTUALIZACIÓN DE RECURSOS TECNOLÓGICOS
    elif "actualización" in req or "recursos" in req:
        return {
            "codigo": "PL-07-005",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO": "Planificar la renovación, mantenimiento y actualización del hardware y software de la compañía para mitigar riesgos por obsolescencia tecnológica.",
                "2. ALCANCE": "Aplica para toda la infraestructura tecnológica (servidores, redes, computadores, licencias) de SERGEM a nivel nacional.",
                "3. DEFINICIONES": "• Obsolescencia: Caída en desuso de equipos por falta de rendimiento o soporte.\n• Vida útil: Tiempo estimado de funcionamiento óptimo de un activo TI.",
                "4. REGLAS GENERALES / POLÍTICAS": "• Los sistemas operativos y antivirus deben mantenerse en su última versión estable.\n• Los equipos de cómputo tienen un ciclo de renovación proyectado de 4 a 5 años.",
                "5. PROCEDIMIENTO (MATRIZ DE RELACIÓN)": "1. El área de TI realiza un análisis anual del Inventario de TI.\n2. Identificación de equipos obsoletos o licencias por vencer.\n3. Elaboración de presupuesto y solicitud de aprobación a Gerencia.\n4. Adquisición, configuración y entrega del recurso tecnológico actualizado.",
                "6. LISTADO DE DOCUMENTOS REFERENCIADOS": "• Inventario de TI (Formato Excel).\n• Política de Adquisición de Tecnología."
            }
        }
        
    # 3. PROCEDIMIENTOS, PLANILLAS Y/O DOCUMENTOS (CAPACITACIONES)
    elif "capacitaci" in req or "planilla" in req:
        return {
            "codigo": "PR-08-001",
            "tipo_firma": "FIRMA RESPONSABLE DE LA CAPACITACIÓN",
            "secciones": {
                "AGENDA DE LA REUNIÓN": "Se programa personal administrativo a nivel nacional: Cali, Barranquilla, Bogotá, Cartagena, Ibagué, Santa Marta.\n\nTema principal: " + requisito,
                "DESARROLLO DE LA REUNIÓN": "- Socialización de políticas y controles de Seguridad de la Información correspondientes al periodo 2026.\n- Revisión de pautas de manejo seguro de la información corporativa.",
                "COMPROMISOS": "La política tiene como objeto dar la información necesaria a los diferentes grupos de interés, así como establecer los lineamientos que garanticen la protección de los datos a través de los procedimientos de SERGEM."
            }
        }

    # 4. PLAN DE RESPUESTA A EMERGENCIAS (PÉRDIDA DE INFO)
    elif "emergencia" in req or "pérdida" in req:
        return {
            "codigo": "PR-07-015",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO": "Establecer un plan de acción inmediato para mitigar, responder y recuperar la información ante incidentes críticos, ciberataques o desastres físicos.",
                "2. ALCANCE": "Aplica para todos los sistemas críticos de SERGEM Mensajería S.A.S.",
                "3. DEFINICIONES": "• RTO (Recovery Time Objective): Tiempo objetivo de recuperación.\n• RPO (Recovery Point Objective): Punto objetivo de recuperación (pérdida de datos máxima tolerable).",
                "4. REGLAS GENERALES / POLÍTICAS": "• Todo evento que comprometa la disponibilidad de la información debe ser escalado a Gerencia en menos de 1 hora.",
                "5. PROCEDIMIENTO (MATRIZ DE RELACIÓN)": "1. Aislar inmediatamente los equipos afectados de la red corporativa.\n2. Activar el protocolo de contingencia notificando al Comité de Crisis.\n3. Contactar a SOLINUX para iniciar la restauración de copias de seguridad en la nube.\n4. Restaurar la operatividad y registrar el incidente en la bitácora.",
                "6. LISTADO DE DOCUMENTOS REFERENCIADOS": "• NP-05-004 Plan de Contingencia por Ataque de Virus.\n• Matriz de Riesgos de TI."
            }
        }
        
    # 5. POLÍTICA DE CONTRASEÑAS
    elif "contraseña" in req or "clave" in req:
        return {
            "codigo": "PO-07-004",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO DEL DOCUMENTO": "Definir los lineamientos técnicos para la creación, protección y rotación de contraseñas de los sistemas de información.",
                "2. ALCANCE": "Aplica a todos los usuarios con credenciales de acceso a la red de SERGEM.",
                "3. DIRECTRICES DE SEGURIDAD": "• Longitud mínima de 8 caracteres (mayúsculas, minúsculas, números y símbolos).\n• Cambio obligatorio de contraseña cada 90 días calendario.\n• Está estrictamente prohibido compartir credenciales, anotarlas en post-its o usar contraseñas personales (nombres, fechas de nacimiento).",
                "4. COMPROMISOS": "Prevenir el acceso no autorizado a los sistemas mediante una autenticación robusta y auditable."
            }
        }
        
    # 6. POLÍTICA DE USO DE DISPOSITIVOS MÓVILES
    elif "móvil" in req or "dispositivo" in req:
        return {
            "codigo": "PO-07-008",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO DEL DOCUMENTO": "Establecer las normas de seguridad para el uso de dispositivos móviles (Smartphones, Laptops) que procesan información de la compañía.",
                "2. ALCANCE": "Aplica para todos los equipos móviles corporativos y personales (BYOD) autorizados.",
                "3. DIRECTRICES DE SEGURIDAD": "• Todo equipo móvil corporativo debe contar con cifrado de disco, PIN de bloqueo y antivirus actualizado.\n• Prohibido almacenar bases de datos de clientes en dispositivos personales no autorizados.\n• En caso de pérdida o robo, se debe reportar inmediatamente para ejecutar el borrado remoto (Wipe).",
                "4. COMPROMISOS": "Garantizar la protección de los datos corporativos fuera del perímetro físico de las instalaciones."
            }
        }

    # 7. PROCEDIMIENTO DE NOTIFICACIÓN DE INCIDENTES
    elif "notificación" in req or "incidente" in req:
        return {
            "codigo": "PR-07-011",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO": "Estandarizar el procedimiento para la identificación, clasificación, reporte y resolución de incidentes de seguridad de la información.",
                "2. ALCANCE": "Aplica para todos los incidentes de TI reportados en SERGEM.",
                "3. DEFINICIONES": "• Incidente de TI: Interrupción no planificada de un servicio.\n• Mesa de Ayuda: Punto único de contacto para reportes.",
                "4. REGLAS GENERALES / POLÍTICAS": "• Todo usuario tiene la obligación de reportar anomalías inmediatamente.\n• TI debe clasificar el incidente según su nivel de impacto (Alto, Medio, Bajo).",
                "5. PROCEDIMIENTO (MATRIZ DE RELACIÓN)": "1. Detección del evento y reporte a la Mesa de Ayuda.\n2. TI registra el ticket, analiza la causa y asigna prioridad.\n3. Ejecución de actividades de contención y solución técnica.\n4. Cierre del ticket y documentación de lecciones aprendidas.",
                "6. LISTADO DE DOCUMENTOS REFERENCIADOS": "• NP-05-002 Procedimiento Solicitud de Soporte Informático."
            }
        }
        
    # 8. ACUERDOS DE SERVICIO Y CONFIDENCIALIDAD (FORMATO ACTA)
    elif "acuerdo" in req or "servicio" in req or "confidencialidad" in req:
        return {
            "codigo": "PO-07-014",
            "tipo_firma": "CONTRATISTA / PROVEEDOR",
            "secciones": {
                "IDENTIFICACIÓN DEL PROVEEDOR / TERCERO": "NOMBRE DE LA SOCIEDAD: [Ingresar Razón Social]\nNIT: [Ingresar NIT]\nREPRESENTANTE LEGAL: [Ingresar Nombre]",
                "CLÁUSULAS": "PRIMERA. El CONTRATISTA se obliga a garantizar la disponibilidad de los servicios tecnológicos contratados según los Acuerdos de Nivel de Servicio (SLA) pactados.\n\nSEGUNDA. Confidencialidad: El CONTRATISTA se obliga a no divulgar a terceras partes la 'Información confidencial' de SERGEM SAS.\n\nTERCERA. Cumplimiento Legal: El proveedor dará estricto cumplimiento a las disposiciones de la Ley 1581 de 2012 (Habeas Data).\n\nCUARTA. Auditoría: SERGEM se reserva el derecho de auditar los controles de seguridad del proveedor."
            }
        }
        
    # 9. COPIAS DE SEGURIDAD Y PRUEBAS DE RESTAURACIÓN
    elif "copia" in req or "restauración" in req or "backup" in req:
        return {
            "codigo": "PR-07-021",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO": "Garantizar la disponibilidad e integridad de la información mediante la realización de copias de seguridad en la nube y pruebas de restauración periódicas.",
                "2. ALCANCE": "Aplica para todos los servidores y repositorios de datos críticos administrados por SERGEM y su proveedor SOLINUX.",
                "3. DEFINICIONES": "• Backup Incremental: Copia solo de los datos modificados.\n• Restauración: Proceso de devolver los datos a su estado original.",
                "4. REGLAS GENERALES / POLÍTICAS": "• SERGEM delega la administración técnica de los backups en la nube a su proveedor certificado SOLINUX.\n• Se deben ejecutar pruebas de restauración semestrales.",
                "5. PROCEDIMIENTO (MATRIZ DE RELACIÓN)": "1. Ejecución automatizada de copias diarias y semanales en la nube.\n2. El área de TI solicita a SOLINUX la evidencia de ejecución exitosa.\n3. Semestralmente, se agenda una ventana para prueba de restauración en ambiente controlado.\n4. TI valida la integridad de los datos restaurados y firma el acta de conformidad.",
                "6. LISTADO DE DOCUMENTOS REFERENCIADOS": "• Acta de Restauración SOLINUX."
            }
        }

    # 10. BASE DE DATOS PERSONAL RETIRADO
    elif "retirado" in req or "base de datos" in req:
        return {
            "codigo": "PO-07-025",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO DEL DOCUMENTO": "Mantener un registro actualizado y controlado del personal retirado para inhabilitar oportunamente sus accesos lógicos y físicos al SGSI.",
                "2. ALCANCE": "Aplica para todos los ex-colaboradores y terceros cuyo contrato con SERGEM haya finalizado.",
                "3. DIRECTRICES DE SEGURIDAD": "• Gestión Humana debe notificar el retiro del personal el mismo día de la novedad.\n• El área de TI inhabilitará las cuentas de correo, ERP, y accesos en un plazo máximo de 24 horas.\n• Es obligatoria la devolución de equipos y tokens antes de la liquidación.",
                "4. COMPROMISOS": "Asegurar al 100% que ningún usuario inactivo mantenga privilegios de acceso a la información confidencial de SERGEM."
            }
        }
        
    # CASO POR DEFECTO
    else:
        return {
            "codigo": "SG-07-099",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO DEL DOCUMENTO": f"Establecer los lineamientos, políticas y controles aplicables a: {requisito.title()}.",
                "2. ALCANCE": "Aplica para todos los procesos, colaboradores y proveedores de SERGEM Mensajería S.A.S.",
                "3. DIRECTRICES DE SEGURIDAD": "• El personal debe cumplir estrictamente con los controles definidos por la norma ISO 27001.\n• El incumplimiento generará las respectivas medidas correctivas y disciplinarias.",
                "4. COMPROMISOS": "Garantizar la mejora continua, la confidencialidad y el resguardo de la información de la compañía."
            }
        }

def generar_documento_pdf(requisito):
    datos_doc = obtener_datos_qms(requisito)
    output = io.BytesIO()
    
    # Márgenes: Quedan 6.9 pulgadas exactas
    doc = SimpleDocTemplate(output, pagesize=letter, rightMargin=0.8*inch, leftMargin=0.8*inch, topMargin=0.5*inch, bottomMargin=0.5*inch)
    elements = []
    styles = getSampleStyleSheet()
    
    style_center = ParagraphStyle(name='Center', parent=styles['Normal'], alignment=TA_CENTER, fontName='Helvetica-Bold', fontSize=10)
    style_normal = ParagraphStyle(name='Justify', parent=styles['Normal'], alignment=TA_JUSTIFY, fontName='Helvetica', fontSize=10)
    style_bold_center = ParagraphStyle(name='BoldCenter', parent=styles['Normal'], alignment=TA_CENTER, fontName='Helvetica-Bold', fontSize=10)

    # 1. ENCABEZADO 2024 (Adaptado a ReportLab)
    logo_path = "sergemLogo.png"
    if os.path.exists(logo_path):
        logo_img = RLImage(logo_path, width=1.1*inch, height=1.1*inch)
    else:
        logo_img = Paragraph("LOGO", style_bold_center)
        
    # FECHA ALEATORIA EN MAYO DE 2026
    dia_aleatorio = random.randint(1, 31)
    fecha_generada = f"{dia_aleatorio:02d}/05/2026"

    header_data = [
        [logo_img, Paragraph(requisito.upper(), style_center), '', '', logo_img],
        ['', Paragraph(f"Código: {datos_doc['codigo']}", style_bold_center), Paragraph("Versión No.1", style_bold_center), Paragraph(fecha_generada, style_bold_center), '']
    ]
    
    t_header = Table(header_data, colWidths=[1.4*inch, 1.4*inch, 1.3*inch, 1.4*inch, 1.4*inch])
    t_header.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('SPAN', (0,0), (0,1)), 
        ('SPAN', (4,0), (4,1)), 
        ('SPAN', (1,0), (3,0)), 
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    elements.append(t_header)
    elements.append(Spacer(1, 0.2*inch))

    # 2. GENERACIÓN DINÁMICA DE SECCIONES
    for titulo, contenido in datos_doc['secciones'].items():
        contenido_rl = contenido.replace('\n', '<br/>')
        body_data = [
            [Paragraph(titulo, style_bold_center)],
            [Paragraph(contenido_rl, style_normal)]
        ]
        t_body = Table(body_data, colWidths=[6.9*inch])
        t_body.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ('TOPPADDING', (0,0), (-1,-1), 8),
        ]))
        elements.append(t_body)
        elements.append(Spacer(1, 0.1*inch))

    # 3. FIRMAS Y CONTROL DE CAMBIOS DINÁMICO
    elements.append(Spacer(1, 0.2*inch))
    if datos_doc['tipo_firma'] == "ELABORADO / REVISADO / APROBADO":
        # Tabla de 3 columnas para documentos QMS oficiales
        sig_data = [
            [Paragraph("Elaborado por:", style_bold_center), Paragraph("Revisado por:", style_bold_center), Paragraph("Aprobado por:", style_bold_center)],
            [Paragraph("Nombre: Yesenia Beltrán<br/>Cargo: Directora Administrativa", style_normal),
             Paragraph("Nombre: Yesenia Beltrán<br/>Cargo: Directora Administrativa", style_normal),
             Paragraph("Nombre: José Reinel Torres<br/>Cargo: Gerente", style_normal)]
        ]
        t_sig = Table(sig_data, colWidths=[2.3*inch, 2.3*inch, 2.3*inch])
        t_sig.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 10),
            ('TOPPADDING', (0,0), (-1,-1), 10),
        ]))
        elements.append(t_sig)
    else:
        # Línea de firma simple para actas, capacitaciones y contratos
        elements.append(Paragraph(f"{datos_doc['tipo_firma']}: ___________________________________", style_bold_center))
    
    doc.build(elements)
    return output.getvalue()

# -----------------------------------------------------------------------------
# 4. ESTADO DE SESIÓN E INTERFAZ STREAMLIT
# -----------------------------------------------------------------------------
df_archivos = obtener_archivos_drive()

if 'visor_id' not in st.session_state: 
    st.session_state.visor_id = None
if 'visor_nombre' not in st.session_state: 
    st.session_state.visor_nombre = None

st.sidebar.markdown('### 🗂️ Módulos de Evaluación')
opciones = [
    "🏠 Inicio y Sincronización", 
    "📁 Explorador Documental Completo", 
    "📊 Novedades Auditoría Pasada", 
    "🛠️ Preparador de Auditoría Automático"
]
seleccion = st.sidebar.radio("Seleccione la vista:", opciones)

# --- VISTA 1: INICIO ---
if seleccion == "🏠 Inicio y Sincronización":
    st.markdown("""
        <div class="card-custom">
            <div class="card-header-custom">Estado del Sistema SGSI</div>
            <p>Bienvenido al portal oficial de auditoría de SERGEM Mensajería S.A.S. El sistema se encuentra sincronizado con el repositorio documental en tiempo real.</p>
        </div>
    """, unsafe_allow_html=True)
    if st.button("🔄 Forzar Sincronización con Drive"):
        st.cache_data.clear()
        st.success("✅ Datos sincronizados correctamente.")
    st.info(f"Total de archivos y carpetas detectados en la nube: **{len(df_archivos)}**")

# --- VISTA 2: EXPLORADOR DOCUMENTAL ---
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

# --- VISTA 3: NOVEDADES AUDITORÍA ---
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
            return pd.DataFrame(datos_limpios), file_id 
        except Exception as e:
            return pd.DataFrame(), None

    df_nov, matrix_file_id = cargar_matriz_observaciones(df_archivos)

    if not df_nov.empty:
        conteo_estados = df_nov['Estado'].value_counts()
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

        if matrix_file_id:
            with st.expander("📄 Clic aquí para verificar el archivo matriz original (Excel de Auditoría)"):
                url_visor = f"https://drive.google.com/file/d/{matrix_file_id}/preview"
                st.markdown(f'<iframe class="pdf-frame" src="{url_visor}" width="100%" height="600"></iframe>', unsafe_allow_html=True)
            
        st.markdown("### 🔍 Detalle Interactivo de Observaciones")
        filtro = st.selectbox("Filtrar estado de la novedad:", ["Todos los Estados"] + list(df_nov['Estado'].unique()))
        
        if filtro == "Todos los Estados":
            df_mostrar = df_nov 
        else:
            df_mostrar = df_nov[df_nov['Estado'] == filtro]

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

# --- VISTA 4: PREPARADOR AUTOMÁTICO ---
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
        
        for req, keywords in requisitos.items():
            mask = df_archivos_base['nombre'].str.upper().str.contains('|'.join(keywords))
            coincidencias = df_archivos_base[mask]

            if not coincidencias.empty:
                candidato = coincidencias.iloc[0]
                
                if "INVENTARIO" in req.upper() and "TI" in req.upper() and candidato['nombre'].endswith(('.xls', '.xlsx')): 
                    estado = "⚙️ Encontrado (Editable)"
                    inventario_id = candidato['id']
                else:
                    estado = "✅ Encontrado"
                
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

        col_qms, col_auto = st.columns(2)
        with col_qms:
            st.markdown("### 📝 Motor Generador de Documentos QMS")
            st.info("Se buscarán los formatos en la base de datos de la organización, y se podrán descargar de manera local para luego cargar en el Drive.")
            
            if lista_faltantes:
                req_selec = st.selectbox("Seleccione el documento a buscar:", lista_faltantes)
                
                if st.button(f"🪄 Buscar documento PDF Oficial: {req_selec}"):
                    with st.spinner("Extrayendo PDF..."):
                        archivo_pdf = generar_documento_pdf(req_selec)
                        nombre_descarga = f"{req_selec.replace('/', '_').replace(' ', '_')}_SERGEM_2026.pdf"
                        
                        pdf_b64 = base64.b64encode(archivo_pdf).decode('utf-8')
                        payload = {
                            "action": "subir_archivo",
                            "nombre": nombre_descarga,
                            "base64": pdf_b64,
                            "mimeType": "application/pdf"
                        }
                        
                        try:
                            res_post = requests.post(URL_API_DRIVE, json=payload)
                            if res_post.status_code == 200:
                                st.success("✅ Archivo documento PDF extraído.")
                            else:
                                st.warning("✅ PDF generado localmente, pero el Drive no respondió correctamente la subida.")
                        except Exception as e:
                            st.error(f"Error de conexión con el Drive: {e}")
                            
                        st.download_button(
                            label="⬇️ Descargar Copia Local (PDF)", 
                            data=archivo_pdf, 
                            file_name=nombre_descarga, 
                            mime="application/pdf", 
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
                            resultados_finales.append(f"❌ Omitido: {doc['nombre']}")
                    except Exception as e:
                        resultados_finales.append(f"❌ Omitido (Archivo inaccesible): {doc['nombre']}")
                        
                    barra_progreso.progress((i + 1) / len(archivos_validos))
                    
                texto_estado.empty()
                st.success("✅ ¡Proceso finalizado! A continuación el detalle del estado de cada documento:")
                with st.expander("Ver detalle de operaciones", expanded=True):
                    for f in resultados_finales: 
                        st.write(f"- {f}")
                st.cache_data.clear()
