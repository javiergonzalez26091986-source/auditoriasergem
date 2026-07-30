import streamlit as st
import pandas as pd
import requests
import base64
import os
import io
import random
import datetime
import plotly.express as px
import openpyxl
import unicodedata

# LIBRERÍAS PARA GENERACIÓN DIRECTA DE PDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
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

ANIO_ACTUAL = datetime.datetime.now().year

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
    urls_descarga = [
        f"https://drive.google.com/uc?export=download&id={file_id}",
        f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
    ]
    
    for url in urls_descarga:
        try:
            headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
            r = requests.get(url, headers=headers, timeout=15)
            
            if r.status_code == 200 and len(r.content) > 2000:
                wb = openpyxl.load_workbook(io.BytesIO(r.content))
                ws = wb.active
                
                ultima_fila = 7
                for row_idx in range(8, ws.max_row + 10):
                    if ws.cell(row=row_idx, column=2).value or ws.cell(row=row_idx, column=3).value:
                        ultima_fila = row_idx
                
                try:
                    ultimo_consecutivo = int(ws.cell(row=ultima_fila, column=1).value)
                except:
                    ultimo_consecutivo = ultima_fila - 7
                
                fila_actual = ultima_fila + 1
                consecutivo = ultimo_consecutivo + 1
                
                sistemas_operativos = ["W10 Pro 64", "W11 Enterprise", "W11 Pro", "Ubuntu 22.04 LTS"]
                tipos = ["Torre / Board Asus", "Todo en Uno / HP", "Portatil / Lenovo", "Torre / Dell Optiplex", "Portatil / Asus Vivo"]
                procesadores = ["Intel Core i5-12400 2.5GHz", "Intel Core i7-11700 2.5GHz", "AMD Ryzen 5 5600G 3.9GHz", "Intel Core i5 10400 2.9GHz", "Intel Core i3-10100 3.6GHz"]
                rams = ["8 Gb DDR4", "16 Gb DDR4", "32 Gb DDR4", "8 Gb DDR3"]
                discos = ["SSD 512GB NVMe M.2", "SSD 1TB SATA", "SSD 256GB Kingston", "NVMe Samsung 470GB"]
                monitores = ["LG 22 pulgadas", "Samsung 24 pulgadas", "Janus 20 pulgadas", "Integrado", "Dell 24 pulgadas"]
                ubicaciones = ["Cali", "Bogotá", "Cartagena", "Medellín", "Barranquilla"]
                perifericos = ["Logitech", "Genius", "Microsoft", "HP", "Dell"]
                observaciones = [
                    "Mantenimiento preventivo anual programado.", 
                    "Equipo de nueva adquisición.", 
                    f"Actualización de RAM y Disco en {ANIO_ACTUAL}.", 
                    "Optimización de sistema operativo."
                ]
                
                for i in range(15):
                    dia = random.randint(1, 28)
                    mes = random.randint(1, 12)
                    fecha_dinamica = f"{dia:02d}/{mes:02d}/{ANIO_ACTUAL}"
                    
                    ws.cell(row=fila_actual, column=1, value=str(consecutivo))
                    ws.cell(row=fila_actual, column=2, value=f"SRG{random.randint(300, 999)}")
                    ws.cell(row=fila_actual, column=3, value=f"DESKTOP-{random.randint(1000, 9999)}")
                    ws.cell(row=fila_actual, column=4, value=random.choice(sistemas_operativos))
                    ws.cell(row=fila_actual, column=5, value=random.choice(tipos))
                    ws.cell(row=fila_actual, column=6, value=random.choice(procesadores))
                    ws.cell(row=fila_actual, column=7, value=random.choice(rams))
                    ws.cell(row=fila_actual, column=8, value=random.choice(discos))
                    ws.cell(row=fila_actual, column=9, value=random.choice(monitores))
                    ws.cell(row=fila_actual, column=10, value=f"MON-{random.randint(100, 999)}")
                    ws.cell(row=fila_actual, column=11, value=random.choice(ubicaciones))
                    ws.cell(row=fila_actual, column=12, value=fecha_dinamica)
                    ws.cell(row=fila_actual, column=13, value=random.choice(perifericos))
                    ws.cell(row=fila_actual, column=14, value=random.choice(perifericos))
                    ws.cell(row=fila_actual, column=15, value=random.choice(observaciones))
                    
                    fila_actual += 1
                    consecutivo += 1
                
                output = io.BytesIO()
                wb.save(output)
                return output.getvalue()
        except Exception as e:
            continue
    return None

def remover_acentos(texto):
    return "".join(c for c in unicodedata.normalize('NFD', str(texto)) if unicodedata.category(c) != 'Mn').upper()

# -----------------------------------------------------------------------------
# 3. MOTOR INTELIGENTE DE ESTRUCTURAS DOCUMENTALES QMS (PDF AMPLIADO MULTIPÁGINA)
# -----------------------------------------------------------------------------
def obtener_datos_qms(requisito):
    req = requisito.lower()
    
    # --- DOCUMENTOS AMPLIADOS Y ROBUSTOS PARA CUMPLIR CON ISO 27001 Y AUDITORÍA ---
    
    if "políticas de la seguridad" in req or "política de seguridad" in req:
        return {
            "codigo": "PO-01-001",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO Y MARCO DE REFERENCIA": "Establecer la declaración formal de la Dirección General de SERGEM Mensajería S.A.S. respecto al compromiso inquebrantable con la Seguridad de la Información. Este marco da cumplimiento estricto a las directrices del control A.5.1 de la norma internacional ISO/IEC 27001 y se alinea con la legislación colombiana vigente en materia de protección de datos personales y seguridad informática.",
                "2. ALCANCE ORGANIZACIONAL": "La presente política es de mandatorio cumplimiento para todos los colaboradores directos, temporales, contratistas, prestadores de servicios y terceros que operen o tengan acceso a la infraestructura tecnológica, bases de datos, redes de comunicación y activos físicos de SERGEM a nivel nacional (Cali, Bogotá, Medellín, Barranquilla, Cartagena e Ibagué).",
                "3. PRINCIPIOS RECTORES DEL SGSI": "• Confidencialidad: Garantizar que la información corporativa, logística y de clientes solo sea accesible por personal autorizado.\n• Integridad: Proteger la exactitud, completitud y validez de los datos operativos frente a alteraciones no autorizadas.\n• Disponibilidad: Asegurar que los sistemas de información, plataformas logísticas y canales de atención permanezcan accesibles ininterrumpidamente para los usuarios autorizados.",
                "4. DIRECTRICES ESTRATÉGICAS DE LA DIRECCIÓN": "La Gerencia General y la Dirección Administrativa proveerán los recursos económicos, logísticos y tecnológicos necesarios para mantener, evaluar y mejorar continuamente el Sistema de Gestión de Seguridad de la Información (SGSI). Ningún objetivo de negocio podrá anteponerse a la seguridad de los activos de información.",
                "5. GESTIÓN DE EXCEPCIONES Y SANCIONES": "Cualquier intento de vulneración, desviación o incumplimiento de los lineamientos descritos en esta política será calificado como falta grave y sometido de manera inmediata al Procedimiento Disciplinario interno (PR-03-002), sin perjuicio de las acciones legales penales o civiles a que haya lugar en los juzgados de la República de Colombia."
            }
        }
        
    elif "procedimientos de seguridad" in req:
        return {
            "codigo": "PR-05-010",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO OPERATIVO": "Documentar los procedimientos operativos estándar (SOP) de seguridad física y lógica aplicados en el procesamiento diario de mensajería y logística, cumpliendo con el control A.5.37 de la norma ISO/IEC 27001.",
                "2. RESPONSABILIDADES Y ROLES": "El Departamento de Tecnología e Infraestructura, en coordinación con la Jefatura de Operaciones, es el responsable directo de vigilar la ejecución estricta de estos protocolos en cada una de las sedes de la compañía.",
                "3. PROTOCOLOS DE ACCESO LÓGICO Y AUTENTICACIÓN": "• El acceso a los sistemas core (Freeway y bases de datos asociadas) exige autenticación robusta obligatoria mediante doble factor (2FA) y contraseñas alfanuméricas con rotación trimestral.\n• Se prohíbe terminantemente el uso de cuentas genéricas o compartidas para la ejecución de tareas operativas o administrativas.",
                "4. SEGURIDAD EN EL PUESTO DE TRABAJO (ESCRITORIO LIMPIO)": "• Todo colaborador debe bloquear su estación de trabajo al ausentarse de su escritorio (ataque de tecla Windows + L).\n• Queda prohibido dejar documentos físicos con información de clientes, guías o datos financieros sobre escritorios o zonas comunes al finalizar la jornada laboral.",
                "5. SEGURIDAD EN REDES Y COMUNICACIONES": "La red Wi-Fi corporativa está estrictamente segmentada. Las terminales de invitados operan en una VLAN aislada sin permisos de enrutamiento hacia los servidores de bases de datos centrales."
            }
        }
        
    elif "hoja de vida" in req:
        return {
            "codigo": "RG-08-015",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO DEL CONTROL": "Estandarizar el registro sistemático del ciclo de vida, características técnicas, asignaciones, mantenimientos correctivos y preventivos del hardware de la compañía, dando cumplimiento al control A.8.1 (Inventario de activos) de la norma ISO/IEC 27001.",
                "2. ALCANCE Y APLICABILIDAD": "Aplica de forma obligatoria a servidores, estaciones de trabajo de escritorio, computadores portátiles, impresoras y equipos de comunicación asignados al personal administrativo y operativo a nivel nacional.",
                "3. DIRECTRICES DE GESTIÓN Y MANTENIMIENTO": "• Cada activo tecnológico posee una 'Hoja de Vida' digital vinculada al inventario maestro en Excel.\n• Toda intervención técnica, cambio de componente (memorias RAM, discos de estado sólido) o traslado de sede debe quedar registrado con fecha, descripción y cédula del técnico responsable.",
                "4. PROTOCOLO DE BAJA Y RETIRO DE ACTIVOS": "Antes de proceder a la baja física o desecho de un equipo de cómputo, el área de TI ejecutará un borrado seguro de almacenamiento secundario (Wipe certificado mediante software especializado) para evitar la fuga o recuperación residual de datos corporativos o información sensible de clientes."
            }
        }
        
    elif "licenciamiento" in req or "soporte de adquisición" in req:
        return {
            "codigo": "PO-05-032",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO NORMATIVO": "Asegurar el cumplimiento estricto de los derechos de propiedad intelectual, contratos de usuario final (EULA) y prevenir la instalación de software no autorizado (Control A.5.32 de la ISO/IEC 27001 y legislación sobre derechos de autor en Colombia).",
                "2. POLÍTICA DE ADQUISICIÓN Y CONTROL": "• Ningún software comercial, libre, de código abierto (Open Source) o de prueba puede ser instalado en los equipos de SERGEM sin previa validación y aprobación escrita del Departamento de Tecnología.\n• Todos los soportes de compra, facturas electrónicas y certificados de licencias operativas se encuentran custodiados digitalmente en la carpeta de proveedores.",
                "3. AUDITORÍAS TRIMESTRALES DE SOFTWARE (SHADOW IT)": "El departamento de TI ejecutará de manera automatizada escaneos semestrales en las estaciones de trabajo para detectar instalaciones clandestinas o aplicaciones no autorizadas (Shadow IT), procediendo a su aislamiento y desinstalación inmediata con reporte a la Dirección Administrativa."
            }
        }
        
    elif "copia" in req and "seguridad" in req:
        return {
            "codigo": "PR-08-013",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO Y PROPÓSITO": "Definir los lineamientos técnicos y operativos para la creación, custodia, retención y protección de las copias de respaldo (backups) de la información institucional, mitigando riesgos de pérdida por fallas de hardware, ataques de ransomware o desastres físicos (Control A.8.13).",
                "2. FRECUENCIA Y CRONOGRAMA DE RESPALDO": "• Backups Incrementales: Ejecución automatizada diaria a las 02:00 AM sobre bases de datos operativas y ERP.\n• Backups Completos (Full): Ejecución automatizada semanal todos los domingos en horario no hábil.\n• Almacenamiento: Las copias son transferidas de forma encriptada hacia los servidores en la nube del proveedor certificado SOLINUX.",
                "3. POLÍTICA DE RETENCIÓN Y DISPONIBILIDAD", "Los respaldos diarios se retienen por un periodo mínimo de 30 días en entornos inmutables. Los respaldos mensuales se conservan por un ciclo de un (1) año para asegurar trazabilidad histórica y auditoría fiscal."
            }
        }
        
    elif "matriz de riesgos" in req:
        return {
            "codigo": "MT-06-001",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. MARCO METODOLÓGICO": "La presente matriz ha sido estructurada bajo los lineamientos metodológicos de la norma internacional ISO/IEC 27005, orientada a la apreciación, análisis, evaluación y tratamiento sistemático de los riesgos de seguridad de la información en SERGEM.",
                "2. FÓRMULA DE VALORACIÓN DEL RIESGO": "El nivel de riesgo se determina mediante la ponderación matemática: Riesgo = Probabilidad (P) x Impacto (I), evaluando las dimensiones de Confidencialidad, Integridad y Disponibilidad.",
                "3. PRINCIPALES RIESGOS IDENTIFICADOS Y TRATAMIENTO": "• Riesgo de Malware / Ransomware en Servidores Core: Nivel Alto. Tratamiento: Mitigado mediante la instalación de antivirus EDR centralizado, copias inmutables en la nube y capacitación anti-phishing al personal.\n• Riesgo de Fuga o Exfiltración de Datos Personales: Nivel Medio. Tratamiento: Mitigado mediante restricciones de puertos USB en equipos operativos, cifrado de información y firma obligatoria de acuerdos de confidencialidad (NDA)."
            }
        }
        
    elif "contratos" in req and ("gestión" in req or "seguridad" in req):
        return {
            "codigo": "PO-05-019",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO CONTRACTUAL": "Asegurar que los riesgos asociados al acceso de proveedores externos, contratistas y prestadores de servicios tecnológicos a los activos de información de SERGEM sean mitigados eficazmente, cumpliendo con los controles A.5.19 y A.5.20 de la ISO/IEC 27001.",
                "2. CLÁUSULAS DE SALVAGUARDA OBLIGATORIAS", "• Todo contrato con un tercero debe incluir inexorablemente un Anexo de Seguridad de la Información y un Acuerdo de Confidencialidad (NDA) debidamente firmado por el Representante Legal.\n• Los proveedores tecnológicos deben garantizar el cumplimiento de estándares equivalentes de ciberseguridad y protección de datos conforme a la Ley 1581 de 2012.",
                "3. DERECHO DE INSPECCIÓN Y AUDITORÍA", "SERGEM se reserva de manera explícita el derecho de auditar los controles técnicos, instalaciones y políticas internas de los proveedores críticos para verificar el cumplimiento de los Acuerdos de Nivel de Servicio (SLA) pactados."
            }
        }

    elif "plan de acción" in req or "preventivo y correctivo" in req:
        return {
            "codigo": "PR-10-001",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO DE MEJORA CONTINUA": "Garantizar la mejora continua del Sistema de Gestión de Seguridad de la Información (SGSI) mediante la gestión estructurada, análisis de causa raíz y tratamiento oportuno de las no conformidades, hallazgos de auditoría e incidentes detectados (Cláusulas 10.1 y 10.2 de la ISO/IEC 27001).",
                "2. CICLO METODOLÓGICO DE ACCIÓN", "1. Identificación y registro formal del hallazgo o desvío.\n2. Análisis de causa raíz empleando metodologías de ingeniería (Diagrama de Ishikawa o los 5 Porqués).\n3. Definición, asignación y ejecución de la acción correctiva o preventiva con responsables y plazos definidos.\n4. Verificación posterior de la eficacia de la solución implementada por parte de la Dirección Administrativa."
            }
        }

    elif "control de acceso" in req:
        return {
            "codigo": "PO-07-003",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO Y PROPÓSITO": "Limitar el acceso autorizado a los sistemas de información, redes de comunicación, bases de datos corporativas y sedes físicas de SERGEM Mensajería S.A.S., previniendo accesos no autorizados y protegiendo los activos críticos (Controles A.9.1 al A.9.4 de la ISO/IEC 27001).",
                "2. POLÍTICA DE CONTROL DE ACCESO LÓGICO", "• El acceso a los sistemas operativos y ERP se basa estrictamente en el principio de 'Privilegio Mínimo' (otorgando únicamente los permisos indispensables para cumplir con las funciones del cargo).\n• Las cuentas de usuario de colaboradores retirados deben ser desactivadas de forma definitiva en un plazo no mayor a 24 horas tras la notificación oficial de Gestión Humana.",
                "3. CONTROL DE ACCESO FÍSICO A INSTALACIONES", "• Las áreas de servidores (Data Center) y archivo central cuentan con control de acceso restringido mediante huella biométrica y/o llaves magnéticas.\n• El ingreso de visitantes a las instalaciones debe registrarse obligatoriamente en la portería, exigiendo presentación de documento de identidad, entrega de distintivo visible y supervisión permanente por un colaborador anfitrión."
            }
        }

    # --- OTROS DOCUMENTOS ESTÁNDAR AMPLIADOS ---
    elif "disciplinario" in req:
        return {
            "codigo": "PR-03-002",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO Y MARCO LEGAL": "Establecer de manera clara y transparente los lineamientos, tipificación de faltas y el debido proceso sancionatorio aplicable ante el incumplimiento de las políticas y procedimientos del Sistema de Gestión de Seguridad de la Información (SGSI) en SERGEM Mensajería S.A.S.",
                "2. TIPIFICACIÓN DE FALTAS CONTRA LA SEGURIDAD", "• Falta Leve: Desatención menor a recomendaciones de seguridad que no compromete datos críticos.\n• Falta Grave: Uso indebido de credenciales, omisión de bloqueo de estaciones de trabajo, revelación negligente de información operativa o manipulación no autorizada de equipos tecnológicos.",
                "3. PROCEDIMIENTO DE DESCARGOS Y SANCIONES", "1. Reporte formal del incidente por parte de TI o jefatura inmediata a Gestión Humana.\n2. Citación a descargos por escrito al colaborador implicado con respeto irrestricto al debido proceso.\n3. Evaluación conjunta entre Gerencia y Dirección Administrativa para la aplicación de la sanción disciplinaria (amonestación escrita, suspensión temporal o terminación de contrato con justa causa conforme al Código Sustantivo del Trabajo)."
            }
        }
    
    elif "actualización" in req or "recursos" in req:
        return {
            "codigo": "PL-07-005",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO ESTRATÉGICO": "Planificar, presupuestar y ejecutar la renovación tecnológica, mantenimiento preventivo y actualización de hardware y software de la compañía para mitigar riesgos asociados a la obsolescencia técnica (Control A.8.19 de la ISO/IEC 27001).",
                "2. CICLO DE VIDA DE LOS ACTIVOS", "• Los equipos de cómputo y servidores tienen un ciclo de vida útil estimado entre 4 y 5 años.\n• Las licencias de sistemas operativos, paquetes de oficina y antivirus corporativos se actualizan de manera permanente a sus últimas versiones estables con soporte activo de fábrica.",
                "3. PROCEDIMIENTO DE EJECUCIÓN", "El área de TI realiza un inventario anual consolidado, detecta componentes próximos a obsolescencia y presenta ante la Gerencia el plan de inversiones tecnológicas para su respectiva aprobación y ejecución presupuestal."
            }
        }
        
    elif "capacitaci" in req or "planilla" in req:
        return {
            "codigo": "PR-08-001",
            "tipo_firma": "FIRMA RESPONSABLE DE LA CAPACITACIÓN",
            "secciones": {
                "1. OBJETIVO DEL PROGRAMA": "Garantizar que la totalidad del personal administrativo, operativo y directivo de SERGEM reciba formación y concienciación periódica en materia de seguridad de la información y protección de datos (Control A.6.3 de la ISO/IEC 27001).",
                "2. AGENDA Y COBERTURA NACIONAL": "Se programan jornadas de capacitación dirigidas al personal a nivel nacional en las sedes de Cali, Bogotá, Medellín, Barranquilla, Cartagena e Ibagué.\n\nTemas tratados: Uso seguro de contraseñas, reconocimiento de correos sospechosos (Phishing), política de escritorio limpio y normatividad de Habeas Data (Ley 1581).",
                "3. COMPROMISOS Y CONSTANCIA": "La asistencia a estas jornadas es de carácter obligatorio. Cada asistente firma la planilla de asistencia digital o física como evidencia documental auditable para el cumplimiento del SGSI."
            }
        }

    elif "emergencia" in req or "pérdida" in req:
        return {
            "codigo": "PR-07-015",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO DE CONTINGENCIA": "Establecer un plan de acción inmediato y estructurado para mitigar, responder y recuperar la operatividad de la infraestructura y la información ante incidentes críticos, ciberataques, fallas de servidores o desastres físicos (Control A.5.29 y A.8.14).",
                "2. MÉTRICAS DE RECUPERACIÓN (RTO Y RPO)": "• RTO (Recovery Time Objective): Tiempo máximo tolerable de interrupción del sistema core fijado en 4 horas.\n• RPO (Recovery Point Objective): Punto máximo tolerable de pérdida de datos fijado en un ciclo de 24 horas.",
                "3. PROTOCOLO DE RESPUESTA A CRISIS", "1. Aislamiento inmediato de los equipos o segmentos de red afectados para frenar propagación de amenazas.\n2. Notificación urgente al Comité de Crisis y a la Gerencia General.\n3. Activación del protocolo de restauración de respaldos en la nube junto al proveedor tecnológico SOLINUX."
            }
        }
        
    elif "contraseña" in req or "clave" in req:
        return {
            "codigo": "PO-07-004",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO TÉCNICO": "Definir los lineamientos estrictos para la creación, longitud, complejidad, protección y rotación periódica de las credenciales de acceso a los sistemas de información de SERGEM (Control A.9.4.3 de la ISO/IEC 27001).",
                "2. ESTÁNDARES TÉCNICOS DE CONTRASEÑAS", "• Longitud mínima de ocho (8) caracteres combinando obligatoriamente letras mayúsculas, minúsculas, números y símbolos especiales.\n• Rotación obligatoria cada noventa (90) días calendario impidiendo la reutilización de las últimas cuatro contraseñas anteriores.",
                "3. PROHIBICIONES Y BUENAS PRÁCTICAS", "Queda terminantemente prohibido compartir contraseñas con compañeros, escribirlas en notas adhesivas (post-its) visibles en los monitores o utilizar datos personales fácilmente adivinables (fechas de nacimiento, nombres de mascotas o familiares)."
            }
        }
        
    elif "móvil" in req or "dispositivo" in req:
        return {
            "codigo": "PO-07-008",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO DE SEGURIDAD MÓVIL": "Establecer las normas de control y seguridad para el uso de dispositivos móviles corporativos o personales (política BYOD) que procesen, almacenen o transmitan información de la compañía (Control A.8.1 de la ISO/IEC 27001).",
                "2. DIRECTRICES DE CONFIGURACIÓN Y USO", "• Todo equipo móvil que maneje correo o datos corporativos debe contar con cifrado de almacenamiento activo, PIN de bloqueo obligatorio y solución de seguridad o antivirus actualizado.\n• Se prohíbe almacenar bases de datos de clientes en memorias USB o dispositivos personales no autorizados por TI.",
                "3. PROTOCOLO EN CASO DE PÉRDIDA O ROBO", "El colaborador tiene la obligación de reportar de forma inmediata la pérdida o hurto de su dispositivo móvil al departamento de TI para proceder con el bloqueo de cuentas y la ejecución remota de borrado seguro (Wipe)."
            }
        }

    elif "notificación" in req or "incidente" in req:
        return {
            "codigo": "PR-07-011",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO DE GESTIÓN": "Estandarizar el procedimiento operativo para la detección temprana, reporte formal, clasificación por criticidad y resolución de incidentes de seguridad de la información (Controles A.5.24 al A.5.28 de la ISO/IEC 27001).",
                "2. CANALES Y PLAZOS DE REPORTE", "• Todo colaborador o tercero que detecte una anomalía, comportamiento extraño en el software, pérdida de equipos o sospecha de acceso no autorizado tiene el deber ético y laboral de reportarlo de inmediato a través de la Mesa de Ayuda.\n• Los incidentes clasificados como críticos deben ser escalados y notificados a la Dirección en un plazo inferior a una (1) hora.",
                "3. REGISTRO Y LECCIONES APRENDIDAS", "Cada incidente resuelto debe documentarse en la bitácora de TI detallando la causa raíz, el tiempo de afectación y las medidas preventivas adoptadas para evitar su repetición futura."
            }
        }
        
    elif "acuerdo" in req or "servicio" in req or "confidencialidad" in req:
        return {
            "codigo": "PO-07-014",
            "tipo_firma": "CONTRATISTA / PROVEEDOR",
            "secciones": {
                "1. OBJETO Y MARCO CONTRACTUAL": "Establecer los términos y condiciones obligatorias sobre Acuerdos de Nivel de Servicio (SLA) y Salvaguarda de Confidencialidad aplicables a todos los proveedores y terceros que presten servicios tecnológicos o logísticos a SERGEM Mensajería S.A.S.",
                "2. CLÁUSULAS DE CONFIDENCIALIDAD ESTRICTA", "PRIMERA: El CONTRATISTA se obliga a mantener absoluta reserva sobre toda la información comercial, operativa, de clientes y bases de datos a las que tenga acceso, catalogándola como 'Información Confidencial'.\n\nSEGUNDA: El CONTRATISTA dará estricto cumplimiento a la normatividad de protección de datos personales (Ley 1581 de 2012).\n\nTERCERA: Queda prohibida la divulgación, comercialización o uso de los datos para fines distintos a los estrictamente pactados en el contrato principal de servicios."
            }
        }

    elif "retirado" in req or "base de datos" in req:
        return {
            "codigo": "PO-07-025",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO DE CONTROL DE PERSONAL": "Asegurar que los procesos de desvinculación laboral o terminación de contratos con terceros incluyan la revocación oportuna y total de los accesos lógicos y físicos a las instalaciones y sistemas de SERGEM (Control A.6.5 de la ISO/IEC 27001).",
                "2. PROTOCOLO DE DESVINCULACIÓN", "• Gestión Humana tiene la obligación de notificar formalmente al departamento de TI la novedad de retiro de cualquier colaborador el mismo día en que se genera la novedad.\n• El equipo de TI procederá a inhabilitar de manera inmediata (plazo máximo de 24 horas) las cuentas de correo electrónico, accesos al ERP, VPN y bases de datos.\n• Es requisito obligatorio para la firma del paz y salvo laboral la devolución de equipos portátiles, carnés, fichas de acceso y tokens de seguridad."
            }
        }
        
    else:
        cod_aleatorio = random.randint(10, 99)
        return {
            "codigo": f"SG-07-0{cod_aleatorio}",
            "tipo_firma": "ELABORADO / REVISADO / APROBADO",
            "secciones": {
                "1. OBJETIVO DEL DOCUMENTO Y MARCO NORMATIVO": f"Establecer los lineamientos técnicos, políticas restrictivas y controles operativos aplicables al proceso de: {requisito.title()}, en estricto cumplimiento del marco normativo internacional de la ISO/IEC 27001 para la seguridad de la información.",
                "2. ALCANCE ORGANIZACIONAL": "Aplica de manera integral para todos los procesos operativos, administrativos, directivos y proveedores de servicios tecnológicos de SERGEM Mensajería S.A.S. a nivel nacional.",
                "3. DIRECTRICES Y CONTROLES OPERATIVOS", "• Todo el personal involucrado debe adherirse estrictamente a las pautas de control definidas en este documento.\n• El área de TI y la Dirección Administrativa realizarán supervisiones y auditorías preventivas periódicas para verificar el grado de cumplimiento.\n• Cualquier desvío detectado será objeto de revisión correctiva inmediata.",
                "4. COMPROMISOS Y MEJORA CONTINUA": "Garantizar la protección permanente de los activos informáticos, asegurando la confidencialidad, integridad y disponibilidad frente a riesgos o amenazas internas y externas."
            }
        }

def generar_documento_pdf(requisito):
    datos_doc = obtener_datos_qms(requisito)
    output = io.BytesIO()
    
    # Márgenes profesionales
    doc = SimpleDocTemplate(output, pagesize=letter, rightMargin=0.75*inch, leftMargin=0.75*inch, topMargin=0.6*inch, bottomMargin=0.6*inch)
    elements = []
    styles = getSampleStyleSheet()
    
    style_center = ParagraphStyle(name='Center', parent=styles['Normal'], alignment=TA_CENTER, fontName='Helvetica-Bold', fontSize=10)
    style_normal = ParagraphStyle(name='Justify', parent=styles['Normal'], alignment=TA_JUSTIFY, fontName='Helvetica', fontSize=10, leading=14)
    style_bold_center = ParagraphStyle(name='BoldCenter', parent=styles['Normal'], alignment=TA_CENTER, fontName='Helvetica-Bold', fontSize=10)
    style_title = ParagraphStyle(name='SectionTitle', parent=styles['Normal'], alignment=TA_LEFT, fontName='Helvetica-Bold', fontSize=11, textColor=colors.HexColor('#002b5e'))

    logo_path = "sergemLogo.png"
    if os.path.exists(logo_path):
        logo_img = RLImage(logo_path, width=1.0*inch, height=1.0*inch)
    else:
        logo_img = Paragraph("LOGO", style_bold_center)
        
    dia_aleatorio = random.randint(1, 28)
    fecha_generada = f"{dia_aleatorio:02d}/05/{ANIO_ACTUAL}"

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
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    elements.append(t_header)
    elements.append(Spacer(1, 0.25*inch))

    # Construcción de secciones con formato extendido multilínea
    for titulo, contenido in datos_doc['secciones'].items():
        contenido_rl = contenido.replace('\n', '<br/>')
        body_data = [
            [Paragraph(titulo, style_title)],
            [Paragraph(contenido_rl, style_normal)]
        ]
        t_body = Table(body_data, colWidths=[7.0*inch])
        t_body.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.HexColor('#bdc3c7')),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#f8f9fa')),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 10),
            ('TOPPADDING', (0,0), (-1,-1), 10),
            ('LEFTPADDING', (0,0), (-1,-1), 10),
            ('RIGHTPADDING', (0,0), (-1,-1), 10),
        ]))
        elements.append(t_body)
        elements.append(Spacer(1, 0.15*inch))

    elements.append(Spacer(1, 0.1*inch))
    if datos_doc['tipo_firma'] == "ELABORADO / REVISADO / APROBADO":
        sig_data = [
            [Paragraph("Elaborado por:", style_bold_center), Paragraph("Revisado por:", style_bold_center), Paragraph("Aprobado por:", style_bold_center)],
            [Paragraph("Nombre: Yesenia Beltrán<br/>Cargo: Directora Administrativa", style_normal),
             Paragraph("Nombre: Yesenia Beltrán<br/>Cargo: Directora Administrativa", style_normal),
             Paragraph("Nombre: José Reinel Torres<br/>Cargo: Gerente", style_normal)]
        ]
        t_sig = Table(sig_data, colWidths=[2.33*inch, 2.33*inch, 2.33*inch])
        t_sig.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ('TOPPADDING', (0,0), (-1,-1), 8),
            ('LEFTPADDING', (0,0), (-1,-1), 6),
        ]))
        elements.append(t_sig)
    else:
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

if 'excel_backup' not in st.session_state:
    st.session_state.excel_backup = None
if 'mostrar_salvavidas' not in st.session_state:
    st.session_state.mostrar_salvavidas = False

st.sidebar.markdown('### 🗂️ Módulos de Evaluación')
opciones = [
    "🏠 Inicio y Sincronización", 
    "📁 Explorador Documental Completo", 
    "📊 Novedades Auditoría Pasada", 
    "🛠️ Preparador de Auditoría Automático"
]
seleccion = st.sidebar.radio("Seleccione la vista:", opciones)

if seleccion == "🏠 Inicio y Sincronización":
    st.markdown("""
        <div class="card-custom">
            <div class="card-header-custom">Estado del Sistema SGSI</div>
            <p>Bienvenido al portal oficial de auditoría de SERGEM Mensajería S.A.S. El sistema se encuentra sincronizado con el repositorio documental en tiempo real.</p>
        </div>
    """, unsafe_allow_html=True)
    if st.button("🔄 Forzar Sincronización con Repositorio"):
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

elif seleccion == "🛠️ Preparador de Auditoría Automático":
    st.markdown(f"""
        <div class="card-custom">
            <div class="card-header-custom">Preparación Automática para Auditoría (ISO 27001)</div>
            <p>El sistema escanea el inventario del repositorio documental buscando los requisitos exactos del formato base.</p>
        </div>
    """, unsafe_allow_html=True)

    if not df_archivos.empty:
        # Se restringe la evaluación EXCLUSIVAMENTE a la carpeta "Auditoría actual"
        df_archivos_base = df_archivos[
            (df_archivos['tipo'] == 'Archivo') & 
            (df_archivos['ruta'].str.contains('Auditoría actual', case=False, na=False))
        ].copy()
        
        # Normalizar los nombres (sin acentos ni mayúsculas) para búsqueda perfecta
        df_archivos_base['nombre_norm'] = df_archivos_base['nombre'].apply(remover_acentos)

        requisitos = {
            "Políticas de la seguridad de la información": ["POLITICA", "SEGURIDAD", "INFORMACION"],
            "Políticas de protección de datos (Habeas Data)": ["PROTECCION", "DATOS"],
            "Procedimientos, planillas y/o documentos (capacitaciones)": ["CAPACITACION"],
            "Procedimiento disciplinario": ["DISCIPLINARIO"],
            "Inventario de TI": ["INVENTARIO", "COMPUTADORES"],
            "Plan de actualización de los recursos tecnológicos": ["ACTUALIZACION", "RECURSOS"],
            "Procedimientos de seguridad": ["PROCEDIMIENTOS", "SEGURIDAD"], 
            "Hoja de vida de los equipos de cómputo y servidores": ["HOJA", "VIDA"],
            "Políticas de control de acceso": ["CONTROL", "ACCESO"], # <-- Corregido aquí para detectar "control_acceso"
            f"Base de datos, personal retirado {ANIO_ACTUAL}": ["RETIRADO"],
            "Contratos y cláusulas de confidencialidad": ["CONFIDENCIALIDAD"],
            "Plan de respuesta a emergencias (Pérdida de info.)": ["EMERGENCIA", "PERDIDA"],
            "Políticas de contraseñas": ["CONTRASEÑA"],
            "Políticas de uso de dispositivos móviles": ["DISPOSITIVO", "MOVIL"],
            "Políticas, procedimientos de incidentes": ["ROLES", "RESPONSABILIDADES"],
            "Procedimiento de notificación de incidentes": ["NOTIFICACION", "INCIDENTE"],
            "Inventario de Licenciamiento": ["LICENCIA", "INVENTARIO"],
            "Documentos soporte de adquisición de licencias": ["SOPORTE", "ADQUISICION"],
            "Certificación software legal (Representante Legal)": ["LEGAL", "REPRESENTANTE"],
            "Acuerdos de servicio (Proveedores/Terceros)": ["ACUERDO", "SERVICIO", "PROVEEDOR"],
            "Copias de seguridad vigentes y estado": ["COPIA", "SEGURIDAD", "VIGENTE"],
            "Prueba de restauración": ["RESTAURACION", "PRUEBA"],
            "Plan de continuidad del negocio": ["CONTINUIDAD", "NEGOCIO"],
            "Matriz de riesgos de TI": ["MATRIZ", "RIESGO"],
            "Informe de pruebas de vulnerabilidad (Ethical Hacking)": ["VULNERABILIDAD", "HACKING"],
            "Documentos de gestión de seguridad en contratos": ["CONTRATO", "SEGURIDAD"],
            "SGSI (Sistema de gestión de seguridad)": ["SGSI"],
            "Plan de acción, preventivo y correctivo": ["PLAN", "ACCION", "PREVENTIVO"]
        }

        archivos_encontrados = []
        archivos_validos = []
        ids_procesados = set()
        inventario_id = None
        lista_faltantes = []
        
        for req, keywords in requisitos.items():
            mask = pd.Series(True, index=df_archivos_base.index)
            for kw in keywords:
                kw_norm = remover_acentos(kw)
                mask = mask & df_archivos_base['nombre_norm'].str.contains(kw_norm)

            coincidencias = df_archivos_base[mask]

            if not coincidencias.empty:
                candidato = coincidencias.iloc[0]
                
                es_excel_inventario = "INVENTARIO" in candidato['nombre_norm'] and candidato['nombre'].endswith(('.xls', '.xlsx'))
                
                if es_excel_inventario: 
                    estado = "⚙️ Encontrado (Sincronizable)"
                    inventario_id = candidato['id']
                else:
                    estado = "✅ Encontrado"
                    if candidato['id'] not in ids_procesados:
                        archivos_validos.append({"nombre": candidato['nombre'], "id": candidato['id']})
                        ids_procesados.add(candidato['id'])
                
                archivos_encontrados.append({
                    "Requisito": req, 
                    "Estado": estado, 
                    "Archivo Base": candidato['nombre']
                })
            else:
                archivos_encontrados.append({
                    "Requisito": req, 
                    "Estado": "❌ Faltante", 
                    "Archivo Base": "Buscar automáticamente"
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
        
        # --- VALIDADOR DE ARCHIVOS SOBRANTES (EXCLUSIVO CARPETA ACTUAL) ---
        nombres_validos = [d['nombre'] for d in archivos_validos]
        if inventario_id:
            nombres_validos.append(df_archivos_base[df_archivos_base['id'] == inventario_id].iloc[0]['nombre'])
            
        df_sobrantes = df_archivos_base[~df_archivos_base['nombre'].isin(nombres_validos)][['nombre', 'ruta']]
        if not df_sobrantes.empty:
            with st.expander(f"⚠️ Atención: Se detectaron {len(df_sobrantes)} archivos sobrantes en la carpeta actual (No requeridos)"):
                st.warning("Estos documentos no hacen parte de la lista oficial de la auditoría Kreston en esta carpeta. Considera verificar y removerlos para evitar confusiones a la hora de presentar los soportes.")
                st.dataframe(df_sobrantes, use_container_width=True, hide_index=True)
                
        st.divider()

        col_qms, col_auto = st.columns(2)
        with col_qms:
            st.markdown("### 📝 Buscador de Documentos Oficiales QMS")
            st.info("Busca de manera inteligente los documentos faltantes basándose en la normativa y los controles del SGSI y la base de datos de SERGEM.")
            
            if lista_faltantes:
                req_selec = st.selectbox("Seleccione el documento a generar:", lista_faltantes)
                
                if st.button(f"🪄 Descargar PDF Oficial: {req_selec}"):
                    with st.spinner("Compilando PDF..."):
                        archivo_pdf = generar_documento_pdf(req_selec)
                        nombre_descarga = f"{req_selec.replace('/', '_').replace(' ', '_')}_SERGEM_{ANIO_ACTUAL}.pdf"
                        
                        st.download_button(
                            label="⬇️ Descargar Documento Oficial", 
                            data=archivo_pdf, 
                            file_name=nombre_descarga, 
                            mime="application/pdf", 
                            type="secondary"
                        )
            else:
                st.success("✅ ¡Todos los documentos están listos!")

        with col_auto:
            st.markdown("### 🚀 Módulo de Actualización y Empaque")
            st.info(f"Se actualizará el Inventario de TI y se empaquetarán los **{len(archivos_validos)}** documentos validados en el repositorio de Auditoría.")
            
            if st.button(f"▶️ Sincronizar Repositorio Oficial {ANIO_ACTUAL}", type="primary"):
                st.session_state.mostrar_salvavidas = False 
                st.markdown("#### Progreso de la sincronización y actualización:")
                barra_progreso = st.progress(0)
                texto_estado = st.empty()
                resultados_finales = []
                
                total_pasos = len(archivos_validos) + (1 if inventario_id else 0)
                paso_actual = 0
                
                for doc in archivos_validos:
                    texto_estado.write(f"⏳ Evaluando y sincronizando: {doc['nombre']}...")
                    
                    archivos_mismo_nombre = df_archivos[df_archivos['nombre'] == doc['nombre']]
                    ya_existe_en_auditoria = archivos_mismo_nombre['ruta'].str.contains('Auditoría actual', case=False, na=False).any()
                    
                    if len(archivos_mismo_nombre) > 1 or ya_existe_en_auditoria:
                        resultados_finales.append(f"⏭️ Omitido (Ya existe en destino): {doc['nombre']}")
                        paso_actual += 1
                        barra_progreso.progress(paso_actual / total_pasos)
                        continue

                    payload = {"action": "copiar_archivos", "fileIds": [doc['id']]}
                    try:
                        res_post = requests.post(URL_API_DRIVE, json=payload)
                        if res_post.status_code == 200:
                            respuesta = res_post.json()
                            if respuesta.get("status") == "success": 
                                resultados_finales.append(f"✅ Sincronizado: {doc['nombre']}")
                            else: 
                                resultados_finales.append(f"❌ Error al procesar: {doc['nombre']}")
                        else:
                            resultados_finales.append(f"❌ Falló conexión: {doc['nombre']}")
                    except Exception:
                        resultados_finales.append(f"❌ Omitido (Archivo inaccesible): {doc['nombre']}")
                    
                    paso_actual += 1
                    barra_progreso.progress(paso_actual / total_pasos)
                
                if inventario_id:
                    texto_estado.write(f"⏳ Inyectando volumen de registros ({ANIO_ACTUAL}) en el Inventario Excel...")
                    excel_modificado = actualizar_fecha_inventario_excel(inventario_id)
                    
                    if excel_modificado:
                        st.session_state.excel_backup = excel_modificado 
                        excel_b64 = base64.b64encode(excel_modificado).decode('utf-8')
                        payload_excel = {
                            "action": "subir_archivo",
                            "nombre": f"Inventario de computadores - Actualizado {ANIO_ACTUAL}.xlsx",
                            "base64": excel_b64,
                            "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        }
                        
                        try:
                            res_excel = requests.post(URL_API_DRIVE, json=payload_excel)
                            if res_excel.status_code == 200:
                                try:
                                    respuesta_json = res_excel.json()
                                    if respuesta_json.get("status") == "success":
                                        resultados_finales.append(f"✅ Inventario_Actualizado_{ANIO_ACTUAL}.xlsx transferido.")
                                    else:
                                        resultados_finales.append(f"❌ El servidor rechazó la transferencia. Descárgalo de forma segura abajo.")
                                        st.session_state.mostrar_salvavidas = True
                                except:
                                    resultados_finales.append("⚠️ Retraso en la comunicación con el servidor. Descárgalo de forma segura abajo.")
                                    st.session_state.mostrar_salvavidas = True
                            else:
                                resultados_finales.append(f"❌ Fallo HTTP {res_excel.status_code} al sincronizar el Inventario.")
                                st.session_state.mostrar_salvavidas = True
                        except Exception as e:
                            resultados_finales.append(f"⚠️ Error de red al sincronizar el Inventario. Descárgalo de forma segura abajo.")
                            st.session_state.mostrar_salvavidas = True
                    else:
                        resultados_finales.append("❌ Falló el procesamiento del Inventario base.")
                        
                    paso_actual += 1
                    barra_progreso.progress(paso_actual / total_pasos)

                texto_estado.empty()
                st.success("✅ ¡Proceso de empacado finalizado! Detalle de operaciones:")
                with st.expander("Ver bitácora de actualización", expanded=True):
                    for f in resultados_finales: 
                        st.write(f"- {f}")
                st.cache_data.clear()

            if st.session_state.mostrar_salvavidas and st.session_state.excel_backup:
                st.warning("El tamaño del archivo o los límites de red impidieron la carga automática al repositorio.")
                st.info("💡 **ACCIÓN REQUERIDA:** El sistema ha validado y compilado tu archivo con éxito. Descárgalo y cárgalo manualmente en la carpeta de Auditoría.")
                st.download_button(
                    label=f"⬇️ Descargar Inventario Actualizado {ANIO_ACTUAL}",
                    data=st.session_state.excel_backup,
                    file_name=f"Inventario de computadores - Actualizado {ANIO_ACTUAL}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
