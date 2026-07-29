import streamlit as st
import pandas as pd

# 1. CONFIGURACIÓN DE LA PÁGINA
st.set_page_config(
    page_title="SGSI - SERGEM Mensajería S.A.S.",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. ESTILOS CSS PERSONALIZADOS (Para darle un toque corporativo y limpio)
st.markdown("""
    <style>
    .main-header {font-size: 2.5rem; color: #004d99; font-weight: bold;}
    .sub-header {font-size: 1.5rem; color: #333333;}
    .pdf-container {border: 1px solid #cccccc; border-radius: 5px; padding: 10px; background-color: #f9f9f9;}
    </style>
""", unsafe_allow_html=True)

# 3. FUNCIONES DE CONEXIÓN CON GOOGLE DRIVE
@st.cache_data(ttl=600) # Cache para no recargar el Excel en cada clic
def cargar_datos_excel(file_id):
    """Lee un archivo Excel público de Google Drive usando su ID"""
    url = f'https://drive.google.com/uc?id={file_id}&export=download'
    df = pd.read_excel(url, engine='openpyxl')
    return df

def mostrar_pdf_drive(file_id):
    """Incrusta un visor de PDF de Google Drive en la web"""
    url = f"https://drive.google.com/file/d/{file_id}/preview"
    iframe = f'<div class="pdf-container"><iframe src="{url}" width="100%" height="600" style="border: none;"></iframe></div>'
    st.markdown(iframe, unsafe_allow_html=True)

# 4. BARRA DE NAVEGACIÓN LATERAL (Basada en la solicitud RM-4901-26)
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/3251/3251522.png", width=100) # Puedes cambiarlo por el logo de SERGEM
st.sidebar.title("Auditoría SGSI 2026")
st.sidebar.subheader("Menú de Navegación")

opciones = [
    "Inicio y Dashboard General",
    "1. Políticas de Seguridad",
    "2. Gestión de Activos de TI",
    "3. Seguridad en RRHH",
    "4. Continuidad y Respaldos",
    "5. Cumplimiento Legal",
    "Cierre Hallazgos 2025"
]
seleccion = st.sidebar.radio("Seleccione un módulo:", opciones)

# 5. LÓGICA DE LAS PÁGINAS
if seleccion == "Inicio y Dashboard General":
    st.markdown('<p class="main-header">Sistema de Gestión de Seguridad de la Información (SGSI)</p>', unsafe_allow_html=True)
    st.write("Bienvenido al portal de auditoría de SERGEM Mensajería S.A.S. Utilice el menú lateral para navegar por los componentes evaluados.")
    
    st.info("💡 La documentación ha sido actualizada y centralizada para la auditoría de la vigencia 2026.")

elif seleccion == "1. Políticas de Seguridad":
    st.markdown('<p class="main-header">Políticas de Seguridad de la Información</p>', unsafe_allow_html=True)
    st.write("Visualización de las políticas de seguridad y protección de datos aprobadas por la gerencia.")
    
    # EJEMPLO: Aquí pones el ID del PDF de las políticas que está en tu Drive
    # Para obtener el ID: Clic derecho en el archivo en Drive -> Obtener enlace -> Copias la cadena de texto larga entre /d/ y /view
    id_pdf_politicas = "AQUI_PONES_EL_ID_DEL_PDF" 
    # mostrar_pdf_drive(id_pdf_politicas) # Descomentar cuando pongas el ID

elif seleccion == "Cierre Hallazgos 2025":
    st.markdown('<p class="main-header">Cierre de Observaciones (Auditoría 2025)</p>', unsafe_allow_html=True)
    st.write("Seguimiento y cierre de los hallazgos reportados en el informe RM-1253-24 y RM-4278-25.")
    
    try:
        # Aquí colocaremos el ID del Excel de la Matriz de Observaciones
        id_excel_observaciones = "AQUI_PONES_EL_ID_DEL_EXCEL"
        # df_obs = cargar_datos_excel(id_excel_observaciones)
        # st.dataframe(df_obs)
        st.success("Módulo de conexión a datos listo para configurar.")
    except Exception as e:
        st.error(f"Error al cargar la matriz: {e}")

# (Puedes agregar las demás secciones siguiendo esta misma lógica)
