import streamlit as st
import pandas as pd

# 1. CONFIGURACIÓN DE PÁGINA (Ocultando la barra lateral por defecto)
st.set_page_config(
    page_title="SGSI - SERGEM",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 2. INYECCIÓN DE HTML5, CSS3 Y BOOTSTRAP 5
st.markdown("""
    <!-- Importar Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    
    <style>
        /* Ocultar la marca de agua y UI genérica de Streamlit */
        #MainMenu {visibility: hidden;}
        header {visibility: hidden;}
        footer {visibility: hidden;}
        [data-testid="collapsedControl"] {display: none;}
        .block-container {padding-top: 0rem; padding-bottom: 0rem; max-width: 100%; padding-left: 0; padding-right: 0;}
        
        /* Estilos Corporativos Personalizados */
        body {background-color: #f4f6f9;}
        .navbar-custom {background-color: #003366;}
        .navbar-brand {color: #ffffff !important; font-weight: bold; font-size: 1.5rem;}
        .nav-link {color: #e0e0e0 !important; font-weight: 500;}
        .nav-link:hover {color: #ffffff !important;}
        
        .dashboard-container {padding: 2rem;}
        
        .card-custom {
            border: none; 
            border-radius: 8px; 
            box-shadow: 0 4px 6px rgba(0,0,0,0.05); 
            background: #ffffff;
            margin-bottom: 1.5rem;
        }
        .card-header-custom {
            background-color: #ffffff; 
            border-bottom: 2px solid #003366; 
            font-weight: 700; 
            color: #003366;
            padding: 1rem 1.5rem;
            border-radius: 8px 8px 0 0;
        }
        .metric-value {font-size: 2rem; font-weight: bold; color: #28a745;}
    </style>
""", unsafe_allow_html=True)

# 3. MAQUETACIÓN DE LA BARRA DE NAVEGACIÓN (NAVBAR)
navbar = """
<nav class="navbar navbar-expand-lg navbar-custom">
  <div class="container-fluid">
    <a class="navbar-brand" href="#">
      🛡️ SERGEM Mensajería - SGSI 2026
    </a>
    <div class="collapse navbar-collapse" id="navbarNav">
      <ul class="navbar-nav ms-auto">
        <li class="nav-item"><a class="nav-link" href="#">Panel General</a></li>
        <li class="nav-item"><a class="nav-link" href="#">Auditoría Kreston</a></li>
        <li class="nav-item"><a class="nav-link" href="#">Cierre 2025</a></li>
      </ul>
    </div>
  </div>
</nav>
"""
st.markdown(navbar, unsafe_allow_html=True)

# 4. ESTRUCTURA DEL CONTENIDO PRINCIPAL (GRID BOOTSTRAP)
st.markdown('<div class="dashboard-container">', unsafe_allow_html=True)

# Crear dos columnas usando las herramientas nativas de Streamlit pero con estilo CSS
col1, col2 = st.columns([1, 3])

with col1:
    st.markdown("""
    <div class="card card-custom">
        <div class="card-header-custom">Módulos de Auditoría</div>
        <div class="card-body">
            <ul class="list-group list-group-flush">
                <li class="list-group-item">1. Políticas de Seguridad</li>
                <li class="list-group-item">2. Gestión de Activos</li>
                <li class="list-group-item">3. Recursos Humanos</li>
                <li class="list-group-item">4. Matriz de Riesgos</li>
            </ul>
        </div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="card card-custom">
        <div class="card-header-custom">Visor de Documentos Oficiales</div>
        <div class="card-body">
            <h5 class="card-title">Seleccione un documento en el panel izquierdo</h5>
            <p class="card-text text-muted">La documentación se sincroniza en tiempo real con el repositorio seguro (Google Drive).</p>
            <hr>
            <!-- Aquí embeberemos el PDF dinámicamente -->
            <div style="text-align: center; padding: 40px; background-color: #f8f9fa; border: 1px dashed #cccccc;">
                <p>Área de previsualización (iframe)</p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# (Aquí conectaremos luego Pandas y los IDs del Excel)
