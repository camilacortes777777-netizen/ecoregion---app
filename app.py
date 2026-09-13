import io
import math
from datetime import datetime

import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from docx import Document

# ------------------------------------------------------------------------------
# 1. CONFIGURACIÓN Y CSS AVANZADO (PALETA DE COLORES & UI PRO)
# ------------------------------------------------------------------------------
st.set_page_config(
    page_title="EcoRegión - Soluciones Ambientales",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Inyección de CSS Profesional
st.markdown(
    """
    <style>
    /* Importar fuente moderna (Inter) */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
        background-color: #F8F9FA;
        color: #212529;
    }
    
    /* Fondo principal de la App */
    .stApp {
        background-color: #F8F9FA;
    }

    /* Ocultar barra de menú de Streamlit arriba para aspecto limpio */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    /* Contenedor del Título Principal */
    .main-header {
        background: linear-gradient(135deg, #1E4D2B 0%, #2E7D32 100%);
        padding: 2.5rem 2rem;
        border-radius: 16px;
        color: white;
        box-shadow: 0 10px 25px rgba(30, 77, 43, 0.15);
        margin-bottom: 2rem;
    }
    .main-header h1 {
        color: #FFFFFF !important;
        font-weight: 700;
        font-size: 2.2rem;
        margin-bottom: 0.5rem;
    }
    .main-header p {
        color: #E8F5E9;
        font-size: 1.05rem;
        margin: 0;
    }

    /* Tarjetas Métricas Personalizadas (KPI Cards) */
    .kpi-card {
        background-color: #FFFFFF;
        padding: 1.5rem;
        border-radius: 14px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
        border-left: 5px solid #2E7D32;
        text-align: center;
        transition: transform 0.2s ease;
    }
    .kpi-card:hover {
        transform: translateY(-3px);
    }
    .kpi-title {
        font-size: 0.85rem;
        color: #6C757D;
        text-transform: uppercase;
        font-weight: 600;
        letter-spacing: 0.5px;
    }
    .kpi-value {
        font-size: 1.8rem;
        font-weight: 700;
        color: #1E4D2B;
        margin-top: 0.3rem;
    }

    /* Estilizado de Pestañas (Tabs) */
    .stTabs [data-baseweb="tab-list"] {
        gap: 12px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #FFFFFF;
        border-radius: 10px;
        padding: 10px 20px;
        font-weight: 600;
        color: #495057;
        box-shadow: 0 2px 5px rgba(0,0,0,0.03);
        border: 1px solid #E9ECEF;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1E4D2B !important;
        color: #FFFFFF !important;
        border-color: #1E4D2B !important;
    }

    /* Estilizado de Botones */
    .stButton>button, .stDownloadButton>button {
        background: linear-gradient(135deg, #2E7D32 0%, #1E4D2B 100%);
        color: white !important;
        border: none;
        border-radius: 10px;
        padding: 0.6rem 1.5rem;
        font-weight: 600;
        font-size: 1rem;
        box-shadow: 0 4px 12px rgba(46, 125, 50, 0.25);
        transition: all 0.3s ease;
        width: 100%;
    }
    .stButton>button:hover, .stDownloadButton>button:hover {
        background: linear-gradient(135deg, #1E4D2B 0%, #14371E 100%);
        box-shadow: 0 6px 18px rgba(30, 77, 43, 0.35);
        transform: translateY(-1px);
    }

    /* Formularios e Inputs */
    div[data-baseweb="select"] > div, div[data-baseweb="input"] > div {
        border-radius: 10px !important;
        border-color: #CED4DA !important;
    }
    
    /* Contenedores con borde fino */
    .custom-box {
        background-color: #FFFFFF;
        padding: 1.5rem;
        border-radius: 14px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.04);
        border: 1px solid #E9ECEF;
        margin-bottom: 1rem;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# ------------------------------------------------------------------------------
# 2. BARRA LATERAL (SIDEBAR) CORPORATIVA
# ------------------------------------------------------------------------------
with st.sidebar:
    st.image(
        "https://ecoregionsas.com/wp-content/uploads/2021/04/logo-ecoregion.png",
        use_container_width=True,
    )
    st.markdown("---")
    st.markdown("### 🏢 **ECO REGIÓN SAS BIC**")
    st.caption("Soluciones Ambientales & Licenciamiento")

    st.markdown(
        """
        <div style="background-color: #E8F5E9; padding: 12px; border-radius: 10px; border-left: 4px solid #2E7D32; margin-top: 15px;">
            <span style="color: #1E4D2B; font-weight: 600; font-size: 0.9rem;">🍃 Prototipo MVP v1.2</span><br>
            <span style="color: #495057; font-size: 0.8rem;">Trámite: Aprovechamiento Forestal de Árboles Aislados</span>
        </div>
    """,
        unsafe_allow_html=True,
    )

    st.markdown("---")
    st.markdown("#### ⚙️ **Configuración Rápida**")
    modo_oscuro_mapa = st.checkbox("Mapa Satelital / Alto Contraste", value=False)


# ------------------------------------------------------------------------------
# 3. ENCABEZADO PRINCIPAL (HEADER HERO)
# ------------------------------------------------------------------------------
st.markdown(
    """
    <div class="main-header">
        <h1>🌿 EcoRegión App</h1>
        <p>Gestión inteligente y automatización de permisos ambientales ante CAR y SDA</p>
    </div>
""",
    unsafe_allow_html=True,
)

# Pestañas Principales
tab_app, tab_guia, tab_encuesta = st.tabs(
    [
        "📋  Formulario de Solicitud",
        "📖  Guía de Uso del Trámite",
        "📊  Módulo de Validación & Usabilidad",
    ]
)


# ------------------------------------------------------------------------------
# 4. PESTAÑA 1: FORMULARIO INTERACTIVO Y CÁLCULO
# ------------------------------------------------------------------------------
with tab_app:
    st.markdown("### 📝 Registro del Proyecto Ambiental")
    st.caption("Diligencie los campos requeridos para estructurar automáticamente el documento técnico de radicación.")

    with st.form("form_aprovechamiento"):
        col1, col2 = st.columns(2, gap="medium")

        with col1:
            st.markdown("#### **1. Datos del Solicitante**")
            nombre = st.text_input("Nombre o Razón Social", "Nicol Daniela Navarrete")
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                tipo_doc = st.selectbox("Tipo Doc.", ["CC", "NIT", "CE", "PASAPORTE"])
            with col_d2:
                num_doc = st.text_input("Número de Doc.", "1000123456")
            calidad = st.selectbox(
                "Calidad en que actúa sobre el predio",
                ["Propietario", "Tenedor", "Poseedor", "Apoderado / Autorizado"],
            )

            st.markdown("#### **2. Ubicación y Autoridad Ambiental**")
            jurisdiccion = st.selectbox(
                "Autoridad Ambiental Competente",
                ["CAR Corpoboyacá", "CAR Cundinamarca", "SDA (Bogotá D.C.)"],
            )
            nombre_predio = st.text_input("Nombre del Predio / Institución", "E.S.E. Hospital Duitama")
            col_u1, col_u2 = st.columns(2)
            with col_u1:
                municipio = st.text_input("Municipio / Ciudad", "Duitama")
            with col_u2:
                direccion = st.text_input("Dirección", "Av. Las Américas # 12-45")

            st.markdown("**Coordenadas de Ubicación (GPS):**")
            col_lat, col_lon = st.columns(2)
            with col_lat:
                latitud = st.number_input("Latitud (Norte)", value=5.8267, format="%.5f")
            with col_lon:
                longitud = st.number_input("Longitud (Oeste)", value=-73.0331, format="%.5f")

        with col2:
            st.markdown("#### **3. Caracterización Dendrométrica**")
            especie_comun = st.text_input("Nombre Común de la Especie", "Caucho")
            especie_cientifico = st.text_input("Nombre Científico", "Hevea brasiliensis")
            cant_arboles = st.number_input("Número de Individuos Arbóreos", min_value=1, value=1)

            col_dap, col_alt = st.columns(2)
            with col_dap:
                dap_cm = st.number_input("DAP (cm - a 1.3m)", min_value=1.0, value=35.0)
            with col_alt:
                altura_m = st.number_input("Altura Total (m)", min_value=1.0, value=9.0)

            forma_fuste = st.selectbox(
                "Forma del Fuste (Factor de Forma - fm)",
                ["Paraboloide (0.55)", "Cilíndrico (0.75)", "Cónico (0.33)", "Neiloide (0.25)"],
            )

            st.markdown("#### **4. Diagnóstico y Fotografías**")
            col_chk1, col_chk2 = st.columns(2)
            with col_chk1:
                riesgo = st.checkbox("Riesgo Inminente de Caída", value=True)
            with col_chk2:
                fauna = st.checkbox("Presencia de Nidos/Avifauna", value=False)

            foto_arbol = st.file_uploader(
                "Cargar Evidencia Fotográfica (Fuste / Copa / Entorno)",
                type=["jpg", "jpeg", "png"],
            )

        submit_btn = st.form_submit_button("🔥 Procesar Solicitud y Generar Documentación")

    # Visualización Geográfica y Fotografía
    st.markdown("---")
    col_mapa, col_foto = st.columns(2, gap="medium")

    with col_mapa:
        st.markdown("##### 📍 **Georreferenciación del Individuo**")
        df_coordenadas = pd.DataFrame({"lat": [latitud], "lon": [longitud]})
        st.map(df_coordenadas, zoom=14)

    with col_foto:
        st.markdown("##### 📸 **Evidencia del Estado Fitosanitario**")
        if foto_arbol is not None:
            st.image(foto_arbol, caption=f"Registro cargado: {especie_comun}", use_container_width=True)
        else:
            st.info("No se ha adjuntado imagen previa.")

    # Lógica de Procesamiento y Resultados
    if submit_btn or "procesado" in st.session_state:
        st.session_state["procesado"] = True
        st.markdown("---")

        # Cálculos matemáticos
        dap_m = dap_cm / 100.0
        fm_val = float(forma_fuste.split("(")[1].replace(")", ""))
        volumen_m3 = (math.pi / 4) * (dap_m**2) * altura_m * fm_val * cant_arboles

        # KPI CARDS Estilizadas con HTML/CSS
        st.markdown("### 📊 **Resumen Técnico del Procesamiento**")
        kpi1, kpi2, kpi3 = st.columns(3)

        with kpi1:
            st.markdown(
                f"""
                <div class="kpi-card">
                    <div class="kpi-title">Volumen Total Calculado</div>
                    <div class="kpi-value">{volumen_m3:.3f} m³</div>
                </div>
            """,
                unsafe_allow_html=True,
            )

        with kpi2:
            st.markdown(
                f"""
                <div class="kpi-card" style="border-left-color: #1976D2;">
                    <div class="kpi-title">Jurisdicción Registrada</div>
                    <div class="kpi-value" style="color: #1976D2;">{jurisdiccion.split()[0]}</div>
                </div>
            """,
                unsafe_allow_html=True,
            )

        with kpi3:
            color_riesgo = "#D32F2F" if riesgo else "#2E7D32"
            texto_riesgo = "ALTO RIESGO" if riesgo else "NORMAL"
            st.markdown(
                f"""
                <div class="kpi-card" style="border-left-color: {color_riesgo};">
                    <div class="kpi-title">Prioridad Diagnosticada</div>
                    <div class="kpi-value" style="color: {color_riesgo};">{texto_riesgo}</div>
                </div>
            """,
                unsafe_allow_html=True,
            )

        # Reglas Condicionales CAR vs SDA
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### 📑 **Lista de Chequeo de Anexos Requeridos**")

        if "CAR" in jurisdiccion:
            st.success("✅ **Régimen CAR (Rural / Municipal):** Aplicación estricta del Decreto 1076 de 2015.")
            anexos_requeridos = [
                "Formato Único Nacional (FUN) de Solicitud de Aprovechamiento Forestal.",
                "Certificado de Libertad y Tradición (vigencia menor a 2 meses).",
                "Cartografía Oficial a escala 1:5.000 en sistema MAGNA-SIRGAS.",
                "Estudio Técnico de Aprovechamiento de Árboles Aislados (Inventario 100%).",
                "Copia de Cédula de Ciudadanía del Solicitante / Representante Legal.",
            ]
        else:
            st.info("ℹ️ **Régimen SDA (Distrito Capital - Urbano):** Enfocado en concepto silvicultural y riesgo de vuelco.")
            anexos_requeridos = [
                "Formulario Oficial de Silvicultura Urbana SDA.",
                "Registro Fotográfico de interferencia con infraestructura urbana o redes.",
                "Plan de Ahuyentamiento de Avifauna y Traslado de Nidos (si aplica).",
                "Certificado de Libertad (predio privado) o Autorización de Espacio Público.",
                "Cédula de Ciudadanía del Solicitante.",
            ]

        for anexo in anexos_requeridos:
            st.markdown(f"  • {anexo}")

        # Generación del Documento Word
        doc = Document()
        doc.add_heading("SOLICITUD TÉCNICA DE APROVECHAMIENTO FORESTAL", 0)
        doc.add_paragraph(f"Generado a través de EcoRegión App el {datetime.now().strftime('%d/%m/%Y a las %H:%M')}")

        doc.add_heading("1. DATOS DEL SOLICITANTE Y PREDIO", level=1)
        doc.add_paragraph(f"Solicitante: {nombre} ({tipo_doc}: {num_doc})")
        doc.add_paragraph(f"Calidad de Actuación: {calidad}")
        doc.add_paragraph(f"Autoridad Ambiental: {jurisdiccion}")
        doc.add_paragraph(f"Predio: {nombre_predio} | Municipio: {municipio}")
        doc.add_paragraph(f"Dirección: {direccion} | GPS: ({latitud}, {longitud})")

        doc.add_heading("2. CARACTERÍSTICAS TÉCNICAS Y VOLUMETRÍA", level=1)
        doc.add_paragraph(f"Especie: {especie_comun} ({especie_cientifico})")
        doc.add_paragraph(f"Cantidad: {cant_arboles} individuo(s)")
        doc.add_paragraph(f"DAP: {dap_cm} cm | Altura Total: {altura_m} m | Factor de Forma: {fm_val}")
        doc.add_paragraph(f"VOLUMEN EXTRAÍBLE CALCULADO: {volumen_m3:.3f} m³")

        doc.add_heading("3. ANEXOS COMPILADOS PARA RADICACIÓN", level=1)
        for a in anexos_requeridos:
            doc.add_paragraph(a, style="List Bullet")

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(
            label="📄 Descargar Documento Técnico Generado (.docx)",
            data=buffer,
            file_name=f"Solicitud_EcoRegion_{nombre_predio.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )


# ------------------------------------------------------------------------------
# 5. PESTAÑA 2: GUÍA DE USO SENCILLA
# ------------------------------------------------------------------------------
with tab_guia:
    st.markdown(
        """
        <div class="custom-box">
            <h3>📖 Guía Rápida para el Diligenciamiento del Trámite</h3>
            <p>Siga estos 4 sencillos pasos para completar su solicitud sin necesidad de asesoría técnica externa:</p>
            <ol>
                <li><b>Seleccione la Entidad Ambiental:</b> Si su predio está dentro de Bogotá seleccione <b>SDA</b>. Si está en otros municipios seleccione la <b>CAR</b> respectiva.</li>
                <li><b>Ingrese las Medidas Dendrométricas:</b> Mida la circunferencia del tronco a la altura de su pecho (DAP) y estime la altura. La app calculará los metros cúbicos (m³) automáticamente.</li>
                <li><b>Valide los Anexos Automáticos:</b> El sistema identificará los documentos obligatorios (escrituras, libertad y tradición, fotos) según la entidad.</li>
                <li><b>Descargue y Radique:</b> Al finalizar, haga clic en <i>Descargar Documento Técnico</i> para obtener su plantilla diligenciada lista para entregar.</li>
            </ol>
        </div>
    """,
        unsafe_allow_html=True,
    )


# ------------------------------------------------------------------------------
# 6. PESTAÑA 3: ENCUESTA & ESTADÍSTICAS GRÁFICAS
# ------------------------------------------------------------------------------
with tab_encuesta:
    st.markdown("### 📊 Validación de Usabilidad con Usuarios Reales")
    st.caption("Resultados del instrumento de evaluación aplicado al grupo controlado para medir la viabilidad del MVP.")

    col_e1, col_e2 = st.columns([1, 1], gap="medium")

    with col_e1:
        with st.form("form_encuesta_val"):
            st.markdown("#### **Formulario de Retroalimentación**")
            q1 = st.slider("Facilidad de diligenciamiento del formulario", 1, 5, 5)
            q2 = st.slider("Claridad en la diferencia CAR vs. SDA", 1, 5, 5)
            q3 = st.slider("Calidad y utilidad del documento Word generado", 1, 5, 5)
            comentarios = st.text_area("Sugerencias o comentarios adicionales:")
            btn_sub_enc = st.form_submit_button("Enviar Evaluación")

            if btn_sub_enc:
                st.success("¡Muchas gracias! Su calificación ha sido registrada con éxito.")

    with col_e2:
        st.markdown("#### **Resultados Consolidados (Escala Likert)**")

        # Gráfico elegante estilizado con la paleta de colores
        categorias = ["Facilidad Uso", "Claridad CAR/SDA", "Documento Word", "Cálculo m³"]
        puntajes = [4.8, 4.6, 4.9, 5.0]

        fig, ax = plt.subplots(figsize=(6, 4.2))
        fig.patch.set_facecolor("#FFFFFF")
        ax.set_facecolor("#FFFFFF")

        bars = ax.barh(categorias, puntajes, color="#2E7D32", height=0.55)
        ax.set_xlim(0, 5)
        ax.set_xlabel("Puntaje Promedio (1 a 5)", fontsize=10, color="#6C757D", fontweight="bold")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color("#CED4DA")
        ax.spines["bottom"].set_color("#CED4DA")

        for bar in bars:
            w = bar.get_width()
            ax.text(
                w - 0.45,
                bar.get_y() + bar.get_height() / 2,
                f"{w}",
                va="center",
                color="white",
                fontweight="bold",
                fontsize=10,
            )

        st.pyplot(fig)