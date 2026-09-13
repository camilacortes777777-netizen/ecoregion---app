import io
import math
from datetime import datetime

import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from docx import Document

# ------------------------------------------------------------------------------
# 1. CONFIGURACIÓN Y CSS GLOBAL DE ALTA VISIBILIDAD
# ------------------------------------------------------------------------------
st.set_page_config(
    page_title="EcoRegión - Soluciones Ambientales",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Insertar justo después de st.set_page_config(...)
st.markdown("""
    <style>
    /* 1. FORZAR FONDO BLANCO Y TEXTO OSCURO EN TODA LA APP */
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
        background-color: #FFFFFF !important;
        color: #1F2937 !important;
    }
    
    /* 2. BARRA LATERAL (SIDEBAR) EN GRIS MUY CLARO */
    [data-testid="stSidebar"] {
        background-color: #F8F9FA !important;
        border-right: 1px solid #E5E7EB !important;
    }
    [data-testid="stSidebar"] * {
        color: #1F2937 !important;
    }

    /* 3. TÍTULOS Y TEXTOS SIEMPRE VISIBLES */
    h1, h2, h3, h4, h5, h6, p, label, span, div, caption {
        color: #1F2937 !important;
    }
    
    /* Subtítulos y textos secundarios */
    .stCaption, caption {
        color: #4B5563 !important;
    }

    /* 4. CAMPOS DE TEXTO E INPUTS CON FONDO BLANCO Y BORDE DEFINIDO */
    input, textarea, div[data-baseweb="select"] > div {
        background-color: #FFFFFF !important;
        color: #1F2937 !important;
        border: 1px solid #D1D5DB !important;
        border-radius: 6px !important;
    }
    
    /* Menús desplegables (Selectbox) */
    div[data-baseweb="popover"], div[role="listbox"], li[role="option"] {
        background-color: #FFFFFF !important;
        color: #1F2937 !important;
    }
    li[role="option"]:hover {
        background-color: #E8F5E9 !important;
        color: #2E7D32 !important;
    }

    /* 5. BOTONES PRINCIPALES EN VERDE CORPORATIVO */
    .stButton>button, .stFormSubmitButton>button {
        background-color: #2E7D32 !important;
        color: #FFFFFF !important;
        font-weight: bold !important;
        border-radius: 8px !important;
        border: none !important;
        padding: 0.5rem 1rem !important;
    }
    .stButton>button:hover, .stFormSubmitButton>button:hover {
        background-color: #1B5E20 !important;
        color: #FFFFFF !important;
    }
    .stButton>button *, .stFormSubmitButton>button * {
        color: #FFFFFF !important;
    }

    /* 6. PESTAÑAS (TABS) */
    button[data-baseweb="tab"] {
        color: #4B5563 !important;
        font-weight: 600 !important;
    }
    button[aria-selected="true"] {
        color: #2E7D32 !important;
        border-bottom-color: #2E7D32 !important;
    }
    </style>
""", unsafe_allow_html=True)
# ------------------------------------------------------------------------------
# BARRA LATERAL (SIDEBAR): LOGO, LEMA E INFORMACIÓN INSTITUCIONAL
# ------------------------------------------------------------------------------
with st.sidebar:
    # 1. LOGO DE LA EMPRESA
    # Asegúrate de guardar la imagen en la misma carpeta de app.py
    st.image("LOGO.png", use_container_width=True)

    # 2. LEMA O FRASE INSTITUCIONAL
    # Cambia el texto entre comillas por tu frase exacta
    st.markdown(
        """
        <div style="text-align: center; margin-top: -10px; margin-bottom: 15px;">
            <span style="font-size: 13px; font-style: italic; color: #2E7D32; font-weight: 600;">
                “Permisos forestales sin complicaciones.”
            </span>
        </div>
        """, 
        unsafe_allow_html=True
    )
    
    st.divider()

    # 3. INFORMACIÓN COMPLEMENTARIA
    st.markdown("### 📌 **Panel de Control**")
    st.info("Herramienta oficial para la automatización de salvoconductos y trámites de aprovechamiento forestal.")
    
    st.divider()
    st.caption("🌿 **ECO REGIÓN S.A.S.** | Versión 1.0 MVP")
# ------------------------------------------------------------------------------
# 3. ENCABEZADO
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

tab_app, tab_guia, tab_encuesta = st.tabs(
    [
        "📝 Formulario de Solicitud",
        "📖 Guía de Uso del Trámite",
        "📊 Módulo de Validación & Usabilidad",
    ]
)


# ------------------------------------------------------------------------------
# 4. PESTAÑA 1: FORMULARIO
# ------------------------------------------------------------------------------
with tab_app:
    st.markdown("### 📝 Registro del Proyecto Ambiental")
    st.caption("Diligencie los campos requeridos para estructurar automáticamente el documento técnico.")

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

    if submit_btn or "procesado" in st.session_state:
        st.session_state["procesado"] = True
        st.markdown("---")

        dap_m = dap_cm / 100.0
        fm_val = float(forma_fuste.split("(")[1].replace(")", ""))
        volumen_m3 = (math.pi / 4) * (dap_m**2) * altura_m * fm_val * cant_arboles

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
                    <div class="kpi-value" style="color: #1976D2 !important;">{jurisdiccion.split()[0]}</div>
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
                    <div class="kpi-value" style="color: {color_riesgo} !important;">{texto_riesgo}</div>
                </div>
            """,
                unsafe_allow_html=True,
            )

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
            st.write(f"• {anexo}")

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
# 5. PESTAÑA 2: GUÍA DE USO (ARREGLADO VISUALMENTE CON NATIVOS DE STREAMLIT)
# ------------------------------------------------------------------------------
with tab_guia:
    st.markdown("### 📖 Guía Rápida para el Diligenciamiento del Trámite")
    st.write(
        "Siga estos 4 sencillos pasos para completar su solicitud sin necesidad de asesoría técnica externa:"
    )

    st.info(
        """
    **1. Seleccione la Entidad Ambiental:**  
    Si su predio está dentro de Bogotá D.C. seleccione **SDA**. Si se encuentra en otros municipios de Cundinamarca o Boyacá, seleccione la **CAR** correspondiente.
    """
    )

    st.info(
        """
    **2. Ingrese las Medidas Dendrométricas:**  
    Mida el grosor del tronco a la altura de su pecho (DAP en centímetros) y la altura total aproximada en metros. La aplicación calculará los metros cúbicos ($m^3$) automáticamente.
    """
    )

    st.info(
        """
    **3. Valide los Anexos Automáticos:**  
    El sistema identificará la lista de chequeo obligatoria (Certificados de Libertad, Fotos de interferencia, Plan de Manejo de Fauna) según la entidad seleccionada.
    """
    )

    st.success(
        """
    **4. Descargue y Radique:**  
    Al finalizar, haga clic en **Descargar Documento Técnico** para obtener su archivo estructurado en formato Word (.docx), listo para entregar ante la autoridad ambiental.
    """
    )


from streamlit_gsheets import GSheetsConnection

# ------------------------------------------------------------------------------
# 6. PESTAÑA 3: ENCUESTA CONECTADA A GOOGLE SHEETS
# ------------------------------------------------------------------------------
with tab_encuesta:
    st.markdown("### 📊 Validación de Usabilidad con Usuarios Reales")
    st.caption("Las respuestas registradas se guardan en tiempo real en nuestra base de datos de Google Sheets.")

    # PEGA AQUÍ TU ENLACE REAL DE GOOGLE SHEETS
    URL_SHEET = "https://docs.google.com/spreadsheets/d/17rVqZExih7E-dhfE3Vg_4tNqlTx5gCg5azDvzerys94/edit?usp=sharing"

    col_e1, col_e2 = st.columns([1, 1], gap="medium")

    # Inicializar conexión
    conn = st.connection("gsheets", type=GSheetsConnection)

    with col_e1:
        with st.form("form_encuesta_val"):
            st.markdown("#### **Formulario de Retroalimentación**")
            q1 = st.slider("1. Facilidad de diligenciamiento del formulario", 1, 5, 5)
            q2 = st.slider("2. Claridad en la diferencia CAR vs. SDA", 1, 5, 5)
            q3 = st.slider("3. Calidad y utilidad del documento Word generado", 1, 5, 5)
            comentarios = st.text_area("Sugerencias o comentarios adicionales:")
            btn_sub_enc = st.form_submit_button("☁️ Enviar a Google Sheets")

            if btn_sub_enc:
                try:
                    # Leer datos actuales
                    df_actual = conn.read(spreadsheet=URL_SHEET, ttl=0)
                    
                    # Estructurar la nueva respuesta
                    nueva_fila = pd.DataFrame([{
                        "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Facilidad_Uso": q1,
                        "Claridad_CAR_SDA": q2,
                        "Calidad_Word": q3,
                        "Comentarios": comentarios
                    }])
                    
                    # Limpiar nulos si la hoja está nueva
                    if df_actual is None or df_actual.empty:
                        df_actualizado = nueva_fila
                    else:
                        df_actualizado = pd.concat([df_actual, nueva_fila], ignore_index=True)
                    
                    # Actualizar hoja
                    conn.update(spreadsheet=URL_SHEET, data=df_actualizado)
                    st.success("¡Excelente! Tu calificación ha sido registrada en Google Sheets en tiempo real.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al conectar con Google Sheets: {e}")

    with col_e2:
        st.markdown("#### **Resultados Acumulados en Tiempo Real**")
        
        try:
            # Consultar datos en vivo especificando la URL
            df_respuestas = conn.read(spreadsheet=URL_SHEET, ttl=0)
            
            if df_respuestas is not None and not df_respuestas.empty:
                # Filtrar filas vacías si las hay
                df_respuestas = df_respuestas.dropna(subset=["Facilidad_Uso"])
                total_respuestas = len(df_respuestas)

                st.metric("Total de Usuarios Encuestados", f"{total_respuestas} respuestas")

                if total_respuestas > 0:
                    prom_q1 = pd.to_numeric(df_respuestas["Facilidad_Uso"]).mean()
                    prom_q2 = pd.to_numeric(df_respuestas["Claridad_CAR_SDA"]).mean()
                    prom_q3 = pd.to_numeric(df_respuestas["Calidad_Word"]).mean()

                    categorias = ["Facilidad Uso", "Claridad CAR/SDA", "Documento Word"]
                    puntajes = [round(prom_q1, 2), round(prom_q2, 2), round(prom_q3, 2)]

                    # Gráfico con promedios reales
                    fig, ax = plt.subplots(figsize=(6, 4))
                    fig.patch.set_facecolor('#FFFFFF')
                    ax.set_facecolor('#FFFFFF')

                    bars = ax.barh(categorias, puntajes, color='#2E7D32', height=0.5)
                    ax.set_xlim(0, 5)
                    ax.set_xlabel("Promedio Real en la Nube (1 a 5)", fontsize=10, color='#6C757D', fontweight='bold')
                    ax.spines['top'].set_visible(False)
                    ax.spines['right'].set_visible(False)

                    for bar in bars:
                        w = bar.get_width()
                        ax.text(w - 0.4, bar.get_y() + bar.get_height()/2, f"{w:.1f}", 
                                va='center', color='white', fontweight='bold', fontsize=10)

                    st.pyplot(fig)
                else:
                    st.info("Aún no hay respuestas guardadas. ¡Sé el primero en calificar!")
            else:
                st.info("Aún no hay respuestas en la hoja de Google Sheets. ¡Sé el primero en calificar!")
        except Exception as e:
            st.error(f"No se pudo cargar la vista previa: {e}")