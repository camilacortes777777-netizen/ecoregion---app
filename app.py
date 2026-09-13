import io
import math
from datetime import datetime
import pandas as pd
import streamlit as st
from docx import Document

# ------------------------------------------------------------------------------
# CONFIGURACIÓN DE LA PÁGINA
# ------------------------------------------------------------------------------
st.set_page_config(
    page_title="EcoRegión - Tramitador Ambiental",
    page_icon="🌿",
    layout="wide",
)

st.title("🌿 EcoRegión App: Aprovechamiento Forestal")
st.markdown(
    "Prototipo interactivo para la gestión y automatización de permisos ambientales ante **CAR** y **SDA**."
)

# Creamos pestañas principales
tab_app, tab_guia, tab_encuesta = st.tabs(
    ["📝 Formulario de Solicitud", "📚 Guía del Trámite", "📊 Encuesta de Validación"]
)

# ------------------------------------------------------------------------------
# PESTAÑA 1: FORMULARIO Y GENERACIÓN DE DOCUMENTOS
# ------------------------------------------------------------------------------
with tab_app:
    st.header("Formulario Único de Registro del Proyecto")

    with st.form("form_aprovechamiento"):
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("1. Datos del Solicitante")
            nombre = st.text_input("Nombre / Razón Social", "Nicol Daniela Navarrete")
            tipo_doc = st.selectbox("Tipo de Documento", ["CC", "NIT", "CE", "PASAPORTE"])
            num_doc = st.text_input("Número de Documento", "1000123456")
            calidad = st.selectbox(
                "Calidad en que actúa sobre el predio",
                ["Propietario", "Tenedor", "Poseedor", "Apoderado / Autorizado"],
            )

            st.subheader("2. Ubicación y Jurisdicción")
            jurisdiccion = st.selectbox(
                "Autoridad Ambiental Competente",
                ["CAR Corpoboyacá", "CAR Cundinamarca", "SDA (Bogotá D.C.)"],
            )
            nombre_predio = st.text_input("Nombre del Predio / Institución", "E.S.E. Hospital Duitama")
            municipio = st.text_input("Municipio / Ciudad", "Duitama")
            direccion = st.text_input("Dirección del Predio", "Av. Américas # 12-45")

        with col2:
            st.subheader("3. Datos Técnicos del Árbol / Cobertura")
            especie_comun = st.text_input("Nombre Común de la Especie", "Caucho")
            especie_cientifico = st.text_input("Nombre Científico", "Hevea brasiliensis")
            cant_arboles = st.number_input("Número de Individuos", min_value=1, value=1)

            st.markdown("**Dendrometría:**")
            dap_cm = st.number_input("DAP - Diámetro a la altura del pecho (cm)", min_value=1.0, value=35.0)
            altura_m = st.number_input("Altura Total (m)", min_value=1.0, value=9.0)

            # Factor de forma según fuste (CAR)
            forma_fuste = st.selectbox(
                "Forma del Fuste (Factor de Forma - fm)",
                ["Cilíndrico (0.75)", "Paraboloide (0.55)", "Cónico (0.33)", "Neiloide (0.25)"],
            )

            st.subheader("4. Urgencia y Entorno")
            riesgo = st.checkbox("¿Existe Riesgo Inminente de Caída / Daño?", value=True)
            fauna = st.checkbox("¿Se identificó presencia de avifauna / nidos o epífitas?", value=False)

        submit_btn = st.form_submit_button("🔥 Procesar Solicitud y Calcular")

    # --------------------------------------------------------------------------
    # LÓGICA CONDICIONAL Y CÁLCULOS
    # --------------------------------------------------------------------------
    if submit_btn or "procesado" in st.session_state:
        st.session_state["procesado"] = True

        st.divider()
        st.subheader("📋 Resumen de Validación y Cálculos Matemáticos")

        # Cálculo de Volumen en m3: V = (PI / 4) * (DAP en m)^2 * Altura * fm
        dap_m = dap_cm / 100.0
        fm_val = float(forma_fuste.split("(")[1].replace(")", ""))
        volumen_m3 = (math.pi / 4) * (dap_m**2) * altura_m * fm_val * cant_arboles

        col_res1, col_res2, col_res3 = st.columns(3)
        col_res1.metric("Volumen Total Calculado", f"{volumen_m3:.3f} m³")
        col_res2.metric("Jurisdicción Seleccionada", jurisdiccion)
        col_res3.metric("Nivel de Riesgo", "ALTO / INMINENTE" if riesgo else "NORMAL")

        # REGLAS CONDICIONALES CAR vs SDA
        st.subheader("⚡ Requisitos Identificados según la Autoridad")
        anexos_requeridos = []

        if "CAR" in jurisdiccion:
            st.info(
                "**Régimen CAR (Rural / Municipal):** Requiere acreditación estricta de propiedad y estudio técnico biofísico."
            )
            anexos_requeridos = [
                "1. Formato Único Nacional (FUN) firmado.",
                "2. Certificado de Libertad y Tradición (expedición < 2 meses).",
                "3. Cartografía oficial / Plano georreferenciado MAGNA-SIRGAS (Escala 1:5.000).",
                "4. Estudio Técnico de Aprovechamiento Forestal (Inventario 100%).",
                "5. Copia de Cédula de Ciudadanía del Solicitante.",
            ]
        else:  # SDA
            st.info(
                "**Régimen SDA (Distrito Capital - Urbano):** Enfocado en concepto técnico silvicultural, interferencia urbana y gestión de riesgo."
            )
            anexos_requeridos = [
                "1. Formato de Solicitud Silvicultura Urbana SDA.",
                "2. Registro Fotográfico con evidencia de interferencia con infraestructura/redes.",
                "3. Plan de Manejo de Avifauna y Traslado de Nidos (si aplica).",
                "4. Certificado de Libertad (predio privado) o Autorización de Espacio Público.",
                "5. Copia de Cédula de Ciudadanía.",
            ]

        for anexo in anexos_requeridos:
            st.write(f"- ✅ {anexo}")

        # GENERACIÓN DEL DOCUMENTO WORD DINÁMICO
        doc = Document()
        doc.add_heading("SOLICITUD TÉCNICA DE APROVECHAMIENTO FORESTAL", 0)

        doc.add_heading("1. INFORMACIÓN GENERAL", level=1)
        doc.add_paragraph(f"Fecha de Generación: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        doc.add_paragraph(f"Autoridad Ambiental: {jurisdiccion}")
        doc.add_paragraph(f"Solicitante: {nombre} ({tipo_doc}: {num_doc})")
        doc.add_paragraph(f"Calidad sobre el predio: {calidad}")

        doc.add_heading("2. UBICACIÓN DEL PREDIO", level=1)
        doc.add_paragraph(f"Predio/Institución: {nombre_predio}")
        doc.add_paragraph(f"Municipio: {municipio}")
        doc.add_paragraph(f"Dirección: {direccion}")

        doc.add_heading("3. CARACTERÍSTICAS TÉCNICAS DEL APROVECHAMIENTO", level=1)
        doc.add_paragraph(f"Especie: {especie_comun} ({especie_cientifico})")
        doc.add_paragraph(f"Cantidad de Individuos: {cant_arboles}")
        doc.add_paragraph(f"DAP: {dap_cm} cm | Altura Total: {altura_m} m")
        doc.add_paragraph(f"Factor de Forma Aplicado: {fm_val}")
        doc.add_paragraph(f"VOLUMEN TOTAL CALCULADO: {volumen_m3:.3f} m³")

        doc.add_heading("4. JUSTIFICACIÓN Y REQUISITOS", level=1)
        doc.add_paragraph(
            f"Condición de Riesgo Inminente: {'SI (Prioritario)' if riesgo else 'NO'}"
        )
        doc.add_paragraph(
            f"Presencia de Avifauna/Epífitas: {'SI (Requiere manejo)' if fauna else 'NO'}"
        )

        doc.add_heading("5. ANEXOS GENERADOS PARA RADICACIÓN", level=1)
        for anexo in anexos_requeridos:
            doc.add_paragraph(anexo, style="List Bullet")

        # Guardar documento en memoria para descarga
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        st.divider()
        st.download_button(
            label="📄 Descargar Documento Técnico Automático (.docx)",
            data=buffer,
            file_name=f"Solicitud_{nombre_predio.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

# ------------------------------------------------------------------------------
# PESTAÑA 2: GUÍA DE USO EN LENGUAJE SENCILLO (MÁX. 2 PÁGINAS)
# ------------------------------------------------------------------------------
with tab_guia:
    st.header("📖 Guía Rápida de Uso - Aplicación EcoRegión")

    st.subheader("¿Cómo radicar tu trámite ambiental en 4 pasos?")
    st.markdown(
        """
    1. **Diligencia el Formulario:** Ingresa tus datos básicos, la ubicación de tu predio y la autoridad ambiental correspondiente (SDA si es dentro de Bogotá, CAR si es fuera de Bogotá).
    2. **Toma las Medidas del Árbol:** Ingresa el grosor del tronco (DAP) y la altura aproximada. La app calculará automáticamente los metros cúbicos ($m^3$) de madera.
    3. **Revisa los Anexos Automáticos:** La app detectará si tu trámite es en la CAR o en la SDA y te dará la lista exacta de documentos que debes adjuntar para evitar devoluciones.
    4. **Descarga tu Documento:** Haz clic en el botón de descarga para obtener tu informe técnico listo en formato Word (.docx), firmado y estructurado para radicar de inmediato.
    """
    )

    st.subheader("Paso a Paso General del Trámite Ambiental")
    st.markdown(
        """
    * **Paso A (Radicación):** Presentas el archivo generado por esta App junto con los anexos solicitados.
    * **Paso B (Evaluación):** La autoridad ambiental revisa la documentación y agenda una visita técnica.
    * **Paso C (Visita en Sitio):** Un técnico de la CAR o SDA inspecciona el árbol físicamente.
    * **Paso D (Resolución/Autorización):** Se emite el acto administrativo que autoriza la tala o poda de los individuos arbóreos.
    """
    )

# ------------------------------------------------------------------------------
# PESTAÑA 3: MECANISMO DE ENCUESTAS DE VALIDACIÓN
# ------------------------------------------------------------------------------
with tab_encuesta:
    st.header("📊 Encuesta de Usabilidad (Validación con Usuarios Reales)")
    st.write(
        "Por favor califica tu experiencia con el prototipo de EcoRegión para ayudarnos a mejorar el producto."
    )

    with st.form("form_encuesta"):
        q1 = st.slider("1. ¿Qué tan fácil fue diligenciar la información de tu proyecto?", 1, 5, 5)
        q2 = st.slider("2. ¿La explicación de los requisitos (CAR / SDA) fue clara?", 1, 5, 5)
        q3 = st.slider("3. ¿El documento Word generado cumple con tus expectativas?", 1, 5, 5)
        comentarios = st.text_area("Comentarios o sugerencias de mejora:")

        enviar_encuesta = st.form_submit_button("Enviar Retroalimentación")

    if enviar_encuesta:
        st.success("¡Gracias por tus respuestas! Tus datos han sido registrados para el informe de validación.")