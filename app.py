import io
import json
import math
import hashlib
import os
import re
import shutil
import zipfile
from datetime import datetime
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from docx import Document

# =============================================================================
# ECORREGION APP v2
# MVP de gestión de permisos de aprovechamiento forestal
# - Navegación tipo panel lateral
# - Solicitud por etapas
# - Requisitos dinámicos por autoridad/modalidad
# - Descarga de formatos oficiales desde carpeta local + enlace fuente
# - Carga de documentos diligenciados
# - Expediente ZIP
# - Cálculo volumétrico por árbol
# - Encuesta de validación
# =============================================================================

APP_TITLE = "EcoRegión App"
DATA_DIR = Path("data")
DOCS_DIR = Path("documentos")
CASE_DIR = DATA_DIR / "expedientes"
SURVEY_FILE = DATA_DIR / "respuestas_encuesta.csv"

for folder in [DATA_DIR, DOCS_DIR, CASE_DIR]:
    folder.mkdir(parents=True, exist_ok=True)

st.set_page_config(
    page_title="EcoRegión App",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================================================
# ESTILOS
# =============================================================================
st.markdown(
    """
    <style>
    :root {
        --eco-dark: #123c2a;
        --eco: #1d6b4b;
        --eco-light: #dff4e8;
        --eco-soft: #eef8f2;
        --ink: #1d2a24;
        --muted: #5b6c63;
        --border: #d8e5de;
        --warning: #a36b00;
        --danger: #a52f2f;
    }

    .stApp { background: #f6faf8; color: var(--ink); }
    [data-testid="stHeader"] { background: transparent; }
    [data-testid="stSidebar"] {
        background: var(--eco-dark) !important;
        border-right: 1px solid #0d2e20 !important;
    }
    [data-testid="stSidebar"] * { color: #f4fbf6 !important; }
    [data-testid="stSidebar"] .stRadio > div { gap: 0.18rem; }
    [data-testid="stSidebar"] .stRadio label {
        padding: 0.45rem 0.55rem;
        border-radius: 0.5rem;
    }
    [data-testid="stSidebar"] .stRadio label:hover { background: rgba(255,255,255,0.08); }

    h1, h2, h3, h4 { color: var(--eco-dark) !important; }
    p, label, span, div { color: var(--ink); }
    .stCaption { color: var(--muted) !important; }

    input, textarea, [data-baseweb="select"] > div {
        background: #ffffff !important;
        color: var(--ink) !important;
        border-color: var(--border) !important;
    }
    [data-baseweb="popover"], [role="listbox"], [role="option"] {
        background: #ffffff !important;
        color: var(--ink) !important;
    }

    .stButton > button, .stFormSubmitButton > button {
        background: var(--eco) !important;
        color: #fff !important;
        border: 0 !important;
        border-radius: 0.6rem !important;
        font-weight: 700 !important;
    }
    .stButton > button:hover, .stFormSubmitButton > button:hover {
        background: #145139 !important;
        color: #fff !important;
    }

    .hero {
        background: linear-gradient(135deg, #123c2a 0%, #1d6b4b 100%);
        color: white;
        border-radius: 1rem;
        padding: 1.35rem 1.5rem;
        margin-bottom: 1rem;
        box-shadow: 0 8px 30px rgba(18,60,42,.14);
    }
    .hero h1, .hero p { color: white !important; margin: 0; }
    .hero h1 { font-size: 2rem; }
    .hero p { opacity: .88; margin-top: .25rem; }

    .card {
        background: #fff;
        border: 1px solid var(--border);
        border-radius: .9rem;
        padding: 1rem;
        margin-bottom: .9rem;
    }
    .kpi {
        background: #fff;
        border: 1px solid var(--border);
        border-left: 5px solid var(--eco);
        border-radius: .8rem;
        padding: .9rem 1rem;
    }
    .kpi .title { font-size: .82rem; color: var(--muted); }
    .kpi .value { font-size: 1.35rem; font-weight: 800; color: var(--eco-dark); }

    .doc-ok { color: #19643f; font-weight: 700; }
    .doc-warn { color: var(--warning); font-weight: 700; }
    .doc-miss { color: var(--danger); font-weight: 700; }
    .small { font-size: .87rem; color: var(--muted); }

    .sidebar-brand { text-align: center; padding: .2rem 0 1rem; }
    .sidebar-brand .mark {
        width: 58px; height: 58px; margin: 0 auto .55rem;
        background: #dff4e8; color: var(--eco-dark);
        border-radius: 1rem; display: flex; align-items: center; justify-content: center;
        font-size: 1.9rem; font-weight: 800;
    }
    .sidebar-brand .name { font-weight: 800; font-size: 1.1rem; }
    .sidebar-brand .tag { font-size: .78rem; opacity: .8; }
    </style>
    """,
    unsafe_allow_html=True,
)


# =============================================================================
# CONFIGURACIÓN DOCUMENTAL
# =============================================================================
def slug(text: str) -> str:
    value = re.sub(r"[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ]+", "_", text.strip())
    return value.strip("_").lower()


SOURCE_SDA_F1 = (
    "https://www.ambientebogota.gov.co/web/transparencia/lineamientos"
)
SOURCE_CAR_FUN = "https://www.car.gov.co/assets/v1/uploads/666b2c9e17175_2e929fc718.pdf"

REQUIREMENTS = {
    "SDA": {
        "label": "SDA (Bogotá D.C.)",
        "procedure_title": "Permiso o autorización para aprovechamiento forestal de árboles aislados",
        "procedure": "La información suministrada para el proyecto identifica 3 documentos oficiales de SDA para este trámite.",
        "documents": [
            {
                "id": "sda_f1",
                "name": "PM04-PR30-F1 Formulario solicitud manejo aprovechamiento forestal.xlsx",
                "local_filename": "PM04-PR30-F1 Formulario solicitud manejo aprovechamiento forestal.xlsx",
                "folder": "SDA",
                "source_url": SOURCE_SDA_F1,
                "note": "Formulario oficial indicado en la información suministrada.",
            },
            {
                "id": "sda_f2",
                "name": "PM04-PR30-F2 Recoleccion informacion silvicultural individuo ficha1.xls",
                "local_filename": "PM04-PR30-F2 Recoleccion informacion silvicultural individuo ficha1.xls",
                "folder": "SDA",
                "source_url": SOURCE_SDA_F1,
                "note": "Ficha 1 de recolección de información silvicultural.",
            },
            {
                "id": "sda_f3",
                "name": "PM04-PR30-F3 Ficha tecnica de registro Ficha 2.docx",
                "local_filename": "PM04-PR30-F3 Ficha tecnica de registro Ficha 2.docx",
                "folder": "SDA",
                "source_url": SOURCE_SDA_F1,
                "note": "Ficha técnica de registro.",
            },
        ],
    },
    "CAR": {
        "label": "CAR Cundinamarca",
        "procedure_title": "Aprovechamiento forestal de árboles aislados",
        "procedure": "La información suministrada indica obtener el Formato Único Nacional de Solicitud de Aprovechamiento Forestal y Manejo Sostenible de Flora Silvestre y los Productos Forestales No Maderables Nuevo/Prórroga (FUN), diligenciarlo y radicarlo en los lugares señalados por la Corporación.",
        "documents": [
            {
                "id": "car_fun",
                "name": "Formato Único Nacional de Solicitud de Aprovechamiento Forestal y Manejo Sostenible de Flora Silvestre y los Productos Forestales No Maderables Nuevo/Prórroga (FUN)",
                "local_filename": "CAR_FUN.pdf",
                "folder": "CAR_Cundinamarca",
                "source_url": SOURCE_CAR_FUN,
                "note": "Formato FUN indicado en la información suministrada.",
            },
        ],
    },
    "CORPO_PRIORITARIA": {
        "label": "Corpoboyacá — prioritaria / emergencia / obra",
        "procedure_title": "Árboles por solicitud prioritaria, tala de emergencia, construcción de obra pública o privada",
        "procedure": "Una vez reunidos los requisitos, se solicita la liquidación para el pago por servicios de evaluación ambiental y, efectuado el pago, se procede con la radicación de la documentación.",
        "documents": [
            {"id": "fun", "name": "Formato único nacional de solicitud de aprovechamiento forestal y manejo sostenible de flora silvestre y productos forestales no maderables.", "local_filename": "FUN_Corpoboyaca.pdf", "folder": "Corpoboyaca/prioritaria_emergencia_obra"},
            {"id": "fgr06a", "name": "Formato FGR-06 Parte A «Inventario Forestal al 100%» (incluir registro fotográfico y georreferenciación por individuo marcado y destinado a aprovechar, anexo en medio magnético).", "local_filename": "FGR-06_Parte_A.pdf", "folder": "Corpoboyaca/prioritaria_emergencia_obra"},
            {"id": "cedulas", "name": "Fotocopias de cédula del o los propietarios del predio o representante legal según aplique.", "local_filename": "", "folder": "Corpoboyaca/prioritaria_emergencia_obra"},
            {"id": "existencia", "name": "Certificado de existencia y representación legal (persona jurídica).", "local_filename": "", "folder": "Corpoboyaca/prioritaria_emergencia_obra"},
            {"id": "escritura", "name": "Fotocopia escritura(s) del predio objeto de aprovechamiento.", "local_filename": "", "folder": "Corpoboyaca/prioritaria_emergencia_obra"},
            {"id": "libertad", "name": "Certificado de libertad y tradición con fecha de expedición no mayor a treinta (30) días.", "local_filename": "", "folder": "Corpoboyaca/prioritaria_emergencia_obra"},
            {"id": "autorizacion", "name": "Autorización escrita del propietario del predio cuando se actúa como tenedor, si aplica.", "local_filename": "", "folder": "Corpoboyaca/prioritaria_emergencia_obra"},
            {"id": "fgr29", "name": "Formato FGR-29 «Auto-declaración costos de inversión y anual de operación».", "local_filename": "FGR-29.pdf", "folder": "Corpoboyaca/prioritaria_emergencia_obra"},
            {"id": "pago", "name": "Copia de recibo de consignación o factura de pago por servicios de evaluación ambiental a favor de Corpoboyacá.", "local_filename": "", "folder": "Corpoboyaca/prioritaria_emergencia_obra"},
        ],
    },
    "CORPO_NATIVA_GT50": {
        "label": "Corpoboyacá — nativas > 50 m³",
        "procedure_title": "Aprovechamiento de árboles aislados fuera de cobertura de bosque natural — nativas > 50 m³",
        "procedure": "La documentación suministrada señala que, reunidos los requisitos, se solicita la liquidación para el pago por servicios de evaluación ambiental y luego se radica la documentación.",
        "documents": [
            {"id": "fun", "name": "Formato único nacional de solicitud de aprovechamiento forestal y manejo sostenible de flora silvestre y productos forestales no maderables.", "local_filename": "FUN_Corpoboyaca.pdf", "folder": "Corpoboyaca/nativas_mayor_50m3"},
            {"id": "estudio", "name": "Estudio técnico para el aprovechamiento de árboles aislados (términos de referencia Corpoboyacá).", "local_filename": "Estudio_Tecnico_Arboles_Aislados.pdf", "folder": "Corpoboyaca/nativas_mayor_50m3"},
            {"id": "cedulas", "name": "Fotocopias de cédula del o los propietarios del predio o representante legal según aplique.", "local_filename": "", "folder": "Corpoboyaca/nativas_mayor_50m3"},
            {"id": "existencia", "name": "Certificado de existencia y representación legal (persona jurídica).", "local_filename": "", "folder": "Corpoboyaca/nativas_mayor_50m3"},
            {"id": "escritura", "name": "Fotocopia escritura(s) del predio objeto de aprovechamiento.", "local_filename": "", "folder": "Corpoboyaca/nativas_mayor_50m3"},
            {"id": "libertad", "name": "Certificado de libertad y tradición con fecha de expedición no mayor a treinta (30) días.", "local_filename": "", "folder": "Corpoboyaca/nativas_mayor_50m3"},
            {"id": "autorizacion", "name": "Autorización escrita del propietario del predio cuando se actúa como tenedor, si aplica.", "local_filename": "", "folder": "Corpoboyaca/nativas_mayor_50m3"},
            {"id": "fgr29", "name": "Formato FGR-29 «Auto declaración, costos de inversión y anual de operación».", "local_filename": "FGR-29.pdf", "folder": "Corpoboyaca/nativas_mayor_50m3"},
            {"id": "pago", "name": "Copia de recibo de consignación o factura de pago por servicios de evaluación ambiental a favor de Corpoboyacá.", "local_filename": "", "folder": "Corpoboyaca/nativas_mayor_50m3"},
        ],
    },
    "CORPO_MENOR50_EXOTICA": {
        "label": "Corpoboyacá — nativas < 50 m³ / exóticas",
        "procedure_title": "Nativas con volumen inferior a 50 m³ y especies exóticas o introducidas sin límite de volumen",
        "procedure": "La documentación suministrada señala que, reunidos los requisitos, se solicita la liquidación para el pago por servicios de evaluación ambiental y posteriormente se radica la documentación.",
        "documents": [
            {"id": "fun", "name": "Formato único Nacional de solicitud de aprovechamiento forestal y manejo sostenible de flora silvestre y productos forestales no maderables.", "local_filename": "FUN_Corpoboyaca.pdf", "folder": "Corpoboyaca/nativas_menor_50m3_exoticas"},
            {"id": "fgr06ab", "name": "Formato FGR-06 “Información para el aprovechamiento de árboles Aislados” Parte A y B.", "local_filename": "FGR-06_Parte_A_y_B.pdf", "folder": "Corpoboyaca/nativas_menor_50m3_exoticas"},
            {"id": "cedulas", "name": "Fotocopias de cédula del o los propietarios del predio o representante legal según aplique.", "local_filename": "", "folder": "Corpoboyaca/nativas_menor_50m3_exoticas"},
            {"id": "existencia", "name": "Certificado de existencia y representación Legal (persona jurídica).", "local_filename": "", "folder": "Corpoboyaca/nativas_menor_50m3_exoticas"},
            {"id": "escritura", "name": "Fotocopia escritura(s) del predio objeto de aprovechamiento.", "local_filename": "", "folder": "Corpoboyaca/nativas_menor_50m3_exoticas"},
            {"id": "libertad", "name": "Certificado de libertad y tradición con fecha de expedición no mayor a treinta (30) días.", "local_filename": "", "folder": "Corpoboyaca/nativas_menor_50m3_exoticas"},
            {"id": "autorizacion", "name": "Autorización escrita del propietario del predio cuando se actúa como tenedor, si aplica.", "local_filename": "", "folder": "Corpoboyaca/nativas_menor_50m3_exoticas"},
            {"id": "fgr29", "name": "Formato FGR-29 «Auto declaración, costos de inversión y anual de operación».", "local_filename": "FGR-29.pdf", "folder": "Corpoboyaca/nativas_menor_50m3_exoticas"},
            {"id": "pago", "name": "Copia de recibo de consignación o factura de pago por servicios de evaluación ambiental a favor de Corpoboyacá.", "local_filename": "", "folder": "Corpoboyaca/nativas_menor_50m3_exoticas"},
        ],
    },
    "CORPO_DOMESTICO": {
        "label": "Corpoboyacá — uso doméstico",
        "procedure_title": "Aprovechamiento de árboles aislados para uso doméstico",
        "procedure": "Aplica para predios privados. El destino de los productos es exclusivamente satisfacer necesidades domésticas y no se permite comercialización. El volumen máximo indicado es 10 m³ para especies nativas y 20 m³ para especies exóticas o introducidas. Según la información suministrada, no requiere pago por servicios de evaluación ambiental.",
        "documents": [
            {"id": "fun", "name": "Formato único Nacional de solicitud de aprovechamiento forestal y manejo sostenible de flora silvestre y productos forestales no maderables.", "local_filename": "FUN_Corpoboyaca.pdf", "folder": "Corpoboyaca/uso_domestico"},
            {"id": "fgr06a", "name": "Formato FGR-06 Parte A “inventario forestal al 100%” (incluir registro fotográfico y georreferenciación por individuo marcado y destinado a aprovechar, anexo en medio magnético).", "local_filename": "FGR-06_Parte_A.pdf", "folder": "Corpoboyaca/uso_domestico"},
            {"id": "cedulas", "name": "Fotocopias de cédula del o los propietarios del predio.", "local_filename": "", "folder": "Corpoboyaca/uso_domestico"},
            {"id": "escritura", "name": "Fotocopia de la Escritura Pública.", "local_filename": "", "folder": "Corpoboyaca/uso_domestico"},
            {"id": "libertad", "name": "Certificado de libertad con fecha de expedición no mayor a 30 días.", "local_filename": "", "folder": "Corpoboyaca/uso_domestico"},
        ],
    },
    "GENERAL": {
        "label": "General / sin autoridad definida",
        "procedure_title": "Prealistamiento general",
        "procedure": "No se suministró una lista oficial independiente para una categoría 'General'. Para no inventar requisitos legales, esta opción se usa para organizar información y documentos mientras se define la autoridad competente.",
        "documents": [],
    },
}

CORPO_OPTIONS = {
    "Prioritaria / emergencia / obra pública o privada": "CORPO_PRIORITARIA",
    "Nativas con volumen superior a 50 m³": "CORPO_NATIVA_GT50",
    "Nativas < 50 m³ / especies exóticas o introducidas": "CORPO_MENOR50_EXOTICA",
    "Uso doméstico": "CORPO_DOMESTICO",
}

AUTHORITY_OPTIONS = {
    "SDA (Bogotá D.C.)": "SDA",
    "CAR Cundinamarca": "CAR",
    "Corpoboyacá": "CORPO",
    "General / sin autoridad definida": "GENERAL",
}

ALLOWED_DOC_EXTENSIONS = {
    ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".csv", ".jpg", ".jpeg", ".png", ".webp"
}


def get_requirement_group():
    authority = st.session_state.get("authority_key", "GENERAL")
    if authority == "SDA":
        return "SDA"
    if authority == "CAR":
        return "CAR"
    if authority == "CORPO":
        return st.session_state.get("corpo_case", "CORPO_PRIORITARIA")
    return "GENERAL"


def get_requirements():
    return REQUIREMENTS[get_requirement_group()]


def current_case_id():
    return st.session_state.get("case_id") or "sin_expediente"


def case_folder() -> Path:
    folder = CASE_DIR / current_case_id()
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def doc_upload_folder() -> Path:
    folder = case_folder() / "documentos_diligenciados"
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def safe_file_name(filename: str) -> str:
    base = os.path.basename(filename)
    return re.sub(r"[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ._() -]", "_", base)


def local_template_path(req: dict) -> Path | None:
    filename = req.get("local_filename", "").strip()
    if not filename:
        return None
    path = DOCS_DIR / req["folder"] / filename
    return path if path.exists() else None


def init_state():
    defaults = {
        "page": "Inicio",
        "request_page": "Proyecto",
        "case_id": datetime.now().strftime("ER-%Y%m%d-%H%M%S"),
        "project_type": "Aprovechamiento forestal de árboles aislados",
        "authority_label": "SDA (Bogotá D.C.)",
        "authority_key": "SDA",
        "corpo_case": "CORPO_PRIORITARIA",
        "solicitud": {},
        "trees": [],
        "uploaded_docs": {},
        "followup_status": {},
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


init_state()


def ensure_tree_count(n: int):
    trees = st.session_state.setdefault("trees", [])
    while len(trees) < n:
        idx = len(trees) + 1
        trees.append(
            {
                "nombre_comun": "",
                "nombre_cientifico": "",
                "dap_cm": 0.0,
                "altura_m": 0.0,
                "factor_nombre": "Paraboloide (0.55)",
                "riesgo": False,
                "fauna": False,
                "epifitas": False,
                "observaciones": "",
                "foto_path": "",
                "foto_name": "",
            }
        )
    if len(trees) > n:
        del trees[n:]


def format_file_size(path: Path) -> str:
    if not path.exists():
        return "0 B"
    size = path.stat().st_size
    if size < 1024:
        return f"{size} B"
    if size < 1024 * 1024:
        return f"{size / 1024:.1f} KB"
    return f"{size / (1024 * 1024):.1f} MB"


def save_uploaded_file(uploaded, folder: Path, prefix: str = "") -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    safe = safe_file_name(uploaded.name)
    stem = Path(safe).stem
    suffix = Path(safe).suffix.lower()
    token = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    final = folder / f"{prefix}{stem}_{token}{suffix}"
    final.write_bytes(uploaded.getvalue())
    return final


def requirement_uploaded(req_id: str) -> list[Path]:
    values = st.session_state.get("uploaded_docs", {}).get(req_id, [])
    paths = [Path(p) for p in values if Path(p).exists()]
    return paths


def uploaded_signature(uploaded) -> str:
    data = uploaded.getvalue()
    return hashlib.sha256(data).hexdigest()


def store_requirement_uploads(req_id: str, uploaded_files):
    uploaded_registry = st.session_state.setdefault("uploaded_docs", {})
    saved = uploaded_registry.setdefault(req_id, [])
    seen = uploaded_registry.setdefault("__signatures__", {}).setdefault(req_id, set())
    for uploaded in uploaded_files:
        if not uploaded:
            continue
        ext = Path(uploaded.name).suffix.lower()
        if ext not in ALLOWED_DOC_EXTENSIONS:
            st.error(f"Formato no permitido: {uploaded.name}")
            continue
        signature = uploaded_signature(uploaded)
        if signature in seen:
            continue
        saved_path = save_uploaded_file(uploaded, doc_upload_folder(), prefix=f"{slug(req_id)}__")
        saved.append(str(saved_path))
        seen.add(signature)


def volume_for_tree(tree: dict) -> float:
    dap = float(tree.get("dap_cm", 0)) / 100.0
    height = float(tree.get("altura_m", 0))
    factor_text = tree.get("factor_nombre", "Paraboloide (0.55)")
    try:
        factor = float(factor_text.split("(")[1].split(")")[0])
    except (IndexError, ValueError):
        factor = 0.55
    if dap <= 0 or height <= 0:
        return 0.0
    return (math.pi / 4) * (dap ** 2) * height * factor


def total_volume() -> float:
    return sum(volume_for_tree(t) for t in st.session_state.get("trees", []))


def priority_text() -> str:
    trees = st.session_state.get("trees", [])
    if any(t.get("riesgo") for t in trees):
        return "CRÍTICA — Riesgo inminente"
    if any(t.get("fauna") or t.get("epifitas") for t in trees):
        return "ALTA — Restricción ambiental"
    return "ESTÁNDAR"


def create_technical_doc() -> bytes:
    doc = Document()
    doc.add_heading("EXPEDIENTE DE PREALISTAMIENTO — ECO REGIÓN APP", 0)
    doc.add_paragraph(
        f"Generado el {datetime.now().strftime('%d/%m/%Y %H:%M')} por EcoRegión App."
    )
    doc.add_paragraph(
        "Este archivo es un resumen técnico generado por el prototipo y no sustituye los formatos oficiales de la autoridad ambiental."
    )

    sol = st.session_state.get("solicitud", {})
    doc.add_heading("1. Datos del proyecto", level=1)
    doc.add_paragraph(f"Tipo de trámite: {st.session_state.get('project_type')}")
    doc.add_paragraph(f"Autoridad: {st.session_state.get('authority_label')}")
    if st.session_state.get("authority_key") == "CORPO":
        doc.add_paragraph(f"Modalidad Corpoboyacá: {REQUIREMENTS[get_requirement_group()]['label']}")

    doc.add_heading("2. Solicitante", level=1)
    for label in [
        ("Nombre / Razón social", "nombre"),
        ("Tipo de documento", "tipo_doc"),
        ("Número de documento", "num_doc"),
        ("Calidad sobre el predio", "calidad"),
        ("Correo", "correo"),
        ("Teléfono", "telefono"),
    ]:
        doc.add_paragraph(f"{label[0]}: {sol.get(label[1], '')}")

    doc.add_heading("3. Predio y ubicación", level=1)
    for label in [
        ("Nombre del predio / institución", "nombre_predio"),
        ("Municipio / ciudad", "municipio"),
        ("Dirección", "direccion"),
        ("Latitud", "latitud"),
        ("Longitud", "longitud"),
        ("Descripción de la intervención", "descripcion"),
    ]:
        doc.add_paragraph(f"{label[0]}: {sol.get(label[1], '')}")

    doc.add_heading("4. Árboles registrados", level=1)
    for i, tree in enumerate(st.session_state.get("trees", []), start=1):
        doc.add_heading(f"Individuo {i}", level=2)
        doc.add_paragraph(f"Nombre común: {tree.get('nombre_comun', '')}")
        doc.add_paragraph(f"Nombre científico: {tree.get('nombre_cientifico', '')}")
        doc.add_paragraph(f"DAP: {tree.get('dap_cm', '')} cm")
        doc.add_paragraph(f"Altura total: {tree.get('altura_m', '')} m")
        doc.add_paragraph(f"Factor de forma: {tree.get('factor_nombre', '')}")
        doc.add_paragraph(f"Volumen estimado: {volume_for_tree(tree):.3f} m³")
        doc.add_paragraph(f"Riesgo inminente: {'Sí' if tree.get('riesgo') else 'No'}")
        doc.add_paragraph(f"Avifauna: {'Sí' if tree.get('fauna') else 'No'}")
        doc.add_paragraph(f"Epífitas: {'Sí' if tree.get('epifitas') else 'No'}")
        if tree.get("observaciones"):
            doc.add_paragraph(f"Observaciones: {tree.get('observaciones')}")

    reqs = get_requirements()["documents"]
    doc.add_heading("5. Requisitos documentales detectados", level=1)
    for req in reqs:
        uploaded = requirement_uploaded(req["id"])
        state = "SUBIDO" if uploaded else "PENDIENTE"
        doc.add_paragraph(f"[{state}] {req['name']}", style="List Bullet")

    doc.add_heading("6. Resultado del procesamiento", level=1)
    doc.add_paragraph(f"Volumen total estimado: {total_volume():.3f} m³")
    doc.add_paragraph(f"Prioridad: {priority_text()}")
    doc.add_paragraph(f"Documentos cargados: {sum(len(v) for v in st.session_state.get('uploaded_docs', {}).values())}")

    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream.getvalue()


def build_case_zip() -> bytes:
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        technical = create_technical_doc()
        zf.writestr("01_Resumen_Tecnico_EcoRegion.docx", technical)

        reqs = get_requirements()["documents"]
        for req in reqs:
            for path in requirement_uploaded(req["id"]):
                zf.write(path, arcname=f"02_Documentos_Diligenciados/{path.name}")

        additional = case_folder() / "documentos_adicionales"
        if additional.exists():
            for path in additional.glob("*"):
                if path.is_file():
                    zf.write(path, arcname=f"03_Documentos_Adicionales/{path.name}")

        meta = {
            "case_id": current_case_id(),
            "generated_at": datetime.now().isoformat(),
            "authority": st.session_state.get("authority_label"),
            "requirements_group": get_requirement_group(),
            "total_volume_m3": round(total_volume(), 3),
            "priority": priority_text(),
        }
        zf.writestr("00_metadata.json", json.dumps(meta, ensure_ascii=False, indent=2))

    zip_buffer.seek(0)
    return zip_buffer.getvalue()


# =============================================================================
# SIDEBAR
# =============================================================================
with st.sidebar:
    st.markdown(
        """
        <div class='sidebar-brand'>
            <div class='mark'>🌿</div>
            <div class='name'>EcoRegión App</div>
            <div class='tag'>Gestión inteligente de permisos forestales</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.divider()

    main_options = [
        "Inicio",
        "Nueva solicitud",
        "Requisitos",
        "Autoridades",
        "Documentación",
        "Seguimiento",
        "Encuesta de validación",
        "Manual de uso",
    ]
    selected = st.radio("MENÚ PRINCIPAL", main_options, index=main_options.index(st.session_state["page"]))
    st.session_state["page"] = selected

    if selected == "Nueva solicitud":
        st.markdown("**Etapas de la solicitud**")
        request_options = ["Proyecto", "Solicitante", "Predio", "Árboles", "Validación"]
        req_page = st.radio(
            "",
            request_options,
            index=request_options.index(st.session_state["request_page"]),
            key="request_nav",
        )
        st.session_state["request_page"] = req_page

    st.divider()
    st.caption(f"Expediente: {current_case_id()}")
    st.caption("Versión 2.0 MVP")


# =============================================================================
# HERO
# =============================================================================
st.markdown(
    f"""
    <div class='hero'>
        <h1>🌿 {APP_TITLE}</h1>
        <p>Gestión estructurada y automatización del trámite de aprovechamiento forestal de árboles aislados.</p>
    </div>
    """,
    unsafe_allow_html=True,
)


# =============================================================================
# PÁGINA: INICIO
# =============================================================================
def page_inicio():
    st.subheader("Panel de control")
    req_group = get_requirement_group()
    req_count = len(get_requirements()["documents"])
    uploaded_count = sum(len(v) for k, v in st.session_state.get("uploaded_docs", {}).items() if k != "__signatures__")
    trees = len(st.session_state.get("trees", []))
    sol = st.session_state.get("solicitud", {})

    cols = st.columns(4)
    metrics = [
        ("Autoridad", st.session_state.get("authority_label", "Sin definir")),
        ("Árboles", str(trees)),
        ("Requisitos", str(req_count)),
        ("Documentos subidos", str(uploaded_count)),
    ]
    for col, (title, value) in zip(cols, metrics):
        with col:
            st.markdown(f"<div class='kpi'><div class='title'>{title}</div><div class='value'>{value}</div></div>", unsafe_allow_html=True)

    st.write("")
    left, right = st.columns([1.2, 1])
    with left:
        st.markdown("### Flujo de trabajo")
        st.markdown(
            "1. **Proyecto** — seleccione autoridad y modalidad.  \n"
            "2. **Solicitante** — registre identificación y contacto.  \n"
            "3. **Predio** — registre ubicación y georreferenciación.  \n"
            "4. **Árboles** — registre cada individuo y su diagnóstico.  \n"
            "5. **Validación** — revise requisitos, cargue documentos y genere el expediente."
        )
    with right:
        st.markdown("### Estado actual")
        st.info(f"**Grupo documental:** {req_group}")
        if sol.get("nombre"):
            st.success(f"Solicitante: {sol['nombre']}")
        else:
            st.warning("Todavía no se ha diligenciado el solicitante.")
        st.write(f"**Volumen estimado:** {total_volume():.3f} m³")
        st.write(f"**Prioridad:** {priority_text()}")

    st.markdown("### Qué cambió frente a la versión anterior")
    st.markdown(
        "- Navegación lateral por módulos, como el prototipo de referencia.\n"
        "- Una solicitud puede registrar **varios árboles**, cada uno con sus datos.\n"
        "- Los requisitos cambian automáticamente según **SDA, CAR Cundinamarca o Corpoboyacá** y, en Corpoboyacá, según la modalidad.\n"
        "- Cada requisito tiene espacio para **descargar la plantilla** y luego **subir el documento diligenciado**.\n"
        "- La validación genera un **expediente ZIP** con el resumen técnico y los archivos cargados."
    )


# =============================================================================
# PÁGINA: NUEVA SOLICITUD
# =============================================================================
def page_nueva_solicitud():
    step = st.session_state["request_page"]
    st.subheader(f"Nueva solicitud · {step}")

    if step == "Proyecto":
        with st.form("form_proyecto"):
            st.markdown("### 1. Proyecto y autoridad")
            project_type = st.selectbox(
                "Tipo de trámite",
                ["Aprovechamiento forestal de árboles aislados"],
                index=0,
            )
            authority_label = st.selectbox(
                "Autoridad ambiental competente",
                list(AUTHORITY_OPTIONS.keys()),
                index=list(AUTHORITY_OPTIONS.keys()).index(st.session_state.get("authority_label", "SDA (Bogotá D.C.)")),
            )
            authority_key = AUTHORITY_OPTIONS[authority_label]

            corpo_case_key = st.session_state.get("corpo_case", "CORPO_PRIORITARIA")
            if authority_key == "CORPO":
                selected_label = st.selectbox(
                    "Modalidad Corpoboyacá",
                    list(CORPO_OPTIONS.keys()),
                    index=max(0, list(CORPO_OPTIONS.values()).index(corpo_case_key)),
                )
                corpo_case_key = CORPO_OPTIONS[selected_label]
                st.caption(REQUIREMENTS[corpo_case_key]["procedure"])

            descripcion = st.text_area(
                "Descripción / objeto de la solicitud",
                value=st.session_state.get("solicitud", {}).get("descripcion", ""),
                placeholder="Describa brevemente la intervención y el motivo de la solicitud.",
            )
            if st.form_submit_button("Guardar proyecto y continuar"):
                st.session_state["project_type"] = project_type
                st.session_state["authority_label"] = authority_label
                st.session_state["authority_key"] = authority_key
                st.session_state["corpo_case"] = corpo_case_key
                st.session_state.setdefault("solicitud", {})["descripcion"] = descripcion
                st.success("Proyecto guardado.")
                st.session_state["request_page"] = "Solicitante"
                st.rerun()

    elif step == "Solicitante":
        sol = st.session_state.setdefault("solicitud", {})
        with st.form("form_solicitante"):
            st.markdown("### 2. Datos del solicitante")
            c1, c2 = st.columns(2)
            with c1:
                sol["nombre"] = st.text_input("Nombre o razón social", sol.get("nombre", ""))
                sol["tipo_doc"] = st.selectbox("Tipo de documento", ["CC", "NIT", "CE", "PASAPORTE"], index=["CC", "NIT", "CE", "PASAPORTE"].index(sol.get("tipo_doc", "CC")))
                sol["num_doc"] = st.text_input("Número de documento", sol.get("num_doc", ""))
                sol["calidad"] = st.selectbox("Calidad sobre el predio", ["Propietario", "Tenedor", "Poseedor", "Apoderado / Autorizado"], index=["Propietario", "Tenedor", "Poseedor", "Apoderado / Autorizado"].index(sol.get("calidad", "Propietario")))
            with c2:
                sol["correo"] = st.text_input("Correo electrónico", sol.get("correo", ""))
                sol["telefono"] = st.text_input("Teléfono", sol.get("telefono", ""))
                sol["representante_legal"] = st.text_input("Representante legal (si aplica)", sol.get("representante_legal", ""))
                sol["observaciones_solicitante"] = st.text_area("Observaciones", sol.get("observaciones_solicitante", ""))

            if st.form_submit_button("Guardar solicitante y continuar"):
                st.success("Datos del solicitante guardados.")
                st.session_state["request_page"] = "Predio"
                st.rerun()

    elif step == "Predio":
        sol = st.session_state.setdefault("solicitud", {})
        with st.form("form_predio"):
            st.markdown("### 3. Predio y georreferenciación")
            c1, c2 = st.columns(2)
            with c1:
                sol["nombre_predio"] = st.text_input("Nombre del predio / institución", sol.get("nombre_predio", ""))
                sol["municipio"] = st.text_input("Municipio / ciudad", sol.get("municipio", ""))
                sol["direccion"] = st.text_input("Dirección", sol.get("direccion", ""))
            with c2:
                sol["latitud"] = st.number_input("Latitud", value=float(sol.get("latitud", 5.0)), format="%.6f")
                sol["longitud"] = st.number_input("Longitud", value=float(sol.get("longitud", -73.0)), format="%.6f")
                sol["matricula"] = st.text_input("Matrícula / referencia predial (si aplica)", sol.get("matricula", ""))

            st.markdown("#### Ubicación")
            st.map(pd.DataFrame({"lat": [sol["latitud"]], "lon": [sol["longitud"]]}), zoom=12)

            if st.form_submit_button("Guardar predio y continuar"):
                st.success("Predio guardado.")
                st.session_state["request_page"] = "Árboles"
                st.rerun()

    elif step == "Árboles":
        st.markdown("### 4. Caracterización de los árboles")
        st.caption("La aplicación permite registrar múltiples individuos. Cada árbol queda guardado por separado con sus medidas y diagnóstico.")

        current = len(st.session_state.get("trees", [])) or 1
        n = st.number_input("Número de individuos arbóreos", min_value=1, max_value=100, value=current, step=1)
        ensure_tree_count(int(n))

        for i, tree in enumerate(st.session_state["trees"], start=1):
            with st.expander(f"🌳 Individuo {i}", expanded=(i == 1)):
                a, b = st.columns(2)
                with a:
                    tree["nombre_comun"] = st.text_input(f"Nombre común · Árbol {i}", tree.get("nombre_comun", ""), key=f"tc_{i}")
                    tree["dap_cm"] = st.number_input(f"DAP (cm) · Árbol {i}", min_value=0.0, value=float(tree.get("dap_cm", 0.0)), step=0.1, key=f"tdap_{i}")
                    tree["altura_m"] = st.number_input(f"Altura total (m) · Árbol {i}", min_value=0.0, value=float(tree.get("altura_m", 0.0)), step=0.1, key=f"th_{i}")
                    tree["factor_nombre"] = st.selectbox(f"Factor de forma · Árbol {i}", ["Paraboloide (0.55)", "Cilíndrico (0.75)", "Cónico (0.33)", "Neiloide (0.25)"], index=["Paraboloide (0.55)", "Cilíndrico (0.75)", "Cónico (0.33)", "Neiloide (0.25)"].index(tree.get("factor_nombre", "Paraboloide (0.55)")), key=f"tf_{i}")
                with b:
                    tree["nombre_cientifico"] = st.text_input(f"Nombre científico · Árbol {i}", tree.get("nombre_cientifico", ""), key=f"tsci_{i}")
                    tree["riesgo"] = st.checkbox(f"Riesgo inminente de caída · Árbol {i}", value=bool(tree.get("riesgo", False)), key=f"trisk_{i}")
                    tree["fauna"] = st.checkbox(f"Presencia de avifauna / nidos · Árbol {i}", value=bool(tree.get("fauna", False)), key=f"tfauna_{i}")
                    tree["epifitas"] = st.checkbox(f"Presencia de epífitas · Árbol {i}", value=bool(tree.get("epifitas", False)), key=f"tepif_{i}")
                    tree["observaciones"] = st.text_area(f"Observaciones · Árbol {i}", tree.get("observaciones", ""), key=f"tobs_{i}")
                    photo = st.file_uploader(f"Evidencia fotográfica · Árbol {i}", type=["jpg", "jpeg", "png", "webp"], key=f"photo_{i}")
                    if photo is not None:
                        ext = Path(photo.name).suffix.lower()
                        if ext in {".jpg", ".jpeg", ".png", ".webp"}:
                            signature = uploaded_signature(photo)
                            if tree.get("foto_signature") != signature or not tree.get("foto_path") or not Path(tree.get("foto_path", "")).exists():
                                saved = save_uploaded_file(photo, case_folder() / "evidencia_arboles", prefix=f"arbol_{i}__")
                                tree["foto_path"] = str(saved)
                                tree["foto_name"] = photo.name
                                tree["foto_signature"] = signature
                            st.image(photo, width=260, caption=photo.name)

                st.info(f"Volumen estimado del individuo: **{volume_for_tree(tree):.3f} m³**")

        if st.button("Guardar árboles y continuar", type="primary"):
            st.session_state["request_page"] = "Validación"
            st.success("Caracterización arbórea guardada.")
            st.rerun()

    elif step == "Validación":
        page_validacion_solicitud()


# =============================================================================
# VALIDACIÓN DE LA SOLICITUD
# =============================================================================
def page_validacion_solicitud():
    st.markdown("### 5. Validación y expediente")

    sol = st.session_state.get("solicitud", {})
    missing_core = []
    required_core = {
        "Solicitante": sol.get("nombre"),
        "Documento": sol.get("num_doc"),
        "Predio": sol.get("nombre_predio"),
        "Municipio": sol.get("municipio"),
    }
    for label, val in required_core.items():
        if not val:
            missing_core.append(label)

    trees_ok = len(st.session_state.get("trees", [])) > 0 and all(
        t.get("nombre_comun") and t.get("dap_cm", 0) > 0 and t.get("altura_m", 0) > 0
        for t in st.session_state.get("trees", [])
    )

    reqs = get_requirements()["documents"]
    missing_docs = [r for r in reqs if not requirement_uploaded(r["id"])]

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f"<div class='kpi'><div class='title'>Información básica</div><div class='value'>{'Completa' if not missing_core else 'Pendiente'}</div></div>", unsafe_allow_html=True)
    with c2:
        st.markdown(f"<div class='kpi'><div class='title'>Árboles</div><div class='value'>{'Completa' if trees_ok else 'Pendiente'}</div></div>", unsafe_allow_html=True)
    with c3:
        st.markdown(f"<div class='kpi'><div class='title'>Documentos</div><div class='value'>{len(reqs) - len(missing_docs)}/{len(reqs)}</div></div>", unsafe_allow_html=True)

    if missing_core:
        st.warning("Faltan datos: " + ", ".join(missing_core))
    if not trees_ok:
        st.warning("Cada árbol debe tener al menos nombre común, DAP y altura.")

    st.markdown("### Requisitos de esta solicitud")
    if not reqs:
        st.info("No hay una lista oficial independiente para 'General' en la información suministrada. Define primero la autoridad competente para activar el checklist correspondiente.")

    for req in reqs:
        uploads = requirement_uploaded(req["id"])
        with st.container(border=True):
            left, mid, right = st.columns([2.1, 1.1, 1.4])
            with left:
                st.markdown(f"**{req['name']}**")
                if req.get("note"):
                    st.caption(req["note"])
            with mid:
                if uploads:
                    st.markdown(f"<span class='doc-ok'>✓ {len(uploads)} archivo(s)</span>", unsafe_allow_html=True)
                else:
                    st.markdown("<span class='doc-miss'>Pendiente</span>", unsafe_allow_html=True)
            with right:
                path = local_template_path(req)
                if path:
                    st.download_button(
                        "⬇️ Descargar plantilla",
                        data=path.read_bytes(),
                        file_name=path.name,
                        key=f"dl_{req['id']}",
                    )
                elif req.get("source_url"):
                    st.link_button("🌐 Fuente oficial", req["source_url"])
                else:
                    st.caption("Sube la plantilla en /documentos")

            uploaded_files = st.file_uploader(
                "Subir documento(s) diligenciado(s)",
                type=[ext.replace(".", "") for ext in sorted(ALLOWED_DOC_EXTENSIONS)],
                accept_multiple_files=True,
                key=f"upl_{req['id']}",
            )
            if uploaded_files:
                store_requirement_uploads(req["id"], uploaded_files)
                st.success("Archivo(s) guardado(s) en el expediente.")

    st.markdown("### Documentos adicionales")
    extra = st.file_uploader(
        "Subir otros soportes que no estén en la lista anterior",
        type=[ext.replace(".", "") for ext in sorted(ALLOWED_DOC_EXTENSIONS)],
        accept_multiple_files=True,
        key="extras_uploader",
    )
    if extra:
        extra_dir = case_folder() / "documentos_adicionales"
        for file in extra:
            saved = save_uploaded_file(file, extra_dir, prefix="extra__")
            st.success(f"Guardado: {saved.name}")

    st.markdown("### Resumen")
    a, b, c = st.columns(3)
    with a:
        st.metric("Volumen total estimado", f"{total_volume():.3f} m³")
    with b:
        st.metric("Prioridad", priority_text())
    with c:
        st.metric("Archivos cargados", sum(len(v) for k, v in st.session_state.get("uploaded_docs", {}).items() if k != "__signatures__"))

    technical_doc = create_technical_doc()
    st.download_button(
        "📄 Descargar resumen técnico (.docx)",
        data=technical_doc,
        file_name=f"EcoRegion_Resumen_Tecnico_{current_case_id()}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
    )
    st.download_button(
        "🗂️ Descargar expediente completo (.zip)",
        data=build_case_zip(),
        file_name=f"EcoRegion_Expediente_{current_case_id()}.zip",
        mime="application/zip",
    )


# =============================================================================
# PÁGINA: REQUISITOS
# =============================================================================
def page_requisitos():
    st.subheader("Requisitos documentales")
    st.caption("El checklist se muestra según la autoridad y, para Corpoboyacá, según la modalidad escogida.")

    st.markdown(f"### {get_requirements()['label']}")
    st.write(get_requirements()["procedure"])

    if st.session_state.get("authority_key") == "GENERAL":
        st.warning("La información entregada no contiene un checklist oficial independiente para la opción General. La app no inventa requisitos.")
        return

    for i, req in enumerate(get_requirements()["documents"], start=1):
        st.markdown(f"**{i}. {req['name']}**")
        if req.get("note"):
            st.caption(req["note"])
        path = local_template_path(req)
        if path:
            st.download_button("Descargar plantilla", path.read_bytes(), file_name=path.name, key=f"req_main_dl_{req['id']}")
        elif req.get("source_url"):
            st.link_button("Abrir fuente oficial", req["source_url"])
        else:
            st.caption("No hay plantilla local cargada. Puedes ponerla en la carpeta indicada en el Manual de uso.")


# =============================================================================
# PÁGINA: AUTORIDADES
# =============================================================================
def page_autoridades():
    st.subheader("Autoridades y reglas documentales")
    st.markdown(
        "La aplicación evita mezclar requisitos: primero se define la autoridad y luego se activa el conjunto documental correspondiente."
    )

    cards = [
        ("SDA (Bogotá D.C.)", REQUIREMENTS["SDA"]),
        ("CAR Cundinamarca", REQUIREMENTS["CAR"]),
        ("Corpoboyacá", None),
    ]
    for title, info in cards:
        with st.expander(title, expanded=(title.startswith(st.session_state.get("authority_label", "")))):
            if info:
                st.write(info["procedure"])
                for req in info["documents"]:
                    st.write("•", req["name"])
            else:
                st.write("Corpoboyacá tiene modalidades distintas en la información suministrada:")
                for label, key in CORPO_OPTIONS.items():
                    st.markdown(f"**{label}**")
                    for req in REQUIREMENTS[key]["documents"]:
                        st.write("•", req["name"])
                    st.write("")


# =============================================================================
# PÁGINA: DOCUMENTACIÓN
# =============================================================================
def page_documentacion():
    st.subheader("Centro documental")
    st.write(
        "Aquí puedes comprobar si las plantillas oficiales están cargadas y revisar los documentos diligenciados que ya pertenecen al expediente."
    )

    st.markdown("### Plantillas oficiales disponibles en el servidor")
    reqs = get_requirements()["documents"]
    if not reqs:
        st.info("Selecciona una autoridad específica para activar el centro documental.")
        return
    for req in reqs:
        path = local_template_path(req)
        status = "Disponible" if path else "No cargada"
        st.write(f"**{status}** · {req['name']}")
        if path:
            st.caption(f"{path} · {format_file_size(path)}")

    st.markdown("### Archivos diligenciados en el expediente")
    uploaded_all = [Path(p) for values in st.session_state.get("uploaded_docs", {}).values() for p in values if Path(p).exists()]
    if uploaded_all:
        for path in uploaded_all:
            st.write(f"📄 {path.name} — {format_file_size(path)}")
            st.download_button("Descargar", path.read_bytes(), file_name=path.name, key=f"existing_{path}")
    else:
        st.info("Aún no hay documentos diligenciados cargados.")

    st.markdown("### Carpeta de evidencias")
    evidence_dir = case_folder() / "evidencia_arboles"
    evidence = list(evidence_dir.glob("*") if evidence_dir.exists() else [])
    if evidence:
        for path in evidence:
            st.write(f"📷 {path.name}")
    else:
        st.info("No se han guardado fotografías de árboles en este expediente.")


# =============================================================================
# PÁGINA: SEGUIMIENTO
# =============================================================================
def page_seguimiento():
    st.subheader("Seguimiento del trámite")
    steps = [
        ("Proyecto", bool(st.session_state.get("authority_label"))),
        ("Solicitante", bool(st.session_state.get("solicitud", {}).get("nombre") and st.session_state.get("solicitud", {}).get("num_doc"))),
        ("Predio", bool(st.session_state.get("solicitud", {}).get("nombre_predio") and st.session_state.get("solicitud", {}).get("municipio"))),
        ("Árboles", bool(st.session_state.get("trees"))),
        ("Requisitos", len(st.session_state.get("uploaded_docs", {})) > 0 or len(get_requirements()["documents"]) == 0),
        ("Validación", st.session_state.get("request_page") == "Validación"),
    ]
    for label, done in steps:
        icon = "✅" if done else "⬜"
        st.markdown(f"### {icon} {label}")

    st.markdown("### Observación")
    if all(x[1] for x in steps):
        st.success("El expediente tiene completas las etapas básicas del prototipo. Revisa los documentos antes de radicar.")
    else:
        st.warning("Hay etapas pendientes. Usa 'Nueva solicitud' para completarlas.")


# =============================================================================
# PÁGINA: ENCUESTA
# =============================================================================
def page_encuesta():
    st.subheader("Encuesta de validación y usabilidad")
    st.caption("Las respuestas se almacenan localmente en data/respuestas_encuesta.csv.")

    col1, col2 = st.columns(2)
    with col1:
        with st.form("survey_form"):
            q1 = st.slider("1. Facilidad de diligenciamiento", 1, 5, 5)
            q2 = st.slider("2. Claridad de requisitos según autoridad", 1, 5, 5)
            q3 = st.slider("3. Facilidad para encontrar y subir documentos", 1, 5, 5)
            q4 = st.slider("4. Comprensión de resultados y prioridad", 1, 5, 5)
            q5 = st.slider("5. Utilidad del resumen/documento generado", 1, 5, 5)
            comments = st.text_area("Comentarios o sugerencias")
            submitted = st.form_submit_button("Guardar respuesta")
            if submitted:
                row = pd.DataFrame([{
                    "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Facilidad": q1,
                    "Claridad_requisitos": q2,
                    "Documentacion": q3,
                    "Resultados": q4,
                    "Documento": q5,
                    "Comentarios": comments,
                }])
                if SURVEY_FILE.exists():
                    row.to_csv(SURVEY_FILE, mode="a", sep=";", index=False, header=False, encoding="utf-8-sig")
                else:
                    row.to_csv(SURVEY_FILE, sep=";", index=False, encoding="utf-8-sig")
                st.success("Respuesta registrada.")

    with col2:
        st.markdown("### Resultados")
        if SURVEY_FILE.exists():
            try:
                df = pd.read_csv(SURVEY_FILE, sep=";", encoding="utf-8-sig")
                if len(df):
                    st.metric("Usuarios encuestados", len(df))
                    cats = ["Facilidad", "Claridad_requisitos", "Documentacion", "Resultados", "Documento"]
                    vals = [pd.to_numeric(df[c], errors="coerce").mean() for c in cats]
                    fig, ax = plt.subplots(figsize=(7, 4))
                    ax.barh(cats, vals)
                    ax.set_xlim(0, 5)
                    ax.set_xlabel("Promedio (1 a 5)")
                    ax.set_title("Validación de usabilidad")
                    for i, value in enumerate(vals):
                        if pd.notna(value):
                            ax.text(float(value) + 0.05, i, f"{value:.2f}", va="center")
                    st.pyplot(fig)
                    st.download_button(
                        "📥 Descargar CSV",
                        SURVEY_FILE.read_bytes(),
                        file_name="Reporte_Respuestas_EcoRegion.csv",
                        mime="text/csv",
                    )
                else:
                    st.info("No hay respuestas todavía.")
            except Exception as exc:
                st.error(f"No fue posible leer la encuesta: {exc}")
        else:
            st.info("No hay respuestas todavía.")


# =============================================================================
# PÁGINA: MANUAL
# =============================================================================
def page_manual():
    st.subheader("Manual de uso y carga de documentos")
    st.markdown("### 1. Estructura de carpetas que debes crear")
    st.code(
        """EcoRegion_app_v2.py
LOGO.png                         # opcional
\ndocumentos/
├── SDA/
│   ├── PM04-PR30-F1 Formulario solicitud manejo aprovechamiento forestal.xlsx
│   ├── PM04-PR30-F2 Recoleccion informacion silvicultural individuo ficha1.xls
│   └── PM04-PR30-F3 Ficha tecnica de registro Ficha 2.docx
│
├── CAR_Cundinamarca/
│   └── CAR_FUN.pdf
│
├── Corpoboyaca/
│   ├── prioritaria_emergencia_obra/
│   │   ├── FUN_Corpoboyaca.pdf
│   │   ├── FGR-06_Parte_A.pdf
│   │   └── FGR-29.pdf
│   ├── nativas_mayor_50m3/
│   │   ├── FUN_Corpoboyaca.pdf
│   │   ├── Estudio_Tecnico_Arboles_Aislados.pdf
│   │   └── FGR-29.pdf
│   ├── nativas_menor_50m3_exoticas/
│   │   ├── FUN_Corpoboyaca.pdf
│   │   ├── FGR-06_Parte_A_y_B.pdf
│   │   └── FGR-29.pdf
│   └── uso_domestico/
│       ├── FUN_Corpoboyaca.pdf
│       └── FGR-06_Parte_A.pdf
│
└── data/
    └── (la app crea aquí expedientes y encuesta)
        """,
        language="text",
    )

    st.markdown("### 2. ¿Qué tienes que subir tú?")
    st.info(
        "Los formatos oficiales en blanco. La app no los inventa ni los descarga de forma automática: tú los colocas una vez en las carpetas anteriores y, desde la aplicación, el usuario podrá descargarlos. Después el usuario diligencia el formato y lo vuelve a subir en la sección 'Validación'."
    )

    st.markdown("### 3. Cómo funciona el flujo documental")
    st.markdown(
        "**Seleccionar autoridad → activar checklist → descargar formato → diligenciar → volver a subir → validar → generar expediente ZIP.**"
    )

    st.markdown("### 4. Dónde se guardan los archivos que sube el usuario")
    st.code("data/expedientes/<ID_DEL_EXPEDIENTE>/documentos_diligenciados/")
    st.markdown("Las fotografías de árboles se guardan en:")
    st.code("data/expedientes/<ID_DEL_EXPEDIENTE>/evidencia_arboles/")

    st.markdown("### 5. Importante para producción")
    st.warning(
        "Esta versión guarda archivos en disco local, adecuado para una demo/MVP. Para producción multiusuario conviene mover documentos, fotos, base de datos y autenticación a servicios persistentes (por ejemplo, almacenamiento de objetos + base de datos administrada) y añadir control de acceso por usuario/rol."
    )


# =============================================================================
# ROUTER
# =============================================================================
if st.session_state["page"] == "Inicio":
    page_inicio()
elif st.session_state["page"] == "Nueva solicitud":
    page_nueva_solicitud()
elif st.session_state["page"] == "Requisitos":
    page_requisitos()
elif st.session_state["page"] == "Autoridades":
    page_autoridades()
elif st.session_state["page"] == "Documentación":
    page_documentacion()
elif st.session_state["page"] == "Seguimiento":
    page_seguimiento()
elif st.session_state["page"] == "Encuesta de validación":
    page_encuesta()
elif st.session_state["page"] == "Manual de uso":
    page_manual()
