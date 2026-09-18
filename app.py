import io
import json
import math
import hashlib
import hmac
import os
import re
import secrets
import shutil
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from docx import Document

# =============================================================================
# ECORREGION APP v3
# Gestión de expedientes para aprovechamiento forestal de árboles aislados
# - Autenticación por usuario/contraseña y roles CLIENTE / ADMIN
# - Expedientes persistentes por usuario
# - Nueva solicitud con subnavegación horizontal
# - Requisitos dinámicos por autoridad/modalidad
# - Centro documental con estado PENDIENTE / SUBIDO
# - Gestión de plantillas oficiales desde ADMIN
# - Seguimiento detallado
# - Documento técnico + matriz documental + ZIP del expediente
# =============================================================================

APP_TITLE = "EcoRegión App"
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DOCS_DIR = BASE_DIR / "documentos"
CASE_DIR = DATA_DIR / "expedientes"
CONFIG_DIR = BASE_DIR / "config"
USERS_FILE = CONFIG_DIR / "usuarios.json"
SURVEY_FILE = DATA_DIR / "respuestas_encuesta.csv"

for folder in [DATA_DIR, DOCS_DIR, CASE_DIR, CONFIG_DIR]:
    folder.mkdir(parents=True, exist_ok=True)

st.set_page_config(
    page_title="EcoRegión App",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================================================
# ESTILOS MODERNOS
# =============================================================================
st.markdown(
    """
    <style>
    :root {
        --bg: #f4f7f5;
        --surface: #ffffff;
        --surface-2: #f8fbf9;
        --text: #16251e;
        --muted: #64736b;
        --green-900: #103b2a;
        --green-800: #14553c;
        --green-700: #1c7651;
        --green-600: #299365;
        --green-100: #e3f5eb;
        --green-050: #f0faf4;
        --border: #dce7e1;
        --warning: #9a6200;
        --warning-bg: #fff5df;
        --danger: #b3261e;
        --danger-bg: #fff0ee;
        --success: #18794e;
        --shadow: 0 8px 28px rgba(20, 55, 40, .08);
    }

    .stApp { background: var(--bg); color: var(--text); }
    [data-testid="stHeader"] { background: transparent; }
    [data-testid="stSidebar"] {
        background: #102f24 !important;
        border-right: 1px solid #0b241b !important;
    }
    [data-testid="stSidebar"] * { color: #effaf4 !important; }
    [data-testid="stSidebar"] .stRadio label { border-radius: 10px; }
    [data-testid="stSidebar"] .stRadio label:hover { background: rgba(255,255,255,.08); }
    [data-testid="stSidebar"] hr { border-color: rgba(255,255,255,.13); }

    h1, h2, h3, h4 { color: var(--green-900) !important; letter-spacing: -.02em; }
    p, label, .stMarkdown, [data-testid="stMetricLabel"] { color: var(--text); }
    .stCaption { color: var(--muted) !important; }

    input, textarea, [data-baseweb="select"] > div,
    [data-testid="stNumberInput"] input {
        background: #fff !important;
        color: var(--text) !important;
        border-color: var(--border) !important;
    }
    [data-baseweb="popover"], [role="listbox"], [role="option"] {
        background: #fff !important;
        color: var(--text) !important;
    }
    [role="option"] * { color: var(--text) !important; }

    .stButton > button, .stFormSubmitButton > button {
        border-radius: 10px !important;
        font-weight: 700 !important;
        border: 1px solid var(--green-700) !important;
        background: var(--green-700) !important;
        color: #fff !important;
        min-height: 42px;
    }
    .stButton > button:hover, .stFormSubmitButton > button:hover {
        background: var(--green-800) !important;
        border-color: var(--green-800) !important;
    }
    .stDownloadButton > button {
        border-radius: 10px !important;
        border: 1px solid var(--border) !important;
        background: #fff !important;
        color: var(--green-800) !important;
        font-weight: 700 !important;
    }
    .stDownloadButton > button:hover { background: var(--green-050) !important; }

    .hero {
        background: linear-gradient(135deg, #103b2a 0%, #1f7b55 72%, #2b9a6a 100%);
        color: white;
        border-radius: 18px;
        padding: 1.5rem 1.7rem;
        margin-bottom: 1.2rem;
        box-shadow: 0 12px 36px rgba(16,59,42,.18);
    }
    .hero h1, .hero p { color: white !important; margin: 0; }
    .hero h1 { font-size: 2.15rem; }
    .hero p { opacity: .9; margin-top: .3rem; }

    .card {
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: 16px;
        padding: 1rem 1.1rem;
        margin-bottom: .9rem;
        box-shadow: var(--shadow);
    }
    .kpi {
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: 15px;
        padding: 1rem;
        min-height: 98px;
        box-shadow: var(--shadow);
    }
    .kpi .title { font-size: .8rem; color: var(--muted); margin-bottom: .3rem; }
    .kpi .value { font-size: 1.3rem; font-weight: 850; color: var(--green-900); }

    .status {
        display: inline-flex; align-items: center; gap: .35rem;
        border-radius: 999px; padding: .32rem .7rem; font-weight: 800; font-size: .82rem;
    }
    .status-ok { color: var(--success); background: #e8f7ef; }
    .status-pending { color: var(--warning); background: var(--warning-bg); }
    .status-danger { color: var(--danger); background: var(--danger-bg); }
    .status-info { color: #245d85; background: #eaf4fb; }

    .stepbar {
        background: #fff; border: 1px solid var(--border); border-radius: 15px;
        padding: .35rem; margin-bottom: 1.1rem; box-shadow: var(--shadow);
    }
    .stepbar [data-testid="stRadio"] > div { gap: .25rem; }
    .stepbar label { border-radius: 10px; }

    .sidebar-brand { text-align: center; padding: .3rem 0 1rem; }
    .sidebar-brand .mark {
        width: 58px; height: 58px; margin: 0 auto .55rem;
        background: #dff4e8; color: var(--green-900);
        border-radius: 16px; display: flex; align-items: center; justify-content: center;
        font-size: 1.9rem; font-weight: 800;
    }
    .sidebar-brand .name { font-weight: 850; font-size: 1.1rem; }
    .sidebar-brand .tag { font-size: .76rem; opacity: .78; }
    .role-chip { text-align:center; font-weight:800; font-size:.78rem; margin-bottom:.4rem; }

    .doc-title { font-weight: 800; color: var(--green-900); }
    .small { font-size: .88rem; color: var(--muted); }
    .muted-box { background: var(--surface-2); border:1px solid var(--border); border-radius:12px; padding:.8rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

# =============================================================================
# AUTENTICACIÓN
# =============================================================================
def hash_password(password: str, salt: bytes | None = None) -> str:
    salt = salt or secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 180_000)
    return f"pbkdf2_sha256$180000${salt.hex()}${digest.hex()}"


def verify_password(password: str, encoded: str) -> bool:
    try:
        algorithm, iterations, salt_hex, digest_hex = encoded.split("$")
        if algorithm != "pbkdf2_sha256":
            return False
        test = hashlib.pbkdf2_hmac(
            "sha256", password.encode("utf-8"), bytes.fromhex(salt_hex), int(iterations)
        ).hex()
        return hmac.compare_digest(test, digest_hex)
    except Exception:
        return False


def init_users():
    if USERS_FILE.exists():
        return
    admin_password = os.environ.get("ECOREGION_ADMIN_PASSWORD", "AdminEco2026!")
    client_password = os.environ.get("ECOREGION_CLIENT_PASSWORD", "ClienteEco2026!")
    users = {
        "admin": {
            "name": "Administrador EcoRegión",
            "role": "admin",
            "password_hash": hash_password(admin_password),
            "active": True,
        },
        "cliente": {
            "name": "Cliente Demo",
            "role": "client",
            "password_hash": hash_password(client_password),
            "active": True,
        },
    }
    USERS_FILE.write_text(json.dumps(users, ensure_ascii=False, indent=2), encoding="utf-8")


init_users()


def load_users() -> dict:
    try:
        return json.loads(USERS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_users(users: dict):
    USERS_FILE.write_text(json.dumps(users, ensure_ascii=False, indent=2), encoding="utf-8")


def authenticate(username: str, password: str):
    users = load_users()
    user = users.get(username.strip())
    if not user or not user.get("active", True):
        return None
    if verify_password(password, user.get("password_hash", "")):
        return {"username": username.strip(), **user}
    return None


def logout():
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()


def login_page():
    st.markdown("<div style='height:8vh'></div>", unsafe_allow_html=True)
    left, center, right = st.columns([1, 1.25, 1])
    with center:
        st.markdown(
            """
            <div class='card' style='text-align:center; padding:2rem;'>
                <div style='font-size:3rem;'>🌿</div>
                <h1 style='margin-bottom:.2rem;'>EcoRegión</h1>
                <p style='color:#64736b;'>Gestión inteligente de permisos forestales</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        with st.form("login_form"):
            username = st.text_input("Usuario", placeholder="Ingrese su usuario")
            password = st.text_input("Contraseña", type="password", placeholder="Ingrese su contraseña")
            if st.form_submit_button("Iniciar sesión", use_container_width=True, type="primary"):
                user = authenticate(username, password)
                if user:
                    st.session_state["authenticated"] = True
                    st.session_state["user"] = user
                    load_or_create_user_case(user["username"])
                    st.rerun()
                else:
                    st.error("Usuario o contraseña incorrectos.")
        st.caption("El acceso del cliente permite diligenciar y cargar su expediente. Las funciones administrativas están restringidas al rol ADMIN.")


# =============================================================================
# DATOS DOCUMENTALES
# =============================================================================
def slug(text: str) -> str:
    value = re.sub(r"[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ]+", "_", str(text).strip())
    return value.strip("_").lower()


SOURCE_SDA = "https://www.ambientebogota.gov.co/web/transparencia/lineamientos/-/document_library_display/eVU7938nZRvM/view/1444163"
SOURCE_CAR = "https://car1.car.gov.co/vercontenido/1178"
SOURCE_CAR_FUN = "https://www.car.gov.co/uploads/files/666b2c9e17175.pdf"
SOURCE_CORPO = "https://www.corpoboyaca.gov.co/ventanilla-atencion/aprovechamiento-forestal-arboles-aislados/"


def req(rid, name, filename="", folder="", note="", source_url="", required=True, conditional=""):
    return {
        "id": rid, "name": name, "local_filename": filename, "folder": folder,
        "note": note, "source_url": source_url, "required": required, "conditional": conditional,
    }


REQUIREMENTS = {
    "SDA": {
        "label": "SDA (Bogotá D.C.)",
        "procedure_title": "Permiso o autorización para aprovechamiento forestal de árboles aislados",
        "procedure": "La SDA publica tres formatos para este trámite: F1, F2 y F3.",
        "documents": [
            req("sda_f1", "PM04-PR30-F1 Formulario solicitud manejo aprovechamiento forestal.xlsx", "PM04-PR30-F1 Formulario solicitud manejo aprovechamiento forestal.xlsx", "SDA", "Formulario oficial de la SDA.", SOURCE_SDA),
            req("sda_f2", "PM04-PR30-F2 Recoleccion informacion silvicultural individuo ficha1.xls", "PM04-PR30-F2 Recoleccion informacion silvicultural individuo ficha1.xls", "SDA", "Ficha 1 de recolección de información silvicultural.", SOURCE_SDA),
            req("sda_f3", "PM04-PR30-F3 Ficha tecnica de registro Ficha 2.docx", "PM04-PR30-F3 Ficha tecnica de registro Ficha 2.docx", "SDA", "Ficha técnica de registro.", SOURCE_SDA),
        ],
    },
    "CAR": {
        "label": "CAR Cundinamarca",
        "procedure_title": "Permiso o autorización para aprovechamiento forestal de árboles aislados",
        "procedure": "Checklist basado en la ficha oficial de la CAR para árboles aislados. La aplicación marca requisitos según la situación del solicitante.",
        "documents": [
            req("car_fun", "Formato Único Nacional de Solicitud de Aprovechamiento Forestal y Manejo Sostenible de Flora Silvestre y los Productos Forestales No Maderables Nuevo/Prórroga, debidamente diligenciado.", "CAR_FUN.pdf", "CAR_Cundinamarca", "Formato FUN oficial.", SOURCE_CAR_FUN),
            req("car_libertad", "Certificado de libertad y tradición expedido dentro de los dos (2) meses inmediatamente anteriores a la presentación de la solicitud. Si se trata de predio ajeno se anexará la prueba de la posesión o tenencia.", required=True, conditional="Predio de propiedad privada / predio ajeno según corresponda."),
            req("car_autorizacion", "Autorización escrita del propietario cuando el solicitante no sea el mismo propietario del predio.", required=True, conditional="Aplica cuando solicitante y propietario no son la misma persona."),
            req("car_croquis", "Croquis a mano alzada para acceso al predio.", required=True),
            req("car_poder", "Poder debidamente otorgado, cuando se actúa mediante abogado.", required=False, conditional="Aplica cuando se actúa mediante abogado."),
            req("car_existencia", "Certificado de existencia y representación legal vigente, para el caso de personas jurídicas.", required=False, conditional="Aplica cuando el solicitante es persona jurídica."),
        ],
    },
    "CORPO_PRIORITARIA": {
        "label": "Corpoboyacá — prioritaria / emergencia / obra pública o privada",
        "procedure_title": "Árboles por solicitud prioritaria, tala de emergencia, construcción de obra pública o privada",
        "procedure": "La página oficial de Corpoboyacá exige FUN, FGR-06 Parte A, soportes de propiedad/identificación, FGR-29 y soporte de pago cuando corresponda.",
        "documents": [
            req("corpo_fun", "Formato único nacional de solicitud de aprovechamiento forestal y manejo sostenible de flora silvestre y productos forestales no maderables.", "FUN_Corpoboyaca.pdf", "Corpoboyaca/prioritaria_emergencia_obra", "Formato oficial.", SOURCE_CORPO),
            req("corpo_fgr06a", "Formato FGR-06 Parte A «Inventario Forestal al 100%» (incluir registro fotográfico y georreferenciación por individuo marcado y destinado a aprovechar).", "FGR-06_Parte_A.xlsx", "Corpoboyaca/prioritaria_emergencia_obra", "Debe acompañarse del registro fotográfico y georreferenciación según la guía.", SOURCE_CORPO),
            req("corpo_cedulas", "Fotocopias de cédula del o los propietarios del predio o representante legal según aplique.", required=True),
            req("corpo_existencia", "Certificado de existencia y representación legal (persona jurídica).", required=False, conditional="Aplica a personas jurídicas."),
            req("corpo_escritura", "Fotocopia de la(s) escritura(s) del predio objeto de aprovechamiento.", required=True),
            req("corpo_libertad", "Certificado de libertad y tradición con fecha de expedición no mayor a treinta (30) días.", required=True),
            req("corpo_autorizacion", "Autorización escrita del propietario del predio cuando se actúa como tenedor, si aplica.", required=False),
            req("corpo_fgr29", "Formato FGR-29 «Autodeclaración de costos de inversión y anual de operación».", "FGR-29.xlsx", "Corpoboyaca/prioritaria_emergencia_obra", "Formato oficial.", SOURCE_CORPO),
            req("corpo_pago", "Copia del recibo de consignación o factura de pago por servicios de evaluación ambiental a favor de Corpoboyacá.", required=False, conditional="Se incorpora después de la liquidación/pago, cuando corresponda."),
        ],
    },
    "CORPO_NATIVA_GT50": {
        "label": "Corpoboyacá — especies nativas con volumen superior a 50 m³",
        "procedure_title": "Aprovechamiento de árboles aislados fuera de cobertura de bosque natural",
        "procedure": "Para nativas >50 m³ se exige, entre otros, estudio técnico conforme a términos de referencia de Corpoboyacá.",
        "documents": [
            req("corpo_fun", "Formato único nacional de solicitud de aprovechamiento forestal y manejo sostenible de flora silvestre y productos forestales no maderables.", "FUN_Corpoboyaca.pdf", "Corpoboyaca/nativas_mayor_50m3", "Formato oficial.", SOURCE_CORPO),
            req("corpo_estudio", "Estudio técnico para el aprovechamiento de árboles aislados, conforme a los términos de referencia de Corpoboyacá.", "Terminos_referencia_arboles_aislados.pdf", "Corpoboyaca/nativas_mayor_50m3", "Documento técnico que debe elaborar el interesado.", SOURCE_CORPO),
            req("corpo_cedulas", "Fotocopias de cédula del o los propietarios del predio o representante legal según aplique.", required=True),
            req("corpo_existencia", "Certificado de existencia y representación legal (persona jurídica).", required=False, conditional="Aplica a personas jurídicas."),
            req("corpo_escritura", "Fotocopia de la(s) escritura(s) del predio objeto de aprovechamiento.", required=True),
            req("corpo_libertad", "Certificado de libertad y tradición con fecha de expedición no mayor a treinta (30) días.", required=True),
            req("corpo_autorizacion", "Autorización escrita del propietario del predio cuando se actúa como tenedor, si aplica.", required=False),
            req("corpo_fgr29", "Formato FGR-29 «Autodeclaración de costos de inversión y anual de operación».", "FGR-29.xlsx", "Corpoboyaca/nativas_mayor_50m3", "Formato oficial.", SOURCE_CORPO),
            req("corpo_pago", "Copia del recibo de consignación o factura de pago por servicios de evaluación ambiental a favor de Corpoboyacá.", required=False, conditional="Se incorpora después de la liquidación/pago, cuando corresponda."),
        ],
    },
    "CORPO_MENOR50_EXOTICA": {
        "label": "Corpoboyacá — nativas <50 m³ / exóticas o introducidas",
        "procedure_title": "Nativas con volumen inferior a 50 m³ y especies exóticas o introducidas sin límite de volumen",
        "procedure": "La guía oficial indica FUN + FGR-06 Parte A y B + soportes prediales/identificación + FGR-29 + soporte de pago cuando corresponda.",
        "documents": [
            req("corpo_fun", "Formato único nacional de solicitud de aprovechamiento forestal y manejo sostenible de flora silvestre y productos forestales no maderables.", "FUN_Corpoboyaca.pdf", "Corpoboyaca/nativas_menor_50m3_exoticas", "Formato oficial.", SOURCE_CORPO),
            req("corpo_fgr06ab", "Formato FGR-06 «Información para el aprovechamiento de árboles aislados» Parte A y B.", "FGR-06_Parte_A_y_B.xlsx", "Corpoboyaca/nativas_menor_50m3_exoticas", "Formato oficial.", SOURCE_CORPO),
            req("corpo_cedulas", "Fotocopias de cédula del o los propietarios del predio o representante legal según aplique.", required=True),
            req("corpo_existencia", "Certificado de existencia y representación legal (persona jurídica).", required=False, conditional="Aplica a personas jurídicas."),
            req("corpo_escritura", "Fotocopia de la(s) escritura(s) del predio objeto de aprovechamiento.", required=True),
            req("corpo_libertad", "Certificado de libertad y tradición con fecha de expedición no mayor a treinta (30) días.", required=True),
            req("corpo_autorizacion", "Autorización escrita del propietario del predio cuando se actúa como tenedor, si aplica.", required=False),
            req("corpo_fgr29", "Formato FGR-29 «Autodeclaración de costos de inversión y anual de operación».", "FGR-29.xlsx", "Corpoboyaca/nativas_menor_50m3_exoticas", "Formato oficial.", SOURCE_CORPO),
            req("corpo_pago", "Copia del recibo de consignación o factura de pago por servicios de evaluación ambiental a favor de Corpoboyacá.", required=False, conditional="Se incorpora después de la liquidación/pago, cuando corresponda."),
        ],
    },
    "CORPO_DOMESTICO": {
        "label": "Corpoboyacá — uso doméstico",
        "procedure_title": "Aprovechamiento de árboles aislados para uso doméstico",
        "procedure": "La guía oficial indica que aplica en predios privados, sin comercialización; volumen máximo 10 m³ para nativas y 20 m³ para exóticas/introducidas. No requiere pago por servicios de evaluación ambiental según la guía.",
        "documents": [
            req("corpo_fun", "Formato único nacional de solicitud de aprovechamiento forestal y manejo sostenible de flora silvestre y productos forestales no maderables.", "FUN_Corpoboyaca.pdf", "Corpoboyaca/uso_domestico", "Formato oficial.", SOURCE_CORPO),
            req("corpo_fgr06a", "Formato FGR-06 Parte A «Inventario Forestal al 100%» (incluir registro fotográfico y georreferenciación por individuo marcado y destinado a aprovechar).", "FGR-06_Parte_A.xlsx", "Corpoboyaca/uso_domestico", "Formato oficial.", SOURCE_CORPO),
            req("corpo_cedulas", "Fotocopias de cédula del o los propietarios del predio.", required=True),
            req("corpo_escritura", "Fotocopia de la Escritura Pública.", required=True),
            req("corpo_libertad", "Certificado de libertad con fecha de expedición no mayor a 30 días.", required=True),
        ],
    },
    "GENERAL": {
        "label": "General / autoridad por definir",
        "procedure_title": "Prealistamiento general",
        "procedure": "No se presenta como un checklist legal. Sirve para organizar la información hasta identificar la autoridad competente.",
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
    "General / autoridad por definir": "GENERAL",
}
ALLOWED_DOC_EXTENSIONS = {".pdf", ".doc", ".docx", ".xls", ".xlsx", ".csv", ".jpg", ".jpeg", ".png", ".webp"}

# =============================================================================
# EXPEDIENTES PERSISTENTES
# =============================================================================
def default_state(username: str):
    return {
        "case_id": datetime.now().strftime("ER-%Y%m%d-%H%M%S") + "-" + secrets.token_hex(2).upper(),
        "owner": username,
        "created_at": datetime.now().isoformat(),
        "updated_at": datetime.now().isoformat(),
        "status": "En diligenciamiento",
        "request_page": "Proyecto",
        "project_type": "Aprovechamiento forestal de árboles aislados",
        "authority_label": "SDA (Bogotá D.C.)",
        "authority_key": "SDA",
        "corpo_case": "CORPO_PRIORITARIA",
        "solicitud": {},
        "trees": [],
        "uploaded_docs": {},
    }


def case_json_path(case_id: str) -> Path:
    folder = CASE_DIR / case_id
    folder.mkdir(parents=True, exist_ok=True)
    return folder / "case.json"


def save_case():
    if not st.session_state.get("authenticated") or st.session_state.get("user", {}).get("role") != "client":
        return
    data = {
        "case_id": st.session_state["case_id"],
        "owner": st.session_state["user"]["username"],
        "created_at": st.session_state.get("created_at", datetime.now().isoformat()),
        "updated_at": datetime.now().isoformat(),
        "status": st.session_state.get("case_status", "En diligenciamiento"),
        "request_page": st.session_state.get("request_page", "Proyecto"),
        "project_type": st.session_state.get("project_type"),
        "authority_label": st.session_state.get("authority_label"),
        "authority_key": st.session_state.get("authority_key"),
        "corpo_case": st.session_state.get("corpo_case"),
        "solicitud": st.session_state.get("solicitud", {}),
        "trees": st.session_state.get("trees", []),
        "uploaded_docs": st.session_state.get("uploaded_docs", {}),
    }
    # Signatures are only used in memory to avoid duplicate uploads.
    data["uploaded_docs"].pop("__signatures__", None)
    path = case_json_path(data["case_id"])
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    st.session_state["updated_at"] = data["updated_at"]


def load_case(case_id: str):
    path = case_json_path(case_id)
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


def load_or_create_user_case(username: str):
    cases = []
    for path in CASE_DIR.glob("*/case.json"):
        data = load_case(path.parent.name)
        if data and data.get("owner") == username:
            cases.append(data)
    if cases:
        data = sorted(cases, key=lambda x: x.get("updated_at", ""), reverse=True)[0]
    else:
        data = default_state(username)
        case_json_path(data["case_id"]).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    for key, value in data.items():
        if key != "uploaded_docs":
            st.session_state[key] = value
    st.session_state["uploaded_docs"] = data.get("uploaded_docs", {})
    st.session_state["authenticated"] = True


def new_case():
    data = default_state(st.session_state["user"]["username"])
    for key, value in data.items():
        st.session_state[key] = value
    save_case()
    st.rerun()


def list_cases():
    rows = []
    for path in CASE_DIR.glob("*/case.json"):
        data = load_case(path.parent.name)
        if not data:
            continue
        reqs = requirements_for_data(data)
        uploaded = data.get("uploaded_docs", {})
        required_count = sum(1 for r in reqs if r.get("required", True))
        uploaded_required = sum(1 for r in reqs if r.get("required", True) and uploaded.get(r["id"]))
        rows.append({
            "Expediente": data.get("case_id"),
            "Cliente": data.get("owner"),
            "Autoridad": data.get("authority_label"),
            "Estado": data.get("status"),
            "Documentos": f"{uploaded_required}/{required_count}",
            "Actualizado": data.get("updated_at", "")[:16].replace("T", " "),
        })
    return pd.DataFrame(rows)


def requirements_for_data(data: dict):
    authority = data.get("authority_key", "GENERAL")
    if authority == "SDA":
        return REQUIREMENTS["SDA"]["documents"]
    if authority == "CAR":
        return REQUIREMENTS["CAR"]["documents"]
    if authority == "CORPO":
        return REQUIREMENTS.get(data.get("corpo_case", "CORPO_PRIORITARIA"), REQUIREMENTS["CORPO_PRIORITARIA"])["documents"]
    return []


def get_requirement_group():
    authority = st.session_state.get("authority_key", "GENERAL")
    if authority == "SDA": return "SDA"
    if authority == "CAR": return "CAR"
    if authority == "CORPO": return st.session_state.get("corpo_case", "CORPO_PRIORITARIA")
    return "GENERAL"


def get_requirements():
    return REQUIREMENTS[get_requirement_group()]


def current_case_id():
    return st.session_state.get("case_id", "sin_expediente")


def case_folder():
    folder = CASE_DIR / current_case_id()
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def doc_upload_folder():
    folder = case_folder() / "documentos_diligenciados"
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def safe_file_name(filename: str) -> str:
    return re.sub(r"[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ._() -]", "_", os.path.basename(filename))


def format_file_size(path: Path) -> str:
    if not path.exists(): return "0 B"
    size = path.stat().st_size
    if size < 1024: return f"{size} B"
    if size < 1024 * 1024: return f"{size / 1024:.1f} KB"
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


def relative_path(path: Path) -> str:
    return str(path.relative_to(BASE_DIR)).replace("\\", "/")


def absolute_from_stored(value: str) -> Path:
    path = Path(value)
    if not path.is_absolute():
        path = BASE_DIR / path
    return path


def requirement_uploaded(req_id: str):
    values = st.session_state.get("uploaded_docs", {}).get(req_id, [])
    paths = []
    for value in values:
        path = absolute_from_stored(value)
        if path.exists(): paths.append(path)
    return paths


def uploaded_signature(uploaded):
    return hashlib.sha256(uploaded.getvalue()).hexdigest()


def store_requirement_uploads(req_id: str, uploaded_files):
    registry = st.session_state.setdefault("uploaded_docs", {})
    saved = registry.setdefault(req_id, [])
    seen = st.session_state.setdefault("upload_signatures", {}).setdefault(req_id, set())
    for uploaded in uploaded_files or []:
        ext = Path(uploaded.name).suffix.lower()
        if ext not in ALLOWED_DOC_EXTENSIONS:
            st.error(f"Formato no permitido: {uploaded.name}")
            continue
        sig = uploaded_signature(uploaded)
        if sig in seen:
            continue
        path = save_uploaded_file(uploaded, doc_upload_folder(), prefix=f"{slug(req_id)}__")
        saved.append(relative_path(path))
        seen.add(sig)
    save_case()


def requirement_status(req):
    uploaded = requirement_uploaded(req["id"])
    if uploaded:
        return "SUBIDO", "ok"
    if req.get("required", True):
        return "PENDIENTE", "pending"
    return "OPCIONAL / PENDIENTE", "info"

# =============================================================================
# CÁLCULOS Y DOCUMENTO FINAL
# =============================================================================
def ensure_tree_count(n: int):
    trees = st.session_state.setdefault("trees", [])
    while len(trees) < n:
        trees.append({
            "nombre_comun": "", "nombre_cientifico": "", "dap_cm": 0.0, "altura_m": 0.0,
            "factor_nombre": "Paraboloide (0.55)", "riesgo": False, "fauna": False,
            "epifitas": False, "observaciones": "", "foto_path": "", "foto_name": "",
        })
    if len(trees) > n:
        del trees[n:]


def volume_for_tree(tree):
    dap = float(tree.get("dap_cm", 0) or 0) / 100.0
    height = float(tree.get("altura_m", 0) or 0)
    factor_text = tree.get("factor_nombre", "Paraboloide (0.55)")
    try: factor = float(factor_text.split("(")[1].split(")")[0])
    except Exception: factor = 0.55
    if dap <= 0 or height <= 0: return 0.0
    return (math.pi / 4) * (dap ** 2) * height * factor


def total_volume():
    return sum(volume_for_tree(t) for t in st.session_state.get("trees", []))


def priority_text():
    trees = st.session_state.get("trees", [])
    if any(t.get("riesgo") for t in trees): return "CRÍTICA — Riesgo inminente reportado"
    if any(t.get("fauna") or t.get("epifitas") for t in trees): return "ALTA — Restricción ambiental reportada"
    return "ESTÁNDAR"


def core_validation():
    sol = st.session_state.get("solicitud", {})
    missing = []
    fields = {
        "Nombre / razón social": sol.get("nombre"),
        "Número de documento": sol.get("num_doc"),
        "Correo": sol.get("correo"),
        "Nombre del predio": sol.get("nombre_predio"),
        "Municipio": sol.get("municipio"),
        "Dirección": sol.get("direccion"),
    }
    for label, value in fields.items():
        if not value: missing.append(label)
    trees_ok = bool(st.session_state.get("trees")) and all(
        t.get("nombre_comun") and t.get("dap_cm", 0) > 0 and t.get("altura_m", 0) > 0
        for t in st.session_state.get("trees", [])
    )
    return missing, trees_ok


def create_technical_doc():
    doc = Document()
    doc.add_heading("EXPEDIENTE DE PREALISTAMIENTO — ECO REGIÓN APP", 0)
    doc.add_paragraph(f"Expediente: {current_case_id()}")
    doc.add_paragraph(f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    doc.add_paragraph(
        "Este documento consolida la información suministrada en EcoRegión App y la matriz documental de prealistamiento. "
        "No sustituye los formatos oficiales ni la decisión de la autoridad ambiental competente. Los formatos oficiales deben diligenciarse y anexarse por separado."
    )
    sol = st.session_state.get("solicitud", {})
    doc.add_heading("1. Identificación del trámite", level=1)
    doc.add_paragraph(f"Tipo de trámite: {st.session_state.get('project_type', '')}")
    doc.add_paragraph(f"Autoridad: {st.session_state.get('authority_label', '')}")
    doc.add_paragraph(f"Modalidad: {get_requirements()['label']}")
    doc.add_paragraph(f"Descripción: {sol.get('descripcion', '')}")

    doc.add_heading("2. Solicitante", level=1)
    for label, key in [
        ("Nombre / razón social", "nombre"), ("Tipo de documento", "tipo_doc"), ("Número", "num_doc"),
        ("Calidad sobre el predio", "calidad"), ("Correo", "correo"), ("Teléfono", "telefono"),
        ("Representante legal", "representante_legal"),
    ]:
        doc.add_paragraph(f"{label}: {sol.get(key, '')}")

    doc.add_heading("3. Predio y ubicación", level=1)
    for label, key in [
        ("Nombre del predio / institución", "nombre_predio"), ("Municipio", "municipio"),
        ("Dirección", "direccion"), ("Matrícula / referencia", "matricula"),
        ("Latitud", "latitud"), ("Longitud", "longitud"),
    ]:
        doc.add_paragraph(f"{label}: {sol.get(key, '')}")

    doc.add_heading("4. Inventario y caracterización arbórea", level=1)
    table = doc.add_table(rows=1, cols=7)
    hdr = table.rows[0].cells
    for i, text in enumerate(["#", "Nombre común", "Nombre científico", "DAP cm", "Altura m", "Volumen m³", "Riesgos / restricciones"]):
        hdr[i].text = text
    for i, tree in enumerate(st.session_state.get("trees", []), start=1):
        row = table.add_row().cells
        flags = []
        if tree.get("riesgo"): flags.append("Riesgo")
        if tree.get("fauna"): flags.append("Avifauna")
        if tree.get("epifitas"): flags.append("Epífitas")
        vals = [i, tree.get("nombre_comun", ""), tree.get("nombre_cientifico", ""), tree.get("dap_cm", ""), tree.get("altura_m", ""), f"{volume_for_tree(tree):.3f}", ", ".join(flags) or "Ninguno reportado"]
        for j, val in enumerate(vals): row[j].text = str(val)
        if tree.get("observaciones"): doc.add_paragraph(f"Individuo {i} — Observaciones: {tree['observaciones']}")
    doc.add_paragraph(f"Volumen total estimado: {total_volume():.3f} m³")
    doc.add_paragraph(f"Prioridad interna de revisión: {priority_text()}")

    doc.add_heading("5. Matriz de requisitos documentales", level=1)
    table2 = doc.add_table(rows=1, cols=4)
    for i, text in enumerate(["Estado", "Requisito", "Condición", "Archivo"]): table2.rows[0].cells[i].text = text
    for req in get_requirements()["documents"]:
        status, _ = requirement_status(req)
        paths = requirement_uploaded(req["id"])
        file_names = ", ".join(p.name for p in paths) if paths else "—"
        row = table2.add_row().cells
        row[0].text = status
        row[1].text = req["name"]
        row[2].text = req.get("conditional", "")
        row[3].text = file_names

    doc.add_heading("6. Revisión final", level=1)
    missing_core, trees_ok = core_validation()
    doc.add_paragraph(f"Información básica: {'COMPLETA' if not missing_core else 'PENDIENTE — ' + ', '.join(missing_core)}")
    doc.add_paragraph(f"Inventario arbóreo: {'COMPLETO' if trees_ok else 'PENDIENTE'}")
    required_reqs = [r for r in get_requirements()["documents"] if r.get("required", True)]
    missing_req = [r["name"] for r in required_reqs if not requirement_uploaded(r["id"])]
    doc.add_paragraph(f"Requisitos documentales obligatorios cargados: {len(required_reqs)-len(missing_req)}/{len(required_reqs)}")
    if missing_req:
        doc.add_paragraph("Pendientes documentales:")
        for item in missing_req: doc.add_paragraph(item, style="List Bullet")

    doc.add_heading("7. Observación de alcance", level=1)
    doc.add_paragraph(
        "La aplicación funciona como herramienta de prealistamiento, organización y seguimiento. La suficiencia final de la documentación, "
        "la procedencia de cada formato, las visitas técnicas, liquidaciones, pagos, requerimientos y la decisión administrativa corresponden a la autoridad competente."
    )
    stream = io.BytesIO(); doc.save(stream); stream.seek(0)
    return stream.getvalue()


def build_case_zip():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("01_Expediente_PreAlistamiento_EcoRegion.docx", create_technical_doc())
        for req in get_requirements()["documents"]:
            for path in requirement_uploaded(req["id"]):
                zf.write(path, arcname=f"02_Documentos_Diligenciados/{path.name}")
        additional = case_folder() / "documentos_adicionales"
        if additional.exists():
            for path in additional.iterdir():
                if path.is_file(): zf.write(path, arcname=f"03_Documentos_Adicionales/{path.name}")
        evidence = case_folder() / "evidencia_arboles"
        if evidence.exists():
            for path in evidence.iterdir():
                if path.is_file(): zf.write(path, arcname=f"04_Evidencia_Arboles/{path.name}")
        meta = {"case_id": current_case_id(), "owner": st.session_state.get("user", {}).get("username"), "generated_at": datetime.now().isoformat(), "authority": st.session_state.get("authority_label"), "group": get_requirement_group(), "volume_m3": round(total_volume(), 3)}
        zf.writestr("00_metadata.json", json.dumps(meta, ensure_ascii=False, indent=2))
    buf.seek(0); return buf.getvalue()

# =============================================================================
# ESTADO DE SESIÓN / LOGIN
# =============================================================================
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    login_page()
    st.stop()

# =============================================================================
# INICIALIZACIÓN VISUAL
# =============================================================================
user = st.session_state["user"]
if "page" not in st.session_state:
    st.session_state["page"] = "Inicio"
if "request_page" not in st.session_state:
    st.session_state["request_page"] = "Proyecto"
if "upload_signatures" not in st.session_state:
    st.session_state["upload_signatures"] = {}

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
        """, unsafe_allow_html=True,
    )
    st.markdown(f"<div class='role-chip'>{'ADMINISTRADOR' if user['role']=='admin' else 'CLIENTE'}</div>", unsafe_allow_html=True)
    st.caption(f"Usuario: {user['username']}")
    st.divider()

    if user["role"] == "admin":
        main_options = ["Inicio", "Nueva solicitud", "Requisitos", "Autoridades", "Documentación", "Seguimiento", "Administración", "Encuesta de validación", "Manual de uso"]
    else:
        main_options = ["Inicio", "Nueva solicitud", "Requisitos", "Autoridades", "Documentación", "Seguimiento", "Encuesta de validación", "Manual de uso"]

    st.session_state["page"] = st.radio("MENÚ PRINCIPAL", main_options, index=main_options.index(st.session_state.get("page", "Inicio")))
    st.divider()
    st.caption(f"Expediente actual: {current_case_id()}")
    if user["role"] == "client":
        if st.button("＋ Nueva solicitud", use_container_width=True):
            new_case()
    if st.button("Cerrar sesión", use_container_width=True):
        logout()

# =============================================================================
# HERO
# =============================================================================
st.markdown(
    f"""
    <div class='hero'>
        <h1>🌿 {APP_TITLE}</h1>
        <p>Expedientes, requisitos, documentación y seguimiento para el aprovechamiento forestal de árboles aislados.</p>
    </div>
    """, unsafe_allow_html=True,
)

# =============================================================================
# PÁGINAS
# =============================================================================
def page_inicio():
    st.subheader("Panel de control")
    reqs = get_requirements()["documents"]
    required = [r for r in reqs if r.get("required", True)]
    uploaded_required = sum(bool(requirement_uploaded(r["id"])) for r in required)
    sol = st.session_state.get("solicitud", {})
    cols = st.columns(4)
    metrics = [
        ("Autoridad", st.session_state.get("authority_label", "Sin definir")),
        ("Árboles", str(len(st.session_state.get("trees", [])))),
        ("Requisitos", f"{uploaded_required}/{len(required)}"),
        ("Estado", st.session_state.get("case_status", "En diligenciamiento")),
    ]
    for col, (title, value) in zip(cols, metrics):
        with col: st.markdown(f"<div class='kpi'><div class='title'>{title}</div><div class='value'>{value}</div></div>", unsafe_allow_html=True)

    st.write("")
    left, right = st.columns([1.25, 1])
    with left:
        st.markdown("### Flujo de trabajo")
        st.markdown("**1. Proyecto** → autoridad y modalidad  \\n**2. Solicitante** → identificación y contacto  \\n**3. Predio** → ubicación y georreferenciación  \\n**4. Árboles** → inventario y caracterización  \\n**5. Validación** → documentos, revisión y expediente")
        st.markdown("### Regla documental")
        st.info("La autoridad seleccionada determina automáticamente el checklist. En Corpoboyacá, la modalidad también cambia los requisitos.")
    with right:
        st.markdown("### Estado del expediente")
        if sol.get("nombre"): st.success(f"Solicitante: {sol['nombre']}")
        else: st.warning("Falta diligenciar el solicitante.")
        st.write(f"**Volumen estimado:** {total_volume():.3f} m³")
        st.write(f"**Revisión interna:** {priority_text()}")
        st.write(f"**Última actualización:** {st.session_state.get('updated_at','')[:16].replace('T',' ')}")


def render_request_tabs():
    options = ["Proyecto", "Solicitante", "Predio", "Árboles", "Validación"]
    with st.container():
        st.markdown("<div class='stepbar'>", unsafe_allow_html=True)
        selected = st.radio("Etapas", options, index=options.index(st.session_state.get("request_page", "Proyecto")), horizontal=True, label_visibility="collapsed", key="request_tabs_v3")
        st.markdown("</div>", unsafe_allow_html=True)
    st.session_state["request_page"] = selected


def page_nueva_solicitud():
    st.subheader("Nueva solicitud")
    st.caption("Siga las cinco etapas. La navegación de la solicitud ahora está arriba para facilitar el recorrido.")
    render_request_tabs()
    step = st.session_state["request_page"]

    if step == "Proyecto":
        with st.form("form_proyecto_v3"):
            st.markdown("### 1 · Proyecto y autoridad")
            project_type = st.selectbox("Tipo de trámite", ["Aprovechamiento forestal de árboles aislados"], index=0)
            authority_label = st.selectbox("Autoridad ambiental competente", list(AUTHORITY_OPTIONS.keys()), index=list(AUTHORITY_OPTIONS.keys()).index(st.session_state.get("authority_label", "SDA (Bogotá D.C.)")))
            authority_key = AUTHORITY_OPTIONS[authority_label]
            corpo_case_key = st.session_state.get("corpo_case", "CORPO_PRIORITARIA")
            if authority_key == "CORPO":
                selected = st.selectbox("Modalidad Corpoboyacá", list(CORPO_OPTIONS.keys()), index=list(CORPO_OPTIONS.values()).index(corpo_case_key))
                corpo_case_key = CORPO_OPTIONS[selected]
                st.info(REQUIREMENTS[corpo_case_key]["procedure"])
            sol = st.session_state.setdefault("solicitud", {})
            descripcion = st.text_area("Descripción / objeto de la solicitud", sol.get("descripcion", ""), placeholder="Describa la intervención y motivo de la solicitud.")
            if st.form_submit_button("Guardar proyecto", type="primary"):
                old_group = get_requirement_group()
                st.session_state.update({"project_type": project_type, "authority_label": authority_label, "authority_key": authority_key, "corpo_case": corpo_case_key})
                st.session_state["solicitud"]["descripcion"] = descripcion
                new_group = get_requirement_group()
                if old_group != new_group:
                    st.session_state["uploaded_docs"] = {}
                    st.warning("Cambió la autoridad/modalidad: se reinició la matriz documental de esta solicitud para evitar mezclar requisitos.")
                save_case(); st.success("Proyecto guardado.")
    elif step == "Solicitante":
        sol = st.session_state.setdefault("solicitud", {})
        with st.form("form_solicitante_v3"):
            st.markdown("### 2 · Datos del solicitante")
            c1, c2 = st.columns(2)
            with c1:
                sol["nombre"] = st.text_input("Nombre o razón social", sol.get("nombre", ""))
                tipos = ["CC", "NIT", "CE", "PASAPORTE"]
                sol["tipo_doc"] = st.selectbox("Tipo de documento", tipos, index=tipos.index(sol.get("tipo_doc", "CC")))
                sol["num_doc"] = st.text_input("Número de documento", sol.get("num_doc", ""))
                calidades = ["Propietario", "Tenedor", "Poseedor", "Apoderado / Autorizado"]
                sol["calidad"] = st.selectbox("Calidad sobre el predio", calidades, index=calidades.index(sol.get("calidad", "Propietario")))
            with c2:
                sol["correo"] = st.text_input("Correo electrónico", sol.get("correo", ""))
                sol["telefono"] = st.text_input("Teléfono", sol.get("telefono", ""))
                sol["representante_legal"] = st.text_input("Representante legal (si aplica)", sol.get("representante_legal", ""))
                sol["es_persona_juridica"] = st.checkbox("El solicitante es persona jurídica", value=bool(sol.get("es_persona_juridica", False)))
            if st.form_submit_button("Guardar solicitante", type="primary"):
                save_case(); st.success("Datos del solicitante guardados.")
    elif step == "Predio":
        sol = st.session_state.setdefault("solicitud", {})
        with st.form("form_predio_v3"):
            st.markdown("### 3 · Predio y georreferenciación")
            c1, c2 = st.columns(2)
            with c1:
                sol["nombre_predio"] = st.text_input("Nombre del predio / institución", sol.get("nombre_predio", ""))
                sol["municipio"] = st.text_input("Municipio / ciudad", sol.get("municipio", ""))
                sol["direccion"] = st.text_input("Dirección", sol.get("direccion", ""))
                sol["matricula"] = st.text_input("Matrícula / referencia predial", sol.get("matricula", ""))
            with c2:
                sol["latitud"] = st.number_input("Latitud", value=float(sol.get("latitud", 4.7110)), format="%.6f")
                sol["longitud"] = st.number_input("Longitud", value=float(sol.get("longitud", -74.0721)), format="%.6f")
                sol["tipo_predio"] = st.selectbox("Tipo de predio", ["Propiedad privada", "Dominio público", "Otro / por verificar"], index=["Propiedad privada", "Dominio público", "Otro / por verificar"].index(sol.get("tipo_predio", "Propiedad privada")))
            st.map(pd.DataFrame({"lat": [sol["latitud"]], "lon": [sol["longitud"]]}), zoom=12)
            if st.form_submit_button("Guardar predio", type="primary"):
                save_case(); st.success("Predio guardado.")
    elif step == "Árboles":
        st.markdown("### 4 · Caracterización de árboles")
        st.caption("Registre los individuos que serán incluidos en el expediente. El volumen mostrado es una estimación interna; los valores oficiales dependen de la metodología y formato aplicables.")
        current = len(st.session_state.get("trees", [])) or 1
        n = st.number_input("Número de individuos", min_value=1, max_value=500, value=current, step=1)
        ensure_tree_count(int(n))
        for i, tree in enumerate(st.session_state["trees"], start=1):
            with st.expander(f"🌳 Individuo {i} · {tree.get('nombre_comun') or 'Sin identificar'}", expanded=(i == 1)):
                a, b = st.columns(2)
                with a:
                    tree["nombre_comun"] = st.text_input(f"Nombre común · {i}", tree.get("nombre_comun", ""), key=f"v3_common_{i}")
                    tree["nombre_cientifico"] = st.text_input(f"Nombre científico · {i}", tree.get("nombre_cientifico", ""), key=f"v3_sci_{i}")
                    tree["dap_cm"] = st.number_input(f"DAP (cm) · {i}", min_value=0.0, value=float(tree.get("dap_cm", 0)), step=0.1, key=f"v3_dap_{i}")
                    tree["altura_m"] = st.number_input(f"Altura total (m) · {i}", min_value=0.0, value=float(tree.get("altura_m", 0)), step=0.1, key=f"v3_h_{i}")
                with b:
                    factors = ["Paraboloide (0.55)", "Cilíndrico (0.75)", "Cónico (0.33)", "Neiloide (0.25)"]
                    tree["factor_nombre"] = st.selectbox(f"Factor de forma · {i}", factors, index=factors.index(tree.get("factor_nombre", "Paraboloide (0.55)")), key=f"v3_factor_{i}")
                    tree["riesgo"] = st.checkbox("Riesgo inminente de caída", bool(tree.get("riesgo", False)), key=f"v3_risk_{i}")
                    tree["fauna"] = st.checkbox("Avifauna / nidos", bool(tree.get("fauna", False)), key=f"v3_fauna_{i}")
                    tree["epifitas"] = st.checkbox("Epífitas", bool(tree.get("epifitas", False)), key=f"v3_epi_{i}")
                    tree["observaciones"] = st.text_area(f"Observaciones · {i}", tree.get("observaciones", ""), key=f"v3_obs_{i}")
                    photo = st.file_uploader(f"Fotografía · {i}", type=["jpg", "jpeg", "png", "webp"], key=f"v3_photo_{i}")
                    if photo:
                        sig = uploaded_signature(photo)
                        if tree.get("foto_signature") != sig:
                            path = save_uploaded_file(photo, case_folder() / "evidencia_arboles", prefix=f"arbol_{i}__")
                            tree["foto_path"] = relative_path(path); tree["foto_name"] = photo.name; tree["foto_signature"] = sig
                        st.image(photo, width=250)
                st.success(f"Volumen estimado: {volume_for_tree(tree):.3f} m³")
        if st.button("Guardar caracterización", type="primary"):
            save_case(); st.session_state["request_page"] = "Validación"; st.rerun()
    elif step == "Validación":
        page_validacion_solicitud()


def page_validacion_solicitud():
    st.markdown("### 5 · Validación y expediente")
    missing_core, trees_ok = core_validation()
    reqs = get_requirements()["documents"]
    required_reqs = [r for r in reqs if r.get("required", True)]
    uploaded_required = [r for r in required_reqs if requirement_uploaded(r["id"])]
    missing_required = [r for r in required_reqs if not requirement_uploaded(r["id"])]

    c1, c2, c3, c4 = st.columns(4)
    for col, title, value in [
        (c1, "Información", "Completa" if not missing_core else "Pendiente"),
        (c2, "Árboles", "Completa" if trees_ok else "Pendiente"),
        (c3, "Documentos", f"{len(uploaded_required)}/{len(required_reqs)}"),
        (c4, "Volumen", f"{total_volume():.3f} m³"),
    ]:
        with col: st.markdown(f"<div class='kpi'><div class='title'>{title}</div><div class='value'>{value}</div></div>", unsafe_allow_html=True)

    if missing_core: st.warning("Información pendiente: " + ", ".join(missing_core))
    if not trees_ok: st.warning("Cada individuo debe tener al menos nombre común, DAP y altura para la revisión interna.")
    if missing_required: st.warning(f"Hay {len(missing_required)} requisito(s) obligatorio(s) todavía pendiente(s).")

    st.markdown("### Checklist documental")
    for req in reqs:
        uploads = requirement_uploaded(req["id"])
        status, css = requirement_status(req)
        with st.container(border=True):
            a, b, c = st.columns([2.25, .9, 1.25])
            with a:
                st.markdown(f"<div class='doc-title'>{req['name']}</div>", unsafe_allow_html=True)
                if req.get("conditional"): st.caption(req["conditional"])
                if req.get("note"): st.caption(req["note"])
            with b:
                icon = "✓" if status == "SUBIDO" else "!" if status == "PENDIENTE" else "i"
                st.markdown(f"<span class='status status-{css}'>{icon} {status}</span>", unsafe_allow_html=True)
                if uploads: st.caption(f"{len(uploads)} archivo(s)")
            with c:
                path = None
                if req.get("local_filename") and req.get("folder"):
                    candidate = DOCS_DIR / req["folder"] / req["local_filename"]
                    if candidate.exists(): path = candidate
                if path:
                    st.download_button("⬇ Descargar formato", path.read_bytes(), file_name=path.name, key=f"v3_dl_{req['id']}")
                elif req.get("source_url"):
                    st.link_button("Fuente oficial", req["source_url"])
                else:
                    st.caption("Plantilla no cargada")
            uploaded = st.file_uploader("Subir documento diligenciado", type=[x[1:] for x in sorted(ALLOWED_DOC_EXTENSIONS)], accept_multiple_files=True, key=f"v3_upl_{req['id']}")
            if uploaded:
                store_requirement_uploads(req["id"], uploaded)
                st.rerun()
            for path in uploads:
                st.caption(f"📄 {path.name} · {format_file_size(path)}")

    st.markdown("### Soportes adicionales")
    extra = st.file_uploader("Cargue aquí documentos adicionales que la autoridad o el caso requiera", type=[x[1:] for x in sorted(ALLOWED_DOC_EXTENSIONS)], accept_multiple_files=True, key="v3_extra")
    if extra:
        extra_dir = case_folder() / "documentos_adicionales"; extra_dir.mkdir(parents=True, exist_ok=True)
        for file in extra:
            save_uploaded_file(file, extra_dir, prefix="extra__")
        save_case(); st.rerun()

    st.markdown("### Resultado final")
    if not missing_core and trees_ok and not missing_required:
        st.success("Expediente documental obligatorio completo según el checklist configurado. Revise también los requisitos condicionales antes de radicar.")
        st.session_state["case_status"] = "Listo para revisión / radicación"
    else:
        st.info("El expediente todavía tiene pendientes. El sistema permite generar el documento para revisión, pero no lo marca como listo.")
        st.session_state["case_status"] = "Pendiente de completar"
    save_case()

    st.download_button("📄 Descargar expediente de prealistamiento (.docx)", create_technical_doc(), file_name=f"EcoRegion_Expediente_{current_case_id()}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary")
    st.download_button("🗂️ Descargar expediente completo (.zip)", build_case_zip(), file_name=f"EcoRegion_Expediente_{current_case_id()}.zip", mime="application/zip")


def page_requisitos():
    st.subheader("Requisitos documentales")
    st.caption("Los requisitos se muestran según la autoridad seleccionada. Las condiciones marcadas como opcionales/condicionales deben revisarse según el caso.")
    info = get_requirements()
    st.markdown(f"### {info['label']}")
    st.write(info["procedure"])
    if not info["documents"]:
        st.warning("No hay checklist oficial cargado para la categoría General.")
        return
    for i, req in enumerate(info["documents"], 1):
        status, css = requirement_status(req)
        st.markdown(f"**{i}. {req['name']}**")
        st.markdown(f"<span class='status status-{css}'>{status}</span>", unsafe_allow_html=True)
        if req.get("conditional"): st.caption(req["conditional"])
        path = DOCS_DIR / req["folder"] / req["local_filename"] if req.get("local_filename") and req.get("folder") else None
        if path and path.exists(): st.download_button("Descargar formato", path.read_bytes(), file_name=path.name, key=f"r3_{req['id']}")
        elif req.get("source_url"): st.link_button("Abrir fuente oficial", req["source_url"])


def page_autoridades():
    st.subheader("Autoridades y reglas documentales")
    st.info("La aplicación separa los checklists para evitar que un documento de una autoridad se mezcle con otra.")
    for title, key in [("SDA (Bogotá D.C.)", "SDA"), ("CAR Cundinamarca", "CAR")]:
        with st.expander(title, expanded=(st.session_state.get("authority_key") == key)):
            st.write(REQUIREMENTS[key]["procedure"])
            for req in REQUIREMENTS[key]["documents"]: st.write("•", req["name"])
            st.link_button("Fuente oficial", SOURCE_SDA if key == "SDA" else SOURCE_CAR)
    with st.expander("Corpoboyacá", expanded=(st.session_state.get("authority_key") == "CORPO")):
        for label, key in CORPO_OPTIONS.items():
            st.markdown(f"**{label}**")
            for req in REQUIREMENTS[key]["documents"]: st.write("•", req["name"])
        st.link_button("Guía oficial Corpoboyacá", SOURCE_CORPO)


def page_documentacion():
    st.subheader("Centro documental")
    reqs = get_requirements()["documents"]
    st.markdown("### Plantillas oficiales")
    for req in reqs:
        path = DOCS_DIR / req["folder"] / req["local_filename"] if req.get("local_filename") and req.get("folder") else None
        if path and path.exists():
            st.markdown(f"<span class='status status-ok'>✓ DISPONIBLE</span> &nbsp; **{req['name']}**", unsafe_allow_html=True)
            st.caption(f"{relative_path(path)} · {format_file_size(path)}")
        else:
            st.markdown(f"<span class='status status-pending'>! NO CARGADA</span> &nbsp; **{req['name']}**", unsafe_allow_html=True)
    st.markdown("### Documentos del expediente")
    any_file = False
    for req in reqs:
        for path in requirement_uploaded(req["id"]):
            any_file = True
            st.write(f"📄 {path.name} — {format_file_size(path)}")
    if not any_file: st.info("Todavía no hay documentos diligenciados en este expediente.")
    evidence = list((case_folder() / "evidencia_arboles").glob("*")) if (case_folder() / "evidencia_arboles").exists() else []
    if evidence:
        st.markdown("### Evidencia fotográfica")
        st.write(" · ".join(p.name for p in evidence))


def page_seguimiento():
    st.subheader("Seguimiento del expediente")
    missing_core, trees_ok = core_validation()
    reqs = get_requirements()["documents"]
    required = [r for r in reqs if r.get("required", True)]
    completed = sum(bool(requirement_uploaded(r["id"])) for r in required)
    steps = [
        ("Proyecto", bool(st.session_state.get("authority_key"))),
        ("Solicitante", not bool(missing_core[:2])),
        ("Predio", not bool(missing_core[2:])),
        ("Árboles", trees_ok),
        ("Documentos", completed == len(required) if required else True),
    ]
    for label, done in steps:
        if done: st.markdown(f"<span class='status status-ok'>✓ COMPLETADO</span> &nbsp; **{label}**", unsafe_allow_html=True)
        else: st.markdown(f"<span class='status status-pending'>! PENDIENTE</span> &nbsp; **{label}**", unsafe_allow_html=True)
    st.markdown("### Estado de cada documento")
    for req in reqs:
        status, css = requirement_status(req)
        paths = requirement_uploaded(req["id"])
        st.markdown(f"<span class='status status-{css}'>{status}</span> &nbsp; {req['name']}", unsafe_allow_html=True)
        if paths: st.caption("Archivos: " + ", ".join(p.name for p in paths))
    if completed == len(required) and not missing_core and trees_ok:
        st.success("Las etapas obligatorias del prototipo están completas. Verifique los requisitos condicionales y la información oficial antes de radicar.")
    else:
        st.warning("El expediente todavía tiene etapas o documentos pendientes.")


def page_encuesta():
    st.subheader("Encuesta de validación y usabilidad")
    with st.form("survey_v3"):
        q1 = st.slider("Facilidad de diligenciamiento", 1, 5, 5)
        q2 = st.slider("Claridad de requisitos según autoridad", 1, 5, 5)
        q3 = st.slider("Facilidad para encontrar y subir documentos", 1, 5, 5)
        q4 = st.slider("Comprensión del seguimiento", 1, 5, 5)
        q5 = st.slider("Utilidad del expediente generado", 1, 5, 5)
        comments = st.text_area("Comentarios")
        if st.form_submit_button("Guardar respuesta", type="primary"):
            row = pd.DataFrame([{"Fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Usuario": user["username"], "Facilidad": q1, "Claridad": q2, "Documentacion": q3, "Seguimiento": q4, "Expediente": q5, "Comentarios": comments}])
            if SURVEY_FILE.exists(): row.to_csv(SURVEY_FILE, mode="a", sep=";", index=False, header=False, encoding="utf-8-sig")
            else: row.to_csv(SURVEY_FILE, sep=";", index=False, encoding="utf-8-sig")
            st.success("Respuesta registrada.")
    if SURVEY_FILE.exists():
        df = pd.read_csv(SURVEY_FILE, sep=";", encoding="utf-8-sig")
        st.metric("Usuarios encuestados", len(df))
        cols = [c for c in ["Facilidad", "Claridad", "Documentacion", "Seguimiento", "Expediente"] if c in df.columns]
        if cols: st.bar_chart(df[cols].mean())


def page_manual():
    st.subheader("Manual de uso")
    st.markdown("### Acceso")
    st.write("Existen dos roles: **CLIENTE**, que diligencia solicitudes y carga documentos; y **ADMIN**, que administra usuarios, plantillas y consulta expedientes.")
    st.markdown("### Acceso inicial")
    if user["role"] == "admin":
        st.code("ADMIN   usuario: admin      contraseña: AdminEco2026!\nCLIENTE usuario: cliente    contraseña: ClienteEco2026!")
        st.warning("Son credenciales de demostración. Cámbielas desde Administración antes de usar información real.")
    else:
        st.info("Las credenciales son administradas por el equipo de EcoRegión. Si necesita acceso o restablecimiento de contraseña, solicítelo al administrador.")
    st.markdown("### Estructura documental")
    st.code("""documentos/\n├── SDA/\n├── CAR_Cundinamarca/\n└── Corpoboyaca/\n    ├── prioritaria_emergencia_obra/\n    ├── nativas_mayor_50m3/\n    ├── nativas_menor_50m3_exoticas/\n    └── uso_domestico/\n\ndata/expedientes/<ID>/\n├── case.json\n├── documentos_diligenciados/\n├── documentos_adicionales/\n└── evidencia_arboles/""", language="text")
    st.markdown("### Flujo documental")
    st.write("Seleccionar autoridad/modalidad → descargar formato → diligenciar → subir → verificar estado → generar expediente.")
    st.markdown("### Alcance")
    st.info("El DOCX generado consolida el prealistamiento y la información del expediente. No reemplaza formularios oficiales, estudios técnicos exigibles, liquidaciones, pagos ni la decisión de la autoridad ambiental.")

# =============================================================================
# ADMINISTRACIÓN
# =============================================================================
def page_admin():
    st.subheader("Administración")
    if user["role"] != "admin":
        st.error("Acceso restringido.")
        return
    tabs = st.tabs(["Usuarios", "Plantillas", "Expedientes"])
    with tabs[0]:
        st.markdown("### Usuarios")
        users = load_users()
        for uname, u in users.items():
            a, b, c = st.columns([1.5, 1, 1])
            with a: st.write(f"**{uname}** — {u.get('name','')}")
            with b: st.write("ADMIN" if u.get("role") == "admin" else "CLIENTE")
            with c: st.write("Activo" if u.get("active", True) else "Inactivo")
        st.markdown("### Crear usuario cliente")
        with st.form("create_client"):
            uname = st.text_input("Usuario nuevo")
            name = st.text_input("Nombre")
            pwd = st.text_input("Contraseña temporal", type="password")
            if st.form_submit_button("Crear cliente", type="primary"):
                if not uname or not pwd: st.error("Usuario y contraseña son obligatorios.")
                elif uname in users: st.error("Ese usuario ya existe.")
                else:
                    users[uname] = {"name": name or uname, "role": "client", "password_hash": hash_password(pwd), "active": True}
                    save_users(users); st.success("Cliente creado."); st.rerun()
        st.markdown("### Cambiar contraseña")
        with st.form("change_pwd"):
            target = st.selectbox("Usuario", list(users.keys()))
            new_pwd = st.text_input("Nueva contraseña", type="password")
            if st.form_submit_button("Actualizar contraseña"):
                if len(new_pwd) < 8: st.error("Use al menos 8 caracteres.")
                else:
                    users[target]["password_hash"] = hash_password(new_pwd); save_users(users); st.success("Contraseña actualizada.")
    with tabs[1]:
        st.markdown("### Carga de formatos oficiales")
        st.info("El administrador puede subir aquí los archivos oficiales. La app los guardará en la carpeta exacta que corresponde al requisito seleccionado.")
        group = st.selectbox("Grupo documental", list(REQUIREMENTS.keys()))
        reqs = REQUIREMENTS[group]["documents"]
        if reqs:
            rid = st.selectbox("Requisito", [r["id"] for r in reqs], format_func=lambda x: next(r["name"] for r in reqs if r["id"] == x))
            chosen = next(r for r in reqs if r["id"] == rid)
            st.caption(f"Destino: documentos/{chosen.get('folder','')}/{chosen.get('local_filename','')}")
            up = st.file_uploader("Archivo oficial", type=[x[1:] for x in sorted(ALLOWED_DOC_EXTENSIONS)], key=f"admin_template_{group}_{rid}")
            if up and st.button("Guardar plantilla oficial", type="primary"):
                if not chosen.get("folder") or not chosen.get("local_filename"):
                    st.error("Este requisito no tiene una plantilla local definida; se maneja como soporte que sube el cliente.")
                else:
                    dest = DOCS_DIR / chosen["folder"]
                    dest.mkdir(parents=True, exist_ok=True)
                    (dest / chosen["local_filename"]).write_bytes(up.getvalue())
                    st.success(f"Plantilla guardada en {dest / chosen['local_filename']}")
        else:
            st.info("Este grupo no tiene documentos configurados.")
    with tabs[2]:
        st.markdown("### Expedientes de clientes")
        df = list_cases()
        if df.empty: st.info("No hay expedientes todavía.")
        else:
            st.dataframe(df, use_container_width=True, hide_index=True)
            selected_case = st.selectbox("Seleccionar expediente", df["Expediente"].tolist())
            data = load_case(selected_case)
            if data:
                st.json({k: data[k] for k in ["case_id", "owner", "authority_label", "status", "updated_at"] if k in data})
                st.download_button("Descargar expediente del cliente", build_admin_case_zip(data), file_name=f"EcoRegion_{selected_case}.zip", mime="application/zip")


def build_admin_case_zip(data):
    buf = io.BytesIO()
    folder = CASE_DIR / data["case_id"]
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        case_doc = create_doc_from_data(data)
        zf.writestr("01_Expediente_PreAlistamiento_EcoRegion.docx", case_doc)
        for req in requirements_for_data(data):
            for value in data.get("uploaded_docs", {}).get(req["id"], []):
                path = absolute_from_stored(value)
                if path.exists(): zf.write(path, arcname=f"02_Documentos_Diligenciados/{path.name}")
        for sub, arc in [("documentos_adicionales", "03_Documentos_Adicionales"), ("evidencia_arboles", "04_Evidencia_Arboles")]:
            d = folder / sub
            if d.exists():
                for p in d.iterdir():
                    if p.is_file(): zf.write(p, arcname=f"{arc}/{p.name}")
    buf.seek(0); return buf.getvalue()


def create_doc_from_data(data):
    # Temporarily render an admin-selected case through a document built from stored data.
    doc = Document()
    doc.add_heading("EXPEDIENTE DE PREALISTAMIENTO — ECO REGIÓN APP", 0)
    doc.add_paragraph(f"Expediente: {data.get('case_id','')}")
    doc.add_paragraph(f"Cliente: {data.get('owner','')}")
    doc.add_paragraph(f"Autoridad: {data.get('authority_label','')}")
    doc.add_paragraph(f"Estado: {data.get('status','')}")
    sol = data.get("solicitud", {})
    doc.add_heading("Solicitante", level=1)
    for label, key in [("Nombre / razón social","nombre"),("Documento","num_doc"),("Correo","correo"),("Teléfono","telefono")]: doc.add_paragraph(f"{label}: {sol.get(key,'')}")
    doc.add_heading("Predio", level=1)
    for label, key in [("Predio","nombre_predio"),("Municipio","municipio"),("Dirección","direccion"),("Latitud","latitud"),("Longitud","longitud")]: doc.add_paragraph(f"{label}: {sol.get(key,'')}")
    doc.add_heading("Árboles", level=1)
    total = 0
    for i, tree in enumerate(data.get("trees", []), 1):
        vol = volume_for_tree(tree); total += vol
        doc.add_paragraph(f"{i}. {tree.get('nombre_comun','')} — {tree.get('nombre_cientifico','')} — DAP {tree.get('dap_cm','')} cm — Altura {tree.get('altura_m','')} m — Volumen {vol:.3f} m³")
    doc.add_paragraph(f"Volumen total estimado: {total:.3f} m³")
    doc.add_heading("Requisitos", level=1)
    for req in requirements_for_data(data):
        uploaded = bool(data.get("uploaded_docs", {}).get(req["id"]))
        doc.add_paragraph(f"[{ 'SUBIDO' if uploaded else 'PENDIENTE' }] {req['name']}")
    stream = io.BytesIO(); doc.save(stream); stream.seek(0); return stream.getvalue()

# =============================================================================
# ROUTER
# =============================================================================
page = st.session_state.get("page", "Inicio")
if page == "Inicio": page_inicio()
elif page == "Nueva solicitud": page_nueva_solicitud()
elif page == "Requisitos": page_requisitos()
elif page == "Autoridades": page_autoridades()
elif page == "Documentación": page_documentacion()
elif page == "Seguimiento": page_seguimiento()
elif page == "Administración": page_admin()
elif page == "Encuesta de validación": page_encuesta()
elif page == "Manual de uso": page_manual()

# Persistir cambios de cliente al finalizar cada ejecución.
save_case()
