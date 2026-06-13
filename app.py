import io
import hashlib
import json
import re
import secrets as py_secrets
from datetime import datetime, date, timedelta
from pathlib import Path

import pandas as pd
import requests
import streamlit as st

try:
    from supabase import create_client
except Exception:
    create_client = None

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A3, landscape
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
except Exception:
    colors = None
    landscape = None
    A3 = None
    getSampleStyleSheet = None
    SimpleDocTemplate = None
    Table = None
    TableStyle = None
    Paragraph = None
    Spacer = None
    Image = None


# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================
APP_VERSION = "DCD 1.1.1"
APP_TITLE = "DATOS CAPACIDAD DOCENTE (DCD 1.0)"
DEFAULT_PASSWORD = "Capacidad2026"
EXCEL_PATH = Path(__file__).parent / "data" / "listado_para_capacidad_docente.xlsx"
ASSETS_DIR = Path(__file__).parent / "assets"
LOGO_PATH = ASSETS_DIR / "logo.png"
INSTITUTIONAL_PHRASE = "Informe desarrollado para la gestión y análisis de los Datos de Capacidad Docente del Servicio correspondiente."

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="📊",
    layout="wide",
)


# =========================================================
# DATOS MAESTROS: ÁREA / UNIDAD DOCENTE
# Reutiliza la estructura del aplicativo de residentes y añade mapeo a columnas del Excel.
# =========================================================
AREA_OPTIONS = [
    "HOSPITAL",
    "ATENCION FAMILIAR Y COMUNITARIA",
]

DIRECCIONES_POR_AREA = {
    "HOSPITAL": [
        "DIRECCIÓN GERENCIA HOSPITAL DOCTOR NEGRIN",
        "DIRECCIÓN GERENCIA COMPLEJO HOSPITALARIO UNIVERSITARIO Y MATERNO INFANTIL",
        "DIRECCIÓN GERENCIA COMPLEJO HOSPITALARIO UNIVERSITARIO DE CANARIAS",
        "DIRECCIÓN GERENCIA HOSPITAL NUESTRA SEÑORA DE CANDELARIA",
    ],
    "ATENCION FAMILIAR Y COMUNITARIA": [
        "GERENCIA DE ATENCIÓN PRIMARIA DE GRAN CANARIA",
        "GERENCIA DE ATENCIÓN PRIMARIA DE TENERIFE NORTE",
        "GERENCIA DE ATENCIÓN PRIMARIA DE TENERIFE SUR",
        "GERENCIA DE SERVICIOS SANITARIOS DE FUERTEVENTURA",
        "GERENCIA DE SERVICIOS SANITARIOS DE LANZAROTE",
        "GERENCIA DE SERVICIOS SANITARIOS DE LA PALMA",
        "GERENCIA DE SERVICIOS SANITARIOS DE LA GOMERA",
        "GERENCIA DE SERVICIOS SANITARIOS DE EL HIERRO",
    ],
}

LEGACY_AREA_MAP = {
    "UNIDAD DOCENTE DE CENTRO HOSPITALARIO": "HOSPITAL",
    "UNIDAD DOCENTE MULTIPROFESIONAL DE ATENCION FAMILIAR Y COMUNITARIA": "ATENCION FAMILIAR Y COMUNITARIA",
}

CODIGOS_DIRECCION = {
    "DIRECCIÓN GERENCIA HOSPITAL DOCTOR NEGRIN": "HUGCNEGRIN",
    "DIRECCIÓN GERENCIA COMPLEJO HOSPITALARIO UNIVERSITARIO Y MATERNO INFANTIL": "CHUIMI",
    "DIRECCIÓN GERENCIA COMPLEJO HOSPITALARIO UNIVERSITARIO DE CANARIAS": "CHUC",
    "DIRECCIÓN GERENCIA HOSPITAL NUESTRA SEÑORA DE CANDELARIA": "HUNSC",
    "GERENCIA DE ATENCIÓN PRIMARIA DE GRAN CANARIA": "GAPGC",
    "GERENCIA DE ATENCIÓN PRIMARIA DE TENERIFE NORTE": "GAPTF_NORTE",
    "GERENCIA DE ATENCIÓN PRIMARIA DE TENERIFE SUR": "GAPTF_SUR",
    "GERENCIA DE SERVICIOS SANITARIOS DE FUERTEVENTURA": "GSSFV",
    "GERENCIA DE SERVICIOS SANITARIOS DE LANZAROTE": "GSSLZ",
    "GERENCIA DE SERVICIOS SANITARIOS DE LA PALMA": "GSSLP",
    "GERENCIA DE SERVICIOS SANITARIOS DE LA GOMERA": "GSSLG",
    "GERENCIA DE SERVICIOS SANITARIOS DE EL HIERRO": "GSSEH",
    "UNIDAD DOCENTE MULTIPROFESIONAL DE SALUD MENTAL DE GRAN CANARIA": "UDM_SM_GC",
    "UNIDAD DOCENTE MULTIPROFESIONAL DE SALUF MENTAL DE TENERIFE": "UDM_SM_TF",
    "UNIDAD DOCENTE MULTIPROFESIONAL DE SALUD LABORAL": "UDM_SL",
    "UNIDAD DOCENTE MULTIPROFESIONAL DE OBSTETRICIA Y GINECOLOGIA DEL CHUIMI": "UDM_OG_CHUIMI",
    "UNIDAD DOCENTE MULTIPROFESIONAL DE PEDIATRIA DEL CHUIMI": "UDM_PED_CHUIMI",
    "UNIDAD DOCENTE MULTIPROFESIONAL DE PEDIATRIA DEL HUNSC": "UDM_PED_HUNSC",
    "UNIDAD DOCENTE MULTIPROFESIONAL DE PEDIATRIA DEL HUC": "UDM_PED_HUC",
    "UNIDAD DOCENTE DE ENFERMERIA OBSTETRICA-GINECOLOGICA DE TENERIFE": "UD_ENFOG_TF",
    "UNIDAD DOCENTE DE MEDICINA PREVENTIVA Y SALUD PUBLICA": "UD_MPYSP",
}

# Mapeo de unidad docente a columna existente en el Excel.
# Algunas unidades de Tenerife comparten columna en el Excel original.
COLUMNA_EXCEL_POR_DIRECCION = {
    "DIRECCIÓN GERENCIA HOSPITAL DOCTOR NEGRIN": "HUGC DN",
    "DIRECCIÓN GERENCIA COMPLEJO HOSPITALARIO UNIVERSITARIO Y MATERNO INFANTIL": "CHUIMI",
    "DIRECCIÓN GERENCIA COMPLEJO HOSPITALARIO UNIVERSITARIO DE CANARIAS": "CHUC",
    "DIRECCIÓN GERENCIA HOSPITAL NUESTRA SEÑORA DE CANDELARIA": "HUNSC",
    "GERENCIA DE ATENCIÓN PRIMARIA DE GRAN CANARIA": "GAP GC",
    "GERENCIA DE ATENCIÓN PRIMARIA DE TENERIFE NORTE": "GAP TF",
    "GERENCIA DE ATENCIÓN PRIMARIA DE TENERIFE SUR": "GAP TF",
    "GERENCIA DE SERVICIOS SANITARIOS DE FUERTEVENTURA": "GSS FV",
    "GERENCIA DE SERVICIOS SANITARIOS DE LANZAROTE": "GSS LZ",
    "GERENCIA DE SERVICIOS SANITARIOS DE LA PALMA": "GSS LP",
    "GERENCIA DE SERVICIOS SANITARIOS DE LA GOMERA": "GSS LG",
    "GERENCIA DE SERVICIOS SANITARIOS DE EL HIERRO": "GSS EH",
}

KEY_COLUMNS = ["Nivel Estudio I", "Nivel Estudio II", "Rama", "Titulación"]
MATRIX_VALUE_COLUMNS = [
    "CHUIMI",
    "HUGC DN",
    "GAP GC",
    "Gran Canaria",
    "GSS FV",
    "GSS LZ",
    "Las Palmas",
    "CHUC",
    "HUNSC",
    "GAP TF",
    "Tenerife",
    "GSS LP",
    "GSS LG",
    "GSS EH",
    "S/C Tenerife",
    "Total",
]
DERIVED_MATRIX_COLUMNS = {"Gran Canaria", "Las Palmas", "Tenerife", "S/C Tenerife", "Total"}

# Columnas reales introducidas por centros. Las columnas agregadas se calculan.
REAL_CENTER_COLUMNS = [
    "CHUIMI",
    "HUGC DN",
    "GAP GC",
    "GSS FV",
    "GSS LZ",
    "CHUC",
    "HUNSC",
    "GAP TF",
    "GSS LP",
    "GSS LG",
    "GSS EH",
]

ISLA_COLUMNAS = {
    "Gran Canaria": ["CHUIMI", "HUGC DN", "GAP GC"],
    "Fuerteventura": ["GSS FV"],
    "Lanzarote": ["GSS LZ"],
    "Tenerife": ["GAP TF", "CHUC", "HUNSC"],
    "La Palma": ["GSS LP"],
    "La Gomera": ["GSS LG"],
    "El Hierro": ["GSS EH"],
}

PROVINCIA_COLUMNAS = {
    "Las Palmas": ["CHUIMI", "HUGC DN", "GAP GC", "GSS FV", "GSS LZ"],
    "S/C Tenerife": ["GAP TF", "CHUC", "HUNSC", "GSS LP", "GSS LG", "GSS EH"],
}

PUBLICATION_BUCKET = "dcd-publicaciones"


# =========================================================
# UTILIDADES DE ESTADO
# =========================================================
def init_session_state() -> None:
    defaults = {
        "logged_in": False,
        "current_user": "",
        "current_user_role": "",
        "current_user_display": "",
        "current_user_area": "",
        "current_user_unidad": "",
        "current_user_codigo_unidad": "",
        "must_change_password": False,
        "current_step": 1,
        "info_understood": False,
        "area_selected": "",
        "direccion_selected": "",
        "confirm_selection": False,
        "registros": {},
        "sel_nivel_i": "",
        "sel_nivel_ii": "",
        "sel_rama": "",
        "sel_titulacion": "",
        "numero_alumnos": 0,
        "codigo_borrador": "",
        "codigo_expediente": "",
        "version_num": 0,
        "draft_estado": "borrador",
        "permitir_editar_finalizado": False,
        "observaciones": "",
        "last_message": "",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def reset_selectores_estudio() -> None:
    st.session_state.sel_nivel_i = ""
    st.session_state.sel_nivel_ii = ""
    st.session_state.sel_rama = ""
    st.session_state.sel_titulacion = ""
    st.session_state.numero_alumnos = 0


def normalizar_area(area: str) -> str:
    return LEGACY_AREA_MAP.get(area, area)


def reset_downstream(level: str) -> None:
    if level == "nivel_i":
        st.session_state.sel_nivel_ii = ""
        st.session_state.sel_rama = ""
        st.session_state.sel_titulacion = ""
    elif level == "nivel_ii":
        st.session_state.sel_rama = ""
        st.session_state.sel_titulacion = ""
    elif level == "rama":
        st.session_state.sel_titulacion = ""


def safe_code(text: str) -> str:
    text = text or ""
    text = text.upper().strip()
    text = re.sub(r"[^A-Z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text[:60] or "DCD"


def build_codigo_borrador(unidad_docente: str) -> str:
    codigo_unidad = CODIGOS_DIRECCION.get(unidad_docente) or safe_code(unidad_docente)
    return f"DCD-{codigo_unidad}-2026"


def build_codigo_version(codigo_expediente: str, version_num: int) -> str:
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    return f"{codigo_expediente}-V{version_num:03d}-{timestamp}"


def expected_centros_docentes() -> list[dict]:
    rows = []
    for area, centros in DIRECCIONES_POR_AREA.items():
        for centro in centros:
            rows.append({
                "Área": area,
                "Centro docente": centro,
                "Código unidad": CODIGOS_DIRECCION.get(centro, safe_code(centro)),
                "Columna Excel": COLUMNA_EXCEL_POR_DIRECCION.get(centro, ""),
            })
    return rows


def registro_key(nivel_i: str, nivel_ii: str, rama: str, titulacion: str) -> str:
    return "||".join([nivel_i, nivel_ii, rama, titulacion])


# =========================================================
# LECTURA DEL EXCEL BASE
# =========================================================
@st.cache_data(show_spinner=False)
def load_catalogo() -> pd.DataFrame:
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"No se encuentra el Excel base en: {EXCEL_PATH}")

    # El Excel tiene una primera fila superior de agrupación y la cabecera real en la fila 2.
    df = pd.read_excel(EXCEL_PATH, header=1)
    df = df.dropna(how="all")

    missing = [col for col in KEY_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(f"Faltan columnas obligatorias en el Excel: {missing}")

    for col in KEY_COLUMNS:
        df[col] = df[col].astype(str).str.strip()

    df = df[df["Titulación"].notna() & (df["Titulación"].astype(str).str.strip() != "")]
    df = df.drop_duplicates(subset=KEY_COLUMNS).reset_index(drop=True)
    return df


def sorted_unique(series: pd.Series) -> list[str]:
    values = [str(x).strip() for x in series.dropna().unique() if str(x).strip()]
    return sorted(values, key=lambda x: x.upper())


# =========================================================
# SUPABASE
# =========================================================
def get_secret(name: str, default: str = "") -> str:
    try:
        return st.secrets.get(name, default)
    except Exception:
        return default


def get_users_config() -> dict:
    """
    Control básico/avanzado de usuarios.

    Por defecto existe un usuario admin con la contraseña APP_PASSWORD.
    Si se define USERS_JSON en secrets, se usan los usuarios indicados allí.
    Ejemplo de USERS_JSON:
    {"admin":{"password":"Capacidad2026","role":"admin","display_name":"Administrador"}}
    """
    default_password = get_secret("APP_PASSWORD", DEFAULT_PASSWORD)
    default_users = {
        "admin": {
            "password": default_password,
            "role": "admin",
            "display_name": "Administrador",
        }
    }

    raw = get_secret("USERS_JSON", "")
    if not raw:
        return default_users

    try:
        parsed = json.loads(raw)
        if isinstance(parsed, dict) and parsed:
            return parsed
    except Exception:
        pass

    return default_users


def is_admin() -> bool:
    return st.session_state.get("current_user_role") == "admin"


def is_external_viewer() -> bool:
    """Usuarios de consulta externa: solo pueden ver la publicación vigente."""
    return st.session_state.get("current_user_role") in {"consulta", "externo", "entidad_externa"}


def user_scope_unidad() -> str:
    return st.session_state.get("current_user_unidad", "")


def hash_password(password: str, salt: str | None = None, iterations: int = 200000) -> str:
    salt = salt or py_secrets.token_hex(16)
    password_bytes = (password or "").encode("utf-8")
    derived = hashlib.pbkdf2_hmac("sha256", password_bytes, salt.encode("utf-8"), iterations)
    return f"pbkdf2_sha256${iterations}${salt}${derived.hex()}"


def verify_password(password: str, stored_hash: str) -> bool:
    try:
        algorithm, iterations, salt, expected = stored_hash.split("$", 3)
        if algorithm != "pbkdf2_sha256":
            return False
        candidate = hash_password(password, salt=salt, iterations=int(iterations)).split("$", 3)[3]
        return py_secrets.compare_digest(candidate, expected)
    except Exception:
        return False


def get_user_from_supabase(username: str) -> dict | None:
    client = get_supabase_client()
    if client is None:
        return None

    try:
        resp = client.table("dcd_usuarios").select("*").eq("username", username).limit(1).execute()
        rows = getattr(resp, "data", []) or []
        return rows[0] if rows else None
    except Exception:
        return None


def login_user(username: str, password: str) -> tuple[bool, str]:
    username = (username or "").strip()
    password = password or ""

    db_user = get_user_from_supabase(username)
    if db_user:
        if not db_user.get("activo", True):
            return False, "Usuario desactivado."
        if not verify_password(password, db_user.get("password_hash", "")):
            return False, "Usuario o contraseña incorrectos."

        st.session_state.logged_in = True
        st.session_state.current_user = username
        st.session_state.current_user_role = str(db_user.get("role", "usuario"))
        st.session_state.current_user_display = str(db_user.get("display_name", username))
        st.session_state.current_user_area = normalizar_area(str(db_user.get("area", "") or ""))
        st.session_state.current_user_unidad = str(db_user.get("unidad_docente", "") or "")
        st.session_state.current_user_codigo_unidad = str(db_user.get("codigo_unidad", "") or "")
        st.session_state.must_change_password = bool(db_user.get("must_change_password", False))
        return True, "Acceso correcto."

    users = get_users_config()
    if username not in users:
        return False, "Usuario o contraseña incorrectos."

    user_data = users.get(username, {})
    expected_password = str(user_data.get("password", ""))
    if password != expected_password:
        return False, "Usuario o contraseña incorrectos."

    st.session_state.logged_in = True
    st.session_state.current_user = username
    st.session_state.current_user_role = str(user_data.get("role", "usuario"))
    st.session_state.current_user_display = str(user_data.get("display_name", username))
    st.session_state.current_user_area = normalizar_area(str(user_data.get("area", "") or ""))
    st.session_state.current_user_unidad = str(user_data.get("unidad_docente", "") or "")
    st.session_state.current_user_codigo_unidad = str(user_data.get("codigo_unidad", "") or "")
    st.session_state.must_change_password = False
    return True, "Acceso correcto."


def audit_event(action: str, detail: str = "", codigo_borrador: str = "") -> None:
    """
    Registra una acción en Supabase si la tabla dcd_auditoria existe.
    La auditoría no debe bloquear el uso del aplicativo si falla.
    """
    client = get_supabase_client()
    if client is None:
        return

    try:
        client.table("dcd_auditoria").insert({
            "created_at": datetime.now().isoformat(),
            "usuario": st.session_state.get("current_user", ""),
            "rol": st.session_state.get("current_user_role", ""),
            "accion": action,
            "detalle": detail,
            "codigo_borrador": codigo_borrador or st.session_state.get("codigo_borrador", ""),
            "area": st.session_state.get("area_selected", ""),
            "unidad_docente": st.session_state.get("direccion_selected", ""),
        }).execute()
    except Exception:
        return


def mailgun_configured() -> bool:
    return all([
        get_secret("MAILGUN_API_KEY", ""),
        get_secret("MAILGUN_DOMAIN", ""),
        get_secret("MAILGUN_SENDER_EMAIL", ""),
        get_secret("MAILGUN_RECIPIENT_EMAIL", ""),
    ])


def send_email_with_mailgun(excel_bytes: bytes, filename: str) -> tuple[bool, str]:
    api_key = get_secret("MAILGUN_API_KEY", "")
    domain = get_secret("MAILGUN_DOMAIN", "")
    sender = get_secret("MAILGUN_SENDER_EMAIL", "")
    recipient = get_secret("MAILGUN_RECIPIENT_EMAIL", "")

    if not all([api_key, domain, sender, recipient]):
        return False, "Mailgun no está configurado en secrets. Puede descargar el Excel y enviarlo manualmente."

    codigo = st.session_state.codigo_borrador or build_codigo_borrador(st.session_state.direccion_selected)
    subject = f"DCD 1.0 - Datos Capacidad Docente - {codigo}"
    text = (
        "Se adjunta el Excel generado por el aplicativo DATOS CAPACIDAD DOCENTE (DCD 1.0).\n\n"
        f"Código de borrador: {codigo}\n"
        f"Área: {st.session_state.area_selected}\n"
        f"Centro docente: {st.session_state.direccion_selected}\n"
        f"Usuario: {st.session_state.get('current_user_display', '')}\n"
        f"Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n"
    )

    try:
        response = requests.post(
            f"https://api.mailgun.net/v3/{domain}/messages",
            auth=("api", api_key),
            data={
                "from": sender,
                "to": recipient,
                "subject": subject,
                "text": text,
            },
            files={
                "attachment": (
                    filename,
                    excel_bytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            },
            timeout=30,
        )
        if response.status_code == 200:
            audit_event("correo_enviado", f"Correo enviado a {recipient}", codigo)
            return True, f"Correo enviado correctamente a {recipient}."
        return False, f"Error Mailgun {response.status_code}: {response.text}"
    except Exception as exc:
        return False, f"Error al enviar correo: {exc}"


@st.cache_resource(show_spinner=False)
def get_supabase_client():
    url = get_secret("SUPABASE_URL", "")
    key = get_secret("SUPABASE_KEY", "")
    if not url or not key or create_client is None:
        return None
    return create_client(url, key)


def supabase_available() -> bool:
    return get_supabase_client() is not None


def registros_to_rows(estado: str = "borrador") -> list[dict]:
    rows = []
    area = st.session_state.area_selected
    unidad = st.session_state.direccion_selected
    codigo_unidad = CODIGOS_DIRECCION.get(unidad, safe_code(unidad))
    columna_excel = COLUMNA_EXCEL_POR_DIRECCION.get(unidad, "")
    codigo_borrador = st.session_state.codigo_borrador or build_codigo_borrador(unidad)
    codigo_expediente = st.session_state.codigo_expediente or build_codigo_borrador(unidad)
    version_num = int(st.session_state.get("version_num") or 1)

    for item in st.session_state.registros.values():
        rows.append({
            "codigo_borrador": codigo_borrador,
            "codigo_expediente": codigo_expediente,
            "version_num": version_num,
            "estado": estado,
            "area": area,
            "unidad_docente": unidad,
            "codigo_unidad": codigo_unidad,
            "columna_excel": columna_excel,
            "usuario_propietario": st.session_state.get("current_user", ""),
            "nivel_i": item["Nivel Estudio I"],
            "nivel_ii": item["Nivel Estudio II"],
            "rama": item["Rama"],
            "titulacion": item["Titulación"],
            "numero_alumnos": int(item["Nº alumnos"]),
        })
    return rows


def get_next_version_num(codigo_expediente: str) -> int:
    client = get_supabase_client()
    if client is None:
        return 1

    try:
        resp = client.table("dcd_borradores").select("version_num").eq("codigo_expediente", codigo_expediente).order("version_num", desc=True).limit(1).execute()
        rows = getattr(resp, "data", []) or []
        if rows and rows[0].get("version_num") is not None:
            return int(rows[0].get("version_num") or 0) + 1

        # Compatibilidad con borradores anteriores a la 1.0.6.
        legacy_resp = client.table("dcd_borradores").select("codigo_borrador").eq("codigo_borrador", codigo_expediente).limit(1).execute()
        legacy_rows = getattr(legacy_resp, "data", []) or []
        return 2 if legacy_rows else 1
    except Exception:
        return 1


def save_draft_to_supabase(estado: str = "borrador") -> tuple[bool, str]:
    client = get_supabase_client()
    if client is None:
        return False, "Supabase no está configurado todavía. Puedes seguir usando el MVP local y descargar el Excel."

    unidad = st.session_state.direccion_selected
    if not is_admin() and user_scope_unidad() and unidad != user_scope_unidad():
        return False, "No puede guardar datos de un centro docente distinto al asignado a su usuario."

    codigo_expediente = st.session_state.codigo_expediente or build_codigo_borrador(unidad)
    version_num = get_next_version_num(codigo_expediente)
    codigo_borrador = build_codigo_version(codigo_expediente, version_num)
    saved_at = datetime.now().isoformat()
    st.session_state.codigo_expediente = codigo_expediente
    st.session_state.codigo_borrador = codigo_borrador
    st.session_state.version_num = version_num
    codigo_unidad = CODIGOS_DIRECCION.get(unidad, safe_code(unidad))

    try:
        client.table("dcd_borradores").update({"is_latest": False}).eq("codigo_expediente", codigo_expediente).execute()

        client.table("dcd_borradores").insert({
            "codigo_borrador": codigo_borrador,
            "codigo_expediente": codigo_expediente,
            "version_num": version_num,
            "saved_at": saved_at,
            "is_latest": True,
            "app_version": APP_VERSION,
            "estado": estado,
            "area": st.session_state.area_selected,
            "unidad_docente": unidad,
            "codigo_unidad": codigo_unidad,
            "observaciones": st.session_state.get("observaciones", ""),
            "usuario_propietario": st.session_state.get("current_user", ""),
            "usuario_ultima_edicion": st.session_state.get("current_user", ""),
        }).execute()

        rows = registros_to_rows(estado=estado)
        if rows:
            client.table("dcd_registros").insert(rows).execute()

        st.session_state.draft_estado = estado
        if estado == "finalizado":
            st.session_state.permitir_editar_finalizado = False

        audit_event("guardar_borrador" if estado == "borrador" else "guardar_finalizado", f"Estado: {estado}. Versión: {version_num}. Registros: {len(rows)}", codigo_borrador)
        if estado == "finalizado":
            send_admin_notification(
                subject=f"DCD - Centro docente finalizado: {unidad}",
                text=(
                    f"Un centro docente ha finalizado una nueva versión de su expediente DCD.\n\n"
                    f"Centro docente: {unidad}\n"
                    f"Código: {codigo_borrador}\n"
                    f"Versión: {version_num}\n"
                    f"Usuario: {st.session_state.get('current_user_display', '') or st.session_state.get('current_user', '')}\n"
                    f"Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n"
                ),
            )
            auto_msg = maybe_auto_publish_if_complete()
            if auto_msg:
                return True, f"Expediente finalizado correctamente como nueva versión: {codigo_borrador}. {auto_msg}"
            return True, f"Expediente finalizado correctamente como nueva versión: {codigo_borrador}"
        return True, f"Borrador guardado correctamente como nueva versión: {codigo_borrador}"
    except Exception as exc:
        return False, f"Error al guardar en Supabase: {exc}"


def load_draft_from_supabase(codigo_borrador: str) -> tuple[bool, str]:
    client = get_supabase_client()
    if client is None:
        return False, "Supabase no está configurado."

    try:
        borrador_resp = client.table("dcd_borradores").select("*").eq("codigo_borrador", codigo_borrador).limit(1).execute()
        borradores = getattr(borrador_resp, "data", []) or []
        if not borradores:
            return False, "No se encontró ese código de borrador."

        borrador = borradores[0]
        if not is_admin() and user_scope_unidad() and borrador.get("unidad_docente", "") != user_scope_unidad():
            return False, "No puede cargar un borrador de un centro docente distinto al asignado a su usuario."

        registros_resp = client.table("dcd_registros").select("*").eq("codigo_borrador", codigo_borrador).execute()
        rows = getattr(registros_resp, "data", []) or []

        st.session_state.area_selected = normalizar_area(borrador.get("area", ""))
        st.session_state.direccion_selected = borrador.get("unidad_docente", "")
        st.session_state.codigo_borrador = codigo_borrador
        st.session_state.codigo_expediente = borrador.get("codigo_expediente") or build_codigo_borrador(st.session_state.direccion_selected)
        st.session_state.version_num = int(borrador.get("version_num") or 1)
        st.session_state.draft_estado = borrador.get("estado", "borrador") or "borrador"
        st.session_state.permitir_editar_finalizado = False
        st.session_state.observaciones = borrador.get("observaciones", "") or ""
        st.session_state.registros = {}

        for row in rows:
            key = registro_key(row["nivel_i"], row["nivel_ii"], row["rama"], row["titulacion"])
            st.session_state.registros[key] = {
                "Nivel Estudio I": row["nivel_i"],
                "Nivel Estudio II": row["nivel_ii"],
                "Rama": row["rama"],
                "Titulación": row["titulacion"],
                "Nº alumnos": int(row.get("numero_alumnos") or 0),
            }

        audit_event("cargar_borrador", f"Registros cargados: {len(rows)}", codigo_borrador)
        return True, f"Borrador cargado: {codigo_borrador}"
    except Exception as exc:
        return False, f"Error al cargar desde Supabase: {exc}"


def list_drafts_from_supabase() -> list[str]:
    client = get_supabase_client()
    if client is None:
        return []
    try:
        query = client.table("dcd_borradores").select("codigo_borrador, codigo_expediente, version_num, saved_at, updated_at, estado, unidad_docente, is_latest")
        if not is_admin() and user_scope_unidad():
            query = query.eq("unidad_docente", user_scope_unidad())
        resp = query.order("saved_at", desc=True).order("updated_at", desc=True).limit(100).execute()
        rows = getattr(resp, "data", []) or []
        result = []
        for r in rows:
            version = r.get("version_num") or "-"
            date_value = r.get("saved_at") or r.get("updated_at") or ""
            latest = " | ÚLTIMO" if r.get("is_latest") else ""
            result.append(f"{r['codigo_borrador']} | v{version} | {r.get('estado', '')} | {date_value}{latest}")
        return result
    except Exception:
        return []


def list_users_from_supabase() -> list[dict]:
    client = get_supabase_client()
    if client is None:
        return []
    try:
        resp = client.table("dcd_usuarios").select(
            "username, display_name, role, area, unidad_docente, codigo_unidad, activo, must_change_password, updated_at"
        ).order("username").execute()
        return getattr(resp, "data", []) or []
    except Exception:
        return []


def save_user_to_supabase(
    username: str,
    display_name: str,
    role: str,
    area: str,
    unidad_docente: str,
    password: str,
    activo: bool = True,
    must_change_password: bool = True,
) -> tuple[bool, str]:
    client = get_supabase_client()
    if client is None:
        return False, "Supabase no está configurado."

    username = (username or "").strip().lower()
    if not username:
        return False, "Debe indicar un nombre de usuario."
    if not password:
        return False, "Debe indicar una contraseña temporal."
    if role == "usuario" and not unidad_docente:
        return False, "Los usuarios de centro docente deben tener un centro docente asignado."

    codigo_unidad = CODIGOS_DIRECCION.get(unidad_docente, safe_code(unidad_docente)) if unidad_docente else ""

    try:
        client.table("dcd_usuarios").upsert({
            "username": username,
            "display_name": display_name or username,
            "role": role,
            "area": normalizar_area(area or ""),
            "unidad_docente": unidad_docente or "",
            "codigo_unidad": codigo_unidad,
            "password_hash": hash_password(password),
            "activo": activo,
            "must_change_password": must_change_password,
        }, on_conflict="username").execute()
        audit_event("guardar_usuario", f"Usuario: {username} | Rol: {role}")
        return True, f"Usuario guardado correctamente: {username}"
    except Exception as exc:
        return False, f"Error al guardar usuario: {exc}"


def reset_user_password(username: str, new_password: str, must_change_password: bool = True) -> tuple[bool, str]:
    client = get_supabase_client()
    if client is None:
        return False, "Supabase no está configurado."
    username = (username or "").strip().lower()
    if not username or not new_password:
        return False, "Debe indicar usuario y nueva contraseña."
    try:
        client.table("dcd_usuarios").update({
            "password_hash": hash_password(new_password),
            "must_change_password": must_change_password,
        }).eq("username", username).execute()
        audit_event("reset_password", f"Usuario: {username}")
        return True, f"Contraseña reseteada para: {username}"
    except Exception as exc:
        return False, f"Error al resetear contraseña: {exc}"


def set_user_active(username: str, active: bool) -> tuple[bool, str]:
    client = get_supabase_client()
    if client is None:
        return False, "Supabase no está configurado."
    username = (username or "").strip().lower()
    if not username:
        return False, "Debe indicar un usuario."
    try:
        client.table("dcd_usuarios").update({"activo": active}).eq("username", username).execute()
        audit_event("activar_usuario" if active else "desactivar_usuario", f"Usuario: {username}")
        return True, f"Usuario {'activado' if active else 'desactivado'}: {username}"
    except Exception as exc:
        return False, f"Error al actualizar usuario: {exc}"


def delete_user_from_supabase(username: str) -> tuple[bool, str]:
    """Elimina físicamente un usuario. Usar solo si no se necesita conservarlo en el listado operativo."""
    client = get_supabase_client()
    if client is None:
        return False, "Supabase no está configurado."
    username = (username or "").strip().lower()
    if not username:
        return False, "Debe indicar un usuario."
    if username == (st.session_state.get("current_user") or "").strip().lower():
        return False, "No puede eliminar el usuario con el que está trabajando ahora mismo."
    try:
        client.table("dcd_usuarios").delete().eq("username", username).execute()
        audit_event("eliminar_usuario", f"Usuario eliminado: {username}")
        return True, f"Usuario eliminado correctamente: {username}"
    except Exception as exc:
        return False, f"Error al eliminar usuario: {exc}"


def change_current_user_password(new_password: str, repeat_password: str) -> tuple[bool, str]:
    if not new_password or len(new_password) < 8:
        return False, "La nueva contraseña debe tener al menos 8 caracteres."
    if new_password != repeat_password:
        return False, "Las contraseñas no coinciden."
    ok, msg = reset_user_password(st.session_state.get("current_user", ""), new_password, must_change_password=False)
    if ok:
        st.session_state.must_change_password = False
    return ok, msg


# =========================================================
# EXPORTACIÓN A EXCEL
# =========================================================
def formula_total_columna(col_letter: str, first_excel_row: int = 2, last_excel_row: int = 196) -> str:
    return f"=SUM({col_letter}{first_excel_row}:{col_letter}{last_excel_row})"


def aplicar_formulas_matriz(ws, workbook, data_row_count: int) -> None:
    """
    Aplica las fórmulas de la matriz:
    H = E+F+G, K = H+I+J, O = L+M+N, S = O+P+Q+R, T = K+S.
    La fila de total queda en la fila Excel 197 cuando la matriz tiene 195 filas de datos.
    """
    formula_format = workbook.add_format({"border": 1, "num_format": "0", "valign": "top"})
    total_label_format = workbook.add_format({
        "bold": True,
        "bg_color": "#D9EAF7",
        "border": 1,
        "align": "right",
        "valign": "top",
    })
    total_format = workbook.add_format({
        "bold": True,
        "bg_color": "#D9EAF7",
        "border": 1,
        "num_format": "0",
        "valign": "top",
    })

    # Filas de datos: Excel 2 a Excel 196 si hay 195 registros en la matriz base.
    for row_idx in range(1, data_row_count + 1):
        excel_row = row_idx + 1
        ws.write_formula(row_idx, 7, f"=SUM(E{excel_row}:G{excel_row})", formula_format)
        ws.write_formula(row_idx, 10, f"=SUM(H{excel_row}:J{excel_row})", formula_format)
        ws.write_formula(row_idx, 14, f"=SUM(L{excel_row}:N{excel_row})", formula_format)
        ws.write_formula(row_idx, 18, f"=SUM(O{excel_row}:R{excel_row})", formula_format)
        ws.write_formula(row_idx, 19, f"=K{excel_row}+S{excel_row}", formula_format)

    total_row_idx = data_row_count + 1
    last_excel_row = data_row_count + 1
    ws.write(total_row_idx, 3, "TOTAL", total_label_format)
    for col_idx, col_letter in enumerate("EFGHIJKLMNOPQRST", start=4):
        ws.write_formula(total_row_idx, col_idx, formula_total_columna(col_letter, last_excel_row=last_excel_row), total_format)


def limpiar_columnas_matriz(matriz: pd.DataFrame) -> pd.DataFrame:
    matriz = matriz.copy()
    for col in MATRIX_VALUE_COLUMNS:
        if col in matriz.columns:
            matriz[col] = pd.NA
    return matriz


def volcar_registro_en_matriz(matriz: pd.DataFrame, item: dict, columna_excel: str) -> None:
    if not columna_excel or columna_excel not in matriz.columns or columna_excel in DERIVED_MATRIX_COLUMNS:
        return

    mask = (
        (matriz["Nivel Estudio I"] == item["Nivel Estudio I"])
        & (matriz["Nivel Estudio II"] == item["Nivel Estudio II"])
        & (matriz["Rama"] == item["Rama"])
        & (matriz["Titulación"] == item["Titulación"])
    )
    if not mask.any():
        return

    current = pd.to_numeric(matriz.loc[mask, columna_excel], errors="coerce").fillna(0)
    matriz.loc[mask, columna_excel] = current + int(item["Nº alumnos"])


def build_output_excel() -> bytes:
    catalogo = load_catalogo()
    matriz = limpiar_columnas_matriz(catalogo)
    registros = list(st.session_state.registros.values())
    registros_df = pd.DataFrame(registros)

    unidad = st.session_state.direccion_selected
    columna_excel = COLUMNA_EXCEL_POR_DIRECCION.get(unidad, "")
    codigo = st.session_state.codigo_borrador or build_codigo_borrador(unidad)
    estado = st.session_state.get("draft_estado", "borrador")
    usuario = st.session_state.get("current_user_display", "") or st.session_state.get("current_user", "")

    for item in registros:
        volcar_registro_en_matriz(matriz, item, columna_excel)

    resumen = pd.DataFrame([{
        "Aplicativo": APP_VERSION,
        "Fecha generación": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Área": st.session_state.area_selected,
        "Centro docente": unidad,
        "Código unidad": CODIGOS_DIRECCION.get(unidad, safe_code(unidad)),
        "Columna Excel": columna_excel,
        "Código borrador": codigo,
        "Estado": estado,
        "Usuario": usuario,
        "Observaciones": st.session_state.get("observaciones", ""),
        "Nº registros": len(registros),
        "Total alumnos": sum(int(x["Nº alumnos"]) for x in registros) if registros else 0,
    }])

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        if registros_df.empty:
            registros_df = pd.DataFrame(columns=KEY_COLUMNS + ["Nº alumnos"])
        else:
            registros_df.insert(0, "Código borrador", codigo)
            registros_df.insert(1, "Estado", estado)
            registros_df.insert(2, "Área", st.session_state.area_selected)
            registros_df.insert(3, "Centro docente", unidad)
            registros_df.insert(4, "Usuario", usuario)
        registros_df.to_excel(writer, sheet_name="Registros_DCD", index=False)
        matriz.to_excel(writer, sheet_name="Matriz_DCD", index=False)

        workbook = writer.book
        header_format = workbook.add_format({
            "bold": True,
            "bg_color": "#1F4E78",
            "font_color": "#FFFFFF",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        })
        body_format = workbook.add_format({"border": 1, "valign": "top"})
        int_format = workbook.add_format({"border": 1, "num_format": "0", "valign": "top"})

        for sheet_name in ["Resumen", "Registros_DCD", "Matriz_DCD"]:
            ws = writer.sheets[sheet_name]
            df_sheet = {"Resumen": resumen, "Registros_DCD": registros_df, "Matriz_DCD": matriz}[sheet_name]
            for col_num, value in enumerate(df_sheet.columns.values):
                ws.write(0, col_num, value, header_format)
                max_len = max([len(str(value))] + [len(str(v)) for v in df_sheet[value].head(200).fillna("").tolist()])
                ws.set_column(col_num, col_num, min(max(max_len + 2, 12), 45))
            ws.freeze_panes(1, 0)
            if not df_sheet.empty:
                rows, cols = df_sheet.shape
                ws.autofilter(0, 0, rows, max(cols - 1, 0))
                # Formato básico para zona usada.
                for row_idx in range(1, min(rows + 1, 500)):
                    for col_idx, col_name in enumerate(df_sheet.columns):
                        value = df_sheet.iloc[row_idx - 1, col_idx]
                        fmt = int_format if isinstance(value, int) else body_format
                        if pd.isna(value):
                            value = ""
                        ws.write(row_idx, col_idx, value, fmt)

            if sheet_name == "Matriz_DCD":
                aplicar_formulas_matriz(ws, workbook, len(matriz))

    output.seek(0)
    return output.getvalue()


def get_latest_finalized_drafts() -> tuple[list[dict], list[dict]]:
    client = get_supabase_client()
    if client is None:
        return [], []

    resp = client.table("dcd_borradores").select("*").eq("estado", "finalizado").order("saved_at", desc=True).order("updated_at", desc=True).execute()
    borradores = getattr(resp, "data", []) or []

    latest_by_unit = {}
    duplicates = []
    for borrador in borradores:
        unidad = borrador.get("unidad_docente", "")
        if not unidad:
            continue
        if unidad not in latest_by_unit:
            latest_by_unit[unidad] = borrador
        else:
            duplicates.append(borrador)

    return list(latest_by_unit.values()), duplicates


def get_admin_centros_status() -> tuple[pd.DataFrame, list[str]]:
    client = get_supabase_client()
    expected = expected_centros_docentes()
    if client is None:
        return pd.DataFrame(expected), []

    try:
        resp = client.table("dcd_borradores").select("*").order("saved_at", desc=True).order("updated_at", desc=True).execute()
        borradores = getattr(resp, "data", []) or []
    except Exception:
        borradores = []

    by_unit: dict[str, list[dict]] = {}
    for borrador in borradores:
        unidad = borrador.get("unidad_docente", "")
        if unidad:
            by_unit.setdefault(unidad, []).append(borrador)

    rows = []
    missing = []
    for item in expected:
        centro = item["Centro docente"]
        versions = by_unit.get(centro, [])
        latest_any = versions[0] if versions else {}
        finalized = [v for v in versions if v.get("estado") == "finalizado"]
        latest_finalized = finalized[0] if finalized else {}

        if latest_finalized:
            if latest_any and latest_any.get("codigo_borrador") != latest_finalized.get("codigo_borrador"):
                estado = "Finalizado con borrador posterior"
            else:
                estado = "Finalizado"
            entra = "Sí"
        elif latest_any:
            estado = "Pendiente de finalizar"
            entra = "No"
            missing.append(centro)
        else:
            estado = "Sin datos"
            entra = "No"
            missing.append(centro)

        rows.append({
            **item,
            "Estado": estado,
            "Última versión guardada": latest_any.get("codigo_borrador", ""),
            "Fecha último guardado": latest_any.get("saved_at") or latest_any.get("updated_at", ""),
            "Último finalizado": latest_finalized.get("codigo_borrador", ""),
            "Fecha último finalizado": latest_finalized.get("saved_at") or latest_finalized.get("updated_at", ""),
            "Entra en consolidado": entra,
            "Versiones guardadas": len(versions),
        })

    return pd.DataFrame(rows), missing


def send_missing_centros_email(missing: list[str], status_df: pd.DataFrame) -> tuple[bool, str]:
    api_key = get_secret("MAILGUN_API_KEY", "")
    domain = get_secret("MAILGUN_DOMAIN", "")
    sender = get_secret("MAILGUN_SENDER_EMAIL", "")
    recipient = get_secret("MAILGUN_RECIPIENT_EMAIL", "")

    if not all([api_key, domain, sender, recipient]):
        return False, "Mailgun no está configurado en secrets. Puede revisar el listado de pendientes en pantalla."

    if not missing:
        return False, "No hay centros pendientes que notificar."

    pending_lines = "\n".join(f"- {centro}" for centro in missing)
    subject = "DCD - Centros docentes pendientes de finalizar"
    text = (
        "Centros docentes pendientes de incorporarse al consolidado DCD:\n\n"
        f"{pending_lines}\n\n"
        f"Fecha del aviso: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n"
        f"Usuario administrador: {st.session_state.get('current_user_display', '')}\n"
    )

    try:
        response = requests.post(
            f"https://api.mailgun.net/v3/{domain}/messages",
            auth=("api", api_key),
            data={
                "from": sender,
                "to": recipient,
                "subject": subject,
                "text": text,
            },
            timeout=30,
        )
        if response.status_code == 200:
            audit_event("aviso_centros_pendientes", f"Pendientes: {len(missing)}")
            return True, f"Aviso enviado correctamente a {recipient}."
        return False, f"Error Mailgun {response.status_code}: {response.text}"
    except Exception as exc:
        return False, f"Error al enviar aviso: {exc}"



def calcular_totales_matriz_df(matriz: pd.DataFrame) -> pd.DataFrame:
    """Devuelve una copia de Matriz_DCD con columnas derivadas calculadas y fila TOTAL."""
    out = matriz.copy()
    for col in MATRIX_VALUE_COLUMNS:
        if col in out.columns:
            out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0).astype(int)

    if all(c in out.columns for c in ["CHUIMI", "HUGC DN", "GAP GC", "Gran Canaria"]):
        out["Gran Canaria"] = out["CHUIMI"] + out["HUGC DN"] + out["GAP GC"]
    if all(c in out.columns for c in ["Gran Canaria", "GSS FV", "GSS LZ", "Las Palmas"]):
        out["Las Palmas"] = out["Gran Canaria"] + out["GSS FV"] + out["GSS LZ"]
    if all(c in out.columns for c in ["CHUC", "HUNSC", "GAP TF", "Tenerife"]):
        out["Tenerife"] = out["CHUC"] + out["HUNSC"] + out["GAP TF"]
    if all(c in out.columns for c in ["Tenerife", "GSS LP", "GSS LG", "GSS EH", "S/C Tenerife"]):
        out["S/C Tenerife"] = out["Tenerife"] + out["GSS LP"] + out["GSS LG"] + out["GSS EH"]
    if all(c in out.columns for c in ["Las Palmas", "S/C Tenerife", "Total"]):
        out["Total"] = out["Las Palmas"] + out["S/C Tenerife"]

    total_row = {col: "" for col in out.columns}
    if "Titulación" in out.columns:
        total_row["Titulación"] = "TOTAL"
    elif len(out.columns) >= 4:
        total_row[out.columns[3]] = "TOTAL"
    for col in MATRIX_VALUE_COLUMNS:
        if col in out.columns:
            total_row[col] = int(pd.to_numeric(out[col], errors="coerce").fillna(0).sum())
    out = pd.concat([out, pd.DataFrame([total_row])], ignore_index=True)
    return out



def matriz_sin_total(matriz: pd.DataFrame) -> pd.DataFrame:
    """Devuelve la matriz calculada sin la fila TOTAL final."""
    df = calcular_totales_matriz_df(matriz)
    if "Titulación" in df.columns:
        df = df[df["Titulación"].astype(str).str.upper().str.strip() != "TOTAL"].copy()
    return df.reset_index(drop=True)


def suma_columnas(df: pd.DataFrame, columnas: list[str]) -> pd.Series:
    existentes = [c for c in columnas if c in df.columns]
    if not existentes:
        return pd.Series([0] * len(df), index=df.index)
    return df[existentes].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)


def total_columna(df: pd.DataFrame, columna: str) -> int:
    if columna not in df.columns:
        return 0
    return int(pd.to_numeric(df[columna], errors="coerce").fillna(0).sum())


def build_analytics_tables(matriz: pd.DataFrame, status_df: pd.DataFrame | None = None) -> dict[str, pd.DataFrame]:
    """Construye tablas analíticas a partir de la Matriz_DCD consolidada."""
    df = matriz_sin_total(matriz)
    for col in MATRIX_VALUE_COLUMNS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    total = total_columna(df, "Total")
    las_palmas = suma_columnas(df, PROVINCIA_COLUMNAS["Las Palmas"]).sum()
    tenerife_prov = suma_columnas(df, PROVINCIA_COLUMNAS["S/C Tenerife"]).sum()
    centros_finalizados = 0
    centros_pendientes = 0
    if status_df is not None and not status_df.empty and "Estado" in status_df.columns:
        centros_finalizados = int(status_df["Estado"].astype(str).str.startswith("Finalizado").sum())
        centros_pendientes = int((~status_df["Estado"].astype(str).str.startswith("Finalizado")).sum())

    resumen_global = pd.DataFrame([
        {"Indicador": "Total plazas", "Valor": int(total)},
        {"Indicador": "Total provincia Las Palmas", "Valor": int(las_palmas)},
        {"Indicador": "Total provincia S/C Tenerife", "Valor": int(tenerife_prov)},
        {"Indicador": "% Las Palmas sobre total", "Valor": round((las_palmas / total * 100), 2) if total else 0},
        {"Indicador": "% S/C Tenerife sobre total", "Valor": round((tenerife_prov / total * 100), 2) if total else 0},
        {"Indicador": "Centros finalizados", "Valor": centros_finalizados},
        {"Indicador": "Centros pendientes/sin finalizar", "Valor": centros_pendientes},
    ])

    resumen_provincia = pd.DataFrame([
        {"Provincia": provincia, "Total plazas": int(suma_columnas(df, columnas).sum()), "% sobre total": round((suma_columnas(df, columnas).sum() / total * 100), 2) if total else 0}
        for provincia, columnas in PROVINCIA_COLUMNAS.items()
    ])

    resumen_isla = pd.DataFrame([
        {"Isla": isla, "Total plazas": int(suma_columnas(df, columnas).sum()), "% sobre total": round((suma_columnas(df, columnas).sum() / total * 100), 2) if total else 0}
        for isla, columnas in ISLA_COLUMNAS.items()
    ]).sort_values("Total plazas", ascending=False).reset_index(drop=True)

    resumen_centro = pd.DataFrame([
        {"Centro/columna": col, "Total plazas": total_columna(df, col), "% sobre total": round((total_columna(df, col) / total * 100), 2) if total else 0}
        for col in REAL_CENTER_COLUMNS if col in df.columns
    ]).sort_values("Total plazas", ascending=False).reset_index(drop=True)

    resumen_rama = (
        df.groupby("Rama", dropna=False)["Total"].sum().reset_index(name="Total plazas")
        if "Rama" in df.columns and "Total" in df.columns else pd.DataFrame(columns=["Rama", "Total plazas"])
    )
    if not resumen_rama.empty:
        resumen_rama["% sobre total"] = resumen_rama["Total plazas"].apply(lambda x: round((x / total * 100), 2) if total else 0)
        resumen_rama = resumen_rama.sort_values("Total plazas", ascending=False).reset_index(drop=True)

    resumen_nivel = (
        df.groupby(["Nivel Estudio I", "Nivel Estudio II"], dropna=False)["Total"].sum().reset_index(name="Total plazas")
        if all(c in df.columns for c in ["Nivel Estudio I", "Nivel Estudio II", "Total"]) else pd.DataFrame(columns=["Nivel Estudio I", "Nivel Estudio II", "Total plazas"])
    )
    if not resumen_nivel.empty:
        resumen_nivel["% sobre total"] = resumen_nivel["Total plazas"].apply(lambda x: round((x / total * 100), 2) if total else 0)
        resumen_nivel = resumen_nivel.sort_values("Total plazas", ascending=False).reset_index(drop=True)

    centro_rama_rows = []
    for col in REAL_CENTER_COLUMNS:
        if col not in df.columns:
            continue
        tmp = df.groupby("Rama", dropna=False)[col].sum().reset_index(name="Total plazas")
        tmp.insert(0, "Centro/columna", col)
        centro_rama_rows.append(tmp)
    resumen_centro_rama = pd.concat(centro_rama_rows, ignore_index=True) if centro_rama_rows else pd.DataFrame(columns=["Centro/columna", "Rama", "Total plazas"])
    if not resumen_centro_rama.empty:
        resumen_centro_rama = resumen_centro_rama[resumen_centro_rama["Total plazas"] > 0].sort_values(["Centro/columna", "Total plazas"], ascending=[True, False]).reset_index(drop=True)

    centro_nivel_rows = []
    for col in REAL_CENTER_COLUMNS:
        if col not in df.columns:
            continue
        tmp = df.groupby(["Nivel Estudio I", "Nivel Estudio II"], dropna=False)[col].sum().reset_index(name="Total plazas")
        tmp.insert(0, "Centro/columna", col)
        centro_nivel_rows.append(tmp)
    resumen_centro_nivel = pd.concat(centro_nivel_rows, ignore_index=True) if centro_nivel_rows else pd.DataFrame(columns=["Centro/columna", "Nivel Estudio I", "Nivel Estudio II", "Total plazas"])
    if not resumen_centro_nivel.empty:
        resumen_centro_nivel = resumen_centro_nivel[resumen_centro_nivel["Total plazas"] > 0].sort_values(["Centro/columna", "Total plazas"], ascending=[True, False]).reset_index(drop=True)

    top_titulaciones = pd.DataFrame(columns=["Nivel Estudio I", "Nivel Estudio II", "Rama", "Titulación", "Total plazas"])
    if all(c in df.columns for c in KEY_COLUMNS + ["Total"]):
        top_titulaciones = (
            df.groupby(KEY_COLUMNS, dropna=False)["Total"].sum().reset_index(name="Total plazas")
            .sort_values("Total plazas", ascending=False)
            .reset_index(drop=True)
        )
        top_titulaciones = top_titulaciones[top_titulaciones["Total plazas"] > 0].head(25)

    return {
        "Dashboard": resumen_global,
        "Resumen_Global": resumen_global,
        "Resumen_Provincia": resumen_provincia,
        "Resumen_Isla": resumen_isla,
        "Resumen_Centro": resumen_centro,
        "Resumen_Rama": resumen_rama,
        "Resumen_Nivel": resumen_nivel,
        "Resumen_Centro_Rama": resumen_centro_rama,
        "Resumen_Centro_Nivel": resumen_centro_nivel,
        "Top_Titulaciones": top_titulaciones,
    }



def build_current_analytics_from_supabase() -> tuple[bool, str, dict[str, pd.DataFrame] | None]:
    """Genera analítica rápida con la última versión finalizada por centro."""
    client = get_supabase_client()
    if client is None:
        return False, "Supabase no está configurado.", None
    try:
        borradores, _duplicados = get_latest_finalized_drafts()
        if not borradores:
            return False, "No hay expedientes finalizados para analizar.", None
        status_df, _missing = get_admin_centros_status()
        matriz = limpiar_columnas_matriz(load_catalogo())
        for borrador in borradores:
            codigo = borrador.get("codigo_borrador", "")
            unidad = borrador.get("unidad_docente", "")
            columna_excel = COLUMNA_EXCEL_POR_DIRECCION.get(unidad, "")
            if not codigo:
                continue
            resp = client.table("dcd_registros").select("*").eq("codigo_borrador", codigo).execute()
            for row in getattr(resp, "data", []) or []:
                item = {
                    "Nivel Estudio I": row.get("nivel_i", ""),
                    "Nivel Estudio II": row.get("nivel_ii", ""),
                    "Rama": row.get("rama", ""),
                    "Titulación": row.get("titulacion", ""),
                    "Nº alumnos": int(row.get("numero_alumnos") or 0),
                }
                volcar_registro_en_matriz(matriz, item, columna_excel)
        return True, "Analítica generada.", build_analytics_tables(matriz, status_df)
    except Exception as exc:
        return False, f"Error al generar analítica: {exc}", None


def render_streamlit_dashboard(analytics: dict[str, pd.DataFrame]) -> None:
    global_df = analytics.get("Resumen_Global", pd.DataFrame())
    values = dict(zip(global_df.get("Indicador", []), global_df.get("Valor", []))) if not global_df.empty else {}
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total plazas", values.get("Total plazas", 0))
    c2.metric("Las Palmas", values.get("Total provincia Las Palmas", 0))
    c3.metric("S/C Tenerife", values.get("Total provincia S/C Tenerife", 0))
    c4.metric("Centros pendientes", values.get("Centros pendientes/sin finalizar", 0))

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("#### Por isla")
        isla = analytics.get("Resumen_Isla", pd.DataFrame())
        if not isla.empty:
            st.bar_chart(isla.set_index("Isla")["Total plazas"])
            st.dataframe(isla, use_container_width=True, hide_index=True)
    with col_b:
        st.markdown("#### Por rama")
        rama = analytics.get("Resumen_Rama", pd.DataFrame())
        if not rama.empty:
            st.bar_chart(rama.head(10).set_index("Rama")["Total plazas"])
            st.dataframe(rama.head(10), use_container_width=True, hide_index=True)

    st.markdown("#### Top titulaciones")
    top = analytics.get("Top_Titulaciones", pd.DataFrame())
    if not top.empty:
        st.dataframe(top.head(15), use_container_width=True, hide_index=True)

def build_dashboard_text(analytics: dict[str, pd.DataFrame]) -> list[str]:
    global_df = analytics.get("Resumen_Global", pd.DataFrame())
    values = dict(zip(global_df.get("Indicador", []), global_df.get("Valor", []))) if not global_df.empty else {}
    return [
        f"Total de plazas: {values.get('Total plazas', 0)}",
        f"Provincia Las Palmas: {values.get('Total provincia Las Palmas', 0)}",
        f"Provincia S/C Tenerife: {values.get('Total provincia S/C Tenerife', 0)}",
        f"Centros finalizados: {values.get('Centros finalizados', 0)}",
        f"Centros pendientes/sin finalizar: {values.get('Centros pendientes/sin finalizar', 0)}",
    ]


def add_pdf_table(story: list, title: str, df: pd.DataFrame, styles, max_rows: int = 20) -> None:
    if df is None or df.empty:
        return
    story.append(Paragraph(title, styles["Heading2"]))
    show = df.head(max_rows).copy().fillna("")
    data = [list(show.columns)] + show.astype(str).values.tolist()
    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 7),
        ("FONTSIZE", (0, 1), (-1, -1), 6.5),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(table)
    story.append(Spacer(1, 8))

def generate_matriz_pdf(matriz: pd.DataFrame, titulo: str, resumen_lineas: list[str], analytics: dict[str, pd.DataFrame] | None = None) -> bytes:
    if SimpleDocTemplate is None or Table is None:
        raise RuntimeError("La librería reportlab no está instalada. Añada reportlab a requirements.txt y reinicie la app.")

    matriz_pdf = calcular_totales_matriz_df(matriz)
    output = io.BytesIO()
    doc = SimpleDocTemplate(
        output,
        pagesize=landscape(A3),
        leftMargin=18,
        rightMargin=18,
        topMargin=18,
        bottomMargin=18,
    )
    styles = getSampleStyleSheet()
    story = []
    if Image is not None and LOGO_PATH.exists():
        try:
            story.append(Image(str(LOGO_PATH), width=120, height=60, kind="proportional"))
            story.append(Spacer(1, 6))
        except Exception:
            pass
    story.append(Paragraph(titulo, styles["Title"]))
    story.append(Paragraph(INSTITUTIONAL_PHRASE, styles["Normal"]))
    story.append(Spacer(1, 6))
    for linea in resumen_lineas:
        story.append(Paragraph(str(linea), styles["Normal"]))
    story.append(Spacer(1, 8))

    if analytics:
        story.append(Paragraph("Dashboard resumen", styles["Heading1"]))
        for linea in build_dashboard_text(analytics):
            story.append(Paragraph(str(linea), styles["Normal"]))
        story.append(Spacer(1, 8))
        add_pdf_table(story, "Resumen por provincia", analytics.get("Resumen_Provincia", pd.DataFrame()), styles, max_rows=10)
        add_pdf_table(story, "Resumen por isla", analytics.get("Resumen_Isla", pd.DataFrame()), styles, max_rows=10)
        add_pdf_table(story, "Top centros docentes / columnas", analytics.get("Resumen_Centro", pd.DataFrame()), styles, max_rows=15)
        add_pdf_table(story, "Top ramas", analytics.get("Resumen_Rama", pd.DataFrame()), styles, max_rows=15)
        add_pdf_table(story, "Top titulaciones", analytics.get("Top_Titulaciones", pd.DataFrame()), styles, max_rows=15)
        story.append(Paragraph("Matriz DCD completa", styles["Heading1"]))

    df = matriz_pdf.copy().fillna("")
    # Reducimos textos largos para que la matriz sea imprimible en PDF.
    for col in ["Nivel Estudio I", "Nivel Estudio II", "Rama", "Titulación"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.slice(0, 70)

    data = [list(df.columns)] + df.astype(str).values.tolist()
    page_width = landscape(A3)[0] - 36
    # 4 columnas descriptivas + 16 columnas numéricas. Ajuste compacto.
    desc_widths = [52, 58, 80, 170]
    remaining = max(page_width - sum(desc_widths), 200)
    numeric_count = max(len(df.columns) - 4, 1)
    col_widths = desc_widths + [remaining / numeric_count] * numeric_count
    col_widths = col_widths[:len(df.columns)]

    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 5.5),
        ("FONTSIZE", (0, 1), (-1, -1), 4.5),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (4, 1), (-1, -1), "RIGHT"),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E2F0D9")),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
    ]))
    story.append(table)
    doc.build(story)
    output.seek(0)
    return output.getvalue()


def build_publication_package() -> tuple[bool, str, dict | None]:
    """Construye los objetos necesarios para una publicación, sin guardarlos todavía."""
    client = get_supabase_client()
    if client is None:
        return False, "Supabase no está configurado.", None

    borradores, duplicados = get_latest_finalized_drafts()
    if not borradores:
        return False, "No hay expedientes finalizados para publicar.", None

    status_df, missing_centros = get_admin_centros_status()
    catalogo = load_catalogo()
    matriz = limpiar_columnas_matriz(catalogo)
    registros_consolidados = []
    codigos_usados = []

    for borrador in borradores:
        codigo = borrador.get("codigo_borrador", "")
        unidad = borrador.get("unidad_docente", "")
        columna_excel = COLUMNA_EXCEL_POR_DIRECCION.get(unidad, "")
        if not codigo:
            continue
        codigos_usados.append(codigo)
        resp = client.table("dcd_registros").select("*").eq("codigo_borrador", codigo).execute()
        rows = getattr(resp, "data", []) or []
        for row in rows:
            item = {
                "Nivel Estudio I": row.get("nivel_i", ""),
                "Nivel Estudio II": row.get("nivel_ii", ""),
                "Rama": row.get("rama", ""),
                "Titulación": row.get("titulacion", ""),
                "Nº alumnos": int(row.get("numero_alumnos") or 0),
            }
            volcar_registro_en_matriz(matriz, item, columna_excel)
            registros_consolidados.append({
                "Código borrador": codigo,
                "Fecha guardado": borrador.get("saved_at") or borrador.get("updated_at", ""),
                "Área": normalizar_area(borrador.get("area", "")),
                "Centro docente": unidad,
                "Columna Excel": columna_excel,
                **item,
            })

    analytics = build_analytics_tables(matriz, status_df)

    ok_xlsx, msg_xlsx, excel_bytes, excel_filename = build_consolidated_excel_from_supabase()
    if not ok_xlsx or not excel_bytes:
        return False, msg_xlsx, None

    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    pdf_bytes = generate_matriz_pdf(
        matriz,
        titulo="DATOS CAPACIDAD DOCENTE (DCD 1.0) - Informe de publicación",
        resumen_lineas=[
            f"Fecha de generación: {fecha}",
            f"Expedientes finalizados incorporados: {len(codigos_usados)}",
            f"Centros docentes pendientes: {len(missing_centros)}",
            "Criterio: última versión finalizada por centro docente.",
        ],
        analytics=analytics,
    )

    return True, "Paquete de publicación generado.", {
        "excel_bytes": excel_bytes,
        "excel_filename": excel_filename,
        "pdf_bytes": pdf_bytes,
        "pdf_filename": f"Matriz_DCD_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
        "status_df": status_df,
        "missing_centros": missing_centros,
        "borradores": borradores,
        "duplicados": duplicados,
        "registros_consolidados": registros_consolidados,
        "analytics": analytics,
        "matriz": matriz,
    }


def get_next_publication_version() -> int:
    client = get_supabase_client()
    if client is None:
        return 1
    try:
        resp = client.table("dcd_publicaciones").select("version_publicacion").order("version_publicacion", desc=True).limit(1).execute()
        rows = getattr(resp, "data", []) or []
        if rows and rows[0].get("version_publicacion") is not None:
            return int(rows[0].get("version_publicacion") or 0) + 1
    except Exception:
        pass
    return 1


def ensure_publication_bucket(client) -> None:
    try:
        client.storage.create_bucket(PUBLICATION_BUCKET, options={"public": False})
    except Exception:
        # Si ya existe, Supabase devuelve error. No debe bloquear.
        pass


def upload_publication_file(client, path: str, data: bytes, content_type: str) -> tuple[bool, str]:
    try:
        ensure_publication_bucket(client)
        try:
            client.storage.from_(PUBLICATION_BUCKET).upload(path, data, file_options={"content-type": content_type, "upsert": "true"})
        except TypeError:
            client.storage.from_(PUBLICATION_BUCKET).upload(path, data, {"content-type": content_type, "upsert": "true"})
        return True, path
    except Exception as exc:
        return False, f"Error al subir archivo a Supabase Storage: {exc}"


def download_publication_file(path: str) -> tuple[bool, bytes | None, str]:
    client = get_supabase_client()
    if client is None:
        return False, None, "Supabase no está configurado."
    try:
        data = client.storage.from_(PUBLICATION_BUCKET).download(path)
        return True, data, "Archivo descargado."
    except Exception as exc:
        return False, None, f"Error al descargar archivo: {exc}"


def fetch_table_for_backup(table_name: str, limit: int = 10000) -> tuple[pd.DataFrame, str]:
    """Descarga una tabla de Supabase para backup administrativo.

    Devuelve DataFrame y mensaje de error si lo hubiera. No bloquea si una tabla no existe.
    """
    client = get_supabase_client()
    if client is None:
        return pd.DataFrame(), "Supabase no está configurado."

    try:
        query = client.table(table_name).select("*").limit(limit)
        if table_name in {"dcd_auditoria", "dcd_borradores", "dcd_registros", "dcd_publicaciones"}:
            query = query.order("created_at", desc=True)
        elif table_name == "dcd_usuarios":
            query = query.order("username")
        resp = query.execute()
        rows = getattr(resp, "data", []) or []
        return pd.DataFrame(rows), ""
    except Exception as exc:
        return pd.DataFrame(), str(exc)


def generate_admin_backup_excel(include_password_hashes: bool = False) -> tuple[bool, bytes | None, str]:
    """Genera un backup administrativo en Excel de las tablas principales.

    Por seguridad, los hashes de contraseña se excluyen por defecto.
    """
    if not supabase_available():
        return False, None, "Supabase no está configurado."

    tables = [
        "dcd_usuarios",
        "dcd_borradores",
        "dcd_registros",
        "dcd_publicaciones",
        "dcd_auditoria",
        "dcd_configuracion",
    ]
    errors = []
    output = io.BytesIO()

    try:
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            resumen_rows = []
            for table in tables:
                df, err = fetch_table_for_backup(table)
                if err:
                    errors.append({"tabla": table, "error": err})
                if table == "dcd_usuarios" and not df.empty and not include_password_hashes:
                    if "password_hash" in df.columns:
                        df = df.drop(columns=["password_hash"])
                    df["password_hash_exportado"] = "No. Excluido por seguridad."
                sheet = table.replace("dcd_", "")[:31]
                df.to_excel(writer, sheet_name=sheet, index=False)
                resumen_rows.append({
                    "tabla": table,
                    "filas_exportadas": len(df),
                    "estado": "OK" if not err else "ERROR",
                    "observacion": err,
                })

            pd.DataFrame(resumen_rows).to_excel(writer, sheet_name="Resumen_backup", index=False)
            if errors:
                pd.DataFrame(errors).to_excel(writer, sheet_name="Errores", index=False)

        output.seek(0)
        audit_event(
            "backup_admin_generado",
            f"Backup generado. Hashes incluidos: {'sí' if include_password_hashes else 'no'}",
        )
        return True, output.getvalue(), "Backup generado correctamente."
    except Exception as exc:
        return False, None, f"Error al generar backup: {exc}"


def send_admin_notification(subject: str, text: str) -> tuple[bool, str]:
    api_key = get_secret("MAILGUN_API_KEY", "")
    domain = get_secret("MAILGUN_DOMAIN", "")
    sender = get_secret("MAILGUN_SENDER_EMAIL", "")
    recipient = get_secret("MAILGUN_RECIPIENT_EMAIL", "")

    if not all([api_key, domain, sender, recipient]):
        return False, "Mailgun no está configurado; no se ha enviado correo de notificación."
    try:
        response = requests.post(
            f"https://api.mailgun.net/v3/{domain}/messages",
            auth=("api", api_key),
            data={"from": sender, "to": recipient, "subject": subject, "text": text},
            timeout=30,
        )
        if response.status_code == 200:
            return True, f"Notificación enviada a {recipient}."
        return False, f"Error Mailgun {response.status_code}: {response.text}"
    except Exception as exc:
        return False, f"Error al enviar notificación: {exc}"


# =========================================================
# CONFIGURACIÓN DE CIERRE / PUBLICACIÓN
# =========================================================
DEFAULT_CIERRE_CONFIG = {
    "modo_cierre": "todos_finalizados",
    "fecha_tope": "",
    "dias_aviso_previo": "7",
    "avisos_admin_activados": "true",
    "ultimo_aviso_previo": "",
    "ultima_publicacion_fecha_tope": "",
}

CIERRE_MODE_LABELS = {
    "todos_finalizados": "Publicar automáticamente solo cuando todos finalicen",
    "fecha_tope_con_pendientes": "Publicar automáticamente al llegar la fecha tope aunque falten centros",
    "manual": "Solo publicación manual por admin",
}


def get_config_value(key: str, default: str = "") -> str:
    client = get_supabase_client()
    if client is None:
        return default
    try:
        resp = client.table("dcd_configuracion").select("valor").eq("clave", key).limit(1).execute()
        rows = getattr(resp, "data", []) or []
        if rows:
            return str(rows[0].get("valor") or default)
    except Exception:
        return default
    return default


def set_config_value(key: str, value: str) -> tuple[bool, str]:
    client = get_supabase_client()
    if client is None:
        return False, "Supabase no está configurado."
    try:
        client.table("dcd_configuracion").upsert({
            "clave": key,
            "valor": str(value),
            "descripcion": "Configuración DCD",
            "updated_by": st.session_state.get("current_user", "sistema"),
        }, on_conflict="clave").execute()
        return True, "Configuración guardada."
    except Exception as exc:
        return False, f"Error al guardar configuración: {exc}"


def get_cierre_config() -> dict:
    cfg = DEFAULT_CIERRE_CONFIG.copy()
    for key, default in DEFAULT_CIERRE_CONFIG.items():
        cfg[key] = get_config_value(key, default)
    if cfg.get("modo_cierre") not in CIERRE_MODE_LABELS:
        cfg["modo_cierre"] = "todos_finalizados"
    try:
        cfg["dias_aviso_previo_int"] = max(0, int(cfg.get("dias_aviso_previo") or 7))
    except Exception:
        cfg["dias_aviso_previo_int"] = 7
    cfg["avisos_admin_activados_bool"] = str(cfg.get("avisos_admin_activados", "true")).lower() in {"1", "true", "sí", "si", "yes"}
    return cfg


def parse_config_date(value: str) -> date | None:
    value = (value or "").strip()
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except Exception:
        return None


def today_iso() -> str:
    return datetime.now().date().isoformat()


def maybe_send_deadline_warning() -> str:
    """Envía aviso previo al admin si la fecha tope está próxima y quedan centros pendientes.

    Nota: Streamlit no ejecuta tareas en segundo plano. Este aviso se evalúa cuando alguien usa la app.
    """
    cfg = get_cierre_config()
    if not cfg.get("avisos_admin_activados_bool"):
        return ""
    fecha_tope = parse_config_date(cfg.get("fecha_tope", ""))
    if not fecha_tope:
        return ""

    status_df, missing = get_admin_centros_status()
    if not missing:
        return ""

    hoy = datetime.now().date()
    dias = (fecha_tope - hoy).days
    if dias < 0 or dias > cfg.get("dias_aviso_previo_int", 7):
        return ""

    aviso_key = f"{hoy.isoformat()}|{fecha_tope.isoformat()}|{len(missing)}"
    if get_config_value("ultimo_aviso_previo", "") == aviso_key:
        return ""

    subject = f"DCD - Aviso previo de cierre: faltan {len(missing)} centros"
    text = (
        f"Aviso previo de cierre DCD.\n\n"
        f"Fecha tope configurada: {fecha_tope.strftime('%d/%m/%Y')}\n"
        f"Días restantes: {dias}\n"
        f"Modo de cierre: {CIERRE_MODE_LABELS.get(cfg.get('modo_cierre'), cfg.get('modo_cierre'))}\n\n"
        f"Centros pendientes:\n- " + "\n- ".join(missing) + "\n\n"
        f"Este aviso se genera automáticamente cuando la app detecta que la fecha tope está próxima."
    )
    ok, msg = send_admin_notification(subject, text)
    if ok:
        set_config_value("ultimo_aviso_previo", aviso_key)
        audit_event("aviso_previo_cierre", f"Fecha tope {fecha_tope.isoformat()}. Pendientes: {len(missing)}")
        return f"Aviso previo enviado al admin. Pendientes: {len(missing)}."
    return f"No se pudo enviar aviso previo: {msg}"


def maybe_deadline_auto_publish() -> str:
    """Publica automáticamente al llegar fecha tope si el modo lo permite.

    Nota: Streamlit no ejecuta cron real; se evalúa al usar la app.
    """
    cfg = get_cierre_config()
    if cfg.get("modo_cierre") != "fecha_tope_con_pendientes":
        return ""
    fecha_tope = parse_config_date(cfg.get("fecha_tope", ""))
    if not fecha_tope:
        return ""
    hoy = datetime.now().date()
    if hoy < fecha_tope:
        return ""
    if get_config_value("ultima_publicacion_fecha_tope", "") == fecha_tope.isoformat():
        return ""

    status_df, missing = get_admin_centros_status()
    tipo = "automatica_fecha_tope_con_pendientes" if missing else "automatica_fecha_tope_completa"
    motivo = (
        f"Publicación automática por fecha tope ({fecha_tope.strftime('%d/%m/%Y')}). "
        f"Centros pendientes en el momento de publicación: {len(missing)}."
    )
    ok, msg = create_publication(tipo_publicacion=tipo, motivo=motivo, allow_missing=True)
    if ok:
        set_config_value("ultima_publicacion_fecha_tope", fecha_tope.isoformat())
        audit_event("publicacion_fecha_tope", motivo)
        return f"Publicación automática por fecha tope generada: {msg}"

    send_admin_notification(
        "DCD - Error en publicación automática por fecha tope",
        f"Se alcanzó la fecha tope {fecha_tope.strftime('%d/%m/%Y')}, pero no se pudo generar la publicación automática.\n\nDetalle: {msg}",
    )
    return f"No se pudo generar publicación por fecha tope: {msg}"


def evaluar_cierre_automatico() -> str:
    mensajes = []
    msg_aviso = maybe_send_deadline_warning()
    if msg_aviso:
        mensajes.append(msg_aviso)
    msg_fecha = maybe_deadline_auto_publish()
    if msg_fecha:
        mensajes.append(msg_fecha)
    return " ".join(mensajes)


def create_publication(tipo_publicacion: str, motivo: str, allow_missing: bool) -> tuple[bool, str]:
    client = get_supabase_client()
    if client is None:
        return False, "Supabase no está configurado."

    ok, msg, package = build_publication_package()
    if not ok or not package:
        return False, msg

    missing = package["missing_centros"]
    if missing and not allow_missing:
        return False, "Hay centros docentes pendientes. Marque la opción de publicar con pendientes o espere a que finalicen."

    version = get_next_publication_version()
    codigo_publicacion = f"DCD-PUB-2026-V{version:03d}-{datetime.now().strftime('%Y%m%d-%H%M%S')}"
    base_path = f"publicaciones/{codigo_publicacion}"
    excel_path = f"{base_path}/{codigo_publicacion}.xlsx"
    pdf_path = f"{base_path}/{codigo_publicacion}_Matriz_DCD.pdf"

    ok_up_xlsx, msg_xlsx = upload_publication_file(
        client,
        excel_path,
        package["excel_bytes"],
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    if not ok_up_xlsx:
        return False, msg_xlsx
    ok_up_pdf, msg_pdf = upload_publication_file(client, pdf_path, package["pdf_bytes"], "application/pdf")
    if not ok_up_pdf:
        return False, msg_pdf

    try:
        client.table("dcd_publicaciones").update({"publicacion_vigente": False}).eq("publicacion_vigente", True).execute()
        centros_incluidos = [b.get("unidad_docente", "") for b in package["borradores"]]
        client.table("dcd_publicaciones").insert({
            "codigo_publicacion": codigo_publicacion,
            "version_publicacion": version,
            "fecha_publicacion": datetime.now().isoformat(),
            "publicacion_vigente": True,
            "tipo_publicacion": tipo_publicacion,
            "motivo_publicacion": motivo,
            "generada_por": st.session_state.get("current_user", "sistema"),
            "generada_automaticamente": tipo_publicacion == "automatica",
            "ruta_excel": excel_path,
            "ruta_pdf": pdf_path,
            "centros_incluidos": centros_incluidos,
            "centros_pendientes": missing,
            "centros_con_borrador_no_finalizado": package["status_df"].to_dict(orient="records"),
            "observaciones": "Publicación vigente generada desde DCD 1.0.8.2.",
        }).execute()
        audit_event("publicacion_generada", f"{codigo_publicacion}. Pendientes: {len(missing)}")
    except Exception as exc:
        return False, f"Los archivos se subieron, pero falló el registro de publicación: {exc}"

    subject = f"DCD - Nueva publicación vigente {codigo_publicacion}"
    text = (
        f"Se ha generado una nueva publicación vigente DCD.\n\n"
        f"Código: {codigo_publicacion}\n"
        f"Tipo: {tipo_publicacion}\n"
        f"Motivo: {motivo}\n"
        f"Centros incluidos: {len(package['borradores'])}\n"
        f"Centros pendientes: {len(missing)}\n"
        f"Publicación anterior: archivada como no vigente.\n"
        f"Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n"
    )
    send_admin_notification(subject, text)
    return True, f"Publicación vigente creada correctamente: {codigo_publicacion}"


def get_publicaciones() -> list[dict]:
    client = get_supabase_client()
    if client is None:
        return []
    try:
        resp = client.table("dcd_publicaciones").select("*").order("fecha_publicacion", desc=True).limit(100).execute()
        return getattr(resp, "data", []) or []
    except Exception:
        return []


def get_publicacion_vigente() -> dict | None:
    """Devuelve la publicación marcada como vigente. Solo hay una vigente por diseño."""
    client = get_supabase_client()
    if client is None:
        return None
    try:
        resp = (
            client.table("dcd_publicaciones")
            .select("*")
            .eq("publicacion_vigente", True)
            .order("fecha_publicacion", desc=True)
            .limit(1)
            .execute()
        )
        rows = getattr(resp, "data", []) or []
        return rows[0] if rows else None
    except Exception:
        pubs = get_publicaciones()
        return next((p for p in pubs if p.get("publicacion_vigente")), None)


def load_analytics_from_publication_excel(pub: dict) -> tuple[bool, str, dict[str, pd.DataFrame] | None]:
    """Lee las hojas analíticas del Excel publicado para que el portal muestre exactamente la publicación vigente."""
    ruta_excel = (pub or {}).get("ruta_excel", "")
    if not ruta_excel:
        return False, "La publicación vigente no tiene ruta de Excel asociada.", None

    ok, data, msg = download_publication_file(ruta_excel)
    if not ok or not data:
        return False, msg, None

    try:
        xls = pd.ExcelFile(io.BytesIO(data))
        sheets = [
            "Resumen_Global",
            "Resumen_Provincia",
            "Resumen_Isla",
            "Resumen_Centro",
            "Resumen_Rama",
            "Resumen_Nivel",
            "Resumen_Centro_Rama",
            "Resumen_Centro_Nivel",
            "Top_Titulaciones",
            "Estado_centros",
        ]
        analytics: dict[str, pd.DataFrame] = {}
        for sheet in sheets:
            if sheet in xls.sheet_names:
                analytics[sheet] = pd.read_excel(xls, sheet_name=sheet)
        if "Resumen_Global" not in analytics:
            return False, "El Excel publicado no contiene hojas analíticas. Genere una nueva publicación con la versión actual.", None
        return True, "Analítica de publicación cargada.", analytics
    except Exception as exc:
        return False, f"Error al leer el Excel publicado: {exc}", None


def render_info_card(label: str, value: str) -> None:
    """Tarjeta compacta para evitar cortes con puntos suspensivos en textos largos."""
    st.markdown(
        f"""
        <div class="dcd-card">
            <div class="dcd-card-label">{label}</div>
            <div class="dcd-card-value">{value or '-'}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_publication_metadata(pub: dict) -> None:
    col1, col2, col3 = st.columns(3)
    with col1:
        render_info_card("Publicación vigente", str(pub.get("codigo_publicacion", "")))
    with col2:
        render_info_card("Versión", str(pub.get("version_publicacion", "")))
    with col3:
        render_info_card("Fecha", str(pub.get("fecha_publicacion", ""))[:19])

    st.caption(f"Tipo: {pub.get('tipo_publicacion', '')} | Motivo: {pub.get('motivo_publicacion', '')}")


def render_publication_downloads(pub: dict, prefix: str = "portal") -> None:
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Preparar PDF", key=f"{prefix}_prepare_pdf"):
            ok, data, msg = download_publication_file(pub.get("ruta_pdf", ""))
            if ok and data:
                st.download_button(
                    "Descargar PDF de informe",
                    data=data,
                    file_name=f"{pub.get('codigo_publicacion', 'DCD')}_Matriz_DCD.pdf",
                    mime="application/pdf",
                    key=f"{prefix}_download_pdf",
                    on_click=audit_event,
                    args=("descarga_pdf_publicacion", f"{prefix} | {pub.get('codigo_publicacion', '')}"),
                )
            else:
                st.error(msg)
    with c2:
        if st.button("Preparar Excel", key=f"{prefix}_prepare_excel"):
            ok, data, msg = download_publication_file(pub.get("ruta_excel", ""))
            if ok and data:
                st.download_button(
                    "Descargar Excel consolidado",
                    data=data,
                    file_name=f"{pub.get('codigo_publicacion', 'DCD')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"{prefix}_download_excel",
                    on_click=audit_event,
                    args=("descarga_excel_publicacion", f"{prefix} | {pub.get('codigo_publicacion', '')}"),
                )
            else:
                st.error(msg)

def maybe_auto_publish_if_complete() -> str:
    """Publica automáticamente al finalizar un centro si la configuración lo permite.

    Nunca debe bloquear la finalización del centro: si falla, se registra/avisa y se devuelve un mensaje informativo.
    """
    try:
        cfg = get_cierre_config()
        modo = cfg.get("modo_cierre")
        if modo == "manual":
            return "Modo de cierre manual: no se genera publicación automática."

        _status_df, missing = get_admin_centros_status()
        if not missing:
            ok, msg = create_publication(
                tipo_publicacion="automatica",
                motivo="Publicación automática generada al estar todos los centros docentes finalizados.",
                allow_missing=False,
            )
            if ok:
                return f"Se ha generado publicación automática: {msg}"
            send_admin_notification(
                "DCD - Error al generar publicación automática",
                f"Todos los centros parecen finalizados, pero no se pudo generar la publicación automática.\n\nDetalle: {msg}",
            )
            return f"Todos los centros están finalizados, pero no se pudo generar la publicación automática: {msg}"

        if modo == "fecha_tope_con_pendientes":
            msg_fecha = maybe_deadline_auto_publish()
            if msg_fecha:
                return msg_fecha

        return f"Quedan {len(missing)} centros pendientes; no se genera publicación automática todavía."
    except Exception as exc:
        send_admin_notification(
            "DCD - Error inesperado en publicación automática",
            f"Se produjo un error inesperado al intentar publicar automáticamente.\n\nDetalle: {exc}",
        )
        return f"No se pudo comprobar/generar la publicación automática: {exc}"


def build_consolidated_excel_from_supabase() -> tuple[bool, str, bytes | None, str]:
    client = get_supabase_client()
    if client is None:
        return False, "Supabase no está configurado.", None, ""

    try:
        borradores, duplicados = get_latest_finalized_drafts()
        if not borradores:
            return False, "No hay expedientes finalizados para consolidar.", None, ""
        status_df, missing_centros = get_admin_centros_status()

        catalogo = load_catalogo()
        matriz = limpiar_columnas_matriz(catalogo)
        registros_consolidados = []
        codigos_usados = []

        for borrador in borradores:
            codigo = borrador.get("codigo_borrador", "")
            unidad = borrador.get("unidad_docente", "")
            columna_excel = COLUMNA_EXCEL_POR_DIRECCION.get(unidad, "")
            if not codigo:
                continue

            codigos_usados.append(codigo)
            resp = client.table("dcd_registros").select("*").eq("codigo_borrador", codigo).execute()
            rows = getattr(resp, "data", []) or []

            for row in rows:
                item = {
                    "Nivel Estudio I": row.get("nivel_i", ""),
                    "Nivel Estudio II": row.get("nivel_ii", ""),
                    "Rama": row.get("rama", ""),
                    "Titulación": row.get("titulacion", ""),
                    "Nº alumnos": int(row.get("numero_alumnos") or 0),
                }
                volcar_registro_en_matriz(matriz, item, columna_excel)
                registros_consolidados.append({
                    "Código borrador": codigo,
                    "Fecha guardado": borrador.get("saved_at") or borrador.get("updated_at", ""),
                    "Área": normalizar_area(borrador.get("area", "")),
                    "Centro docente": unidad,
                    "Columna Excel": columna_excel,
                    **item,
                })

        registros_df = pd.DataFrame(registros_consolidados)
        if registros_df.empty:
            registros_df = pd.DataFrame(columns=[
                "Código borrador",
                "Fecha guardado",
                "Área",
                "Centro docente",
                "Columna Excel",
                *KEY_COLUMNS,
                "Nº alumnos",
            ])

        resumen = pd.DataFrame([{
            "Aplicativo": APP_VERSION,
            "Fecha generación": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Tipo": "Consolidado finalizados",
            "Expedientes usados": len(codigos_usados),
            "Registros consolidados": len(registros_consolidados),
            "Centros docentes pendientes": len(missing_centros),
            "Duplicados finalizados ignorados": len(duplicados),
            "Criterio": "Último expediente finalizado por centro docente según fecha de guardado",
            "Usuario generación": st.session_state.get("current_user_display", "") or st.session_state.get("current_user", ""),
        }])

        borradores_df = pd.DataFrame([{
            "Código borrador": b.get("codigo_borrador", ""),
            "Estado": b.get("estado", ""),
            "Área": normalizar_area(b.get("area", "")),
            "Centro docente": b.get("unidad_docente", ""),
            "Código unidad": b.get("codigo_unidad", ""),
            "Actualizado": b.get("updated_at", ""),
            "Columna Excel": COLUMNA_EXCEL_POR_DIRECCION.get(b.get("unidad_docente", ""), ""),
        } for b in borradores])

        duplicados_df = pd.DataFrame([{
            "Código borrador ignorado": b.get("codigo_borrador", ""),
            "Área": normalizar_area(b.get("area", "")),
            "Centro docente": b.get("unidad_docente", ""),
            "Actualizado": b.get("updated_at", ""),
        } for b in duplicados])
        if duplicados_df.empty:
            duplicados_df = pd.DataFrame(columns=["Código borrador ignorado", "Área", "Centro docente", "Actualizado"])

        analytics = build_analytics_tables(matriz, status_df)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            sheets = {
                "Resumen": resumen,
                "Dashboard": analytics["Dashboard"],
                "Resumen_Global": analytics["Resumen_Global"],
                "Resumen_Provincia": analytics["Resumen_Provincia"],
                "Resumen_Isla": analytics["Resumen_Isla"],
                "Resumen_Centro": analytics["Resumen_Centro"],
                "Resumen_Rama": analytics["Resumen_Rama"],
                "Resumen_Nivel": analytics["Resumen_Nivel"],
                "Resumen_Centro_Rama": analytics["Resumen_Centro_Rama"],
                "Resumen_Centro_Nivel": analytics["Resumen_Centro_Nivel"],
                "Top_Titulaciones": analytics["Top_Titulaciones"],
                "Estado_centros": status_df,
                "Borradores_usados": borradores_df,
                "Registros_DCD": registros_df,
                "Duplicados_ignorados": duplicados_df,
                "Matriz_DCD": matriz,
            }

            workbook = writer.book
            header_format = workbook.add_format({
                "bold": True,
                "bg_color": "#1F4E78",
                "font_color": "#FFFFFF",
                "border": 1,
                "align": "center",
                "valign": "vcenter",
            })
            body_format = workbook.add_format({"border": 1, "valign": "top"})
            int_format = workbook.add_format({"border": 1, "num_format": "0", "valign": "top"})

            for sheet_name, df_sheet in sheets.items():
                df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]
                for col_num, value in enumerate(df_sheet.columns.values):
                    ws.write(0, col_num, value, header_format)
                    max_len = max([len(str(value))] + [len(str(v)) for v in df_sheet[value].head(200).fillna("").tolist()])
                    ws.set_column(col_num, col_num, min(max(max_len + 2, 12), 45))
                ws.freeze_panes(1, 0)
                if not df_sheet.empty:
                    rows, cols = df_sheet.shape
                    ws.autofilter(0, 0, rows, max(cols - 1, 0))
                    for row_idx in range(1, min(rows + 1, 500)):
                        for col_idx, col_name in enumerate(df_sheet.columns):
                            value = df_sheet.iloc[row_idx - 1, col_idx]
                            fmt = int_format if isinstance(value, int) else body_format
                            if pd.isna(value):
                                value = ""
                            ws.write(row_idx, col_idx, value, fmt)

                if sheet_name == "Matriz_DCD":
                    aplicar_formulas_matriz(ws, workbook, len(matriz))

            # Dashboard visual básico dentro del Excel: KPIs + gráficos de barras.
            try:
                dash_ws = writer.sheets.get("Dashboard")
                if dash_ws is not None:
                    title_fmt = workbook.add_format({"bold": True, "font_size": 14, "font_color": "#1F4E78"})
                    dash_ws.write("D1", "Dashboard DCD - publicación consolidada", title_fmt)
                    chart1 = workbook.add_chart({"type": "column"})
                    chart1.add_series({
                        "name": "Plazas por provincia",
                        "categories": "=Resumen_Provincia!$A$2:$A$3",
                        "values": "=Resumen_Provincia!$B$2:$B$3",
                    })
                    chart1.set_title({"name": "Plazas por provincia"})
                    chart1.set_y_axis({"name": "Plazas"})
                    dash_ws.insert_chart("D3", chart1, {"x_scale": 1.25, "y_scale": 1.15})

                    isla_rows = max(2, min(len(analytics["Resumen_Isla"]) + 1, 8))
                    chart2 = workbook.add_chart({"type": "bar"})
                    chart2.add_series({
                        "name": "Plazas por isla",
                        "categories": f"=Resumen_Isla!$A$2:$A${isla_rows}",
                        "values": f"=Resumen_Isla!$B$2:$B${isla_rows}",
                    })
                    chart2.set_title({"name": "Plazas por isla"})
                    chart2.set_x_axis({"name": "Plazas"})
                    dash_ws.insert_chart("D20", chart2, {"x_scale": 1.25, "y_scale": 1.4})

                    centro_rows = max(2, min(len(analytics["Resumen_Centro"]) + 1, 12))
                    chart3 = workbook.add_chart({"type": "bar"})
                    chart3.add_series({
                        "name": "Top centros/columnas",
                        "categories": f"=Resumen_Centro!$A$2:$A${centro_rows}",
                        "values": f"=Resumen_Centro!$B$2:$B${centro_rows}",
                    })
                    chart3.set_title({"name": "Top centros docentes / columnas"})
                    chart3.set_x_axis({"name": "Plazas"})
                    dash_ws.insert_chart("L3", chart3, {"x_scale": 1.35, "y_scale": 1.45})
            except Exception:
                # Los gráficos son un complemento. Si Excel no pudiera crearlos, no debe bloquear el consolidado.
                pass

        output.seek(0)
        filename = f"DCD_CONSOLIDADO_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        audit_event("generar_consolidado", f"Expedientes: {len(codigos_usados)}. Registros: {len(registros_consolidados)}")
        return True, f"Consolidado generado con {len(codigos_usados)} expedientes finalizados.", output.getvalue(), filename
    except Exception as exc:
        return False, f"Error al generar consolidado: {exc}", None, ""


# =========================================================
# COMPONENTES DE INTERFAZ
# =========================================================
def app_sidebar() -> None:
    st.sidebar.markdown("---")
    st.sidebar.markdown("### ℹ️ Información")
    st.sidebar.write(f"Versión: {APP_VERSION}")
    st.sidebar.write("Desarrollado para: F.S.E. – S.C.S.")
    if st.session_state.get("current_user_display"):
        st.sidebar.write(f"Usuario: {st.session_state.current_user_display}")
        st.sidebar.caption(f"Rol: {st.session_state.get('current_user_role', '')}")
        if st.session_state.get("current_user_unidad"):
            st.sidebar.caption(f"Centro asignado: {st.session_state.current_user_unidad}")

    if supabase_available():
        st.sidebar.success("Supabase configurado")
    else:
        st.sidebar.warning("Supabase no configurado")
        st.sidebar.caption("El aplicativo funciona como MVP local, pero no guarda borradores entre sesiones.")

    if mailgun_configured():
        st.sidebar.success("Mailgun configurado")
    else:
        st.sidebar.info("Mailgun no configurado")

    if st.session_state.get("current_user_role") == "admin":
        st.sidebar.markdown("---")
        if st.sidebar.button("Panel administrador"):
            st.session_state.current_step = 6
            st.rerun()
    elif is_external_viewer():
        st.sidebar.markdown("---")
        st.sidebar.info("Acceso de consulta: solo publicación vigente.")

    st.sidebar.markdown("---")
    if st.sidebar.button("Salir del aplicativo 🚪"):
        st.session_state.clear()
        st.rerun()


def page_login() -> None:
    st.title(APP_TITLE)
    st.subheader("🔐 Acceso al aplicativo")
    st.write("Introduce usuario y contraseña para continuar.")
    st.caption("Si no se configuran usuarios avanzados, el usuario por defecto es admin.")

    username = st.text_input("Usuario", value="admin", key="username_input")
    password = st.text_input("Contraseña", type="password", key="password_input")
    if st.button("Iniciar sesión"):
        ok, msg = login_user(username, password)
        if ok:
            audit_event("login", "Inicio de sesión correcto")
            st.session_state.current_step = 1
            st.rerun()
        else:
            st.error(msg)
            audit_event("login_fallido", f"Intento fallido. Usuario: {username}")

    st.markdown("---")
    st.markdown("##### Historial de versiones")
    st.markdown("- **DCD 1.0:** MVP inicial con contraseña, instrucciones, selección de centro docente, selectores dependientes y preparación para Supabase.")
    st.markdown("- **DCD 1.0.1:** Pantalla de recordatorio, correo automático opcional, usuarios configurables, auditoría y revisión de mapeo.")
    st.markdown("- **DCD 1.0.2:** Ajuste de áreas: Hospital, Atención Familiar y Comunitaria, y retirada de Otras Unidades Docentes.")
    st.markdown("- **DCD 1.0.3:** Cierre estable: exportación más completa, estado finalizado y bloqueo suave de edición.")
    st.markdown("- **DCD 1.0.4:** Totales en Matriz_DCD y panel administrador para Excel consolidado desde Supabase.")
    st.markdown("- **DCD 1.0.5:** Usuarios por centro docente, contraseñas hasheadas, reset por admin y borradores filtrados.")
    st.markdown("- **DCD 1.0.5.1:** Corrección del selector de centro docente al crear usuarios.")
    st.markdown("- **DCD 1.0.6:** Guardado versionado, control de centros pendientes y cambio visible a Centros Docentes.")
    st.markdown("- **DCD 1.0.7:** Publicaciones oficiales: PDF Matriz_DCD, histórico/vigente, Supabase Storage y notificación al administrador.")
    st.markdown("- **DCD 1.0.8:** Dashboard y análisis de publicación: resúmenes por provincia, isla, centro, rama, nivel y titulaciones en Excel/PDF/panel admin.")
    st.markdown("- **DCD 1.0.8.1:** Cierre configurable: fecha tope, modos de cierre, avisos previos y publicación automática por vencimiento si procede.")
    st.markdown("- **DCD 1.0.8.2:** Refuerzo de evaluación de cierre al acceder cualquier usuario, al finalizar centros y desde botón admin.")
    st.markdown("- **DCD 1.0.9:** Portal externo de consulta de publicación vigente con dashboard y descargas limitadas.")
    st.markdown("- **DCD 1.1.0:** Mejora visual del dashboard, tarjetas compactas y PDF preparado para logo/frase institucional.")
    st.markdown("- **DCD 1.1.1:** Auditoría ampliada, registro de accesos/descargas y backup completo admin.")
    st.markdown("- **DCD 1.0.9.1:** Ajustes de interfaz del portal y mantenimiento avanzado de usuarios.")


def page_change_password() -> None:
    st.title(APP_TITLE)
    st.subheader("Cambio obligatorio de contraseña")
    st.warning("Su usuario tiene una contraseña temporal. Debe cambiarla antes de continuar.")

    new_password = st.text_input("Nueva contraseña", type="password")
    repeat_password = st.text_input("Repetir nueva contraseña", type="password")

    if st.button("Guardar nueva contraseña"):
        ok, msg = change_current_user_password(new_password, repeat_password)
        if ok:
            st.success("Contraseña actualizada correctamente. Ya puede continuar.")
            st.rerun()
        else:
            st.error(msg)


def page_instrucciones() -> None:
    st.header("Paso 1: Información, instrucciones y aceptación")
    st.markdown(
        """
        **Bienvenido al aplicativo DATOS CAPACIDAD DOCENTE (DCD 1.0).**

        Este aplicativo tiene como finalidad recoger datos de capacidad docente por centro docente, nivel de estudios,
        rama y titulación.

        **Instrucciones básicas:**

        1. Seleccione primero el área y el centro docente correspondiente.
        2. Revise y confirme los datos de la unidad seleccionada.
        3. Introduzca los datos mediante los selectores encadenados:
           **Nivel Estudio I → Nivel Estudio II → Rama → Titulación**.
        4. Para cada titulación, indique el **número de alumnos**.
        5. Puede añadir varios registros antes de generar el Excel.
        6. Si Supabase está configurado, podrá guardar y recuperar borradores.
        7. Antes de finalizar, el aplicativo mostrará una pantalla específica de recordatorio y descarga/envío.

        **Advertencias:**

        - Revise bien la titulación seleccionada antes de añadirla.
        - Si introduce de nuevo una misma combinación, se actualizará el número de alumnos.
        - El envío automático por correo solo funcionará si Mailgun está configurado en los secretos de Streamlit.
        """
    )

    st.session_state.info_understood = st.checkbox(
        "He comprendido las instrucciones del aplicativo y deseo continuar",
        value=st.session_state.info_understood,
    )

    if st.button("CONTINUAR"):
        if st.session_state.info_understood:
            st.session_state.current_step = 2
            st.rerun()
        else:
            st.warning("Debe marcar la casilla de comprensión para continuar.")


def page_seleccion_unidad() -> None:
    st.header("Paso 2: Selección de Área y Centro Docente")

    st.session_state.area_selected = normalizar_area(st.session_state.area_selected)

    if not is_admin() and user_scope_unidad():
        st.session_state.area_selected = st.session_state.get("current_user_area", "")
        st.session_state.direccion_selected = user_scope_unidad()
        st.session_state.codigo_expediente = build_codigo_borrador(st.session_state.direccion_selected)
        st.info("Su usuario está vinculado a este centro docente. Solo podrá crear, cargar y modificar borradores de este centro.")
        st.markdown(f"**Área asignada:** {st.session_state.area_selected}")
        st.markdown(f"**Centro docente asignado:** {st.session_state.direccion_selected}")

        col_next, col_back = st.columns(2)
        with col_next:
            if st.button("Continuar con mi centro docente"):
                st.session_state.current_step = 3
                st.rerun()
        with col_back:
            if st.button("ATRÁS"):
                st.session_state.current_step = 1
                st.rerun()

        st.markdown("---")
        st.subheader("Cargar borrador existente")
        if supabase_available():
            drafts = list_drafts_from_supabase()
            if drafts:
                selected = st.selectbox("Borradores disponibles", options=[""] + drafts)
                codigo = selected.split(" | ")[0] if selected else ""
                if st.button("Cargar borrador"):
                    if codigo:
                        ok, msg = load_draft_from_supabase(codigo)
                        if ok:
                            st.success(msg)
                            st.session_state.current_step = 4
                            st.rerun()
                        else:
                            st.error(msg)
                    else:
                        st.warning("Debe seleccionar un código de borrador.")
            else:
                st.info("No se encontraron borradores guardados de su centro docente.")
        else:
            st.info("Supabase no está configurado. No se pueden cargar borradores guardados.")
        return

    st.session_state.area_selected = st.selectbox(
        "**SELECCIONE ÁREA**",
        options=[""] + AREA_OPTIONS,
        index=([""] + AREA_OPTIONS).index(st.session_state.area_selected) if st.session_state.area_selected in ([""] + AREA_OPTIONS) else 0,
    )

    direccion_options = DIRECCIONES_POR_AREA.get(st.session_state.area_selected, [])
    if st.session_state.direccion_selected not in direccion_options:
        st.session_state.direccion_selected = ""

    st.session_state.direccion_selected = st.selectbox(
        "**SELECCIONE CENTRO DOCENTE / DIRECCIÓN / GERENCIA**",
        options=[""] + direccion_options,
        index=([""] + direccion_options).index(st.session_state.direccion_selected) if st.session_state.direccion_selected in ([""] + direccion_options) else 0,
    )

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Siguiente"):
            if st.session_state.area_selected and st.session_state.direccion_selected:
                nuevo_codigo = build_codigo_borrador(st.session_state.direccion_selected)
                if st.session_state.codigo_expediente and st.session_state.codigo_expediente != nuevo_codigo:
                    st.session_state.registros = {}
                    st.session_state.observaciones = ""
                    st.session_state.draft_estado = "borrador"
                    st.session_state.permitir_editar_finalizado = False
                    reset_selectores_estudio()
                st.session_state.codigo_expediente = nuevo_codigo
                st.session_state.codigo_borrador = ""
                st.session_state.version_num = 0
                st.session_state.current_step = 3
                st.rerun()
            else:
                st.warning("Debe seleccionar un área y un centro docente para continuar.")
    with col2:
        if st.button("ATRÁS"):
            st.session_state.current_step = 1
            st.rerun()

    st.markdown("---")
    st.subheader("Cargar borrador existente")
    if supabase_available():
        drafts = list_drafts_from_supabase()
        if drafts:
            selected = st.selectbox("Borradores disponibles", options=[""] + drafts)
            manual = st.text_input("O introduce el código de borrador manualmente", value="")
            codigo = manual.strip() or (selected.split(" | ")[0] if selected else "")
            if st.button("Cargar borrador"):
                if codigo:
                    ok, msg = load_draft_from_supabase(codigo)
                    if ok:
                        st.success(msg)
                        st.session_state.current_step = 4
                        st.rerun()
                    else:
                        st.error(msg)
                else:
                    st.warning("Debe seleccionar o introducir un código de borrador.")
        else:
            st.info("No se encontraron borradores guardados o aún no hay conexión válida.")
    else:
        st.info("Cuando configuremos Supabase, aquí aparecerá la carga de borradores.")


def page_confirmacion() -> None:
    st.header("Paso 3: Confirmación de datos")
    unidad = st.session_state.direccion_selected
    columna_excel = COLUMNA_EXCEL_POR_DIRECCION.get(unidad, "")

    st.markdown(f"**Área:** <span style='color:#198754'>{st.session_state.area_selected}</span>", unsafe_allow_html=True)
    st.markdown(f"**Centro docente:** <span style='color:#0d6efd'>{unidad}</span>", unsafe_allow_html=True)
    st.markdown(f"**Expediente base:** `{st.session_state.codigo_expediente or build_codigo_borrador(unidad)}`")
    if st.session_state.codigo_borrador:
        st.markdown(f"**Versión cargada:** `{st.session_state.codigo_borrador}`")

    if columna_excel:
        st.success(f"El centro seleccionado se corresponde con la columna del Excel: **{columna_excel}**")
    else:
        st.warning(
            "Este centro no tiene una columna directa identificada en el Excel base. "
            "Se podrán guardar registros y generar hoja de registros, pero quizá haya que revisar el mapeo para la matriz final."
        )

    with st.expander("Revisión del mapeo centro docente ↔ columna Excel"):
        mapping_rows = []
        for direccion, codigo in CODIGOS_DIRECCION.items():
            mapping_rows.append({
                "Centro docente": direccion,
                "Código interno": codigo,
                "Columna Excel": COLUMNA_EXCEL_POR_DIRECCION.get(direccion, "PENDIENTE DE MAPEAR"),
            })
        st.dataframe(pd.DataFrame(mapping_rows), use_container_width=True, hide_index=True)
        st.caption("Esta tabla permite revisar qué centros vuelcan datos directamente en la matriz Excel y cuáles quedan pendientes de correspondencia definitiva.")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("SÍ, confirmar"):
            st.session_state.current_step = 4
            st.rerun()
    with col2:
        if st.button("ATRÁS"):
            st.session_state.current_step = 2
            st.rerun()


def page_entrada_datos() -> None:
    st.header("Paso 4: Introducción de datos")
    catalogo = load_catalogo()
    expediente_finalizado = st.session_state.get("draft_estado") == "finalizado"

    st.write(f"Centro docente seleccionado: **{st.session_state.direccion_selected}**")
    st.write(f"Expediente base: `{st.session_state.codigo_expediente or build_codigo_borrador(st.session_state.direccion_selected)}`")
    if st.session_state.codigo_borrador:
        st.write(f"Versión cargada: `{st.session_state.codigo_borrador}`")
    st.write(f"Estado actual: **{st.session_state.get('draft_estado', 'borrador').upper()}**")

    if expediente_finalizado:
        st.warning(
            "Este expediente está marcado como FINALIZADO. Para evitar cambios accidentales, la edición queda bloqueada "
            "hasta que active la casilla de edición."
        )
        st.checkbox(
            "Permitir edición de este expediente finalizado",
            key="permitir_editar_finalizado",
        )

    edicion_bloqueada = expediente_finalizado and not st.session_state.get("permitir_editar_finalizado", False)

    st.markdown("### Añadir o actualizar titulación")

    nivel_i_options = sorted_unique(catalogo["Nivel Estudio I"])
    st.selectbox(
        "Nivel Estudio I",
        options=[""] + nivel_i_options,
        key="sel_nivel_i",
        on_change=reset_downstream,
        args=("nivel_i",),
        disabled=edicion_bloqueada,
    )

    df2 = catalogo[catalogo["Nivel Estudio I"] == st.session_state.sel_nivel_i] if st.session_state.sel_nivel_i else catalogo.iloc[0:0]
    nivel_ii_options = sorted_unique(df2["Nivel Estudio II"]) if not df2.empty else []
    if st.session_state.sel_nivel_ii not in nivel_ii_options:
        st.session_state.sel_nivel_ii = ""
    st.selectbox(
        "Nivel Estudio II",
        options=[""] + nivel_ii_options,
        key="sel_nivel_ii",
        on_change=reset_downstream,
        args=("nivel_ii",),
        disabled=edicion_bloqueada or not bool(st.session_state.sel_nivel_i),
    )

    df3 = df2[df2["Nivel Estudio II"] == st.session_state.sel_nivel_ii] if st.session_state.sel_nivel_ii else df2.iloc[0:0]
    rama_options = sorted_unique(df3["Rama"]) if not df3.empty else []
    if st.session_state.sel_rama not in rama_options:
        st.session_state.sel_rama = ""
    st.selectbox(
        "Rama",
        options=[""] + rama_options,
        key="sel_rama",
        on_change=reset_downstream,
        args=("rama",),
        disabled=edicion_bloqueada or not bool(st.session_state.sel_nivel_ii),
    )

    df4 = df3[df3["Rama"] == st.session_state.sel_rama] if st.session_state.sel_rama else df3.iloc[0:0]
    titulacion_options = sorted_unique(df4["Titulación"]) if not df4.empty else []
    if st.session_state.sel_titulacion not in titulacion_options:
        st.session_state.sel_titulacion = ""
    st.selectbox(
        "Titulación",
        options=[""] + titulacion_options,
        key="sel_titulacion",
        disabled=edicion_bloqueada or not bool(st.session_state.sel_rama),
    )

    st.number_input(
        "Número de alumnos",
        min_value=0,
        step=1,
        key="numero_alumnos",
        disabled=edicion_bloqueada or not bool(st.session_state.sel_titulacion),
    )

    col_add, col_clear = st.columns(2)
    with col_add:
        if st.button("Añadir / actualizar registro", disabled=edicion_bloqueada):
            if not all([
                st.session_state.sel_nivel_i,
                st.session_state.sel_nivel_ii,
                st.session_state.sel_rama,
                st.session_state.sel_titulacion,
            ]):
                st.warning("Debe completar Nivel I, Nivel II, Rama y Titulación.")
            else:
                key = registro_key(
                    st.session_state.sel_nivel_i,
                    st.session_state.sel_nivel_ii,
                    st.session_state.sel_rama,
                    st.session_state.sel_titulacion,
                )
                st.session_state.registros[key] = {
                    "Nivel Estudio I": st.session_state.sel_nivel_i,
                    "Nivel Estudio II": st.session_state.sel_nivel_ii,
                    "Rama": st.session_state.sel_rama,
                    "Titulación": st.session_state.sel_titulacion,
                    "Nº alumnos": int(st.session_state.numero_alumnos),
                }
                audit_event("registro_actualizado", f"{st.session_state.sel_titulacion} | {int(st.session_state.numero_alumnos)} alumnos")
                st.success("Registro añadido/actualizado correctamente.")
    with col_clear:
        if st.button("Limpiar selectores", disabled=edicion_bloqueada):
            reset_selectores_estudio()
            st.rerun()

    st.markdown("---")
    st.markdown("### Registros introducidos")
    registros = list(st.session_state.registros.values())
    if registros:
        registros_df = pd.DataFrame(registros)
        st.dataframe(registros_df, use_container_width=True, hide_index=True)

        registro_labels = [
            f"{i + 1}. {r['Nivel Estudio II']} | {r['Rama']} | {r['Titulación']} | {r['Nº alumnos']} alumnos"
            for i, r in enumerate(registros)
        ]
        delete_label = st.selectbox("Seleccionar registro para eliminar", options=[""] + registro_labels, disabled=edicion_bloqueada)
        if st.button("Eliminar registro seleccionado", disabled=edicion_bloqueada):
            if delete_label:
                idx = int(delete_label.split(".", 1)[0]) - 1
                item = registros[idx]
                key = registro_key(item["Nivel Estudio I"], item["Nivel Estudio II"], item["Rama"], item["Titulación"])
                st.session_state.registros.pop(key, None)
                audit_event("registro_eliminado", item.get("Titulación", ""))
                st.success("Registro eliminado.")
                st.rerun()
            else:
                st.warning("Debe seleccionar un registro.")
    else:
        st.info("Todavía no se ha introducido ningún registro.")

    st.markdown("---")
    st.text_area("Observaciones internas del borrador", key="observaciones", disabled=edicion_bloqueada)

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        if st.button("Guardar borrador"):
            if expediente_finalizado and not st.session_state.get("permitir_editar_finalizado", False):
                st.warning("El expediente está finalizado. Active la edición si necesita reabrirlo como borrador.")
            else:
                if expediente_finalizado:
                    st.info("El expediente pasará de finalizado a borrador para poder corregirse.")
                ok, msg = save_draft_to_supabase(estado="borrador")
                if ok:
                    st.success(msg)
                else:
                    st.warning(msg)
    with col2:
        if st.button("Finalizar"):
            if not registros:
                st.warning("No hay registros introducidos.")
            else:
                st.session_state.current_step = 5
                st.rerun()
    with col3:
        if st.button("ATRÁS"):
            st.session_state.current_step = 3
            st.rerun()
    with col4:
        if st.button("Reiniciar datos", disabled=edicion_bloqueada):
            audit_event("reiniciar_datos", "Se eliminaron todos los registros de la sesión")
            st.session_state.registros = {}
            reset_selectores_estudio()
            st.rerun()


def page_resumen_descarga() -> None:
    st.header("Paso 5: Recordatorio, cierre y envío")
    registros = list(st.session_state.registros.values())
    if not registros:
        st.warning("No hay registros para generar.")
        if st.button("Volver a introducir datos"):
            st.session_state.current_step = 4
            st.rerun()
        return

    df = pd.DataFrame(registros)
    total_alumnos = int(df["Nº alumnos"].sum())
    codigo = st.session_state.codigo_borrador or build_codigo_borrador(st.session_state.direccion_selected)
    filename = f"{codigo}.xlsx"
    excel_bytes = build_output_excel()

    st.subheader("Resumen de datos introducidos")
    st.dataframe(df, use_container_width=True, hide_index=True)
    col_estado, col_total = st.columns(2)
    with col_estado:
        st.metric("Estado actual", st.session_state.get("draft_estado", "borrador").upper())
    with col_total:
        st.metric("Total alumnos introducidos", total_alumnos)

    st.markdown("---")
    st.subheader("Recordatorio antes de finalizar")
    st.warning(
        "Antes de cerrar el expediente, revise que el centro docente, la titulación y el número de alumnos son correctos. "
        "Si falta alguna titulación o está esperando datos de alguna especialidad, use 'REVISAR / VOLVER' o 'Guardar borrador'."
    )

    unidad = st.session_state.direccion_selected
    columna_excel = COLUMNA_EXCEL_POR_DIRECCION.get(unidad, "")
    if columna_excel:
        st.info(f"Los datos de este centro se volcarán en la columna del Excel: **{columna_excel}**.")
    else:
        st.error(
            "Este centro no tiene columna Excel mapeada. Se generará el Excel con la hoja de registros, "
            "pero la matriz principal puede no quedar volcada en una columna específica."
        )

    check1 = st.checkbox("He revisado que el Área y el Centro Docente son correctos", key="check_recordatorio_unidad")
    check2 = st.checkbox("He revisado que las titulaciones y el número de alumnos son correctos", key="check_recordatorio_datos")
    check3 = st.checkbox("Entiendo que debo descargar el Excel y/o enviarlo por correo según el procedimiento indicado", key="check_recordatorio_envio")

    st.markdown("---")
    st.subheader("Acciones finales")

    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("Guardar borrador"):
            ok, msg = save_draft_to_supabase(estado="borrador")
            if ok:
                st.success(msg)
            else:
                st.warning(msg)

    with col2:
        st.download_button(
            label="Descargar Excel",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption("La descarga no marca por sí sola el expediente como finalizado. Si quiere cerrar el expediente, use el botón de finalización.")

    with col3:
        if st.button("REVISAR / VOLVER"):
            st.session_state.current_step = 4
            st.rerun()

    col4, col5 = st.columns(2)
    with col4:
        if st.button("Finalizar expediente en Supabase"):
            if not all([check1, check2, check3]):
                st.warning("Debe marcar las tres casillas de revisión antes de finalizar.")
            else:
                ok, msg = save_draft_to_supabase(estado="finalizado")
                if ok:
                    st.success(msg)
                else:
                    st.warning(msg)

    with col5:
        if st.button("Enviar correo automático"):
            if not all([check1, check2, check3]):
                st.warning("Debe marcar las tres casillas de revisión antes de enviar el correo.")
            else:
                ok_save, msg_save = save_draft_to_supabase(estado="finalizado")
                if ok_save:
                    st.success(msg_save)
                    excel_bytes = build_output_excel()
                    filename = f"{st.session_state.codigo_borrador}.xlsx"
                    ok_mail, msg_mail = send_email_with_mailgun(excel_bytes, filename)
                    if ok_mail:
                        st.success(msg_mail)
                    else:
                        st.warning(msg_mail)
                else:
                    st.warning(msg_save)

    st.markdown("---")
    st.caption("Si Mailgun no está configurado, el botón de correo mostrará un aviso y podrá seguir usando la descarga manual.")


def render_admin_consolidado() -> None:
    st.subheader("Consolidado DCD")

    st.markdown(
        """
        Esta pantalla genera un Excel consolidado a partir de los expedientes guardados como **finalizado** en Supabase.

        Criterio aplicado:
        - Se toma el último expediente finalizado por cada centro docente.
        - Si dos unidades vuelcan en la misma columna, sus registros se suman en la matriz.
        - Las columnas de totales se calculan automáticamente:
          H = E+F+G, K = H+I+J, O = L+M+N, S = O+P+Q+R y T = K+S.
        """
    )

    if not supabase_available():
        st.warning("Supabase no está configurado. No se puede generar el consolidado.")
        return

    borradores, duplicados = get_latest_finalized_drafts()
    status_df, missing_centros = get_admin_centros_status()
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Expedientes finalizados usados", len(borradores))
    with col2:
        st.metric("Duplicados finalizados ignorados", len(duplicados))
    with col3:
        st.metric("Centros pendientes", len(missing_centros))

    st.subheader("Estado de centros docentes")
    st.dataframe(status_df, use_container_width=True, hide_index=True)

    if missing_centros:
        st.warning("Hay centros docentes pendientes de incorporarse al consolidado.")
        with st.expander("Ver centros pendientes"):
            for centro in missing_centros:
                st.write(f"- {centro}")
        if st.button("Enviar aviso de centros pendientes por correo"):
            ok, msg = send_missing_centros_email(missing_centros, status_df)
            if ok:
                st.success(msg)
            else:
                st.warning(msg)
    else:
        st.success("Todos los centros docentes previstos tienen expediente finalizado.")

    if borradores:
        preview = pd.DataFrame([{
            "Código borrador": b.get("codigo_borrador", ""),
            "Área": normalizar_area(b.get("area", "")),
            "Centro docente": b.get("unidad_docente", ""),
            "Guardado": b.get("saved_at") or b.get("updated_at", ""),
            "Columna Excel": COLUMNA_EXCEL_POR_DIRECCION.get(b.get("unidad_docente", ""), ""),
        } for b in borradores])
        st.subheader("Expedientes que entrarán en el consolidado")
        st.dataframe(preview, use_container_width=True, hide_index=True)
    else:
        st.info("Todavía no hay expedientes finalizados para consolidar.")

    if duplicados:
        with st.expander("Ver expedientes finalizados duplicados que se ignorarán"):
            st.dataframe(pd.DataFrame([{
                "Código borrador ignorado": b.get("codigo_borrador", ""),
                "Área": normalizar_area(b.get("area", "")),
                "Centro docente": b.get("unidad_docente", ""),
                "Guardado": b.get("saved_at") or b.get("updated_at", ""),
            } for b in duplicados]), use_container_width=True, hide_index=True)

    st.markdown("---")
    if st.button("Generar Excel consolidado"):
        ok, msg, excel_bytes, filename = build_consolidated_excel_from_supabase()
        if ok and excel_bytes:
            st.success(msg)
            st.download_button(
                label="Descargar Excel consolidado",
                data=excel_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning(msg)



def render_admin_publicaciones() -> None:
    st.subheader("Publicaciones oficiales")
    st.markdown(
        """
        Esta pantalla gestiona la publicación oficial de la Matriz DCD.

        - El administrador ve el histórico completo.
        - Las entidades externas, en una futura fase, solo verán la publicación marcada como **vigente**.
        - Una nueva publicación no borra la anterior: la archiva como no vigente.
        """
    )

    if not supabase_available():
        st.warning("Supabase no está configurado. No se pueden gestionar publicaciones.")
        return

    status_df, missing_centros = get_admin_centros_status()
    publicaciones = get_publicaciones()

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Centros pendientes", len(missing_centros))
    with col2:
        st.metric("Publicaciones históricas", len(publicaciones))
    with col3:
        vigente = next((p for p in publicaciones if p.get("publicacion_vigente")), None)
        st.metric("Publicación vigente", vigente.get("codigo_publicacion", "No") if vigente else "No")

    st.markdown("### Dashboard actual de datos finalizados")
    ok_dash, msg_dash, analytics_dash = build_current_analytics_from_supabase()
    if ok_dash and analytics_dash:
        render_streamlit_dashboard(analytics_dash)
    else:
        st.info(msg_dash)

    st.markdown("### Estado actual de centros")
    st.dataframe(status_df, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### Crear publicación")
    if missing_centros:
        st.warning("Hay centros docentes pendientes. Puede publicar con los datos disponibles, pero quedará registrado.")
        with st.expander("Ver centros pendientes"):
            for centro in missing_centros:
                st.write(f"- {centro}")
    else:
        st.success("Todos los centros docentes esperados tienen una versión finalizada.")

    motivo = st.text_area(
        "Motivo de publicación",
        value="Publicación generada desde Panel Administrador.",
        key="motivo_publicacion_admin",
    )
    permitir_pendientes = st.checkbox(
        "Permitir publicación con centros pendientes",
        value=False,
        help="Use esta opción solo si se ha alcanzado el vencimiento o si administrativamente procede publicar con los datos disponibles.",
    )
    confirm_publication = st.checkbox(
        "Confirmo que deseo generar una nueva publicación vigente. La anterior quedará archivada solo para administrador.",
        value=False,
    )

    if st.button("Generar publicación vigente"):
        if not confirm_publication:
            st.warning("Debe marcar la confirmación antes de publicar.")
        else:
            tipo = "manual_con_pendientes" if missing_centros else "manual_completa"
            ok, msg = create_publication(tipo_publicacion=tipo, motivo=motivo, allow_missing=permitir_pendientes)
            if ok:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)

    st.markdown("---")
    st.markdown("### Histórico de publicaciones")
    if not publicaciones:
        st.info("Todavía no hay publicaciones generadas.")
        return

    publicaciones_df = pd.DataFrame([{
        "Código": p.get("codigo_publicacion", ""),
        "Versión": p.get("version_publicacion", ""),
        "Fecha": p.get("fecha_publicacion", ""),
        "Vigente": "Sí" if p.get("publicacion_vigente") else "No",
        "Tipo": p.get("tipo_publicacion", ""),
        "Motivo": p.get("motivo_publicacion", ""),
        "Generada por": p.get("generada_por", ""),
        "Pendientes": len(p.get("centros_pendientes") or []),
    } for p in publicaciones])
    st.dataframe(publicaciones_df, use_container_width=True, hide_index=True)

    codigos = [p.get("codigo_publicacion", "") for p in publicaciones if p.get("codigo_publicacion")]
    selected = st.selectbox("Seleccionar publicación para descargar", options=[""] + codigos)
    pub = next((p for p in publicaciones if p.get("codigo_publicacion") == selected), None)
    if pub:
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Preparar descarga PDF"):
                ok, data, msg = download_publication_file(pub.get("ruta_pdf", ""))
                if ok and data:
                    st.download_button(
                        "Descargar PDF Matriz_DCD",
                        data=data,
                        file_name=f"{selected}_Matriz_DCD.pdf",
                        mime="application/pdf",
                        on_click=audit_event,
                        args=("descarga_pdf_historico_admin", selected),
                    )
                else:
                    st.error(msg)
        with c2:
            if st.button("Preparar descarga Excel"):
                ok, data, msg = download_publication_file(pub.get("ruta_excel", ""))
                if ok and data:
                    st.download_button(
                        "Descargar Excel consolidado",
                        data=data,
                        file_name=f"{selected}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        on_click=audit_event,
                        args=("descarga_excel_historico_admin", selected),
                    )
                else:
                    st.error(msg)

def render_admin_cierre() -> None:
    st.subheader("Cierre y publicación automática")
    st.markdown(
        """
        Configure cómo debe comportarse el sistema cuando los centros docentes finalizan sus expedientes.

        Importante: Streamlit no ejecuta tareas en segundo plano de forma permanente. La fecha tope se evalúa cuando cualquier usuario accede a la app,
        cuando un centro finaliza expediente o cuando el admin pulsa **Evaluar ahora cierre automático**.

        Esto significa que no es necesario que entre el administrador para que se dispare la comprobación: también puede dispararla internamente un centro docente al acceder o finalizar. El centro no ve datos ajenos; solo activa la comprobación automática del sistema.
        """
    )

    if not supabase_available():
        st.warning("Supabase no está configurado. No se puede guardar configuración de cierre.")
        return

    cfg = get_cierre_config()
    status_df, missing_centros = get_admin_centros_status()

    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Centros pendientes", len(missing_centros))
    with c2:
        fecha_txt = cfg.get("fecha_tope") or "No definida"
        st.metric("Fecha tope", fecha_txt)
    with c3:
        st.metric("Modo", CIERRE_MODE_LABELS.get(cfg.get("modo_cierre"), cfg.get("modo_cierre")))

    st.markdown("### Configuración")
    mode_keys = list(CIERRE_MODE_LABELS.keys())
    mode_labels = [CIERRE_MODE_LABELS[k] for k in mode_keys]
    current_mode = cfg.get("modo_cierre", "todos_finalizados")
    selected_label = st.selectbox(
        "Modo de cierre",
        options=mode_labels,
        index=mode_keys.index(current_mode) if current_mode in mode_keys else 0,
    )
    selected_mode = mode_keys[mode_labels.index(selected_label)]

    fecha_actual = parse_config_date(cfg.get("fecha_tope", ""))
    usar_fecha = st.checkbox("Definir fecha tope", value=fecha_actual is not None)
    fecha_tope_val = None
    if usar_fecha:
        fecha_tope_val = st.date_input(
            "Fecha tope",
            value=fecha_actual or (datetime.now().date() + timedelta(days=30)),
            format="DD/MM/YYYY",
        )
    dias_aviso = st.number_input(
        "Días de aviso previo al admin",
        min_value=0,
        max_value=60,
        value=int(cfg.get("dias_aviso_previo_int", 7)),
        step=1,
    )
    avisos_activos = st.checkbox("Activar avisos por correo al admin", value=bool(cfg.get("avisos_admin_activados_bool", True)))

    if st.button("Guardar configuración de cierre"):
        errors = []
        for key, value in [
            ("modo_cierre", selected_mode),
            ("fecha_tope", fecha_tope_val.isoformat() if usar_fecha and fecha_tope_val else ""),
            ("dias_aviso_previo", str(int(dias_aviso))),
            ("avisos_admin_activados", "true" if avisos_activos else "false"),
        ]:
            ok, msg = set_config_value(key, value)
            if not ok:
                errors.append(msg)
        if errors:
            st.error(" | ".join(errors))
        else:
            audit_event("configuracion_cierre_actualizada", f"Modo: {selected_mode}. Fecha tope: {fecha_tope_val if usar_fecha else 'sin fecha'}")
            st.success("Configuración de cierre guardada correctamente.")
            st.rerun()

    st.markdown("---")
    st.markdown("### Evaluación manual")
    st.dataframe(status_df, use_container_width=True, hide_index=True)

    col_eval, col_aviso = st.columns(2)
    with col_eval:
        if st.button("Evaluar ahora cierre automático"):
            msg = evaluar_cierre_automatico()
            if msg:
                st.info(msg)
            else:
                st.success("Evaluación realizada. No hay acciones automáticas pendientes ahora mismo.")
    with col_aviso:
        if st.button("Enviar aviso de pendientes ahora"):
            if missing_centros:
                ok, msg = send_missing_centros_email(missing_centros, status_df)
                if ok:
                    st.success(msg)
                else:
                    st.warning(msg)
            else:
                st.success("No hay centros pendientes que avisar.")

    st.caption("Los avisos automáticos dependen de Mailgun. Si Mailgun no está configurado, la app dejará constancia del estado pero no enviará correo.")
    st.caption("La publicación por fecha tope se evalúa al usar la app o al pulsar 'Evaluar ahora cierre automático'. Para automatismo 24/7 sin accesos habría que añadir una tarea programada externa.")


def render_admin_usuarios() -> None:
    st.subheader("Usuarios y permisos")
    st.info("Las contraseñas se guardan hasheadas. El administrador puede resetearlas, pero no verlas.")

    if not supabase_available():
        st.warning("Supabase no está configurado. No se pueden gestionar usuarios.")
        return

    users = list_users_from_supabase()
    if users:
        st.dataframe(pd.DataFrame(users), use_container_width=True, hide_index=True)
    else:
        st.warning("No se encontraron usuarios en Supabase. Puede crear el primer usuario admin desde este panel usando el acceso actual.")

    st.markdown("---")
    st.markdown("### Crear o actualizar usuario")

    role = st.selectbox("Rol", options=["usuario", "consulta", "admin"], key="admin_user_role")
    area = ""
    unidad_docente = ""

    if role == "usuario":
        area = st.selectbox("Área asignada", options=[""] + AREA_OPTIONS, key="admin_user_area")
        direccion_options = DIRECCIONES_POR_AREA.get(area, [])
        unidad_docente = st.selectbox("Centro docente asignado", options=[""] + direccion_options, key="admin_user_unidad")
    elif role == "consulta":
        st.info("Los usuarios de consulta externa solo pueden ver la publicación vigente. No quedan vinculados a un centro docente.")
    else:
        st.info("Los usuarios administradores no quedan vinculados a un centro docente concreto.")

    with st.form("form_usuario"):
        username = st.text_input("Usuario", placeholder="chuimi")
        display_name = st.text_input("Nombre visible", placeholder="CHUIMI")
        activo = st.checkbox("Usuario activo", value=True)
        must_change = st.checkbox("Obligar a cambiar contraseña en el primer acceso", value=True)
        password = st.text_input("Contraseña temporal", type="password")

        submitted = st.form_submit_button("Guardar usuario")
        if submitted:
            ok, msg = save_user_to_supabase(
                username=username,
                display_name=display_name,
                role=role,
                area=area,
                unidad_docente=unidad_docente,
                password=password,
                activo=activo,
                must_change_password=must_change,
            )
            if ok:
                st.success(msg)
            else:
                st.error(msg)

    st.markdown("---")
    st.markdown("### Mantenimiento de usuarios")
    st.caption("Recomendación: use primero Desactivar para conservar trazabilidad. Eliminar borra el usuario de la tabla operativa, aunque las auditorías y registros históricos mantienen el nombre de usuario como texto.")
    usernames = [u.get("username", "") for u in users if u.get("username")]
    selected_user = st.selectbox("Usuario existente", options=[""] + usernames)
    new_password = st.text_input("Nueva contraseña temporal", type="password", key="admin_reset_password")

    col_reset, col_activate, col_deactivate = st.columns(3)
    with col_reset:
        if st.button("Resetear contraseña"):
            ok, msg = reset_user_password(selected_user, new_password, must_change_password=True)
            if ok:
                st.success(msg)
            else:
                st.error(msg)
    with col_activate:
        if st.button("Activar"):
            ok, msg = set_user_active(selected_user, True)
            if ok:
                st.success(msg)
            else:
                st.error(msg)
    with col_deactivate:
        if st.button("Desactivar"):
            if selected_user == st.session_state.get("current_user"):
                st.warning("No puede desactivar el usuario con el que está trabajando ahora mismo.")
            else:
                ok, msg = set_user_active(selected_user, False)
                if ok:
                    st.success(msg)
                else:
                    st.error(msg)

    st.markdown("#### Eliminar usuario")
    st.warning("La eliminación es definitiva para la tabla de usuarios. Para bajas ordinarias, es preferible desactivar el usuario.")
    confirm_delete = st.text_input("Para eliminar el usuario seleccionado escriba ELIMINAR", key="admin_delete_confirm")
    if st.button("Eliminar usuario definitivamente", type="secondary"):
        if not selected_user:
            st.error("Seleccione primero un usuario.")
        elif selected_user == st.session_state.get("current_user"):
            st.warning("No puede eliminar el usuario con el que está trabajando ahora mismo.")
        elif confirm_delete != "ELIMINAR":
            st.error("Debe escribir ELIMINAR exactamente para confirmar.")
        else:
            ok, msg = delete_user_from_supabase(selected_user)
            if ok:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)


def render_admin_auditoria_backup() -> None:
    st.subheader("Auditoría y backup")
    st.info("Esta pantalla permite revisar acciones recientes y exportar una copia administrativa de las tablas principales.")

    if not supabase_available():
        st.warning("Supabase no está configurado. No se puede consultar auditoría ni generar backup.")
        return

    st.markdown("### Auditoría reciente")
    audit_df, err = fetch_table_for_backup("dcd_auditoria", limit=1000)
    if err:
        st.warning(f"No se pudo cargar la auditoría: {err}")
    elif audit_df.empty:
        st.info("No hay eventos de auditoría registrados todavía.")
    else:
        if "created_at" in audit_df.columns:
            audit_df = audit_df.sort_values("created_at", ascending=False)
        acciones = sorted([a for a in audit_df.get("accion", pd.Series(dtype=str)).dropna().unique().tolist() if a])
        usuarios = sorted([u for u in audit_df.get("usuario", pd.Series(dtype=str)).dropna().unique().tolist() if u])

        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            filtro_accion = st.selectbox("Filtrar por acción", options=["Todas"] + acciones)
        with col_f2:
            filtro_usuario = st.selectbox("Filtrar por usuario", options=["Todos"] + usuarios)
        with col_f3:
            limite = st.number_input("Eventos a mostrar", min_value=50, max_value=1000, value=200, step=50)

        filtered = audit_df.copy()
        if filtro_accion != "Todas" and "accion" in filtered.columns:
            filtered = filtered[filtered["accion"] == filtro_accion]
        if filtro_usuario != "Todos" and "usuario" in filtered.columns:
            filtered = filtered[filtered["usuario"] == filtro_usuario]
        st.dataframe(filtered.head(int(limite)), use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### Backup administrativo")
    st.warning("El backup puede contener información administrativa sensible. Descárguelo y consérvelo en una ubicación segura.")
    include_hashes = st.checkbox(
        "Incluir hashes de contraseña en el backup (no recomendado)",
        value=False,
        help="Los hashes no son contraseñas visibles, pero siguen siendo material sensible. Déjelo desmarcado salvo necesidad técnica.",
    )

    if st.button("Generar backup completo"):
        ok, data, msg = generate_admin_backup_excel(include_password_hashes=include_hashes)
        if ok and data:
            filename = f"DCD_backup_admin_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            st.success(msg)
            st.download_button(
                "Descargar backup Excel",
                data=data,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                on_click=audit_event,
                args=("descarga_backup_admin", filename),
            )
        else:
            st.error(msg)



def page_portal_externo() -> None:
    st.title("Portal de consulta - DATOS CAPACIDAD DOCENTE")
    st.markdown("---")
    app_sidebar()

    st.subheader("Publicación vigente")
    st.info("Este portal muestra únicamente la publicación vigente. Las versiones históricas solo están disponibles para el administrador.")

    if not supabase_available():
        st.warning("Supabase no está configurado. No se puede consultar la publicación vigente.")
        return

    pub = get_publicacion_vigente()
    if not pub:
        st.warning("Todavía no hay una publicación vigente disponible.")
        return

    render_publication_metadata(pub)

    ok, msg, analytics = load_analytics_from_publication_excel(pub)
    if ok and analytics:
        st.markdown("### Resumen de datos")
        render_streamlit_dashboard(analytics)

        with st.expander("Ver tablas resumen completas"):
            for sheet_name in ["Resumen_Provincia", "Resumen_Isla", "Resumen_Centro", "Resumen_Rama", "Resumen_Nivel", "Top_Titulaciones"]:
                df = analytics.get(sheet_name, pd.DataFrame())
                if not df.empty:
                    st.markdown(f"#### {sheet_name.replace('_', ' ')}")
                    st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.warning(msg)

    st.markdown("---")
    st.markdown("### Descargas")
    render_publication_downloads(pub, prefix="external_portal")

    audit_key = f"auditoria_consulta_publicacion_{pub.get('codigo_publicacion', '')}"
    if not st.session_state.get(audit_key):
        audit_event("consulta_publicacion_vigente", f"Consulta externa de {pub.get('codigo_publicacion', '')}")
        st.session_state[audit_key] = True


def page_admin_consolidado() -> None:
    st.header("Panel administrador")

    if st.session_state.get("current_user_role") != "admin":
        st.error("Esta pantalla solo está disponible para usuarios administradores.")
        if st.button("Volver"):
            st.session_state.current_step = 2
            st.rerun()
        return

    tab_consolidado, tab_publicaciones, tab_cierre, tab_usuarios, tab_auditoria = st.tabs(["Consolidado", "Publicaciones", "Cierre", "Usuarios", "Auditoría/Backup"])
    with tab_consolidado:
        render_admin_consolidado()
    with tab_publicaciones:
        render_admin_publicaciones()
    with tab_cierre:
        render_admin_cierre()
    with tab_usuarios:
        render_admin_usuarios()
    with tab_auditoria:
        render_admin_auditoria_backup()

    st.markdown("---")
    if st.button("Volver al aplicativo"):
        st.session_state.current_step = 2
        st.rerun()


# =========================================================
# EJECUCIÓN PRINCIPAL
# =========================================================
init_session_state()

if not st.session_state.logged_in:
    page_login()
    st.stop()

if st.session_state.get("must_change_password"):
    page_change_password()
    st.stop()

if is_external_viewer():
    page_portal_externo()
    st.stop()

st.title(APP_TITLE)
st.markdown("---")
app_sidebar()

try:
    # Precarga para detectar problemas del Excel cuanto antes.
    _ = load_catalogo()
except Exception as exc:
    st.error(f"No se pudo cargar el Excel base: {exc}")
    st.stop()

# Evaluación oportunista de cierre automático/fecha tope.
# Streamlit no ejecuta procesos en segundo plano; esta comprobación se realiza al usar la app.
if supabase_available():
    try:
        cierre_msg = evaluar_cierre_automatico()
        if cierre_msg and is_admin():
            st.sidebar.info(cierre_msg)
    except Exception:
        pass

if st.session_state.current_step == 1:
    page_instrucciones()
elif st.session_state.current_step == 2:
    page_seleccion_unidad()
elif st.session_state.current_step == 3:
    page_confirmacion()
elif st.session_state.current_step == 4:
    page_entrada_datos()
elif st.session_state.current_step == 5:
    page_resumen_descarga()
elif st.session_state.current_step == 6:
    page_admin_consolidado()
else:
    st.session_state.current_step = 1
    st.rerun()
