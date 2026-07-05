# ============================================================
# DATOS CAPACIDAD DOCENTE (DCD)
# Aplicativo desarrollado para la gestión y análisis de los Datos
# de Capacidad Docente del Servicio Canario de la Salud.
#
# Desarrollador / creador del programa: Alberto Cabrera
# Responsable funcional del proyecto: Alberto Cabrera
# Versión: DCD 1.2.2 estable
# Año: 2026
#
# Nota de autoría:
# Este bloque identifica la autoría funcional y de desarrollo del
# aplicativo. Debe mantenerse en las versiones derivadas del proyecto.
# ============================================================

import io
import base64
import hashlib
import json
import re
import secrets as py_secrets
from datetime import datetime, date, timedelta
from pathlib import Path

import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components

try:
    from supabase import create_client
except Exception:
    create_client = None

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A3, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
except Exception:
    colors = None
    landscape = None
    A3 = None
    getSampleStyleSheet = None
    ParagraphStyle = None
    SimpleDocTemplate = None
    Table = None
    TableStyle = None
    Paragraph = None
    Spacer = None
    Image = None


# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================
APP_VERSION = "DCD 1.2.2 estable"
APP_TITLE = "DATOS CAPACIDAD DOCENTE (DCD 1.0)"
APP_AUTHOR = "Alberto Cabrera"
APP_CREATOR = "Alberto Cabrera"
APP_DEVELOPED_FOR = "F.S.E. – S.C.S."
APP_BUILD_ID = "DCD-2026-FSE-SCS"
DEFAULT_PASSWORD = "Capacidad2026"
EXCEL_PATH = Path(__file__).parent / "data" / "listado_para_capacidad_docente.xlsx"
ASSETS_DIR = Path(__file__).parent / "assets"
LOGO_PATH = ASSETS_DIR / "logo.png"
MAPA_CANARIAS_PATH = ASSETS_DIR / "mapa_canarias.png"
INSTITUTIONAL_PHRASE = "Informe desarrollado para la gestión y análisis de los Datos de Capacidad Docente del Servicio Canario de la Salud."
SIGNATURE_FOOTER = "Jefatura del Servicio de Formacion Sanitaria Especializada"


def make_json_safe(value):
    """Convierte estructuras Python/Pandas en JSON válido para Supabase.

    Pandas puede generar NaN/NaT/inf en DataFrames y diccionarios.
    JSON/Supabase no acepta esos valores, por lo que se normalizan a None.
    """
    try:
        if value is None:
            return None
        if value is pd.NA or value is pd.NaT:
            return None
        if isinstance(value, float):
            if pd.isna(value) or value in (float("inf"), float("-inf")):
                return None
            return value
        if isinstance(value, (int, bool, str)):
            return value
        if isinstance(value, (datetime, date)):
            return value.isoformat()
        if isinstance(value, dict):
            return {str(k): make_json_safe(v) for k, v in value.items()}
        if isinstance(value, (list, tuple, set)):
            return [make_json_safe(v) for v in value]
        if pd.isna(value):
            return None
        return value
    except Exception:
        return str(value) if value is not None else None


def dataframe_records_json_safe(df: pd.DataFrame) -> list[dict]:
    """Devuelve registros de un DataFrame limpiando NaN/NaT/inf para JSON."""
    if df is None or df.empty:
        return []
    cleaned = df.replace([float("inf"), float("-inf")], pd.NA)
    return make_json_safe(cleaned.to_dict(orient="records"))


CONSULTA_EXCEL_SHEETS = [
    "Dashboard",
    "Resumen_Global",
    "Resumen_Provincia",
    "Resumen_Isla",
    "Resumen_Centro",
    "Resumen_Rama",
    "Resumen_Nivel",
    "Resumen_Centro_Rama",
    "Resumen_Centro_Nivel",
    "Top_Titulaciones",
    "Turnos_Resumen",
    "Turnos_Resumen_Centro",
    "Turnos_Resumen_Isla",
    "Turnos_Detalle",
    "Observaciones_Titulacion",
    "Detalle_Alumnos",
    "Matriz_DCD",
]

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
        "GERENCIA DE ATENCIÓN PRIMARIA DE TENERIFE",
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

# Compatibilidad histórica: en versiones anteriores aparecían dos centros para GAP Tenerife.
# La matriz oficial solo tiene una columna GAP TF, por lo que ambos nombres antiguos
# se normalizan al único centro válido.
LEGACY_UNIDAD_DOCENTE_MAP = {
    "GERENCIA DE ATENCIÓN PRIMARIA DE TENERIFE NORTE": "GERENCIA DE ATENCIÓN PRIMARIA DE TENERIFE",
    "GERENCIA DE ATENCIÓN PRIMARIA DE TENERIFE SUR": "GERENCIA DE ATENCIÓN PRIMARIA DE TENERIFE",
}

CODIGOS_DIRECCION = {
    "DIRECCIÓN GERENCIA HOSPITAL DOCTOR NEGRIN": "HUGCNEGRIN",
    "DIRECCIÓN GERENCIA COMPLEJO HOSPITALARIO UNIVERSITARIO Y MATERNO INFANTIL": "CHUIMI",
    "DIRECCIÓN GERENCIA COMPLEJO HOSPITALARIO UNIVERSITARIO DE CANARIAS": "CHUC",
    "DIRECCIÓN GERENCIA HOSPITAL NUESTRA SEÑORA DE CANDELARIA": "HUNSC",
    "GERENCIA DE ATENCIÓN PRIMARIA DE GRAN CANARIA": "GAPGC",
    "GERENCIA DE ATENCIÓN PRIMARIA DE TENERIFE": "GAPTF",
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
    "GERENCIA DE ATENCIÓN PRIMARIA DE TENERIFE": "GAP TF",
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

TURNOS_ALUMNOS_DISPLAY = [
    "Alumnos mañana",
    "Alumnos tarde",
    "Alumnos rotatorio",
    "Alumnos deslizante",
]

TURNOS_DB_MAP = {
    "Alumnos mañana": "alumnos_manana",
    "Alumnos tarde": "alumnos_tarde",
    "Alumnos rotatorio": "alumnos_rotatorio",
    "Alumnos deslizante": "alumnos_deslizante",
}

DESLIZANTE_DISPLAY = [
    "Deslizante lunes",
    "Deslizante martes",
    "Deslizante miércoles",
    "Deslizante jueves",
    "Deslizante viernes",
]

DESLIZANTE_DB_MAP = {
    "Deslizante lunes": "deslizante_lunes",
    "Deslizante martes": "deslizante_martes",
    "Deslizante miércoles": "deslizante_miercoles",
    "Deslizante jueves": "deslizante_jueves",
    "Deslizante viernes": "deslizante_viernes",
}

OBS_TITULACION_DISPLAY = "Observaciones titulación"
OBS_TITULACION_DB = "observaciones_titulacion"
DETALLE_ALUMNOS_DISPLAY = "Detalle alumnos"
DETALLE_ALUMNOS_DB = "detalle_alumnos"
TURNO_DIA_OPTIONS = ["", "M", "T", "R"]


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
        "sel_titulacion_libre_catalogo": "",
        "numero_alumnos": 0,
        "alumnos_manana": 0,
        "alumnos_tarde": 0,
        "alumnos_rotatorio": 0,
        "alumnos_deslizante": 0,
        "deslizante_lunes": "",
        "deslizante_martes": "",
        "deslizante_miercoles": "",
        "deslizante_jueves": "",
        "deslizante_viernes": "",
        "observaciones_titulacion": "",
        "detalle_alumnos": [],
        "codigo_borrador": "",
        "codigo_expediente": "",
        "version_num": 0,
        "draft_estado": "borrador",
        "permitir_editar_finalizado": False,
        "observaciones": "",
        "last_message": "",
        "editing_registro_key": "",
        "pending_edit_registro": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value



def normalizar_detalle_alumnos(value, total: int | None = None) -> list[dict]:
    """Normaliza el detalle opcional por alumno.

    Formato interno/JSON ampliado desde DCD 1.2.2 estable:
    [{"alumno": 1, "servicio": "...", "curso": "...", "es_deslizante": true,
      "deslizante_lunes": "M", ...}, ...]

    Solo se conservan filas con servicio, curso o patrón deslizante informado,
    salvo que se solicite expansión a un total para pintar el formulario.
    """
    data = []
    if value is None or value == "":
        data = []
    elif isinstance(value, str):
        try:
            parsed = json.loads(value)
            data = parsed if isinstance(parsed, list) else []
        except Exception:
            data = []
    elif isinstance(value, list):
        data = value
    else:
        data = []

    por_alumno: dict[int, dict] = {}
    for idx, entry in enumerate(data, start=1):
        if not isinstance(entry, dict):
            continue
        alumno = safe_int(entry.get("alumno") or entry.get("Alumno") or idx, idx)
        if alumno <= 0:
            alumno = idx
        servicio = str(entry.get("servicio") or entry.get("Servicio") or "").strip()
        curso = str(entry.get("curso") or entry.get("Curso") or entry.get("Curso/año") or "").strip()
        des_lunes = normalizar_turno_dia(entry.get("deslizante_lunes") or entry.get("Deslizante lunes") or "")
        des_martes = normalizar_turno_dia(entry.get("deslizante_martes") or entry.get("Deslizante martes") or "")
        des_miercoles = normalizar_turno_dia(entry.get("deslizante_miercoles") or entry.get("deslizante_miércoles") or entry.get("Deslizante miércoles") or "")
        des_jueves = normalizar_turno_dia(entry.get("deslizante_jueves") or entry.get("Deslizante jueves") or "")
        des_viernes = normalizar_turno_dia(entry.get("deslizante_viernes") or entry.get("Deslizante viernes") or "")
        es_deslizante = bool(entry.get("es_deslizante") or entry.get("Es deslizante") or any([des_lunes, des_martes, des_miercoles, des_jueves, des_viernes]))
        row = {
            "alumno": alumno,
            "servicio": servicio,
            "curso": curso,
            "es_deslizante": es_deslizante,
            "deslizante_lunes": des_lunes,
            "deslizante_martes": des_martes,
            "deslizante_miercoles": des_miercoles,
            "deslizante_jueves": des_jueves,
            "deslizante_viernes": des_viernes,
        }
        if servicio or curso or es_deslizante or total is not None:
            por_alumno[alumno] = row

    if total is not None:
        total = max(safe_int(total, 0), 0)
        return [por_alumno.get(i, {
            "alumno": i,
            "servicio": "",
            "curso": "",
            "es_deslizante": False,
            "deslizante_lunes": "",
            "deslizante_martes": "",
            "deslizante_miercoles": "",
            "deslizante_jueves": "",
            "deslizante_viernes": "",
        }) for i in range(1, total + 1)]

    return [
        por_alumno[k]
        for k in sorted(por_alumno)
        if por_alumno[k].get("servicio") or por_alumno[k].get("curso") or por_alumno[k].get("es_deslizante")
    ]


def normalizar_turno_dia(value) -> str:
    value = str(value or "").strip().upper()
    return value if value in {"M", "T", "R"} else ""


def patron_deslizante_entry(entry: dict) -> dict:
    """Devuelve el patrón semanal normalizado de una entrada de detalle_alumnos."""
    return {
        "deslizante_lunes": normalizar_turno_dia(entry.get("deslizante_lunes", "")),
        "deslizante_martes": normalizar_turno_dia(entry.get("deslizante_martes", "")),
        "deslizante_miercoles": normalizar_turno_dia(entry.get("deslizante_miercoles", "")),
        "deslizante_jueves": normalizar_turno_dia(entry.get("deslizante_jueves", "")),
        "deslizante_viernes": normalizar_turno_dia(entry.get("deslizante_viernes", "")),
    }


def entry_tiene_patron_deslizante_completo(entry: dict) -> bool:
    patron = patron_deslizante_entry(entry)
    return all(patron.values())


def formatear_patron_deslizante_entry(entry: dict) -> str:
    patron = patron_deslizante_entry(entry)
    if not any(patron.values()):
        return ""
    return " / ".join([
        f"L:{patron['deslizante_lunes']}",
        f"M:{patron['deslizante_martes']}",
        f"X:{patron['deslizante_miercoles']}",
        f"J:{patron['deslizante_jueves']}",
        f"V:{patron['deslizante_viernes']}",
    ])


def deslizantes_desde_detalle(detalle: list[dict], n_deslizantes: int = 0) -> list[dict]:
    """Obtiene, por orden, las entradas marcadas como deslizantes."""
    detalle = normalizar_detalle_alumnos(detalle)
    rows = [entry for entry in detalle if bool(entry.get("es_deslizante"))]
    if n_deslizantes > 0:
        rows = rows[:n_deslizantes]
    return rows


def formatear_patrones_deslizantes_detalle(detalle: list[dict], n_deslizantes: int = 0) -> str:
    partes = []
    for idx, entry in enumerate(deslizantes_desde_detalle(detalle, n_deslizantes), start=1):
        patron = formatear_patron_deslizante_entry(entry)
        if patron:
            partes.append(f"Alumno deslizante {idx}: {patron}")
    return "; ".join(partes)


def detalle_con_fallback_deslizante(detalle: list[dict], n_deslizantes: int, item: dict | None = None) -> list[dict]:
    """Garantiza patrones deslizantes para registros antiguos con un único patrón global."""
    n_deslizantes = max(safe_int(n_deslizantes, 0), 0)
    detalle_norm = normalizar_detalle_alumnos(detalle)
    if n_deslizantes <= 0:
        return detalle_norm
    actuales = deslizantes_desde_detalle(detalle_norm, n_deslizantes)
    if len(actuales) >= n_deslizantes and all(entry_tiene_patron_deslizante_completo(e) for e in actuales):
        return detalle_norm

    item = item or {}
    fallback = {
        "deslizante_lunes": normalizar_turno_dia(item.get("Deslizante lunes", "")),
        "deslizante_martes": normalizar_turno_dia(item.get("Deslizante martes", "")),
        "deslizante_miercoles": normalizar_turno_dia(item.get("Deslizante miércoles", "")),
        "deslizante_jueves": normalizar_turno_dia(item.get("Deslizante jueves", "")),
        "deslizante_viernes": normalizar_turno_dia(item.get("Deslizante viernes", "")),
    }
    if not all(fallback.values()):
        return detalle_norm

    por_alumno = {safe_int(e.get("alumno", 0)): dict(e) for e in detalle_norm if safe_int(e.get("alumno", 0)) > 0}
    for alumno in range(1, n_deslizantes + 1):
        row = por_alumno.get(alumno, {"alumno": alumno, "servicio": "", "curso": ""})
        row.update({"es_deslizante": True, **fallback})
        por_alumno[alumno] = row
    return [por_alumno[k] for k in sorted(por_alumno)]


def limpiar_detalle_alumnos_widgets() -> None:
    """Limpia los campos dinámicos de detalle por alumno del formulario."""
    for key in list(st.session_state.keys()):
        if (
            key.startswith("detalle_alumno_servicio_")
            or key.startswith("detalle_alumno_curso_")
            or key.startswith("detalle_alumno_deslizante_")
        ):
            del st.session_state[key]
    st.session_state.detalle_alumnos = []


def preparar_detalle_alumnos_widgets(total: int, detalle: list[dict] | None = None) -> None:
    """Inicializa, sin sobrescribir lo ya escrito, los widgets de detalle por alumno."""
    total = max(safe_int(total, 0), 0)
    detalle_norm = normalizar_detalle_alumnos(detalle if detalle is not None else st.session_state.get("detalle_alumnos", []), total=total)
    for entry in detalle_norm:
        alumno = safe_int(entry.get("alumno", 0), 0)
        if alumno <= 0:
            continue
        servicio_key = f"detalle_alumno_servicio_{alumno}"
        curso_key = f"detalle_alumno_curso_{alumno}"
        if servicio_key not in st.session_state:
            st.session_state[servicio_key] = str(entry.get("servicio", "") or "")
        if curso_key not in st.session_state:
            st.session_state[curso_key] = str(entry.get("curso", "") or "")


def preparar_deslizantes_alumno_widgets(n_deslizantes: int, detalle: list[dict] | None = None, item_fallback: dict | None = None) -> None:
    """Inicializa los patrones semanales por cada alumno en turno deslizante."""
    n_deslizantes = max(safe_int(n_deslizantes, 0), 0)
    detalle_norm = detalle_con_fallback_deslizante(
        detalle if detalle is not None else st.session_state.get("detalle_alumnos", []),
        n_deslizantes,
        item_fallback,
    )
    deslizantes = deslizantes_desde_detalle(detalle_norm, n_deslizantes)
    por_orden = {idx: entry for idx, entry in enumerate(deslizantes, start=1)}
    for idx in range(1, n_deslizantes + 1):
        entry = por_orden.get(idx, {})
        for dia in ["lunes", "martes", "miercoles", "jueves", "viernes"]:
            key = f"detalle_alumno_deslizante_{idx}_{dia}"
            if key not in st.session_state:
                st.session_state[key] = normalizar_turno_dia(entry.get(f"deslizante_{dia}", ""))


def recoger_detalle_alumnos_desde_widgets(total: int) -> list[dict]:
    """Recoge Servicio/Curso y el patrón deslizante individual de cada alumno deslizante."""
    total = max(safe_int(total, 0), 0)
    n_deslizantes = max(safe_int(st.session_state.get("alumnos_deslizante", 0), 0), 0)
    rows = []
    for alumno in range(1, total + 1):
        servicio = str(st.session_state.get(f"detalle_alumno_servicio_{alumno}", "") or "").strip()
        curso = str(st.session_state.get(f"detalle_alumno_curso_{alumno}", "") or "").strip()
        row = {"alumno": alumno, "servicio": servicio, "curso": curso}
        if alumno <= n_deslizantes:
            patron = {
                "deslizante_lunes": normalizar_turno_dia(st.session_state.get(f"detalle_alumno_deslizante_{alumno}_lunes", "")),
                "deslizante_martes": normalizar_turno_dia(st.session_state.get(f"detalle_alumno_deslizante_{alumno}_martes", "")),
                "deslizante_miercoles": normalizar_turno_dia(st.session_state.get(f"detalle_alumno_deslizante_{alumno}_miercoles", "")),
                "deslizante_jueves": normalizar_turno_dia(st.session_state.get(f"detalle_alumno_deslizante_{alumno}_jueves", "")),
                "deslizante_viernes": normalizar_turno_dia(st.session_state.get(f"detalle_alumno_deslizante_{alumno}_viernes", "")),
            }
            if any(patron.values()):
                row.update({"es_deslizante": True, **patron})
        if row.get("servicio") or row.get("curso") or row.get("es_deslizante"):
            rows.append(row)
    return normalizar_detalle_alumnos(rows)


def patrones_deslizantes_widgets_completos(n_deslizantes: int) -> bool:
    n_deslizantes = max(safe_int(n_deslizantes, 0), 0)
    for idx in range(1, n_deslizantes + 1):
        valores = [
            normalizar_turno_dia(st.session_state.get(f"detalle_alumno_deslizante_{idx}_lunes", "")),
            normalizar_turno_dia(st.session_state.get(f"detalle_alumno_deslizante_{idx}_martes", "")),
            normalizar_turno_dia(st.session_state.get(f"detalle_alumno_deslizante_{idx}_miercoles", "")),
            normalizar_turno_dia(st.session_state.get(f"detalle_alumno_deslizante_{idx}_jueves", "")),
            normalizar_turno_dia(st.session_state.get(f"detalle_alumno_deslizante_{idx}_viernes", "")),
        ]
        if not all(valores):
            return False
    return True


def primer_patron_deslizante_para_campos_globales(detalle: list[dict]) -> dict:
    deslizantes = deslizantes_desde_detalle(detalle, 1)
    if not deslizantes:
        return {"lunes": "", "martes": "", "miercoles": "", "jueves": "", "viernes": ""}
    entry = deslizantes[0]
    return {
        "lunes": normalizar_turno_dia(entry.get("deslizante_lunes", "")),
        "martes": normalizar_turno_dia(entry.get("deslizante_martes", "")),
        "miercoles": normalizar_turno_dia(entry.get("deslizante_miercoles", "")),
        "jueves": normalizar_turno_dia(entry.get("deslizante_jueves", "")),
        "viernes": normalizar_turno_dia(entry.get("deslizante_viernes", "")),
    }




def formatear_detalle_alumnos_display(value) -> str:
    """Devuelve un texto compacto para visualizar detalle_alumnos en tablas de Streamlit.

    Evita que Streamlit muestre listas de diccionarios como [object Object].
    No altera el dato real guardado ni la exportación a Excel.
    """
    detalle = normalizar_detalle_alumnos(value)
    partes = []
    for entry in detalle:
        alumno = safe_int(entry.get("alumno", 0), 0)
        servicio = str(entry.get("servicio", "") or "").strip()
        curso = str(entry.get("curso", "") or "").strip()
        patron = formatear_patron_deslizante_entry(entry) if entry.get("es_deslizante") else ""
        if not servicio and not curso and not patron:
            continue
        etiqueta = f"Alumno {alumno}" if alumno else "Alumno"
        valores = []
        if servicio:
            valores.append(servicio)
        if curso:
            valores.append(curso)
        if patron:
            valores.append(f"Deslizante {patron}")
        partes.append(f"{etiqueta}: " + " / ".join(valores))
    return "; ".join(partes)


def preparar_registros_para_visualizacion(df: pd.DataFrame) -> pd.DataFrame:
    """Prepara la tabla de registros introducidos para pantalla, sin tocar datos internos."""
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.copy()
    if DETALLE_ALUMNOS_DISPLAY in out.columns:
        out[DETALLE_ALUMNOS_DISPLAY] = out[DETALLE_ALUMNOS_DISPLAY].apply(formatear_detalle_alumnos_display)
    return out

def dataframe_excel_safe(df: pd.DataFrame) -> pd.DataFrame:
    """Convierte listas/dicts a texto JSON para que Excel/XlsxWriter no falle."""
    if df is None:
        return pd.DataFrame()
    out = df.copy()
    for col in out.columns:
        out[col] = out[col].apply(lambda v: json.dumps(make_json_safe(v), ensure_ascii=False) if isinstance(v, (list, dict)) else v)
    return out

def reset_selectores_estudio() -> None:
    st.session_state.sel_nivel_i = ""
    st.session_state.sel_nivel_ii = ""
    st.session_state.sel_rama = ""
    st.session_state.sel_titulacion = ""
    st.session_state.numero_alumnos = 0
    st.session_state.alumnos_manana = 0
    st.session_state.alumnos_tarde = 0
    st.session_state.alumnos_rotatorio = 0
    st.session_state.alumnos_deslizante = 0
    st.session_state.deslizante_lunes = ""
    st.session_state.deslizante_martes = ""
    st.session_state.deslizante_miercoles = ""
    st.session_state.deslizante_jueves = ""
    st.session_state.deslizante_viernes = ""
    st.session_state.observaciones_titulacion = ""
    limpiar_detalle_alumnos_widgets()


def cargar_registro_en_formulario(item: dict) -> None:
    """Carga un registro ya introducido en el formulario del Paso 4 para poder editarlo."""
    st.session_state.editing_registro_key = registro_key(
        item.get("Nivel Estudio I", ""),
        item.get("Nivel Estudio II", ""),
        item.get("Rama", ""),
        item.get("Titulación", ""),
    )
    st.session_state.sel_nivel_i = item.get("Nivel Estudio I", "")
    st.session_state.sel_nivel_ii = item.get("Nivel Estudio II", "")
    st.session_state.sel_rama = item.get("Rama", "")
    st.session_state.sel_titulacion = item.get("Titulación", "")
    st.session_state.numero_alumnos = safe_int(item.get("Nº alumnos", 0))
    st.session_state.alumnos_manana = safe_int(item.get("Alumnos mañana", 0))
    st.session_state.alumnos_tarde = safe_int(item.get("Alumnos tarde", 0))
    st.session_state.alumnos_rotatorio = safe_int(item.get("Alumnos rotatorio", 0))
    st.session_state.alumnos_deslizante = safe_int(item.get("Alumnos deslizante", 0))
    st.session_state.deslizante_lunes = str(item.get("Deslizante lunes", "") or "")
    st.session_state.deslizante_martes = str(item.get("Deslizante martes", "") or "")
    st.session_state.deslizante_miercoles = str(item.get("Deslizante miércoles", "") or "")
    st.session_state.deslizante_jueves = str(item.get("Deslizante jueves", "") or "")
    st.session_state.deslizante_viernes = str(item.get("Deslizante viernes", "") or "")
    st.session_state.observaciones_titulacion = str(item.get(OBS_TITULACION_DISPLAY, "") or "")
    limpiar_detalle_alumnos_widgets()
    detalle = detalle_con_fallback_deslizante(
        item.get(DETALLE_ALUMNOS_DISPLAY, []),
        safe_int(item.get("Alumnos deslizante", 0)),
        item,
    )
    detalle_total = normalizar_detalle_alumnos(detalle, total=safe_int(item.get("Nº alumnos", 0)))
    st.session_state.detalle_alumnos = normalizar_detalle_alumnos(detalle_total)
    preparar_detalle_alumnos_widgets(safe_int(item.get("Nº alumnos", 0)), detalle_total)
    preparar_deslizantes_alumno_widgets(safe_int(item.get("Alumnos deslizante", 0)), detalle_total, item)


def aplicar_edicion_pendiente_en_formulario() -> None:
    """Aplica, antes de renderizar widgets, el registro seleccionado para edición.

    Streamlit no permite modificar st.session_state de una key asociada a un widget
    después de que ese widget se haya instanciado en la misma ejecución. Por eso el
    botón de edición deja el registro en pending_edit_registro, hace rerun, y aquí
    se cargan los valores antes de pintar selectbox/number_input/text_area.
    """
    item = st.session_state.get("pending_edit_registro")
    if not item:
        return
    st.session_state.pending_edit_registro = None
    cargar_registro_en_formulario(item)


def cancelar_edicion_registro() -> None:
    """Sale del modo edición de registro del Paso 4 y limpia el formulario."""
    st.session_state.editing_registro_key = ""
    reset_selectores_estudio()


def normalizar_area(area: str) -> str:
    return LEGACY_AREA_MAP.get(area, area)


def normalizar_unidad_docente(unidad_docente: str) -> str:
    return LEGACY_UNIDAD_DOCENTE_MAP.get(unidad_docente, unidad_docente)


def reset_downstream(level: str) -> None:
    if level == "nivel_i":
        st.session_state.sel_nivel_ii = ""
        st.session_state.sel_rama = ""
        st.session_state.sel_titulacion = ""
    elif level == "nivel_ii":
        st.session_state.sel_rama = ""
        st.session_state.sel_titulacion = ""
        st.session_state.sel_titulacion_libre_catalogo = ""
    elif level == "rama":
        st.session_state.sel_titulacion = ""
        st.session_state.sel_titulacion_libre_catalogo = ""


def safe_code(text: str) -> str:
    text = text or ""
    text = text.upper().strip()
    text = re.sub(r"[^A-Z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text[:60] or "DCD"


def build_codigo_borrador(unidad_docente: str) -> str:
    unidad_docente = normalizar_unidad_docente(unidad_docente)
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


def normalizar_texto_opcion(value: str) -> str:
    return str(value or "").strip().upper().replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")


def es_modo_titulacion_libre(nivel_i: str, nivel_ii: str) -> bool:
    return normalizar_texto_opcion(nivel_i) == "UNIVERSITARIO" and normalizar_texto_opcion(nivel_ii) in {"MASTER", "OTRO", "OTRO (DEFINIR)"}


def aplicar_titulacion_libre_catalogo() -> None:
    seleccion = str(st.session_state.get("sel_titulacion_libre_catalogo", "") or "").strip()
    if seleccion:
        st.session_state.sel_titulacion = seleccion


def safe_int(value, default: int = 0) -> int:
    try:
        if pd.isna(value):
            return default
        return int(value or 0)
    except Exception:
        return default


def suma_turnos_registro(item: dict) -> int:
    return sum(safe_int(item.get(col, 0)) for col in TURNOS_ALUMNOS_DISPLAY)


def registro_deslizante_completo(item: dict) -> bool:
    n_deslizantes = safe_int(item.get("Alumnos deslizante", 0))
    if n_deslizantes <= 0:
        return True
    detalle = detalle_con_fallback_deslizante(item.get(DETALLE_ALUMNOS_DISPLAY, []), n_deslizantes, item)
    deslizantes = deslizantes_desde_detalle(detalle, n_deslizantes)
    if len(deslizantes) >= n_deslizantes and all(entry_tiene_patron_deslizante_completo(e) for e in deslizantes[:n_deslizantes]):
        return True
    return all(str(item.get(col, "") or "").strip() in {"M", "T", "R"} for col in DESLIZANTE_DISPLAY)


def registro_turnos_cuadra(item: dict) -> bool:
    return safe_int(item.get("Nº alumnos", 0)) == suma_turnos_registro(item)


def registro_from_db_row(row: dict) -> dict:
    """Convierte una fila de dcd_registros al formato interno/visual de la app."""
    item = {
        "Nivel Estudio I": row.get("nivel_i", ""),
        "Nivel Estudio II": row.get("nivel_ii", ""),
        "Rama": row.get("rama", ""),
        "Titulación": row.get("titulacion", ""),
        "Nº alumnos": safe_int(row.get("numero_alumnos", 0)),
        "Alumnos mañana": safe_int(row.get("alumnos_manana", 0)),
        "Alumnos tarde": safe_int(row.get("alumnos_tarde", 0)),
        "Alumnos rotatorio": safe_int(row.get("alumnos_rotatorio", 0)),
        "Alumnos deslizante": safe_int(row.get("alumnos_deslizante", 0)),
        "Deslizante lunes": row.get("deslizante_lunes", "") or "",
        "Deslizante martes": row.get("deslizante_martes", "") or "",
        "Deslizante miércoles": row.get("deslizante_miercoles", "") or "",
        "Deslizante jueves": row.get("deslizante_jueves", "") or "",
        "Deslizante viernes": row.get("deslizante_viernes", "") or "",
        OBS_TITULACION_DISPLAY: row.get("observaciones_titulacion", "") or "",
        DETALLE_ALUMNOS_DISPLAY: normalizar_detalle_alumnos(row.get("detalle_alumnos", [])),
    }
    return item


def registro_to_db_extra_fields(item: dict) -> dict:
    out = {db_col: safe_int(item.get(display_col, 0)) for display_col, db_col in TURNOS_DB_MAP.items()}
    for display_col, db_col in DESLIZANTE_DB_MAP.items():
        out[db_col] = str(item.get(display_col, "") or "").strip()
    out[OBS_TITULACION_DB] = str(item.get(OBS_TITULACION_DISPLAY, "") or "").strip()
    out[DETALLE_ALUMNOS_DB] = normalizar_detalle_alumnos(item.get(DETALLE_ALUMNOS_DISPLAY, []))
    return out


def build_turnos_tables(registros_df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Construye hojas/tablas de turnos y observaciones para Excel/PDF de consulta."""
    base_cols = ["Código borrador", "Fecha guardado", "Área", "Centro docente", "Columna Excel", *KEY_COLUMNS, "Nº alumnos"]
    detalle_cols = base_cols + TURNOS_ALUMNOS_DISPLAY + DESLIZANTE_DISPLAY + [OBS_TITULACION_DISPLAY, DETALLE_ALUMNOS_DISPLAY]
    detalle_pdf_cols = ["Centro docente", "Nivel I", "Nivel II", "Rama", "Titulación", "Nº", "Mañana", "Tarde", "Rot.", "Desl.", "Patrón deslizante"]
    obs_pdf_cols = ["Centro docente", "Nivel I", "Nivel II", "Rama", "Titulación", "Nº", "Observaciones"]
    if registros_df is None or registros_df.empty:
        return {
            "Turnos_Detalle": pd.DataFrame(columns=detalle_cols),
            "Turnos_Detalle_PDF": pd.DataFrame(columns=detalle_pdf_cols),
            "Turnos_Resumen": pd.DataFrame(columns=["Turno", "Total alumnos"]),
            "Turnos_Resumen_Centro": pd.DataFrame(columns=["Centro docente", "Alumnos mañana", "Alumnos tarde", "Alumnos rotatorio", "Alumnos deslizante", "Total turnos"]),
            "Turnos_Resumen_Isla": pd.DataFrame(columns=["Isla", "Alumnos mañana", "Alumnos tarde", "Alumnos rotatorio", "Alumnos deslizante", "Total turnos"]),
            "Observaciones_Titulacion": pd.DataFrame(columns=base_cols + [OBS_TITULACION_DISPLAY]),
            "Observaciones_Titulacion_PDF": pd.DataFrame(columns=obs_pdf_cols),
        }

    df = registros_df.copy()
    for col in base_cols + TURNOS_ALUMNOS_DISPLAY + DESLIZANTE_DISPLAY + [OBS_TITULACION_DISPLAY]:
        if col not in df.columns:
            df[col] = 0 if col in TURNOS_ALUMNOS_DISPLAY or col == "Nº alumnos" else ""
    for col in TURNOS_ALUMNOS_DISPLAY + ["Nº alumnos"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    detalle = df[detalle_cols].copy()
    detalle["Total turnos"] = detalle[TURNOS_ALUMNOS_DISPLAY].sum(axis=1)
    obs_mask = detalle[OBS_TITULACION_DISPLAY].astype(str).str.strip() != ""
    detalle = detalle[(detalle["Total turnos"] > 0) | obs_mask].reset_index(drop=True)

    resumen_turnos = pd.DataFrame([
        {"Turno": "Mañana", "Total alumnos": int(df["Alumnos mañana"].sum())},
        {"Turno": "Tarde", "Total alumnos": int(df["Alumnos tarde"].sum())},
        {"Turno": "Rotatorio", "Total alumnos": int(df["Alumnos rotatorio"].sum())},
        {"Turno": "Deslizante", "Total alumnos": int(df["Alumnos deslizante"].sum())},
    ])
    resumen_turnos = resumen_turnos[resumen_turnos["Total alumnos"] > 0].reset_index(drop=True)

    resumen_centro = pd.DataFrame(columns=["Centro docente", *TURNOS_ALUMNOS_DISPLAY, "Total turnos"])
    if "Centro docente" in df.columns:
        resumen_centro = df.groupby("Centro docente", dropna=False)[TURNOS_ALUMNOS_DISPLAY].sum().reset_index()
        resumen_centro["Total turnos"] = resumen_centro[TURNOS_ALUMNOS_DISPLAY].sum(axis=1)
        resumen_centro = resumen_centro[resumen_centro["Total turnos"] > 0].sort_values("Total turnos", ascending=False).reset_index(drop=True)

    isla_por_columna = {}
    for isla, columnas in ISLA_COLUMNAS.items():
        for col in columnas:
            isla_por_columna[col] = isla
    df["Isla"] = df.get("Columna Excel", "").map(isla_por_columna).fillna("") if "Columna Excel" in df.columns else ""
    resumen_isla = pd.DataFrame(columns=["Isla", *TURNOS_ALUMNOS_DISPLAY, "Total turnos"])
    if "Isla" in df.columns:
        tmp = df[df["Isla"].astype(str).str.strip() != ""].copy()
        if not tmp.empty:
            resumen_isla = tmp.groupby("Isla", dropna=False)[TURNOS_ALUMNOS_DISPLAY].sum().reset_index()
            resumen_isla["Total turnos"] = resumen_isla[TURNOS_ALUMNOS_DISPLAY].sum(axis=1)
            resumen_isla = resumen_isla[resumen_isla["Total turnos"] > 0].sort_values("Total turnos", ascending=False).reset_index(drop=True)

    obs = df[df[OBS_TITULACION_DISPLAY].astype(str).str.strip() != ""][base_cols + [OBS_TITULACION_DISPLAY]].copy()

    detalle_pdf = detalle.copy()
    if not detalle_pdf.empty:
        def _patron_deslizante(row) -> str:
            n_deslizantes = safe_int(row.get("Alumnos deslizante", 0))
            detalle_txt = formatear_patrones_deslizantes_detalle(row.get(DETALLE_ALUMNOS_DISPLAY, []), n_deslizantes)
            if detalle_txt:
                return detalle_txt
            valores = [str(row.get(col, "") or "").strip() for col in DESLIZANTE_DISPLAY]
            if not any(valores):
                return ""
            dias = ["L", "M", "X", "J", "V"]
            return " / ".join(f"{dia}:{valor}" for dia, valor in zip(dias, valores) if valor)

        detalle_pdf["Patrón deslizante"] = detalle_pdf.apply(_patron_deslizante, axis=1)
        detalle_pdf = detalle_pdf[["Centro docente", "Nivel Estudio I", "Nivel Estudio II", "Rama", "Titulación", "Nº alumnos", *TURNOS_ALUMNOS_DISPLAY, "Patrón deslizante"]].copy()
        detalle_pdf = detalle_pdf.rename(columns={
            "Nivel Estudio I": "Nivel I",
            "Nivel Estudio II": "Nivel II",
            "Nº alumnos": "Nº",
            "Alumnos rotatorio": "Rot.",
            "Alumnos deslizante": "Desl.",
        })

    obs_pdf = obs.copy()
    if not obs_pdf.empty:
        obs_pdf = obs_pdf[["Centro docente", "Nivel Estudio I", "Nivel Estudio II", "Rama", "Titulación", "Nº alumnos", OBS_TITULACION_DISPLAY]].copy()
        obs_pdf = obs_pdf.rename(columns={
            "Nivel Estudio I": "Nivel I",
            "Nivel Estudio II": "Nivel II",
            "Nº alumnos": "Nº",
            OBS_TITULACION_DISPLAY: "Observaciones",
        })

    return {
        "Turnos_Detalle": detalle.drop(columns=[DETALLE_ALUMNOS_DISPLAY], errors="ignore"),
        "Turnos_Detalle_PDF": detalle_pdf.reset_index(drop=True),
        "Turnos_Resumen": resumen_turnos,
        "Turnos_Resumen_Centro": resumen_centro,
        "Turnos_Resumen_Isla": resumen_isla,
        "Observaciones_Titulacion": obs.reset_index(drop=True),
        "Observaciones_Titulacion_PDF": obs_pdf.reset_index(drop=True),
    }


def build_detalle_alumnos_table(registros_df: pd.DataFrame) -> pd.DataFrame:
    """Construye la hoja Detalle_Alumnos para Excel admin/consulta.

    Desde DCD 1.2.2 estable incluye, además de Servicio y Curso/año,
    el patrón semanal individual de cada alumno en turno deslizante.
    """
    cols = [
        "Código borrador", "Fecha guardado", "Área", "Centro docente", "Usuario aportación",
        "Nivel Estudio I", "Nivel Estudio II", "Rama", "Titulación",
        "Alumno", "Servicio", "Curso/año", "Es deslizante",
        "Deslizante lunes", "Deslizante martes", "Deslizante miércoles", "Deslizante jueves", "Deslizante viernes",
        "Patrón deslizante",
        "Nº alumnos titulación",
        "Alumnos mañana", "Alumnos tarde", "Alumnos rotatorio", "Alumnos deslizante",
        OBS_TITULACION_DISPLAY,
    ]
    if registros_df is None or registros_df.empty:
        return pd.DataFrame(columns=cols)

    rows = []
    for _, row in registros_df.iterrows():
        item_dict = row.to_dict()
        n_deslizantes = safe_int(row.get("Alumnos deslizante", 0))
        detalle = detalle_con_fallback_deslizante(row.get(DETALLE_ALUMNOS_DISPLAY, []), n_deslizantes, item_dict)
        if not detalle:
            continue
        for entry in detalle:
            servicio = str(entry.get("servicio", "") or "").strip()
            curso = str(entry.get("curso", "") or "").strip()
            es_deslizante = bool(entry.get("es_deslizante"))
            patron = patron_deslizante_entry(entry)
            patron_txt = formatear_patron_deslizante_entry(entry)
            if not servicio and not curso and not es_deslizante:
                continue
            rows.append({
                "Código borrador": row.get("Código borrador", ""),
                "Fecha guardado": row.get("Fecha guardado", ""),
                "Área": row.get("Área", ""),
                "Centro docente": row.get("Centro docente", ""),
                "Usuario aportación": row.get("Usuario aportación", row.get("Usuario", "")),
                "Nivel Estudio I": row.get("Nivel Estudio I", ""),
                "Nivel Estudio II": row.get("Nivel Estudio II", ""),
                "Rama": row.get("Rama", ""),
                "Titulación": row.get("Titulación", ""),
                "Alumno": safe_int(entry.get("alumno", 0)),
                "Servicio": servicio,
                "Curso/año": curso,
                "Es deslizante": "Sí" if es_deslizante else "No",
                "Deslizante lunes": patron["deslizante_lunes"],
                "Deslizante martes": patron["deslizante_martes"],
                "Deslizante miércoles": patron["deslizante_miercoles"],
                "Deslizante jueves": patron["deslizante_jueves"],
                "Deslizante viernes": patron["deslizante_viernes"],
                "Patrón deslizante": patron_txt,
                "Nº alumnos titulación": safe_int(row.get("Nº alumnos", 0)),
                "Alumnos mañana": safe_int(row.get("Alumnos mañana", 0)),
                "Alumnos tarde": safe_int(row.get("Alumnos tarde", 0)),
                "Alumnos rotatorio": safe_int(row.get("Alumnos rotatorio", 0)),
                "Alumnos deslizante": n_deslizantes,
                OBS_TITULACION_DISPLAY: row.get(OBS_TITULACION_DISPLAY, ""),
            })
    if not rows:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame(rows, columns=cols)


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

    # DCD 1.2.2 estable: opciones universitarias adicionales con titulación libre.
    # Se añaden al catálogo en memoria para que aparezcan en los selectores sin modificar el Excel base.
    extra_rows = []
    for nivel_ii_extra in ["Máster", "Otro"]:
        for rama_extra in ["Sanidad", "Otro (Definir)"]:
            extra_rows.append({
                "Nivel Estudio I": "Universitario",
                "Nivel Estudio II": nivel_ii_extra,
                "Rama": rama_extra,
                "Titulación": "Titulación libre",
            })
    if extra_rows:
        df = pd.concat([df, pd.DataFrame(extra_rows)], ignore_index=True)

    df = df.drop_duplicates(subset=KEY_COLUMNS).reset_index(drop=True)
    return df


def sorted_unique(series: pd.Series) -> list[str]:
    values = [str(x).strip() for x in series.dropna().unique() if str(x).strip()]
    return sorted(values, key=lambda x: x.upper())


def obtener_titulaciones_libres_usadas(nivel_i: str, nivel_ii: str, rama: str) -> list[str]:
    """Recupera titulaciones libres ya usadas para reutilizarlas como sugerencia.

    No requiere tabla nueva: consulta los registros ya guardados y añade los valores
    del expediente en edición.
    """
    values = set()
    for item in st.session_state.get("registros", {}).values():
        if (
            str(item.get("Nivel Estudio I", "")).strip() == str(nivel_i).strip()
            and str(item.get("Nivel Estudio II", "")).strip() == str(nivel_ii).strip()
            and str(item.get("Rama", "")).strip() == str(rama).strip()
        ):
            val = str(item.get("Titulación", "") or "").strip()
            if val and val != "Titulación libre":
                values.add(val)

    client = get_supabase_client()
    if client is not None and nivel_i and nivel_ii and rama:
        try:
            resp = (
                client.table("dcd_registros")
                .select("titulacion")
                .eq("nivel_i", nivel_i)
                .eq("nivel_ii", nivel_ii)
                .eq("rama", rama)
                .limit(1000)
                .execute()
            )
            for row in getattr(resp, "data", []) or []:
                val = str(row.get("titulacion", "") or "").strip()
                if val and val != "Titulación libre":
                    values.add(val)
        except Exception:
            pass

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


def is_consulta_user() -> bool:
    """Compatibilidad para descargas limitadas del rol consulta/visor externo."""
    return is_external_viewer()


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
        st.session_state.current_user_unidad = normalizar_unidad_docente(str(db_user.get("unidad_docente", "") or ""))
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
    st.session_state.current_user_unidad = normalizar_unidad_docente(str(user_data.get("unidad_docente", "") or ""))
    st.session_state.current_user_codigo_unidad = str(user_data.get("codigo_unidad", "") or "")
    st.session_state.must_change_password = False
    return True, "Acceso correcto."


def audit_event(action: str, detail: str = "", codigo_borrador: str = "") -> bool:
    """
    Registra una acción en Supabase si la tabla dcd_auditoria existe.
    La auditoría no debe bloquear el uso del aplicativo si falla, pero deja
    trazado el último error para que el administrador pueda diagnosticarlo.
    """
    client = get_supabase_client()
    if client is None:
        st.session_state["last_audit_error"] = "La Base de Datos no está configurada."
        return False

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
        st.session_state["last_audit_error"] = ""
        return True
    except Exception as exc:
        st.session_state["last_audit_error"] = str(exc)
        return False


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
        return False, "El servidor de correo no está configurado. Puede descargar el Excel y enviarlo manualmente."

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
            return True, "Correo enviado correctamente a servicio de FSE"
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
            "usuario_aportacion": st.session_state.get("current_user", ""),
            "centro_multiusuario": is_centro_multiusuario_activo(unidad),
            "codigo_consolidado_centro": "",
            "aportacion_finalizada": estado == "finalizado",
            "fecha_finalizacion_aportacion": datetime.now().isoformat() if estado == "finalizado" else None,
            "nivel_i": item["Nivel Estudio I"],
            "nivel_ii": item["Nivel Estudio II"],
            "rama": item["Rama"],
            "titulacion": item["Titulación"],
            "numero_alumnos": int(item["Nº alumnos"]),
            **registro_to_db_extra_fields(item),
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
        return False, "La Base de Datos no está configurada todavía. Puede seguir usando la descarga del Excel."

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
            "usuario_aportacion": st.session_state.get("current_user", ""),
            "centro_multiusuario": is_centro_multiusuario_activo(unidad),
            "codigo_consolidado_centro": "",
            "aportacion_finalizada": estado == "finalizado",
            "fecha_finalizacion_aportacion": saved_at if estado == "finalizado" else None,
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
            maybe_auto_publish_if_complete()
            return True, f"Expediente finalizado correctamente como nueva versión: {codigo_borrador}."
        return True, f"Borrador guardado correctamente como nueva versión: {codigo_borrador}"
    except Exception as exc:
        return False, f"Error al guardar en Base de Datos: {exc}"


def load_draft_from_supabase(codigo_borrador: str) -> tuple[bool, str]:
    client = get_supabase_client()
    if client is None:
        return False, "La Base de Datos no está configurada."

    try:
        borrador_resp = client.table("dcd_borradores").select("*").eq("codigo_borrador", codigo_borrador).limit(1).execute()
        borradores = getattr(borrador_resp, "data", []) or []
        if not borradores:
            return False, "No se encontró ese código de borrador."

        borrador = borradores[0]
        if not is_admin() and user_scope_unidad() and borrador.get("unidad_docente", "") != user_scope_unidad():
            return False, "No puede cargar un borrador de un centro docente distinto al asignado a su usuario."
        if not is_admin() and str(borrador.get("usuario_propietario", "")).strip().lower() != str(st.session_state.get("current_user", "")).strip().lower():
            return False, "Este borrador pertenece a otro usuario del centro. Cada usuario debe cargar y finalizar su propia aportación."

        registros_resp = client.table("dcd_registros").select("*").eq("codigo_borrador", codigo_borrador).execute()
        rows = getattr(registros_resp, "data", []) or []

        st.session_state.area_selected = normalizar_area(borrador.get("area", ""))
        st.session_state.direccion_selected = normalizar_unidad_docente(borrador.get("unidad_docente", ""))
        st.session_state.codigo_borrador = codigo_borrador
        st.session_state.codigo_expediente = borrador.get("codigo_expediente") or build_codigo_borrador(st.session_state.direccion_selected)
        st.session_state.version_num = int(borrador.get("version_num") or 1)
        st.session_state.draft_estado = borrador.get("estado", "borrador") or "borrador"
        st.session_state.permitir_editar_finalizado = False
        st.session_state.observaciones = borrador.get("observaciones", "") or ""
        st.session_state.registros = {}

        for row in rows:
            item = registro_from_db_row(row)
            key = registro_key(item["Nivel Estudio I"], item["Nivel Estudio II"], item["Rama"], item["Titulación"])
            st.session_state.registros[key] = item

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
        if not is_admin():
            query = query.eq("usuario_propietario", st.session_state.get("current_user", ""))
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
        return False, "La Base de Datos no está configurada."

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
        return False, "La Base de Datos no está configurada."
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
        return False, "La Base de Datos no está configurada."
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
        return False, "La Base de Datos no está configurada."
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
        turnos_tables = build_turnos_tables(registros_df)
        detalle_alumnos = build_detalle_alumnos_table(registros_df)
        registros_df_excel = dataframe_excel_safe(registros_df)
        registros_df_excel.to_excel(writer, sheet_name="Registros_DCD", index=False)
        turnos_tables["Turnos_Detalle"].to_excel(writer, sheet_name="Turnos_Detalle", index=False)
        turnos_tables["Observaciones_Titulacion"].to_excel(writer, sheet_name="Observaciones_Titulacion", index=False)
        detalle_alumnos.to_excel(writer, sheet_name="Detalle_Alumnos", index=False)
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

        local_sheets = {
            "Resumen": resumen,
            "Registros_DCD": registros_df_excel,
            "Turnos_Detalle": turnos_tables["Turnos_Detalle"],
            "Observaciones_Titulacion": turnos_tables["Observaciones_Titulacion"],
            "Detalle_Alumnos": detalle_alumnos,
            "Matriz_DCD": matriz,
        }
        for sheet_name, df_sheet in local_sheets.items():
            ws = writer.sheets[sheet_name]
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
    """Devuelve los expedientes finalizados que deben entrar en el consolidado.

    Regla DCD 1.2.1 beta 4:
    - Centros ordinarios: se mantiene el criterio histórico de última versión finalizada por centro.
    - Centros multiusuario activos: se toma la última aportación finalizada de cada usuario activo asignado,
      pero solo cuando todos los usuarios activos asignados han finalizado. Si falta algún usuario,
      el centro queda pendiente y no entra todavía en el consolidado general.
    """
    client = get_supabase_client()
    if client is None:
        return [], []

    resp = client.table("dcd_borradores").select("*").eq("estado", "finalizado").order("saved_at", desc=True).order("updated_at", desc=True).execute()
    borradores = getattr(resp, "data", []) or []

    configs = {c.get("centro_docente", ""): c for c in list_multiusuario_configs() if c.get("multiusuario_activo")}
    multi_centros = set(configs.keys())
    active_users_by_centro: dict[str, list[str]] = {}
    for centro in multi_centros:
        active_users_by_centro[centro] = [
            str(a.get("username", "")).strip().lower()
            for a in list_multiusuario_assignments(centro)
            if a.get("activo") and a.get("username")
        ]

    latest_by_unit: dict[str, dict] = {}
    latest_by_multi_user: dict[tuple[str, str], dict] = {}
    duplicates: list[dict] = []

    for borrador in borradores:
        unidad = borrador.get("unidad_docente", "")
        usuario = str(borrador.get("usuario_propietario", "") or borrador.get("usuario_aportacion", "")).strip().lower()
        if not unidad:
            continue

        if unidad in multi_centros:
            active_users = active_users_by_centro.get(unidad, [])
            if active_users and usuario not in active_users:
                duplicates.append(borrador)
                continue
            key = (unidad, usuario)
            if key not in latest_by_multi_user:
                latest_by_multi_user[key] = borrador
            else:
                duplicates.append(borrador)
        else:
            if unidad not in latest_by_unit:
                latest_by_unit[unidad] = borrador
            else:
                duplicates.append(borrador)

    selected = list(latest_by_unit.values())

    for centro, cfg in configs.items():
        active_users = active_users_by_centro.get(centro, [])
        expected = int(cfg.get("usuarios_previstos") or len(active_users) or 1)
        if not active_users:
            continue
        finalized_users = [u for u in active_users if (centro, u) in latest_by_multi_user]
        # Para garantizar que el consolidado general no incluye centros incompletos, exigimos que finalicen
        # todos los usuarios activos asignados y que se cubra al menos el número previsto por el admin.
        if len(finalized_users) >= expected and all(u in finalized_users for u in active_users):
            selected.extend([latest_by_multi_user[(centro, u)] for u in active_users])
        else:
            partial_auth = get_latest_multiusuario_consolidation(centro)
            if partial_auth and bool(partial_auth.get("consolidacion_parcial", False)) and finalized_users:
                # DCD 1.2.1 beta 4: el admin puede autorizar que entren las aportaciones finalizadas
                # aunque falten usuarios, dejando trazabilidad de consolidación parcial.
                selected.extend([latest_by_multi_user[(centro, u)] for u in finalized_users])
            else:
                # El centro queda pendiente. Las aportaciones finalizadas se conservan, pero no entran todavía.
                for u in finalized_users:
                    duplicates.append(latest_by_multi_user[(centro, u)])

    return selected, duplicates


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

    configs = {c.get("centro_docente", ""): c for c in list_multiusuario_configs() if c.get("multiusuario_activo")}

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

        if centro in configs:
            cfg = configs.get(centro, {})
            assignments = [a for a in list_multiusuario_assignments(centro) if a.get("activo") and a.get("username")]
            active_users = [str(a.get("username", "")).strip().lower() for a in assignments]
            expected_users = int(cfg.get("usuarios_previstos") or len(active_users) or 1)
            finalized_by_user = {}
            for v in finalized:
                user = str(v.get("usuario_propietario", "") or v.get("usuario_aportacion", "")).strip().lower()
                if user and user in active_users and user not in finalized_by_user:
                    finalized_by_user[user] = v
            done = len(finalized_by_user)
            complete = bool(active_users) and done >= expected_users and all(u in finalized_by_user for u in active_users)
            partial_auth = get_latest_multiusuario_consolidation(centro)
            if complete:
                estado = f"Multiusuario finalizado ({done}/{expected_users})"
                entra = "Sí"
            elif partial_auth and bool(partial_auth.get("consolidacion_parcial", False)) and done > 0:
                estado = f"Multiusuario consolidado parcial ({done}/{expected_users})"
                entra = "Sí (parcial)"
            elif versions:
                estado = f"Multiusuario pendiente ({done}/{expected_users})"
                entra = "No"
                missing.append(centro)
            else:
                estado = f"Multiusuario sin datos (0/{expected_users})"
                entra = "No"
                missing.append(centro)
            latest_finalized = max(finalized_by_user.values(), key=lambda x: x.get("saved_at") or x.get("updated_at") or "", default={}) if finalized_by_user else {}
            rows.append({
                **item,
                "Estado": estado,
                "Última versión guardada": latest_any.get("codigo_borrador", ""),
                "Fecha último guardado": latest_any.get("saved_at") or latest_any.get("updated_at", ""),
                "Último finalizado": latest_finalized.get("codigo_borrador", ""),
                "Fecha último finalizado": latest_finalized.get("saved_at") or latest_finalized.get("updated_at", ""),
                "Entra en consolidado": entra,
                "Versiones guardadas": len(versions),
                "Usuarios previstos": expected_users,
                "Usuarios finalizados": done,
            })
            continue

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
        return False, "El servidor de correo no está configurado. Puede revisar el listado de pendientes en pantalla."

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
    # Seguridad: si la matriz procede de un Excel ya publicado, puede traer una fila TOTAL.
    # La retiramos antes de recalcular para evitar duplicar/sumar dos veces el total.
    if "Titulación" in out.columns:
        out = out[out["Titulación"].astype(str).str.upper().str.strip() != "TOTAL"].copy()
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


def filtrar_matriz_total_positivo(matriz: pd.DataFrame, mantener_fila_total: bool = False) -> pd.DataFrame:
    """
    Devuelve una copia calculada de Matriz_DCD sin filas de titulaciones con Total = 0.

    Es un filtro solo de presentación para dashboard/PDF: no modifica registros,
    borradores, publicaciones, cálculos de consolidación ni la matriz base.
    """
    df = calcular_totales_matriz_df(matriz)
    if "Total" not in df.columns:
        return df.reset_index(drop=True)

    total_num = pd.to_numeric(df["Total"], errors="coerce").fillna(0)
    is_total_row = pd.Series(False, index=df.index)
    if "Titulación" in df.columns:
        is_total_row = df["Titulación"].astype(str).str.upper().str.strip().eq("TOTAL")

    mask = total_num > 0
    if mantener_fila_total:
        mask = mask | is_total_row
    else:
        mask = mask & ~is_total_row
    return df[mask].reset_index(drop=True)


def filtrar_resumen_total_plazas_positivo(df: pd.DataFrame) -> pd.DataFrame:
    """Oculta filas analíticas con Total plazas = 0 para mejorar la visualización del portal de consulta."""
    if df is None or df.empty or "Total plazas" not in df.columns:
        return df
    out = df.copy()
    totals = pd.to_numeric(out["Total plazas"], errors="coerce").fillna(0)
    return out[totals > 0].reset_index(drop=True)


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

    # Limpieza visual: las filas con Total = 0 no aportan capacidad docente
    # y se ocultan en dashboard/PDF sin modificar la matriz oficial ni la base de datos.
    if "Total" in df.columns:
        df = df[pd.to_numeric(df["Total"], errors="coerce").fillna(0) > 0].reset_index(drop=True)

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
        return False, "La Base de Datos no está configurada.", None
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
                item = registro_from_db_row(row)
                volcar_registro_en_matriz(matriz, item, columna_excel)
        return True, "Analítica generada.", build_analytics_tables(matriz, status_df)
    except Exception as exc:
        return False, f"Error al generar analítica: {exc}", None



# =========================================================
# CALIDAD DE DATOS / VALIDACIONES
# =========================================================
def build_current_dataset_from_supabase() -> tuple[bool, str, pd.DataFrame | None, pd.DataFrame | None, pd.DataFrame | None, list[dict], list[dict]]:
    """Construye la matriz y registros consolidados con la última versión finalizada por centro."""
    client = get_supabase_client()
    if client is None:
        return False, "La Base de Datos no está configurada.", None, None, None, [], []
    try:
        borradores, duplicados = get_latest_finalized_drafts()
        status_df, _missing = get_admin_centros_status()
        if not borradores:
            return False, "No hay expedientes finalizados.", None, pd.DataFrame(), status_df, [], duplicados

        matriz = limpiar_columnas_matriz(load_catalogo())
        registros_consolidados = []

        for borrador in borradores:
            codigo = borrador.get("codigo_borrador", "")
            unidad = borrador.get("unidad_docente", "")
            columna_excel = COLUMNA_EXCEL_POR_DIRECCION.get(unidad, "")
            if not codigo:
                continue
            resp = client.table("dcd_registros").select("*").eq("codigo_borrador", codigo).execute()
            rows = getattr(resp, "data", []) or []
            for row in rows:
                item = registro_from_db_row(row)
                volcar_registro_en_matriz(matriz, item, columna_excel)
                registros_consolidados.append({
                    "Código borrador": codigo,
                    "Estado": borrador.get("estado", ""),
                    "Fecha guardado": borrador.get("saved_at") or borrador.get("updated_at", ""),
                    "Área": normalizar_area(borrador.get("area", "")),
                    "Centro docente": unidad,
                    "Columna Excel": columna_excel,
                    **item,
                })

        registros_df = pd.DataFrame(registros_consolidados)
        return True, "Dataset generado.", matriz, registros_df, status_df, borradores, duplicados
    except Exception as exc:
        return False, f"Error al construir dataset de calidad: {exc}", None, None, None, [], []


def build_quality_tables(matriz: pd.DataFrame, registros_df: pd.DataFrame, status_df: pd.DataFrame, duplicados: list[dict] | None = None) -> dict[str, pd.DataFrame]:
    """Genera tablas de calidad de datos para revisión administrativa."""
    duplicados = duplicados or []
    if matriz is None or matriz.empty:
        matriz_calc = pd.DataFrame()
    else:
        matriz_calc = matriz_sin_total(matriz)
        for col in MATRIX_VALUE_COLUMNS:
            if col in matriz_calc.columns:
                matriz_calc[col] = pd.to_numeric(matriz_calc[col], errors="coerce").fillna(0).astype(int)

    if registros_df is None:
        registros_df = pd.DataFrame()
    if status_df is None:
        status_df = pd.DataFrame()

    pendientes = pd.DataFrame()
    if not status_df.empty and "Estado" in status_df.columns:
        pendientes = status_df[~status_df["Estado"].astype(str).str.startswith("Finalizado")].copy()

    borrador_posterior = pd.DataFrame()
    if not status_df.empty and "Estado" in status_df.columns:
        borrador_posterior = status_df[status_df["Estado"].astype(str).str.contains("borrador posterior", case=False, na=False)].copy()

    duplicados_titulacion = pd.DataFrame(columns=["Centro docente", *KEY_COLUMNS, "Nº líneas", "Total alumnos"])
    if not registros_df.empty and all(c in registros_df.columns for c in ["Centro docente", *KEY_COLUMNS, "Nº alumnos"]):
        duplicados_titulacion = (
            registros_df.groupby(["Centro docente", *KEY_COLUMNS], dropna=False)
            .agg(**{"Nº líneas": ("Nº alumnos", "size"), "Total alumnos": ("Nº alumnos", "sum")})
            .reset_index()
        )
        duplicados_titulacion = duplicados_titulacion[duplicados_titulacion["Nº líneas"] > 1].sort_values(["Centro docente", "Nº líneas"], ascending=[True, False])

    valores_altos = pd.DataFrame(columns=list(registros_df.columns) + ["Motivo alerta"] if not registros_df.empty else ["Motivo alerta"])
    if not registros_df.empty and "Nº alumnos" in registros_df.columns:
        tmp = registros_df.copy()
        tmp["Nº alumnos"] = pd.to_numeric(tmp["Nº alumnos"], errors="coerce").fillna(0).astype(int)
        positivos = tmp[tmp["Nº alumnos"] > 0]
        p95 = int(positivos["Nº alumnos"].quantile(0.95)) if not positivos.empty else 0
        threshold = max(100, p95 * 2) if p95 else 100
        valores_altos = tmp[tmp["Nº alumnos"] >= threshold].copy()
        if not valores_altos.empty:
            valores_altos["Motivo alerta"] = f"Nº alumnos >= {threshold}. Revisar si es correcto."

    titulaciones_sin_plazas = pd.DataFrame(columns=KEY_COLUMNS + ["Total plazas"])
    if not matriz_calc.empty and all(c in matriz_calc.columns for c in KEY_COLUMNS + ["Total"]):
        titulaciones_sin_plazas = matriz_calc[matriz_calc["Total"] == 0][KEY_COLUMNS + ["Total"]].copy()
        titulaciones_sin_plazas = titulaciones_sin_plazas.rename(columns={"Total": "Total plazas"}).head(500)

    centros_finalizados_sin_registros = pd.DataFrame(columns=["Centro docente", "Estado", "Último finalizado"])
    if not status_df.empty and not registros_df.empty and "Centro docente" in registros_df.columns:
        centros_con_registros = set(registros_df["Centro docente"].dropna().astype(str).unique().tolist())
        if "Centro docente" in status_df.columns and "Estado" in status_df.columns:
            centros_finalizados_sin_registros = status_df[
                status_df["Estado"].astype(str).str.startswith("Finalizado")
                & ~status_df["Centro docente"].astype(str).isin(centros_con_registros)
            ].copy()
    elif not status_df.empty and registros_df.empty and "Estado" in status_df.columns:
        centros_finalizados_sin_registros = status_df[status_df["Estado"].astype(str).str.startswith("Finalizado")].copy()

    resumen = pd.DataFrame([
        {"Control": "Centros pendientes/sin finalizar", "Resultado": len(pendientes), "Nivel": "Aviso" if len(pendientes) else "OK"},
        {"Control": "Centros finalizados con borrador posterior", "Resultado": len(borrador_posterior), "Nivel": "Aviso" if len(borrador_posterior) else "OK"},
        {"Control": "Centros finalizados sin registros", "Resultado": len(centros_finalizados_sin_registros), "Nivel": "Alerta" if len(centros_finalizados_sin_registros) else "OK"},
        {"Control": "Titulaciones duplicadas dentro del mismo centro", "Resultado": len(duplicados_titulacion), "Nivel": "Revisar" if len(duplicados_titulacion) else "OK"},
        {"Control": "Registros con valores altos", "Resultado": len(valores_altos), "Nivel": "Revisar" if len(valores_altos) else "OK"},
        {"Control": "Titulaciones del catálogo sin plazas", "Resultado": len(titulaciones_sin_plazas), "Nivel": "Informativo"},
        {"Control": "Expedientes finalizados duplicados ignorados", "Resultado": len(duplicados), "Nivel": "Aviso" if len(duplicados) else "OK"},
    ])

    return {
        "Calidad_Resumen": resumen,
        "Calidad_Pendientes": pendientes,
        "Calidad_Borrador_Posterior": borrador_posterior,
        "Calidad_Finalizados_0reg": centros_finalizados_sin_registros,
        "Calidad_Duplicados": duplicados_titulacion,
        "Calidad_Valores_Altos": valores_altos,
        "Calidad_Sin_Plazas": titulaciones_sin_plazas,
    }


def build_quality_from_supabase() -> tuple[bool, str, dict[str, pd.DataFrame] | None]:
    ok, msg, matriz, registros_df, status_df, _borradores, duplicados = build_current_dataset_from_supabase()
    if not ok or matriz is None or registros_df is None or status_df is None:
        return False, msg, None
    return True, "Validaciones de calidad generadas.", build_quality_tables(matriz, registros_df, status_df, duplicados)


def compare_latest_publications() -> tuple[bool, str, dict[str, pd.DataFrame] | None]:
    pubs = get_publicaciones()
    if len(pubs) < 2:
        return False, "No hay al menos dos publicaciones para comparar.", None
    pub_actual, pub_anterior = pubs[0], pubs[1]
    try:
        ok1, data1, msg1 = download_publication_file(pub_actual.get("ruta_excel", ""))
        ok2, data2, msg2 = download_publication_file(pub_anterior.get("ruta_excel", ""))
        if not ok1 or not ok2 or not data1 or not data2:
            return False, f"No se pudieron descargar ambas publicaciones. Actual: {msg1}. Anterior: {msg2}", None
        x1 = pd.ExcelFile(io.BytesIO(data1))
        x2 = pd.ExcelFile(io.BytesIO(data2))
        if "Matriz_DCD" not in x1.sheet_names or "Matriz_DCD" not in x2.sheet_names:
            return False, "Alguna publicación no contiene hoja Matriz_DCD.", None
        m1 = pd.read_excel(x1, sheet_name="Matriz_DCD")
        m2 = pd.read_excel(x2, sheet_name="Matriz_DCD")
        m1 = matriz_sin_total(m1)
        m2 = matriz_sin_total(m2)
        for df in (m1, m2):
            if "Total" in df.columns:
                df["Total"] = pd.to_numeric(df["Total"], errors="coerce").fillna(0).astype(int)
        cols = [c for c in KEY_COLUMNS if c in m1.columns and c in m2.columns]
        if not cols or "Total" not in m1.columns or "Total" not in m2.columns:
            return False, "No se pudieron localizar columnas clave o Total para comparar.", None
        a = m1[cols + ["Total"]].rename(columns={"Total": "Total actual"})
        b = m2[cols + ["Total"]].rename(columns={"Total": "Total anterior"})
        comp = a.merge(b, on=cols, how="outer").fillna(0)
        comp["Total actual"] = pd.to_numeric(comp["Total actual"], errors="coerce").fillna(0).astype(int)
        comp["Total anterior"] = pd.to_numeric(comp["Total anterior"], errors="coerce").fillna(0).astype(int)
        comp["Diferencia"] = comp["Total actual"] - comp["Total anterior"]
        cambios = comp[comp["Diferencia"] != 0].copy().sort_values("Diferencia", key=lambda s: s.abs(), ascending=False)
        resumen = pd.DataFrame([
            {"Indicador": "Publicación actual", "Valor": pub_actual.get("codigo_publicacion", "")},
            {"Indicador": "Publicación anterior", "Valor": pub_anterior.get("codigo_publicacion", "")},
            {"Indicador": "Total actual", "Valor": int(comp["Total actual"].sum())},
            {"Indicador": "Total anterior", "Valor": int(comp["Total anterior"].sum())},
            {"Indicador": "Diferencia total", "Valor": int(comp["Diferencia"].sum())},
            {"Indicador": "Titulaciones con cambios", "Valor": int(len(cambios))},
        ])
        return True, "Comparativa generada.", {
            "Comparativa_Resumen": resumen,
            "Comparativa_Cambios": cambios.head(200),
        }
    except Exception as exc:
        return False, f"Error al comparar publicaciones: {exc}", None


def render_canarias_capacity_map(analytics: dict[str, pd.DataFrame]) -> None:
    """
    Muestra un mapa visual de Canarias con la capacidad docente total por isla.

    Es un bloque exclusivamente visual para el rol consulta. No modifica datos,
    cálculos, publicaciones, Excel ni PDF.
    """
    isla_df = analytics.get("Resumen_Isla", pd.DataFrame())
    if isla_df is None or isla_df.empty or "Isla" not in isla_df.columns or "Total plazas" not in isla_df.columns:
        return
    if not MAPA_CANARIAS_PATH.exists():
        return

    valores = {}
    for _, row in isla_df.iterrows():
        isla = str(row.get("Isla", "")).strip()
        try:
            valor = int(pd.to_numeric(row.get("Total plazas", 0), errors="coerce"))
        except Exception:
            valor = 0
        valores[isla] = valor

    # Coordenadas en porcentaje sobre la imagen base mapa_canarias.png (1600 x 912).
    # DCD 1.1.3.9: microajuste final adicional del recuadro de Fuerteventura.
    # Solo se modifican posiciones visuales del dashboard de consulta; no afecta a datos ni cálculos.
    puntos = [
        {"isla": "La Palma", "x": 12.0, "y": 32.0, "dot_x": 12.9, "dot_y": 45.0},
        {"isla": "Tenerife", "x": 40.2, "y": 44.2, "dot_x": 41.5, "dot_y": 55.0},
        {"isla": "La Gomera", "x": 21.8, "y": 57.5, "dot_x": 24.0, "dot_y": 68.2},
        {"isla": "El Hierro", "x": 9.0, "y": 72.0, "dot_x": 9.5, "dot_y": 82.0},
        {"isla": "Gran Canaria", "x": 55.0, "y": 61.5, "dot_x": 55.0, "dot_y": 71.5},
        {"isla": "Fuerteventura", "x": 75.6, "y": 49.8, "dot_x": 83.2, "dot_y": 54.8},
        {"isla": "Lanzarote", "x": 91.0, "y": 18.5, "dot_x": 92.3, "dot_y": 30.5},
    ]

    try:
        encoded_map = base64.b64encode(MAPA_CANARIAS_PATH.read_bytes()).decode("utf-8")
    except Exception:
        return

    labels_html = []
    for punto in puntos:
        isla = punto["isla"]
        valor = valores.get(isla, 0)
        labels_html.append(f"""
            <div class="map-label" style="left:{punto['x']}%; top:{punto['y']}%;">
                <div class="island-name">{isla}</div>
                <div class="island-value">{valor}</div>
            </div>
            <div class="map-dot" style="left:{punto['dot_x']}%; top:{punto['dot_y']}%;"></div>
        """)

    legend_html = []
    for isla in ["La Palma", "Tenerife", "Gran Canaria", "Fuerteventura", "Lanzarote", "La Gomera", "El Hierro"]:
        legend_html.append(f"""
            <div class="legend-item">
                <span class="legend-dot"></span>
                <span class="legend-name">{isla}</span>
                <span class="legend-value">{valores.get(isla, 0)}</span>
            </div>
        """)

    html = f"""
    <div class="capacity-map-card">
        <div class="capacity-map-title">Capacidad Docente por Isla</div>
        <div class="map-wrap">
            <img class="map-img" src="data:image/png;base64,{encoded_map}" />
            {''.join(labels_html)}
        </div>
        <div class="legend-wrap">
            {''.join(legend_html)}
        </div>
    </div>
    <style>
        .capacity-map-card {{
            width: 100%;
            box-sizing: border-box;
            background: #ffffff;
            border: 1px solid rgba(49, 61, 83, 0.16);
            border-radius: 18px;
            padding: 18px 22px 20px 22px;
            box-shadow: 0 8px 28px rgba(15, 23, 42, 0.08);
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        }}
        .capacity-map-title {{
            text-align: center;
            font-size: 30px;
            line-height: 1.2;
            font-weight: 750;
            color: #17233d;
            margin: 4px 0 6px 0;
        }}
        .map-wrap {{
            position: relative;
            width: 100%;
            aspect-ratio: 1600 / 912;
            overflow: hidden;
        }}
        .map-img {{
            position: absolute;
            left: 0;
            top: 0;
            width: 100%;
            height: 100%;
            object-fit: contain;
            opacity: 0.92;
            pointer-events: none;
        }}
        .map-label {{
            position: absolute;
            transform: translate(-50%, -50%);
            min-width: 86px;
            background: rgba(255, 255, 255, 0.94);
            border: 1px solid rgba(37, 99, 235, 0.15);
            border-radius: 12px;
            padding: 7px 10px;
            text-align: center;
            box-shadow: 0 4px 14px rgba(15, 23, 42, 0.10);
            z-index: 2;
        }}
        .island-name {{
            font-size: 14px;
            font-weight: 650;
            color: #1f2937;
            white-space: nowrap;
        }}
        .island-value {{
            font-size: 24px;
            line-height: 1.05;
            font-weight: 800;
            color: #2563eb;
            margin-top: 3px;
        }}
        .map-dot {{
            position: absolute;
            transform: translate(-50%, -50%);
            width: 8px;
            height: 8px;
            border-radius: 50%;
            background: #2563eb;
            box-shadow: 0 0 0 4px rgba(37, 99, 235, 0.12);
            z-index: 3;
        }}
        .legend-wrap {{
            display: grid;
            grid-template-columns: repeat(7, minmax(0, 1fr));
            gap: 8px;
            margin-top: 6px;
            padding: 12px 10px;
            border: 1px solid rgba(49, 61, 83, 0.12);
            border-radius: 14px;
            background: #fbfdff;
        }}
        .legend-item {{
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 2px;
            border-right: 1px solid rgba(49, 61, 83, 0.10);
        }}
        .legend-item:last-child {{ border-right: none; }}
        .legend-dot {{
            width: 8px;
            height: 8px;
            border-radius: 50%;
            background: #2563eb;
            margin-bottom: 2px;
        }}
        .legend-name {{
            font-size: 13px;
            font-weight: 600;
            color: #1f2937;
            text-align: center;
        }}
        .legend-value {{
            font-size: 20px;
            font-weight: 800;
            color: #2563eb;
        }}
        @media (max-width: 900px) {{
            .capacity-map-title {{ font-size: 24px; }}
            .map-label {{ min-width: 72px; padding: 5px 7px; }}
            .island-name {{ font-size: 11px; }}
            .island-value {{ font-size: 18px; }}
            .legend-wrap {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
            .legend-item {{ border-right: none; }}
        }}
    </style>
    """
    components.html(html, height=780, scrolling=False)

def render_streamlit_dashboard(analytics: dict[str, pd.DataFrame]) -> None:
    global_df = analytics.get("Resumen_Global", pd.DataFrame())
    values = dict(zip(global_df.get("Indicador", []), global_df.get("Valor", []))) if not global_df.empty else {}
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total plazas", values.get("Total plazas", 0))
    c2.metric("Las Palmas", values.get("Total provincia Las Palmas", 0))
    c3.metric("S/C Tenerife", values.get("Total provincia S/C Tenerife", 0))
    c4.metric("Centros pendientes", values.get("Centros pendientes/sin finalizar", 0))

    if is_consulta_user():
        st.markdown("### Visualización territorial")
        render_canarias_capacity_map(analytics)
        st.markdown("")

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("#### Por isla")
        isla = analytics.get("Resumen_Isla", pd.DataFrame())
        if not isla.empty:
            isla = filtrar_resumen_total_plazas_positivo(isla)
        if not isla.empty:
            st.bar_chart(isla.set_index("Isla")["Total plazas"])
            st.dataframe(isla, use_container_width=True, hide_index=True)
    with col_b:
        st.markdown("#### Por rama")
        rama = analytics.get("Resumen_Rama", pd.DataFrame())
        if not rama.empty:
            rama = filtrar_resumen_total_plazas_positivo(rama)
        if not rama.empty:
            st.bar_chart(rama.head(10).set_index("Rama")["Total plazas"])
            st.dataframe(rama.head(10), use_container_width=True, hide_index=True)

    st.markdown("#### Top titulaciones")
    top = analytics.get("Top_Titulaciones", pd.DataFrame())
    if not top.empty:
        top = filtrar_resumen_total_plazas_positivo(top)
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
    show = filtrar_resumen_total_plazas_positivo(df).head(max_rows).copy().fillna("")
    # Si tras filtrar quedan solo filas con total 0, no se muestra ni el título ni la tabla.
    if show.empty:
        return
    story.append(Paragraph(title, styles["Heading2"]))
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



def add_pdf_table_wrapped(story: list, title: str, df: pd.DataFrame, styles, max_rows: int = 20, col_widths: list[float] | None = None, font_size: float = 6.0) -> None:
    """Añade una tabla PDF con ajuste automático de texto, pensada para observaciones y detalle de turnos."""
    if df is None or df.empty:
        return
    show = df.head(max_rows).copy().fillna("")
    if show.empty:
        return
    story.append(Paragraph(title, styles["Heading2"]))

    if ParagraphStyle is None:
        data = [list(show.columns)] + show.astype(str).values.tolist()
    else:
        header_style = ParagraphStyle(
            f"Header_{title}",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=6.2,
            leading=7.2,
            textColor=colors.white,
        )
        cell_style = ParagraphStyle(
            f"Cell_{title}",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=font_size,
            leading=font_size + 1.2,
        )
        data = [[Paragraph(str(col), header_style) for col in show.columns]]
        for _, row in show.iterrows():
            data.append([Paragraph(str(row.get(col, "")), cell_style) for col in show.columns])

    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 6.2),
        ("FONTSIZE", (0, 1), (-1, -1), font_size),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    story.append(table)
    story.append(Spacer(1, 8))

def draw_pdf_footer(canvas, doc):
    """Cabecera y pie de página común en los informes PDF publicados."""
    canvas.saveState()
    width, height = landscape(A3)
    if LOGO_PATH.exists():
        try:
            # Logo superior derecho, ligeramente más pequeño que en versiones anteriores.
            canvas.drawImage(
                str(LOGO_PATH),
                width - doc.rightMargin - 72,
                height - 30,
                width=72,
                height=24,
                preserveAspectRatio=True,
                mask="auto",
            )
        except Exception:
            pass
    footer_y = 12
    canvas.setStrokeColor(colors.HexColor("#BFBFBF"))
    canvas.setLineWidth(0.4)
    canvas.line(doc.leftMargin, footer_y + 10, width - doc.rightMargin, footer_y + 10)
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(colors.HexColor("#4F4F4F"))
    canvas.drawString(doc.leftMargin, footer_y, SIGNATURE_FOOTER)
    canvas.drawRightString(width - doc.rightMargin, footer_y, f"Página {doc.page}")
    canvas.restoreState()


def generate_matriz_pdf(
    matriz: pd.DataFrame,
    titulo: str,
    resumen_lineas: list[str],
    analytics: dict[str, pd.DataFrame] | None = None,
    include_quality: bool = True,
) -> bytes:
    if SimpleDocTemplate is None or Table is None:
        raise RuntimeError("La librería reportlab no está instalada. Añada reportlab a requirements.txt y reinicie la app.")

    matriz_pdf = filtrar_matriz_total_positivo(matriz, mantener_fila_total=True)
    output = io.BytesIO()
    doc = SimpleDocTemplate(
        output,
        pagesize=landscape(A3),
        leftMargin=18,
        rightMargin=18,
        topMargin=34,
        bottomMargin=28,
    )
    styles = getSampleStyleSheet()
    story = []
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
        if any(key in analytics for key in ["Turnos_Resumen", "Turnos_Resumen_Centro", "Turnos_Resumen_Isla", "Turnos_Detalle", "Observaciones_Titulacion"]):
            story.append(Spacer(1, 8))
            story.append(Paragraph("Distribución por turnos y observaciones", styles["Heading1"]))
            add_pdf_table(story, "Resumen global por turnos", analytics.get("Turnos_Resumen", pd.DataFrame()), styles, max_rows=10)
            add_pdf_table(story, "Resumen de turnos por isla", analytics.get("Turnos_Resumen_Isla", pd.DataFrame()), styles, max_rows=10)
            add_pdf_table(story, "Resumen de turnos por centro docente", analytics.get("Turnos_Resumen_Centro", pd.DataFrame()), styles, max_rows=20)
            add_pdf_table_wrapped(
                story,
                "Observaciones por titulación",
                analytics.get("Observaciones_Titulacion_PDF", pd.DataFrame()),
                styles,
                max_rows=30,
                col_widths=[145, 58, 58, 70, 205, 30, 500],
                font_size=5.8,
            )
            add_pdf_table_wrapped(
                story,
                "Detalle de turnos por titulación",
                analytics.get("Turnos_Detalle_PDF", pd.DataFrame()),
                styles,
                max_rows=30,
                col_widths=[145, 54, 54, 65, 195, 28, 42, 42, 42, 42, 250],
                font_size=5.6,
            )
        if include_quality and "Calidad_Resumen" in analytics:
            story.append(Spacer(1, 8))
            story.append(Paragraph("Validaciones de calidad de datos", styles["Heading1"]))
            add_pdf_table(story, "Resumen de controles", analytics.get("Calidad_Resumen", pd.DataFrame()), styles, max_rows=20)
            add_pdf_table(story, "Registros con valores altos", analytics.get("Calidad_Valores_Altos", pd.DataFrame()), styles, max_rows=15)
            add_pdf_table(story, "Duplicados por centro/titulación", analytics.get("Calidad_Duplicados", pd.DataFrame()), styles, max_rows=15)
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
    doc.build(story, onFirstPage=draw_pdf_footer, onLaterPages=draw_pdf_footer)
    output.seek(0)
    return output.getvalue()


def build_publication_package() -> tuple[bool, str, dict | None]:
    """Construye los objetos necesarios para una publicación, sin guardarlos todavía."""
    client = get_supabase_client()
    if client is None:
        return False, "La Base de Datos no está configurada.", None

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
            item = registro_from_db_row(row)
            volcar_registro_en_matriz(matriz, item, columna_excel)
            registros_consolidados.append({
                "Código borrador": codigo,
                "Fecha guardado": borrador.get("saved_at") or borrador.get("updated_at", ""),
                "Área": normalizar_area(borrador.get("area", "")),
                "Centro docente": unidad,
                "Columna Excel": columna_excel,
                "Usuario aportación": row.get("usuario_aportacion", "") or row.get("usuario_propietario", ""),
                **item,
            })

    analytics = build_analytics_tables(matriz, status_df)
    registros_df = pd.DataFrame(registros_consolidados)
    turnos_tables = build_turnos_tables(registros_df)
    detalle_alumnos = build_detalle_alumnos_table(registros_df)
    quality_tables = build_quality_tables(matriz, registros_df, status_df, duplicados)
    analytics.update(turnos_tables)
    analytics["Detalle_Alumnos"] = detalle_alumnos
    analytics.update(quality_tables)

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
        return False, None, "La Base de Datos no está configurada."
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
        return pd.DataFrame(), "La Base de Datos no está configurada."

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
        return False, None, "La Base de Datos no está configurada."

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
        return False, "El servidor de correo no está configurado; no se ha enviado correo de notificación."
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
        return False, "La Base de Datos no está configurada."
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
        return False, "La Base de Datos no está configurada."

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
        centros_incluidos = make_json_safe([b.get("unidad_docente", "") for b in package["borradores"]])
        centros_pendientes = make_json_safe(missing)
        estado_centros = dataframe_records_json_safe(package.get("status_df", pd.DataFrame()))

        payload_publicacion = make_json_safe({
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
            "centros_pendientes": centros_pendientes,
            "centros_con_borrador_no_finalizado": estado_centros,
            "observaciones": "Publicación vigente generada desde DCD 1.2.1 beta 4.",
        })
        client.table("dcd_publicaciones").insert(payload_publicacion).execute()
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
                df_sheet = pd.read_excel(xls, sheet_name=sheet)
                if sheet != "Resumen_Global":
                    df_sheet = filtrar_resumen_total_plazas_positivo(df_sheet)
                analytics[sheet] = df_sheet
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


def filter_excel_for_consulta(excel_bytes: bytes) -> bytes:
    """Devuelve una copia del Excel solo con las hojas permitidas para usuarios de consulta."""
    try:
        xls = pd.ExcelFile(io.BytesIO(excel_bytes))
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            written = False
            for sheet_name in CONSULTA_EXCEL_SHEETS:
                if sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet_name)
                    safe_sheet = sheet_name[:31]
                    df.to_excel(writer, sheet_name=safe_sheet, index=False)
                    worksheet = writer.sheets[safe_sheet]
                    workbook = writer.book
                    header_format = workbook.add_format({
                        "bold": True,
                        "bg_color": "#1F4E78",
                        "font_color": "#FFFFFF",
                        "border": 1,
                        "align": "center",
                        "valign": "vcenter",
                    })
                    for col_idx, col_name in enumerate(df.columns):
                        worksheet.write(0, col_idx, col_name, header_format)
                        try:
                            sample = df[col_name].head(200).fillna("").astype(str).tolist()
                        except Exception:
                            sample = []
                        max_len = max([len(str(col_name))] + [len(x) for x in sample]) if len(df.columns) else 12
                        worksheet.set_column(col_idx, col_idx, min(max(max_len + 2, 12), 45))
                    worksheet.freeze_panes(1, 0)
                    if not df.empty and len(df.columns) > 0:
                        worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
                    written = True
            if not written:
                # Fallback defensivo: si no hay hojas analíticas, se entrega la primera hoja para no romper la descarga.
                first = xls.sheet_names[0]
                pd.read_excel(xls, sheet_name=first).to_excel(writer, sheet_name=first[:31], index=False)
        output.seek(0)
        return output.getvalue()
    except Exception:
        # Si el filtrado fallara por cualquier motivo, devolvemos el archivo original para no bloquear la consulta.
        return excel_bytes


def build_consulta_pdf_from_publication(pub: dict) -> tuple[bool, bytes | None, str]:
    """Genera un PDF específico para rol consulta desde el Excel vigente.

    No usa el PDF administrativo almacenado porque este puede contener bloques
    internos de calidad de datos. Este PDF es de presentación externa: sin calidad
    y con filas/resúmenes de total 0 filtrados.
    """
    ruta_excel = (pub or {}).get("ruta_excel", "")
    if not ruta_excel:
        return False, None, "La publicación vigente no tiene Excel asociado."

    ok, data, msg = download_publication_file(ruta_excel)
    if not ok or not data:
        return False, None, msg

    try:
        xls = pd.ExcelFile(io.BytesIO(data))
        if "Matriz_DCD" not in xls.sheet_names:
            return False, None, "El Excel publicado no contiene la hoja Matriz_DCD."

        matriz = pd.read_excel(xls, sheet_name="Matriz_DCD")
        if "Titulación" in matriz.columns:
            matriz = matriz[matriz["Titulación"].astype(str).str.upper().str.strip() != "TOTAL"].copy()

        analytics: dict[str, pd.DataFrame] = {}
        for sheet in [
            "Resumen_Global",
            "Resumen_Provincia",
            "Resumen_Isla",
            "Resumen_Centro",
            "Resumen_Rama",
            "Resumen_Nivel",
            "Resumen_Centro_Rama",
            "Resumen_Centro_Nivel",
            "Top_Titulaciones",
            "Turnos_Resumen",
            "Turnos_Resumen_Centro",
            "Turnos_Resumen_Isla",
            "Turnos_Detalle",
            "Turnos_Detalle_PDF",
            "Observaciones_Titulacion",
            "Observaciones_Titulacion_PDF",
        ]:
            if sheet in xls.sheet_names:
                df_sheet = pd.read_excel(xls, sheet_name=sheet)
                if sheet != "Resumen_Global":
                    df_sheet = filtrar_resumen_total_plazas_positivo(df_sheet)
                analytics[sheet] = df_sheet

        resumen_lineas = [
            f"Publicación vigente: {pub.get('codigo_publicacion', '')}",
            f"Fecha de publicación: {str(pub.get('fecha_publicacion', ''))[:19]}",
            "Informe de consulta: se ocultan filas y resúmenes sin capacidad docente.",
        ]
        pdf = generate_matriz_pdf(
            matriz,
            titulo="DATOS CAPACIDAD DOCENTE (DCD) - Informe de consulta",
            resumen_lineas=resumen_lineas,
            analytics=analytics,
            include_quality=False,
        )
        return True, pdf, "PDF de consulta generado."
    except Exception as exc:
        return False, None, f"Error al generar PDF de consulta: {exc}"


def render_publication_downloads(pub: dict, prefix: str = "portal") -> None:
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Preparar PDF", key=f"{prefix}_prepare_pdf"):
            if is_consulta_user():
                ok, data, msg = build_consulta_pdf_from_publication(pub)
                pdf_filename = f"{pub.get('codigo_publicacion', 'DCD')}_consulta.pdf"
            else:
                ok, data, msg = download_publication_file(pub.get("ruta_pdf", ""))
                pdf_filename = f"{pub.get('codigo_publicacion', 'DCD')}_Matriz_DCD.pdf"
            if ok and data:
                st.download_button(
                    "Descargar PDF de informe",
                    data=data,
                    file_name=pdf_filename,
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
                excel_data = data
                excel_filename = f"{pub.get('codigo_publicacion', 'DCD')}.xlsx"
                if is_consulta_user():
                    excel_data = filter_excel_for_consulta(data)
                    excel_filename = f"{pub.get('codigo_publicacion', 'DCD')}_consulta.xlsx"
                st.download_button(
                    "Descargar Excel consolidado",
                    data=excel_data,
                    file_name=excel_filename,
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
        return False, "La Base de Datos no está configurada.", None, ""

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
                item = registro_from_db_row(row)
                volcar_registro_en_matriz(matriz, item, columna_excel)
                registros_consolidados.append({
                    "Código borrador": codigo,
                    "Fecha guardado": borrador.get("saved_at") or borrador.get("updated_at", ""),
                    "Área": normalizar_area(borrador.get("area", "")),
                    "Centro docente": unidad,
                    "Columna Excel": columna_excel,
                    "Usuario aportación": row.get("usuario_aportacion", "") or row.get("usuario_propietario", ""),
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
        quality_tables = build_quality_tables(matriz, registros_df, status_df, duplicados)
        turnos_tables = build_turnos_tables(registros_df)
        detalle_alumnos = build_detalle_alumnos_table(registros_df)

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
                "Turnos_Resumen": turnos_tables["Turnos_Resumen"],
                "Turnos_Resumen_Centro": turnos_tables["Turnos_Resumen_Centro"],
                "Turnos_Resumen_Isla": turnos_tables["Turnos_Resumen_Isla"],
                "Turnos_Detalle": turnos_tables["Turnos_Detalle"],
                "Turnos_Detalle_PDF": turnos_tables.get("Turnos_Detalle_PDF", pd.DataFrame()),
                "Observaciones_Titulacion": turnos_tables["Observaciones_Titulacion"],
                "Observaciones_Titulacion_PDF": turnos_tables.get("Observaciones_Titulacion_PDF", pd.DataFrame()),
                "Detalle_Alumnos": detalle_alumnos,
                "Calidad_Resumen": quality_tables["Calidad_Resumen"],
                "Calidad_Pendientes": quality_tables["Calidad_Pendientes"],
                "Calidad_Duplicados": quality_tables["Calidad_Duplicados"],
                "Calidad_Valores_Altos": quality_tables["Calidad_Valores_Altos"],
                "Calidad_Sin_Plazas": quality_tables["Calidad_Sin_Plazas"],
                "Estado_centros": status_df,
                "Borradores_usados": borradores_df,
                "Registros_DCD": dataframe_excel_safe(registros_df),
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
def scroll_to_top() -> None:
    """Intenta llevar la vista a la parte superior tras cambios de paso en Streamlit."""
    components.html(
        """
        <script>
            const doc = window.parent.document;
            const main = doc.querySelector('section.main') || doc.querySelector('[data-testid="stAppViewContainer"]');
            if (main) { main.scrollTo({top: 0, behavior: 'smooth'}); }
            window.parent.scrollTo({top: 0, behavior: 'smooth'});
        </script>
        """,
        height=0,
    )


def app_sidebar() -> None:
    st.sidebar.markdown("---")
    st.sidebar.markdown("### ℹ️ Información")
    st.sidebar.write(f"Versión: {APP_VERSION}")
    st.sidebar.write(f"Desarrollado para: {APP_DEVELOPED_FOR}")
    if st.session_state.get("current_user_role") == "admin":
        st.sidebar.caption(f"Desarrollador / creador: {APP_AUTHOR}")
    if st.session_state.get("current_user_display"):
        st.sidebar.write(f"Usuario: {st.session_state.current_user_display}")
        st.sidebar.caption(f"Rol: {st.session_state.get('current_user_role', '')}")
        if st.session_state.get("current_user_unidad"):
            st.sidebar.caption(f"Centro asignado: {st.session_state.current_user_unidad}")

    if supabase_available():
        st.sidebar.success("Base de Datos configurada")
    else:
        st.sidebar.warning("Base de Datos no configurada")
        st.sidebar.caption("El aplicativo funciona en modo local, pero no guarda borradores entre sesiones.")

    if mailgun_configured():
        st.sidebar.success("Servidor correo configurado")
    else:
        st.sidebar.info("Servidor correo no configurado")

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
            # DCD 1.2.1 beta 2: el rol admin entra directamente al panel administrador.
            # Los usuarios de centro mantienen el flujo normal de instrucciones/entrada de datos.
            if st.session_state.get("current_user_role") == "admin":
                st.session_state.current_step = 6
            else:
                st.session_state.current_step = 1
            st.rerun()
        else:
            st.error(msg)
            audit_event("login_fallido", f"Intento fallido. Usuario: {username}")

    st.markdown(
        f"""
        <div style="position: fixed; right: 18px; bottom: 12px; color: #888; font-size: 0.80rem; z-index: 999;">
            Versión {APP_VERSION.replace('DCD ', '')}
        </div>
        """,
        unsafe_allow_html=True,
    )


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
        6. Si la Base de Datos está configurada, podrá guardar y recuperar borradores.
        7. Antes de finalizar, el aplicativo mostrará una pantalla específica de recordatorio y descarga/envío.

        **Advertencias:**

        - Revise bien la titulación seleccionada antes de añadirla.
        - Si introduce de nuevo una misma combinación, se actualizará el número de alumnos.
        - El envío automático por correo solo funcionará si el servidor de correo aparece en verde.
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
            st.info("La Base de Datos no está configurada. No se pueden cargar borradores guardados.")
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
        st.info("Cuando esté configurada la Base de Datos, aquí aparecerá la carga de borradores.")


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

    if is_admin():
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

    # Si se pulsó "Editar registro seleccionado" en la ejecución anterior,
    # cargar ahora los valores en session_state antes de crear los widgets.
    aplicar_edicion_pendiente_en_formulario()

    editing_mode = bool(st.session_state.get("editing_registro_key"))
    st.markdown("### Añadir o actualizar titulación")
    if editing_mode:
        st.info("Está editando un registro ya introducido. Modifique los datos necesarios y pulse **Actualizar registro**.")

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
    titulacion_libre = es_modo_titulacion_libre(st.session_state.sel_nivel_i, st.session_state.sel_nivel_ii)
    if titulacion_libre:
        st.caption("Para Universitario + Máster/Otro la titulación se introduce como texto libre. Las ya usadas aparecerán como sugerencias.")
        titulaciones_usadas = obtener_titulaciones_libres_usadas(
            st.session_state.sel_nivel_i,
            st.session_state.sel_nivel_ii,
            st.session_state.sel_rama,
        ) if st.session_state.sel_rama else []
        st.selectbox(
            "Titulaciones libres ya utilizadas",
            options=[""] + titulaciones_usadas,
            key="sel_titulacion_libre_catalogo",
            on_change=aplicar_titulacion_libre_catalogo,
            disabled=edicion_bloqueada or not bool(st.session_state.sel_rama) or not bool(titulaciones_usadas),
            help="Opcional. Si selecciona una titulación ya usada, se copiará al campo de texto.",
        )
        st.text_input(
            "Titulación",
            key="sel_titulacion",
            disabled=edicion_bloqueada or not bool(st.session_state.sel_rama),
            placeholder="Escriba el nombre del máster o de la titulación",
        )
    else:
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

    st.markdown("#### Distribución por turnos")
    st.caption("Indique cuántos alumnos de esta titulación corresponden a cada turno. La suma debe coincidir con el número total de alumnos.")
    t1, t2, t3, t4 = st.columns(4)
    with t1:
        st.number_input("Mañana", min_value=0, step=1, key="alumnos_manana", disabled=edicion_bloqueada or not bool(st.session_state.sel_titulacion))
    with t2:
        st.number_input("Tarde", min_value=0, step=1, key="alumnos_tarde", disabled=edicion_bloqueada or not bool(st.session_state.sel_titulacion))
    with t3:
        st.number_input("Rotatorio", min_value=0, step=1, key="alumnos_rotatorio", disabled=edicion_bloqueada or not bool(st.session_state.sel_titulacion))
    with t4:
        st.number_input("Deslizante", min_value=0, step=1, key="alumnos_deslizante", disabled=edicion_bloqueada or not bool(st.session_state.sel_titulacion))

    suma_turnos_form = (
        int(st.session_state.get("alumnos_manana", 0) or 0)
        + int(st.session_state.get("alumnos_tarde", 0) or 0)
        + int(st.session_state.get("alumnos_rotatorio", 0) or 0)
        + int(st.session_state.get("alumnos_deslizante", 0) or 0)
    )
    total_form = int(st.session_state.get("numero_alumnos", 0) or 0)
    if bool(st.session_state.sel_titulacion):
        if suma_turnos_form == total_form:
            st.success(f"Suma de turnos correcta: {suma_turnos_form} / {total_form}")
        else:
            st.warning(f"La suma de turnos ({suma_turnos_form}) debe coincidir con el número total de alumnos ({total_form}).")

    n_deslizantes_form = int(st.session_state.get("alumnos_deslizante", 0) or 0)
    if n_deslizantes_form > 0:
        st.markdown("#### Patrón semanal del turno deslizante")
        st.caption("Indique un patrón por cada alumno en turno deslizante: M = mañana, T = tarde, R = rotatorio.")
        preparar_deslizantes_alumno_widgets(n_deslizantes_form, st.session_state.get("detalle_alumnos", []))
        for desl_idx in range(1, n_deslizantes_form + 1):
            st.write(f"Alumno deslizante {desl_idx}")
            d1, d2, d3, d4, d5 = st.columns(5)
            with d1:
                st.selectbox("Lunes", options=TURNO_DIA_OPTIONS, key=f"detalle_alumno_deslizante_{desl_idx}_lunes", disabled=edicion_bloqueada)
            with d2:
                st.selectbox("Martes", options=TURNO_DIA_OPTIONS, key=f"detalle_alumno_deslizante_{desl_idx}_martes", disabled=edicion_bloqueada)
            with d3:
                st.selectbox("Miércoles", options=TURNO_DIA_OPTIONS, key=f"detalle_alumno_deslizante_{desl_idx}_miercoles", disabled=edicion_bloqueada)
            with d4:
                st.selectbox("Jueves", options=TURNO_DIA_OPTIONS, key=f"detalle_alumno_deslizante_{desl_idx}_jueves", disabled=edicion_bloqueada)
            with d5:
                st.selectbox("Viernes", options=TURNO_DIA_OPTIONS, key=f"detalle_alumno_deslizante_{desl_idx}_viernes", disabled=edicion_bloqueada)
    else:
        st.session_state.deslizante_lunes = ""
        st.session_state.deslizante_martes = ""
        st.session_state.deslizante_miercoles = ""
        st.session_state.deslizante_jueves = ""
        st.session_state.deslizante_viernes = ""

    st.text_area(
        "Observaciones de esta titulación",
        key="observaciones_titulacion",
        disabled=edicion_bloqueada or not bool(st.session_state.sel_titulacion),
        help="Campo específico por titulación/especialidad. Se mostrará en Excel y PDF de consulta.",
    )

    st.markdown("#### Detalle opcional por alumno")
    st.caption("Campos voluntarios. No introduzca nombres ni datos personales; use solo Alumno 1, Alumno 2, etc.")
    total_detalle_form = int(st.session_state.get("numero_alumnos", 0) or 0)
    preparar_detalle_alumnos_widgets(total_detalle_form, st.session_state.get("detalle_alumnos", []))
    with st.expander("Servicio y curso/año por alumno", expanded=False):
        if not bool(st.session_state.sel_titulacion):
            st.info("Seleccione una titulación para activar el detalle por alumno.")
        elif total_detalle_form <= 0:
            st.info("Indique primero el número de alumnos para poder rellenar el detalle opcional.")
        else:
            st.info("Estos campos son voluntarios y se exportarán en la hoja Detalle_Alumnos del Excel.")
            for alumno_idx in range(1, total_detalle_form + 1):
                a1, a2, a3 = st.columns([1, 3, 2])
                with a1:
                    st.write(f"Alumno {alumno_idx}")
                with a2:
                    st.text_input(
                        f"Servicio alumno {alumno_idx}",
                        key=f"detalle_alumno_servicio_{alumno_idx}",
                        disabled=edicion_bloqueada or not bool(st.session_state.sel_titulacion),
                        label_visibility="collapsed",
                        placeholder="Servicio",
                    )
                with a3:
                    st.text_input(
                        f"Curso/año alumno {alumno_idx}",
                        key=f"detalle_alumno_curso_{alumno_idx}",
                        disabled=edicion_bloqueada or not bool(st.session_state.sel_titulacion),
                        label_visibility="collapsed",
                        placeholder="Curso/año",
                    )

    col_add, col_clear, col_cancel = st.columns(3)
    with col_add:
        boton_registro = "Actualizar registro" if editing_mode else "Añadir registro"
        if st.button(boton_registro, disabled=edicion_bloqueada):
            if not all([
                st.session_state.sel_nivel_i,
                st.session_state.sel_nivel_ii,
                st.session_state.sel_rama,
                st.session_state.sel_titulacion,
            ]):
                st.warning("Debe completar Nivel I, Nivel II, Rama y Titulación.")
            elif suma_turnos_form != total_form:
                st.error("No se puede añadir el registro: la suma de alumnos por turnos debe coincidir con el número total de alumnos.")
            elif int(st.session_state.get("alumnos_deslizante", 0) or 0) > 0 and not patrones_deslizantes_widgets_completos(int(st.session_state.get("alumnos_deslizante", 0) or 0)):
                st.error("Si hay alumnos en turno deslizante, debe indicar el patrón completo de lunes a viernes para cada alumno deslizante.")
            else:
                key = registro_key(
                    st.session_state.sel_nivel_i,
                    st.session_state.sel_nivel_ii,
                    st.session_state.sel_rama,
                    st.session_state.sel_titulacion,
                )
                original_key = st.session_state.get("editing_registro_key", "")
                if editing_mode and original_key and original_key != key:
                    st.session_state.registros.pop(original_key, None)
                detalle_recogido = recoger_detalle_alumnos_desde_widgets(total_form)
                patron_global = primer_patron_deslizante_para_campos_globales(detalle_recogido)
                st.session_state.registros[key] = {
                    "Nivel Estudio I": st.session_state.sel_nivel_i,
                    "Nivel Estudio II": st.session_state.sel_nivel_ii,
                    "Rama": st.session_state.sel_rama,
                    "Titulación": st.session_state.sel_titulacion,
                    "Nº alumnos": int(st.session_state.numero_alumnos),
                    "Alumnos mañana": int(st.session_state.get("alumnos_manana", 0) or 0),
                    "Alumnos tarde": int(st.session_state.get("alumnos_tarde", 0) or 0),
                    "Alumnos rotatorio": int(st.session_state.get("alumnos_rotatorio", 0) or 0),
                    "Alumnos deslizante": int(st.session_state.get("alumnos_deslizante", 0) or 0),
                    "Deslizante lunes": patron_global["lunes"] if int(st.session_state.get("alumnos_deslizante", 0) or 0) > 0 else "",
                    "Deslizante martes": patron_global["martes"] if int(st.session_state.get("alumnos_deslizante", 0) or 0) > 0 else "",
                    "Deslizante miércoles": patron_global["miercoles"] if int(st.session_state.get("alumnos_deslizante", 0) or 0) > 0 else "",
                    "Deslizante jueves": patron_global["jueves"] if int(st.session_state.get("alumnos_deslizante", 0) or 0) > 0 else "",
                    "Deslizante viernes": patron_global["viernes"] if int(st.session_state.get("alumnos_deslizante", 0) or 0) > 0 else "",
                    OBS_TITULACION_DISPLAY: st.session_state.get("observaciones_titulacion", ""),
                    DETALLE_ALUMNOS_DISPLAY: detalle_recogido,
                }
                audit_event("registro_actualizado", f"{st.session_state.sel_titulacion} | {int(st.session_state.numero_alumnos)} alumnos | turnos {suma_turnos_form}")
                if editing_mode:
                    st.session_state.editing_registro_key = ""
                    st.success("Registro actualizado correctamente.")
                else:
                    st.success("Registro añadido correctamente.")
    with col_clear:
        if st.button("Limpiar selectores", disabled=edicion_bloqueada):
            cancelar_edicion_registro()
            st.rerun()
    with col_cancel:
        if st.button("Cancelar edición", disabled=edicion_bloqueada or not editing_mode):
            cancelar_edicion_registro()
            st.rerun()

    st.markdown("---")
    st.markdown("### Registros introducidos")
    registros = list(st.session_state.registros.values())
    if registros:
        registros_df = pd.DataFrame(registros)
        registros_df_display = preparar_registros_para_visualizacion(registros_df)
        st.dataframe(registros_df_display, use_container_width=True, hide_index=True)

        registro_labels = [
            f"{i + 1}. {r['Nivel Estudio II']} | {r['Rama']} | {r['Titulación']} | {r['Nº alumnos']} alumnos"
            for i, r in enumerate(registros)
        ]
        action_label = st.selectbox("Seleccionar registro para editar o eliminar", options=[""] + registro_labels, disabled=edicion_bloqueada)
        col_edit_reg, col_delete_reg = st.columns(2)
        with col_edit_reg:
            if st.button("Editar registro seleccionado", disabled=edicion_bloqueada):
                if action_label:
                    idx = int(action_label.split(".", 1)[0]) - 1
                    item = registros[idx]
                    st.session_state.pending_edit_registro = dict(item)
                    audit_event("registro_edicion_iniciada", item.get("Titulación", ""))
                    st.rerun()
                else:
                    st.warning("Debe seleccionar un registro.")
        with col_delete_reg:
            if st.button("Eliminar registro seleccionado", disabled=edicion_bloqueada):
                if action_label:
                    idx = int(action_label.split(".", 1)[0]) - 1
                    item = registros[idx]
                    key = registro_key(item["Nivel Estudio I"], item["Nivel Estudio II"], item["Rama"], item["Titulación"])
                    st.session_state.registros.pop(key, None)
                    if st.session_state.get("editing_registro_key") == key:
                        cancelar_edicion_registro()
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
    scroll_to_top()
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
    total_turnos = sum(suma_turnos_registro(item) for item in registros)
    turnos_errors = []
    deslizante_errors = []
    for item in registros:
        if not registro_turnos_cuadra(item):
            turnos_errors.append(
                f"{item.get('Titulación', '')}: total {safe_int(item.get('Nº alumnos', 0))}, turnos {suma_turnos_registro(item)}"
            )
        if not registro_deslizante_completo(item):
            deslizante_errors.append(str(item.get("Titulación", "")))
    codigo = st.session_state.codigo_borrador or build_codigo_borrador(st.session_state.direccion_selected)
    filename = f"{codigo}.xlsx"
    excel_bytes = build_output_excel()

    st.subheader("Resumen de datos introducidos")
    st.dataframe(preparar_registros_para_visualizacion(df), use_container_width=True, hide_index=True)
    col_estado, col_total = st.columns(2)
    with col_estado:
        st.metric("Estado actual", st.session_state.get("draft_estado", "borrador").upper())
    with col_total:
        st.metric("Total alumnos introducidos", total_alumnos)
    st.metric("Total alumnos distribuidos por turnos", total_turnos)
    if not turnos_errors and not deslizante_errors:
        st.success("La distribución por turnos cuadra con el total de alumnos introducido.")
    else:
        if turnos_errors:
            st.error("Hay titulaciones cuya suma de turnos no coincide con el número total de alumnos.")
            for err in turnos_errors[:20]:
                st.write(f"- {err}")
        if deslizante_errors:
            st.error("Hay titulaciones con turno deslizante sin patrón completo de lunes a viernes.")
            for err in deslizante_errors[:20]:
                st.write(f"- {err}")

    # Control preventivo de calidad antes de finalizar. No bloquea, pero avisa de situaciones a revisar.
    st.markdown("### Revisión automática de calidad")
    quality_warnings = []
    if total_alumnos == 0:
        quality_warnings.append("El expediente tiene total de alumnos igual a 0.")
    if all(c in df.columns for c in KEY_COLUMNS):
        dup_count = int(df.duplicated(subset=KEY_COLUMNS, keep=False).sum())
        if dup_count:
            quality_warnings.append(f"Hay {dup_count} líneas con titulaciones repetidas dentro del mismo expediente.")
    if "Nº alumnos" in df.columns and not df.empty:
        max_val = int(pd.to_numeric(df["Nº alumnos"], errors="coerce").fillna(0).max())
        if max_val >= 100:
            quality_warnings.append(f"Existe al menos un registro con {max_val} alumnos. Revise si es correcto.")
    if quality_warnings:
        for warning in quality_warnings:
            st.warning(warning)
        st.caption("Estos avisos no impiden finalizar, pero ayudan a detectar posibles errores antes del cierre.")
    else:
        st.success("No se han detectado avisos básicos de calidad en este expediente.")

    bloqueos_turnos = bool(turnos_errors or deslizante_errors)
    if bloqueos_turnos:
        st.error("No se podrá finalizar el expediente hasta corregir la distribución por turnos.")

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
        if st.button("Finalizar expediente en Base de Datos"):
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
            elif bloqueos_turnos:
                st.warning("Debe corregir los errores de distribución por turnos antes de enviar el correo.")
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
    st.caption("Si el Servidor de Correo no está configurado, el botón de correo mostrará un aviso y podrá seguir usando la descarga manual.")


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
        st.warning("La Base de Datos no está configurada. No se puede generar el consolidado.")
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
        st.warning("La Base de Datos no está configurada. No se pueden gestionar publicaciones.")
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
        st.warning("La Base de Datos no está configurada. No se puede guardar configuración de cierre.")
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

    st.caption("Los avisos automáticos dependen del servidor de correo. Si no está configurado, la app dejará constancia del estado pero no enviará correo.")
    st.caption("La publicación por fecha tope se evalúa al usar la app o al pulsar 'Evaluar ahora cierre automático'. Para automatismo 24/7 sin accesos habría que añadir una tarea programada externa.")



# =========================================================
# ADMIN MULTIUSUARIO POR CENTRO
# =========================================================
def list_multiusuario_configs() -> list[dict]:
    client = get_supabase_client()
    if client is None:
        return []
    try:
        resp = client.table("dcd_centros_multiusuario_config").select("*").order("centro_docente").execute()
        return getattr(resp, "data", []) or []
    except Exception as exc:
        st.session_state["last_multiusuario_error"] = str(exc)
        return []


def get_multiusuario_config(centro_docente: str) -> dict:
    client = get_supabase_client()
    if client is None or not centro_docente:
        return {}
    try:
        resp = client.table("dcd_centros_multiusuario_config").select("*").eq("centro_docente", centro_docente).limit(1).execute()
        rows = getattr(resp, "data", []) or []
        return rows[0] if rows else {}
    except Exception as exc:
        st.session_state["last_multiusuario_error"] = str(exc)
        return {}


def is_centro_multiusuario_activo(centro_docente: str) -> bool:
    """Indica si un centro está configurado como multiusuario.

    Devuelve False ante cualquier problema para no bloquear el flujo ordinario.
    """
    try:
        cfg = get_multiusuario_config(centro_docente)
        return bool(cfg.get("multiusuario_activo", False))
    except Exception:
        return False


def list_active_multiusuario_usernames(centro_docente: str) -> list[str]:
    try:
        return [
            str(a.get("username", "")).strip().lower()
            for a in list_multiusuario_assignments(centro_docente)
            if a.get("activo") and a.get("username")
        ]
    except Exception:
        return []


def save_multiusuario_config(
    centro_docente: str,
    area: str,
    multiusuario_activo: bool,
    usuarios_previstos: int,
    permite_consolidacion_parcial: bool = True,
    observaciones_config: str = "",
) -> tuple[bool, str]:
    client = get_supabase_client()
    if client is None:
        return False, "La Base de Datos no está configurada."
    centro_docente = (centro_docente or "").strip()
    if not centro_docente:
        return False, "Debe seleccionar un centro docente."
    usuarios_previstos = max(1, int(usuarios_previstos or 1))
    try:
        client.table("dcd_centros_multiusuario_config").upsert({
            "centro_docente": centro_docente,
            "area": normalizar_area(area or ""),
            "multiusuario_activo": bool(multiusuario_activo),
            "usuarios_previstos": usuarios_previstos,
            "permite_consolidacion_parcial": bool(permite_consolidacion_parcial),
            "observaciones_config": observaciones_config or "",
            "actualizado_por": st.session_state.get("current_user", ""),
            "updated_at": datetime.now().isoformat(),
        }, on_conflict="centro_docente").execute()
        audit_event("guardar_config_multiusuario", f"Centro: {centro_docente} | Activo: {multiusuario_activo} | Usuarios previstos: {usuarios_previstos}")
        return True, "Configuración multiusuario guardada correctamente."
    except Exception as exc:
        return False, f"Error al guardar configuración multiusuario: {exc}"


def list_multiusuario_assignments(centro_docente: str) -> list[dict]:
    client = get_supabase_client()
    if client is None or not centro_docente:
        return []
    try:
        resp = client.table("dcd_centros_multiusuario_usuarios").select("*").eq("centro_docente", centro_docente).order("orden_participante").execute()
        return getattr(resp, "data", []) or []
    except Exception as exc:
        st.session_state["last_multiusuario_error"] = str(exc)
        return []


def sync_multiusuario_assignments(centro_docente: str, usernames: list[str]) -> tuple[bool, str]:
    client = get_supabase_client()
    if client is None:
        return False, "La Base de Datos no está configurada."
    centro_docente = (centro_docente or "").strip()
    usernames = [(u or "").strip().lower() for u in usernames if (u or "").strip()]
    if not centro_docente:
        return False, "Debe seleccionar un centro docente."
    try:
        existing = list_multiusuario_assignments(centro_docente)
        existing_by_user = {str(x.get("username", "")).strip().lower(): x for x in existing if x.get("username")}

        # Activar/crear seleccionados.
        for idx, username in enumerate(usernames, start=1):
            payload = {
                "centro_docente": centro_docente,
                "username": username,
                "email": username if "@" in username else "",
                "activo": True,
                "orden_participante": idx,
                "updated_at": datetime.now().isoformat(),
            }
            client.table("dcd_centros_multiusuario_usuarios").upsert(
                payload,
                on_conflict="centro_docente,username"
            ).execute()

        # Desactivar los que ya no estén seleccionados, manteniendo trazabilidad.
        for username, row in existing_by_user.items():
            if username not in usernames:
                client.table("dcd_centros_multiusuario_usuarios").update({
                    "activo": False,
                    "updated_at": datetime.now().isoformat(),
                }).eq("id", row.get("id")).execute()

        audit_event("sincronizar_usuarios_multiusuario", f"Centro: {centro_docente} | Usuarios activos: {', '.join(usernames)}")
        return True, "Usuarios asignados al centro sincronizados correctamente."
    except Exception as exc:
        return False, f"Error al sincronizar usuarios del centro: {exc}"




def get_multiusuario_progress(centro_docente: str) -> dict:
    """Devuelve el estado operativo de un centro multiusuario.

    Esta función no modifica datos. Se usa para saber qué usuarios activos
    asignados han finalizado su aportación y cuáles siguen pendientes.
    """
    client = get_supabase_client()
    cfg = get_multiusuario_config(centro_docente)
    assignments = [
        a for a in list_multiusuario_assignments(centro_docente)
        if a.get("activo") and a.get("username")
    ]
    active_users = [str(a.get("username", "")).strip().lower() for a in assignments if a.get("username")]
    expected = int(cfg.get("usuarios_previstos") or len(active_users) or 1)
    latest_by_user: dict[str, dict] = {}

    if client is not None:
        for username in active_users:
            try:
                resp = client.table("dcd_borradores").select("*").eq("unidad_docente", centro_docente).eq("usuario_propietario", username).eq("estado", "finalizado").order("saved_at", desc=True).order("updated_at", desc=True).limit(1).execute()
                rows = getattr(resp, "data", []) or []
                if rows:
                    latest_by_user[username] = rows[0]
            except Exception:
                pass

    finalized_users = [u for u in active_users if u in latest_by_user]
    pending_users = [u for u in active_users if u not in latest_by_user]
    complete = bool(active_users) and len(finalized_users) >= expected and all(u in finalized_users for u in active_users)
    return {
        "config": cfg,
        "assignments": assignments,
        "active_users": active_users,
        "expected": expected,
        "latest_by_user": latest_by_user,
        "finalized_users": finalized_users,
        "pending_users": pending_users,
        "complete": complete,
    }


def get_latest_multiusuario_consolidation(centro_docente: str) -> dict:
    """Devuelve la última consolidación operativa registrada para un centro.

    Se considera operativa si su estado es consolidado o consolidado_parcial.
    No se usa para centros no multiusuario.
    """
    client = get_supabase_client()
    if client is None or not centro_docente:
        return {}
    try:
        resp = client.table("dcd_centros_multiusuario_consolidados").select("*").eq("centro_docente", centro_docente).order("created_at", desc=True).limit(10).execute()
        rows = getattr(resp, "data", []) or []
        for row in rows:
            if str(row.get("estado", "")).lower() in {"consolidado", "consolidado_parcial"}:
                return row
    except Exception as exc:
        st.session_state["last_multiusuario_error"] = str(exc)
    return {}


def create_partial_multiusuario_consolidation(centro_docente: str, motivo: str = "") -> tuple[bool, str]:
    """Autoriza que un centro multiusuario entre parcialmente en el consolidado.

    No crea datos nuevos ni suma registros en tablas intermedias. Registra una autorización
    administrativa para que, al generar publicación/consolidado, entren las aportaciones
    finalizadas disponibles aunque falten usuarios asignados.
    """
    client = get_supabase_client()
    if client is None:
        return False, "La Base de Datos no está configurada."
    if not centro_docente:
        return False, "Debe seleccionar un centro docente."

    progress = get_multiusuario_progress(centro_docente)
    cfg = progress.get("config", {}) or {}
    if not bool(cfg.get("multiusuario_activo", False)):
        return False, "El centro no está marcado como multiusuario."
    if not bool(cfg.get("permite_consolidacion_parcial", True)):
        return False, "La consolidación parcial no está permitida para este centro."

    finalized_users = progress.get("finalized_users", []) or []
    pending_users = progress.get("pending_users", []) or []
    expected = int(progress.get("expected") or 1)

    if progress.get("complete"):
        return False, "El centro ya está completo; no necesita consolidación parcial."
    if not finalized_users:
        return False, "No hay aportaciones finalizadas para consolidar parcialmente."

    codigo = f"CONS-{safe_code(centro_docente)}-{datetime.now().strftime('%Y%m%d%H%M%S')}"
    payload = {
        "centro_docente": centro_docente,
        "codigo_consolidado": codigo,
        "estado": "consolidado_parcial",
        "total_usuarios_previstos": expected,
        "total_usuarios_finalizados": len(finalized_users),
        "consolidacion_parcial": True,
        "usuarios_pendientes": ", ".join(pending_users),
        "consolidado_por": st.session_state.get("current_user", "admin"),
        "motivo_consolidacion": motivo or "Consolidación parcial autorizada por administrador.",
        "created_at": datetime.now().isoformat(),
        "updated_at": datetime.now().isoformat(),
    }
    try:
        client.table("dcd_centros_multiusuario_consolidados").insert(payload).execute()
        audit_event(
            "consolidacion_parcial_multiusuario",
            f"Centro: {centro_docente} | Finalizados: {len(finalized_users)}/{expected} | Pendientes: {', '.join(pending_users)}",
            codigo,
        )
        return True, f"Consolidación parcial registrada para {centro_docente}. Entrarán {len(finalized_users)} aportaciones finalizadas."
    except Exception as exc:
        return False, f"Error al registrar consolidación parcial: {exc}"

def build_multiusuario_estado_df(centro_docente: str, assignments: list[dict]) -> pd.DataFrame:
    if not assignments:
        return pd.DataFrame(columns=["Usuario", "Activo", "Orden", "Estado aportación", "Último finalizado", "Fecha finalización", "Registros"])
    client = get_supabase_client()
    rows = []
    for a in assignments:
        username = str(a.get("username", "")).strip().lower()
        registros = 0
        codigo_finalizado = ""
        fecha_finalizacion = ""
        estado = "Pendiente"
        if client is not None and username:
            try:
                borr_resp = client.table("dcd_borradores").select("codigo_borrador,saved_at,updated_at,estado").eq("unidad_docente", centro_docente).eq("usuario_propietario", username).eq("estado", "finalizado").order("saved_at", desc=True).order("updated_at", desc=True).limit(1).execute()
                borr_rows = getattr(borr_resp, "data", []) or []
                if borr_rows:
                    codigo_finalizado = borr_rows[0].get("codigo_borrador", "")
                    fecha_finalizacion = borr_rows[0].get("saved_at") or borr_rows[0].get("updated_at") or ""
                    estado = "Aportación finalizada"
                    reg_resp = client.table("dcd_registros").select("id").eq("codigo_borrador", codigo_finalizado).execute()
                    registros = len(getattr(reg_resp, "data", []) or [])
                else:
                    draft_resp = client.table("dcd_borradores").select("codigo_borrador").eq("unidad_docente", centro_docente).eq("usuario_propietario", username).limit(1).execute()
                    if getattr(draft_resp, "data", []) or []:
                        estado = "Con borrador pendiente"
            except Exception:
                pass
        rows.append({
            "Usuario": username,
            "Activo": "Sí" if a.get("activo") else "No",
            "Orden": a.get("orden_participante", ""),
            "Estado aportación": estado,
            "Último finalizado": codigo_finalizado,
            "Fecha finalización": fecha_finalizacion,
            "Registros": registros,
        })
    return pd.DataFrame(rows)


def render_admin_multiusuario_centros() -> None:
    st.subheader("Multiusuario por centro docente")
    st.info(
        "Esta beta permite configurar centros multiusuario, revisar el estado de aportaciones por usuario "
        "y autorizar una consolidación parcial si llega la fecha límite y faltan usuarios por finalizar."
    )

    if not supabase_available():
        st.warning("La Base de Datos no está configurada. No se puede gestionar multiusuario.")
        return

    last_error = st.session_state.get("last_multiusuario_error", "")
    if last_error:
        st.warning(f"Último aviso multiusuario: {last_error}")

    users = list_users_from_supabase()
    users_df = pd.DataFrame(users)
    centros_configs = list_multiusuario_configs()

    if centros_configs:
        st.markdown("### Centros configurados")
        df_cfg = pd.DataFrame(centros_configs)
        cols = [c for c in ["centro_docente", "area", "multiusuario_activo", "usuarios_previstos", "permite_consolidacion_parcial", "actualizado_por", "updated_at"] if c in df_cfg.columns]
        st.dataframe(df_cfg[cols], use_container_width=True, hide_index=True)
    else:
        st.caption("Todavía no hay centros configurados específicamente como multiusuario.")

    st.markdown("---")
    st.markdown("### Configurar centro")

    area_sel = st.selectbox("Área", options=[""] + AREA_OPTIONS, key="multi_area_sel")
    centros_area = DIRECCIONES_POR_AREA.get(area_sel, []) if area_sel else []
    centro_sel = st.selectbox("Centro docente", options=[""] + centros_area, key="multi_centro_sel")

    if not centro_sel:
        st.info("Seleccione un área y un centro docente para configurar el modo multiusuario.")
        return

    current_cfg = get_multiusuario_config(centro_sel)
    current_assignments = list_multiusuario_assignments(centro_sel)

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        multi_activo = st.checkbox(
            "Centro multiusuario",
            value=bool(current_cfg.get("multiusuario_activo", False)),
            help="Si está desactivado, el centro funciona como hasta ahora: un expediente ordinario por centro.",
            key=f"multi_activo_{safe_code(centro_sel)}",
        )
    with col_b:
        usuarios_previstos = st.number_input(
            "Nº usuarios previstos",
            min_value=1,
            max_value=20,
            value=int(current_cfg.get("usuarios_previstos") or 1),
            step=1,
            key=f"multi_previstos_{safe_code(centro_sel)}",
        )
    with col_c:
        parcial = st.checkbox(
            "Permitir consolidación parcial por admin",
            value=bool(current_cfg.get("permite_consolidacion_parcial", True)),
            key=f"multi_parcial_{safe_code(centro_sel)}",
        )

    obs_cfg = st.text_area(
        "Observaciones internas de configuración",
        value=current_cfg.get("observaciones_config", "") if current_cfg else "",
        key=f"multi_obs_{safe_code(centro_sel)}",
    )

    eligible_users = []
    if not users_df.empty:
        tmp = users_df.copy()
        if "role" in tmp.columns:
            tmp = tmp[tmp["role"].astype(str).str.lower().eq("usuario")]
        if "unidad_docente" in tmp.columns:
            tmp = tmp[tmp["unidad_docente"].astype(str).eq(centro_sel)]
        if "activo" in tmp.columns:
            tmp = tmp[tmp["activo"].fillna(True).astype(bool)]
        eligible_users = sorted(tmp["username"].dropna().astype(str).str.lower().unique().tolist()) if "username" in tmp.columns else []

    assigned_active = sorted([str(a.get("username", "")).strip().lower() for a in current_assignments if a.get("activo") and a.get("username")])
    selected_users = st.multiselect(
        "Usuarios asignados a este centro",
        options=eligible_users,
        default=[u for u in assigned_active if u in eligible_users],
        help="Solo se muestran usuarios activos de rol usuario asignados previamente a este centro en Usuarios y permisos.",
        key=f"multi_users_{safe_code(centro_sel)}",
    )

    if multi_activo and len(selected_users) != int(usuarios_previstos):
        st.warning(
            f"El centro está marcado como multiusuario con {int(usuarios_previstos)} usuarios previstos, "
            f"pero actualmente hay {len(selected_users)} usuarios seleccionados."
        )

    col_save1, col_save2 = st.columns(2)
    with col_save1:
        if st.button("Guardar configuración multiusuario", key=f"btn_save_multi_{safe_code(centro_sel)}"):
            ok, msg = save_multiusuario_config(
                centro_docente=centro_sel,
                area=area_sel,
                multiusuario_activo=multi_activo,
                usuarios_previstos=int(usuarios_previstos),
                permite_consolidacion_parcial=parcial,
                observaciones_config=obs_cfg,
            )
            if ok:
                st.success(msg)
            else:
                st.error(msg)
    with col_save2:
        if st.button("Guardar usuarios asignados", key=f"btn_save_multi_users_{safe_code(centro_sel)}"):
            ok, msg = sync_multiusuario_assignments(centro_sel, selected_users)
            if ok:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)

    st.markdown("### Estado de usuarios asignados")
    assignments_after = list_multiusuario_assignments(centro_sel)
    if assignments_after:
        estado_df = build_multiusuario_estado_df(centro_sel, assignments_after)
        st.dataframe(estado_df, use_container_width=True, hide_index=True)
    else:
        st.info("No hay usuarios asignados a este centro todavía.")

    if multi_activo:
        st.markdown("### Consolidación del centro multiusuario")
        progress = get_multiusuario_progress(centro_sel)
        expected = int(progress.get("expected") or usuarios_previstos or 1)
        finalized_users = progress.get("finalized_users", []) or []
        pending_users = progress.get("pending_users", []) or []
        partial_auth = get_latest_multiusuario_consolidation(centro_sel)

        if progress.get("complete"):
            st.success(f"Centro completo: {len(finalized_users)}/{expected} usuarios han finalizado. Entrará en el consolidado general.")
        elif partial_auth and bool(partial_auth.get("consolidacion_parcial", False)):
            st.warning(
                f"Centro con consolidación parcial autorizada. Finalizados: {len(finalized_users)}/{expected}. "
                f"Pendientes: {partial_auth.get('usuarios_pendientes', '') or ', '.join(pending_users) or 'sin pendientes identificados'}."
            )
            cols_auth = [c for c in ["codigo_consolidado", "estado", "total_usuarios_finalizados", "total_usuarios_previstos", "usuarios_pendientes", "consolidado_por", "created_at", "motivo_consolidacion"] if c in partial_auth]
            st.dataframe(pd.DataFrame([{c: partial_auth.get(c, "") for c in cols_auth}]), use_container_width=True, hide_index=True)
        elif finalized_users:
            st.warning(
                f"Centro pendiente: {len(finalized_users)}/{expected} usuarios han finalizado. "
                f"Pendientes: {', '.join(pending_users) if pending_users else 'sin pendientes identificados'}."
            )
            if parcial:
                motivo_parcial = st.text_area(
                    "Motivo de consolidación parcial",
                    value="Consolidación parcial autorizada por fecha límite o instrucción administrativa.",
                    key=f"multi_motivo_parcial_{safe_code(centro_sel)}",
                )
                confirmar_parcial = st.checkbox(
                    "Confirmo que quiero consolidar parcialmente este centro con las aportaciones finalizadas disponibles.",
                    key=f"multi_confirm_parcial_{safe_code(centro_sel)}",
                )
                if st.button("Consolidar parcialmente este centro", key=f"btn_multi_partial_{safe_code(centro_sel)}"):
                    if not confirmar_parcial:
                        st.error("Debe marcar la casilla de confirmación para consolidar parcialmente.")
                    else:
                        ok, msg = create_partial_multiusuario_consolidation(centro_sel, motivo_parcial)
                        if ok:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)
            else:
                st.info("La consolidación parcial no está permitida en la configuración de este centro.")
        else:
            st.info(f"Centro pendiente: todavía no hay aportaciones finalizadas ({len(finalized_users)}/{expected}).")

    st.caption(
        "Regla operativa: un centro multiusuario entra automáticamente en el consolidado cuando todos los usuarios activos asignados finalizan. "
        "Si falta algún usuario, el admin puede autorizar consolidación parcial solo cuando esté permitido en la configuración del centro."
    )


def render_admin_usuarios() -> None:
    st.subheader("Usuarios y permisos")
    st.info("Las contraseñas se guardan hasheadas. El administrador puede resetearlas, pero no verlas.")

    if not supabase_available():
        st.warning("La Base de Datos no está configurada. No se pueden gestionar usuarios.")
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
        st.warning("La Base de Datos no está configurada. No se puede consultar auditoría ni generar backup.")
        return

    st.markdown("### Auditoría reciente")

    col_diag1, col_diag2 = st.columns([1, 2])
    with col_diag1:
        if st.button("Registrar evento de prueba", key="btn_test_auditoria"):
            ok = audit_event("test_auditoria", "Prueba manual desde Panel administrador > Auditoría/Backup")
            if ok:
                st.success("Evento de prueba registrado correctamente.")
                st.rerun()
            else:
                st.error(f"No se pudo registrar auditoría: {st.session_state.get('last_audit_error', 'Error desconocido')}")
    with col_diag2:
        last_error = st.session_state.get("last_audit_error", "")
        if last_error:
            st.caption(f"Último error de auditoría detectado: {last_error}")

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
        st.warning("La Base de Datos no está configurada. No se puede consultar la publicación vigente.")
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
                    df = filtrar_resumen_total_plazas_positivo(df)
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



def render_admin_calidad_datos() -> None:
    st.subheader("Calidad de datos")
    st.info(
        "Esta pantalla no valida el contenido material declarado por cada centro, pero ayuda a detectar situaciones que conviene revisar: "
        "centros pendientes, expedientes finalizados sin registros, duplicados, valores altos y cambios entre publicaciones."
    )

    if not supabase_available():
        st.warning("La Base de Datos no está configurada. No se pueden generar validaciones de calidad.")
        return

    ok, msg, quality = build_quality_from_supabase()
    if not ok or not quality:
        st.warning(msg)
        return

    resumen = quality.get("Calidad_Resumen", pd.DataFrame())
    if not resumen.empty:
        alertas = int((resumen["Nivel"].astype(str).isin(["Alerta", "Revisar", "Aviso"])).sum()) if "Nivel" in resumen.columns else 0
        c1, c2, c3 = st.columns(3)
        c1.metric("Controles ejecutados", len(resumen))
        c2.metric("Controles con aviso/revisión", alertas)
        c3.metric("Titulaciones sin plazas", int(resumen.loc[resumen["Control"] == "Titulaciones del catálogo sin plazas", "Resultado"].iloc[0]) if "Control" in resumen.columns and (resumen["Control"] == "Titulaciones del catálogo sin plazas").any() else 0)
        st.dataframe(resumen, use_container_width=True, hide_index=True)

    tabs = st.tabs(["Pendientes", "Borrador posterior", "Finalizados 0 registros", "Duplicados", "Valores altos", "Sin plazas", "Comparativa publicaciones"])
    tab_names = [
        "Calidad_Pendientes",
        "Calidad_Borrador_Posterior",
        "Calidad_Finalizados_0reg",
        "Calidad_Duplicados",
        "Calidad_Valores_Altos",
        "Calidad_Sin_Plazas",
    ]
    for tab, sheet in zip(tabs[:6], tab_names):
        with tab:
            df = quality.get(sheet, pd.DataFrame())
            if df is None or df.empty:
                st.success("Sin incidencias para este control.")
            else:
                st.dataframe(df, use_container_width=True, hide_index=True)
                st.caption(f"{len(df)} filas encontradas.")

    with tabs[6]:
        ok_comp, msg_comp, comp = compare_latest_publications()
        if not ok_comp or not comp:
            st.info(msg_comp)
        else:
            st.markdown("#### Resumen comparativo")
            st.dataframe(comp.get("Comparativa_Resumen", pd.DataFrame()), use_container_width=True, hide_index=True)
            st.markdown("#### Cambios por titulación")
            cambios = comp.get("Comparativa_Cambios", pd.DataFrame())
            if cambios.empty:
                st.success("No se detectan cambios entre las dos últimas publicaciones.")
            else:
                st.dataframe(cambios, use_container_width=True, hide_index=True)

    st.markdown("---")
    if st.button("Registrar revisión de calidad"):
        audit_event("revision_calidad_datos", "Revisión manual de calidad ejecutada desde panel admin")
        st.success("Revisión registrada en auditoría.")

def render_admin_historial_versiones() -> None:
    st.subheader("Historial de versiones")
    st.caption("Información visible solo para usuarios administradores.")
    st.info(f"Desarrollador / creador del programa: {APP_AUTHOR}. Responsable funcional del proyecto: {APP_CREATOR}.")
    versiones = [
        ("DCD 1.0", "MVP inicial con contraseña, instrucciones, selección de centro docente, selectores dependientes y preparación para Base de Datos."),
        ("DCD 1.0.1", "Pantalla de recordatorio, correo automático opcional, usuarios configurables, auditoría y revisión de mapeo."),
        ("DCD 1.0.2", "Ajuste de áreas: Hospital, Atención Familiar y Comunitaria, y retirada de Otras Unidades Docentes."),
        ("DCD 1.0.3", "Cierre estable: exportación más completa, estado finalizado y bloqueo suave de edición."),
        ("DCD 1.0.4", "Totales en Matriz_DCD y panel administrador para Excel consolidado desde Base de Datos."),
        ("DCD 1.0.5", "Usuarios por centro docente, contraseñas hasheadas, reset por admin y borradores filtrados."),
        ("DCD 1.0.5.1", "Corrección del selector de centro docente al crear usuarios."),
        ("DCD 1.0.6", "Guardado versionado, control de centros pendientes y cambio visible a Centros Docentes."),
        ("DCD 1.0.7", "Publicaciones oficiales: PDF Matriz_DCD, histórico/vigente, Storage y notificación al administrador."),
        ("DCD 1.0.8", "Dashboard y análisis de publicación: resúmenes por provincia, isla, centro, rama, nivel y titulaciones."),
        ("DCD 1.0.8.1", "Cierre configurable: fecha tope, modos de cierre, avisos previos y publicación automática por vencimiento si procede."),
        ("DCD 1.0.8.2", "Refuerzo de evaluación de cierre al acceder cualquier usuario, al finalizar centros y desde botón admin."),
        ("DCD 1.0.9", "Portal externo de consulta de publicación vigente con dashboard y descargas limitadas."),
        ("DCD 1.0.9.1", "Ajustes de interfaz del portal y mantenimiento avanzado de usuarios."),
        ("DCD 1.1.0", "Mejora visual del dashboard, tarjetas compactas y PDF preparado para logo/frase institucional."),
        ("DCD 1.1.1", "Auditoría ampliada, registro de accesos/descargas y backup completo admin."),
        ("DCD 1.1.1.1", "Corrección de políticas RLS de auditoría y diagnóstico manual de eventos."),
        ("DCD 1.1.2", "Calidad de datos: avisos, validaciones, valores atípicos y comparativa entre publicaciones."),
        ("DCD 1.1.2.1", "Pulido de interfaz, ocultación de historial en login y lenguaje no técnico para usuarios."),
        ("DCD 1.1.2.2", "Ajustes finales de PDF: logo, frase institucional, pie de firma y numeración de páginas."),
        ("DCD 1.1.2.3", "Excel limitado para usuarios de consulta y ajuste de logo superior derecho en PDF."),
        ("DCD 1.1.3", "Cierre documental y formalización de autoría: cabecera de código, README, CHANGELOG, AUTHORSHIP y huella SHA256 del paquete."),
        ("DCD 1.1.3.1", "Corrección GAP TF: centro docente único para Atención Primaria de Tenerife."),
        ("DCD 1.1.3.2", "Corrección rol consulta: preparación de Excel limitado sin error NameError."),
        ("DCD 1.1.3.3", "Ajustes finales de visualización: mensaje de correo y ocultación de filas con total 0 en dashboard/PDF."),
        ("DCD 1.1.3.4", "Mapa visual de capacidad docente por isla en el dashboard del rol consulta."),
        ("DCD 1.1.3.5", "Corrección PDF de consulta: sin calidad interna, sin filas total 0 y ajuste de coordenadas del mapa."),
        ("DCD 1.1.3.6", "Ajuste fino de los puntos azules del mapa de Canarias en el dashboard de consulta."),
        ("DCD 1.1.3.7", "Ajuste fino final de puntos y tarjetas del mapa de Canarias tras pilotaje."),
        ("DCD 1.1.3.8", "Microajuste final de tarjetas del mapa de Canarias (La Gomera, Tenerife y Fuerteventura)."),
        ("DCD 1.1.3.9", "Microajuste final adicional del recuadro de Fuerteventura en el mapa de Canarias."),
        ("DCD 1.2.0 beta 1", "Nueva fase funcional: captura de turnos y observaciones por titulación, con persistencia y exportación inicial."),
        ("DCD 1.2.0 beta 2", "Edición de registros del Paso 4 y pulido del PDF de consulta para turnos y observaciones."),
        ("DCD 1.2.0 beta 3", "Corrección de carga de registros en modo edición antes del renderizado de widgets Streamlit."),
        ("DCD 1.2.0 RC1", "Versión candidata estable: turnos y observaciones validados en rol usuario, admin, consulta, Excel y PDF."),
        ("DCD 1.2.1 beta 1", "Nueva fase: configuración admin de centros multiusuario, usuarios previstos y asignaciones por centro."),
        ("DCD 1.2.1 beta 2", "Estado de aportaciones por usuario, trazabilidad de aportación y entrada al consolidado solo cuando el centro multiusuario está completo."),
        ("DCD 1.2.1 beta 3", "Corrección de valores NaN al registrar publicaciones multiusuario en Supabase."),
        ("DCD 1.2.1 beta 4", "Consolidación parcial manual por admin para centros multiusuario incompletos."),
        ("DCD 1.2.2 beta 1", "Servicios y curso/año voluntarios por alumno, con exportación a Excel en hoja Detalle_Alumnos."),
        ("DCD 1.2.2 estable", "Ajuste de visualización de Detalle alumnos en la tabla de registros introducidos."),
    ]
    df_versiones = pd.DataFrame(versiones, columns=["Versión", "Cambios principales"])
    st.dataframe(df_versiones, use_container_width=True, hide_index=True)

def page_admin_consolidado() -> None:
    st.header("Panel administrador")

    if st.session_state.get("current_user_role") != "admin":
        st.error("Esta pantalla solo está disponible para usuarios administradores.")
        if st.button("Volver"):
            st.session_state.current_step = 2
            st.rerun()
        return

    tab_consolidado, tab_publicaciones, tab_cierre, tab_calidad, tab_usuarios, tab_multiusuario, tab_auditoria, tab_historial = st.tabs(["Consolidado", "Publicaciones", "Cierre", "Calidad de datos", "Usuarios", "Multiusuario", "Auditoría/Backup", "Historial versiones"])
    with tab_consolidado:
        render_admin_consolidado()
    with tab_publicaciones:
        render_admin_publicaciones()
    with tab_cierre:
        render_admin_cierre()
    with tab_calidad:
        render_admin_calidad_datos()
    with tab_usuarios:
        render_admin_usuarios()
    with tab_multiusuario:
        render_admin_multiusuario_centros()
    with tab_auditoria:
        render_admin_auditoria_backup()
    with tab_historial:
        render_admin_historial_versiones()

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
