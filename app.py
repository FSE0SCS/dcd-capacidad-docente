import io
import json
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
import streamlit as st

try:
    from supabase import create_client
except Exception:
    create_client = None


# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================
APP_VERSION = "DCD 1.0.4"
APP_TITLE = "DATOS CAPACIDAD DOCENTE (DCD 1.0)"
DEFAULT_PASSWORD = "Capacidad2026"
EXCEL_PATH = Path(__file__).parent / "data" / "listado_para_capacidad_docente.xlsx"

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


# =========================================================
# UTILIDADES DE ESTADO
# =========================================================
def init_session_state() -> None:
    defaults = {
        "logged_in": False,
        "current_user": "",
        "current_user_role": "",
        "current_user_display": "",
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


def login_user(username: str, password: str) -> tuple[bool, str]:
    users = get_users_config()
    username = (username or "").strip()
    password = password or ""

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
        f"Unidad docente: {st.session_state.direccion_selected}\n"
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

    for item in st.session_state.registros.values():
        rows.append({
            "codigo_borrador": codigo_borrador,
            "estado": estado,
            "area": area,
            "unidad_docente": unidad,
            "codigo_unidad": codigo_unidad,
            "columna_excel": columna_excel,
            "nivel_i": item["Nivel Estudio I"],
            "nivel_ii": item["Nivel Estudio II"],
            "rama": item["Rama"],
            "titulacion": item["Titulación"],
            "numero_alumnos": int(item["Nº alumnos"]),
        })
    return rows


def save_draft_to_supabase(estado: str = "borrador") -> tuple[bool, str]:
    client = get_supabase_client()
    if client is None:
        return False, "Supabase no está configurado todavía. Puedes seguir usando el MVP local y descargar el Excel."

    unidad = st.session_state.direccion_selected
    codigo_borrador = st.session_state.codigo_borrador or build_codigo_borrador(unidad)
    st.session_state.codigo_borrador = codigo_borrador
    codigo_unidad = CODIGOS_DIRECCION.get(unidad, safe_code(unidad))

    try:
        client.table("dcd_borradores").upsert(
            {
                "codigo_borrador": codigo_borrador,
                "app_version": APP_VERSION,
                "estado": estado,
                "area": st.session_state.area_selected,
                "unidad_docente": unidad,
                "codigo_unidad": codigo_unidad,
                "observaciones": st.session_state.get("observaciones", ""),
            },
            on_conflict="codigo_borrador",
        ).execute()

        client.table("dcd_registros").delete().eq("codigo_borrador", codigo_borrador).execute()
        rows = registros_to_rows(estado=estado)
        if rows:
            client.table("dcd_registros").insert(rows).execute()

        st.session_state.draft_estado = estado
        if estado == "finalizado":
            st.session_state.permitir_editar_finalizado = False

        audit_event("guardar_borrador" if estado == "borrador" else "guardar_finalizado", f"Estado: {estado}. Registros: {len(rows)}", codigo_borrador)
        if estado == "finalizado":
            return True, f"Expediente finalizado correctamente: {codigo_borrador}"
        return True, f"Borrador guardado correctamente: {codigo_borrador}"
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
        registros_resp = client.table("dcd_registros").select("*").eq("codigo_borrador", codigo_borrador).execute()
        rows = getattr(registros_resp, "data", []) or []

        st.session_state.area_selected = normalizar_area(borrador.get("area", ""))
        st.session_state.direccion_selected = borrador.get("unidad_docente", "")
        st.session_state.codigo_borrador = codigo_borrador
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
        resp = client.table("dcd_borradores").select("codigo_borrador, updated_at, estado").order("updated_at", desc=True).limit(50).execute()
        rows = getattr(resp, "data", []) or []
        return [f"{r['codigo_borrador']} | {r.get('estado', '')} | {r.get('updated_at', '')}" for r in rows]
    except Exception:
        return []


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
        "Unidad docente": unidad,
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
            registros_df.insert(3, "Unidad docente", unidad)
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

    resp = client.table("dcd_borradores").select("*").eq("estado", "finalizado").order("updated_at", desc=True).execute()
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


def build_consolidated_excel_from_supabase() -> tuple[bool, str, bytes | None, str]:
    client = get_supabase_client()
    if client is None:
        return False, "Supabase no está configurado.", None, ""

    try:
        borradores, duplicados = get_latest_finalized_drafts()
        if not borradores:
            return False, "No hay expedientes finalizados para consolidar.", None, ""

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
                    "Fecha actualización": borrador.get("updated_at", ""),
                    "Área": normalizar_area(borrador.get("area", "")),
                    "Unidad docente": unidad,
                    "Columna Excel": columna_excel,
                    **item,
                })

        registros_df = pd.DataFrame(registros_consolidados)
        if registros_df.empty:
            registros_df = pd.DataFrame(columns=[
                "Código borrador",
                "Fecha actualización",
                "Área",
                "Unidad docente",
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
            "Duplicados finalizados ignorados": len(duplicados),
            "Criterio": "Último expediente finalizado por unidad docente según updated_at",
            "Usuario generación": st.session_state.get("current_user_display", "") or st.session_state.get("current_user", ""),
        }])

        borradores_df = pd.DataFrame([{
            "Código borrador": b.get("codigo_borrador", ""),
            "Estado": b.get("estado", ""),
            "Área": normalizar_area(b.get("area", "")),
            "Unidad docente": b.get("unidad_docente", ""),
            "Código unidad": b.get("codigo_unidad", ""),
            "Actualizado": b.get("updated_at", ""),
            "Columna Excel": COLUMNA_EXCEL_POR_DIRECCION.get(b.get("unidad_docente", ""), ""),
        } for b in borradores])

        duplicados_df = pd.DataFrame([{
            "Código borrador ignorado": b.get("codigo_borrador", ""),
            "Área": normalizar_area(b.get("area", "")),
            "Unidad docente": b.get("unidad_docente", ""),
            "Actualizado": b.get("updated_at", ""),
        } for b in duplicados])
        if duplicados_df.empty:
            duplicados_df = pd.DataFrame(columns=["Código borrador ignorado", "Área", "Unidad docente", "Actualizado"])

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            sheets = {
                "Resumen": resumen,
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

    st.markdown("---")
    st.markdown("##### Historial de versiones")
    st.markdown("- **DCD 1.0:** MVP inicial con contraseña, instrucciones, selección de unidad docente, selectores dependientes y preparación para Supabase.")
    st.markdown("- **DCD 1.0.1:** Pantalla de recordatorio, correo automático opcional, usuarios configurables, auditoría y revisión de mapeo.")
    st.markdown("- **DCD 1.0.2:** Ajuste de áreas: Hospital, Atención Familiar y Comunitaria, y retirada de Otras Unidades Docentes.")
    st.markdown("- **DCD 1.0.3:** Cierre estable: exportación más completa, estado finalizado y bloqueo suave de edición.")
    st.markdown("- **DCD 1.0.4:** Totales en Matriz_DCD y panel administrador para Excel consolidado desde Supabase.")


def page_instrucciones() -> None:
    st.header("Paso 1: Información, instrucciones y aceptación")
    st.markdown(
        """
        **Bienvenido al aplicativo DATOS CAPACIDAD DOCENTE (DCD 1.0).**

        Este aplicativo tiene como finalidad recoger datos de capacidad docente por unidad docente, nivel de estudios,
        rama y titulación.

        **Instrucciones básicas:**

        1. Seleccione primero el área y la unidad docente correspondiente.
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
    st.header("Paso 2: Selección de Área y Unidad Docente")

    st.session_state.area_selected = normalizar_area(st.session_state.area_selected)

    st.session_state.area_selected = st.selectbox(
        "**SELECCIONE ÁREA**",
        options=[""] + AREA_OPTIONS,
        index=([""] + AREA_OPTIONS).index(st.session_state.area_selected) if st.session_state.area_selected in ([""] + AREA_OPTIONS) else 0,
    )

    direccion_options = DIRECCIONES_POR_AREA.get(st.session_state.area_selected, [])
    if st.session_state.direccion_selected not in direccion_options:
        st.session_state.direccion_selected = ""

    st.session_state.direccion_selected = st.selectbox(
        "**SELECCIONE UNIDAD DOCENTE / DIRECCIÓN / GERENCIA**",
        options=[""] + direccion_options,
        index=([""] + direccion_options).index(st.session_state.direccion_selected) if st.session_state.direccion_selected in ([""] + direccion_options) else 0,
    )

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Siguiente"):
            if st.session_state.area_selected and st.session_state.direccion_selected:
                nuevo_codigo = build_codigo_borrador(st.session_state.direccion_selected)
                if st.session_state.codigo_borrador and st.session_state.codigo_borrador != nuevo_codigo:
                    st.session_state.registros = {}
                    st.session_state.observaciones = ""
                    st.session_state.draft_estado = "borrador"
                    st.session_state.permitir_editar_finalizado = False
                    reset_selectores_estudio()
                st.session_state.codigo_borrador = nuevo_codigo
                st.session_state.current_step = 3
                st.rerun()
            else:
                st.warning("Debe seleccionar un área y una unidad docente para continuar.")
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
    st.markdown(f"**Unidad docente:** <span style='color:#0d6efd'>{unidad}</span>", unsafe_allow_html=True)
    st.markdown(f"**Código borrador:** `{st.session_state.codigo_borrador or build_codigo_borrador(unidad)}`")

    if columna_excel:
        st.success(f"La unidad seleccionada se corresponde con la columna del Excel: **{columna_excel}**")
    else:
        st.warning(
            "Esta unidad no tiene una columna directa identificada en el Excel base. "
            "Se podrán guardar registros y generar hoja de registros, pero quizá haya que revisar el mapeo para la matriz final."
        )

    with st.expander("Revisión del mapeo unidad docente ↔ columna Excel"):
        mapping_rows = []
        for direccion, codigo in CODIGOS_DIRECCION.items():
            mapping_rows.append({
                "Unidad docente": direccion,
                "Código interno": codigo,
                "Columna Excel": COLUMNA_EXCEL_POR_DIRECCION.get(direccion, "PENDIENTE DE MAPEAR"),
            })
        st.dataframe(pd.DataFrame(mapping_rows), use_container_width=True, hide_index=True)
        st.caption("Esta tabla permite revisar qué unidades vuelcan datos directamente en la matriz Excel y cuáles quedan pendientes de correspondencia definitiva.")

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

    st.write(f"Unidad docente seleccionada: **{st.session_state.direccion_selected}**")
    st.write(f"Código de borrador: `{st.session_state.codigo_borrador or build_codigo_borrador(st.session_state.direccion_selected)}`")
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
        "Antes de cerrar el expediente, revise que la unidad docente, la titulación y el número de alumnos son correctos. "
        "Si falta alguna titulación o está esperando datos de alguna especialidad, use 'REVISAR / VOLVER' o 'Guardar borrador'."
    )

    unidad = st.session_state.direccion_selected
    columna_excel = COLUMNA_EXCEL_POR_DIRECCION.get(unidad, "")
    if columna_excel:
        st.info(f"Los datos de esta unidad se volcarán en la columna del Excel: **{columna_excel}**.")
    else:
        st.error(
            "Esta unidad no tiene columna Excel mapeada. Se generará el Excel con la hoja de registros, "
            "pero la matriz principal puede no quedar volcada en una columna específica."
        )

    check1 = st.checkbox("He revisado que el Área y la Unidad Docente son correctas", key="check_recordatorio_unidad")
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
                else:
                    st.warning(msg_save)
                ok_mail, msg_mail = send_email_with_mailgun(excel_bytes, filename)
                if ok_mail:
                    st.success(msg_mail)
                else:
                    st.warning(msg_mail)

    st.markdown("---")
    st.caption("Si Mailgun no está configurado, el botón de correo mostrará un aviso y podrá seguir usando la descarga manual.")


def page_admin_consolidado() -> None:
    st.header("Panel administrador: consolidado DCD")

    if st.session_state.get("current_user_role") != "admin":
        st.error("Esta pantalla solo está disponible para usuarios administradores.")
        if st.button("Volver"):
            st.session_state.current_step = 2
            st.rerun()
        return

    st.markdown(
        """
        Esta pantalla genera un Excel consolidado a partir de los expedientes guardados como **finalizado** en Supabase.

        Criterio aplicado:
        - Se toma el último expediente finalizado por cada unidad docente.
        - Si dos unidades vuelcan en la misma columna, sus registros se suman en la matriz.
        - Las columnas de totales se calculan automáticamente:
          H = E+F+G, K = H+I+J, O = L+M+N, S = O+P+Q+R y T = K+S.
        """
    )

    if not supabase_available():
        st.warning("Supabase no está configurado. No se puede generar el consolidado.")
        return

    borradores, duplicados = get_latest_finalized_drafts()
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Expedientes finalizados usados", len(borradores))
    with col2:
        st.metric("Duplicados finalizados ignorados", len(duplicados))

    if borradores:
        preview = pd.DataFrame([{
            "Código borrador": b.get("codigo_borrador", ""),
            "Área": normalizar_area(b.get("area", "")),
            "Unidad docente": b.get("unidad_docente", ""),
            "Actualizado": b.get("updated_at", ""),
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
                "Unidad docente": b.get("unidad_docente", ""),
                "Actualizado": b.get("updated_at", ""),
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

st.title(APP_TITLE)
st.markdown("---")
app_sidebar()

try:
    # Precarga para detectar problemas del Excel cuanto antes.
    _ = load_catalogo()
except Exception as exc:
    st.error(f"No se pudo cargar el Excel base: {exc}")
    st.stop()

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
