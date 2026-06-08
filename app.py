import io
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

try:
    from supabase import create_client
except Exception:
    create_client = None


# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================
APP_VERSION = "DCD 1.0"
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
    "UNIDAD DOCENTE DE CENTRO HOSPITALARIO",
    "UNIDAD DOCENTE MULTIPROFESIONAL DE ATENCION FAMILIAR Y COMUNITARIA",
    "OTRAS UNIDADES DOCENTES",
]

DIRECCIONES_POR_AREA = {
    "UNIDAD DOCENTE DE CENTRO HOSPITALARIO": [
        "DIRECCIÓN GERENCIA HOSPITAL DOCTOR NEGRIN",
        "DIRECCIÓN GERENCIA COMPLEJO HOSPITALARIO UNIVERSITARIO Y MATERNO INFANTIL",
        "DIRECCIÓN GERENCIA COMPLEJO HOSPITALARIO UNIVERSITARIO DE CANARIAS",
        "DIRECCIÓN GERENCIA HOSPITAL NUESTRA SEÑORA DE CANDELARIA",
    ],
    "UNIDAD DOCENTE MULTIPROFESIONAL DE ATENCION FAMILIAR Y COMUNITARIA": [
        "GERENCIA DE ATENCIÓN PRIMARIA DE GRAN CANARIA",
        "GERENCIA DE ATENCIÓN PRIMARIA DE TENERIFE NORTE",
        "GERENCIA DE ATENCIÓN PRIMARIA DE TENERIFE SUR",
        "GERENCIA DE SERVICIOS SANITARIOS DE FUERTEVENTURA",
        "GERENCIA DE SERVICIOS SANITARIOS DE LANZAROTE",
        "GERENCIA DE SERVICIOS SANITARIOS DE LA PALMA",
        "GERENCIA DE SERVICIOS SANITARIOS DE LA GOMERA",
        "GERENCIA DE SERVICIOS SANITARIOS DE EL HIERRO",
    ],
    "OTRAS UNIDADES DOCENTES": [
        "UNIDAD DOCENTE MULTIPROFESIONAL DE SALUD MENTAL DE GRAN CANARIA",
        "UNIDAD DOCENTE MULTIPROFESIONAL DE SALUF MENTAL DE TENERIFE",
        "UNIDAD DOCENTE MULTIPROFESIONAL DE SALUD LABORAL",
        "UNIDAD DOCENTE MULTIPROFESIONAL DE OBSTETRICIA Y GINECOLOGIA DEL CHUIMI",
        "UNIDAD DOCENTE MULTIPROFESIONAL DE PEDIATRIA DEL CHUIMI",
        "UNIDAD DOCENTE MULTIPROFESIONAL DE PEDIATRIA DEL HUNSC",
        "UNIDAD DOCENTE MULTIPROFESIONAL DE PEDIATRIA DEL HUC",
        "UNIDAD DOCENTE DE ENFERMERIA OBSTETRICA-GINECOLOGICA DE TENERIFE",
        "UNIDAD DOCENTE DE MEDICINA PREVENTIVA Y SALUD PUBLICA",
    ],
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


# =========================================================
# UTILIDADES DE ESTADO
# =========================================================
def init_session_state() -> None:
    defaults = {
        "logged_in": False,
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

        st.session_state.area_selected = borrador.get("area", "")
        st.session_state.direccion_selected = borrador.get("unidad_docente", "")
        st.session_state.codigo_borrador = codigo_borrador
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
def build_output_excel() -> bytes:
    catalogo = load_catalogo()
    matriz = catalogo.copy()
    registros = list(st.session_state.registros.values())
    registros_df = pd.DataFrame(registros)

    unidad = st.session_state.direccion_selected
    columna_excel = COLUMNA_EXCEL_POR_DIRECCION.get(unidad, "")

    if columna_excel and columna_excel in matriz.columns and registros:
        # Limpiamos solo la columna seleccionada y volcamos los registros del usuario.
        matriz[columna_excel] = pd.NA
        for item in registros:
            mask = (
                (matriz["Nivel Estudio I"] == item["Nivel Estudio I"])
                & (matriz["Nivel Estudio II"] == item["Nivel Estudio II"])
                & (matriz["Rama"] == item["Rama"])
                & (matriz["Titulación"] == item["Titulación"])
            )
            matriz.loc[mask, columna_excel] = int(item["Nº alumnos"])

    resumen = pd.DataFrame([{
        "Aplicativo": APP_VERSION,
        "Fecha generación": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Área": st.session_state.area_selected,
        "Unidad docente": unidad,
        "Código unidad": CODIGOS_DIRECCION.get(unidad, safe_code(unidad)),
        "Columna Excel": columna_excel,
        "Código borrador": st.session_state.codigo_borrador or build_codigo_borrador(unidad),
        "Nº registros": len(registros),
        "Total alumnos": sum(int(x["Nº alumnos"]) for x in registros) if registros else 0,
    }])

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        if registros_df.empty:
            registros_df = pd.DataFrame(columns=KEY_COLUMNS + ["Nº alumnos"])
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

    output.seek(0)
    return output.getvalue()


# =========================================================
# COMPONENTES DE INTERFAZ
# =========================================================
def app_sidebar() -> None:
    st.sidebar.markdown("---")
    st.sidebar.markdown("### ℹ️ Información")
    st.sidebar.write(f"Versión: {APP_VERSION}")
    st.sidebar.write("Desarrollado para: F.S.E. – S.C.S.")

    if supabase_available():
        st.sidebar.success("Supabase configurado")
    else:
        st.sidebar.warning("Supabase no configurado")
        st.sidebar.caption("El aplicativo funciona como MVP local, pero no guarda borradores entre sesiones.")

    st.sidebar.markdown("---")
    if st.sidebar.button("Salir del aplicativo 🚪"):
        st.session_state.clear()
        st.rerun()


def page_login() -> None:
    st.title(APP_TITLE)
    st.subheader("🔐 Acceso al aplicativo")
    st.write("Introduce la contraseña para continuar.")

    password = st.text_input("Contraseña", type="password", key="password_input")
    if st.button("Iniciar sesión"):
        app_password = get_secret("APP_PASSWORD", DEFAULT_PASSWORD)
        if password == app_password:
            st.session_state.logged_in = True
            st.session_state.current_step = 1
            st.rerun()
        else:
            st.error("Contraseña incorrecta. Por favor, inténtalo de nuevo.")

    st.markdown("---")
    st.markdown("##### Historial de versiones")
    st.markdown("- **DCD 1.0:** MVP inicial con contraseña, instrucciones, selección de unidad docente, selectores dependientes y preparación para Supabase.")


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

        **Advertencias:**

        - Revise bien la titulación seleccionada antes de añadirla.
        - Si introduce de nuevo una misma combinación, se actualizará el número de alumnos.
        - En esta primera versión, el envío por correo queda pendiente para una fase posterior.
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
                st.session_state.codigo_borrador = build_codigo_borrador(st.session_state.direccion_selected)
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

    st.write(f"Unidad docente seleccionada: **{st.session_state.direccion_selected}**")
    st.write(f"Código de borrador: `{st.session_state.codigo_borrador or build_codigo_borrador(st.session_state.direccion_selected)}`")

    st.markdown("### Añadir o actualizar titulación")

    nivel_i_options = sorted_unique(catalogo["Nivel Estudio I"])
    st.selectbox(
        "Nivel Estudio I",
        options=[""] + nivel_i_options,
        key="sel_nivel_i",
        on_change=reset_downstream,
        args=("nivel_i",),
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
        disabled=not bool(st.session_state.sel_nivel_i),
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
        disabled=not bool(st.session_state.sel_nivel_ii),
    )

    df4 = df3[df3["Rama"] == st.session_state.sel_rama] if st.session_state.sel_rama else df3.iloc[0:0]
    titulacion_options = sorted_unique(df4["Titulación"]) if not df4.empty else []
    if st.session_state.sel_titulacion not in titulacion_options:
        st.session_state.sel_titulacion = ""
    st.selectbox(
        "Titulación",
        options=[""] + titulacion_options,
        key="sel_titulacion",
        disabled=not bool(st.session_state.sel_rama),
    )

    st.number_input(
        "Número de alumnos",
        min_value=0,
        step=1,
        key="numero_alumnos",
        disabled=not bool(st.session_state.sel_titulacion),
    )

    col_add, col_clear = st.columns(2)
    with col_add:
        if st.button("Añadir / actualizar registro"):
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
                st.success("Registro añadido/actualizado correctamente.")
    with col_clear:
        if st.button("Limpiar selectores"):
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
        delete_label = st.selectbox("Seleccionar registro para eliminar", options=[""] + registro_labels)
        if st.button("Eliminar registro seleccionado"):
            if delete_label:
                idx = int(delete_label.split(".", 1)[0]) - 1
                item = registros[idx]
                key = registro_key(item["Nivel Estudio I"], item["Nivel Estudio II"], item["Rama"], item["Titulación"])
                st.session_state.registros.pop(key, None)
                st.success("Registro eliminado.")
                st.rerun()
            else:
                st.warning("Debe seleccionar un registro.")
    else:
        st.info("Todavía no se ha introducido ningún registro.")

    st.markdown("---")
    st.text_area("Observaciones internas del borrador", key="observaciones")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        if st.button("Guardar borrador"):
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
        if st.button("Reiniciar datos"):
            st.session_state.registros = {}
            reset_selectores_estudio()
            st.rerun()


def page_resumen_descarga() -> None:
    st.header("Paso 5: Resumen, guardado final y descarga")
    registros = list(st.session_state.registros.values())
    if not registros:
        st.warning("No hay registros para generar.")
        if st.button("Volver a introducir datos"):
            st.session_state.current_step = 4
            st.rerun()
        return

    df = pd.DataFrame(registros)
    st.dataframe(df, use_container_width=True, hide_index=True)

    total_alumnos = int(df["Nº alumnos"].sum())
    st.metric("Total alumnos introducidos", total_alumnos)

    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("Guardar como finalizado en Supabase"):
            ok, msg = save_draft_to_supabase(estado="finalizado")
            if ok:
                st.success(msg)
            else:
                st.warning(msg)
    with col2:
        excel_bytes = build_output_excel()
        filename = f"{st.session_state.codigo_borrador or build_codigo_borrador(st.session_state.direccion_selected)}.xlsx"
        st.download_button(
            label="Descargar Excel",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col3:
        if st.button("REVISAR / VOLVER"):
            st.session_state.current_step = 4
            st.rerun()

    st.markdown("---")
    st.warning("La pantalla de recordatorio y envío automático por correo queda pendiente para la siguiente fase.")


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
else:
    st.session_state.current_step = 1
    st.rerun()
