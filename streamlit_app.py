import os
import re
import io
import json
import shutil
import sqlite3
import tempfile
import subprocess
import base64
import zipfile
from xml.sax.saxutils import escape
from pathlib import Path
from datetime import datetime, date
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================
st.set_page_config(
    page_title="Cotizaciones Charles Servicio Automotriz",
    page_icon="🔧",
    layout="wide",
)

try:
    import gspread
    from google.oauth2.service_account import Credentials
except Exception:
    gspread = None
    Credentials = None

APP_DIR = Path(__file__).parent
DATA_DIR = APP_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)

TEMPLATE_FILE = APP_DIR / "plantilla_cotizacion_charles_profesional.docx"
LOGO_FILE = APP_DIR / "marca charles.png"
LOGO_FONDO = APP_DIR / "logo charles blanco 21 jun 2025, 23_06_15.png"
DB_FILE = DATA_DIR / "cotizaciones_charles.db"
EXCEL_FILE = DATA_DIR / "base_cotizaciones_charles.xlsx"

PREFIJO_COTIZACION = "CSA"
ESTADOS = ["Pendiente", "Aceptada", "Rechazada", "En revisión", "Vencida", "Anulada por error"]
CATEGORIAS = ["Mano de obra", "Repuestos", "Otros"]

CONDICIONES_DEFAULT = """- Validez de la cotización: 10 días hábiles.
- Forma de pago: transferencia / efectivo.
- Valores expresados en pesos chilenos.
- Cotización sujeta a disponibilidad de repuestos."""

CATALOGO_DEFAULT = [
    # Mano de obra
    ("Mano de obra", "Diagnóstico scanner", 25000),
    ("Mano de obra", "Revisión general", 20000),
    ("Mano de obra", "Mantención preventiva", 45000),
    ("Mano de obra", "Cambio aceite motor", 20000),
    ("Mano de obra", "Revisión sistema eléctrico", 30000),
    ("Mano de obra", "Cambio pastillas de freno", 45000),
    ("Mano de obra", "Reparación sistema de frenos", 65000),
    ("Mano de obra", "Cambio amortiguadores", 60000),
    ("Mano de obra", "Revisión suspensión", 30000),
    ("Mano de obra", "Cambio batería", 15000),
    # Repuestos
    ("Repuestos", "Aceite motor", 0),
    ("Repuestos", "Filtro de aceite", 0),
    ("Repuestos", "Filtro de aire", 0),
    ("Repuestos", "Pastillas de freno", 0),
    ("Repuestos", "Bujías", 0),
    ("Repuestos", "Batería", 0),
    ("Repuestos", "Amortiguador", 0),
    ("Repuestos", "Correa accesorios", 0),
    ("Repuestos", "Ampolleta", 0),
    # Otros
    ("Otros", "Insumos de taller", 5000),
    ("Otros", "Traslado", 15000),
    ("Otros", "Lavado técnico", 10000),
    ("Otros", "Servicio externo", 0),
    ("Otros", "Rectificación", 0),
    ("Otros", "Scanner externo", 25000),
]

HEADERS_COTIZACIONES = [
    "id", "numero_cotizacion", "correlativo", "fecha", "estado",
    "cliente", "contacto_cel", "mail_cliente",
    "patente", "marca", "modelo", "anio", "kilometraje", "vin",
    "subtotal_mano_obra", "subtotal_repuestos", "subtotal_otros", "total",
    "observaciones", "condiciones", "creado_en", "actualizado_en",
]

HEADERS_DETALLE = [
    "id_item", "cotizacion_id", "numero_cotizacion", "categoria",
    "descripcion", "cantidad", "valor_unitario", "total",
]

HEADERS_CATALOGO = ["id_cat", "categoria", "descripcion", "valor_sugerido", "activo", "actualizado_en"]

# =========================================================
# ESTILO VISUAL
# =========================================================
def img_to_base64(path: Path) -> str:
    if not path.exists():
        return ""
    return base64.b64encode(path.read_bytes()).decode("utf-8")


def aplicar_estilo():
    fondo = img_to_base64(LOGO_FONDO)
    marca = ""
    if fondo:
        marca = f"""
        .stApp::after {{
            content: "";
            position: fixed;
            top: 58%;
            left: 50%;
            transform: translate(-50%, -50%);
            width: 520px;
            height: 520px;
            background-image: url("data:image/png;base64,{fondo}");
            background-repeat: no-repeat;
            background-position: center;
            background-size: contain;
            opacity: 0.035;
            z-index: 0;
            pointer-events: none;
        }}
        """
    st.markdown(
        f"""
        <style>
        .stApp {{ background-color: #ffffff; }}
        {marca}
        .block-container {{ padding-top: 1.2rem; }}
        h1, h2, h3 {{ color: #111827; }}
        .charles-title {{ font-size: 34px; font-weight: 800; color: #111827; margin-bottom: 0px; }}
        .charles-subtitle {{ color: #6b7280; font-size: 15px; margin-top: 0px; }}
        .metric-card {{
            border: 1px solid #e5e7eb; border-radius: 12px; padding: 16px;
            background: rgba(255,255,255,0.88);
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def encabezado():
    c1, c2 = st.columns([1, 5])
    with c1:
        if LOGO_FILE.exists():
            st.image(str(LOGO_FILE), width=140)
    with c2:
        st.markdown("<div class='charles-title'>Cotizaciones Charles Servicio Automotriz</div>", unsafe_allow_html=True)
        st.markdown("<div class='charles-subtitle'>Generación de cotizaciones, historial, correlativos, estados y base comercial.</div>", unsafe_allow_html=True)
    st.divider()

# =========================================================
# UTILIDADES
# =========================================================
def ahora_santiago() -> datetime:
    return datetime.now(ZoneInfo("America/Santiago"))


def fecha_hoy() -> date:
    return ahora_santiago().date()


def clp_fmt(valor) -> str:
    try:
        v = float(valor or 0)
        return "$ " + f"{v:,.0f}".replace(",", ".")
    except Exception:
        return "$ 0"


def to_float(valor) -> float:
    try:
        if valor is None or valor == "":
            return 0.0
        if isinstance(valor, str):
            valor = valor.replace("$", "").replace(".", "").replace(",", ".").strip()
        return float(valor)
    except Exception:
        return 0.0


def limpiar_nombre_archivo(texto: str) -> str:
    texto = str(texto or "").strip().replace(" ", "_")
    texto = re.sub(r"[^A-Za-z0-9_\-]", "", texto)
    return texto[:70] or "cotizacion"


def normalizar_texto(texto: str) -> str:
    return str(texto or "").strip().lower()


def safe_int(valor) -> int:
    try:
        return int(float(valor))
    except Exception:
        return 0


def modo_gsheets_activo() -> bool:
    if gspread is None or Credentials is None:
        return False
    try:
        return "gsheets" in st.secrets and "spreadsheet_id" in st.secrets["gsheets"]
    except Exception:
        return False

# =========================================================
# GOOGLE SHEETS
# =========================================================
GSHEET_SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

@st.cache_resource(show_spinner=False)
def get_gsheet():
    creds_info = dict(st.secrets["gsheets"]["service_account"])
    if "private_key" in creds_info:
        creds_info["private_key"] = creds_info["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_info, scopes=GSHEET_SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(st.secrets["gsheets"]["spreadsheet_id"])


def asegurar_ws(nombre: str, headers: list[str]):
    sh = get_gsheet()
    try:
        ws = sh.worksheet(nombre)
    except Exception:
        ws = sh.add_worksheet(title=nombre, rows=2000, cols=max(20, len(headers) + 2))
        ws.append_row(headers)
        return ws
    values = ws.get_all_values()
    if not values:
        ws.append_row(headers)
    return ws


def gs_get_records(nombre: str, headers: list[str]) -> pd.DataFrame:
    ws = asegurar_ws(nombre, headers)
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    for h in headers:
        if h not in df.columns:
            df[h] = ""
    return df[headers]


def gs_append_row(nombre: str, headers: list[str], row_dict: dict):
    ws = asegurar_ws(nombre, headers)
    ws.append_row([row_dict.get(h, "") for h in headers])


def gs_rewrite(nombre: str, headers: list[str], df: pd.DataFrame):
    ws = asegurar_ws(nombre, headers)
    clean = df.copy()
    for h in headers:
        if h not in clean.columns:
            clean[h] = ""
    clean = clean[headers].fillna("")
    ws.clear()
    ws.update("A1", [headers] + clean.astype(str).values.tolist())


def gs_update_cotizacion_estado(cotizacion_id: int, estado: str):
    ws = asegurar_ws("cotizaciones", HEADERS_COTIZACIONES)
    values = ws.get_all_values()
    if len(values) < 2:
        return False
    header = values[0]
    idx_id = header.index("id")
    idx_estado = header.index("estado")
    idx_update = header.index("actualizado_en")
    for row_num, row in enumerate(values[1:], start=2):
        if len(row) > idx_id and str(row[idx_id]) == str(cotizacion_id):
            ws.update_cell(row_num, idx_estado + 1, estado)
            ws.update_cell(row_num, idx_update + 1, ahora_santiago().strftime("%Y-%m-%d %H:%M:%S"))
            return True
    return False

# =========================================================
# BASE LOCAL SQLite + Excel
# =========================================================
def get_conn():
    return sqlite3.connect(DB_FILE, check_same_thread=False)


def init_local_db():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(f"""
        CREATE TABLE IF NOT EXISTS cotizaciones (
            id INTEGER PRIMARY KEY,
            numero_cotizacion TEXT UNIQUE,
            correlativo INTEGER,
            fecha TEXT,
            estado TEXT,
            cliente TEXT,
            contacto_cel TEXT,
            mail_cliente TEXT,
            patente TEXT,
            marca TEXT,
            modelo TEXT,
            anio TEXT,
            kilometraje TEXT,
            vin TEXT,
            subtotal_mano_obra REAL,
            subtotal_repuestos REAL,
            subtotal_otros REAL,
            total REAL,
            observaciones TEXT,
            condiciones TEXT,
            creado_en TEXT,
            actualizado_en TEXT
        )
    """)
    cur.execute(f"""
        CREATE TABLE IF NOT EXISTS detalle_items (
            id_item INTEGER PRIMARY KEY,
            cotizacion_id INTEGER,
            numero_cotizacion TEXT,
            categoria TEXT,
            descripcion TEXT,
            cantidad REAL,
            valor_unitario REAL,
            total REAL
        )
    """)
    cur.execute(f"""
        CREATE TABLE IF NOT EXISTS catalogos (
            id_cat INTEGER PRIMARY KEY,
            categoria TEXT,
            descripcion TEXT,
            valor_sugerido REAL,
            activo TEXT,
            actualizado_en TEXT
        )
    """)
    conn.commit()
    # Insertar catálogo base si está vacío
    cur.execute("SELECT COUNT(*) FROM catalogos")
    if cur.fetchone()[0] == 0:
        now = ahora_santiago().strftime("%Y-%m-%d %H:%M:%S")
        for i, (cat, desc, val) in enumerate(CATALOGO_DEFAULT, start=1):
            cur.execute(
                "INSERT INTO catalogos (id_cat, categoria, descripcion, valor_sugerido, activo, actualizado_en) VALUES (?, ?, ?, ?, ?, ?)",
                (i, cat, desc, val, "SI", now),
            )
        conn.commit()
    conn.close()


def local_df(tabla: str, headers: list[str]) -> pd.DataFrame:
    init_local_db()
    conn = get_conn()
    try:
        df = pd.read_sql_query(f"SELECT * FROM {tabla}", conn)
    finally:
        conn.close()
    for h in headers:
        if h not in df.columns:
            df[h] = ""
    return df[headers]


def local_rewrite_excel():
    init_local_db()
    df_cot = local_df("cotizaciones", HEADERS_COTIZACIONES)
    df_det = local_df("detalle_items", HEADERS_DETALLE)
    df_cat = local_df("catalogos", HEADERS_CATALOGO)
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_cot.to_excel(writer, sheet_name="cotizaciones", index=False)
        df_det.to_excel(writer, sheet_name="detalle_items", index=False)
        df_cat.to_excel(writer, sheet_name="catalogos", index=False)


def local_insert_cotizacion(cot: dict, items: list[dict]):
    init_local_db()
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        f"INSERT INTO cotizaciones ({', '.join(HEADERS_COTIZACIONES)}) VALUES ({', '.join(['?']*len(HEADERS_COTIZACIONES))})",
        [cot.get(h, "") for h in HEADERS_COTIZACIONES],
    )
    for item in items:
        cur.execute(
            f"INSERT INTO detalle_items ({', '.join(HEADERS_DETALLE)}) VALUES ({', '.join(['?']*len(HEADERS_DETALLE))})",
            [item.get(h, "") for h in HEADERS_DETALLE],
        )
    conn.commit()
    conn.close()
    local_rewrite_excel()


def local_update_estado(cotizacion_id: int, estado: str):
    init_local_db()
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "UPDATE cotizaciones SET estado=?, actualizado_en=? WHERE id=?",
        (estado, ahora_santiago().strftime("%Y-%m-%d %H:%M:%S"), int(cotizacion_id)),
    )
    conn.commit()
    conn.close()
    local_rewrite_excel()


def local_delete_cotizacion(cotizacion_id: int):
    init_local_db()
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM detalle_items WHERE cotizacion_id=?", (int(cotizacion_id),))
    cur.execute("DELETE FROM cotizaciones WHERE id=?", (int(cotizacion_id),))
    conn.commit()
    conn.close()
    local_rewrite_excel()


def local_add_catalogo(categoria: str, descripcion: str, valor: float):
    init_local_db()
    df = local_df("catalogos", HEADERS_CATALOGO)
    next_id = 1 if df.empty else int(pd.to_numeric(df["id_cat"], errors="coerce").fillna(0).max()) + 1
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO catalogos (id_cat, categoria, descripcion, valor_sugerido, activo, actualizado_en) VALUES (?, ?, ?, ?, ?, ?)",
        (next_id, categoria, descripcion, valor, "SI", ahora_santiago().strftime("%Y-%m-%d %H:%M:%S")),
    )
    conn.commit()
    conn.close()
    local_rewrite_excel()


def local_update_catalogo(catalogo_id: int, valor: float | None = None, activo: str | None = None):
    init_local_db()
    conn = get_conn()
    cur = conn.cursor()
    if valor is not None:
        cur.execute("UPDATE catalogos SET valor_sugerido=?, actualizado_en=? WHERE id_cat=?", (valor, ahora_santiago().strftime("%Y-%m-%d %H:%M:%S"), int(catalogo_id)))
    if activo is not None:
        cur.execute("UPDATE catalogos SET activo=?, actualizado_en=? WHERE id_cat=?", (activo, ahora_santiago().strftime("%Y-%m-%d %H:%M:%S"), int(catalogo_id)))
    conn.commit()
    conn.close()
    local_rewrite_excel()

# =========================================================
# CAPA DE DATOS GENERAL
# =========================================================
def inicializar_base():
    if modo_gsheets_activo():
        asegurar_ws("cotizaciones", HEADERS_COTIZACIONES)
        asegurar_ws("detalle_items", HEADERS_DETALLE)
        asegurar_ws("catalogos", HEADERS_CATALOGO)
        df_cat = gs_get_records("catalogos", HEADERS_CATALOGO)
        if df_cat.empty:
            now = ahora_santiago().strftime("%Y-%m-%d %H:%M:%S")
            for i, (cat, desc, val) in enumerate(CATALOGO_DEFAULT, start=1):
                gs_append_row("catalogos", HEADERS_CATALOGO, {
                    "id_cat": i, "categoria": cat, "descripcion": desc,
                    "valor_sugerido": val, "activo": "SI", "actualizado_en": now,
                })
    else:
        init_local_db()
        local_rewrite_excel()


def cargar_cotizaciones() -> pd.DataFrame:
    if modo_gsheets_activo():
        return gs_get_records("cotizaciones", HEADERS_COTIZACIONES)
    return local_df("cotizaciones", HEADERS_COTIZACIONES)


def cargar_detalle() -> pd.DataFrame:
    if modo_gsheets_activo():
        return gs_get_records("detalle_items", HEADERS_DETALLE)
    return local_df("detalle_items", HEADERS_DETALLE)


def cargar_catalogo() -> pd.DataFrame:
    if modo_gsheets_activo():
        return gs_get_records("catalogos", HEADERS_CATALOGO)
    return local_df("catalogos", HEADERS_CATALOGO)


def siguiente_id(df: pd.DataFrame, col: str) -> int:
    if df.empty or col not in df.columns:
        return 1
    return int(pd.to_numeric(df[col], errors="coerce").fillna(0).max()) + 1


def siguiente_correlativo() -> int:
    """Devuelve el siguiente correlativo sin reutilizar números anteriores.

    Para mantener trazabilidad comercial, el correlativo avanza siempre
    desde el máximo histórico guardado en la base. Aunque una cotización
    se anule o se elimine, no se recomienda reutilizar el número.
    """
    df = cargar_cotizaciones()
    if df.empty:
        return 1
    corr = pd.to_numeric(df["correlativo"], errors="coerce").fillna(0)
    max_corr = int(corr.max()) if len(corr) else 0
    return max_corr + 1


def guardar_cotizacion(cot: dict, detalle_items: list[dict]):
    if modo_gsheets_activo():
        gs_append_row("cotizaciones", HEADERS_COTIZACIONES, cot)
        for item in detalle_items:
            gs_append_row("detalle_items", HEADERS_DETALLE, item)
    else:
        local_insert_cotizacion(cot, detalle_items)


def actualizar_estado(cotizacion_id: int, estado: str):
    if modo_gsheets_activo():
        return gs_update_cotizacion_estado(cotizacion_id, estado)
    local_update_estado(cotizacion_id, estado)
    return True


def eliminar_cotizacion(cotizacion_id: int):
    if modo_gsheets_activo():
        df_cot = cargar_cotizaciones()
        df_det = cargar_detalle()
        df_cot = df_cot[df_cot["id"].astype(str) != str(cotizacion_id)]
        df_det = df_det[df_det["cotizacion_id"].astype(str) != str(cotizacion_id)]
        gs_rewrite("cotizaciones", HEADERS_COTIZACIONES, df_cot)
        gs_rewrite("detalle_items", HEADERS_DETALLE, df_det)
    else:
        local_delete_cotizacion(cotizacion_id)


def agregar_catalogo(categoria: str, descripcion: str, valor: float):
    if modo_gsheets_activo():
        df = cargar_catalogo()
        gs_append_row("catalogos", HEADERS_CATALOGO, {
            "id_cat": siguiente_id(df, "id_cat"),
            "categoria": categoria,
            "descripcion": descripcion,
            "valor_sugerido": valor,
            "activo": "SI",
            "actualizado_en": ahora_santiago().strftime("%Y-%m-%d %H:%M:%S"),
        })
    else:
        local_add_catalogo(categoria, descripcion, valor)


def actualizar_catalogo(catalogo_id: int, valor: float | None = None, activo: str | None = None):
    if modo_gsheets_activo():
        df = cargar_catalogo()
        if df.empty:
            return
        for idx, row in df.iterrows():
            if str(row["id_cat"]) == str(catalogo_id):
                if valor is not None:
                    df.at[idx, "valor_sugerido"] = valor
                if activo is not None:
                    df.at[idx, "activo"] = activo
                df.at[idx, "actualizado_en"] = ahora_santiago().strftime("%Y-%m-%d %H:%M:%S")
                break
        gs_rewrite("catalogos", HEADERS_CATALOGO, df)
    else:
        local_update_catalogo(catalogo_id, valor=valor, activo=activo)

# =========================================================
# GENERACIÓN DOCUMENTO
# =========================================================
def preparar_items_contexto(items: list[dict], categoria: str) -> list[dict]:
    filtrados = [i for i in items if i["categoria"] == categoria]
    salida = []
    for i in filtrados:
        salida.append({
            "descripcion": i.get("descripcion", ""),
            "cantidad": f"{to_float(i.get('cantidad')):,.0f}".replace(",", "."),
            "valor_unitario": clp_fmt(i.get("valor_unitario", 0)),
            "total": clp_fmt(i.get("total", 0)),
        })
    return salida


def construir_contexto(cot: dict, items: list[dict]) -> dict:
    return {
        "numero_cotizacion": cot.get("numero_cotizacion", ""),
        "fecha": cot.get("fecha", ""),
        "estado": cot.get("estado", "Pendiente"),
        "cliente": cot.get("cliente", ""),
        "contacto_cel": cot.get("contacto_cel", ""),
        "mail_cliente": cot.get("mail_cliente", ""),
        "patente": cot.get("patente", ""),
        "marca": cot.get("marca", ""),
        "modelo": cot.get("modelo", ""),
        "anio": cot.get("anio", ""),
        "kilometraje": cot.get("kilometraje", ""),
        "vin": cot.get("vin", ""),
        "mano_obra": preparar_items_contexto(items, "Mano de obra"),
        "repuestos": preparar_items_contexto(items, "Repuestos"),
        "otros": preparar_items_contexto(items, "Otros"),
        "total_mano_obra": clp_fmt(cot.get("subtotal_mano_obra", 0)),
        "total_repuestos": clp_fmt(cot.get("subtotal_repuestos", 0)),
        "total_otros": clp_fmt(cot.get("subtotal_otros", 0)),
        "total": clp_fmt(cot.get("total", 0)),
        "observaciones": cot.get("observaciones", "") or "Sin observaciones.",
        "condiciones": cot.get("condiciones", "") or CONDICIONES_DEFAULT,
    }


def xml_text(valor) -> str:
    """Texto seguro para XML Word."""
    return escape(str(valor or ""))


def w_p(texto="", bold=False, size=20, color="111111", align="left", spacing_after=120):
    jc = {"left": "left", "center": "center", "right": "right"}.get(align, "left")
    b = "<w:b/>" if bold else ""
    return f"""
    <w:p>
      <w:pPr><w:jc w:val=\"{jc}\"/><w:spacing w:after=\"{spacing_after}\"/></w:pPr>
      <w:r><w:rPr>{b}<w:sz w:val=\"{size}\"/><w:color w:val=\"{color}\"/></w:rPr><w:t>{xml_text(texto)}</w:t></w:r>
    </w:p>"""


def w_cell(texto="", bold=False, bg=None, align="left", width="2400", color="111111"):
    fill = f'<w:shd w:fill=\"{bg}\"/>' if bg else ""
    b = "<w:b/>" if bold else ""
    jc = {"left": "left", "center": "center", "right": "right"}.get(align, "left")
    return f"""
    <w:tc>
      <w:tcPr><w:tcW w:w=\"{width}\" w:type=\"dxa\"/>{fill}<w:tcMar><w:top w:w=\"80\" w:type=\"dxa\"/><w:left w:w=\"90\" w:type=\"dxa\"/><w:bottom w:w=\"80\" w:type=\"dxa\"/><w:right w:w=\"90\" w:type=\"dxa\"/></w:tcMar></w:tcPr>
      <w:p><w:pPr><w:jc w:val=\"{jc}\"/></w:pPr><w:r><w:rPr>{b}<w:sz w:val=\"18\"/><w:color w:val=\"{color}\"/></w:rPr><w:t>{xml_text(texto)}</w:t></w:r></w:p>
    </w:tc>"""


def w_row(cells):
    return "<w:tr>" + "".join(cells) + "</w:tr>"


def w_table(rows, widths=None):
    grid = ""
    if widths:
        grid = "<w:tblGrid>" + "".join(f'<w:gridCol w:w=\"{w}\"/>' for w in widths) + "</w:tblGrid>"
    return f"""
    <w:tbl>
      <w:tblPr>
        <w:tblStyle w:val=\"TableGrid\"/>
        <w:tblW w:w=\"0\" w:type=\"auto\"/>
        <w:tblBorders>
          <w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"D0D0D0\"/>
          <w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"D0D0D0\"/>
          <w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"D0D0D0\"/>
          <w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"D0D0D0\"/>
          <w:insideH w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"D0D0D0\"/>
          <w:insideV w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"D0D0D0\"/>
        </w:tblBorders>
      </w:tblPr>
      {grid}
      {''.join(rows)}
    </w:tbl>"""


def tabla_datos(contexto: dict) -> str:
    rows = []
    data = [
        ("Fecha", contexto.get("fecha", ""), "Estado", contexto.get("estado", "")),
        ("Cliente", contexto.get("cliente", ""), "Patente", contexto.get("patente", "")),
        ("Contacto", contexto.get("contacto_cel", ""), "Marca", contexto.get("marca", "")),
        ("Mail", contexto.get("mail_cliente", ""), "Modelo", contexto.get("modelo", "")),
        ("Año", contexto.get("anio", ""), "Kilometraje", contexto.get("kilometraje", "")),
        ("VIN", contexto.get("vin", ""), "", ""),
    ]
    for a, b, c, d in data:
        rows.append(w_row([
            w_cell(a, bold=True, bg="F3F4F6", width="1500"), w_cell(b, width="3300"),
            w_cell(c, bold=True, bg="F3F4F6", width="1500"), w_cell(d, width="3300"),
        ]))
    return w_table(rows, [1500, 3300, 1500, 3300])


def tabla_items(titulo: str, items: list[dict]) -> str:
    rows = [w_row([
        w_cell("Descripción", bold=True, bg="111111", color="FFFFFF", width="5200"),
        w_cell("Cant.", bold=True, bg="111111", color="FFFFFF", align="center", width="1000"),
        w_cell("V. unitario", bold=True, bg="111111", color="FFFFFF", align="right", width="1700"),
        w_cell("Total", bold=True, bg="111111", color="FFFFFF", align="right", width="1700"),
    ])]
    if not items:
        rows.append(w_row([w_cell("Sin ítems", width="5200"), w_cell("", width="1000"), w_cell("", width="1700"), w_cell("", width="1700")]))
    else:
        for it in items:
            rows.append(w_row([
                w_cell(it.get("descripcion", ""), width="5200"),
                w_cell(it.get("cantidad", ""), align="center", width="1000"),
                w_cell(it.get("valor_unitario", ""), align="right", width="1700"),
                w_cell(it.get("total", ""), align="right", width="1700"),
            ]))
    return w_p(titulo, bold=True, size=22, color="D96816", spacing_after=80) + w_table(rows, [5200, 1000, 1700, 1700])


def generar_docx_bytes(contexto: dict) -> bytes:
    """Genera un Word profesional sin docxtpl ni python-docx.
    Así evita el error de dependencia en Streamlit Cloud/Codespaces.
    """
    resumen_rows = [
        w_row([w_cell("Total mano de obra", bold=True, width="3500"), w_cell(contexto.get("total_mano_obra", "$ 0"), align="right", width="2500")]),
        w_row([w_cell("Total repuestos", bold=True, width="3500"), w_cell(contexto.get("total_repuestos", "$ 0"), align="right", width="2500")]),
        w_row([w_cell("Total otros", bold=True, width="3500"), w_cell(contexto.get("total_otros", "$ 0"), align="right", width="2500")]),
        w_row([w_cell("TOTAL COTIZACIÓN", bold=True, bg="D96816", color="FFFFFF", width="3500"), w_cell(contexto.get("total", "$ 0"), bold=True, bg="D96816", color="FFFFFF", align="right", width="2500")]),
    ]

    body = f"""
    {w_p("CHARLES SERVICIO AUTOMOTRIZ", bold=True, size=32, color="111111", align="center", spacing_after=40)}
    {w_p("San Bernardo, Chile | Matías Vejar Reyes | charlesautomotriz@gmail.com | +56 9 3453 3841", size=18, color="444444", align="center", spacing_after=180)}
    {w_p("COTIZACIÓN N° " + str(contexto.get("numero_cotizacion", "")), bold=True, size=28, color="D96816", align="center", spacing_after=180)}
    {tabla_datos(contexto)}
    {w_p("", spacing_after=80)}
    {tabla_items("MANO DE OBRA", contexto.get("mano_obra", []))}
    {w_p("", spacing_after=60)}
    {tabla_items("REPUESTOS", contexto.get("repuestos", []))}
    {w_p("", spacing_after=60)}
    {tabla_items("OTROS", contexto.get("otros", []))}
    {w_p("", spacing_after=80)}
    {w_p("OBSERVACIONES", bold=True, size=22, color="D96816", spacing_after=60)}
    {w_table([w_row([w_cell(contexto.get("observaciones", "Sin observaciones."), width="9600")])], [9600])}
    {w_p("", spacing_after=80)}
    {w_p("RESUMEN", bold=True, size=22, color="D96816", spacing_after=60)}
    {w_table(resumen_rows, [3500, 2500])}
    {w_p("CONDICIONES", bold=True, size=22, color="D96816", spacing_after=60)}
    {w_p(contexto.get("condiciones", ""), size=18, color="111111", spacing_after=80)}
    {w_p("¡Gracias por confiar en Charles Servicio Automotriz!", bold=True, size=18, color="444444", align="center", spacing_after=40)}
    {w_p("Servicio profesional, atención cercana y compromiso con tu vehículo.", size=18, color="444444", align="center", spacing_after=40)}
    """

    document_xml = f"""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
    <w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 wp14\">
      <w:body>
        {body}
        <w:sectPr>
          <w:pgSz w:w=\"12240\" w:h=\"15840\"/>
          <w:pgMar w:top=\"720\" w:right=\"720\" w:bottom=\"720\" w:left=\"720\" w:header=\"360\" w:footer=\"360\" w:gutter=\"0\"/>
        </w:sectPr>
      </w:body>
    </w:document>"""

    styles_xml = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
    <w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">
      <w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/><w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/><w:sz w:val=\"20\"/></w:rPr></w:style>
      <w:style w:type=\"table\" w:styleId=\"TableGrid\"><w:name w:val=\"Table Grid\"/><w:tblPr><w:tblBorders><w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"D0D0D0\"/><w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"D0D0D0\"/><w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"D0D0D0\"/><w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"D0D0D0\"/><w:insideH w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"D0D0D0\"/><w:insideV w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"D0D0D0\"/></w:tblBorders></w:tblPr></w:style>
    </w:styles>"""
    content_types = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
    <Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/><Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/></Types>"""
    rels = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>"""
    doc_rels = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rIdStyles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>"""

    bio = io.BytesIO()
    with zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/document.xml", document_xml)
        z.writestr("word/styles.xml", styles_xml)
    return bio.getvalue()

def convertir_pdf_bytes(docx_bytes: bytes) -> bytes | None:
    posibles = [shutil.which("soffice"), "/usr/bin/soffice", "/usr/local/bin/soffice"]
    soffice = next((p for p in posibles if p and os.path.exists(p)), None)
    if not soffice:
        return None
    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        docx_path = tmpdir / "cotizacion.docx"
        docx_path.write_bytes(docx_bytes)
        cmd = [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(tmpdir), str(docx_path)]
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=90)
        pdf_path = tmpdir / "cotizacion.pdf"
        if res.returncode != 0 or not pdf_path.exists():
            return None
        return pdf_path.read_bytes()

# =========================================================
# COMPONENTES UI
# =========================================================
def init_session():
    if "items_cotizacion" not in st.session_state:
        st.session_state.items_cotizacion = []
    if "ultima_cotizacion" not in st.session_state:
        st.session_state.ultima_cotizacion = None


def add_item_ui(categoria: str, catalogo: pd.DataFrame):
    st.markdown(f"#### {categoria}")
    cat = catalogo[(catalogo["categoria"] == categoria) & (catalogo["activo"].astype(str).str.upper().eq("SI"))].copy()
    cat["valor_sugerido"] = pd.to_numeric(cat["valor_sugerido"], errors="coerce").fillna(0)
    opciones = ["Seleccionar..."] + cat["descripcion"].astype(str).tolist() + ["Otro manual"]
    c1, c2, c3, c4 = st.columns([3.2, 1, 1.4, 1.2])
    with c1:
        seleccion = st.selectbox("Descripción", opciones, key=f"sel_{categoria}")
        descripcion = ""
        if seleccion == "Otro manual":
            descripcion = st.text_input("Escribir descripción", key=f"manual_{categoria}")
        elif seleccion != "Seleccionar...":
            descripcion = seleccion
    sugerido = 0.0
    if seleccion not in ["Seleccionar...", "Otro manual"]:
        match = cat[cat["descripcion"].astype(str) == seleccion]
        if not match.empty:
            sugerido = float(match.iloc[0]["valor_sugerido"])
    with c2:
        cantidad = st.number_input("Cantidad", min_value=1.0, value=1.0, step=1.0, key=f"cant_{categoria}")
    with c3:
        valor_unitario = st.number_input("Valor unitario $", min_value=0.0, value=float(sugerido), step=1000.0, key=f"valor_{categoria}")
    with c4:
        st.write("")
        st.write("")
        if st.button("Agregar", key=f"btn_{categoria}", use_container_width=True):
            if not descripcion.strip():
                st.error("Debes seleccionar o escribir una descripción.")
            else:
                st.session_state.items_cotizacion.append({
                    "categoria": categoria,
                    "descripcion": descripcion.strip(),
                    "cantidad": float(cantidad),
                    "valor_unitario": float(valor_unitario),
                    "total": float(cantidad) * float(valor_unitario),
                })
                st.success("Ítem agregado.")
                st.rerun()


def tabla_items_actuales():
    items = st.session_state.items_cotizacion
    if not items:
        st.info("Aún no hay ítems agregados a la cotización.")
        return 0, 0, 0, 0
    df = pd.DataFrame(items)
    df_vista = df.copy()
    df_vista["valor_unitario"] = df_vista["valor_unitario"].apply(clp_fmt)
    df_vista["total"] = df_vista["total"].apply(clp_fmt)
    st.dataframe(df_vista, use_container_width=True, hide_index=True)
    c1, c2 = st.columns([2, 1])
    with c1:
        idx = st.number_input("N° de fila a eliminar", min_value=1, max_value=len(items), value=1, step=1)
    with c2:
        st.write("")
        st.write("")
        if st.button("Eliminar ítem seleccionado", type="secondary", use_container_width=True):
            st.session_state.items_cotizacion.pop(int(idx) - 1)
            st.rerun()
    subtotal_mo = sum(to_float(i["total"]) for i in items if i["categoria"] == "Mano de obra")
    subtotal_rep = sum(to_float(i["total"]) for i in items if i["categoria"] == "Repuestos")
    subtotal_otros = sum(to_float(i["total"]) for i in items if i["categoria"] == "Otros")
    total = subtotal_mo + subtotal_rep + subtotal_otros
    return subtotal_mo, subtotal_rep, subtotal_otros, total

# =========================================================
# APP PRINCIPAL
# =========================================================
aplicar_estilo()
encabezado()
inicializar_base()
init_session()

modo = "Google Sheets / Drive" if modo_gsheets_activo() else "SQLite local + Excel local"
st.sidebar.success(f"Base activa: {modo}")
st.sidebar.caption(f"Hora Santiago: {ahora_santiago().strftime('%d-%m-%Y %H:%M:%S')}")
st.sidebar.caption("Tip: para persistencia real en Streamlit Cloud, configura Google Sheets en secrets.")

tab_nueva, tab_hist, tab_cat, tab_dash, tab_base = st.tabs([
    "🧾 Nueva cotización", "📚 Historial", "🧩 Catálogos", "📊 Dashboard", "💾 Base / Drive"
])

# =========================================================
# NUEVA COTIZACIÓN
# =========================================================
with tab_nueva:
    st.subheader("Nueva cotización")
    st.caption("Flujo simple: datos del cliente y vehículo → seleccionar ítems → revisar total → generar Word/PDF.")

    catalogo = cargar_catalogo()

    with st.form("form_datos_cotizacion"):
        st.markdown("### 1) Datos del cliente y vehículo")
        c1, c2 = st.columns(2)
        with c1:
            fecha = st.date_input("Fecha", value=fecha_hoy())
            cliente = st.text_input("Cliente")
            contacto_cel = st.text_input("Contacto Cel")
            mail_cliente = st.text_input("Mail del cliente")
        with c2:
            patente = st.text_input("Patente")
            marca = st.text_input("Marca")
            modelo = st.text_input("Modelo")
            anio = st.text_input("Año")
            kilometraje = st.text_input("Kilometraje")
            vin = st.text_input("VIN")
        observaciones = st.text_area("Observaciones", height=80)
        condiciones = st.text_area("Condiciones", value=CONDICIONES_DEFAULT, height=110)
        submitted_data = st.form_submit_button("Guardar datos temporales", use_container_width=True)
        if submitted_data:
            st.success("Datos actualizados. Ahora agrega mano de obra, repuestos u otros.")

    st.markdown("### 2) Agregar ítems desde catálogo")
    i1, i2, i3 = st.tabs(["Mano de obra", "Repuestos", "Otros"])
    with i1:
        add_item_ui("Mano de obra", catalogo)
    with i2:
        add_item_ui("Repuestos", catalogo)
    with i3:
        add_item_ui("Otros", catalogo)

    st.markdown("### 3) Resumen de ítems")
    subtotal_mo, subtotal_rep, subtotal_otros, total = tabla_items_actuales()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Mano de obra", clp_fmt(subtotal_mo))
    m2.metric("Repuestos", clp_fmt(subtotal_rep))
    m3.metric("Otros", clp_fmt(subtotal_otros))
    m4.metric("Total", clp_fmt(total))

    st.markdown("### 4) Generar cotización")
    if st.button("Generar y guardar cotización", type="primary", use_container_width=True):
        if not cliente.strip():
            st.error("Debes ingresar el cliente.")
        elif not st.session_state.items_cotizacion:
            st.error("Debes agregar al menos un ítem de mano de obra, repuestos u otros.")
        else:
            df_cot = cargar_cotizaciones()
            df_det = cargar_detalle()
            cotizacion_id = siguiente_id(df_cot, "id")
            correlativo = siguiente_correlativo()
            numero = f"{PREFIJO_COTIZACION}-{correlativo:04d}"
            now = ahora_santiago().strftime("%Y-%m-%d %H:%M:%S")
            cot = {
                "id": cotizacion_id,
                "numero_cotizacion": numero,
                "correlativo": correlativo,
                "fecha": fecha.strftime("%d-%m-%Y"),
                "estado": "Pendiente",
                "cliente": cliente.strip(),
                "contacto_cel": contacto_cel.strip(),
                "mail_cliente": mail_cliente.strip(),
                "patente": patente.strip(),
                "marca": marca.strip(),
                "modelo": modelo.strip(),
                "anio": anio.strip(),
                "kilometraje": kilometraje.strip(),
                "vin": vin.strip(),
                "subtotal_mano_obra": subtotal_mo,
                "subtotal_repuestos": subtotal_rep,
                "subtotal_otros": subtotal_otros,
                "total": total,
                "observaciones": observaciones.strip(),
                "condiciones": condiciones.strip(),
                "creado_en": now,
                "actualizado_en": now,
            }
            detalle_para_guardar = []
            next_item_id = siguiente_id(df_det, "id_item")
            for idx, item in enumerate(st.session_state.items_cotizacion, start=0):
                detalle_para_guardar.append({
                    "id_item": next_item_id + idx,
                    "cotizacion_id": cotizacion_id,
                    "numero_cotizacion": numero,
                    "categoria": item["categoria"],
                    "descripcion": item["descripcion"],
                    "cantidad": item["cantidad"],
                    "valor_unitario": item["valor_unitario"],
                    "total": item["total"],
                })
            try:
                guardar_cotizacion(cot, detalle_para_guardar)
                contexto = construir_contexto(cot, detalle_para_guardar)
                docx_bytes = generar_docx_bytes(contexto)
                pdf_bytes = convertir_pdf_bytes(docx_bytes)
                nombre = f"Cotizacion_{numero}_{limpiar_nombre_archivo(cliente)}"
                st.session_state.ultima_cotizacion = {
                    "numero": numero,
                    "nombre": nombre,
                    "docx": docx_bytes,
                    "pdf": pdf_bytes,
                }
                st.session_state.items_cotizacion = []
                st.success(f"Cotización {numero} generada y guardada correctamente.")
                st.rerun()
            except Exception as e:
                st.error(f"Error al generar la cotización: {e}")

    if st.session_state.ultima_cotizacion:
        data = st.session_state.ultima_cotizacion
        st.success(f"Última cotización generada: {data['numero']}")
        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                "Descargar Word",
                data=data["docx"],
                file_name=f"{data['nombre']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        with d2:
            if data["pdf"]:
                st.download_button(
                    "Descargar PDF",
                    data=data["pdf"],
                    file_name=f"{data['nombre']}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
            else:
                st.warning("PDF no disponible. Instala LibreOffice/soffice o agrega packages.txt con libreoffice.")

# =========================================================
# HISTORIAL
# =========================================================
with tab_hist:
    st.subheader("Historial de cotizaciones")
    df = cargar_cotizaciones()
    det = cargar_detalle()
    if df.empty:
        st.info("Aún no hay cotizaciones guardadas.")
    else:
        df["total_num"] = pd.to_numeric(df["total"], errors="coerce").fillna(0)
        c1, c2, c3 = st.columns([1.5, 2, 2])
        with c1:
            estado_f = st.selectbox("Filtrar por estado", ["Todos"] + ESTADOS)
        with c2:
            texto_f = st.text_input("Buscar cliente, patente o número")
        with c3:
            st.write("")
            st.write("")
            if st.button("Actualizar historial", use_container_width=True):
                st.rerun()
        vista = df.copy()
        if estado_f != "Todos":
            vista = vista[vista["estado"] == estado_f]
        if texto_f.strip():
            f = normalizar_texto(texto_f)
            mask = (
                vista["cliente"].astype(str).str.lower().str.contains(f, na=False) |
                vista["patente"].astype(str).str.lower().str.contains(f, na=False) |
                vista["numero_cotizacion"].astype(str).str.lower().str.contains(f, na=False)
            )
            vista = vista[mask]
        vista_display = vista.copy().sort_values("id", ascending=False)
        for col in ["subtotal_mano_obra", "subtotal_repuestos", "subtotal_otros", "total"]:
            vista_display[col] = vista_display[col].apply(clp_fmt)
        st.dataframe(
            vista_display[["id", "numero_cotizacion", "fecha", "estado", "cliente", "patente", "marca", "modelo", "total"]],
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("### Gestionar cotización")
        opciones = [f"{r['id']} | {r['numero_cotizacion']} | {r['cliente']} | {r['patente']} | {clp_fmt(r['total_num'])}" for _, r in vista.iterrows()]
        if opciones:
            sel = st.selectbox("Selecciona cotización", opciones)
            cot_id = int(sel.split("|")[0].strip())
            row = df[df["id"].astype(str) == str(cot_id)].iloc[0].to_dict()
            items_sel = det[det["cotizacion_id"].astype(str) == str(cot_id)].to_dict("records")

            g1, g2, g3 = st.columns(3)
            with g1:
                nuevo_estado = st.selectbox("Nuevo estado", ESTADOS, index=ESTADOS.index(row.get("estado", "Pendiente")) if row.get("estado", "Pendiente") in ESTADOS else 0)
                if st.button("Actualizar estado", use_container_width=True):
                    actualizar_estado(cot_id, nuevo_estado)
                    st.success("Estado actualizado.")
                    st.rerun()
            with g2:
                if st.button("Recrear Word/PDF", use_container_width=True):
                    try:
                        contexto = construir_contexto(row, items_sel)
                        docx_bytes = generar_docx_bytes(contexto)
                        pdf_bytes = convertir_pdf_bytes(docx_bytes)
                        nombre = f"Cotizacion_{row['numero_cotizacion']}_{limpiar_nombre_archivo(row['cliente'])}"
                        st.session_state.recrear = {"nombre": nombre, "docx": docx_bytes, "pdf": pdf_bytes}
                    except Exception as e:
                        st.error(f"No fue posible recrear el documento: {e}")
            with g3:
                if st.button("Anular por error", type="secondary", use_container_width=True):
                    actualizar_estado(cot_id, "Anulada por error")
                    st.warning("Cotización anulada por error. El correlativo queda trazable.")
                    st.rerun()

            if "recrear" in st.session_state:
                rec = st.session_state.recrear
                d1, d2 = st.columns(2)
                with d1:
                    st.download_button("Descargar Word recreado", rec["docx"], file_name=f"{rec['nombre']}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
                with d2:
                    if rec["pdf"]:
                        st.download_button("Descargar PDF recreado", rec["pdf"], file_name=f"{rec['nombre']}.pdf", mime="application/pdf", use_container_width=True)
                    else:
                        st.warning("PDF no disponible: falta LibreOffice/soffice.")

            with st.expander("Eliminar definitivamente una cotización"):
                st.warning("Recomendación: usa 'Anulada por error' para no perder trazabilidad. Eliminar definitivamente solo si fue una prueba o carga equivocada.")
                confirmar = st.checkbox("Confirmo que quiero eliminar definitivamente esta cotización", key=f"conf_del_{cot_id}")
                if confirmar and st.button("Eliminar definitivamente", type="secondary"):
                    eliminar_cotizacion(cot_id)
                    st.success("Cotización eliminada definitivamente.")
                    st.rerun()

# =========================================================
# CATÁLOGOS
# =========================================================
with tab_cat:
    st.subheader("Catálogos de mano de obra, repuestos y otros")
    st.caption("Estos listados alimentan los selectores de la cotización. Puedes agregar ítems y actualizar precios sugeridos.")
    cat = cargar_catalogo()
    if not cat.empty:
        cat_view = cat.copy()
        cat_view["valor_sugerido"] = pd.to_numeric(cat_view["valor_sugerido"], errors="coerce").fillna(0)
        cat_view["valor_sugerido_fmt"] = cat_view["valor_sugerido"].apply(clp_fmt)
        st.dataframe(cat_view[["id_cat", "categoria", "descripcion", "valor_sugerido_fmt", "activo"]], use_container_width=True, hide_index=True)

    st.markdown("### Agregar nuevo ítem al catálogo")
    with st.form("nuevo_catalogo"):
        c1, c2, c3 = st.columns([1.4, 3, 1.2])
        with c1:
            cat_nueva = st.selectbox("Categoría", CATEGORIAS)
        with c2:
            desc_nueva = st.text_input("Descripción")
        with c3:
            valor_nuevo = st.number_input("Valor sugerido $", min_value=0.0, value=0.0, step=1000.0)
        if st.form_submit_button("Agregar al catálogo", use_container_width=True):
            if not desc_nueva.strip():
                st.error("Debes escribir una descripción.")
            else:
                agregar_catalogo(cat_nueva, desc_nueva.strip(), valor_nuevo)
                st.success("Ítem agregado al catálogo.")
                st.rerun()

    st.markdown("### Editar precio o activar/desactivar")
    if not cat.empty:
        opciones_cat = [f"{r['id_cat']} | {r['categoria']} | {r['descripcion']}" for _, r in cat.iterrows()]
        sel_cat = st.selectbox("Selecciona ítem", opciones_cat)
        cat_id = int(sel_cat.split("|")[0].strip())
        fila_cat = cat[cat["id_cat"].astype(str) == str(cat_id)].iloc[0]
        c1, c2, c3 = st.columns(3)
        with c1:
            nuevo_valor = st.number_input("Nuevo valor sugerido $", min_value=0.0, value=float(to_float(fila_cat.get("valor_sugerido", 0))), step=1000.0)
            if st.button("Actualizar precio", use_container_width=True):
                actualizar_catalogo(cat_id, valor=nuevo_valor)
                st.success("Precio actualizado.")
                st.rerun()
        with c2:
            if st.button("Desactivar", use_container_width=True):
                actualizar_catalogo(cat_id, activo="NO")
                st.warning("Ítem desactivado.")
                st.rerun()
        with c3:
            if st.button("Activar", use_container_width=True):
                actualizar_catalogo(cat_id, activo="SI")
                st.success("Ítem activado.")
                st.rerun()

# =========================================================
# DASHBOARD
# =========================================================
with tab_dash:
    st.subheader("Dashboard comercial")
    df = cargar_cotizaciones()
    if df.empty:
        st.info("No hay datos para visualizar.")
    else:
        df["total"] = pd.to_numeric(df["total"], errors="coerce").fillna(0)
        total_cotizado = df["total"].sum()
        total_aceptado = df[df["estado"] == "Aceptada"]["total"].sum()
        cant_total = len(df)
        cant_aceptada = len(df[df["estado"] == "Aceptada"])
        tasa = (cant_aceptada / cant_total * 100) if cant_total else 0
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Cotizaciones", cant_total)
        c2.metric("Monto cotizado", clp_fmt(total_cotizado))
        c3.metric("Monto aceptado", clp_fmt(total_aceptado))
        c4.metric("Tasa aceptación", f"{tasa:.1f}%")

        resumen_estado = df.groupby("estado", dropna=False).agg(cantidad=("id", "count"), monto=("total", "sum")).reset_index()
        resumen_estado["monto"] = resumen_estado["monto"].apply(clp_fmt)
        st.markdown("### Resumen por estado")
        st.dataframe(resumen_estado, use_container_width=True, hide_index=True)

        resumen_marca = df.groupby(["marca", "modelo"], dropna=False).agg(cantidad=("id", "count"), monto_total=("total", "sum")).reset_index().sort_values("monto_total", ascending=False)
        resumen_marca["monto_total"] = resumen_marca["monto_total"].apply(clp_fmt)
        st.markdown("### Resumen por marca/modelo")
        st.dataframe(resumen_marca, use_container_width=True, hide_index=True)

# =========================================================
# BASE / DRIVE
# =========================================================
with tab_base:
    st.subheader("Base de datos y configuración")
    if modo_gsheets_activo():
        spreadsheet_id = st.secrets["gsheets"].get("spreadsheet_id", "")
        st.success("La app está usando Google Sheets como base principal. Esto permite mantener historial, correlativos y estados aunque se reinicie Streamlit Cloud.")
        st.code(f"Spreadsheet ID: {spreadsheet_id}")
    else:
        st.warning("La app está usando base local SQLite + Excel. En Codespaces funciona, pero en Streamlit Cloud puede reiniciarse si no configuras Google Sheets.")
        local_rewrite_excel()
        c1, c2 = st.columns(2)
        with c1:
            if EXCEL_FILE.exists():
                st.download_button("Descargar base Excel", EXCEL_FILE.read_bytes(), file_name="base_cotizaciones_charles.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with c2:
            if DB_FILE.exists():
                st.download_button("Descargar base SQLite", DB_FILE.read_bytes(), file_name="cotizaciones_charles.db", mime="application/octet-stream", use_container_width=True)

    st.markdown("### Instrucciones para Google Sheets")
    st.markdown(
        """
        1. Crea una planilla en Google Drive.
        2. Crea una cuenta de servicio en Google Cloud y comparte la planilla con el correo de esa cuenta.
        3. Agrega los datos en `.streamlit/secrets.toml` siguiendo el archivo `secrets.toml.example`.
        4. La app creará automáticamente las hojas: `cotizaciones`, `detalle_items` y `catalogos`.
        """
    )
