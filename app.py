import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from pathlib import Path
import unicodedata
import re

# ========== CONFIG ==========
st.set_page_config(page_title="BU Breakdown per Site", layout="centered")
TITLE = "BU Breakdown per Site"

# 👉 Ajusta esta ruta a tu fichero:
DATA_PATH = Path(r"C:\Users\Miguel\OneDrive - MBD 2005 SL\Documentos\Javier Rua\clone github\BUpercentage\IFF Directory 2025-08-04 01_05 EDT.xlsx")

# Si dejas SHEET_NAME = None, probará todas las hojas hasta encontrar cabeceras válidas
SHEET_NAME = None           # None | índice (0,1,...) | nombre ("Sheet1")
MAX_HEADER_SEARCH_ROWS = 40 # filas superiores a inspeccionar para hallar cabeceras

# Columnas requeridas (claves lógicas internas)
REQUIRED_KEYS = ["business_unit", "location", "real_estate_id"]

# Sinónimos aceptados (se normalizan antes de comparar)
SYNONYMS = {
    "business_unit": [
        "business unit", "bu", "b.u.", "unidad negocio", "unidad de negocio",
        "division", "segment", "business_unit"
    ],
    "location": [
        "location", "site", "city", "ciudad", "ubicacion", "ubicación",
        "localizacion", "localización", "loc"
    ],
    "real_estate_id": [
        "real estate id", "realestateid", "real estate code", "realestate code",
        "re id", "reid", "property id", "property code", "codigo inmueble",
        "código inmueble", "realestate", "re code"
    ]
}

# ========== NORMALIZACIÓN/MATCHING ==========
def normalize(s: str) -> str:
    """Minúsculas, sin tildes, y sin caracteres no alfanuméricos."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.lower()
    s = re.sub(r"[^a-z0-9]", "", s)
    return s

def best_match(required_key: str, candidates: list[str]) -> str | None:
    syns_norm = [normalize(x) for x in SYNONYMS.get(required_key, [])]
    cand_norm = {col: normalize(col) for col in candidates}

    # 1) Coincidencia exacta con algún sinónimo
    for col, ncol in cand_norm.items():
        if ncol in syns_norm:
            return col

    # 2) Contiene (por si el encabezado es más largo)
    for col, ncol in cand_norm.items():
        if any(sn in ncol for sn in syns_norm):
            return col

    # 3) Tokens clave
    tokens_map = {
        "business_unit": ["business", "unit", "bu", "segment", "division"],
        "location": ["location", "site", "city", "loc"],
        "real_estate_id": ["real", "estate", "id", "code", "property", "re"]
    }
    tokens = [normalize(t) for t in tokens_map.get(required_key, [])]
    for col, ncol in cand_norm.items():
        if sum(t in ncol for t in tokens) >= 2:
            return col

    return None

# ========== DETECCIÓN DE CABECERAS ==========
def try_build_df(raw_df: pd.DataFrame, header_row: int) -> pd.DataFrame | None:
    """Usa header_row como fila de cabeceras y construye df con columnas requeridas."""
    if header_row >= len(raw_df):
        return None

    # Tomamos la fila como cabecera
    cols = [str(c) for c in raw_df.iloc[header_row]]
    df = raw_df.iloc[header_row + 1:].copy()
    df.columns = cols
    df = df.reset_index(drop=True)

    # Buscar mapeo para las 3 columnas
    mapping = {}
    for key in REQUIRED_KEYS:
        match = best_match(key, [str(c) for c in df.columns])
        if match:
            mapping[key] = match

    if len(mapping) < len(REQUIRED_KEYS):
        return None

    out = df[[mapping["business_unit"], mapping["location"], mapping["real_estate_id"]]].copy()
    out.columns = ["Business Unit", "Location", "Real Estate ID"]

    # Limpieza básica
    for c in out.columns:
        if out[c].dtype == object:
            out[c] = out[c].astype(str).str.strip()

    # Filtrar filas totalmente vacías en las requeridas
    out = out[(out["Business Unit"] != "") | (out["Location"] != "") | (out["Real Estate ID"] != "")]
    return out

def find_header_row(raw_df: pd.DataFrame) -> tuple[int | None, pd.DataFrame | None]:
    n = min(MAX_HEADER_SEARCH_ROWS, len(raw_df))
    for r in range(n):
        trial = try_build_df(raw_df, r)
        if trial is not None:
            return r, trial
    return None, None

@st.cache_data
def load_excel_auto(path: Path, sheet_name):
    if not path.exists():
        raise FileNotFoundError(f"No existe el archivo: {path}")

    def load_from_sheet(sh):
        with open(path, "rb") as fh:
            # Leemos sin cabecera para poder escanear filas
            raw = pd.read_excel(fh, engine="openpyxl", sheet_name=sh, header=None)
        hr, df_ok = find_header_row(raw)
        return hr, df_ok, raw

    # 1) Si se especifica hoja, probar solo esa
    if sheet_name is not None:
        hr, df_ok, raw = load_from_sheet(sheet_name)
        if df_ok is None:
            first_row = [str(c) for c in raw.iloc[0]] if len(raw) else []
            raise ValueError(
                "No se localizaron columnas requeridas.\n"
                f"Hoja: {sheet_name} | Exploradas hasta fila: {MAX_HEADER_SEARCH_ROWS}\n"
                f"Ejemplo de fila 0: {first_row}\n"
                f"Revisa que existan columnas equivalentes a: {SYNONYMS}"
            )
        return sheet_name, hr, df_ok

    # 2) Probar todas las hojas
    with open(path, "rb") as fh:
        xls = pd.ExcelFile(fh, engine="openpyxl")
    last_first_row = None
    for sh in xls.sheet_names:
        hr, df_ok, raw = load_from_sheet(sh)
        if df_ok is not None:
            return sh, hr, df_ok
        last_first_row = [str(c) for c in raw.iloc[0]] if len(raw) else []

    raise ValueError(
