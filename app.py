import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from pathlib import Path
import unicodedata
import re

# ---------- Config ----------
st.set_page_config(page_title="BU Breakdown per Site", layout="centered")
TITLE = "BU Breakdown per Site"

# 👉 Ajusta estos 3 valores:
DATA_PATH = Path(r"C:\Users\Miguel\OneDrive - MBD 2005 SL\Documentos\Javier Rua\clone github\BUpercentage\IFF Directory 2025-08-04 01_05 EDT.xlsx")
SHEET_NAME = 0               # nombre de hoja o índice (ej. "Sheet1" o 0)
HEADER_ROW = 0               # fila (0-index) donde están las cabeceras reales

# Columnas requeridas lógicas (claves internas)
REQUIRED_KEYS = ["business_unit", "location", "real_estate_id"]

# Sinónimos aceptados para cada columna (formato flexible; se normaliza)
SYNONYMS = {
    "business_unit": [
        "business unit", "bu", "b.u.", "unidad negocio", "unidad de negocio",
        "division", "segment", "business_unit"
    ],
    "location": [
        "location", "site", "city", "ciudad", "ubicacion", "ubicación", "localizacion", "localización"
    ],
    "real_estate_id": [
        "real estate id", "realestateid", "real estate code", "realestate code",
        "re id", "reid", "property id", "property code", "codigo inmueble", "código inmueble"
    ]
}

def normalize(s: str) -> str:
    """Minúsculas, sin tildes, sin espacios/guiones/puntuación."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.lower()
    s = re.sub(r"[^a-z0-9]", "", s)  # fuera espacios, _, -, ., etc.
    return s

def best_match(required_key: str, candidates: list[str]) -> str | None:
    """
    Devuelve el nombre ORIGINAL de la columna del Excel que mejor coincide
    con el required_key usando sinónimos + normalización.
    """
    syns_norm = [normalize(x) for x in SYNONYMS.get(required_key, [])]
    cand_norm = {col: normalize(col) for col in candidates}

    # 1) Coincidencia exacta con algún sinónimo
    for col, ncol in cand_norm.items():
        if ncol in syns_norm:
            return col

    # 2) Contiene (para encabezados largos)
    for col, ncol in cand_norm.items():
        if any(sn in ncol for sn in syns_norm):
            return col

    # 3) Coincidencia por tokens clave (por ejemplo "real","estate","id")
    tokens_map = {
        "business_unit": ["business", "unit", "bu"],
        "location": ["location", "site", "city"],
        "real_estate_id": ["real", "estate", "id", "code", "property"]
    }
    tokens = [normalize(t) for t in tokens_map.get(required_key, [])]
    for col, ncol in cand_norm.items():
        if sum(t in ncol for t in tokens) >= 2:  # al menos 2 tokens presentes
            return col

    return None

@st.cache_data
def load_excel_flexible(path: Path, sheet_name, header_row: int):
    if not path.exists():
        raise FileNotFoundError(f"No existe el archivo en: {path}")

    # Abrir en binario ayuda con locks de OneDrive/Excel
    with open(path, "rb") as fh:
        df = pd.read_excel(fh, engine="openpyxl", sheet_name=sheet_name, header=header_row)

    # Columnas disponibles
    cols = [str(c) for c in df.columns]
    mapping = {}
    for key in REQUIRED_KEYS:
        match = best_match(key, cols)
        if match:
            mapping[key] = match

    # Si falta alguna, informar claramente
    missing_keys = [k for k in REQUIRED_KEYS if k not in mapping]
    if missing_keys:
        detail = "\n".join([f"- {k}: {SYNONYMS[k]}" for k in missing_keys])
        raise ValueError(
            "No se localizaron algunas columnas requeridas.\n"
            f"Hoja: {sheet_name} | Fila cabecera: {header_row}\n"
            f"Columnas detectadas: {cols}\n\n"
            f"Faltantes y sinónimos aceptados:\n{detail}"
        )

    # Subconjunto + renombrado a nombres estándar para el resto del código
    out = df[[mapping["business_unit"], mapping["location"], mapping["real_estate_id"]]].copy()
    out.columns = ["Business Unit", "Location", "Real Estate ID"]

    # Limpieza básica
    for c in out.columns:
        if out[c].dtype == object:
            out[c] = out[c].astype(str).str.strip()

    return out

def categorize_bu(bu: str) -> str:
    if not isinstance(bu, str):
        return "Corporate"
    s = bu.lower()
    if "taste" in s:
        return "Taste"
    if "scent" in s:
        return "Scent"
    if "health" in s or "biosc" in s or "h&b" in s or "h+b" in s:
        return "Health & Biosciences"
    if ("food" in s and "ingredient" in s) or "food ing" in s or "ingredien" in s:
        return "Food Ingredients"
    return "Corporate"

def plot_pie(counts: dict):
    labels = list(counts.keys())
    sizes = list(counts.values())
    fig, ax = plt.subplots()
    if sum(sizes) == 0:
        ax.text(0.5, 0.5, "No data", ha="center", va="center")
        ax.axis("off")
        return fig
    ax.pie(
        sizes,
        labels=labels,
        autopct=lambda p: f"{p:.0f}%\n({int(round(p/100*sum(sizes)))})" if p > 0 else "",
        startangle=90
    )
    ax.axis('equal')
    return fig

# ---------- UI ----------
st.markdown(f"## {TITLE}")
st.caption(f"Reading from: `{DATA_PATH}` | sheet: `{SHEET_NAME}` | header row: {HEADER_ROW}")

# Cargar datos desde ruta fija con matching flexible
try:
    df = load_excel_flexible(DATA_PATH, SHEET_NAME, HEADER_ROW)
except PermissionError:
    st.error(
        "No se pudo abrir el archivo (Permission denied). ¿Está abierto en Excel/OneDrive? "
