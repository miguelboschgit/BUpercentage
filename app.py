import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from matplotlib import colors as mcolors
from matplotlib import cm
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
    for col, ncol in cand_norm.items():
        if ncol in syns_norm:
            return col
    for col, ncol in cand_norm.items():
        if any(sn in ncol for sn in syns_norm):
            return col
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
    if header_row >= len(raw_df):
        return None
    cols = [str(c) for c in raw_df.iloc[header_row]]
    df = raw_df.iloc[header_row + 1:].copy()
    df.columns = cols
    df = df.reset_index(drop=True)
    mapping = {}
    for key in REQUIRED_KEYS:
        match = best_match(key, [str(c) for c in df.columns])
        if match:
            mapping[key] = match
    if len(mapping) < len(REQUIRED_KEYS):
        return None
    out = df[[mapping["business_unit"], mapping["location"], mapping["real_estate_id"]]].copy()
    out.columns = ["Business Unit", "Location", "Real Estate ID"]
    for c in out.columns:
        if out[c].dtype == object:
            out[c] = out[c].astype(str).str.strip()
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
            raw = pd.read_excel(fh, engine="openpyxl", sheet_name=sh, header=None)
        hr, df_ok = find_header_row(raw)
        return hr, df_ok, raw

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

    with open(path, "rb") as fh:
        xls = pd.ExcelFile(fh, engine="openpyxl")
    last_first_row = None
    for sh in xls.sheet_names:
        hr, df_ok, raw = load_from_sheet(sh)
        if df_ok is not None:
            return sh, hr, df_ok
        last_first_row = [str(c) for c in raw.iloc[0]] if len(raw) else []
    raise ValueError(
        "No se localizaron columnas requeridas en ninguna hoja.\n"
        f"Hojas probadas: {xls.sheet_names}\n"
        f"Ejemplo de fila 0 de una hoja: {last_first_row}\n"
        f"Revisa que existan columnas equivalentes a: {SYNONYMS}"
    )

# ========== CLASIFICACIÓN BU ==========
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

# ========== PIE + COLOR MAP (AZULES CLAROS) ==========
def plot_pie_and_colors_blue(counts: dict):
    """
    Pie chart en escala 'Blues' clara.
    - Oculta categorías con 0 en el pie.
    - Devuelve (fig, color_map) donde 0 -> blanco en la tabla.
    - Mayor 'Users' => azul más oscuro (pero no muy oscuro).
    """
    labels_all = list(counts.keys())
    sizes_all  = list(counts.values())
    pairs = [(l, s) for l, s in zip(labels_all, sizes_all) if s > 0]

    fig, ax = plt.subplots()
    if not pairs:
        ax.text(0.5, 0.5, "No data", ha="center", va="center")
        ax.axis("off")
        return fig, {k: "#ffffff" for k in labels_all}

    labels = [p[0] for p in pairs]
    sizes  = [p[1] for p in pairs]

    vmin, vmax = min(sizes), max(sizes)
    if vmin == vmax:
        # todos iguales → tono medio-claro constante
        def norm(_): return 0.5
    else:
        _norm = mcolors.Normalize(vmin=vmin, vmax=vmax)
        def norm(v): return _norm(v)

    # sub-rango claro de Blues para evitar tonos muy oscuros
    cmap_full = cm.get_cmap("Blues")
    def cmap_light_unit(u: float):
        # u en [0,1] → mapea a [0.25, 0.75] (claro → medio)
        x = 0.25 + 0.50 * max(0.0, min(1.0, u))
        return cmap_full(x)

    slice_colors = [mcolors.to_hex(cmap_light_unit(norm(v))) for v in sizes]

    ax.pie(
        sizes,
        labels=labels,             # solo >0
        colors=slice_colors,
        autopct=lambda p: f"{p:.0f}%\n({int(round(p/100*sum(sizes)))})" if p > 0 else "",
        startangle=90
    )
    ax.axis("equal")

    # mapa de colores para la tabla (0 → blanco)
    color_map = {}
    for k, v in counts.items():
        color_map[k] = "#ffffff" if v == 0 else mcolors.to_hex(cmap_light_unit(norm(v)))
    return fig, color_map

# ========== UI ==========
st.markdown(f"## {TITLE}")
st.caption(f"Reading from: `{DATA_PATH}` | sheet: `{SHEET_NAME if SHEET_NAME is not None else 'auto'}`")

# Cargar datos con autodetección
try:
    found_sheet, header_row, df = load_excel_auto(DATA_PATH, SHEET_NAME)
    st.caption(f"Detected sheet: `{found_sheet}` | header row: {header_row}")
except PermissionError:
    st.error("Permission denied. Cierra el Excel/OneDrive y recarga.")
    st.stop()
except Exception as e:
    st.error(str(e))
    st.stop()

# ----- Filtro Location -----
locations = sorted(df["Location"].dropna().unique().tolist())
selected_loc = st.selectbox("Location", options=locations, index=0 if locations else None)
if selected_loc is None:
    st.stop()

df_loc = df[df["Location"] == selected_loc].copy()

# ----- Real Estate ID -----
unique_ids = sorted(df_loc["Real Estate ID"].dropna().astype(str).unique().tolist())
reid_text = "—" if not unique_ids else (unique_ids[0] if len(unique_ids) == 1 else ", ".join(unique_ids))
st.markdown(f"**Real Estate ID:** {reid_text}")

# ----- Cómputo por BU -----
df_loc["BU Category"] = df_loc["Business Unit"].apply(categorize_bu)
order = ["Taste", "Scent", "Food Ingredients", "Health & Biosciences", "Corporate"]
counts = {k: int((df_loc["BU Category"] == k).sum()) for k in order}

total_users = sum(counts.values())
if total_users == 0:
    st.info("No hay usuarios en esta Location.")
    st.stop()

# ----- Pie chart + colores (azules claros y sin categorías 0) -----
st.divider()
st.write("### Breakdown by Business Unit")
fig, color_map = plot_pie_and_colors_blue(counts)
st.pyplot(fig, clear_figure=True)

# ----- Resumen (tabla con el mismo esquema de color) -----
st.write("#### Resumen")
summary_df = pd.DataFrame({"Business Unit": order, "Users": [counts[k] for k in order]})
summary_df["Share %"] = (summary_df["Users"] / total_users * 100).round(1)

def highlight_row(row):
    bg = color_map.get(row["Business Unit"], "#ffffff")  # 0 → blanco
    return [f"background-color: {bg}; color: black" for _ in row]

styled = summary_df.style.apply(highlight_row, axis=1)
st.table(styled)

st.caption(f"Total users in {selected_loc}: {total_users}")
