import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from matplotlib import colors as mcolors
from matplotlib import cm
from pathlib import Path
import unicodedata
import re
import numpy as np  # <-- para ajustar posiciones de etiquetas

# ========== CONFIG ==========
st.set_page_config(page_title="BU Breakdown per Site", layout="centered")
TITLE = "BU Breakdown per Site"

# 👉 Ajusta esta ruta a tu fichero:
#DATA_PATH = Path(r"C:\Users\Miguel\OneDrive - MBD 2005 SL\Documentos\Javier Rua\clone github\BUpercentage\IFF Directory 2025-08-04 01_05 EDT.xlsx")
DATA_PATH = Path("IFF Directory 2025-08-04 01_05 EDT.xlsx")

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

# ====== PALETA FIJA DE AZULES (SUAVE) + PIE ======
def fixed_blue_palette(order):
    cmap = cm.get_cmap("Blues")
    n = len(order)
    if n <= 1:
        levels = [0.60]
    else:
        start, end = 0.40, 0.68  # rango claro→medio
        step = (end - start) / (n - 1)
        levels = [end - i * step for i in range(n)]
    return {bu: mcolors.to_hex(cmap(lv)) for bu, lv in zip(order, levels)}

def plot_pie_fixed_blue(counts: dict, order):
    """
    Pie con azules fijos por BU y sin categorías con 0.
    - Labels fuera y % dentro
    - Evita solapes desplazando etiquetas vecinas en el mismo lado
    - Líneas guía hacia cada etiqueta
    """
    palette = fixed_blue_palette(order)
    pairs = [(bu, counts.get(bu, 0)) for bu in order if counts.get(bu, 0) > 0]
    fig, ax = plt.subplots()

    if not pairs:
        ax.text(0.5, 0.5, "No data", ha="center", va="center")
        ax.axis("off")
        table_colors = {bu: "#ffffff" if counts.get(bu, 0) == 0 else palette[bu] for bu in order}
        return fig, table_colors

    labels = [bu for bu, v in pairs]
    sizes  = [v  for _,  v in pairs]
    colors = [palette[bu] for bu in labels]

    # % y conteo dentro; oculta porciones < 2%
    def _autopct_factory(vals, cutoff=2.0):
        total = sum(vals)
        def _fmt(pct):
            if pct < cutoff:
                return ""
            count = int(round(pct/100 * total))
            return f"{pct:.0f}%\n({count})"
        return _fmt

    wedges, texts, autotexts = ax.pie(
        sizes,
        labels=labels,                         # etiquetas FUERA
        colors=colors,
        autopct=_autopct_factory(sizes, 2.0),  # % DENTRO
        startangle=90,
        labeldistance=1.28,                    # un poco más lejos del centro
        pctdistance=0.62,                      # % algo más hacia dentro
        textprops={"fontsize": 10},
        wedgeprops={"linewidth": 1, "edgecolor": "white"},
    )
    ax.axis("equal")

    # --------- Desplazar etiquetas vecinas para evitar solape ----------
    pos = np.array([t.get_position() for t in texts])

    # Lado derecho (x>0) y lado izquierdo (x<0)
    for side in (True, False):
        idx = [i for i, (x, _) in enumerate(pos) if (x > 0) == side]
        if not idx:
            continue
        # Orden por Y
        idx = sorted(idx, key=lambda i: pos[i, 1])

        min_gap = 0.08  # separación vertical mínima entre etiquetas
        for j in range(1, len(idx)):
            i_prev, i_cur = idx[j-1], idx[j]
            y_prev, y_cur = pos[i_prev, 1], pos[i_cur, 1]
            if y_cur - y_prev < min_gap:
                pos[i_cur, 1] = y_prev + min_gap  # empuja hacia abajo

    # Aplicar nuevas posiciones y alinear según lado
    for (x, y), t in zip(pos, texts):
        t.set_position((x, y))
        t.set_ha("left" if x > 0 else "right")

    # --------- Líneas guía (leader lines) hacia las etiquetas ----------
    for w, t in zip(wedges, texts):
        ang = np.deg2rad((w.theta1 + w.theta2) / 2.0)
        # punto en el borde de la porción
        x0, y0 = np.cos(ang) * 1.02, np.sin(ang) * 1.02
        # punto en la etiqueta
        xt, yt = t.get_position()
        ax.plot([x0, xt], [y0, yt], linewidth=0.8, color="#777777")

    plt.tight_layout()

    # Colores para la tabla (0 → blanco)
    table_colors = {bu: ("#ffffff" if counts.get(bu, 0) == 0 else palette[bu]) for bu in order}
    return fig, table_colors

# ========== UI ==========
st.markdown(f"## {TITLE}")
# (ocultamos 'Reading from' y hoja detectada)

try:
    found_sheet, header_row, df = load_excel_auto(DATA_PATH, SHEET_NAME)
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

# ----- Pie chart + colores (azules fijos y sin categorías 0) -----
st.divider()
#st.write("### Breakdown by Business Unit")
fig, color_map = plot_pie_fixed_blue(counts, order)
st.pyplot(fig, clear_figure=True)

# ----- Resumen (tabla con el mismo esquema de color) -----
st.write("#### Resumen")
summary_df = pd.DataFrame({"Business Unit": order, "Users": [counts[k] for k in order]})
summary_df["Share %"] = (summary_df["Users"] / total_users * 100).round(1)

def highlight_row(row):
    bg = color_map.get(row["Business Unit"], "#ffffff")  # 0 → blanco
    return [f"background-color: {bg}; color: black" for _ in row]

styled = summary_df.style.apply(highlight_row, axis=1).format({"Share %": "{:.1f}"})
st.table(styled)

st.caption(f"Total users in {selected_loc}: {total_users}")
