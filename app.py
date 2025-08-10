import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from pathlib import Path

# ---------- Config ----------
st.set_page_config(page_title="BU Breakdown per Site", layout="centered")

TITLE = "BU Breakdown per Site"
REQUIRED_COLS = ["Business Unit", "Location", "Real Estate ID"]

# 👉👉 Cambia esta ruta por la tuya (usa r'...' o dobles barras \\ en Windows)
DATA_PATH = Path(r"C:\Users\Miguel\OneDrive - MBD 2005 SL\Documentos\Javier Rua\clone github\BUpercentage\IFF Directory 2025-08-04 01_05 EDT.xlsx")

# ---------- Helpers ----------
@st.cache_data
def load_excel_from_path(path: Path):
    if not path.exists():
        raise FileNotFoundError(f"No existe el archivo en: {path}")
    # Abrimos en binario para evitar algunos locks raros
    with open(path, "rb") as fh:
        df = pd.read_excel(fh, engine="openpyxl")
    # Limpieza de nombres de columnas
    df.columns = df.columns.str.strip()
    colmap = {c.lower(): c for c in df.columns}
    def find_col(name):
        return colmap.get(name.lower(), name)
    needed = [find_col(c) for c in REQUIRED_COLS if find_col(c) in df.columns]
    if len(needed) < 3:
        missing = set(REQUIRED_COLS) - set(needed)
        raise ValueError(f"Faltan columnas requeridas en el Excel: {', '.join(missing)}")
    out = df[needed].copy()
    out.columns = REQUIRED_COLS
    for c in REQUIRED_COLS:
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
st.caption(f"Reading from: `{DATA_PATH}`")

# Cargar datos desde la ruta fija
try:
    df = load_excel_from_path(DATA_PATH)
except PermissionError as e:
    st.error(f"No se pudo abrir el archivo (Permission denied). "
             f"¿Está abierto en Excel/OneDrive? Ciérralo y actualiza. Ruta: {DATA_PATH}")
    st.stop()
except FileNotFoundError as e:
    st.error(str(e))
    st.stop()
except Exception as e:
    st.error(f"No se pudo cargar el Excel: {e}")
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

# ----- Pie chart -----
st.divider()
st.write("### Breakdown by Business Unit")
fig = plot_pie(counts)
st.pyplot(fig, clear_figure=True)

# ----- Resumen tabla -----
st.write("#### Resumen")
summary_df = pd.DataFrame({"Business Unit": order, "Users": [counts[k] for k in order]})
summary_df["Share %"] = (summary_df["Users"] / total_users * 100).round(1)
st.dataframe(summary_df, use_container_width=True)
st.caption(f"Total users in {selected_loc}: {total_users}")


