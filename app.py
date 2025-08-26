# app.py
# Streamlit app: Autofill KPIs from an uploaded CSV/XLSX/XLS
# Core fields: SKILL, CALLS, Agents Staffed, AHT, Abandon %
# - Robust file reading (openpyxl for .xlsx, xlrd for .xls)
# - Config load/save (JSON)
# - Copy & Paste helpers + Word/PDF exports (graceful if deps missing)
# - Fortress -> PM Connect aliasing
# - Broad auto-detect synonyms + normalized matching
# - Abandon % by skill in Filled Report
# - NEW: Optional 2nd report (no skill filter) provides Total Agents

import io
import re
import json
import pandas as pd
import streamlit as st

# Optional exports (fail gracefully if missing)
try:
    from docx import Document
    from docx.shared import Inches  # noqa: F401
    HAS_DOCX = True
except Exception:
    HAS_DOCX = False

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import inch
    HAS_PDF = True
except Exception:
    HAS_PDF = False

st.set_page_config(page_title="Autofill Numbers (Core)", layout="wide")

st.title("🧮 Autofill Numbers — Core Fields")
st.caption("Upload your main skill-level report, optionally upload a second (non-skill) report to provide Total Agents. Outputs: Skills, Calls, Agents Staffed, AHT, and Abandon %.")

# ---- Sidebar: Config load/save & Branding ----
st.sidebar.header("Config")
cfg_file = st.sidebar.file_uploader("Load config (JSON)", type=["json"], key="cfg")
company_name = st.sidebar.text_input("Company name (optional)", value="")

loaded_cfg = {}
if cfg_file is not None:
    try:
        loaded_cfg = json.load(cfg_file)
        st.sidebar.success("Config loaded.")
        if not company_name and "company_name" in loaded_cfg:
            company_name = loaded_cfg["company_name"]
    except Exception as e:
        st.sidebar.error(f"Invalid config: {e}")

# Optional: load a sample config committed in your repo at config/sample_report_config.json
if st.sidebar.button("Load sample config from repo", use_container_width=True):
    try:
        with open("config/sample_report_config.json", "r") as f:
            loaded_cfg = json.load(f)
        st.sidebar.success("Loaded sample config from repo.")
        if not company_name and "company_name" in loaded_cfg:
            company_name = loaded_cfg["company_name"]
    except Exception as e:
        st.sidebar.error(f"Couldn't load sample config: {e}")

# ---- Helpers ----
def norm(s):
    """Normalize strings: lowercase, collapse to alnum only (remove spaces/punct)."""
    return re.sub(r"[^a-z0-9]+", "", str(s).lower())

def find_col(df, synonyms):
    """
    Find a column by synonyms using:
      1) exact normalized match
      2) contains match on normalized name
    """
    cols = list(df.columns)
    norm_map = {norm(c): c for c in cols}
    syn_norm = [norm(x) for x in synonyms if str(x).strip()]
    for s in syn_norm:  # exact
        if s in norm_map:
            return norm_map[s]
    for c in cols:  # contains
        nc = norm(c)
        for s in syn_norm:
            if s and (s in nc or nc in s):
                return c
    return None

def idx_or_default(options, value):
    try:
        return options.index(value) if value in options else 0
    except Exception:
        return 0

def read_any(uploaded):
    """Read CSV/XLSX/XLS robustly with explicit engines for Streamlit Cloud."""
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded)
    if name.endswith(".xlsx"):
        return pd.read_excel(uploaded, engine="openpyxl")
    if name.endswith(".xls"):
        return pd.read_excel(uploaded, engine="xlrd")
    try:
        return pd.read_csv(uploaded)
    except Exception:
        uploaded.seek(0)
        return pd.read_excel(uploaded)

def to_percent(series_like):
    """Parse a percent series that may contain '%' or be fractional (0-1). Returns float % (0-100)."""
    s = pd.Series(series_like).astype(str).str.replace('%', '', regex=False)
    vals = pd.to_numeric(s, errors='coerce')
    if vals.dropna().max() is not None and vals.dropna().max() <= 1.0:
        vals = vals * 100.0
    return vals

# ------------------------------
# 1) Main report (skill-level)
# ------------------------------
uploaded = st.file_uploader("Main report (CSV/XLSX/XLS) — contains skill-level rows", type=["csv", "xlsx", "xls"], key="main")
if uploaded is None:
    st.info("Upload the main CSV/Excel file to begin.")
    st.stop()

try:
    df = read_any(uploaded)
except Exception as e:
    st.error(f"Could not read main file: {e}")
    st.stop()

if df.empty:
    st.warning("The main uploaded file appears to be empty.")
    st.stop()

st.subheader("Preview — Main Report (first 20 rows)")
preview_df = df.head(20).copy()
st.dataframe(preview_df, use_container_width=True)

# ---- Column Mapping (main) ----
st.subheader("Column Mapping — Main Report")

# Synonyms
SKILL_SYNS = ["skill", "skill name", "skill group", "group", "queue", "split", "team", "program", "department", "dept", "category", "line of business", "lob"]
CALLS_SYNS = ["calls", "total calls", "calls offered", "offered", "inbound calls", "in calls", "total contacts", "contacts", "total interactions", "volume"]
AGENTS_SYNS = ["agents staffed", "agents", "agent count", "staffed agents", "distinct agents", "distinct agent count", "unique agents", "logged in agents", "logged-in agents", "logged in", "agents (distinct)", "agents (unique)"]
AHT_SYNS = ["average handle time", "aht", "avg handle time", "avg handling time", "avg handle", "average handling time", "aht (s)", "aht (sec)", "talk+hold+acw", "handle time"]
ABAND_CNT_SYNS = ["abandoned count", "abandoned", "abandon count", "aband count", "abandoned calls", "aband qty", "aband num", "aband total"]
ABAND_PCT_SYNS = ["abandon %", "a]()
