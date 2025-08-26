# app.py
# Streamlit app: Autofill KPIs from an uploaded CSV/XLSX/XLS
# Core fields: SKILL, CALLS, Agents Staffed, AHT, Abandon %
# - Robust file reading (openpyxl for .xlsx, xlrd for .xls)
# - Config load/save (JSON)
# - Copy & Paste helpers + Word/PDF exports (graceful if deps missing)
# - Fortress -> PM Connect aliasing
# - Broad auto-detect synonyms + normalized matching

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
st.caption("Upload a CSV or Excel report, map columns once (or load a config), and get: Skills, Calls, Agents Staffed, AHT, and Abandon %.")

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

# ---- Upload data file ----
uploaded = st.file_uploader("Attach your report (CSV/XLSX/XLS)", type=["csv", "xlsx", "xls"])
if uploaded is None:
    st.info("Upload a CSV or Excel file to begin.")
    st.stop()

# Read
try:
    df = read_any(uploaded)
except Exception as e:
    st.error(f"Could not read file: {e}")
    st.stop()

if df.empty:
    st.warning("The uploaded file appears to be empty.")
    st.stop()

st.subheader("Preview (first 20 rows)")
preview_df = df.head(20).copy()
st.dataframe(preview_df, use_container_width=True)

# ---- Column Mapping ----
st.subheader("Column Mapping")

# Synonyms
SKILL_SYNS = ["skill", "skill name", "skill group", "group", "queue", "split", "team", "program", "department", "dept", "category", "line of business", "lob"]
CALLS_SYNS = ["calls", "total calls", "calls offered", "offered", "inbound calls", "in calls", "total contacts", "contacts", "total interactions", "volume"]
AGENTS_SYNS = ["agents staffed", "agents", "agent count", "staffed agents", "distinct agents", "distinct agent count", "unique agents", "logged in agents", "logged-in agents", "logged in", "agents (distinct)", "agents (unique)"]
AHT_SYNS = ["average handle time", "aht", "avg handle time", "avg handling time", "avg handle", "average handling time", "aht (s)", "aht (sec)", "talk+hold+acw", "handle time"]
ABAND_CNT_SYNS = ["abandoned count", "abandoned", "abandon count", "aband count", "abandoned calls", "aband qty", "aband num", "aband total"]
ABAND_PCT_SYNS = ["abandon %", "abandoned (%rec)", "abandoned percent", "abandoned %", "abandonment rate", "abandon rate", "aband %", "aband pct", "abandonment %", "abandonment pct", "abn %", "abn pct"]

# Auto-guesses
skill_guess   = find_col(df, SKILL_SYNS)
calls_guess   = find_col(df, CALLS_SYNS)
agents_guess  = find_col(df, AGENTS_SYNS)
aht_guess     = find_col(df, AHT_SYNS)
aband_cnt_guess = find_col(df, ABAND_CNT_SYNS)
aband_pct_guess = find_col(df, ABAND_PCT_SYNS)

cols = list(df.columns)
def cfg_get(key, default):
    return loaded_cfg.get(key, default) if loaded_cfg else default

skill_col  = st.selectbox("Skill / Group column", cols, index=idx_or_default(cols, cfg_get("skill_col",  skill_guess  or cols[0])))
calls_col  = st.selectbox("Calls column",        cols, index=idx_or_default(cols, cfg_get("calls_col",  calls_guess  or cols[0])))
agents_col = st.selectbox("Agents Staffed column", cols, index=idx_or_default(cols, cfg_get("agents_col", agents_guess or cols[0])))
aht_col    = st.selectbox("Average Handle Time (AHT) column", cols, index=idx_or_default(cols, cfg_get("aht_col",    aht_guess    or cols[0])))

# NOTE: label now uses "Abandon %"
abandoned_pct_col = st.selectbox("Abandon % column (optional)", ["<none>"] + cols,
                                 index=idx_or_default(["<none>"]+cols, cfg_get("abandoned_rate_col", aband_pct_guess if aband_pct_guess else "<none>")))
abandoned_count_col = st.selectbox("Abandoned (count) column (optional, used if % is missing)", ["<none>"] + cols,
                                   index=idx_or_default(["<none>"]+cols, cfg_get("abandoned_count_col", aband_cnt_guess if aband_cnt_guess else "<none>")))

# ---- Skills of interest (optional list for sectioned output) ----
# NOTE: "Fortress" will be aliased to "PM Connect"
default_skills = ["B2B Member Success", "B2B Success Activation", "B2B Success Info", "B2B Success Tech Support", "MS Activation", "MS Info", "MS Loyalty", "MS Tech Support", "PM Connect"]
skills_list = st.text_area("Skills of interest (one per line)", value="\n".join(cfg_get("skills", default_skills)))

# Parse user skills, alias and dedupe
raw_skills = [s.strip() for s in skills_list.splitlines() if s.strip()]
skills_wanted = []
for s in raw_skills:
    if s.lower() == "fortress":
        s = "PM Connect"
    if s not in skills_wanted:
        skills_wanted.append(s)

# ---- Save current config ----
cfg = {
    "company_name": company_name,
    "skill_col": skill_col,
    "calls_col": calls_col,
    "agents_col": agents_col,
    "aht_col": aht_col,
    "abandoned_rate_col": abandoned_pct_col,
    "abandoned_count_col": abandoned_count_col,
    "skills": skills_wanted,
}
st.download_button("⬇️ Download current config (JSON)", data=json.dumps(cfg, indent=2).encode("utf-8"),
                   file_name="autofill_config.json", mime="application/json")

# ---- Defensive: verify required columns exist ----
missing = [c for c in [skill_col, calls_col, agents_col, aht_col] if c not in df.columns]
if missing:
    st.error(f"Selected columns not found in file: {missing}")
    st.stop()

# ---- Normalize & alias skills in the DATAFRAME ----
df[skill_col] = df[skill_col].astype(str).str.strip()
df.loc[df[skill_col].str.lower() == "fortress", skill_col] = "PM Connect"

# ---- Calculations ----
calls_num  = pd.to_numeric(df[calls_col],  errors="coerce").fillna(0)
agents_num = pd.to_numeric(df[agents_col], errors="coerce").fillna(0)

# Abandon % series: prefer provided %, else compute from counts
rates = None
if abandoned_pct_col != "<none>" and abandoned_pct_col in df.columns:
    rates = to_percent(df[abandoned_pct_col])

if rates is None and abandoned_count_col != "<none>" and abandoned_count_col in df.columns:
    aband_num = pd.to_numeric(df[abandoned_count_col], errors="coerce")
    with pd.option_context('mode.use_inf_as_na', True):
        rates = (aband_num / calls_num.replace(0, pd.NA)) * 100

# Totals
total_calls  = int(calls_num.sum())
total_agents = int(agents_num.sum())

# Total Abandon %: prefer counts for accuracy if present, else weighted by calls
if abandoned_count_col != "<none>" and abandoned_count_col in df.columns and total_calls > 0:
    aband_num_total = pd.to_numeric(df[abandoned_count_col], errors="coerce").fillna(0).sum()
    total_abandon_pct = (aband_num_total / total_calls) * 100
elif rates is not None and total_calls > 0:
    total_abandon_pct = ((rates.fillna(0) / 100.0) * calls_num).sum() / total_calls * 100
else:
    total_abandon_pct = None

# By-skill table with "Abandon %"
by_skill_core = pd.DataFrame({
    "SKILL": df[skill_col].astype(str),
    "CALLS": calls_num.astype("Int64"),
    "Agents Staffed": agents_num.astype("Int64"),
    "AHT": df[aht_col].astype(str),
})

if rates is not None:
    by_skill_core["Abandon %"] = rates.round(2).astype(str) + "%"
else:
    by_skill_core["Abandon %"] = "N/A"

# ---- Build the filled report (Markdown) ----
md = io.StringIO()
def writeln(s=""):
    md.write(s + "\n")

title = (company_name + " — " if company_name else "") + "Autofilled Metrics (Core)"
writeln(f"## {title}\n")

writeln(f"### 3. Total Calls\n**{total_calls}**\n")
writeln(f"### 4. Agents Staffed (sum of per-skill)\n**{total_agents}**\n")
writeln("### 6. Abandon %")
writeln(f"**{(str(round(total_abandon_pct, 2)) + '%') if total_abandon_pct is not None else 'N/A'}**\n")

writeln("### 7. AHT (By Group)")
for sk in skills_wanted:
    mask = by_skill_core["SKILL"].astype(str).str.lower() == sk.lower()
    val = by_skill_core.loc[mask, "AHT"]
    writeln(f"- **{sk}:** {val.iloc[0] if len(val) else 'Not found in this report'}")

report_md = md.getvalue()

# ==============================
# UI SECTIONS
# ==============================

st.subheader("Filled Report (Core)")
st.markdown(report_md)

# ---- Copy & Paste helpers ----
st.subheader("Copy & Paste")

tabs = st.tabs([
    "📋 Report (Markdown)",
    "📋 KPIs (CSV)",
    "📋 By-Skill Core Table (CSV)",
    "📋 By-Skill Core Table (TSV)",
    "📋 Preview (first 20 rows CSV)"
])

with tabs[0]:
    st.write("Use the copy icon in the code block to copy the entire Markdown report:")
    st.code(report_md, language="markdown")

with tabs[1]:
    kpi_df = pd.DataFrame([{
        "Total Calls": total_calls,
        "Agents Staffed (sum of per-skill)": total_agents,
        "Total Abandon %": (round(total_abandon_pct, 2) if total_abandon_pct is not None else None)
    }])
    # Pretty display with % suffix
    kpi_df_display = kpi_df.copy()
    if kpi_df_display.loc[0, "Total Abandon %"] is not None:
        kpi_df_display.loc[0, "Total Abandon %"] = f"{kpi_df_display.loc[0, 'Total Abandon %']:.2f}%"
    st.dataframe(kpi_df_display, use_container_width=True)
    st.write("Copy the CSV below:")
    st.code(kpi_df.to_csv(index=False), language="text")

with tabs[2]:
    csv_text = by_skill_core.to_csv(index=False)
    st.dataframe(by_skill_core, use_container_width=True)
    st.write("Copy the CSV below:")
    st.code(csv_text, language="text")

with tabs[3]:
    tsv_text = by_skill_core.to_csv(index=False, sep="\t")
    st.dataframe(by_skill_core, use_container_width=True)
    st.write("Copy the TSV below:")
    st.code(tsv_text, language="text")

with tabs[4]:
    prev_csv = preview_df.to_csv(index=False)
    st.dataframe(preview_df, use_container_width=True)
    st.write("Copy the CSV below:")
    st.code(prev_csv, language="text")

# ---- Downloads ----
st.subheader("Downloads")
st.download_button(
    label="⬇️ Download report (Markdown)",
    data=report_md.encode("utf-8"),
    file_name="filled_report_core.md",
    mime="text/markdown",
)

# Visual table
st.subheader("By-Skill Table — Core Fields")
st.dataframe(by_skill_core, use_container_width=True)

# ---- Word export (if available) ----
if HAS_DOCX:
    def build_docx(md_text, company=""):
        doc = Document()
        doc.add_heading(company + (" — " if company else "") + "Autofilled Metrics (Core)", level=1)
        for line in md_text.splitlines():
            if line.startswith("### "):
                doc.add_heading(line.replace("### ", ""), level=2)
            elif line.startswith("## "):
                continue
            else:
                if line.strip():
                    doc.add_paragraph(line)
        bio = io.BytesIO()
        doc.save(bio)
        bio.seek(0)
        return bio.getvalue()

    docx_bytes = build_docx(report_md, company_name)
    st.download_button(
        label="⬇️ Download report (Word .docx)",
        data=docx_bytes,
        file_name="filled_report_core.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
else:
    st.info("python-docx not installed; Word export disabled.")

# ---- PDF export (if available) ----
if HAS_PDF:
    def build_pdf(md_text, company=""):
        bio = io.BytesIO()
        c = canvas.Canvas(bio, pagesize=letter)
        width, height = letter
        left_margin = 0.75 * inch
        top_margin = 0.75 * inch
        bottom_margin = 0.75 * inch

        y = height - top_margin
        title = company + (" — " if company else "") + "Autofilled Metrics (Core)"
        c.setFont("Times-Bold", 14)
        c.drawString(left_margin, y, title)
        y -= 0.3*inch

        c.setFont("Times-Roman", 11)
        max_width_chars = 95

        import textwrap as _tw
        for line in md_text.splitlines():
            if not line.strip():
                y -= 0.18*inch
                if y < bottom_margin:
                    c.showPage(); y = height - top_margin; c.setFont("Times-Roman", 11)
                continue
            for w in _tw.wrap(line, width=max_width_chars):
                c.drawString(left_margin, y, w)
                y -= 0.18*inch
                if y < bottom_margin:
                    c.showPage(); y = height - top_margin; c.setFont("Times-Roman", 11)

        c.showPage()
        c.save()
        bio.seek(0)
        return bio.getvalue()

    pdf_bytes = build_pdf(report_md, company_name)
    st.download_button(
        label="⬇️ Download report (PDF)",
        data=pdf_bytes,
        file_name="filled_report_core.pdf",
        mime="application/pdf",
    )
else:
    st.info("reportlab not installed; PDF export disabled.")
