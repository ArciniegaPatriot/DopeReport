# app.py
import io
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

st.set_page_config(page_title="Autofill Numbers", layout="wide")

st.title("🧮 Autofill Numbers from Report")
st.caption("Upload a CSV or Excel report, map columns once (or load a config), and auto-fill your metrics.")

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

# ---- Helpers ----
def find_col(df, candidates):
    cols_lower = {c.strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().lower()
        if key in cols_lower:
            return cols_lower[key]
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
        # Requires openpyxl
        return pd.read_excel(uploaded, engine="openpyxl")
    if name.endswith(".xls"):
        # Requires xlrd (xls only)
        return pd.read_excel(uploaded, engine="xlrd")
    # Fallback try CSV then Excel
    try:
        return pd.read_csv(uploaded)
    except Exception:
        uploaded.seek(0)
        return pd.read_excel(uploaded)

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

st.subheader("Preview")
st.dataframe(df.head(20))

# ---- Column Mapping ----
st.subheader("Column Mapping")

# Auto-guesses
skill_guess = find_col(df, ["SKILL", "Group", "Queue", "Split"])
calls_guess = find_col(df, ["CALLS", "Total Calls", "Calls"])
agents_guess = find_col(df, ["AGENT distinct count", "Agents", "Agents Staffed", "Agent Count"])
aband_count_guess = find_col(df, ["ABANDONED count", "Abandoned Count"])
aband_rate_guess = find_col(df, ["ABANDONED (%rec)", "Abandon Rate", "Abandonment Rate"])
aht_guess = find_col(df, ["Average HANDLE TIME", "AHT", "Avg Handle Time"])

cols = list(df.columns)
def cfg_get(key, default):
    return loaded_cfg.get(key, default) if loaded_cfg else default

skill_col = st.selectbox("Skill / Group column", cols, index=idx_or_default(cols, cfg_get("skill_col", skill_guess or cols[0])))
calls_col = st.selectbox("Calls column", cols, index=idx_or_default(cols, cfg_get("calls_col", calls_guess or cols[0])))
agents_col = st.selectbox("Agents staffed column", cols, index=idx_or_default(cols, cfg_get("agents_col", agents_guess or cols[0])))
aht_col = st.selectbox("AHT column", cols, index=idx_or_default(cols, cfg_get("aht_col", aht_guess or cols[0])))

abandoned_count_col = st.selectbox("Abandoned count column (optional)", ["<none>"] + cols,
                                   index=idx_or_default(["<none>"]+cols, cfg_get("abandoned_count_col", aband_count_guess if aband_count_guess else "<none>")))
abandoned_rate_col = st.selectbox("Abandoned % column (optional)", ["<none>"] + cols,
                                  index=idx_or_default(["<none>"]+cols, cfg_get("abandoned_rate_col", aband_rate_guess if aband_rate_guess else "<none>")))

# ---- Skills of interest (customizable list) ----
default_skills = [
    "B2B Member Success",
    "B2B Success Activation",
    "B2B Success Info",
    "B2B Success Tech Support",
    "Fortress",
    "MS Activation",
    "MS Info",
    "MS Loyalty",
    "MS Tech Support",
    "PM Connect",
]
skills_list = st.text_area("Skills of interest (one per line)", value="\n".join(cfg_get("skills", default_skills)))
skills_wanted = [s.strip() for s in skills_list.splitlines() if s.strip()]

# ---- Save current config ----
cfg = {
    "company_name": company_name,
    "skill_col": skill_col,
    "calls_col": calls_col,
    "agents_col": agents_col,
    "aht_col": aht_col,
    "abandoned_count_col": abandoned_count_col,
    "abandoned_rate_col": abandoned_rate_col,
    "skills": skills_wanted,
}
st.download_button("⬇️ Download current config (JSON)", data=json.dumps(cfg, indent=2).encode("utf-8"),
                   file_name="autofill_config.json", mime="application/json")

# ---- Defensive: verify required columns exist ----
missing = [c for c in [skill_col, calls_col, agents_col, aht_col] if c not in df.columns]
if missing:
    st.error(f"Selected columns not found in file: {missing}")
    st.stop()

# ---- Calculations ----
calls_num = pd.to_numeric(df[calls_col], errors="coerce")
agents_num = pd.to_numeric(df[agents_col], errors="coerce")
aband_num = pd.to_numeric(df[abandoned_count_col], errors="coerce") if (abandoned_count_col != "<none>" and abandoned_count_col in df.columns) else None

total_calls = int(calls_num.fillna(0).sum())
total_agents = int(agents_num.fillna(0).sum())

# Per-skill abandon rate
by_skill = df[[skill_col]].copy()
if abandoned_rate_col != "<none>" and abandoned_rate_col in df.columns:
    by_skill["Abandonment Rate"] = df[abandoned_rate_col].astype(str)
else:
    if aband_num is not None and calls_num is not None:
        with pd.option_context('mode.use_inf_as_na', True):
            rate = (aband_num / calls_num) * 100
        by_skill["Abandonment Rate"] = rate.round(2).astype(str) + "%"
    else:
        by_skill["Abandonment Rate"] = "N/A"

# AHT by skill
by_skill["AHT"] = df[aht_col].astype(str)
by_skill.rename(columns={skill_col: "SKILL"}, inplace=True)

# Total Abandoned %rec
if aband_num is not None and calls_num.fillna(0).sum() > 0:
    total_abandoned_str = f"{(aband_num.fillna(0).sum() / calls_num.fillna(0).sum())*100:.2f}%"
elif abandoned_col != "<none>" and abandoned_col in df.columns:
    total_abandoned_str = "N/A (needs counts)"
else:
    total_abandoned_str = "N/A"

# ---- Build the filled report (Markdown) ----
md = io.StringIO()
def writeln(s=""):
    md.write(s + "\n")

title = "Autofilled Metrics"
if company_name:
    title = f"{Patriot_Mobile} — " + title

writeln(f"## {title}\n")
writeln(f"### 3. Total Calls\n**{total_calls}**\n")
writeln(f"### 4. Total Agents Staffed\n**{total_agents}**\n")
writeln("### 5. Shrinkage")
writeln("*Not available unless shrinkage columns are present in your file.*")
writeln("- Daily Shrinkage – N/A")
writeln("- Total Shrinkage – N/A")
writeln("- Discretionary Shrinkage – N/A")
writeln("- Non-Discretionary Shrinkage – N/A")

writeln("### 6. Abandoned (all in & by split)")
writeln(f"- **Total Abandoned:** **{total_abandonment_rate_str}**")
for sk in skills_wanted:
    mask = by_skill["SKILL"].astype(str).str.lower() == sk.lower()
    val = by_skill.loc[mask, "Abandoned"]
    writeln(f"- **SKILL: {sk}:** {val.iloc[0] if len(val) else 'Not found in this report'}")
writeln("")

writeln("### 7. AHT (By Group)")
for sk in skills_wanted:
    mask = by_skill["SKILL"].astype(str).str.lower() == sk.lower()
    val = by_skill.loc[mask, "AHT"]
    writeln(f"- **{sk}:** {val.iloc[0] if len(val) else 'Not found in this report'}")

report_md = md.getvalue()

st.subheader("Filled Report")
st.markdown(report_md)

# Download markdown
st.download_button(
    label="⬇️ Download report (Markdown)",
    data=report_md.encode("utf-8"),
    file_name="filled_report.md",
    mime="text/markdown",
)

# Show by-skill table for visibility
st.subheader("By-Skill Table")
st.dataframe(by_skill)

# ---- Word export (if available) ----
if HAS_DOCX:
    def build_docx(md_text, company=""):
        doc = Document()
        doc.add_heading(company + (" — " if company else "") + "Autofilled Metrics", level=1)
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
        file_name="filled_report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
else:
    st.info("python-docx not installed; Word export disabled.")

# ---- PDF export (if available) ----
if HAS_PDF:
    def build_pdf(md_text, company=""):
        from reportlab.lib.pagesizes import letter
        from reportlab.pdfgen import canvas
        from reportlab.lib.units import inch
        bio = io.BytesIO()
        c = canvas.Canvas(bio, pagesize=letter)
        width, height = letter
        left_margin = 0.75 * inch
        right_margin = 0.75 * inch
        top_margin = 0.75 * inch
        bottom_margin = 0.75 * inch

        y = height - top_margin
        title = company + (" — " if company else "") + "Autofilled Metrics"
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
        file_name="filled_report.pdf",
        mime="application/pdf",
    )
else:
    st.info("reportlab not installed; PDF export disabled.")
