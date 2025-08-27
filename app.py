# app.py — Protected KPI app with login + Five9/SFTP sources
# -------------------------------------------------------------------
# Auth: streamlit-authenticator (bcrypt hashes in secrets or auth.yaml)
# Data sources: Manual | Public CSV URL | Local folder | SFTP (Five9 Scheduled) | Five9 SOAP
# Outputs: Skills, Calls, Agents Staffed, AHT, Abandon %, per-skill sections, downloads

import os
import io
import re
import glob
import time
import json
import datetime as dt
import pandas as pd
import streamlit as st

# ---------- Optional deps (graceful fallback) ----------
try:
    from docx import Document
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

try:
    import requests
    HAS_REQUESTS = True
except Exception:
    HAS_REQUESTS = False

try:
    from zeep import Client, Settings
    from zeep.transports import Transport
    HAS_ZEEP = True
except Exception:
    HAS_ZEEP = False

try:
    import paramiko
    HAS_SFTP = True
except Exception:
    HAS_SFTP = False

# ---------- AUTH BLOCK ----------
import streamlit_authenticator as stauth
try:
    import yaml
    from yaml.loader import SafeLoader
    HAS_YAML = True
except Exception:
    HAS_YAML = False

st.set_page_config(page_title="Autofill Numbers (Protected)", layout="wide")

def _load_auth_cfg():
    # Preferred: from .streamlit/secrets.toml
    if "auth" in st.secrets:
        return st.secrets["auth"]
    # Fallback: auth.yaml in repo root
    if HAS_YAML and os.path.exists("auth.yaml"):
        with open("auth.yaml", "r") as f:
            return yaml.load(f, Loader=SafeLoader)
    raise RuntimeError("No auth config found. Add [auth] to secrets or provide auth.yaml.")

def _allowed(username: str, email: str, cfg: dict) -> bool:
    allow = cfg.get("allowlist", {}) or {}
    allow_u = set(allow.get("usernames", []) or [])
    allow_e = set(allow.get("emails", []) or [])
    # If both are empty → all authenticated users allowed
    if not allow_u and not allow_e:
        return True
    return (username in allow_u) or (email in allow_e)

def require_login() -> bool:
    try:
        cfg = _load_auth_cfg()
    except Exception as e:
        st.error(str(e))
        return False

    cookie = cfg.get("cookie", {})
    credentials = cfg.get("credentials", {})
    preauth = cfg.get("preauthorized", {})  # optional key supported by the lib

    authenticator = stauth.Authenticate(
        credentials,
        cookie.get("name", "auth_cookie"),
        cookie.get("key", "please_change_me"),
        cookie.get("expiry_days", 30),
        preauth
    )

    st.markdown("<h2 style='margin-top:0'>🔐 Sign in</h2>", unsafe_allow_html=True)
    name, auth_status, username = authenticator.login("main")

    if auth_status:
        user_rec = (credentials.get("usernames", {}) or {}).get(username, {})
        email = user_rec.get("email", "")
        if not _allowed(username, email, cfg):
            st.error("You are authenticated but not authorized for this app.")
            return False
        authenticator.logout("Logout", "sidebar")
        st.sidebar.success(f"Welcome, {name}")
        if user_rec.get("role"):
            st.sidebar.caption(f"Role: {user_rec['role']}")
        return True
    elif auth_status is False:
        st.error("Username/password is incorrect")
        return False
    else:
        st.info("Enter your credentials to continue.")
        return False

# ---------- Helpers (shared by the app) ----------
def norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(s).lower())

def find_col(df: pd.DataFrame, synonyms) -> str | None:
    cols = list(df.columns)
    norm_map = {norm(c): c for c in cols}
    syn_norm = [norm(x) for x in synonyms if str(x).strip()]
    for s in syn_norm:
        if s in norm_map:
            return norm_map[s]
    for c in cols:
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

def read_any(uploaded_or_bytes, name_hint: str | None = None):
    # Bytes-like: try CSV then Excel
    if isinstance(uploaded_or_bytes, (bytes, bytearray)):
        bio = io.BytesIO(uploaded_or_bytes)
        try:
            bio.seek(0); return pd.read_csv(bio)
        except Exception:
            bio.seek(0); return pd.read_excel(bio)
    # File-like / UploadedFile
    name = (getattr(uploaded_or_bytes, "name", None) or name_hint or "").lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_or_bytes)
    if name.endswith(".xlsx"):
        return pd.read_excel(uploaded_or_bytes, engine="openpyxl")
    if name.endswith(".xls"):
        return pd.read_excel(uploaded_or_bytes, engine="xlrd")
    try:
        return pd.read_csv(uploaded_or_bytes)
    except Exception:
        if hasattr(uploaded_or_bytes, "seek"):
            uploaded_or_bytes.seek(0)
        return pd.read_excel(uploaded_or_bytes)

def to_percent(series_like):
    s = pd.Series(series_like).astype(str).str.replace('%', '', regex=False)
    vals = pd.to_numeric(s, errors='coerce')
    if vals.dropna().max() is not None and vals.dropna().max() <= 1.0:
        vals = vals * 100.0
    return vals

def fetch_csv_url(url: str) -> tuple[pd.DataFrame | None, dict]:
    if not HAS_REQUESTS:
        return None, {"error": "requests not installed"}
    if not url:
        return None, {"error": "No URL provided"}
    try:
        r = requests.get(url, timeout=45)
        if r.status_code >= 300:
            return None, {"error": f"{r.status_code}: {r.text[:300]}", "source": url}
        return read_any(r.content, name_hint=url), {"source": url, "bytes": len(r.content)}
    except Exception as e:
        return None, {"error": str(e), "source": url}

def load_latest_local_csv(folder: str, pattern: str = "*.csv") -> tuple[pd.DataFrame | None, dict]:
    try:
        paths = sorted(glob.glob(os.path.join(folder, pattern)), key=lambda p: os.path.getmtime(p))
        if not paths:
            return None, {"error": f"No files matching {pattern} in {folder}"}
        latest = paths[-1]
        with open(latest, "rb") as f:
            df = read_any(f, name_hint=latest)
        return df, {"source": latest, "mtime": os.path.getmtime(latest)}
    except Exception as e:
        return None, {"error": str(e)}

def load_latest_sftp_csv(host: str, port: int, username: str, password: str,
                         remote_dir: str, pattern: str = "*.csv") -> tuple[pd.DataFrame | None, dict]:
    if not HAS_SFTP:
        return None, {"error": "paramiko not installed"}
    try:
        import posixpath
        transport = paramiko.Transport((host, port))
        transport.connect(username=username, password=password)
        sftp = paramiko.SFTPClient.from_transport(transport)
        try:
            sftp.chdir(remote_dir)
        except IOError:
            sftp.close(); transport.close()
            return None, {"error": f"Remote directory not found: {remote_dir}"}
        # Simple fnmatch using regex
        rx = re.compile("^" + pattern.replace(".", r"\.").replace("*", ".*") + "$")
        files = [f for f in sftp.listdir_attr(".") if rx.match(f.filename)]
        if not files:
            sftp.close(); transport.close()
            return None, {"error": f"No files matching {pattern} in {remote_dir}"}
        latest = max(files, key=lambda f: f.st_mtime)
        with sftp.file(latest.filename, "rb") as fh:
            blob = fh.read()
        sftp.close(); transport.close()
        df = read_any(blob, name_hint=latest.filename)
        return df, {"source": f"sftp://{host}{posixpath.join(remote_dir, latest.filename)}", "mtime": latest.st_mtime}
    except Exception as e:
        return None, {"error": str(e)}

def _dt_to_five9(dt_obj: dt.datetime | dt.date) -> str:
    if isinstance(dt_obj, dt.date) and not isinstance(dt_obj, dt.datetime):
        dt_obj = dt.datetime(dt_obj.year, dt_obj.month, dt_obj.day, 0, 0, 0)
    return dt_obj.strftime("%Y-%m-%d %H:%M:%S GMT")

def five9_run_report_to_df(wsdl_url: str, username: str, password: str,
                           report_name: str, folder: str | None,
                           date_from: dt.date, date_to: dt.date) -> tuple[pd.DataFrame, dict]:
    if not (HAS_ZEEP and HAS_REQUESTS):
        raise RuntimeError("Missing dependency: zeep and/or requests")

    session = requests.Session()
    session.auth = (username, password)   # SOAP reporting uses Basic Auth
    transport = Transport(session=session, timeout=120)
    settings = Settings(strict=False, xml_huge_tree=True)
    client = Client(wsdl=wsdl_url, transport=transport, settings=settings)
    svc = client.service

    start_str = _dt_to_five9(dt.datetime.combine(date_from, dt.time.min))
    end_str   = _dt_to_five9(dt.datetime.combine(date_to, dt.time.max).replace(microsecond=0))

    # Try two common signatures
    try:
        result_id = svc.runReport(reportName=report_name, folderName=(folder or None),
                                  timeFrom=start_str, timeTo=end_str)
    except Exception:
        try:
            req_type = None
            try: req_type = client.get_type("ns0:reportRequest")
            except Exception: pass
            payload = {"reportName": report_name, "timeFrom": start_str, "timeTo": end_str}
            if folder: payload["folderName"] = folder
            req_obj = req_type(**payload) if req_type else payload
            result_id = svc.runReport(req_obj)
        except Exception as e2:
            raise RuntimeError(f"runReport failed: {e2}")

    # Poll until complete
    for _ in range(120):
        if not svc.isReportRunning(result_id):
            break
        time.sleep(1.2)
    else:
        raise RuntimeError("Report timed out waiting for completion.")

    # Fetch CSV
    csv_blob = svc.getReportResultCsv(result_id)
    raw = csv_blob if isinstance(csv_blob, (bytes, bytearray)) else str(csv_blob).encode("utf-8", "ignore")
    try:
        import csv
        sample = raw[:4096].decode("utf-8", "ignore")
        dialect = csv.Sniffer().sniff(sample)
        df = pd.read_csv(io.BytesIO(raw), dialect=dialect)
    except Exception:
        df = pd.read_csv(io.BytesIO(raw))
    meta = {"source": f"Five9 report '{report_name}'", "from": start_str, "to": end_str}
    return df, meta

# ---------- The main app (runs only after login) ----------
def run_app():
    st.title("🧮 Autofill Numbers — Core Fields")
    st.caption("Protected app. Data sources: Manual, URL, Local folder, SFTP (Five9 scheduled), Five9 SOAP.")

    # ---------------- Sidebar: Main report source ----------------
    st.sidebar.header("Main Report — Data Source")
    source_type = st.sidebar.radio(
        "Choose source",
        ["Manual upload", "Public CSV URL", "Local folder (latest *.csv)", "SFTP (Five9 Scheduled)", "Five9 API (SOAP)"],
        index=0,
    )

    refresh_secs = st.sidebar.number_input("Auto-refresh seconds (manual reload)", 10, 3600, 60, 5)
    if st.sidebar.button("🔄 Reload now"):
        st.rerun()

    # URL/Local inputs
    main_url    = st.sidebar.text_input("Main CSV URL", os.getenv("MAIN_CSV_URL", ""))
    main_folder = st.sidebar.text_input("Main local folder", "./data")
    main_glob   = st.sidebar.text_input("Main filename pattern", "*.csv")

    # SFTP inputs
    with st.sidebar.expander("SFTP settings (Five9 Scheduled)", expanded=(source_type == "SFTP (Five9 Scheduled)")):
        sftp_host = st.text_input("Host", os.getenv("SFTP_HOST", ""))
        sftp_port = st.number_input("Port", 1, 65535, int(os.getenv("SFTP_PORT", "22") or 22))
        sftp_user = st.text_input("Username", os.getenv("SFTP_USER", ""))
        sftp_pass = st.text_input("Password", os.getenv("SFTP_PASS", ""), type="password")
        sftp_dir  = st.text_input("Remote directory", os.getenv("SFTP_DIR", "/"))
        sftp_pat  = st.text_input("Filename pattern", "*.csv")

    # Five9 SOAP inputs
    with st.sidebar.expander("Five9 (SOAP) settings", expanded=(source_type == "Five9 API (SOAP)")):
        five9_wsdl   = st.text_input("Five9 WSDL URL", os.getenv("FIVE9_WSDL", "https://api.five9.com/wsadmin/v12/AdminWebService?wsdl"))
        five9_user   = st.text_input("Five9 username", os.getenv("FIVE9_USER", ""))
        five9_pass   = st.text_input("Five9 password", os.getenv("FIVE9_PASS", ""), type="password")
        five9_folder = st.text_input("Report folder (optional)", "")
        five9_report = st.text_input("Report name", "")
        today        = dt.date.today()
        d_from       = st.date_input("From date", today)
        d_to         = st.date_input("To date", today)

    # ---------------- Load main df ----------------
    df, source_meta = None, {}
    if source_type == "Public CSV URL":
        df, source_meta = fetch_csv_url(main_url)
        if df is None: st.error(f"URL load failed: {source_meta.get('error','')}"); st.stop()
    elif source_type == "Local folder (latest *.csv)":
        df, source_meta = load_latest_local_csv(main_folder, main_glob)
        if df is None: st.error(f"Local load failed: {source_meta.get('error','')}"); st.stop()
    elif source_type == "SFTP (Five9 Scheduled)":
        if not HAS_SFTP:
            st.error("Please add 'paramiko' to requirements.txt."); st.stop()
        df, source_meta = load_latest_sftp_csv(sftp_host, sftp_port, sftp_user, sftp_pass, sftp_dir, sftp_pat)
        if df is None: st.error(f"SFTP load failed: {source_meta.get('error','')}"); st.stop()
    elif source_type == "Five9 API (SOAP)":
        if not HAS_ZEEP:
            st.error("Please add 'zeep' to requirements.txt."); st.stop()
        if not (five9_wsdl and five9_user and five9_pass and five9_report):
            st.error("Provide WSDL, username, password, and report name."); st.stop()
        try:
            df, source_meta = five9_run_report_to_df(five9_wsdl, five9_user, five9_pass,
                                                     five9_report, five9_folder or None, d_from, d_to)
        except Exception as e:
            st.error(f"Five9 fetch failed: {e}"); st.stop()
    else:
        uploaded = st.file_uploader("Main report (CSV/XLSX/XLS)", type=["csv", "xlsx", "xls"], key="main")
        if uploaded is None:
            st.info("Upload the main CSV/Excel file, or choose another source.")
            st.stop()
        df = read_any(uploaded); source_meta = {"source": "uploaded file"}

    if df is None or df.empty:
        st.warning("The main report appears to be empty."); st.stop()

    # Alias "Abandoned (%rec)" -> "Abandon %"
    for c in list(df.columns):
        if norm(c) == norm("Abandoned (%rec)"):
            df.rename(columns={c: "Abandon %"}, inplace=True)

    st.caption(f"Loaded main report from: **{source_meta.get('source','(unknown)')}**")
    st.subheader("Preview — Main Report (first 20 rows)")
    st.dataframe(df.head(20), use_container_width=True)

    # ---------------- Column Mapping — Main ----------------
    st.subheader("Column Mapping — Main Report")

    SKILL_SYNS  = ["skill", "skill name", "skill group", "group", "queue", "split", "team", "program", "department", "dept", "category", "line of business", "lob"]
    CALLS_SYNS  = ["calls", "total calls", "calls offered", "offered", "inbound calls", "in calls", "total contacts", "contacts", "total interactions", "volume"]
    AGENTS_SYNS = ["agents staffed", "agents", "agent count", "staffed agents", "distinct agents", "distinct agent count", "unique agents", "logged in agents", "logged-in agents", "logged in", "agents (distinct)", "agents (unique)"]
    AHT_SYNS    = ["aht", "average handle time", "avg handle time", "avg handling time", "avg handle", "average handling time", "aht (s)", "aht (sec)", "talk+hold+acw", "handle time", "a.h.t", "avg hdl time", "avg handle-time"]
    ABAND_CNT_SYNS = ["abandoned count", "abandoned", "abandon count", "aband count", "abandoned calls", "aband qty", "aband num", "aband total"]
    ABAND_PCT_SYNS = ["abandon %", "abandoned (%rec)", "abandoned percent", "abandoned %", "abandonment rate", "abandon rate", "aband %", "aband pct", "abandonment %", "abandonment pct", "abn %", "abn pct"]

    skill_guess     = find_col(df, SKILL_SYNS)
    calls_guess     = find_col(df, CALLS_SYNS)
    agents_guess    = find_col(df, AGENTS_SYNS)
    aht_guess       = find_col(df, AHT_SYNS)
    aband_cnt_guess = find_col(df, ABAND_CNT_SYNS)
    aband_pct_guess = find_col(df, ABAND_PCT_SYNS)

    cols = list(df.columns)
    skill_col  = st.selectbox("Skill / Group column", cols, index=idx_or_default(cols, skill_guess or cols[0]))
    calls_col  = st.selectbox("Calls column",        cols, index=idx_or_default(cols, calls_guess or cols[0]))
    agents_col = st.selectbox("Agents Staffed column (per-skill)", cols, index=idx_or_default(cols, agents_guess or cols[0]))
    aht_col    = st.selectbox("AHT column", cols, index=idx_or_default(cols, aht_guess or cols[0]))
    abandoned_pct_col = st.selectbox("Abandon % column (optional)", ["<none>"] + cols,
                                     index=idx_or_default(["<none>"]+cols, aband_pct_guess if aband_pct_guess else "<none>"))
    abandoned_count_col = st.selectbox("Abandoned (count) column (optional, used if % is missing)", ["<none>"] + cols,
                                       index=idx_or_default(["<none>"]+cols, aband_cnt_guess if aband_cnt_guess else "<none>"))

    # Skills list (Fortress → PM Connect)
    default_skills = ["B2B Member Success", "B2B Success Activation", "B2B Success Info", "B2B Success Tech Support",
                      "MS Activation", "MS Info", "MS Loyalty", "MS Tech Support", "PM Connect"]
    skills_list = st.text_area("Skills of interest (one per line)", value="\n".join(default_skills))
    raw_skills = [s.strip() for s in skills_list.splitlines() if s.strip()]
    skills_wanted = []
    for s in raw_skills:
        if s.lower() == "fortress": s = "PM Connect"
        if s not in skills_wanted: skills_wanted.append(s)

    # ---------------- Secondary report (Agents total) ----------------
    st.sidebar.header("Second Report (Agents total) — Data Source")
    second_source_type = st.sidebar.radio(
        "Choose source",
        ["Manual upload", "Public CSV URL", "Local folder (latest *.csv)", "SFTP (Five9 Scheduled)", "Five9 API (SOAP)"],
        index=0,
    )

    second_df, second_meta = None, {}
    if second_source_type == "Public CSV URL":
        url2 = st.sidebar.text_input("2nd CSV URL", os.getenv("SECOND_CSV_URL", ""))
        if url2: second_df, second_meta = fetch_csv_url(url2)
    elif second_source_type == "Local folder (latest *.csv)":
        fold2 = st.sidebar.text_input("2nd local folder", "./data2")
        pat2  = st.sidebar.text_input("2nd filename pattern", "*.csv")
        second_df, second_meta = load_latest_local_csv(fold2, pat2)
    elif second_source_type == "SFTP (Five9 Scheduled)":
        with st.sidebar.expander("SFTP settings (2nd report)", expanded=False):
            sftp2_host = st.text_input("Host (2nd)", "")
            sftp2_port = st.number_input("Port (2nd)", 1, 65535, 22)
            sftp2_user = st.text_input("Username (2nd)", "")
            sftp2_pass = st.text_input("Password (2nd)", "", type="password")
            sftp2_dir  = st.text_input("Remote directory (2nd)", "/")
            sftp2_pat  = st.text_input("Filename pattern (2nd)", "*.csv")
        if HAS_SFTP and sftp2_host and sftp2_user:
            second_df, second_meta = load_latest_sftp_csv(sftp2_host, sftp2_port, sftp2_user, sftp2_pass, sftp2_dir, sftp2_pat)
    elif second_source_type == "Five9 API (SOAP)":
        with st.sidebar.expander("Five9 settings (2nd)", expanded=False):
            f_wsdl = st.text_input("WSDL URL (2nd)", value=os.getenv("FIVE9_WSDL", "https://api.five9.com/wsadmin/v12/AdminWebService?wsdl"))
            f_user = st.text_input("Username (2nd)", value=os.getenv("FIVE9_USER", ""))
            f_pass = st.text_input("Password (2nd)", value=os.getenv("FIVE9_PASS", ""), type="password")
            f_folder = st.text_input("Report folder (2nd) optional", value="")
            f_report = st.text_input("Report name (2nd)", value="")
            f_from   = st.date_input("From (2nd)", dt.date.today())
            f_to     = st.date_input("To (2nd)",   dt.date.today())
        if f_report and HAS_ZEEP:
            try:
                second_df, second_meta = five9_run_report_to_df(f_wsdl, f_user, f_pass, f_report, f_folder or None, f_from, f_to)
            except Exception as e:
                st.sidebar.error(f"Five9 2nd fetch failed: {e}")
    else:
        uploaded2 = st.file_uploader("Second report (CSV/XLSX/XLS) — overall totals / no skill filter (optional)", type=["csv", "xlsx", "xls"], key="second")
        if uploaded2 is not None:
            second_df = read_any(uploaded2)

    if second_df is not None and not second_df.empty:
        for c in list(second_df.columns):
            if norm(c) == norm("Abandoned (%rec)"):
                second_df.rename(columns={c: "Abandon %"}, inplace=True)
        st.caption(f"Loaded 2nd report from: **{second_meta.get('source','uploaded file')}**")
        st.dataframe(second_df.head(10), use_container_width=True)

    # ---------------- Calculations ----------------
    for c in [skill_col, calls_col, agents_col, aht_col]:
        if c not in df.columns:
            st.error(f"Selected column not found: {c}")
            st.stop()

    df[skill_col] = df[skill_col].astype(str).str.strip()
    df.loc[df[skill_col].str.lower() == "fortress", skill_col] = "PM Connect"

    calls_num  = pd.to_numeric(df[calls_col],  errors="coerce").fillna(0)
    agents_num = pd.to_numeric(df[agents_col], errors="coerce").fillna(0)

    rates = None
    if abandoned_pct_col != "<none>" and abandoned_pct_col in df.columns:
        rates = to_percent(df[abandoned_pct_col])

    if rates is None and abandoned_count_col != "<none>" and abandoned_count_col in df.columns:
        aband_num = pd.to_numeric(df[abandoned_count_col], errors="coerce")
        with pd.option_context('mode.use_inf_as_na', True):
            rates = (aband_num / calls_num.replace(0, pd.NA)) * 100

    total_calls = int(calls_num.sum())

    total_agents = int(agents_num.sum())
    agents_label = "Agents Staffed (sum of per-skill)"
    if second_df is not None and not second_df.empty:
        agents2_guess = find_col(second_df, AGENTS_SYNS) or next((c for c in second_df.columns if "agent" in c.lower()), None)
        if agents2_guess:
            agents2_num = pd.to_numeric(second_df[agents2_guess], errors="coerce").fillna(0)
            total_agents = int(agents2_num.sum())
            agents_label = "Agents Staffed (from 2nd report)"

    if (abandoned_count_col != "<none>" and abandoned_count_col in df.columns and total_calls > 0):
        aband_num_total = pd.to_numeric(df[abandoned_count_col], errors="coerce").fillna(0).sum()
        total_abandon_pct = (aband_num_total / total_calls) * 100
    elif rates is not None and total_calls > 0:
        total_abandon_pct = ((rates.fillna(0) / 100.0) * calls_num).sum() / total_calls * 100
    else:
        total_abandon_pct = None

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

    # ---------------- Filled Report (Markdown) ----------------
    md = io.StringIO()
    def writeln(s=""): md.write(s + "\n")
    writeln("## Autofilled Metrics (Core)\n")
    writeln(f"### 3. Total Calls\n**{total_calls}**\n")
    writeln(f"### 4. {agents_label}\n**{total_agents}**\n")
    writeln("### 6. Abandon %")
    writeln(f"**{(str(round(total_abandon_pct, 2)) + '%') if total_abandon_pct is not None else 'N/A'}**\n")
    writeln("### 7. AHT (By Group)")
    for sk in skills_wanted:
        mask = by_skill_core["SKILL"].str.lower() == sk.lower()
        val = by_skill_core.loc[mask, "AHT"]
        writeln(f"- **{sk}:** {val.iloc[0] if len(val) else 'Not found in this report'}")
    writeln("\n### 8. Abandon % (By Group)")
    for sk in skills_wanted:
        mask = by_skill_core["SKILL"].str.lower() == sk.lower()
        val = by_skill_core.loc[mask, "Abandon %"]
        writeln(f"- **{sk}:** {val.iloc[0] if len(val) else 'Not found in this report'}")
    report_md = md.getvalue()

    # ---------------- UI ----------------
    st.subheader("Filled Report (Core)")
    st.markdown(report_md)

    st.subheader("Copy & Paste")
    tabs = st.tabs(["📋 Report (Markdown)", "📋 KPIs (CSV)", "📋 By-Skill Core Table (CSV)", "📋 By-Skill Core Table (TSV)", "📋 Preview (first 20 rows CSV)"])
    with tabs[0]:
        st.code(report_md, language="markdown")
    with tabs[1]:
        kpi_df = pd.DataFrame([{
            "Total Calls": total_calls,
            "Agents Staffed": total_agents,
            "Total Abandon %": (round(total_abandon_pct, 2) if total_abandon_pct is not None else None)
        }])
        disp = kpi_df.copy()
        if disp.loc[0, "Total Abandon %"] is not None:
            disp.loc[0, "Total Abandon %"] = f"{disp.loc[0, 'Total Abandon %']:.2f}%"
        st.dataframe(disp, use_container_width=True)
        st.code(kpi_df.to_csv(index=False), language="text")
    with tabs[2]:
        st.dataframe(by_skill_core, use_container_width=True)
        st.code(by_skill_core.to_csv(index=False), language="text")
    with tabs[3]:
        st.dataframe(by_skill_core, use_container_width=True)
        st.code(by_skill_core.to_csv(index=False, sep="\t"), language="text")
    with tabs[4]:
        prev_csv = df.head(20).to_csv(index=False)
        st.dataframe(df.head(20), use_container_width=True)
        st.code(prev_csv, language="text")

    st.subheader("Downloads")
    st.download_button("⬇️ Download report (Markdown)", data=report_md.encode("utf-8"),
                       file_name="filled_report_core.md", mime="text/markdown")

    st.subheader("By-Skill Table — Core Fields")
    st.dataframe(by_skill_core, use_container_width=True)

    # Word export
    if HAS_DOCX:
        def build_docx(md_text):
            doc = Document(); doc.add_heading("Autofilled Metrics (Core)", level=1)
            for line in md_text.splitlines():
                if line.startswith("### "): doc.add_heading(line.replace("### ", ""), level=2)
                elif line.startswith("## "): continue
                else:
                    if line.strip(): doc.add_paragraph(line)
            bio = io.BytesIO(); doc.save(bio); bio.seek(0); return bio.getvalue()
        st.download_button("⬇️ Download report (Word .docx)", data=build_docx(report_md),
                           file_name="filled_report_core.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    else:
        st.info("python-docx not installed; Word export disabled.")

    # PDF export
    if HAS_PDF:
        def build_pdf(md_text):
            bio = io.BytesIO(); c = canvas.Canvas(bio, pagesize=letter)
            width, height = letter; L = 0.75 * inch; T = 0.75 * inch; B = 0.75 * inch
            y = height - T; c.setFont("Times-Bold", 14); c.drawString(L, y, "Autofilled Metrics (Core)"); y -= 0.3*inch
            c.setFont("Times-Roman", 11); import textwrap as _tw
            for line in md_text.splitlines():
                if not line.strip(): y -= 0.18*inch
                else:
                    for w in _tw.wrap(line, width=95): c.drawString(L, y, w); y -= 0.18*inch
                if y < B: c.showPage(); y = height - T; c.setFont("Times-Roman", 11)
            c.showPage(); c.save(); bio.seek(0); return bio.getvalue()
        st.download_button("⬇️ Download report (PDF)", data=build_pdf(report_md),
                           file_name="filled_report_core.pdf", mime="application/pdf")
    else:
        st.info("reportlab not installed; PDF export disabled.")

# ---------- Boot ----------
if require_login():
    run_app()
