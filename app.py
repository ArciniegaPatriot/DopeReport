# app.py — Minimal "Metrics Report" shell (Patriot-inspired colors + logo)
# - Title is white
# - Header bar uses Patriot Mobile navy
# - Upload a logo OR use a logo URL (upload wins)
# - Paste your report UI where indicated

import base64
import streamlit as st

# Optional: only needed if you’ll use a logo URL
try:
    import requests
    HAS_REQUESTS = True
except Exception:
    HAS_REQUESTS = False

# ---------------- Brand Colors ----------------
PM_RED   = "#C8102E"
PM_NAVY  = "#0B2D52"
PM_WHITE = "#FFFFFF"
PM_GRAY  = "#D7DBE2"

st.set_page_config(page_title="Metrics Report", layout="wide")

# ---------------- Sidebar: Logo ----------------
with st.sidebar:
    st.header("Brand & Logo")
    logo_file = st.file_uploader(
        "Upload logo (.svg/.png/.jpg)", type=["svg", "png", "jpg", "jpeg"], key="logo_upload"
    )
    logo_url = st.text_input(
        "…or paste a logo URL", placeholder="https://example.com/logo.svg", key="logo_url"
    )
    st.caption("Tip: SVG preferred. If both are provided, the upload is used.")

# ---------------- Helpers ----------------
def _guess_ext(b: bytes) -> str:
    if b[:4] == b"\x89PNG":
        return "png"
    if b[:3] == b"\xFF\xD8\xFF":
        return "jpg"
    head = b[:200].lstrip()
    if head.startswith(b"<svg") or head.startswith(b"<?xml"):
        return "svg"
    return "bin"

def _fetch_logo_bytes():
    # Upload wins
    if logo_file is not None:
        data = logo_file.read()
        return data, (_guess_ext(data) or "svg")
    # Else try URL
    if logo_url.strip():
        if not HAS_REQUESTS:
            st.sidebar.warning("Add 'requests' to requirements.txt to use a logo URL.")
            return None, None
        try:
            r = requests.get(logo_url.strip(), timeout=20)
            r.raise_for_status()
            data = r.content
            return data, _guess_ext(data)
        except Exception as e:
            st.sidebar.error(f"Logo URL failed: {e}")
            return None, None
    return None, None

def _to_data_uri(b: bytes, ext: str) -> str:
    if ext == "svg":
        # Prefer UTF-8 text form for crisp rendering
        try:
            txt = b.decode("utf-8", errors="ignore")
            return f"data:image/svg+xml;utf8,{txt}"
        except Exception:
            pass
    mime = "image/svg+xml" if ext == "svg" else ("image/png" if ext == "png" else "image/jpeg")
    return f"data:{mime};base64,{base64.b64encode(b).decode('ascii')}"

# Fetch / fallback logo
logo_bytes, logo_ext = _fetch_logo_bytes()
if logo_bytes is None:
    placeholder_svg = f'''<svg xmlns="http://www.w3.org/2000/svg" width="220" height="40" viewBox="0 0 440 80">
  <rect width="440" height="80" rx="12" fill="{PM_NAVY}"/>
  <text x="50%" y="52%" dominant-baseline="middle" text-anchor="middle"
        font-family="Poppins, Arial, sans-serif" font-weight="700" font-size="26" fill="{PM_WHITE}">
    YOUR LOGO
  </text>
</svg>'''
    logo_bytes = placeholder_svg.encode("utf-8")
    logo_ext = "svg"

logo_data_uri = _to_data_uri(logo_bytes, logo_ext or "svg")

# ---------------- Light CSS (balanced & safe) ----------------
st.markdown(
    f"""
<style>
/* App background */
html, body, .stApp {{
  background: #ffffff;
  color: #0B1020;
}}

/* Top header bar */
.pm-header {{
  position: sticky; top: 0; z-index: 10;
  background: {PM_NAVY};
  border-bottom: 1px solid {PM_GRAY};
  padding: 10px 0;
}}
.pm-wrap {{
  width: min(1120px, 92vw); margin: 0 auto;
  display: flex; align-items: center; gap: 16px;
}}
.pm-logo {{
  height: 36px; width: auto; display: block;
}}
.pm-title {{
  margin: 0; padding: 0;
  color: {PM_WHITE}; /* Title in white */
  font: 700 22px/1.2 Poppins, Inter, system-ui, -apple-system, Segoe UI, Roboto, Arial;
}}

/* Optional accents */
h2 {{
  border-bottom: 2px solid {PM_RED};
  padding-bottom: 4px;
}}
.stButton > button {{
  background: {PM_RED}; color: {PM_WHITE}; border: 1px solid transparent;
  border-radius: 10px; padding: 0.5rem 0.9rem; font-weight: 600;
}}
.stButton > button:hover {{ filter: brightness(0.95); }}
</style>
    """,
    unsafe_allow_html=True,
)

# ---------------- Header with Logo + Title ----------------
st.markdown(
    f"""
<div class="pm-header">
  <div class="pm-wrap">
    <img src="{logo_data_uri}" alt="Logo" class="pm-logo" />
    <h1 class="pm-title">Metrics Report</h1>
  </div>
</div>
""",
    unsafe_allow_html=True,
)

# ---------------- Your Report Area ----------------
st.write("")  # small spacer
st.markdown("#### Overview")

# >>> Paste your existing report UI below this line. <<<
# Example placeholders (safe to delete):
c1, c2, c3 = st.columns(3)
c1.metric("Total Calls", "—")
c2.metric("Agents Staffed", "—")
c3.metric("Abandon %", "—")

st.markdown("---")
st.caption("Page chrome uses Patriot-inspired colors (navy/white with red accents). Replace with official values if you have a brand guide.")
