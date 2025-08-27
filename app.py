# app.py — Minimal page chrome with Patriot Mobile-inspired colors + logo
# - Page title: "Metrics Report"
# - Simple color styling (red/navy/white), no heavy theme
# - Logo from upload OR URL (upload wins)
# - Drop your existing report UI where indicated

import io
import base64
from urllib.parse import urlparse
import streamlit as st

# Optional: only needed if you use a logo URL
try:
    import requests
    HAS_REQUESTS = True
except Exception:
    HAS_REQUESTS = False

st.set_page_config(page_title="Metrics Report", layout="wide")

# ---- Brand colors (inspired by Patriot Mobile) ----
PM_RED  = "#C8102E"
PM_NAVY = "#0B2D52"
PM_WHITE = "#FFFFFF"
PM_GRAY = "#D7DBE2"

# ---- Sidebar controls for logo ----
with st.sidebar:
    st.header("Brand & Logo")
    logo_file = st.file_uploader("Upload logo (.svg/.png/.jpg)", type=["svg", "png", "jpg", "jpeg"])
    logo_url = st.text_input("…or paste a logo URL", placeholder="https://example.com/logo.svg")
    st.caption("Tip: SVG preferred. If both provided, the upload is used.")

def _guess_ext(b: bytes) -> str:
    if b[:4] == b"\x89PNG": return "png"
    if b[:3] == b"\xFF\xD8\xFF": return "jpg"
    head = b[:200].lstrip()
    if head.startswith(b"<svg") or head.startswith(b"<?xml"): return "svg"
    return "bin"

def _fetch_logo_bytes():
    if logo_file is not None:
        data = logo_file.read()
        return data, (_guess_ext(data) or "svg")
    if logo_url.strip():
        if not HAS_REQUESTS:
            st.sidebar.warning("Add 'requests' to requirements.txt to use a logo URL.")
            return None, None
        try:
            r = requests.get(logo_url, timeout=20)
            r.raise_for_status()
            data = r.content
            return data, _guess_ext(data)
        except Exception as e:
            st.sidebar.error(f"Logo URL failed: {e}")
            return None, None
    return None, None

logo_bytes, logo_ext = _fetch_logo_bytes()

# Fallback placeholder SVG (simple wordmark box) if no logo provided
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

def _to_data_uri(b: bytes, ext: str) -> str:
    # Prefer text form for SVG
    if ext == "svg":
        try:
            txt = b.decode("utf-8", errors="ignore")
            return f"data:image/svg+xml;utf8,{txt}"
        except Exception:
            pass
    mime = "image/svg+xml" if ext == "svg" else ("image/png" if ext == "png" else "image/jpeg")
    b64 = base64.b64encode(b).decode("ascii")
    return f"data:{mime};base64,{b64}"

logo_data_uri = _to_data_uri(logo_bytes, logo_ext or "svg")

# ---- Light CSS to color the page (no full theme) ----
st.markdown(f"""
<style>
/* Page background + default text */
html, body, .main, .stApp {{
  background: #ffffff;
  color: #0B1020;
}}

/* Top header bar */
.pm-header {{
  position: sticky; top: 0; z-index: 10;
  background: linear-gradient(90deg, {PM_NAVY}, {PM_NAVY});
  border-bottom: 1px solid {PM_GRAY};
  padding: 10px 0;
}}
.pm-wrap {{
  width: min(1120px, 92vw); margin: 0 auto;
  display: flex; align-items: center; gap: 16px;
}}
.pm-logo {{
  height: 36px; width: auto; display:block;
}}
.pm-title {{
  margin: 0; padding: 0; color: {PM_WHITE};
  font: 700 22px/1.2 Poppins, Inter, system-ui, -apple-system, Segoe UI, Roboto, Arial;
}}
/* Accent underline for section titles */
h2, .stMarkdown h2 {{
  border-bottom: 2px solid {PM_RED}; padding-bottom: 4px;
}}
/* Buttons */
.stButton > button {{
  background: {PM_RED}; color: {PM_WHITE}; border: 1px solid transparent;
  border-radius: 10px; padding: 0.5rem 0.9rem; font-weight: 600;
}}
.stButton > button:hover {{
  filter: brightness(0.95);
}}
/* Checkboxes, radios focus rings */
input:focus-visible, select:focus-visible, textarea:focus-visible {{
  outline: 3px solid {PM_RED}33 !important;
}}
/* Tables border tone */
[data-testid="stTable"] table, .stDataFrame div[role="grid"] {{
  border-color: {PM_GRAY} !important;
}}
</style>
""", unsafe_allow_html=True)

# ---- Header bar with logo + title ----
st.markdown(f"""
<style>
/* ...other styles... */

.pm-title {{
  margin: 0; padding: 0;
  color: {PM_RED};  /* was {PM_WHITE} */
  font: 700 22px/1.2 Poppins, Inter, system-ui, -apple-system, Segoe UI, Roboto, Arial;
}}

/* ...other styles... */
</style>
""", unsafe_allow_html=True)

)

# ---- Page body (put your report below) ----
st.write("")  # small spacer

# Example “hero” title line (kept very minimal)
st.markdown("#### Overview")

# >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
# Your report UI goes here.
# For example, if you already have code that renders KPIs/tables/charts,
# paste it below this line. The header above will keep the brand color/Logo.
# >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

# Small demo block (safe to delete)
with st.container():
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Calls", "12,480", "+3.2%")
    c2.metric("Agents Staffed", "142", "+4")
    c3.metric("Abandon %", "2.14%", "-0.3 pts")
    st.caption("Replace these demo metrics with your real ones.")

st.markdown("---")
st.caption("Brand colors are inspired by Patriot Mobile (red/navy/white). Replace with official values if you have a brand guide.")
