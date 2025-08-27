# app.py — Patriot-themed site generator (Streamlit)
# Features:
# - Upload or URL for logo
# - Brand color pickers (defaults derived from Patriot Mobile styling)
# - Live preview (inline CSS + inline logo)
# - Download production files as a zip: index.html, assets/css/theme.css, assets/img/logo.*
# Only needs: streamlit (and requests if you use a logo URL)

import io
import os
import base64
import zipfile
from urllib.parse import urlparse
import streamlit as st

# requests is optional; only used if you provide a logo URL
try:
    import requests
    HAS_REQUESTS = True
except Exception:
    HAS_REQUESTS = False

st.set_page_config(page_title="Patriot-Themed Site Builder", page_icon="🛠️", layout="wide")

# ----------------------------- Defaults ------------------------------------
DEFAULTS = {
    "pm_red":  "#C8102E",
    "pm_navy": "#0B2D52",
    "ink":     "#0A0D12",
    "snow":    "#FFFFFF",
    "g90":     "#1F2430",
    "g70":     "#3A4152",
    "g20":     "#D7DBE2",
    "heading": "Nationwide coverage, simple plans, modern experience.",
    "tag":     "Mobilizing Freedom",
    "sub":     "Bring your own device or pick from the latest phones. Switch in minutes and keep your number.",
    "cta1":    "Get Started",
    "cta2":    "See Plans",
    "nav":     ["Plans", "Devices", "Coverage", "Contact"],
    "dark":    False,
}

# --------------------------- UI: Sidebar -----------------------------------
st.title("🎨 Patriot-Themed Site Builder")
st.caption("Upload your logo, tune colors & copy, preview instantly, and download a ready-to-host site.")

with st.sidebar:
    st.header("Brand")
    colA, colB = st.columns(2)
    pm_red  = colA.color_picker("Patriot Red", DEFAULTS["pm_red"])
    pm_navy = colB.color_picker("Patriot Navy", DEFAULTS["pm_navy"])
    colC, colD = st.columns(2)
    g70 = colC.color_picker("Cool Gray 70", DEFAULTS["g70"])
    g20 = colD.color_picker("Cool Gray 20", DEFAULTS["g20"])
    dark_mode = st.toggle("Dark mode preview", DEFAULTS["dark"])

    st.markdown("---")
    st.subheader("Logo")
    logo_file = st.file_uploader("Upload logo (.svg/.png/.jpg)", type=["svg", "png", "jpg", "jpeg"])
    logo_url  = st.text_input("…or Logo URL (https://…)", value="", placeholder="https://example.com/logo.svg")
    st.caption("Tip: Use an SVG for best results. If both are provided, the upload wins.")

    st.markdown("---")
    st.subheader("Hero Copy")
    tag = st.text_input("Badge", DEFAULTS["tag"])
    heading = st.text_area("Headline", DEFAULTS["heading"])
    sub = st.text_area("Subheading", DEFAULTS["sub"])
    colE, colF = st.columns(2)
    cta1 = colE.text_input("Primary CTA", DEFAULTS["cta1"])
    cta2 = colF.text_input("Secondary CTA", DEFAULTS["cta2"])

    st.markdown("---")
    st.subheader("Navigation")
    nav_raw = st.text_area("Links (one per line)", "\n".join(DEFAULTS["nav"]))

# -------------------------- Helpers ----------------------------------------
def guess_ext_from_bytes(b: bytes) -> str:
    # simple magic checks
    if b[:4] == b"\x89PNG": return "png"
    if b[:3] == b"\xFF\xD8\xFF": return "jpg"
    # svg often starts with XML or <svg
    head = b[:200].lstrip()
    if head.startswith(b"<svg") or head.startswith(b"<?xml"):
        return "svg"
    return "bin"

def fetch_logo_bytes(logo_file, logo_url: str) -> tuple[bytes|None, str|None]:
    """Return (bytes, ext) or (None, None). Uploaded file wins; else try URL (if requests available)."""
    if logo_file is not None:
        data = logo_file.read()
        ext = os.path.splitext(logo_file.name)[1].lower().strip(".") or guess_ext_from_bytes(data)
        return data, ext
    if logo_url.strip():
        if not HAS_REQUESTS:
            st.warning("To use a logo URL, please add 'requests' to requirements.txt.")
            return None, None
        try:
            r = requests.get(logo_url, timeout=20)
            r.raise_for_status()
            data = r.content
            # try ext from URL path first
            path = urlparse(logo_url).path
            ext = os.path.splitext(path)[1].lower().strip(".") or guess_ext_from_bytes(data)
            return data, ext or "bin"
        except Exception as e:
            st.error(f"Failed to fetch logo from URL: {e}")
            return None, None
    return None, None

def to_data_uri(b: bytes, ext: str) -> str:
    if ext == "svg":
        try:
            # Prefer utf-8 text for crisp SVG rendering (no base64 needed)
            svg_text = b.decode("utf-8", errors="ignore")
            return f"data:image/svg+xml;utf8,{svg_text}"
        except Exception:
            pass
    mime = "image/svg+xml" if ext == "svg" else ("image/png" if ext == "png" else "image/jpeg")
    b64 = base64.b64encode(b).decode("ascii")
    return f"data:{mime};base64,{b64}"

def build_theme_css(pm_red, pm_navy, g70, g20):
    # Core design system CSS (derived from Patriot styling)
    return f"""
/* === Theme (Patriot-inspired) ========================================== */
:root {{
  --pm-red: {pm_red};
  --pm-navy: {pm_navy};
  --pm-ink: #0A0D12;
  --pm-snow: #FFFFFF;
  --pm-gray-90: #1F2430;
  --pm-gray-70: {g70};
  --pm-gray-20: {g20};

  --bg: var(--pm-snow);
  --text: #0B1020;
  --muted: #5B6373;
  --surface: #F7F8FA;
  --card: #FFFFFF;
  --border: var(--pm-gray-20);
  --brand: var(--pm-red);
  --brand-ink: var(--pm-navy);
  --focus: var(--pm-red);
}}
:root[data-theme="dark"] {{
  --bg: #0B0E15;
  --text: #E9EDF5;
  --muted: #A9B1C1;
  --surface: #10141C;
  --card: #121723;
  --border: #2B3241;
}}
* {{ box-sizing: border-box; }}
html, body {{ height: 100%; }}
body {{
  margin: 0;
  background: var(--bg);
  color: var(--text);
  font: 400 16px/1.6 Inter, system-ui, -apple-system, Segoe UI, Roboto, "Helvetica Neue", Arial, "Apple Color Emoji", "Segoe UI Emoji";
  text-rendering: optimizeLegibility;
}}
img {{ max-width: 100%; display: block; }}
a {{ color: var(--brand); text-decoration: none; }}
a:hover {{ text-decoration: underline; }}

.container {{ width: min(1120px, 92vw); margin: 0 auto; }}
.grid {{ display: grid; gap: 24px; }}
.grid-2 {{ grid-template-columns: repeat(2, minmax(0,1fr)); }}
.grid-3 {{ grid-template-columns: repeat(3, minmax(0,1fr)); }}
@media (max-width: 900px){{ .grid-2, .grid-3 {{ grid-template-columns: 1fr; }} }}

.btn {{
  --btn-bg: var(--brand); --btn-fg: var(--pm-snow); --btn-bd: transparent;
  display:inline-flex; align-items:center; gap:10px;
  padding: 12px 18px; border-radius: 12px; border:1px solid var(--btn-bd);
  background: var(--btn-bg); color: var(--btn-fg); font-weight:600;
  box-shadow: 0 1px 0 rgba(0,0,0,.05), 0 8px 16px rgba(200,16,46,.12);
  transition: transform .16s ease, box-shadow .2s ease, background .2s ease;
}}
.btn:hover {{ transform: translateY(-1px); box-shadow: 0 2px 0 rgba(0,0,0,.06), 0 12px 20px rgba(200,16,46,.18); }}
.btn:focus-visible {{ outline: 3px solid var(--focus); outline-offset: 2px; }}
.btn.secondary {{ --btn-bg: var(--pm-navy); box-shadow: 0 8px 16px rgba(11,45,82,.12); }}
.btn.ghost {{ --btn-bg: transparent; --btn-fg: var(--text); --btn-bd: var(--border); box-shadow: none; }}

.badge {{
  display:inline-block; padding:6px 10px; border-radius: 999px;
  border:1px solid var(--border); background: var(--surface); color: var(--brand-ink);
  font-weight:600; font-size: 12px; letter-spacing: .3px;
}}

.header {{
  position: sticky; top:0; z-index:50;
  backdrop-filter:saturate(180%) blur(6px);
  background: color-mix(in srgb, var(--bg) 86%, transparent);
  border-bottom: 1px solid var(--border);
}}
.nav {{ display:flex; align-items:center; justify-content:space-between; padding: 14px 0; }}
.brand {{ display:flex; align-items:center; gap:12px; }}
.brand img {{ height: 34px; width:auto; }}
.nav a {{ color: var(--text); font-weight:600; padding:8px 12px; border-radius:8px; }}
.nav a:hover {{ background: var(--surface); }}

.hero {{
  padding: 64px 0;
  background:
    radial-gradient(1200px 600px at 10% -10%, color-mix(in srgb, var(--brand) 30%, transparent), transparent 60%),
    linear-gradient(180deg, color-mix(in srgb, var(--brand-ink) 8%, transparent), transparent 40%);
}}
.hero h1 {{
  font: 700 clamp(28px, 4vw, 48px)/1.1 Poppins, Inter, system-ui;
  letter-spacing:-.02em; margin: 0 0 16px;
}}
.hero p {{ max-width: 60ch; margin:0 0 24px; color: var(--muted); }}
.hero .cta {{ display:flex; gap: 12px; flex-wrap: wrap; }}

.card {{
  background: var(--card); border:1px solid var(--border); border-radius: 16px;
  padding: 20px; transition: box-shadow .2s, transform .16s, border-color .2s;
}}
.card:hover {{ transform: translateY(-2px); box-shadow: 0 10px 28px rgba(0,0,0,.08); border-color: color-mix(in srgb, var(--border), var(--brand) 25%); }}
.card h3 {{ margin:0 0 6px; font:700 20px/1.3 Poppins, Inter, system-ui; }}
.card p {{ margin:0; color: var(--muted); }}

.footer {{ margin-top: 64px; padding: 32px 0; border-top: 1px solid var(--border); color: var(--muted); font-size: 14px; }}
input[type="text"], input[type="email"], input[type="search"], select {{
  width:100%; padding: 12px 14px; border-radius: 12px; border:1px solid var(--border);
  background: var(--card); color: var(--text); transition: border-color .2s, box-shadow .2s;
}}
input:focus-visible, select:focus-visible {{
  outline: none; border-color: color-mix(in srgb, var(--brand) 30%, var(--border));
  box-shadow: 0 0 0 3px color-mix(in srgb, var(--brand) 20%, transparent);
}}
.sr-only {{ position:absolute; width:1px; height:1px; padding:0; margin:-1px; overflow:hidden; clip:rect(0,0,0,0); white-space:nowrap; border:0; }}
"""

def build_index_html(dark_mode: bool, nav_links, heading, sub, tag, cta1, cta2):
    nav_html = "\n        ".join([f'<a href="#{l.lower()}">{l}</a>' for l in nav_links if l.strip()])
    dark_attr = ' data-theme="dark"' if dark_mode else ''
    return f"""<!doctype html>
<html lang="en"{dark_attr}>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Patriot-themed Site</title>
  <!-- Fonts -->
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=Poppins:wght@600;700&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="assets/css/theme.css" />
</head>
<body>
  <header class="header">
    <div class="container nav">
      <a class="brand" href="/"><img src="assets/img/logo.svg" alt="Logo" /><span class="sr-only">Home</span></a>
      <nav>
        {nav_html}
        <a href="#contact" class="btn ghost">Contact</a>
      </nav>
    </div>
  </header>

  <main>
    <section class="hero">
      <div class="container">
        <span class="badge">{tag}</span>
        <h1>{heading}</h1>
        <p>{sub}</p>
        <div class="cta">
          <a class="btn" href="#get-started">{cta1}</a>
          <a class="btn secondary" href="#plans">{cta2}</a>
        </div>
      </div>
    </section>

    <section class="container" id="highlights" style="padding: 48px 0;">
      <div class="grid grid-3">
        <article class="card"><h3>Nationwide Coverage</h3><p>Reliable service across major U.S. networks.</p></article>
        <article class="card"><h3>Keep Your Number</h3><p>Seamless port-in with helpful support.</p></article>
        <article class="card"><h3>Transparent Pricing</h3><p>No hidden fees. Change plans anytime.</p></article>
      </div>
    </section>

    <section class="container" id="signup" style="padding: 24px 0 64px;">
      <div class="grid grid-2">
        <div class="card">
          <h3>Check Coverage</h3>
          <p>Enter your ZIP to see signal strength in your area.</p>
          <form onsubmit="return false">
            <label class="sr-only" for="zip">ZIP code</label>
            <input id="zip" type="text" placeholder="ZIP code" inputmode="numeric" />
            <div style="margin-top:12px;"><button class="btn">Check</button></div>
          </form>
        </div>
        <div class="card">
          <h3>Join the Newsletter</h3>
          <p>Get updates on plans, promos, and devices.</p>
          <form onsubmit="return false">
            <label class="sr-only" for="email">Email</label>
            <input id="email" type="email" placeholder="you@example.com" />
            <div style="margin-top:12px;"><button class="btn secondary">Subscribe</button></div>
          </form>
        </div>
      </div>
    </section>
  </main>

  <footer class="footer"><div class="container">© <span id="y"></span> Your Company. All rights reserved.</div></footer>
  <script>document.getElementById("y").textContent = new Date().getFullYear();</script>
</body>
</html>"""

def build_preview_html_inline(css: str, logo_data_uri: str, dark_mode: bool,
                              nav_links, heading, sub, tag, cta1, cta2):
    nav_html = "\n              ".join([f'<a href="#{l.lower()}">{l}</a>' for l in nav_links if l.strip()])
    dark_attr = ' data-theme="dark"' if dark_mode else ''
    return f"""<!doctype html>
<html lang="en"{dark_attr}>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Preview</title>
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=Poppins:wght@600;700&display=swap" rel="stylesheet">
  <style>{css}</style>
</head>
<body>
  <header class="header">
    <div class="container nav">
      <a class="brand" href="/"><img src="{logo_data_uri}" alt="Logo" /><span class="sr-only">Home</span></a>
      <nav>
        {nav_html}
        <a href="#contact" class="btn ghost">Contact</a>
      </nav>
    </div>
  </header>

  <main>
    <section class="hero">
      <div class="container">
        <span class="badge">{tag}</span>
        <h1>{heading}</h1>
        <p>{sub}</p>
        <div class="cta">
          <a class="btn" href="#get-started">{cta1}</a>
          <a class="btn secondary" href="#plans">{cta2}</a>
        </div>
      </div>
    </section>

    <section class="container" id="highlights" style="padding: 48px 0;">
      <div class="grid grid-3">
        <article class="card"><h3>Nationwide Coverage</h3><p>Reliable service across major U.S. networks.</p></article>
        <article class="card"><h3>Keep Your Number</h3><p>Seamless port-in with helpful support.</p></article>
        <article class="card"><h3>Transparent Pricing</h3><p>No hidden fees. Change plans anytime.</p></article>
      </div>
    </section>
  </main>

  <footer class="footer"><div class="container">© <span id="y"></span> Your Company. All rights reserved.</div></footer>
  <script>document.getElementById("y").textContent = new Date().getFullYear();</script>
</body>
</html>"""

def make_zip(css_text: str, html_text: str, logo_bytes: bytes, logo_ext: str) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_DEFLATED) as z:
        # index.html
        z.writestr("index.html", html_text)
        # assets
        z.writestr("assets/css/theme.css", css_text)
        # Normalize logo to .svg if svg, else keep ext
        logo_name = "assets/img/logo." + (logo_ext if logo_ext else "svg")
        z.writestr(logo_name, logo_bytes)
    buf.seek(0)
    return buf.read()

# --------------------------- Build content ----------------------------------
nav_links = [n.strip() for n in nav_raw.splitlines() if n.strip()]
css_text = build_theme_css(pm_red, pm_navy, g70, g20)

# Logo handling
logo_bytes, logo_ext = fetch_logo_bytes(logo_file, logo_url)
if logo_bytes is None:
    # simple placeholder SVG logo
    placeholder_svg = f'''<svg xmlns="http://www.w3.org/2000/svg" width="240" height="36" viewBox="0 0 480 72">
  <rect width="480" height="72" rx="12" fill="{pm_navy}"/>
  <text x="50%" y="50%" dominant-baseline="middle" text-anchor="middle"
        font-family="Poppins, Arial, sans-serif" font-weight="700" font-size="28" fill="{DEFAULTS['snow']}">
    YOUR LOGO
  </text>
</svg>'''
    logo_bytes = placeholder_svg.encode("utf-8")
    logo_ext = "svg"
logo_data_uri = to_data_uri(logo_bytes, logo_ext)

# HTML for preview (inline CSS + inline logo)
preview_html = build_preview_html_inline(css_text, logo_data_uri, dark_mode, nav_links, heading, sub, tag, cta1, cta2)

# HTML for download (external CSS + file logo)
# We always save the logo as "assets/img/logo.<ext>" and reference it in index.html
download_index_html = build_index_html(dark_mode, nav_links, heading, sub, tag, cta1, cta2)

# --------------------------- Preview & Download -----------------------------
st.subheader("Preview")
st.caption("This is an inline preview. Download the zip to get clean files for hosting.")
st.components.v1.html(preview_html, height=750, scrolling=True)

st.markdown("### Download your site")
zip_bytes = make_zip(css_text, download_index_html, logo_bytes, logo_ext)
st.download_button(
    "⬇️ Download website (.zip)",
    data=zip_bytes,
    file_name="patriot_theme_site.zip",
    mime="application/zip"
)

st.markdown("---")
st.info(
    "Trademark note: The Patriot Mobile name and logo are trademarks of their respective owner. "
    "Ensure you have permission to use any third-party marks and follow brand guidelines."
)
