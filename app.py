"""
Affle Awards Dashboard — Streamlit app backed by a Google Sheet.

Edit the Google Sheet -> the app picks up the change (auto every 5 min, or the
"Refresh data" button). No code changes needed for new awards, BUs or quarters.
"""
from __future__ import annotations

import base64
import io
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st

# ----------------------------------------------------------------------------
# Config
# ----------------------------------------------------------------------------
APP_DIR = Path(__file__).parent
LOCAL_FILE = APP_DIR / "awards_data_template.xlsx"
ASSETS = APP_DIR / "assets"
CACHE_TTL_SECONDS = 300  # how often the app re-reads the Google Sheet

st.set_page_config(page_title="Affle Awards Dashboard", page_icon="🏆", layout="wide")

BU_COLORS = {
    "appnext": "#0ea99b", "mdsp": "#ff7a4a", "mediasmart": "#7c6ff0",
    "newton": "#f2a53c", "revx": "#ef5da8", "ultra iris": "#3b8fd6",
}
EXTRA_COLORS = ["#5bb85c", "#8d6e63", "#26a0c9", "#c2185b", "#7cb342", "#455a64"]
STATUS_COLORS = {"Declared": "#0ea99b", "Pending": "#f2a53c", "Upcoming": "#8a8f98"}
px.defaults.color_discrete_sequence = ["#0ea99b", "#ff7a4a", "#7c6ff0", "#f2a53c", "#ef5da8", "#3b8fd6", "#5bb85c"]
LOGO_FILES = {
    "affle": "affle.png",
    "appnext": "appnext.png", "mdsp": "mdsp.png", "mediasmart": "mediasmart.png",
    "newton": "newton.webp", "revx": "revx.webp", "ultra": "ultra.png",
}

# Canonical column name -> accepted spellings in the sheet (case-insensitive)
COLUMN_ALIASES = {
    "Award": ["award", "award name", "awards name", "award program"],
    "BU": ["bu", "business unit", "platform"],
    "Status": ["status", "result status"],
    "Submissions": ["submissions", "entries", "no. of submissions"],
    "Won": ["won", "awards won", "wins"],
    "Cost (USD)": ["cost (usd)", "cost", "cost usd", "spend", "spend (usd)", "fee (usd)"],
    "Submission Date": ["submission date", "submitted on", "due date", "deadline"],
    "Result Date": ["result date", "declared on", "result declared date", "announcement date"],
    "Quarter": ["quarter", "qtr"],
    "Case Studies Submitted": ["case studies submitted", "case studies", "submitted case studies"],
    "Case Studies Won": ["case studies won", "won case studies"],
    "Notes": ["notes", "remarks", "comments"],
}
REQUIRED = ["Award", "BU", "Status", "Submissions", "Won"]


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def secret(key: str, default=None):
    try:
        return st.secrets.get(key, default)
    except Exception:  # no secrets.toml present
        return default


def bu_color_map(bus) -> dict:
    out, extra = {}, iter(EXTRA_COLORS * 5)
    for bu in bus:
        out[bu] = BU_COLORS.get(str(bu).strip().lower()) or next(extra)
    return out


def logo_data_uri(name: str) -> str | None:
    key = next((k for k in LOGO_FILES if k in str(name).lower()), None)
    if not key:
        return None
    path = ASSETS / LOGO_FILES[key]
    if not path.exists():
        return None
    mime = "image/webp" if path.suffix == ".webp" else "image/png"
    return f"data:{mime};base64,{base64.b64encode(path.read_bytes()).decode()}"


def sheet_id_from(url_or_id: str) -> str:
    m = re.search(r"/d/([a-zA-Z0-9_-]+)", url_or_id or "")
    return m.group(1) if m else (url_or_id or "").strip()


# ----------------------------------------------------------------------------
# Data loading
# ----------------------------------------------------------------------------
@st.cache_data(ttl=CACHE_TTL_SECONDS, show_spinner="Loading data from Google Sheets…")
def load_workbook(source: str) -> tuple[dict[str, pd.DataFrame], str]:
    """Return ({tab_name: DataFrame}, loaded_at). source = sheet URL/ID or '__local__'."""
    loaded_at = datetime.now().strftime("%d %b %Y, %H:%M")
    if source == "__local__":
        return pd.read_excel(LOCAL_FILE, sheet_name=None), loaded_at

    # "File -> Share -> Publish to web" link (…/spreadsheets/d/e/2PACX…/pubhtml)
    pub = re.search(r"/spreadsheets/d/e/([a-zA-Z0-9_-]+)", source)
    if pub:
        url = f"https://docs.google.com/spreadsheets/d/e/{pub.group(1)}/pub?output=xlsx"
        resp = requests.get(url, timeout=30)
        if resp.status_code != 200 or "text/html" in resp.headers.get("Content-Type", ""):
            raise PermissionError(
                "Couldn't read the published sheet. In Google Sheets check File → Share → "
                "Publish to web → 'Entire document' is published and 'Automatically republish' is on."
            )
        return pd.read_excel(io.BytesIO(resp.content), sheet_name=None), loaded_at

    sid = sheet_id_from(source)
    sa_info = secret("gcp_service_account")
    if sa_info:  # private sheet, shared with a service account
        import gspread

        gc = gspread.service_account_from_dict(dict(sa_info))
        sh = gc.open_by_key(sid)
        tabs = {}
        for ws in sh.worksheets():
            values = ws.get_all_values()
            if values:
                tabs[ws.title] = pd.DataFrame(values[1:], columns=values[0])
        return tabs, loaded_at

    # "Anyone with the link can view" sheet: download the whole workbook in one go
    url = f"https://docs.google.com/spreadsheets/d/{sid}/export?format=xlsx"
    resp = requests.get(url, timeout=30)
    if resp.status_code != 200 or "text/html" in resp.headers.get("Content-Type", ""):
        raise PermissionError(
            "Couldn't read the Google Sheet. Make sure it's shared as "
            "'Anyone with the link – Viewer', or set up a service account (see README)."
        )
    return pd.read_excel(io.BytesIO(resp.content), sheet_name=None), loaded_at


def find_tab(tabs: dict, name: str):
    for k, v in tabs.items():
        if k.strip().lower() == name.lower():
            return v
    return None


def read_settings(tabs: dict) -> dict:
    df = find_tab(tabs, "Settings")
    out = {"fx": 95.42, "fy_start": 4, "title": "Affle Awards Dashboard"}
    if df is None or df.shape[1] < 2:
        return out
    for _, r in df.iterrows():
        key, val = str(r.iloc[0]).strip().lower(), r.iloc[1]
        try:
            if key.startswith("fx"):
                out["fx"] = float(val)
            elif "fiscal" in key or "fy start" in key:
                out["fy_start"] = int(float(val))
            elif "title" in key and str(val).strip() and str(val) != "nan":
                out["title"] = str(val).strip()
        except (TypeError, ValueError):
            pass
    return out


def quarter_label(d: pd.Timestamp, fy_start: int) -> tuple[str, tuple]:
    q = (d.month - fy_start) % 12 // 3 + 1
    if fy_start == 1:
        return f"Q{q} {d.year}", (d.year, q)
    fy = d.year + (1 if d.month >= fy_start else 0)
    return f"Q{q} FY{str(fy)[-2:]}", (fy, q)


def quarter_sort_key(label: str) -> tuple:
    m = re.search(r"Q\s*([1-4])\D*?(\d{2,4})", str(label), re.I)
    if not m:
        return (9999, 9)
    year = int(m.group(2))
    return (year + 2000 if year < 100 else year, int(m.group(1)))


def normalise_status(s: str) -> str:
    s = str(s).strip().lower()
    if s.startswith(("declared", "result", "won", "lost", "closed")):
        return "Declared"
    if s.startswith(("upcoming", "planned", "to submit", "not submitted")):
        return "Upcoming"
    return "Pending" if s else "Pending"


def parse_date(value):
    """Parse one date cell: real dates, ISO text (2026-08-05) or day-first text (05/08/2026)."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return pd.NaT
    if isinstance(value, (pd.Timestamp, datetime)):
        return pd.Timestamp(value)
    s = str(value).strip()
    if not s or s.lower() in {"nan", "nat", "none", "-"}:
        return pd.NaT
    iso = bool(re.match(r"^\d{4}-\d{1,2}-\d{1,2}", s))
    return pd.to_datetime(s, errors="coerce", dayfirst=not iso)


def normalise_quarter(value, fy_start: int) -> str | None:
    """Turn whatever was typed in the Quarter column into 'Q1 FY27' (or 'Q1 2026').
    Accepts quarter text in many styles, or a date typed in the wrong column."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (pd.Timestamp, datetime)):
        return quarter_label(pd.Timestamp(value), fy_start)[0]
    s = str(value).strip()
    if not s or s.lower() in {"nan", "none", "nat", "-"}:
        return None
    m = re.search(r"\bQ(?:TR|UARTER)?\s*-?\s*([1-4])\b", s, re.I)
    if m:
        q = int(m.group(1))
        rest = s[:m.start()] + " " + s[m.end():]
        y = re.search(r"(\d{4}|\d{2})", rest)
        if not y:
            return f"Q{q}"
        yr = int(y.group(1))
        yr = yr + 2000 if yr < 100 else yr
        return f"Q{q} {yr}" if fy_start == 1 else f"Q{q} FY{str(yr)[-2:]}"
    d = parse_date(s)  # a date typed in the Quarter column
    if pd.notna(d):
        return quarter_label(d, fy_start)[0]
    return s


def clean_entries(raw: pd.DataFrame, fy_start: int) -> tuple[pd.DataFrame, list[str]]:
    df = raw.copy()
    lookup = {str(c).strip().lower(): c for c in df.columns}
    rename = {}
    for canon, aliases in COLUMN_ALIASES.items():
        for a in aliases:
            if a in lookup:
                rename[lookup[a]] = canon
                break
    df = df.rename(columns=rename)
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"The 'Entries' tab is missing column(s): {', '.join(missing)}")
    for c in COLUMN_ALIASES:
        if c not in df.columns:
            df[c] = None

    df["_q_raw"] = df["Quarter"]
    text_cols = ["Award", "BU", "Status", "Quarter", "Case Studies Submitted", "Case Studies Won", "Notes"]
    for c in text_cols:
        df[c] = df[c].fillna("").astype(str).str.strip().replace({"nan": "", "None": ""})
    df = df[df["Award"] != ""].copy()
    df["BU"] = df["BU"].replace("", "Unassigned")
    for c in ["Submissions", "Won", "Cost (USD)"]:
        df[c] = pd.to_numeric(
            df[c].astype(str).str.replace(r"[,$₹\s]", "", regex=True), errors="coerce"
        ).fillna(0)
    df["Status"] = df["Status"].map(normalise_status)
    for c in ["Submission Date", "Result Date"]:
        if not pd.api.types.is_datetime64_any_dtype(df[c]):
            df[c] = pd.to_datetime(df[c].map(parse_date), errors="coerce")

    # Quarter: manual override > Result Date > Submission Date
    labels = []
    for _, r in df.iterrows():
        manual = normalise_quarter(r["_q_raw"], fy_start)
        if manual:
            labels.append(manual)
        elif pd.notna(r["Result Date"]):
            labels.append(quarter_label(r["Result Date"], fy_start)[0])
        elif pd.notna(r["Submission Date"]):
            labels.append(quarter_label(r["Submission Date"], fy_start)[0])
        else:
            labels.append("Unassigned")
    df["Quarter"] = labels
    df = df.drop(columns="_q_raw")

    warnings = []
    n_unq = int(((df["Quarter"] == "Unassigned") & (df["Status"] == "Declared")).sum())
    if n_unq:
        warnings.append(f"{n_unq} declared row(s) have no Result Date or Quarter, so they show as 'Unassigned' in quarterly views.")
    bad = df[df["Won"] > df["Submissions"]]
    if len(bad):
        warnings.append(f"{len(bad)} row(s) have more wins than submissions: " + ", ".join(bad["Award"].head(3)))
    n_cost = int(((df["Cost (USD)"] == 0) & (df["Status"] != "Upcoming") & (df["Submissions"] > 0)).sum())
    if n_cost:
        warnings.append(f"{n_cost} submitted row(s) have no cost entered, so spend totals are understated.")
    return df, warnings


def parse_case_studies(text: str) -> list[tuple[str, int]]:
    """'Kissht – 2, Redbus' -> [('Kissht', 2), ('Redbus', 1)]"""
    out = []
    if str(text or "").strip().lower() in {"", "0", "-", "–", "none", "na", "n/a", "nil", "nan", "no", "not won"}:
        return out
    for part in str(text or "").split(","):
        part = part.strip()
        if not part or part.lower() in {"0", "-", "–", "none", "na", "n/a", "nil"}:
            continue
        m = re.match(r"^(.*?)\s*[–—-]\s*(\d+)\s*$", part)
        out.append((m.group(1).strip(), int(m.group(2))) if m else (part, 1))
    return out



# ----------------------------------------------------------------------------
# Visual system
# ----------------------------------------------------------------------------
import html as _html
import plotly.io as pio

NAVY, INK, MUTED, LINE, GOLD = "#0E1A3A", "#1D2740", "#6A7389", "#E4E8EF", "#E8B44A"
pio.templates["affle"] = go.layout.Template(layout=dict(
    font=dict(family="DM Sans, Segoe UI, sans-serif", color=INK, size=12),
    colorway=["#0ea99b", "#ff7a4a", "#7c6ff0", "#f2a53c", "#ef5da8", "#3b8fd6", "#5bb85c"],
    plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
    xaxis=dict(showgrid=False, linecolor=LINE, ticks="", title=None),
    yaxis=dict(gridcolor="#EEF1F5", zeroline=False, ticks=""),
    hoverlabel=dict(bgcolor=NAVY, font=dict(color="white", family="DM Sans"), bordercolor=NAVY),
    barcornerradius=5,
    title=dict(font=dict(family="Sora, Segoe UI, Helvetica, Arial, sans-serif", size=15, color=NAVY, weight=700), x=0, xanchor="left"),
    legend=dict(orientation="h", yanchor="top", y=-0.14, x=0, title=None, font=dict(size=11)),
    margin=dict(l=8, r=8, t=44, b=8),
))
pio.templates.default = "plotly_white+affle"

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@500;600;700;800&family=DM+Sans:opsz,wght@9..40,400;9..40,500;9..40,600;9..40,700&display=swap');
.stApp {font-family:'DM Sans', 'Segoe UI', sans-serif; background:#F4F6FA;}
.block-container {padding-top:1.4rem; max-width:1320px;}
h1,h2,h3,.sora {font-family:'Sora', 'Segoe UI', sans-serif !important; letter-spacing:-0.02em;}
header[data-testid="stHeader"] {background:transparent;}

/* cards = bordered containers */
div[data-testid="stLayoutWrapper"] > div[data-testid="stVerticalBlock"],
div[data-testid="stVerticalBlockBorderWrapper"] {
  background:#fff; border-color:#E4E8EF !important; border-radius:14px !important;
  box-shadow:0 1px 2px rgba(14,26,58,.04), 0 6px 18px rgba(14,26,58,.04);}

/* pill tabs */
.stTabs div[role="tablist"] {display:flex; gap:4px; background:#fff; padding:5px; border-radius:12px;
  border:1px solid #E4E8EF; overflow-x:auto;}
.stTabs [role="tab"] {height:38px; padding:0 16px !important; border-radius:9px; white-space:nowrap; margin:0 !important;}
.stTabs [role="tab"] p {font-weight:600; font-size:14px; color:#4A5470;}
.stTabs [role="tab"]:hover p {color:#0E1A3A;}
.stTabs [role="tab"][aria-selected="true"] {background:#0E1A3A;}
.stTabs [role="tab"][aria-selected="true"] p {color:#fff !important;}
.stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"], .stTabs .react-aria-SelectionIndicator {display:none;}
.stTabs [role="tab"] {display:flex; align-items:center; cursor:pointer; transition:background .15s;}
.stTabs [data-baseweb="tab-panel"], .stTabs div[role="tabpanel"] {padding-top:18px;}

/* section titles */
.sec {font-family:'Sora',sans-serif; font-size:18px; font-weight:700; color:#0E1A3A; margin:6px 0 2px;}
.sec-sub {color:#6A7389; font-size:13px; margin-bottom:10px;}

/* sidebar */
section[data-testid="stSidebar"] {background:#fff; border-right:1px solid #E4E8EF;}
.side-brand {display:flex; align-items:center; gap:10px; padding:4px 0 14px; border-bottom:1px solid #E4E8EF; margin-bottom:14px;}
.side-brand img {height:30px; border-radius:6px;}
.side-brand b {font-family:'Sora',sans-serif; font-size:15px; color:#0E1A3A;}

/* metric (quarter compare) */
div[data-testid="stMetric"] {background:#fff; border:1px solid #E4E8EF; border-radius:12px; padding:12px 14px;}
div[data-testid="stMetricValue"] {font-family:'Sora',sans-serif; font-weight:700; color:#0E1A3A;}

/* status pill */
.pill {display:inline-block; padding:3px 10px; border-radius:999px; font-size:12px; font-weight:600;}
</style>
""",
    unsafe_allow_html=True,
)

COMMON_HTML_CSS = """
<style>
.kgrid {display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:10px; margin:2px 0 6px;}
.k {background:#fff; border:1px solid #E4E8EF; border-radius:12px; padding:13px 14px 12px; position:relative; overflow:hidden;}
.k::before {content:""; position:absolute; left:0; top:0; bottom:0; width:3px; background:var(--c,#0ea99b);}
.k .l {font-size:12.5px; color:#6A7389; font-weight:600;}
.k .v {font-family:'Sora',sans-serif; font-size:24px; font-weight:700; color:#0E1A3A; margin-top:4px; line-height:1.15; white-space:nowrap;}
.k .s {font-size:12px; color:#8A92A6; margin-top:3px;}
</style>
"""


def html_block(body: str, js: str = ""):
    """Render HTML (+ optional JS) without an iframe; falls back to static HTML."""
    if js:
        try:
            st.html(body + f"<script>{js}</script>", unsafe_allow_javascript=True)
            return
        except TypeError:  # older Streamlit without JS support
            pass
    st.html(body)


COUNT_UP_JS = """
(function(){
  if (window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
  document.querySelectorAll('.cu:not([data-done])').forEach(function(el){
    el.dataset.done = '1';
    var to = parseFloat(el.dataset.to||'0'), dec = parseInt(el.dataset.dec||'0'),
        pre = el.dataset.pre||'', suf = el.dataset.suf||'', loc = el.dataset.loc||'en-US';
    var fmt = function(v){ return pre + v.toLocaleString(loc,{minimumFractionDigits:dec, maximumFractionDigits:dec}) + suf; };
    var t0 = performance.now(), dur = 1400;
    function step(t){ var p = Math.min(1,(t-t0)/dur), e = 1-Math.pow(1-p,3); el.textContent = fmt(to*e); if(p<1) requestAnimationFrame(step); }
    requestAnimationFrame(step);
  });
})();
"""


def esc(s) -> str:
    return _html.escape(str(s))


def cu(value: float, dec: int = 0, pre: str = "", suf: str = "", loc: str = "en-US") -> str:
    """A number that counts up on load; shows the final value if JS is unavailable."""
    final = f"{pre}{value:,.{dec}f}{suf}"
    return (f'<span class="cu" data-to="{value}" data-dec="{dec}" data-pre="{esc(pre)}" '
            f'data-suf="{esc(suf)}" data-loc="{loc}">{esc(final)}</span>')


def mini_cards(items):
    """items: list of (label, value_html, sub, color)."""
    cards = "".join(
        f'<div class="k" style="--c:{c}"><div class="l">{esc(l)}</div><div class="v">{v}</div>'
        f'<div class="s">{esc(s)}</div></div>' for l, v, s, c in items)
    html_block(COMMON_HTML_CSS + f'<div class="kgrid">{cards}</div>', COUNT_UP_JS)


def section(title: str, sub: str = ""):
    st.markdown(f'<div class="sec">{esc(title)}</div>' + (f'<div class="sec-sub">{esc(sub)}</div>' if sub else ""),
                unsafe_allow_html=True)


def show_df(data, container=None, **kw):
    if isinstance(data, pd.DataFrame):
        data = data.copy()
        num = data.select_dtypes("number").columns
        data[num] = data[num].round(2)
    (container or st).dataframe(data, **kw)


def chart(fig, height=340, container=None):
    fig.update_layout(template="plotly_white+affle", height=height, legend_title_text=None,
                      paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
    fig.update_xaxes(title=None)
    (container or st).plotly_chart(fig, width="stretch", theme=None, config={"displaylogo": False})


def donut(labels, values, colors, title, center_big, center_small):
    fig = go.Figure(go.Pie(labels=labels, values=values, hole=0.68, sort=False,
                           marker=dict(colors=colors, line=dict(color="#fff", width=2)),
                           textinfo="none", hovertemplate="%{label}: %{value:,.0f} (%{percent})<extra></extra>"))
    fig.update_layout(title=title, showlegend=True,
                      annotations=[dict(text=f"<b style='font-size:22px'>{center_big}</b><br>"
                                              f"<span style='font-size:11px;color:{MUTED}'>{center_small}</span>",
                                        x=0.5, y=0.5, showarrow=False, font=dict(family="Sora, Segoe UI, Helvetica, Arial, sans-serif", color=NAVY))])
    return fig


def status_pill(status: str) -> str:
    bg = {"Declared": "#E3F6F3", "Pending": "#FDF1DE", "Upcoming": "#EEF0F4"}.get(status, "#EEF0F4")
    fg = {"Declared": "#0B7F74", "Pending": "#A86A0C", "Upcoming": "#59617A"}.get(status, "#59617A")
    return f'<span class="pill" style="background:{bg};color:{fg}">{esc(status)}</span>'


# ----------------------------------------------------------------------------
# Sidebar: data source + filters
# ----------------------------------------------------------------------------
default_sheet = secret("sheet_url", "")
with st.sidebar:
    logo_uri = logo_data_uri("affle")
    st.markdown(
        f'<div class="side-brand">{f"<img src={chr(34)}{logo_uri}{chr(34)}>" if logo_uri else ""}'
        f'<b>Awards dashboard</b></div>', unsafe_allow_html=True)
    if default_sheet:
        source = default_sheet
    else:
        typed = st.text_input("Google Sheet link", placeholder="https://docs.google.com/spreadsheets/d/…",
                              help="Or set sheet_url in the app's Secrets.")
        source = typed.strip() or "__local__"
        if source == "__local__" and not LOCAL_FILE.exists():
            st.info("Add your Google Sheet link above, or set `sheet_url` in the app's Secrets.")
            st.stop()

try:
    tabs, loaded_at = load_workbook(source)
    settings = read_settings(tabs)
    entries_raw = find_tab(tabs, "Entries")
    if entries_raw is None:
        raise ValueError("No tab named 'Entries' found in the sheet.")
    df_all, data_warnings = clean_entries(entries_raw, settings["fy_start"])
except Exception as e:
    st.error(f"**Couldn't load the data.** {e}")
    st.stop()

quarters_all = sorted(df_all["Quarter"].unique(), key=quarter_sort_key)
bus_all = sorted(df_all["BU"].unique(), key=str.lower)
COLORS = bu_color_map(bus_all)

with st.sidebar:
    f_quarter = st.multiselect("Quarter", quarters_all, placeholder="All quarters")
    f_bu = st.multiselect("Business unit", bus_all, placeholder="All business units")
    f_status = st.multiselect("Status", ["Declared", "Pending", "Upcoming"], placeholder="All statuses")
    f_award = st.multiselect("Award", sorted(df_all["Award"].unique()), placeholder="All awards")
    currency = st.segmented_control("Currency", ["US$", "₹"], default="US$") or "US$"
    st.divider()
    if st.button("Refresh data", icon=":material/refresh:", width="stretch"):
        load_workbook.clear()
        st.rerun()
    st.caption(f"{'Live Google Sheet' if source != '__local__' else 'Sample file'} · loaded {loaded_at} · "
               f"auto-refresh every {CACHE_TTL_SECONDS // 60} min")

df = df_all.copy()
for col, sel in [("Quarter", f_quarter), ("BU", f_bu), ("Status", f_status), ("Award", f_award)]:
    if sel:
        df = df[df[col].isin(sel)]

INR = currency == "₹"
FX = settings["fx"] if INR else 1.0
SYM = "₹" if INR else "$"
LOC = "en-IN" if INR else "en-US"


def money(v: float, decimals: int = 0) -> str:
    v = v * FX
    if INR:  # Indian digit grouping
        s = f"{abs(v):.{decimals}f}"
        whole, _, frac = s.partition(".")
        if len(whole) > 3:
            head, tail = whole[:-3], whole[-3:]
            head = ",".join([head[max(0, i - 2):i] for i in range(len(head), 0, -2)][::-1])
            whole = f"{head},{tail}"
        return f"{'-' if v < 0 else ''}₹{whole}{'.' + frac if frac else ''}"
    return f"${v:,.{decimals}f}"


if df.empty:
    st.info("No rows match the current filters. Clear a filter in the sidebar.")
    st.stop()

# ----------------------------------------------------------------------------
# Core numbers
# ----------------------------------------------------------------------------
declared = df[df["Status"] == "Declared"]
total_subs = int(df["Submissions"].sum())
declared_subs = int(declared["Submissions"].sum())
won = int(df["Won"].sum())
spend = float(df["Cost (USD)"].sum())
declared_spend = float(declared["Cost (USD)"].sum())
win_rate = won / declared_subs if declared_subs else 0
cost_per_win = declared_spend / won if won else 0
pending_subs = int(df.loc[df.Status == "Pending", "Submissions"].sum())
participated = df[df["Status"] != "Upcoming"]
programs = participated["Award"].nunique()  # programs actually entered
upcoming_programs = df.loc[df["Status"] == "Upcoming", "Award"].nunique()

# ----------------------------------------------------------------------------
# Hero: headline sentence + animated podium of top BUs by wins
# ----------------------------------------------------------------------------
bu_wins = df.groupby("BU")["Won"].sum().sort_values(ascending=False)
top = [(b, int(w)) for b, w in bu_wins.items() if w > 0][:3]
order = [1, 0, 2] if len(top) == 3 else ([1, 0] if len(top) == 2 else [0])
medal = {0: GOLD, 1: "#C3CCD8", 2: "#C98A5A"}
max_w = max([w for _, w in top], default=1)
podium_cols = ""
for idx in order:
    if idx >= len(top):
        continue
    bu, w = top[idx]
    uri = logo_data_uri(bu)
    name = f'<img src="{uri}" alt="{esc(bu)}">' if uri else f"<span>{esc(bu)}</span>"
    h = 46 + 104 * (w / max_w)
    podium_cols += (
        f'<div class="pc"><div class="pn">{name}</div><div class="pw">{cu(w)}</div>'
        f'<div class="pb" style="height:{h:.0f}px;--m:{medal[idx]};animation-delay:{0.15 * (2 - idx):.2f}s">'
        f'<i>{idx + 1}</i></div></div>')
podium = (f'<div class="podium"><div class="pt">Most awards won</div><div class="prow">{podium_cols}</div></div>'
          if top else "")

lede = (f"{cu(won)} awards won from {cu(declared_subs)} declared entries"
        if declared_subs else f"{cu(total_subs)} submissions in progress")
hero = f"""
<style>
.hero {{position:relative; display:grid; grid-template-columns:minmax(0,1.35fr) minmax(0,1fr); gap:28px;
  background:radial-gradient(120% 140% at 100% 0%, #1F3A75 0%, #0E1A3A 55%, #0A132B 100%);
  border-radius:18px; padding:26px 30px 0; color:#fff; overflow:hidden; margin-bottom:14px;}}
.hero::after {{content:""; position:absolute; inset:0; pointer-events:none;
  background:repeating-linear-gradient(115deg, rgba(255,255,255,.025) 0 2px, transparent 2px 22px);}}
.hl {{padding-bottom:24px; position:relative; z-index:1;}}
.live {{display:inline-flex; align-items:center; gap:7px; font-size:12px; color:#B9C3DA;
  background:rgba(255,255,255,.07); border:1px solid rgba(255,255,255,.1); padding:4px 10px; border-radius:999px;}}
.live i {{width:7px; height:7px; border-radius:50%; background:#3DDC97; box-shadow:0 0 0 0 rgba(61,220,151,.6);
  animation:pulse 2s infinite;}}
@keyframes pulse {{0%{{box-shadow:0 0 0 0 rgba(61,220,151,.55)}}70%{{box-shadow:0 0 0 8px rgba(61,220,151,0)}}100%{{box-shadow:0 0 0 0 rgba(61,220,151,0)}}}}
.ht {{font-family:'Sora',sans-serif; font-size:15px; font-weight:600; color:#B9C3DA; margin:16px 0 6px;}}
.hh {{font-family:'Sora',sans-serif; font-size:clamp(26px,3.1vw,40px); font-weight:700; line-height:1.15;
  letter-spacing:-0.025em; margin:0; max-width:22ch;}}
.hh .cu {{color:{GOLD};}}
.hs {{display:flex; flex-wrap:wrap; gap:26px; margin-top:22px;}}
.hsv {{display:block; font-family:'Sora','Segoe UI',sans-serif; font-size:24px; font-weight:700; color:#fff; line-height:1.2;}}
.hsl {{display:block; font-size:12.5px; color:#9FAAC4; margin-top:2px;}}
.podium {{position:relative; z-index:1; display:flex; flex-direction:column; justify-content:flex-end;}}
.pt {{font-size:12.5px; color:#9FAAC4; text-align:center; margin-bottom:10px;}}
.prow {{display:flex; align-items:flex-end; justify-content:center; gap:12px;}}
.pc {{flex:1; max-width:150px; display:flex; flex-direction:column; align-items:center;}}
.pn {{height:30px; display:flex; align-items:center;}}
.pn img {{max-height:22px; max-width:110px;}}
.pn span {{font-weight:600; font-size:13px;}}
.pw {{font-family:'Sora',sans-serif; font-size:26px; font-weight:700; margin:2px 0 8px;}}
.pb {{width:100%; border-radius:10px 10px 0 0; position:relative; transform-origin:bottom;
  background:linear-gradient(180deg, rgba(255,255,255,.16), rgba(255,255,255,.05));
  border:1px solid rgba(255,255,255,.14); border-bottom:none; animation:rise .9s cubic-bezier(.2,.8,.2,1) both;}}
.pb i {{position:absolute; top:10px; left:50%; transform:translateX(-50%); width:26px; height:26px; border-radius:50%;
  background:var(--m); color:#0E1A3A; font-style:normal; font-weight:800; font-size:13px;
  display:flex; align-items:center; justify-content:center; font-family:'Sora',sans-serif;}}
@keyframes rise {{from {{transform:scaleY(0); opacity:.3;}} to {{transform:scaleY(1); opacity:1;}}}}
@media (prefers-reduced-motion: reduce) {{ .pb, .live i {{animation:none;}} }}
@media (max-width: 760px) {{ .hero {{grid-template-columns:1fr; padding:22px 20px 0;}} }}
</style>
<div class="hero">
  <div class="hl">
    <span class="live"><i></i>{'Live from Google Sheets' if source != '__local__' else 'Sample data'} · as of {datetime.now():%d %b %Y}</span>
    <div class="ht">{esc(settings['title'])}</div>
    <h1 class="hh">{lede}</h1>
    <div class="hs">
      <div><span class="hsv">{cu(win_rate * 100, 0, suf='%')}</span><span class="hsl">win rate on declared</span></div>
      <div><span class="hsv">{cu(spend * FX, 0, pre=SYM, loc=LOC)}</span><span class="hsl">total spend</span></div>
      <div><span class="hsv">{cu(cost_per_win * FX, 0, pre=SYM, loc=LOC)}</span><span class="hsl">cost per award won</span></div>
      <div><span class="hsv">{cu(programs)}</span><span class="hsl">award programs entered{f" · {upcoming_programs} upcoming" if upcoming_programs else ""}</span></div>
    </div>
  </div>
  {podium}
</div>
"""
html_block(hero, COUNT_UP_JS)

mini_cards([
    ("Total submissions", cu(total_subs), f"{programs} award programs entered", "#0ea99b"),
    ("Results declared", cu(declared_subs), f"{declared['Award'].nunique()} programs", "#3b8fd6"),
    ("Awards won", cu(won), f"{win_rate:.0%} win rate", GOLD),
    ("Pending results", cu(pending_subs), f"{df.loc[df.Status == 'Pending', 'Award'].nunique()} programs", "#f2a53c"),
    ("Total spend", cu(spend * FX, 0, pre=SYM, loc=LOC), "all entries with cost", "#7c6ff0"),
    ("Cost per award won", cu(cost_per_win * FX, 2, pre=SYM, loc=LOC), "declared spend ÷ wins", "#ef5da8"),
])

if data_warnings:
    with st.expander(f"{len(data_warnings)} data note(s) to fix in the Google Sheet", icon=":material/info:"):
        for w in data_warnings:
            st.markdown(f"- {w}")

t_over, t_award, t_bu, t_q, t_client, t_pipe, t_custom, t_data = st.tabs(
    ["Overview", "Awards", "Business units", "Quarters", "Clients", "Pipeline", "Build your own", "Data"])

# ----------------------------------------------------------------------------
# Overview
# ----------------------------------------------------------------------------
with t_over:
    c1, c2, c3 = st.columns(3)
    with c1.container(border=True):
        ss = df.groupby("Status")["Submissions"].sum().reindex(["Declared", "Pending", "Upcoming"]).fillna(0)
        ss = ss[ss > 0]
        chart(donut(ss.index, ss.values, [STATUS_COLORS[s] for s in ss.index], "Submission status",
                    f"{total_subs:,}", "submissions"), 300)
    with c2.container(border=True):
        g = go.Figure(go.Indicator(
            mode="gauge+number", value=win_rate * 100, number=dict(suffix="%", font=dict(family="Sora, Segoe UI, Helvetica, Arial, sans-serif", size=40, color=NAVY)),
            gauge=dict(axis=dict(range=[0, 100], tickwidth=0, tickcolor="white", tickfont=dict(size=10, color=MUTED)),
                       bar=dict(color="#0ea99b", thickness=0.32), bgcolor="#EEF1F5", borderwidth=0,
                       threshold=dict(line=dict(color=GOLD, width=3), thickness=0.8, value=win_rate * 100))))
        g.update_layout(title="Win rate on declared entries",
                        annotations=[dict(text=f"{won} won of {declared_subs} declared", x=0.5, y=-0.05,
                                          showarrow=False, font=dict(size=12, color=MUTED))])
        chart(g, 300)
    with c3.container(border=True):
        sb = df.groupby("BU")["Cost (USD)"].sum().sort_values(ascending=False)
        sb = sb[sb > 0]
        chart(donut(sb.index, (sb * FX).values, [COLORS[b] for b in sb.index], f"Spend by BU ({currency})",
                    money(sb.sum()), "total spend"), 300)

    c4, c5 = st.columns([1, 1.15])
    with c4.container(border=True):
        q = df.groupby(["Quarter", "BU"], as_index=False)["Won"].sum()
        fig = px.bar(q, x="Quarter", y="Won", color="BU", color_discrete_map=COLORS,
                     category_orders={"Quarter": quarters_all}, title="Awards won per quarter")
        chart(fig, 360)
    with c5.container(border=True):
        ac = df.groupby(["Award", "BU"], as_index=False)["Cost (USD)"].sum()
        ac = ac[ac["Cost (USD)"] > 0]
        ac["Spend"] = ac["Cost (USD)"] * FX
        top_aw = ac.groupby("Award")["Spend"].sum().nlargest(8).index
        ac = ac[ac["Award"].isin(top_aw)]
        ac["Short"] = ac["Award"].str.slice(0, 38) + ac["Award"].map(lambda s: "…" if len(s) > 38 else "")
        fig = px.bar(ac, y="Short", x="Spend", color="BU", orientation="h", color_discrete_map=COLORS,
                     title=f"Award-wise total cost, top 8 ({currency})", hover_data={"Award": True, "Short": False})
        fig.update_yaxes(categoryorder="total ascending", title=None)
        fig.update_xaxes(title=None)
        chart(fig, 360)

    with st.container(border=True):
        section("Business unit leaderboard", "Ranked by awards won")
        t = df.groupby("BU").agg(Programs=("Award", lambda s: s[df.loc[s.index, "Status"] != "Upcoming"].nunique()), Submissions=("Submissions", "sum"),
                                 Won=("Won", "sum"), Spend=("Cost (USD)", "sum")).reset_index()
        dsub = declared.groupby("BU")["Submissions"].sum()
        t["Win rate"] = (t["Won"] / t["BU"].map(dsub)).where(t["BU"].map(dsub) > 0)
        t["Share of spend"] = t["Spend"] / t["Spend"].sum() if t["Spend"].sum() else 0
        t["Cost per win"] = (t["Spend"] / t["Won"]).where(t["Won"] > 0) * FX
        t["Spend"] = t["Spend"] * FX
        show_df(t.sort_values(["Won", "Spend"], ascending=False), hide_index=True, width="stretch", column_config={
            "Spend": st.column_config.NumberColumn(f"Spend ({currency})", format="localized"),
            "Cost per win": st.column_config.NumberColumn(f"Cost per win ({currency})", format="localized"),
            "Win rate": st.column_config.ProgressColumn(format="percent", min_value=0, max_value=1),
            "Share of spend": st.column_config.ProgressColumn(format="percent", min_value=0, max_value=1),
        })

# ----------------------------------------------------------------------------
# Awards explorer
# ----------------------------------------------------------------------------
with t_award:
    aw = df.groupby("Award").agg(
        Status=("Status", lambda s: s.mode().iat[0]),
        Quarter=("Quarter", lambda s: ", ".join(sorted(set(s), key=quarter_sort_key))),
        BUs=("BU", "nunique"), Submissions=("Submissions", "sum"), Won=("Won", "sum"),
        Cost=("Cost (USD)", "sum")).reset_index()
    aw["Win rate"] = (aw["Won"] / aw["Submissions"]).where((aw["Status"] == "Declared") & (aw["Submissions"] > 0))
    aw["Cost per win"] = (aw["Cost"] / aw["Won"]).where(aw["Won"] > 0)
    aw = aw.sort_values("Cost", ascending=False)

    pick = st.selectbox("Choose an award", aw["Award"], index=0, key="award_pick")
    a = df[df["Award"] == pick]
    ar = aw[aw["Award"] == pick].iloc[0]
    st.markdown(f"{status_pill(ar['Status'])} &nbsp;<span style='color:{MUTED};font-size:13px'>"
                f"Quarter: {esc(ar['Quarter'])}</span>", unsafe_allow_html=True)
    mini_cards([
        ("Total cost", cu(ar["Cost"] * FX, 2, pre=SYM, loc=LOC), f"across {int(ar['BUs'])} BU(s)", "#7c6ff0"),
        ("Submissions", cu(ar["Submissions"]), "entries", "#0ea99b"),
        ("Won", cu(ar["Won"]), "awards", GOLD),
        ("Win rate", f"{ar['Win rate']:.0%}" if pd.notna(ar["Win rate"]) else "–", "on this award", "#3b8fd6"),
        ("Cost per win", money(ar["Cost per win"], 2) if pd.notna(ar["Cost per win"]) else "–", "", "#ef5da8"),
    ])
    c1, c2 = st.columns(2)
    with c1.container(border=True):
        cs = a.groupby("BU")["Cost (USD)"].sum()
        cs = cs[cs > 0]
        if cs.empty:
            st.caption("No cost entered for this award yet.")
        else:
            chart(donut(cs.index, (cs * FX).values, [COLORS[b] for b in cs.index], f"Cost split by BU ({currency})",
                        money(cs.sum()), "total"), 300)
    with c2.container(border=True):
        sw = a.groupby("BU", as_index=False)[["Submissions", "Won"]].sum().melt(id_vars="BU", var_name="Metric")
        fig = px.bar(sw, x="BU", y="value", color="Metric", barmode="group", text="value",
                     color_discrete_map={"Submissions": "#C9D2E0", "Won": GOLD}, title="Submissions vs won by BU")
        fig.update_traces(textposition="outside", cliponaxis=False)
        fig.update_yaxes(title=None)
        chart(fig, 300)
    detail = a[["BU", "Status", "Submissions", "Won", "Cost (USD)", "Case Studies Submitted", "Case Studies Won", "Notes"]].copy()
    detail["Cost (USD)"] = detail["Cost (USD)"] * FX
    show_df(detail.rename(columns={"Cost (USD)": f"Cost ({currency})"}), hide_index=True, width="stretch")

    st.write("")
    with st.container(border=True):
        section("All awards by total cost", "Every award program, highest spend first")
        tbl = aw.copy()
        tbl["Cost"] = tbl["Cost"] * FX
        tbl["Cost per win"] = tbl["Cost per win"] * FX
        show_df(tbl, hide_index=True, width="stretch", column_config={
            "Award": st.column_config.TextColumn(width="large"),
            "Cost": st.column_config.ProgressColumn(f"Total cost ({currency})", format="localized",
                                                    min_value=0, max_value=float(tbl["Cost"].max() or 1)),
            "Cost per win": st.column_config.NumberColumn(f"Cost per win ({currency})", format="localized"),
            "Win rate": st.column_config.ProgressColumn(format="percent", min_value=0, max_value=1),
        })

# ----------------------------------------------------------------------------
# Business unit explorer
# ----------------------------------------------------------------------------
with t_bu:
    bu_rank = df.groupby("BU")["Cost (USD)"].sum().sort_values(ascending=False).index.tolist()
    bpick = st.selectbox("Choose a business unit", bu_rank, key="bu_pick")
    b = df[df["BU"] == bpick]
    bd = b[b["Status"] == "Declared"]
    b_won, b_dsubs, b_spend = int(b["Won"].sum()), int(bd["Submissions"].sum()), float(b["Cost (USD)"].sum())
    uri = logo_data_uri(bpick)
    if uri:
        st.markdown(f'<span style="display:inline-flex;background:{NAVY};padding:6px 12px;border-radius:8px">'
                    f'<img src="{uri}" style="max-height:20px"></span>', unsafe_allow_html=True)
    mini_cards([
        ("Programs entered", cu(b.loc[b["Status"] != "Upcoming", "Award"].nunique()), "excludes upcoming", "#0ea99b"),
        ("Submissions", cu(b["Submissions"].sum()), f"{b_dsubs} declared", "#3b8fd6"),
        ("Won", cu(b_won), f"{(b_won / b_dsubs if b_dsubs else 0):.0%} win rate", GOLD),
        ("Spend", cu(b_spend * FX, 0, pre=SYM, loc=LOC), f"{(b_spend / spend if spend else 0):.0%} of total", "#7c6ff0"),
        ("Cost per win", money(bd["Cost (USD)"].sum() / b_won, 2) if b_won else "–", "declared spend ÷ wins", "#ef5da8"),
    ])
    c1, c2 = st.columns([1.2, 1])
    with c1.container(border=True):
        bc = b.groupby(["Award", "Status"], as_index=False)["Cost (USD)"].sum()
        bc = bc[bc["Cost (USD)"] > 0]
        bc["Spend"] = bc["Cost (USD)"] * FX
        bc["Short"] = bc["Award"].str.slice(0, 36) + bc["Award"].map(lambda s: "…" if len(s) > 36 else "")
        if bc.empty:
            st.caption("No cost entered for this BU yet.")
        else:
            fig = px.bar(bc, y="Short", x="Spend", color="Status", orientation="h", color_discrete_map=STATUS_COLORS,
                         title=f"Cost by award ({currency})", hover_data={"Award": True, "Short": False})
            fig.update_yaxes(categoryorder="total ascending", title=None)
            fig.update_xaxes(title=None)
            chart(fig, max(300, 46 * len(bc) + 80))
    with c2.container(border=True):
        bq = b.groupby("Quarter", as_index=False)["Won"].sum()
        fig = px.bar(bq, x="Quarter", y="Won", text="Won", category_orders={"Quarter": quarters_all},
                     color_discrete_sequence=[COLORS.get(bpick, "#0ea99b")], title="Awards won by quarter")
        fig.update_traces(textposition="outside", cliponaxis=False)
        fig.update_yaxes(title=None)
        chart(fig, 300)
    bt = b[["Award", "Status", "Quarter", "Submissions", "Won", "Cost (USD)", "Case Studies Won"]].copy()
    bt["Cost (USD)"] = bt["Cost (USD)"] * FX
    show_df(bt.rename(columns={"Cost (USD)": f"Cost ({currency})"}).sort_values("Won", ascending=False),
            hide_index=True, width="stretch", column_config={"Award": st.column_config.TextColumn(width="large")})

# ----------------------------------------------------------------------------
# Quarters: compare two quarters + BU × quarter grid
# ----------------------------------------------------------------------------
with t_q:
    qs = [x for x in quarters_all if x in df["Quarter"].unique()]
    dated = [x for x in qs if x != "Unassigned"]
    if len(qs) < 2:
        st.info("Add Result Dates (or a Quarter) in the sheet to compare quarters. "
                "Right now everything falls in one quarter.")
    else:
        c1, c2 = st.columns(2)
        qa = c1.selectbox("Compare", qs, index=max(0, len(dated) - 2) if len(dated) >= 2 else 0, key="qa")
        qb = c2.selectbox("with", qs, index=qs.index(dated[-1]) if dated else len(qs) - 1, key="qb")

        def qstats(qq):
            d = df[df["Quarter"] == qq]
            dd = d[d["Status"] == "Declared"]
            w, ds = d["Won"].sum(), dd["Submissions"].sum()
            return dict(subs=d["Submissions"].sum(), won=w, spend=d["Cost (USD)"].sum() * FX,
                        rate=(w / ds * 100) if ds else 0, prog=d["Award"].nunique())

        A, B = qstats(qa), qstats(qb)
        m = st.columns(5)
        m[0].metric(f"Submissions · {qb}", f"{B['subs']:,.0f}", f"{B['subs'] - A['subs']:+,.0f} vs {qa}")
        m[1].metric(f"Awards won · {qb}", f"{B['won']:,.0f}", f"{B['won'] - A['won']:+,.0f} vs {qa}")
        m[2].metric(f"Spend · {qb}", money(B["spend"] / FX), f"{'+' if B['spend'] >= A['spend'] else '-'}{SYM}{abs(B['spend'] - A['spend']):,.0f} vs {qa}",
                    delta_color="inverse")
        m[3].metric(f"Win rate · {qb}", f"{B['rate']:.0f}%", f"{B['rate'] - A['rate']:+.0f} pts vs {qa}")
        m[4].metric(f"Programs · {qb}", f"{B['prog']}", f"{B['prog'] - A['prog']:+d} vs {qa}")

        with st.container(border=True):
            cmp_ = df[df["Quarter"].isin([qa, qb])].groupby(["BU", "Quarter"], as_index=False)["Won"].sum()
            fig = px.bar(cmp_, x="BU", y="Won", color="Quarter", barmode="group", text="Won",
                         color_discrete_sequence=["#C9D2E0", "#0ea99b"], category_orders={"Quarter": [qa, qb]},
                         title=f"Awards won by BU: {qa} vs {qb}")
            fig.update_traces(textposition="outside", cliponaxis=False)
            fig.update_yaxes(title=None)
            chart(fig, 340)

    with st.container(border=True):
        section("Business unit × quarter grid")
        metric_q = st.segmented_control("Show", ["Won", "Submissions", "Spend"], default="Won", key="qmetric") or "Won"
        val_col = "Cost (USD)" if metric_q == "Spend" else metric_q
        piv = df.pivot_table(index="BU", columns="Quarter", values=val_col, aggfunc="sum", fill_value=0)
        piv = piv[[c for c in quarters_all if c in piv.columns]] * (FX if metric_q == "Spend" else 1)
        piv["Total"] = piv.sum(axis=1)
        piv = piv.sort_values("Total", ascending=False)
        piv.loc["Total"] = piv.sum()
        st.dataframe(piv.style.format("{:,.0f}"), width="stretch")

# ----------------------------------------------------------------------------
# Clients explorer
# ----------------------------------------------------------------------------
with t_client:
    recs = []
    for _, r in df.iterrows():
        for name, n in parse_case_studies(r["Case Studies Submitted"]):
            recs.append(dict(Client=name, BU=r["BU"], Award=r["Award"], Status=r["Status"], Submitted=n, Won=0))
        for name, n in parse_case_studies(r["Case Studies Won"]):
            recs.append(dict(Client=name, BU=r["BU"], Award=r["Award"], Status=r["Status"], Submitted=0, Won=n))
    if not recs:
        st.info("No clients listed yet. Fill the 'Case Studies Submitted / Won' columns in the sheet.")
    else:
        cs = pd.DataFrame(recs)
        cl = cs.groupby("Client", as_index=False)[["Submitted", "Won"]].sum().sort_values(["Won", "Submitted"], ascending=False)
        c1, c2 = st.columns([1, 1.25])
        with c1:
            cpick = st.selectbox("Choose a client / case study", cl["Client"], key="client_pick")
            c = cs[cs["Client"] == cpick]
            mini_cards([
                ("Entries", cu(c["Submitted"].sum()), "case-study submissions", "#0ea99b"),
                ("Wins", cu(c["Won"].sum()), "awards won", GOLD),
                ("Awards", cu(c["Award"].nunique()), f"via {', '.join(sorted(c['BU'].unique()))}", "#3b8fd6"),
            ])
            ct = c.groupby(["Award", "BU", "Status"], as_index=False)[["Submitted", "Won"]].sum()
            show_df(ct, hide_index=True, width="stretch")
        with c2.container(border=True):
            topc = cl.head(12).melt(id_vars="Client", var_name="Metric")
            fig = px.bar(topc, y="Client", x="value", color="Metric", orientation="h", barmode="group",
                         color_discrete_map={"Submitted": "#C9D2E0", "Won": GOLD}, title="Top clients by awards won")
            fig.update_yaxes(categoryorder="array", categoryarray=cl.head(12)["Client"].tolist()[::-1], title=None)
            fig.update_xaxes(title=None)
            chart(fig, max(360, 34 * min(12, len(cl)) + 90))

# ----------------------------------------------------------------------------
# Pipeline: pending + upcoming
# ----------------------------------------------------------------------------
with t_pipe:
    today = pd.Timestamp(datetime.now().date())
    up = df[df["Status"] == "Upcoming"]
    if not up.empty:
        section("Upcoming submissions")
        cards = ""
        for award, g in up.groupby("Award"):
            d = g["Submission Date"].min()
            if pd.notna(d):
                days = (d - today).days
                when = "Due today" if days == 0 else (f"Due in {days} day{'s' if days != 1 else ''}" if days > 0
                                                      else f"Deadline passed {-days} day(s) ago")
                date_s = d.strftime("%d %b %Y")
            else:
                when, date_s = "No deadline set", ""
            note = "; ".join(sorted({n for n in g["Notes"] if n}))
            bus = ", ".join(sorted(b for b in g["BU"].unique()))
            cards += (f'<div class="k" style="--c:#f2a53c"><div class="l">{esc(when)}</div>'
                      f'<div class="v" style="font-size:16px;white-space:normal">{esc(award)}</div>'
                      f'<div class="s">{esc(date_s)} · BU: {esc(bus)}{" · " + esc(note) if note else ""}</div></div>')
        html_block(COMMON_HTML_CSS.replace("minmax(150px,1fr)", "minmax(260px,1fr)") + f'<div class="kgrid">{cards}</div>')
    pe = df[df["Status"] == "Pending"]
    section("Submitted, awaiting results", f"{int(pe['Submissions'].sum())} submissions across {pe['Award'].nunique()} programs")
    if pe.empty:
        st.caption("Nothing pending in the current filter.")
    else:
        pt = pe.groupby("Award").apply(lambda g: pd.Series({
            "Submitted on": g["Submission Date"].min(), "Submissions": g["Submissions"].sum(),
            "BU split": ", ".join(f"{b} – {int(n)}" for b, n in zip(g["BU"], g["Submissions"])),
            f"Cost ({currency})": g["Cost (USD)"].sum() * FX,
        }), include_groups=False).reset_index()
        show_df(pt, hide_index=True, width="stretch", column_config={
            "Submitted on": st.column_config.DateColumn(format="DD MMM YYYY"),
            "Award": st.column_config.TextColumn(width="large")})

# ----------------------------------------------------------------------------
# Build your own view
# ----------------------------------------------------------------------------
with t_custom:
    st.caption("Pick any combination. The chart and table update instantly and respect the sidebar filters.")
    dims = ["BU", "Quarter", "Award", "Status"]
    metrics = ["Awards won", "Submissions", f"Spend ({currency})", "Win rate", f"Cost per win ({currency})", "Award programs"]
    c1, c2, c3, c4 = st.columns(4)
    row_dim = c1.selectbox("Group by", dims, index=1)
    split_dim = c2.selectbox("Split by", ["(none)"] + [d_ for d_ in dims if d_ != row_dim])
    metric = c3.selectbox("Metric", metrics)
    chart_t = c4.selectbox("Chart", ["Bar", "Stacked bar", "Line", "Donut", "Table only"])

    keys = [row_dim] + ([] if split_dim == "(none)" else [split_dim])
    g = df.groupby(keys).agg(Won=("Won", "sum"), Submissions=("Submissions", "sum"),
                             Cost=("Cost (USD)", "sum"), Programs=("Award", "nunique")).reset_index()
    decl = declared.groupby(keys).agg(DSubs=("Submissions", "sum"), DCost=("Cost (USD)", "sum")).reset_index()
    g = g.merge(decl, on=keys, how="left").fillna({"DSubs": 0, "DCost": 0})
    if metric == "Awards won":
        g["Value"] = g["Won"]
    elif metric == "Submissions":
        g["Value"] = g["Submissions"]
    elif metric == "Award programs":
        g["Value"] = g["Programs"]
    elif metric == "Win rate":
        g["Value"] = (g["Won"] / g["DSubs"]).where(g["DSubs"] > 0)
    elif metric.startswith("Spend"):
        g["Value"] = g["Cost"] * FX
    else:
        g["Value"] = (g["DCost"] / g["Won"]).where(g["Won"] > 0) * FX

    orders = {"Quarter": quarters_all, "Status": ["Declared", "Pending", "Upcoming"]}
    color = None if split_dim == "(none)" else split_dim
    cmap = COLORS if color == "BU" else (STATUS_COLORS if color == "Status" else None)
    if chart_t != "Table only":
        with st.container(border=True):
            if chart_t == "Donut":
                gg = g.groupby(row_dim, as_index=False)["Value"].sum()
                cols = [COLORS.get(x) if row_dim == "BU" else STATUS_COLORS.get(x) for x in gg[row_dim]]
                fig = donut(gg[row_dim], gg["Value"], cols if all(cols) else None, f"{metric} by {row_dim}",
                            f"{gg['Value'].sum():,.0f}", metric.lower())
            elif chart_t == "Line":
                fig = px.line(g, x=row_dim, y="Value", color=color, markers=True, color_discrete_map=cmap,
                              category_orders=orders)
            else:
                fig = px.bar(g, x=row_dim, y="Value", color=color,
                             color_discrete_map=cmap or (COLORS if row_dim == "BU" else None),
                             category_orders=orders, barmode="stack" if chart_t == "Stacked bar" else "group")
            if chart_t != "Donut":
                fig.update_layout(title=f"{metric} by {row_dim}" + (f", split by {color}" if color else ""),
                                  yaxis_title=metric)
            if metric == "Win rate" and chart_t != "Donut":
                fig.update_yaxes(tickformat=".0%")
            chart(fig, 420)
    if color:
        tbl = g.pivot_table(index=row_dim, columns=split_dim, values="Value", aggfunc="sum")
    else:
        tbl = g.set_index(row_dim)[["Value"]].rename(columns={"Value": metric})
    if row_dim in orders:
        tbl = tbl.reindex([x for x in orders[row_dim] if x in tbl.index])
    st.dataframe(tbl.style.format("{:.0%}" if metric == "Win rate" else "{:,.0f}", na_rep="–"), width="stretch")

# ----------------------------------------------------------------------------
# Raw data
# ----------------------------------------------------------------------------
with t_data:
    st.caption("Rows from the 'Entries' tab after the sidebar filters. To change anything, edit the Google Sheet "
               "and press Refresh data.")
    show = df[[c for c in df.columns if c in COLUMN_ALIASES]]
    show_df(show, hide_index=True, width="stretch", column_config={
        "Submission Date": st.column_config.DateColumn(format="DD MMM YYYY"),
        "Result Date": st.column_config.DateColumn(format="DD MMM YYYY"),
        "Cost (USD)": st.column_config.NumberColumn(format="localized"),
    })
    st.download_button("Download filtered data (CSV)", show.to_csv(index=False).encode(), "awards_filtered.csv",
                       "text/csv", icon=":material/download:")
