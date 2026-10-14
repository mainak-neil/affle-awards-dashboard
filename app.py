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
            df[c] = pd.to_datetime(df[c].replace("", None), errors="coerce", dayfirst=True)

    # Quarter: manual override > Result Date > Submission Date
    labels = []
    for _, r in df.iterrows():
        if r["Quarter"]:
            labels.append(r["Quarter"].upper().replace("FY ", "FY"))
        elif pd.notna(r["Result Date"]):
            labels.append(quarter_label(r["Result Date"], fy_start)[0])
        elif pd.notna(r["Submission Date"]):
            labels.append(quarter_label(r["Submission Date"], fy_start)[0])
        else:
            labels.append("Unassigned")
    df["Quarter"] = labels

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
    for part in str(text or "").split(","):
        part = part.strip()
        if not part:
            continue
        m = re.match(r"^(.*?)\s*[–—-]\s*(\d+)\s*$", part)
        out.append((m.group(1).strip(), int(m.group(2))) if m else (part, 1))
    return out


# ----------------------------------------------------------------------------
# Styling
# ----------------------------------------------------------------------------
st.markdown(
    """
<style>
  .block-container {padding-top: 2rem; max-width: 1280px;}
  h1 {font-weight: 800 !important; letter-spacing: -0.015em;}
  .kpi {background:#fff; border:1px solid #e9eaec; border-top:3px solid var(--c);
        border-radius:4px; padding:14px 16px 12px; min-height:112px;}
  .kpi .l {font-size:12px; color:#6b7280; font-weight:600; margin-bottom:6px;}
  .kpi .n {font-size:26px; font-weight:700; color:#0f1115; line-height:1.15;}
  .kpi .s {font-size:12px; color:#8a8f98; margin-top:2px;}
  .section {font-size:19px; font-weight:800; border-bottom:2px solid #0f1115;
            padding-bottom:6px; margin:18px 0 12px;}
  .chip {display:inline-flex; align-items:center; background:#152a4a; border-radius:5px;
         padding:4px 10px; height:28px; margin:0 6px 6px 0;}
  .chip img {max-height:16px; max-width:80px;}
  .chip span {color:#fff; font-size:12px; font-weight:600;}
  div[data-testid="stSidebar"] {background:#fff;}
</style>
""",
    unsafe_allow_html=True,
)


def kpi(col, label, value, sub="", color="#0ea99b"):
    col.markdown(
        f'<div class="kpi" style="--c:{color}"><div class="l">{label}</div>'
        f'<div class="n">{value}</div><div class="s">{sub}</div></div>',
        unsafe_allow_html=True,
    )


def show_df(data, container=None, **kw):
    """st.dataframe with money/ratio floats rounded to 2 dp."""
    if isinstance(data, pd.DataFrame):
        num = data.select_dtypes("number").columns
        data = data.copy()
        data[num] = data[num].round(2)
    (container or st).dataframe(data, **kw)


def section(title):
    st.markdown(f'<div class="section">{title}</div>', unsafe_allow_html=True)


def style_fig(fig, height=340):
    fig.update_layout(
        height=height, margin=dict(l=10, r=10, t=40, b=10), plot_bgcolor="white",
        paper_bgcolor="white", font=dict(family="-apple-system, Segoe UI, Arial", size=12),
        legend=dict(orientation="h", yanchor="top", y=-0.12, x=0, title=None),
        title=dict(font=dict(size=14)),
    )
    fig.update_xaxes(showgrid=False, title=None)
    fig.update_yaxes(gridcolor="#eef0f2")
    fig.update_traces(textangle=0, textposition="outside", cliponaxis=False,
                      selector=lambda t: t.type == "bar" and t.text is not None and t.orientation != "h")
    return fig


# ----------------------------------------------------------------------------
# Sidebar: data source + filters
# ----------------------------------------------------------------------------
default_sheet = secret("sheet_url", "")
with st.sidebar:
    st.subheader("Data source")
    if default_sheet:
        source = default_sheet
        st.caption("Connected to the Google Sheet set in app secrets.")
    else:
        typed = st.text_input("Google Sheet link", placeholder="https://docs.google.com/spreadsheets/d/…",
                              help="Share the sheet as 'Anyone with the link – Viewer', or set sheet_url in secrets.")
        source = typed.strip() or "__local__"
        if source == "__local__":
            st.caption("Showing the bundled sample file. Paste your sheet link above, or set it in secrets.")
    if st.button("🔄 Refresh data", width="stretch"):
        load_workbook.clear()
        st.rerun()

try:
    tabs, loaded_at = load_workbook(source)
    settings = read_settings(tabs)
    entries_raw = find_tab(tabs, "Entries")
    if entries_raw is None:
        raise ValueError("No tab named 'Entries' found in the sheet.")
    df_all, data_warnings = clean_entries(entries_raw, settings["fy_start"])
except Exception as e:  # show a friendly error rather than a stack trace
    st.error(f"**Couldn't load the data.** {e}")
    st.stop()

quarters_all = sorted(df_all["Quarter"].unique(), key=quarter_sort_key)
bus_all = sorted(df_all["BU"].unique(), key=str.lower)
COLORS = bu_color_map(bus_all)

with st.sidebar:
    st.caption(f"Last loaded: {loaded_at} · auto-refresh every {CACHE_TTL_SECONDS // 60} min")
    st.divider()
    st.subheader("Filters")
    f_quarter = st.multiselect("Quarter", quarters_all, placeholder="All quarters")
    f_bu = st.multiselect("Business unit", bus_all, placeholder="All BUs")
    f_status = st.multiselect("Status", ["Declared", "Pending", "Upcoming"], placeholder="All statuses")
    f_award = st.multiselect("Award", sorted(df_all["Award"].unique()), placeholder="All awards")
    st.divider()
    currency = st.radio("Currency", ["US$", "₹"], horizontal=True)

df = df_all.copy()
if f_quarter:
    df = df[df["Quarter"].isin(f_quarter)]
if f_bu:
    df = df[df["BU"].isin(f_bu)]
if f_status:
    df = df[df["Status"].isin(f_status)]
if f_award:
    df = df[df["Award"].isin(f_award)]

FX = settings["fx"] if currency == "₹" else 1.0
SYM = "₹" if currency == "₹" else "$"


def money(v: float, decimals: int = 0) -> str:
    return f"{SYM}{v * FX:,.{decimals}f}"


# ----------------------------------------------------------------------------
# Header + KPIs
# ----------------------------------------------------------------------------
head_l, head_r = st.columns([0.75, 0.25])
with head_l:
    affle_logo = (ASSETS / "affle.png")
    if affle_logo.exists():
        st.image(str(affle_logo), width=64)
    st.title(settings["title"])
    st.caption("Award submissions, wins and cost by business unit.")
with head_r:
    st.markdown(
        f"<div style='text-align:right;color:#8a8f98;font-size:13px;padding-top:70px'>"
        f"As of {datetime.now():%d %b %Y}<br>FX: ₹{settings['fx']:.2f} / US$1</div>",
        unsafe_allow_html=True,
    )

if df.empty:
    st.info("No rows match the current filters.")
    st.stop()

declared = df[df["Status"] == "Declared"]
total_subs = int(df["Submissions"].sum())
declared_subs = int(declared["Submissions"].sum())
won = int(df["Won"].sum())
spend = float(df["Cost (USD)"].sum())
declared_spend = float(declared["Cost (USD)"].sum())
win_rate = won / declared_subs if declared_subs else 0
cost_per_win = declared_spend / won if won else 0

k = st.columns(6)
kpi(k[0], "Total submissions", f"{total_subs:,}", f"{df['Award'].nunique()} award programs", "#0ea99b")
kpi(k[1], "Results declared", f"{declared_subs:,}", f"{declared['Award'].nunique()} programs", "#ff7a4a")
kpi(k[2], "Awards won", f"{won:,}", f"{win_rate:.0%} win rate on declared", "#7c6ff0")
kpi(k[3], "Pending results", f"{int(df.loc[df.Status == 'Pending', 'Submissions'].sum()):,}",
    f"{df.loc[df.Status == 'Pending', 'Award'].nunique()} programs", "#f2a53c")
kpi(k[4], "Total spend", money(spend), "all entries with cost", "#3b8fd6")
kpi(k[5], "Cost / award won", money(cost_per_win, 2), "declared spend ÷ wins", "#ef5da8")

if data_warnings:
    with st.expander(f"⚠️ {len(data_warnings)} data note(s) — fix these in the Google Sheet"):
        for w in data_warnings:
            st.markdown(f"- {w}")

st.write("")
tab_overview, tab_q, tab_bu, tab_award, tab_pipe, tab_cs, tab_custom, tab_data = st.tabs(
    ["Overview", "By quarter", "By BU", "By award", "Pending & upcoming", "Case studies", "Build your own view", "Data"]
)

# ----------------------------------------------------------------------------
# Overview
# ----------------------------------------------------------------------------
with tab_overview:
    status_sum = df.groupby("Status")["Submissions"].sum()
    fig = go.Figure()
    for stt in ["Declared", "Pending", "Upcoming"]:
        v = int(status_sum.get(stt, 0))
        if v:
            fig.add_bar(y=["Status"], x=[v], name=f"{stt} ({v})", orientation="h",
                        marker_color=STATUS_COLORS[stt], text=[f"{stt} · {v}"],
                        textposition="inside", insidetextanchor="middle", textfont_color="white")
    fig.update_layout(barmode="stack", title="Submission status", bargap=0.1)
    fig.update_yaxes(visible=False)
    fig.update_xaxes(visible=False)
    st.plotly_chart(style_fig(fig, 150).update_layout(showlegend=False, margin=dict(t=36, b=0)),
                    width="stretch")

    by_bu = df.groupby("BU", as_index=False)[["Submissions", "Won", "Cost (USD)"]].sum()
    by_bu["Spend"] = by_bu["Cost (USD)"] * FX
    c1, c2 = st.columns(2)
    fig = px.bar(by_bu.sort_values("Won", ascending=False), x="BU", y="Won", color="BU",
                 color_discrete_map=COLORS, text="Won", title="Awards won per BU")
    c1.plotly_chart(style_fig(fig).update_layout(showlegend=False), width="stretch")
    fig = px.bar(by_bu.sort_values("Spend", ascending=False), x="BU", y="Spend", color="BU",
                 color_discrete_map=COLORS, text_auto=",.0f", title=f"Spend per BU ({currency})")
    c2.plotly_chart(style_fig(fig).update_layout(showlegend=False), width="stretch")

    c3, c4 = st.columns(2)
    q = df.groupby(["Quarter", "BU"], as_index=False)["Won"].sum()
    fig = px.bar(q, x="Quarter", y="Won", color="BU", color_discrete_map=COLORS,
                 category_orders={"Quarter": quarters_all}, title="Awards won per quarter")
    c3.plotly_chart(style_fig(fig), width="stretch")
    fig = px.bar(by_bu.sort_values("Submissions", ascending=False), x="BU", y="Submissions", color="BU",
                 color_discrete_map=COLORS, text="Submissions", title="Submissions per BU")
    c4.plotly_chart(style_fig(fig).update_layout(showlegend=False), width="stretch")

# ----------------------------------------------------------------------------
# By quarter
# ----------------------------------------------------------------------------
with tab_q:
    section("Quarter-wise summary")
    qs = df.groupby("Quarter").agg(
        Programs=("Award", "nunique"), Submissions=("Submissions", "sum"), Won=("Won", "sum"),
        Spend=("Cost (USD)", "sum"),
    ).reindex([x for x in quarters_all if x in df["Quarter"].unique()]).reset_index()
    qs["Spend"] = qs["Spend"] * FX
    qs["Cost / win"] = (qs["Spend"] / qs["Won"]).where(qs["Won"] > 0)
    show_df(qs, hide_index=True, width="stretch", column_config={
        "Spend": st.column_config.NumberColumn(f"Spend ({currency})", format="localized"),
        "Cost / win": st.column_config.NumberColumn(f"Cost / win ({currency})", format="localized"),
    })

    metric_q = st.selectbox("Show", ["Won", "Submissions", "Spend"], key="qmetric")
    val_col = "Cost (USD)" if metric_q == "Spend" else metric_q
    piv = df.pivot_table(index="BU", columns="Quarter", values=val_col, aggfunc="sum", fill_value=0)
    piv = piv[[c for c in quarters_all if c in piv.columns]] * (FX if metric_q == "Spend" else 1)
    piv["Total"] = piv.sum(axis=1)
    piv.loc["Total"] = piv.sum()
    section(f"{metric_q} — BU × quarter")
    piv = piv.sort_index(key=lambda i: [(x == "Total", str(x).lower()) for x in i])
    show_df(piv.round(0), width="stretch")

    trend = piv.drop(index="Total", columns=["Total", "Unassigned"], errors="ignore")
    trend = trend.reset_index().melt(id_vars="BU", var_name="Quarter", value_name=metric_q)
    fig = px.line(trend, x="Quarter", y=metric_q, color="BU", markers=True, color_discrete_map=COLORS,
                  category_orders={"Quarter": quarters_all}, title=f"{metric_q} trend by BU (dated rows only)")
    st.plotly_chart(style_fig(fig), width="stretch")

# ----------------------------------------------------------------------------
# By BU
# ----------------------------------------------------------------------------
with tab_bu:
    section("Spend & submissions by BU")
    chips = ""
    for bu in sorted(df["BU"].unique(), key=str.lower):
        uri = logo_data_uri(bu)
        chips += f'<span class="chip">{f"<img src={chr(34)}{uri}{chr(34)}>" if uri else f"<span>{bu}</span>"}</span>'
    st.markdown(chips, unsafe_allow_html=True)

    t = df.groupby("BU").agg(
        Submissions=("Submissions", "sum"), Declared=("Submissions", lambda s: s[df.loc[s.index, "Status"] == "Declared"].sum()),
        Won=("Won", "sum"), Spend=("Cost (USD)", "sum"), Programs=("Award", "nunique"),
    ).reset_index()
    t["Win rate"] = (t["Won"] / t["Declared"]).where(t["Declared"] > 0)
    t["% of spend"] = t["Spend"] / t["Spend"].sum() if t["Spend"].sum() else 0
    t["Cost / submission"] = (t["Spend"] / t["Submissions"]).where(t["Submissions"] > 0) * FX
    t["Cost / win"] = (t["Spend"] / t["Won"]).where(t["Won"] > 0) * FX
    t["Spend"] = t["Spend"] * FX
    t = t.sort_values("Spend", ascending=False)
    show_df(t, hide_index=True, width="stretch", column_config={
        "Spend": st.column_config.NumberColumn(f"Spend ({currency})", format="localized"),
        "% of spend": st.column_config.ProgressColumn("% of spend", format="percent", min_value=0, max_value=1),
        "Win rate": st.column_config.ProgressColumn("Win rate", format="percent", min_value=0, max_value=1),
        "Cost / submission": st.column_config.NumberColumn(f"Cost / submission ({currency})", format="localized"),
        "Cost / win": st.column_config.NumberColumn(f"Cost / win ({currency})", format="localized"),
    })
    st.caption(f"Totals: {int(t['Submissions'].sum())} submissions · {int(t['Won'].sum())} won · "
               f"{money(t['Spend'].sum() / FX, 2)} spend")

# ----------------------------------------------------------------------------
# By award
# ----------------------------------------------------------------------------
with tab_award:
    section("Cost / awards won (declared awards)")
    d = df[df["Status"] == "Declared"]
    if d.empty:
        st.info("No declared awards in the current filter.")
    else:
        aw = d.groupby("Award").agg(
            Quarter=("Quarter", lambda s: ", ".join(sorted(set(s), key=quarter_sort_key))),
            Submissions=("Submissions", "sum"), Won=("Won", "sum"), Cost=("Cost (USD)", "sum"),
        ).reset_index()
        split = d.groupby("Award").apply(
            lambda g: ", ".join(f"{b} – {int(s)}" for b, s in zip(g["BU"], g["Submissions"])), include_groups=False
        )
        aw.insert(2, "BU split", aw["Award"].map(split))
        aw["Win rate"] = (aw["Won"] / aw["Submissions"]).where(aw["Submissions"] > 0)
        aw["Cost / win"] = (aw["Cost"] / aw["Won"]).where(aw["Won"] > 0) * FX
        aw["Cost"] = aw["Cost"] * FX
        show_df(aw.sort_values("Won", ascending=False), hide_index=True, width="stretch", column_config={
            "Award": st.column_config.TextColumn(width="large"),
            "Cost": st.column_config.NumberColumn(f"Cost ({currency})", format="localized"),
            "Cost / win": st.column_config.NumberColumn(f"Cost / win ({currency})", format="localized"),
            "Win rate": st.column_config.ProgressColumn("Win rate", format="percent", min_value=0, max_value=1),
        })
        st.caption(f"Total: {int(aw['Submissions'].sum())} submissions · {int(aw['Won'].sum())} won · "
                   f"{SYM}{aw['Cost'].sum():,.2f} · cost per award won {SYM}{(aw['Cost'].sum() / aw['Won'].sum() if aw['Won'].sum() else 0):,.2f}")

        fig = px.bar(d.groupby(["Award", "BU"], as_index=False)["Won"].sum(), y="Award", x="Won", color="BU",
                     orientation="h", color_discrete_map=COLORS, title="Wins per award, split by BU")
        fig.update_yaxes(categoryorder="total ascending", title=None)
        st.plotly_chart(style_fig(fig, max(320, 42 * d["Award"].nunique())), width="stretch")

# ----------------------------------------------------------------------------
# Pending & upcoming
# ----------------------------------------------------------------------------
with tab_pipe:
    for status, title in [("Pending", "Awards submitted — not yet declared"), ("Upcoming", "Upcoming award submissions")]:
        section(title)
        p = df[df["Status"] == status]
        if p.empty:
            st.caption("None in the current filter.")
            continue
        pt = p.groupby("Award").agg(
            Date=("Submission Date", "min"), Submissions=("Submissions", "sum"),
            BUs=("BU", lambda s: ", ".join(f"{b} – {int(n)}" if n else b for b, n in zip(s, p.loc[s.index, "Submissions"]))),
            Notes=("Notes", lambda s: "; ".join(sorted({x for x in s if x}))),
        ).reset_index().rename(columns={"Date": "Submission / due date"})
        show_df(pt, hide_index=True, width="stretch", column_config={
            "Submission / due date": st.column_config.DateColumn(format="DD MMM YYYY"),
            "Award": st.column_config.TextColumn(width="large"),
        })

# ----------------------------------------------------------------------------
# Case studies
# ----------------------------------------------------------------------------
with tab_cs:
    recs = []
    for _, r in df.iterrows():
        for name, n in parse_case_studies(r["Case Studies Submitted"]):
            recs.append({"BU": r["BU"], "Award": r["Award"], "Case study": name, "Submitted": n, "Won": 0})
        for name, n in parse_case_studies(r["Case Studies Won"]):
            recs.append({"BU": r["BU"], "Award": r["Award"], "Case study": name, "Submitted": 0, "Won": n})
    if not recs:
        st.info("No case studies listed for the current filter. Fill the 'Case Studies Submitted / Won' columns in the sheet.")
    else:
        cs = pd.DataFrame(recs)
        section("Platform / case-study wise")
        cs_bu = cs.groupby(["BU", "Case study"], as_index=False)[["Submitted", "Won"]].sum()
        cs_bu = cs_bu.sort_values(["BU", "Submitted"], ascending=[True, False])
        c1, c2 = st.columns([0.45, 0.55])
        show_df(cs_bu, c1, hide_index=True, width="stretch", height=420)
        fig = px.bar(cs_bu.sort_values("Won"), y="Case study", x=["Submitted", "Won"], barmode="group",
                     orientation="h", title="Submitted vs won by case study",
                     color_discrete_sequence=["#c9ced6", "#0ea99b"])
        fig.update_yaxes(title=None)
        c2.plotly_chart(style_fig(fig, max(420, 22 * cs_bu["Case study"].nunique())), width="stretch")

        section("Award-wise case studies submitted / won")
        aw_cs = df[(df["Case Studies Submitted"] != "") | (df["Case Studies Won"] != "")][
            ["BU", "Award", "Submissions", "Case Studies Submitted", "Case Studies Won"]]
        show_df(aw_cs, hide_index=True, width="stretch")

# ----------------------------------------------------------------------------
# Build your own view (pivot builder)
# ----------------------------------------------------------------------------
with tab_custom:
    st.caption("Pick any combination — the chart and table update instantly. Uses the sidebar filters too.")
    dims = ["BU", "Quarter", "Award", "Status"]
    metrics = {"Awards won": "Won", "Submissions": "Submissions", f"Spend ({currency})": "Cost (USD)",
               "Win rate": None, f"Cost per win ({currency})": None, "Award programs": None}
    c1, c2, c3, c4 = st.columns(4)
    row_dim = c1.selectbox("Group by", dims, index=1)
    split_dim = c2.selectbox("Split by", ["(none)"] + [d_ for d_ in dims if d_ != row_dim])
    metric = c3.selectbox("Metric", list(metrics))
    chart = c4.selectbox("Chart", ["Bar", "Stacked bar", "Line", "Pie", "Table only"])

    keys = [row_dim] + ([] if split_dim == "(none)" else [split_dim])
    g = df.groupby(keys).agg(Won=("Won", "sum"), Submissions=("Submissions", "sum"),
                             Cost=("Cost (USD)", "sum"), Programs=("Award", "nunique")).reset_index()
    decl = df[df["Status"] == "Declared"].groupby(keys).agg(
        DSubs=("Submissions", "sum"), DCost=("Cost (USD)", "sum")).reset_index()
    g = g.merge(decl, on=keys, how="left").fillna({"DSubs": 0, "DCost": 0})
    if metric == "Awards won":
        g["Value"] = g["Won"]
    elif metric == "Submissions":
        g["Value"] = g["Submissions"]
    elif metric.startswith("Spend"):
        g["Value"] = g["Cost"] * FX
    elif metric == "Win rate":
        g["Value"] = (g["Won"] / g["DSubs"]).where(g["DSubs"] > 0)
    elif metric.startswith("Cost per win"):
        g["Value"] = (g["DCost"] / g["Won"]).where(g["Won"] > 0) * FX
    else:
        g["Value"] = g["Programs"]

    orders = {"Quarter": quarters_all, "Status": ["Declared", "Pending", "Upcoming"]}
    color = None if split_dim == "(none)" else split_dim
    cmap = COLORS if color == "BU" else (STATUS_COLORS if color == "Status" else None)
    if chart != "Table only":
        if chart == "Pie":
            fig = px.pie(g.groupby(row_dim, as_index=False)["Value"].sum(), names=row_dim, values="Value",
                         hole=0.45, color=row_dim, color_discrete_map=COLORS if row_dim == "BU" else None)
        elif chart == "Line":
            fig = px.line(g, x=row_dim, y="Value", color=color, markers=True, color_discrete_map=cmap,
                          category_orders=orders)
        else:
            fig = px.bar(g, x=row_dim, y="Value", color=color, color_discrete_map=cmap or
                         (COLORS if row_dim == "BU" else None), category_orders=orders,
                         barmode="stack" if chart == "Stacked bar" else "group")
        fig.update_layout(title=f"{metric} by {row_dim}" + ("" if not color else f", split by {color}"),
                          yaxis_title=metric)
        if metric == "Win rate":
            fig.update_yaxes(tickformat=".0%")
        st.plotly_chart(style_fig(fig, 420), width="stretch")

    if color:
        tbl = g.pivot_table(index=row_dim, columns=split_dim, values="Value", aggfunc="sum")
        if row_dim in orders:
            tbl = tbl.reindex([x for x in orders[row_dim] if x in tbl.index])
    else:
        tbl = g.set_index(row_dim)[["Value"]].rename(columns={"Value": metric})
        if row_dim in orders:
            tbl = tbl.reindex([x for x in orders[row_dim] if x in tbl.index])
    st.dataframe(tbl.style.format("{:.0%}" if metric == "Win rate" else "{:,.0f}", na_rep="–"), width="stretch")

# ----------------------------------------------------------------------------
# Raw data
# ----------------------------------------------------------------------------
with tab_data:
    st.caption("These are the rows from the 'Entries' tab after applying the sidebar filters. "
               "To change anything, edit the Google Sheet and press Refresh data.")
    show = df.drop(columns=[c for c in df.columns if c not in COLUMN_ALIASES])
    show_df(show, hide_index=True, width="stretch", column_config={
        "Submission Date": st.column_config.DateColumn(format="DD MMM YYYY"),
        "Result Date": st.column_config.DateColumn(format="DD MMM YYYY"),
        "Cost (USD)": st.column_config.NumberColumn(format="localized"),
    })
    st.download_button("⬇️ Download filtered data (CSV)", show.to_csv(index=False).encode(),
                       "awards_filtered.csv", "text/csv")
