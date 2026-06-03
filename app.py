"""
Dashboard Đăng Ký Khám Online — BVĐK Tâm Đức Cầu Quan
Fully responsive: PC + Mobile
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import gspread
from google.oauth2.service_account import Credentials
import json, os
from datetime import datetime

# ── CONFIG ────────────────────────────────────
SHEET_ID   = "1EYiRA3ar41aue8DlbWA7JTKoLL0M2tiLTcZINhdMfTs"
SHEET_NAME = "Câu trả lời biểu mẫu 1"

COL_TIMESTAMP = "THỜI GIAN ĐĂNG KÝ"
COL_STATUS    = "TRẠNG THÁI"
COL_EXAM_DATE = "NGÀY KHÁM"
COL_NAME      = "1. HỌ VÀ TÊN BỆNH NHÂN"
COL_GENDER    = "3. GIỚI TÍNH"
COL_SPECIALTY = "CHUYÊN KHOA MONG MUỐN KHÁM"
COL_DOCTOR    = "BÁC SĨ MONG MUỐN ( nếu có)"

STATUS_ATTENDED = "BỆNH NHÂN ĐÃ KHÁM"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

# ── PAGE CONFIG ───────────────────────────────
st.set_page_config(
    page_title="Dashboard · Tâm Đức Cầu Quan",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="collapsed",   # collapsed by default = better on mobile
)

# ── COLORS ────────────────────────────────────
C_BLUE   = "#3b82f6"
C_GREEN  = "#10b981"
C_ROSE   = "#f43f5e"
C_VIOLET = "#8b5cf6"
C_TEAL   = "#14b8a6"
C_AMBER  = "#f59e0b"
C_SLATE  = "#94a3b8"

# ── CSS — FULLY RESPONSIVE ───────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:ital,wght@0,300;0,400;0,500;0,600;0,700;1,400&family=JetBrains+Mono:wght@500;600&display=swap');

*, html, body, [class*="css"] {
    font-family: 'Inter', sans-serif !important;
    box-sizing: border-box;
}

/* ══ LAYOUT ══ */
.stApp { background: #f1f5f9; }

.main .block-container {
    padding: 1rem 1rem 3rem 1rem !important;
    max-width: 1440px;
}

/* Desktop padding */
@media (min-width: 768px) {
    .main .block-container {
        padding: 1.5rem 2rem 3rem 2rem !important;
    }
}

/* ══ SIDEBAR ══ */
section[data-testid="stSidebar"] {
    background: linear-gradient(175deg, #0f172a 0%, #1e3a5f 100%) !important;
}
section[data-testid="stSidebar"] * { color: #94b8d4 !important; }
section[data-testid="stSidebar"] strong,
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 { color: #e0f0ff !important; }
section[data-testid="stSidebar"] hr { border-color: #1e3a5f !important; }

/* ══ BUTTON ══ */
.stButton > button {
    background: linear-gradient(135deg, #3b82f6, #2563eb) !important;
    color: white !important; border: none !important;
    border-radius: 12px !important; font-weight: 600 !important;
    padding: 0.65rem 1rem !important; width: 100% !important;
    font-size: 0.9rem !important;
    box-shadow: 0 4px 12px rgba(59,130,246,0.3) !important;
    transition: all 0.2s !important;
}
.stButton > button:hover {
    background: linear-gradient(135deg, #60a5fa, #3b82f6) !important;
    box-shadow: 0 6px 18px rgba(59,130,246,0.4) !important;
    transform: translateY(-1px) !important;
}

/* ══ HEADER ══ */
.app-header {
    background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 50%, #0e4d7a 100%);
    border-radius: 20px;
    padding: 1.2rem 1.4rem;
    margin-bottom: 1.2rem;
    box-shadow: 0 4px 24px rgba(15,23,42,0.2);
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 0.8rem;
}
.app-header-left { display: flex; align-items: center; gap: 0.8rem; flex: 1; min-width: 0; }
.app-header-emoji {
    font-size: 2rem;
    background: rgba(255,255,255,0.1);
    border-radius: 12px;
    padding: 0.4rem 0.55rem;
    line-height: 1; flex-shrink: 0;
}
.app-header-title {
    font-size: 1.1rem; font-weight: 700;
    color: #f0f9ff; line-height: 1.25; margin: 0;
}
.app-header-sub {
    font-size: 0.75rem; color: #7dd3fc;
    margin-top: 0.15rem; line-height: 1.4;
}
.app-header-time {
    background: rgba(59,130,246,0.2);
    border: 1px solid rgba(59,130,246,0.35);
    border-radius: 20px; padding: 0.3rem 0.8rem;
    color: #93c5fd; font-size: 0.72rem; font-weight: 500;
    white-space: nowrap; flex-shrink: 0;
}

@media (min-width: 768px) {
    .app-header { padding: 1.6rem 2rem; }
    .app-header-title { font-size: 1.4rem; }
    .app-header-sub { font-size: 0.82rem; }
    .app-header-emoji { font-size: 2.4rem; }
}

/* ══ KPI CARDS — Responsive Grid ══ */
.kpi-grid {
    display: grid;
    grid-template-columns: repeat(2, 1fr);   /* 2 cols on mobile */
    gap: 0.75rem;
    margin-bottom: 1.2rem;
}
@media (min-width: 900px) {
    .kpi-grid { grid-template-columns: repeat(4, 1fr); gap: 1rem; }
}

.kpi-card {
    background: white;
    border-radius: 16px;
    padding: 1rem 1.1rem;
    border: 1px solid #e2e8f0;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05), 0 4px 16px rgba(0,0,0,0.04);
    position: relative; overflow: hidden;
    transition: transform 0.15s, box-shadow 0.15s;
}
.kpi-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(0,0,0,0.08), 0 8px 24px rgba(0,0,0,0.05);
}
.kpi-card::after {
    content: '';
    position: absolute; top: 0; left: 0; right: 0; height: 3px;
    border-radius: 16px 16px 0 0;
}
.kc-blue::after   { background: linear-gradient(90deg,#3b82f6,#60a5fa); }
.kc-green::after  { background: linear-gradient(90deg,#10b981,#34d399); }
.kc-rose::after   { background: linear-gradient(90deg,#f43f5e,#fb7185); }
.kc-violet::after { background: linear-gradient(90deg,#8b5cf6,#a78bfa); }

.kpi-bg-icon {
    position: absolute; bottom: -0.4rem; right: 0.5rem;
    font-size: 3rem; opacity: 0.07; line-height: 1;
    pointer-events: none;
}
.kpi-label {
    font-size: 0.65rem; font-weight: 700;
    text-transform: uppercase; letter-spacing: 0.09em;
    color: #94a3b8; margin-bottom: 0.4rem;
}
.kpi-num {
    font-size: 2rem; font-weight: 700; color: #0f172a;
    font-family: 'JetBrains Mono', monospace !important;
    line-height: 1;
}
.kpi-sub {
    font-size: 0.7rem; color: #94a3b8;
    margin-top: 0.3rem; font-weight: 500;
}

@media (min-width: 768px) {
    .kpi-card { padding: 1.3rem 1.5rem; }
    .kpi-num  { font-size: 2.4rem; }
    .kpi-label { font-size: 0.7rem; }
}

/* ══ CHART CARDS ══ */
.chart-wrap {
    background: white;
    border-radius: 16px;
    padding: 1rem 0.8rem 0.5rem;
    border: 1px solid #e2e8f0;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05), 0 4px 16px rgba(0,0,0,0.04);
    margin-bottom: 1rem;
}
@media (min-width: 768px) {
    .chart-wrap { padding: 1.3rem 1.2rem 0.6rem; }
}

/* ══ SECTION HEADER ══ */
.sec-hdr {
    display: flex; align-items: center; gap: 0.55rem;
    margin: 1.2rem 0 0.7rem;
    padding-bottom: 0.5rem;
    border-bottom: 1.5px solid #e2e8f0;
}
.sec-hdr-dot {
    width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0;
}
.sec-hdr-text {
    font-size: 0.88rem; font-weight: 700;
    color: #1e293b; letter-spacing: 0.01em;
}
@media (min-width: 768px) {
    .sec-hdr-text { font-size: 0.95rem; }
}

/* ══ LEGEND ROW ══ */
.legend-row {
    display: flex; flex-wrap: wrap;
    gap: 0.8rem 1.5rem;
    justify-content: center;
    padding: 0.4rem 0 0.8rem;
}
.legend-item {
    display: flex; align-items: center; gap: 0.4rem;
    font-size: 0.78rem; color: #475569; font-weight: 500;
}
.legend-dot {
    width: 10px; height: 10px;
    border-radius: 3px; flex-shrink: 0;
}

/* ══ SIDEBAR LOGO ══ */
.sb-logo {
    background: rgba(59,130,246,0.1);
    border: 1px solid rgba(59,130,246,0.2);
    border-radius: 14px; padding: 1rem 1.1rem;
    text-align: center; margin-bottom: 0.8rem;
}
.sb-logo-title { font-size: 0.9rem !important; font-weight: 700 !important; color: #e0f0ff !important; }
.sb-logo-sub   { font-size: 0.7rem !important; color: #5b92b5 !important; margin-top: 0.2rem; }

/* ══ EMPTY STATE ══ */
.empty-wrap {
    text-align: center; padding: 4rem 1.5rem;
    background: white; border-radius: 20px;
    border: 2px dashed #e2e8f0; margin-top: 1.5rem;
}
.empty-icon  { font-size: 3rem; margin-bottom: 0.8rem; }
.empty-title { font-size: 1.1rem; font-weight: 700; color: #1e293b; margin-bottom: 0.5rem; }
.empty-desc  { font-size: 0.85rem; color: #94a3b8; max-width: 340px; margin: 0 auto; line-height: 1.6; }

/* Fix Streamlit default padding on mobile */
@media (max-width: 768px) {
    .main .block-container { padding: 0.6rem 0.6rem 2rem !important; }
    div[data-testid="stHorizontalBlock"] > div { min-width: 0 !important; }
}

/* Dataframe rounded */
div[data-testid="stDataFrame"] > div { border-radius: 12px; overflow: hidden; }
div[data-testid="stExpander"] { border-radius: 12px !important; border: 1px solid #e2e8f0 !important; }
</style>
""", unsafe_allow_html=True)


# ── CHART HELPERS ─────────────────────────────
BASE = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter, sans-serif", color="#64748b"),
    hoverlabel=dict(bgcolor="#0f172a", font_color="#f1f5f9",
                    font_size=12, bordercolor="#334155"),
)

def chart_donut(m):
    fig = go.Figure(go.Pie(
        labels=["Đã khám", "Chưa / Vắng"],
        values=[m["attended_count"], m["noshow_count"]],
        hole=0.68,
        marker=dict(colors=[C_GREEN, C_ROSE], line=dict(color="#fff", width=3)),
        textinfo="percent",
        textfont=dict(size=13, family="Inter"),
        hovertemplate="<b>%{label}</b><br>%{value} người · %{percent}<extra></extra>",
        pull=[0.025, 0.025],
        direction="clockwise",
    ))
    fig.update_layout(
        **BASE, height=300, showlegend=False,
        margin=dict(t=10, b=10, l=20, r=20),
        annotations=[dict(
            text=f"<b>{m['total']}</b><br><span style='font-size:11px;color:#94a3b8'>bệnh nhân</span>",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=24, family="Inter", color="#0f172a"),
        )],
    )
    return fig

def chart_gender(m):
    df = m["gen"]
    if df is None or df.empty: return None
    clr = {"NAM": C_BLUE, "NỮ": "#f472b6", "NU": "#f472b6"}
    colors = [clr.get(g.upper(), C_SLATE) for g in df["Giới tính"]]
    fig = go.Figure(go.Bar(
        x=df["Giới tính"], y=df["Số lượng"],
        marker=dict(color=colors, line=dict(color="#fff", width=2)),
        text=df["Số lượng"], textposition="outside",
        textfont=dict(size=13, family="JetBrains Mono", color="#0f172a"),
        hovertemplate="<b>%{x}</b> — %{y} người<extra></extra>",
        width=0.5,
    ))
    fig.update_layout(
        **BASE, height=240,
        margin=dict(t=10, b=10, l=10, r=10),
        xaxis=dict(showgrid=False, tickfont=dict(size=12, color="#475569")),
        yaxis=dict(gridcolor="#f8fafc", zeroline=False, tickfont=dict(size=10)),
    )
    return fig

def chart_daily(m):
    df = m["daily"]
    if df is None or df.empty: return None
    n = len(df)
    alphas = [0.4 + 0.6 * i / max(n-1, 1) for i in range(n)]
    colors = [f"rgba(59,130,246,{a:.2f})" for a in alphas]
    fig = go.Figure(go.Bar(
        x=df["Ngày khám"], y=df["Lịch hẹn"],
        marker=dict(color=colors, line=dict(color="#fff", width=1.5)),
        text=df["Lịch hẹn"], textposition="outside",
        textfont=dict(size=10, family="JetBrains Mono", color="#0f172a"),
        hovertemplate="<b>%{x}</b><br>%{y} lịch hẹn<extra></extra>",
    ))
    fig.update_layout(
        **BASE, height=260,
        margin=dict(t=10, b=10, l=10, r=10),
        xaxis=dict(tickangle=-35, tickfont=dict(size=9, color="#64748b"), showgrid=False),
        yaxis=dict(gridcolor="#f8fafc", zeroline=False, tickfont=dict(size=9)),
        bargap=0.3,
    )
    return fig

def chart_specialty(m):
    df = m["spec"]
    if df is None or df.empty: return None
    df = df.copy().sort_values("Số lượng")
    df["label"] = df["Chuyên khoa"].apply(
        lambda x: (x[:32] + "…") if len(x) > 32 else x
    )
    palette = [C_TEAL, C_BLUE, C_VIOLET, C_AMBER, C_GREEN, "#06b6d4", "#6366f1", "#f97316"]
    fig = go.Figure(go.Bar(
        y=df["label"], x=df["Số lượng"], orientation="h",
        marker=dict(color=palette[:len(df)], line=dict(color="#fff", width=1.5)),
        text=df["Số lượng"], textposition="outside",
        textfont=dict(size=11, family="JetBrains Mono", color="#0f172a"),
        hovertemplate="<b>%{y}</b><br>%{x} bệnh nhân<extra></extra>",
    ))
    fig.update_layout(
        **BASE, height=max(260, len(df) * 46),
        margin=dict(t=10, b=10, l=10, r=48),
        xaxis=dict(gridcolor="#f8fafc", zeroline=False, tickfont=dict(size=9)),
        yaxis=dict(tickfont=dict(size=9.5, color="#475569"), showgrid=False),
    )
    return fig


# ── DATA FUNCTIONS ────────────────────────────

def authenticate_gspread(src) -> gspread.Client:
    if isinstance(src, str):
        if not os.path.exists(src):
            raise FileNotFoundError(f"Không tìm thấy: {src}")
        creds = Credentials.from_service_account_file(src, scopes=SCOPES)
    elif isinstance(src, dict):
        creds = Credentials.from_service_account_info(src, scopes=SCOPES)
    else:
        raise ValueError("Credentials phải là str hoặc dict.")
    return gspread.authorize(creds)


def fetch_data(client, sheet_id, sheet_name):
    ws = client.open_by_key(sheet_id).worksheet(sheet_name)
    all_vals = ws.get_all_values()
    if not all_vals or len(all_vals) < 2:
        raise ValueError("Sheet trống hoặc không có dữ liệu.")
    headers, rows = all_vals[0], all_vals[1:]
    seen, clean = {}, []
    for i, h in enumerate(headers):
        h = h.strip()
        if h == "": h = f"_col_{i}"
        elif h in seen: seen[h] += 1; h = f"{h}_{seen[h]}"
        else: seen[h] = 0
        clean.append(h)
    df = pd.DataFrame(rows, columns=clean)
    df = df.loc[:, ~df.columns.str.startswith("_col_")]
    df = df.replace("", pd.NA).dropna(how="all").fillna("")
    return df


def process_data(df):
    if COL_STATUS not in df.columns:
        raise KeyError(f"Không thấy cột '{COL_STATUS}'. Có: {list(df.columns)}")
    df = df.copy()
    df[COL_STATUS] = df[COL_STATUS].astype(str).str.strip()
    df = df[~df[COL_STATUS].isin(["", "nan", "N/A", "\u200b"])]
    total = len(df)
    if total == 0:
        return _empty(df)

    attended = int((df[COL_STATUS].str.upper() == STATUS_ATTENDED.upper()).sum())
    noshow   = total - attended

    spec = None
    if COL_SPECIALTY in df.columns:
        s = df[df[COL_SPECIALTY].astype(str).str.strip() != ""][COL_SPECIALTY]
        if not s.empty:
            spec = s.value_counts().head(8).reset_index()
            spec.columns = ["Chuyên khoa", "Số lượng"]

    gen = None
    if COL_GENDER in df.columns:
        g = df[COL_GENDER].astype(str).str.strip()
        g = g[g.str.upper().isin(["NAM", "NỮ", "NU"])]
        if not g.empty:
            gen = g.value_counts().reset_index()
            gen.columns = ["Giới tính", "Số lượng"]

    daily = None
    if COL_EXAM_DATE in df.columns:
        ds = df[COL_EXAM_DATE].astype(str).str.strip()
        ds = ds[ds.str.match(r'\d{2}/\d{2}/\d{4}')]
        if not ds.empty:
            daily = ds.value_counts().sort_index().reset_index()
            daily.columns = ["Ngày khám", "Lịch hẹn"]

    status_tbl = df[COL_STATUS].value_counts().reset_index()
    status_tbl.columns = ["Trạng thái", "Số lượng"]

    return dict(total=total, attended_count=attended, noshow_count=noshow,
                attended_pct=round(attended/total*100,1),
                noshow_pct=round(noshow/total*100,1),
                spec=spec, gen=gen, daily=daily, status_tbl=status_tbl, df=df)


def _empty(df):
    return dict(total=0, attended_count=0, noshow_count=0,
                attended_pct=0.0, noshow_pct=0.0,
                spec=None, gen=None, daily=None, status_tbl=None, df=df)


# ── CREDENTIALS ──────────────────────────────

def get_credentials():
    try:
        if "gcp_service_account" in st.secrets:
            return dict(st.secrets["gcp_service_account"])
    except Exception:
        pass
    for path in [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "credentials.json"),
        os.path.join(os.getcwd(), "credentials.json"),
    ]:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8-sig") as f:
                return json.load(f)
    return None

creds_data = get_credentials()


# ── SIDEBAR ───────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div class="sb-logo">
      <div style="font-size:1.6rem;margin-bottom:0.3rem">🏥</div>
      <div class="sb-logo-title">BVĐK Tâm Đức Cầu Quan</div>
      <div class="sb-logo-sub">Hệ thống theo dõi lịch hẹn</div>
    </div>
    """, unsafe_allow_html=True)

    if creds_data:
        st.success("🔒 Google API: Đã kết nối")
    else:
        st.error("⚠️ Chưa có credentials")
        st.caption("Thêm `gcp_service_account` vào Streamlit Secrets.")

    st.markdown("---")
    st.markdown("**⚡ Tự động cập nhật**")
    refresh_interval = st.selectbox("", [30, 60, 120, 300, 600], index=1,
        format_func=lambda x: f"{x} giây" if x < 60 else f"{x//60} phút",
        label_visibility="collapsed")
    auto_refresh = st.toggle("Bật tự động làm mới", value=True)

    st.markdown("---")
    fetch_btn = st.button("🔄  Tải dữ liệu ngay")

    st.markdown("---")
    st.markdown("**📋 Nguồn dữ liệu**")
    st.code(f"Sheet: ...{SHEET_ID[-10:]}\nTab: {SHEET_NAME[:20]}", language=None)


# ── SESSION STATE ─────────────────────────────
for k, v in [("metrics", None), ("fetch_time", None), ("err", None), ("last_auto", 0)]:
    if k not in st.session_state:
        st.session_state[k] = v


def do_fetch():
    if not creds_data:
        st.session_state.err = "⚠️ Chưa có credentials."
        st.session_state.metrics = None
        return
    try:
        cl = authenticate_gspread(creds_data)
        raw = fetch_data(cl, SHEET_ID, SHEET_NAME)
        st.session_state.metrics    = process_data(raw)
        st.session_state.fetch_time = datetime.now().strftime("%H:%M:%S · %d/%m/%Y")
        st.session_state.last_auto  = datetime.now().timestamp()
        st.session_state.err        = None
    except Exception as e:
        st.session_state.err     = f"❌ {type(e).__name__}: {e}"
        st.session_state.metrics = None


# Initial load
if st.session_state.metrics is None and st.session_state.err is None:
    with st.spinner("Đang tải dữ liệu…"):
        do_fetch()

if fetch_btn:
    with st.spinner("Đang tải dữ liệu…"):
        do_fetch()

# Auto refresh
now_ts = datetime.now().timestamp()
if auto_refresh and creds_data and (now_ts - st.session_state.last_auto) >= refresh_interval and not fetch_btn:
    with st.spinner("Đang cập nhật…"):
        do_fetch()

# Meta refresh (browser-level, no sleep)
if auto_refresh and creds_data and st.session_state.metrics is not None:
    elapsed   = int(datetime.now().timestamp() - st.session_state.last_auto)
    remaining = max(0, refresh_interval - elapsed)
    st.sidebar.caption(f"🔁 Cập nhật sau {remaining}s")
    st.markdown(f'<meta http-equiv="refresh" content="{refresh_interval}">',
                unsafe_allow_html=True)


# ── HEADER ────────────────────────────────────
fetch_ts = st.session_state.fetch_time or "Chưa tải"
st.markdown(f"""
<div class="app-header">
  <div class="app-header-left">
    <div class="app-header-emoji">🏥</div>
    <div>
      <div class="app-header-title">Dashboard Đăng Ký Khám Online</div>
      <div class="app-header-sub">BVĐK Tâm Đức Cầu Quan · Theo dõi lịch hẹn & tình trạng</div>
    </div>
  </div>
  <div class="app-header-time">🕐 {fetch_ts}</div>
</div>
""", unsafe_allow_html=True)

if st.session_state.err:
    st.error(st.session_state.err)


# ── DASHBOARD ─────────────────────────────────
def sec(dot_color, label):
    return f"""<div class="sec-hdr">
      <div class="sec-hdr-dot" style="background:{dot_color}"></div>
      <span class="sec-hdr-text">{label}</span>
    </div>"""

if st.session_state.metrics:
    m = st.session_state.metrics

    # Unique exam dates
    unique_dates = 0
    if COL_EXAM_DATE in m["df"].columns:
        unique_dates = m["df"][COL_EXAM_DATE].astype(str).str.strip().replace("", pd.NA).dropna().nunique()

    # ── KPI Grid (pure CSS — 2 cols mobile, 4 cols desktop) ──
    st.markdown(f"""
    <div class="kpi-grid">

      <div class="kpi-card kc-blue">
        <div class="kpi-bg-icon">📋</div>
        <div class="kpi-label">Tổng Đăng Ký</div>
        <div class="kpi-num">{m['total']}</div>
        <div class="kpi-sub">Toàn bộ lịch hẹn</div>
      </div>

      <div class="kpi-card kc-green">
        <div class="kpi-bg-icon">✅</div>
        <div class="kpi-label">Đã Khám</div>
        <div class="kpi-num">{m['attended_count']}</div>
        <div class="kpi-sub">{m['attended_pct']}% tổng đăng ký</div>
      </div>

      <div class="kpi-card kc-rose">
        <div class="kpi-bg-icon">❌</div>
        <div class="kpi-label">Chưa / Vắng</div>
        <div class="kpi-num">{m['noshow_count']}</div>
        <div class="kpi-sub">{m['noshow_pct']}% tổng đăng ký</div>
      </div>

      <div class="kpi-card kc-violet">
        <div class="kpi-bg-icon">📅</div>
        <div class="kpi-label">Số Ngày Khám</div>
        <div class="kpi-num">{unique_dates}</div>
        <div class="kpi-sub">Ngày có lịch hẹn</div>
      </div>

    </div>
    """, unsafe_allow_html=True)

    # ── Donut + Gender — stack on mobile, side-by-side on desktop ──
    st.markdown(sec(C_GREEN, "Tỷ Lệ Đã Khám / Vắng Khám"), unsafe_allow_html=True)
    col1, col2 = st.columns([3, 2])

    with col1:
        st.markdown('<div class="chart-wrap">', unsafe_allow_html=True)
        st.plotly_chart(chart_donut(m), use_container_width=True,
                        config={"displayModeBar": False})
        st.markdown(f"""
        <div class="legend-row">
          <div class="legend-item">
            <div class="legend-dot" style="background:{C_GREEN}"></div>
            Đã khám — <b style="color:#0f172a">&nbsp;{m['attended_count']}</b>&nbsp;({m['attended_pct']}%)
          </div>
          <div class="legend-item">
            <div class="legend-dot" style="background:{C_ROSE}"></div>
            Chưa / Vắng — <b style="color:#0f172a">&nbsp;{m['noshow_count']}</b>&nbsp;({m['noshow_pct']}%)
          </div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with col2:
        st.markdown(sec(C_BLUE, "Giới Tính"), unsafe_allow_html=True)
        fig_g = chart_gender(m)
        if fig_g:
            st.markdown('<div class="chart-wrap">', unsafe_allow_html=True)
            st.plotly_chart(fig_g, use_container_width=True,
                            config={"displayModeBar": False})
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.info("Không có dữ liệu giới tính.")

    # ── Daily bookings ──
    fig_d = chart_daily(m)
    if fig_d:
        st.markdown(sec(C_BLUE, "Lịch Hẹn Theo Ngày Khám"), unsafe_allow_html=True)
        st.markdown('<div class="chart-wrap">', unsafe_allow_html=True)
        st.plotly_chart(fig_d, use_container_width=True,
                        config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)

    # ── Specialty ──
    fig_s = chart_specialty(m)
    if fig_s:
        st.markdown(sec(C_VIOLET, "Chuyên Khoa Được Đăng Ký Nhiều Nhất"),
                    unsafe_allow_html=True)
        st.markdown('<div class="chart-wrap">', unsafe_allow_html=True)
        st.plotly_chart(fig_s, use_container_width=True,
                        config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)

    # ── Status table ──
    st.markdown(sec(C_AMBER, "Chi Tiết Trạng Thái"), unsafe_allow_html=True)
    if m["status_tbl"] is not None:
        st.dataframe(m["status_tbl"], use_container_width=True,
                     hide_index=True, height=150)

    # ── Raw data ──
    with st.expander("📄 Xem dữ liệu thô"):
        show_cols = [c for c in [COL_TIMESTAMP, COL_NAME, COL_EXAM_DATE,
                                  COL_STATUS, COL_SPECIALTY, COL_GENDER]
                     if c in m["df"].columns]
        st.dataframe(m["df"][show_cols].reset_index(drop=True),
                     use_container_width=True, height=340)

else:
    if not st.session_state.err:
        st.markdown("""
        <div class="empty-wrap">
          <div class="empty-icon">🏥</div>
          <div class="empty-title">Đang tải dữ liệu...</div>
          <div class="empty-desc">
            Nếu không thấy dữ liệu, hãy nhấn <strong>🔄 Tải dữ liệu ngay</strong> trong sidebar.
          </div>
        </div>
        """, unsafe_allow_html=True)
