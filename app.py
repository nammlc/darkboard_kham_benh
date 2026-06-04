"""
Hệ Thống Theo Dõi Đặt Khám Trực Tuyến — BVĐK Tâm Đức Cầu Quan
- Không sidebar
- Navbar inline với nút Làm mới
- Tab: Tổng quan | Báo cáo thống kê | Chi tiết bệnh nhân
- Báo cáo: Ngày · Tuần · Tháng · Quý · Năm
- Fully responsive PC + Mobile
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import gspread
from google.oauth2.service_account import Credentials
import json, os
from datetime import datetime, timedelta

# ── CONFIG ────────────────────────────────────
SHEET_ID   = "1EYiRA3ar41aue8DlbWA7JTKoLL0M2tiLTcZINhdMfTs"
SHEET_NAME = "Câu trả lời biểu mẫu 1"

COL_TIMESTAMP = "THỜI GIAN ĐĂNG KÝ"
COL_STATUS    = "TRẠNG THÁI"
COL_EXAM_DATE = "NGÀY KHÁM"
COL_NAME      = "1. HỌ VÀ TÊN BỆNH NHÂN"
COL_GENDER    = "3. GIỚI TÍNH"
COL_SPECIALTY = "CHUYÊN KHOA MONG MUỐN KHÁM"

STATUS_ATTENDED = "BỆNH NHÂN ĐÃ KHÁM"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

st.set_page_config(
    page_title="Theo Dõi Đặt Khám · BVĐK Tâm Đức",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── COLORS ────────────────────────────────────
CB = "#3b82f6"; CG = "#10b981"; CR = "#f43f5e"
CV = "#8b5cf6"; CT = "#14b8a6"; CA = "#f59e0b"; CS = "#94a3b8"

# ── HIDE SIDEBAR COMPLETELY + FULL CSS ────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@500;600&display=swap');

*, html, body, [class*="css"] {
    font-family: 'Inter', sans-serif !important;
    box-sizing: border-box;
}

/* ── Hide sidebar entirely ── */
section[data-testid="stSidebar"],
[data-testid="stSidebarCollapsedControl"],
button[data-testid="baseButton-headerNoPadding"] {
    display: none !important;
}

/* ── Layout ── */
.stApp { background: #f0f4f8; }
.main .block-container {
    padding: 0 !important;
    max-width: 100% !important;
}

/* ── NAVBAR ── */
.navbar {
    background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 55%, #0e4d7a 100%);
    padding: 0.9rem 1.2rem;
    display: flex; align-items: center;
    justify-content: space-between;
    flex-wrap: wrap; gap: 0.6rem;
    position: sticky; top: 0; z-index: 1000;
    box-shadow: 0 2px 16px rgba(15,23,42,0.25);
}
.navbar-brand {
    display: flex; align-items: center; gap: 0.7rem;
}
.navbar-icon {
    font-size: 1.7rem;
    background: rgba(255,255,255,0.1);
    border-radius: 10px; padding: 0.3rem 0.45rem;
    line-height: 1;
}
.navbar-title {
    font-size: 1rem; font-weight: 700;
    color: #f0f9ff; line-height: 1.2;
}
.navbar-sub {
    font-size: 0.68rem; color: #7dd3fc; margin-top: 0.1rem;
}
.navbar-right {
    display: flex; align-items: center; gap: 0.6rem; flex-wrap: wrap;
}
.navbar-time {
    background: rgba(59,130,246,0.2);
    border: 1px solid rgba(59,130,246,0.35);
    border-radius: 20px; padding: 0.28rem 0.75rem;
    color: #93c5fd; font-size: 0.7rem; font-weight: 500;
    white-space: nowrap;
}
.navbar-status-ok {
    background: rgba(16,185,129,0.15);
    border: 1px solid rgba(16,185,129,0.35);
    border-radius: 20px; padding: 0.28rem 0.75rem;
    color: #6ee7b7; font-size: 0.7rem; font-weight: 600;
    white-space: nowrap;
}
.navbar-status-err {
    background: rgba(244,63,94,0.15);
    border: 1px solid rgba(244,63,94,0.35);
    border-radius: 20px; padding: 0.28rem 0.75rem;
    color: #fca5a5; font-size: 0.7rem; font-weight: 600;
}
@media (min-width: 768px) {
    .navbar { padding: 1rem 2rem; }
    .navbar-title { font-size: 1.15rem; }
}

/* ── CONTENT WRAPPER ── */
.content { padding: 1rem 1rem 3rem; }
@media (min-width: 768px) { .content { padding: 1.5rem 2rem 3rem; } }

/* ── TAB BAR ── */
.tab-bar {
    display: flex; gap: 0.3rem;
    background: white;
    border-radius: 14px; padding: 0.35rem;
    margin-bottom: 1.4rem;
    box-shadow: 0 1px 4px rgba(0,0,0,0.07);
    border: 1px solid #e2e8f0;
    overflow-x: auto; -webkit-overflow-scrolling: touch;
    scrollbar-width: none;
}
.tab-bar::-webkit-scrollbar { display: none; }
.tab-btn {
    flex: 1; min-width: 90px;
    padding: 0.55rem 0.8rem;
    border: none; border-radius: 10px;
    font-size: 0.82rem; font-weight: 600;
    cursor: pointer; white-space: nowrap;
    transition: all 0.18s ease;
    background: transparent; color: #64748b;
    font-family: 'Inter', sans-serif;
}
.tab-btn:hover { background: #f1f5f9; color: #1e293b; }
.tab-btn.active {
    background: linear-gradient(135deg, #3b82f6, #2563eb);
    color: white;
    box-shadow: 0 3px 10px rgba(59,130,246,0.35);
}

/* ── KPI GRID ── */
.kpi-grid {
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 0.75rem; margin-bottom: 1.3rem;
}
@media (min-width: 900px) {
    .kpi-grid { grid-template-columns: repeat(4, 1fr); gap: 1rem; }
}
.kc {
    background: white; border-radius: 16px;
    padding: 1rem 1.1rem;
    border: 1px solid #e2e8f0;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05), 0 4px 14px rgba(0,0,0,0.04);
    position: relative; overflow: hidden;
    transition: transform 0.15s, box-shadow 0.15s;
}
.kc:hover { transform: translateY(-2px); box-shadow: 0 6px 20px rgba(0,0,0,0.08); }
.kc::after {
    content: ''; position: absolute;
    top: 0; left: 0; right: 0; height: 3px;
    border-radius: 16px 16px 0 0;
}
.kc-b::after { background: linear-gradient(90deg,#3b82f6,#60a5fa); }
.kc-g::after { background: linear-gradient(90deg,#10b981,#34d399); }
.kc-r::after { background: linear-gradient(90deg,#f43f5e,#fb7185); }
.kc-v::after { background: linear-gradient(90deg,#8b5cf6,#a78bfa); }
.kc-t::after { background: linear-gradient(90deg,#14b8a6,#2dd4bf); }
.kc-bg { position: absolute; bottom: -0.3rem; right: 0.5rem;
          font-size: 2.8rem; opacity: 0.07; line-height: 1; pointer-events: none; }
.kc-lbl { font-size: 0.65rem; font-weight: 700; text-transform: uppercase;
           letter-spacing: 0.09em; color: #94a3b8; margin-bottom: 0.35rem; }
.kc-val { font-size: 2rem; font-weight: 700; color: #0f172a;
          font-family: 'JetBrains Mono', monospace !important; line-height: 1; }
.kc-sub { font-size: 0.68rem; color: #94a3b8; margin-top: 0.25rem; font-weight: 500; }
@media (min-width: 768px) {
    .kc { padding: 1.3rem 1.5rem; }
    .kc-val { font-size: 2.3rem; }
}

/* ── SECTION HEADER ── */
.sh {
    display: flex; align-items: center; gap: 0.5rem;
    margin: 1.4rem 0 0.8rem;
    padding-bottom: 0.5rem;
    border-bottom: 2px solid #e2e8f0;
}
.sh-dot { width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; }
.sh-txt { font-size: 0.9rem; font-weight: 700; color: #1e293b; }

/* ── CHART CARD ── */
.cc {
    background: white; border-radius: 16px;
    padding: 1.2rem 1rem 0.6rem;
    border: 1px solid #e2e8f0;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05), 0 4px 14px rgba(0,0,0,0.04);
    margin-bottom: 1rem;
}

/* ── LEGEND ── */
.lgd { display: flex; flex-wrap: wrap; gap: 0.7rem 1.4rem;
       justify-content: center; padding: 0.3rem 0 0.7rem; }
.lgd-i { display: flex; align-items: center; gap: 0.4rem;
          font-size: 0.78rem; color: #475569; font-weight: 500; }
.lgd-dot { width: 10px; height: 10px; border-radius: 3px; flex-shrink: 0; }

/* ── PERIOD FILTER ── */
.period-bar {
    display: flex; gap: 0.4rem; flex-wrap: wrap;
    margin-bottom: 1rem;
}
.period-pill {
    padding: 0.38rem 0.9rem;
    border-radius: 20px; border: 1.5px solid #e2e8f0;
    font-size: 0.78rem; font-weight: 600; cursor: pointer;
    background: white; color: #64748b;
    transition: all 0.15s;
    font-family: 'Inter', sans-serif;
}
.period-pill:hover { border-color: #3b82f6; color: #3b82f6; }
.period-pill.active {
    background: #3b82f6; border-color: #3b82f6;
    color: white; box-shadow: 0 2px 8px rgba(59,130,246,0.3);
}

/* ── STAT REPORT TABLE ── */
.report-tbl {
    width: 100%; border-collapse: collapse;
    font-size: 0.82rem; background: white;
    border-radius: 12px; overflow: hidden;
}
.report-tbl th {
    background: #1e3a5f; color: #e0f0ff;
    padding: 0.7rem 0.9rem; text-align: left;
    font-size: 0.72rem; font-weight: 700;
    text-transform: uppercase; letter-spacing: 0.06em;
    white-space: nowrap;
}
.report-tbl td {
    padding: 0.6rem 0.9rem; color: #1e293b;
    border-bottom: 1px solid #f1f5f9;
    white-space: nowrap;
}
.report-tbl tr:nth-child(even) td { background: #f8fafc; }
.report-tbl tr:hover td { background: #eff6ff; }
.report-tbl .num { font-family: 'JetBrains Mono', monospace; font-weight: 600; color: #0f172a; }
.report-tbl .pct-g { color: #059669; font-weight: 600; }
.report-tbl .pct-r { color: #e11d48; font-weight: 600; }
.tbl-scroll { overflow-x: auto; -webkit-overflow-scrolling: touch;
              border-radius: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.07); }

/* ── DATAFRAME ── */
div[data-testid="stDataFrame"] > div {
    border-radius: 12px; overflow: auto !important;
    -webkit-overflow-scrolling: touch;
}
div[data-testid="stDataFrame"] th {
    background: #1e3a5f !important; color: #e0f0ff !important;
    font-size: 0.73rem !important; font-weight: 700 !important;
    text-transform: uppercase !important; white-space: nowrap !important;
    padding: 0.6rem 0.8rem !important;
}
div[data-testid="stDataFrame"] td {
    font-size: 0.8rem !important; color: #1e293b !important;
    padding: 0.5rem 0.8rem !important;
    border-bottom: 1px solid #f1f5f9 !important;
}
div[data-testid="stDataFrame"] tr:nth-child(even) td { background: #f8fafc !important; }
div[data-testid="stDataFrame"] tr:hover td { background: #eff6ff !important; }

/* ── EXPANDER ── */
div[data-testid="stExpander"] {
    border-radius: 14px !important; border: 1px solid #e2e8f0 !important;
    background: white !important; overflow: hidden !important;
}
div[data-testid="stExpander"] summary {
    font-weight: 600 !important; font-size: 0.88rem !important;
    color: #1e293b !important; padding: 0.8rem 1rem !important;
}

/* ── EMPTY STATE ── */
.empty {
    text-align: center; padding: 4rem 1.5rem;
    background: white; border-radius: 20px;
    border: 2px dashed #e2e8f0; margin-top: 1rem;
}
.empty-icon { font-size: 3rem; margin-bottom: 0.8rem; }
.empty-title { font-size: 1.1rem; font-weight: 700; color: #1e293b; }
.empty-desc { font-size: 0.85rem; color: #94a3b8; margin-top: 0.4rem; line-height: 1.6; }

/* ── MOBILE tweaks ── */
@media (max-width: 640px) {
    .navbar-title { font-size: 0.88rem; }
    .kc-val { font-size: 1.75rem; }
}

/* Streamlit tab style override */
div[data-testid="stTabs"] button {
    font-weight: 600 !important; font-size: 0.85rem !important;
}
div[data-testid="stTabs"] button[aria-selected="true"] {
    color: #2563eb !important; border-bottom-color: #2563eb !important;
}
</style>
""", unsafe_allow_html=True)


# ── CHART BASE ────────────────────────────────
BASE = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter, sans-serif", color="#64748b"),
    hoverlabel=dict(bgcolor="#0f172a", font_color="#f1f5f9", font_size=12),
)

def ch_donut(m):
    fig = go.Figure(go.Pie(
        labels=["Đã đến khám", "Vắng / Chưa khám"],
        values=[m["att"], m["nos"]],
        hole=0.68,
        marker=dict(colors=[CG, CR], line=dict(color="#fff", width=3)),
        textinfo="percent",
        textfont=dict(size=13),
        hovertemplate="<b>%{label}</b><br>%{value} người · %{percent}<extra></extra>",
        pull=[0.025, 0.025], direction="clockwise",
    ))
    fig.update_layout(**BASE, height=290, showlegend=False,
        margin=dict(t=8,b=8,l=20,r=20),
        annotations=[dict(
            text=f"<b>{m['total']}</b><br><span style='font-size:11px;color:#94a3b8'>bệnh nhân</span>",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=22, color="#0f172a"),
        )])
    return fig

def ch_gender(m):
    df = m["gen"]
    if df is None or df.empty: return None
    clr = {"NAM": CB, "NỮ": "#f472b6", "NU": "#f472b6"}
    colors = [clr.get(g.upper(), CS) for g in df["Giới tính"]]
    fig = go.Figure(go.Bar(
        x=df["Giới tính"], y=df["Số lượng"],
        marker=dict(color=colors, line=dict(color="#fff", width=2)),
        text=df["Số lượng"], textposition="outside",
        textfont=dict(size=13, family="JetBrains Mono", color="#0f172a"),
        hovertemplate="<b>%{x}</b> — %{y} người<extra></extra>",
        width=0.5,
    ))
    fig.update_layout(**BASE, height=230, margin=dict(t=8,b=8,l=8,r=8),
        xaxis=dict(showgrid=False, tickfont=dict(size=12, color="#475569")),
        yaxis=dict(gridcolor="#f8fafc", zeroline=False))
    return fig

def ch_specialty(m):
    df = m["spec"]
    if df is None or df.empty: return None
    df = df.copy().sort_values("Số lượng")
    df["lbl"] = df["Chuyên khoa"].apply(lambda x: (x[:30]+"…") if len(x)>30 else x)
    pal = [CT, CB, CV, CA, CG, "#06b6d4", "#6366f1", "#f97316"]
    fig = go.Figure(go.Bar(
        y=df["lbl"], x=df["Số lượng"], orientation="h",
        marker=dict(color=pal[:len(df)], line=dict(color="#fff", width=1.5)),
        text=df["Số lượng"], textposition="outside",
        textfont=dict(size=11, family="JetBrains Mono", color="#0f172a"),
        hovertemplate="<b>%{y}</b><br>%{x} bệnh nhân<extra></extra>",
    ))
    fig.update_layout(**BASE, height=max(240, len(df)*46),
        margin=dict(t=8,b=8,l=8,r=44),
        xaxis=dict(gridcolor="#f8fafc", zeroline=False, tickfont=dict(size=9)),
        yaxis=dict(tickfont=dict(size=9.5, color="#475569"), showgrid=False))
    return fig

def ch_trend(df_trend, title, color):
    """Line + bar combo for time-series trend."""
    if df_trend is None or df_trend.empty: return None
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df_trend["Kỳ"], y=df_trend["Đăng ký"],
        name="Lượt đặt",
        marker=dict(color=color, opacity=0.75, line=dict(color="#fff", width=1)),
        hovertemplate="<b>%{x}</b><br>Đăng ký: %{y}<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=df_trend["Kỳ"], y=df_trend["Đã khám"],
        mode="lines+markers", name="Đã khám",
        line=dict(color=CG, width=2.5),
        marker=dict(size=6, color=CG),
        hovertemplate="<b>%{x}</b><br>Đã khám: %{y}<extra></extra>",
    ))
    fig.update_layout(**BASE, height=300,
        margin=dict(t=10,b=10,l=8,r=8),
        xaxis=dict(tickangle=-30, tickfont=dict(size=9, color="#64748b"), showgrid=False),
        yaxis=dict(gridcolor="#f1f5f9", zeroline=False, tickfont=dict(size=9)),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                    font=dict(size=11)),
        bargap=0.3,
    )
    return fig


# ── DATA FUNCTIONS ────────────────────────────

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

def authenticate_gspread(src):
    if isinstance(src, str):
        creds = Credentials.from_service_account_file(src, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_info(src, scopes=SCOPES)
    return gspread.authorize(creds)

def fetch_data(client):
    ws = client.open_by_key(SHEET_ID).worksheet(SHEET_NAME)
    vals = ws.get_all_values()
    if not vals or len(vals) < 2:
        raise ValueError("Sheet trống hoặc không có dữ liệu.")
    headers, rows = vals[0], vals[1:]
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

def parse_date_col(series):
    """Parse dd/mm/yyyy safely, return datetime series."""
    return pd.to_datetime(series, format="%d/%m/%Y", errors="coerce")

def process_data(df):
    if COL_STATUS not in df.columns:
        raise KeyError(f"Không thấy cột '{COL_STATUS}'. Có: {list(df.columns)}")
    df = df.copy()
    df[COL_STATUS] = df[COL_STATUS].astype(str).str.strip()
    df = df[~df[COL_STATUS].isin(["", "nan", "N/A", "\u200b"])]
    total = len(df)
    if total == 0: return _empty(df)

    att = int((df[COL_STATUS].str.upper() == STATUS_ATTENDED.upper()).sum())
    nos = total - att

    # Parse exam date
    if COL_EXAM_DATE in df.columns:
        df["_date"] = parse_date_col(df[COL_EXAM_DATE].astype(str).str.strip())
    else:
        df["_date"] = pd.NaT

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
    date_ok = df["_date"].notna()
    if date_ok.any():
        dg = df[date_ok].copy()
        daily = dg.groupby("_date").size().reset_index(name="Lịch hẹn")
        daily.columns = ["Ngày khám", "Lịch hẹn"]
        daily["Ngày khám"] = daily["Ngày khám"].dt.strftime("%d/%m/%Y")

    status_tbl = df[COL_STATUS].value_counts().reset_index()
    status_tbl.columns = ["Trạng thái", "Số lượng"]

    return dict(total=total, att=att, nos=nos,
                att_pct=round(att/total*100,1),
                nos_pct=round(nos/total*100,1),
                spec=spec, gen=gen, daily=daily,
                status_tbl=status_tbl, df=df)

def _empty(df):
    df = df.copy()
    df["_date"] = pd.NaT   # always ensure _date column exists
    return dict(total=0, att=0, nos=0, att_pct=0.0, nos_pct=0.0,
                spec=None, gen=None, daily=None, status_tbl=None, df=df)


# ── STATISTICS BUILDER ───────────────────────

def build_stats(df, period):
    """Build time-series stats grouped by period."""
    if "_date" not in df.columns:
        return None
    date_ok = df["_date"].notna()
    if not date_ok.any():
        return None
    d = df[date_ok].copy()
    att_mask = d[COL_STATUS].str.upper() == STATUS_ATTENDED.upper()

    if period == "Ngày":
        d["Kỳ"] = d["_date"].dt.strftime("%d/%m/%Y")
    elif period == "Tuần":
        # Week label: Mon dd/mm – Sun dd/mm
        d["Kỳ"] = d["_date"].apply(
            lambda x: f"Tuần {x.isocalendar()[1]}/{x.year}\n({(x - timedelta(days=x.weekday())).strftime('%d/%m')}–{(x + timedelta(days=6-x.weekday())).strftime('%d/%m')})"
        )
    elif period == "Tháng":
        d["Kỳ"] = d["_date"].dt.strftime("Tháng %m/%Y")
    elif period == "Quý":
        d["Kỳ"] = d["_date"].apply(
            lambda x: f"Q{((x.month-1)//3)+1}/{x.year}"
        )
    elif period == "Năm":
        d["Kỳ"] = d["_date"].dt.strftime("Năm %Y")

    grp = d.groupby("Kỳ", sort=False)
    stats = grp.size().reset_index(name="Đăng ký")
    stats["Đã khám"] = grp.apply(lambda g: (g[COL_STATUS].str.upper() == STATUS_ATTENDED.upper()).sum()).values
    stats["Vắng / Chưa"] = stats["Đăng ký"] - stats["Đã khám"]
    stats["Tỷ lệ đến (%)"] = (stats["Đã khám"] / stats["Đăng ký"] * 100).round(1)
    stats["Tỷ lệ vắng (%)"] = (stats["Vắng / Chưa"] / stats["Đăng ký"] * 100).round(1)

    # Sort chronologically by first date in each group
    first_date = d.groupby("Kỳ")["_date"].min().reset_index()
    first_date.columns = ["Kỳ", "_sort"]
    stats = stats.merge(first_date, on="Kỳ").sort_values("_sort").drop(columns="_sort")
    return stats

def kpi_for_period(stats):
    """Summary KPIs for a given period stats table."""
    if stats is None or stats.empty:
        return 0, 0, 0.0, 0
    total = int(stats["Đăng ký"].sum())
    att   = int(stats["Đã khám"].sum())
    pct   = round(att/total*100, 1) if total > 0 else 0.0
    peaks = int(stats["Đăng ký"].max())
    return total, att, pct, peaks


# ── CREDENTIALS & SESSION ────────────────────
creds_data = get_credentials()

for k, v in [("metrics", None), ("fetch_time", None), ("err", None), ("active_tab", 0)]:
    if k not in st.session_state:
        st.session_state[k] = v

def do_fetch():
    if not creds_data:
        st.session_state.err = "⚠️ Chưa có thông tin xác thực. Kiểm tra Streamlit Secrets."
        st.session_state.metrics = None
        return
    try:
        cl  = authenticate_gspread(creds_data)
        raw = fetch_data(cl)
        st.session_state.metrics    = process_data(raw)
        st.session_state.fetch_time = datetime.now().strftime("%H:%M · %d/%m/%Y")
        st.session_state.err        = None
    except Exception as e:
        st.session_state.err     = f"❌ {type(e).__name__}: {e}"
        st.session_state.metrics = None

if st.session_state.metrics is None and st.session_state.err is None:
    with st.spinner("Đang kết nối và tải dữ liệu…"):
        do_fetch()


# ── NAVBAR ───────────────────────────────────
status_html = (
    '<span class="navbar-status-ok">● Đã kết nối</span>' if creds_data
    else '<span class="navbar-status-err">● Chưa kết nối</span>'
)
fetch_ts = st.session_state.fetch_time or "Chưa tải"

st.markdown(f"""
<div class="navbar">
  <div class="navbar-brand">
    <div class="navbar-icon">🏥</div>
    <div>
      <div class="navbar-title">BVĐK Tâm Đức Cầu Quan</div>
      <div class="navbar-sub">Hệ thống theo dõi đặt khám trực tuyến</div>
    </div>
  </div>
  <div class="navbar-right">
    {status_html}
    <span class="navbar-time">🕐 {fetch_ts}</span>
  </div>
</div>
""", unsafe_allow_html=True)

# ── CONTENT ───────────────────────────────────
st.markdown('<div class="content">', unsafe_allow_html=True)

# Error banner
if st.session_state.err:
    st.error(st.session_state.err)

# Refresh button (top of content)
col_r1, col_r2, col_r3 = st.columns([3, 1, 3])
with col_r2:
    if st.button("🔄 Làm mới"):
        with st.spinner("Đang tải…"):
            do_fetch()
        st.rerun()

st.markdown("<div style='margin-bottom:0.5rem'></div>", unsafe_allow_html=True)

# ── TABS ──────────────────────────────────────
if st.session_state.metrics:
    m = st.session_state.metrics
    df = m["df"]

    tab1, tab2, tab3 = st.tabs(["📊 Tổng Quan", "📈 Báo Cáo Thống Kê", "👤 Chi Tiết Bệnh Nhân"])

    # ════════════════════════════════════════════
    # TAB 1 — TỔNG QUAN
    # ════════════════════════════════════════════
    with tab1:
        if "_date" in df.columns and df["_date"].notna().any():
            unique_dates = int(df["_date"].dropna().dt.date.nunique())
        else:
            unique_dates = 0

        st.markdown(f"""
        <div class="kpi-grid">
          <div class="kc kc-b">
            <div class="kc-bg">📋</div>
            <div class="kc-lbl">Tổng Lượt Đặt Khám</div>
            <div class="kc-val">{m['total']}</div>
            <div class="kc-sub">Tổng số lịch hẹn đã đặt</div>
          </div>
          <div class="kc kc-g">
            <div class="kc-bg">✅</div>
            <div class="kc-lbl">Đã Đến Khám</div>
            <div class="kc-val">{m['att']}</div>
            <div class="kc-sub">{m['att_pct']}% tổng lượt đặt</div>
          </div>
          <div class="kc kc-r">
            <div class="kc-bg">❌</div>
            <div class="kc-lbl">Vắng / Chưa Khám</div>
            <div class="kc-val">{m['nos']}</div>
            <div class="kc-sub">{m['nos_pct']}% tổng lượt đặt</div>
          </div>
          <div class="kc kc-v">
            <div class="kc-bg">📅</div>
            <div class="kc-lbl">Số Ngày Có Lịch</div>
            <div class="kc-val">{unique_dates}</div>
            <div class="kc-sub">Ngày có lịch hẹn khám</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        # Donut + Gender
        st.markdown('<div class="sh"><div class="sh-dot" style="background:#10b981"></div><span class="sh-txt">Tỷ Lệ Đến Khám và Vắng Mặt</span></div>', unsafe_allow_html=True)
        c1, c2 = st.columns([3, 2])
        with c1:
            st.markdown('<div class="cc">', unsafe_allow_html=True)
            st.plotly_chart(ch_donut(m), use_container_width=True, config={"displayModeBar": False})
            st.markdown(f"""
            <div class="lgd">
              <div class="lgd-i"><div class="lgd-dot" style="background:{CG}"></div>
                Đã đến khám — <b style="color:#0f172a">&nbsp;{m['att']}</b>&nbsp;({m['att_pct']}%)
              </div>
              <div class="lgd-i"><div class="lgd-dot" style="background:{CR}"></div>
                Vắng / Chưa khám — <b style="color:#0f172a">&nbsp;{m['nos']}</b>&nbsp;({m['nos_pct']}%)
              </div>
            </div>
            """, unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c2:
            st.markdown('<div class="sh"><div class="sh-dot" style="background:#3b82f6"></div><span class="sh-txt">Phân Bố Giới Tính</span></div>', unsafe_allow_html=True)
            fg = ch_gender(m)
            if fg:
                st.markdown('<div class="cc">', unsafe_allow_html=True)
                st.plotly_chart(fg, use_container_width=True, config={"displayModeBar": False})
                st.markdown('</div>', unsafe_allow_html=True)
            else:
                st.info("Không có dữ liệu giới tính bệnh nhân.")

        # Specialty
        fs = ch_specialty(m)
        if fs:
            st.markdown('<div class="sh"><div class="sh-dot" style="background:#8b5cf6"></div><span class="sh-txt">Chuyên Khoa Được Đặt Khám Nhiều Nhất</span></div>', unsafe_allow_html=True)
            st.markdown('<div class="cc">', unsafe_allow_html=True)
            st.plotly_chart(fs, use_container_width=True, config={"displayModeBar": False})
            st.markdown('</div>', unsafe_allow_html=True)

    # ════════════════════════════════════════════
    # TAB 2 — BÁO CÁO THỐNG KÊ
    # ════════════════════════════════════════════
    with tab2:
        period_opts = ["Ngày", "Tuần", "Tháng", "Quý", "Năm"]
        period_colors = [CB, CT, CV, CA, CG]
        period_icons  = ["📆", "🗓️", "🗂️", "📊", "🏆"]

        sel_period = st.radio(
            "Xem theo:",
            period_opts,
            horizontal=True,
            index=2,
            label_visibility="collapsed",
        )
        color_idx = period_opts.index(sel_period)
        p_color   = period_colors[color_idx]
        p_icon    = period_icons[color_idx]

        stats = build_stats(df, sel_period)
        tot_p, att_p, pct_p, peak_p = kpi_for_period(stats)

        # KPI mini row
        st.markdown(f"""
        <div class="kpi-grid" style="grid-template-columns:repeat(2,1fr);gap:0.65rem;margin-bottom:1.2rem">
          <div class="kc kc-b">
            <div class="kc-bg">{p_icon}</div>
            <div class="kc-lbl">Tổng Lượt Đặt — {sel_period}</div>
            <div class="kc-val">{tot_p}</div>
            <div class="kc-sub">Trong tất cả kỳ hiển thị</div>
          </div>
          <div class="kc kc-g">
            <div class="kc-bg">✅</div>
            <div class="kc-lbl">Đã Đến Khám</div>
            <div class="kc-val">{att_p}</div>
            <div class="kc-sub">{pct_p}% tổng lượt đặt</div>
          </div>
          <div class="kc kc-r">
            <div class="kc-bg">📉</div>
            <div class="kc-lbl">Vắng / Chưa Khám</div>
            <div class="kc-val">{tot_p - att_p}</div>
            <div class="kc-sub">{round(100-pct_p,1)}% tổng lượt đặt</div>
          </div>
          <div class="kc kc-t">
            <div class="kc-bg">🔝</div>
            <div class="kc-lbl">Kỳ Đông Nhất</div>
            <div class="kc-val">{peak_p}</div>
            <div class="kc-sub">Lượt đặt cao nhất / kỳ</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        if stats is not None and not stats.empty:
            # Trend chart
            st.markdown(f'<div class="sh"><div class="sh-dot" style="background:{p_color}"></div><span class="sh-txt">Biểu Đồ Xu Hướng Theo {sel_period}</span></div>', unsafe_allow_html=True)
            st.markdown('<div class="cc">', unsafe_allow_html=True)
            ft = ch_trend(stats[["Kỳ","Đăng ký","Đã khám"]], sel_period, p_color)
            if ft:
                st.plotly_chart(ft, use_container_width=True, config={"displayModeBar": False})
            st.markdown('</div>', unsafe_allow_html=True)

            # Detailed table
            st.markdown(f'<div class="sh"><div class="sh-dot" style="background:{p_color}"></div><span class="sh-txt">Bảng Chi Tiết Theo {sel_period}</span></div>', unsafe_allow_html=True)

            rows_html = ""
            for _, row in stats.iterrows():
                pct_g_cls = "pct-g" if row["Tỷ lệ đến (%)"] >= 50 else "pct-r"
                pct_r_cls = "pct-r" if row["Tỷ lệ vắng (%)"] >= 50 else "pct-g"
                rows_html += f"""
                <tr>
                  <td>{row['Kỳ']}</td>
                  <td class="num">{int(row['Đăng ký'])}</td>
                  <td class="num" style="color:#059669">{int(row['Đã khám'])}</td>
                  <td class="num" style="color:#e11d48">{int(row['Vắng / Chưa'])}</td>
                  <td class="{pct_g_cls}">{row['Tỷ lệ đến (%)']}%</td>
                  <td class="{pct_r_cls}">{row['Tỷ lệ vắng (%)']}%</td>
                </tr>"""

            st.markdown(f"""
            <div class="tbl-scroll">
              <table class="report-tbl">
                <thead>
                  <tr>
                    <th>Kỳ</th>
                    <th>Tổng Đăng Ký</th>
                    <th>Đã Đến Khám</th>
                    <th>Vắng / Chưa</th>
                    <th>Tỷ Lệ Đến</th>
                    <th>Tỷ Lệ Vắng</th>
                  </tr>
                </thead>
                <tbody>{rows_html}</tbody>
              </table>
            </div>
            """, unsafe_allow_html=True)

            # Download CSV
            st.markdown("<br>", unsafe_allow_html=True)
            csv = stats.drop(columns=["_sort"], errors="ignore").to_csv(index=False, encoding="utf-8-sig")
            st.download_button(
                label=f"⬇️ Tải báo cáo theo {sel_period} (.csv)",
                data=csv.encode("utf-8-sig"),
                file_name=f"baocao_{sel_period.lower()}_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
            )
        else:
            st.markdown("""
            <div class="empty">
              <div class="empty-icon">📭</div>
              <div class="empty-title">Không có dữ liệu cho kỳ này</div>
              <div class="empty-desc">Dữ liệu ngày khám chưa được nhập hoặc không đúng định dạng dd/mm/yyyy.</div>
            </div>
            """, unsafe_allow_html=True)

    # ════════════════════════════════════════════
    # TAB 3 — CHI TIẾT BỆNH NHÂN
    # ════════════════════════════════════════════
    with tab3:
        show_cols = [c for c in [COL_TIMESTAMP, COL_NAME, COL_EXAM_DATE,
                                  COL_STATUS, COL_SPECIALTY, COL_GENDER]
                     if c in df.columns]

        # Filter by status
        st.markdown('<div class="sh"><div class="sh-dot" style="background:#f59e0b"></div><span class="sh-txt">Lọc Dữ Liệu</span></div>', unsafe_allow_html=True)
        fc1, fc2 = st.columns([1, 2])
        with fc1:
            status_opts = ["Tất cả"] + list(m["status_tbl"]["Trạng thái"].unique())
            sel_status  = st.selectbox("Trạng thái", status_opts, label_visibility="collapsed")
        with fc2:
            if COL_SPECIALTY in df.columns:
                spec_opts = ["Tất cả chuyên khoa"] + sorted(
                    df[df[COL_SPECIALTY].str.strip() != ""][COL_SPECIALTY].unique().tolist()
                )
                sel_spec = st.selectbox("Chuyên khoa", spec_opts, label_visibility="collapsed")
            else:
                sel_spec = "Tất cả chuyên khoa"

        filtered = df.copy()
        if sel_status != "Tất cả":
            filtered = filtered[filtered[COL_STATUS] == sel_status]
        if sel_spec != "Tất cả chuyên khoa" and COL_SPECIALTY in filtered.columns:
            filtered = filtered[filtered[COL_SPECIALTY] == sel_spec]

        st.markdown(f"""
        <div style="margin:0.6rem 0 0.8rem;padding:0.5rem 0.9rem;background:white;
             border-radius:10px;border:1px solid #e2e8f0;font-size:0.82rem;color:#475569">
          Hiển thị <b style="color:#0f172a">{len(filtered)}</b> / {m['total']} bệnh nhân
        </div>
        """, unsafe_allow_html=True)

        st.dataframe(
            filtered[show_cols].reset_index(drop=True),
            use_container_width=True, hide_index=True, height=420,
        )

        csv2 = filtered[show_cols].to_csv(index=False, encoding="utf-8-sig")
        st.download_button(
            label="⬇️ Tải danh sách bệnh nhân (.csv)",
            data=csv2.encode("utf-8-sig"),
            file_name=f"danhsach_benhnhan_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
        )

else:
    # Empty / error state
    if not st.session_state.err:
        st.markdown("""
        <div class="empty">
          <div class="empty-icon">🏥</div>
          <div class="empty-title">Đang tải dữ liệu, vui lòng chờ...</div>
          <div class="empty-desc">Nếu dữ liệu không hiển thị sau vài giây,
            hãy nhấn nút <strong>🔄 Làm mới</strong> ở trên.</div>
        </div>
        """, unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)   # close .content
