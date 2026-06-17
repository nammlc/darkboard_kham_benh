"""
BVĐK Tâm Đức Cầu Quan — Hệ Thống Theo Dõi Đặt Khám Trực Tuyến
Mobile-first · No sidebar · Today's stats · Date search
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import gspread
from google.oauth2.service_account import Credentials
import json, os
from datetime import datetime, timedelta, date

SHEET_ID        = "1EYiRA3ar41aue8DlbWA7JTKoLL0M2tiLTcZINhdMfTs"
SHEET_NAME      = "Câu trả lời biểu mẫu 1"
COL_STATUS      = "TRẠNG THÁI"
COL_EXAM_DATE   = "NGÀY KHÁM"
COL_NAME        = "1. HỌ VÀ TÊN BỆNH NHÂN"
COL_GENDER      = "3. GIỚI TÍNH"
COL_SPECIALTY   = "CHUYÊN KHOA MONG MUỐN KHÁM"
COL_TIMESTAMP   = "THỜI GIAN ĐĂNG KÝ"
COL_DOCTOR      = "BÁC SĨ MONG MUỐN ( nếu có)"
COL_SOURCE      = "NGUỒN BỆNH NHÂN"
STATUS_ATTENDED = "BỆNH NHÂN ĐÃ KHÁM"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

st.set_page_config(
    page_title="Theo Dõi Đặt Khám · BVĐK Tâm Đức",
    page_icon="🏥", layout="wide",
    initial_sidebar_state="collapsed",
)

CB="#3b82f6"; CG="#10b981"; CR="#f43f5e"; CV="#8b5cf6"
CT="#14b8a6"; CA="#f59e0b"; CS="#94a3b8"

# ═══════════════════════════════════════════════
# CSS — MOBILE FIRST
# ═══════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@500;600&display=swap');

*, html, body, [class*="css"] {
    font-family: 'Inter', sans-serif !important;
    box-sizing: border-box;
    -webkit-tap-highlight-color: transparent;
}

/* Hide sidebar */
section[data-testid="stSidebar"],
[data-testid="stSidebarCollapsedControl"],
button[data-testid="baseButton-headerNoPadding"] { display:none !important; }

.stApp { background:#f0f4f8; }
.main .block-container { padding:0 !important; max-width:100% !important; }

/* ── NAVBAR ── */
.nb {
    background: linear-gradient(135deg,#0f172a 0%,#1e3a5f 60%,#0e4d7a 100%);
    padding: 0.8rem 1rem;
    display: flex; align-items: center;
    justify-content: space-between; gap: 0.5rem;
    position: sticky; top:0; z-index:1000;
    box-shadow: 0 2px 16px rgba(15,23,42,0.3);
}
.nb-brand { display:flex; align-items:center; gap:0.6rem; min-width:0; }
.nb-ico {
    font-size:1.5rem; background:rgba(255,255,255,0.1);
    border-radius:9px; padding:0.28rem 0.4rem;
    line-height:1; flex-shrink:0;
}
.nb-name { font-size:0.85rem; font-weight:700; color:#f0f9ff; line-height:1.2; }
.nb-sub  { font-size:0.62rem; color:#7dd3fc; margin-top:0.08rem; }
.nb-right { display:flex; align-items:center; gap:0.4rem; flex-shrink:0; }
.badge {
    border-radius:20px; padding:0.22rem 0.6rem;
    font-size:0.65rem; font-weight:600; white-space:nowrap;
}
.badge-ok  { background:rgba(16,185,129,0.15); border:1px solid rgba(16,185,129,0.4); color:#6ee7b7; }
.badge-err { background:rgba(244,63,94,0.15);  border:1px solid rgba(244,63,94,0.4);  color:#fca5a5; }
.badge-time{ background:rgba(59,130,246,0.15); border:1px solid rgba(59,130,246,0.4); color:#93c5fd; }
@media(min-width:640px){
    .nb { padding:1rem 2rem; }
    .nb-name { font-size:1.05rem; }
    .badge { font-size:0.72rem; padding:0.28rem 0.75rem; }
}

/* ── CONTENT ── */
.content { padding:0.75rem 0.75rem 4rem; }
@media(min-width:640px){ .content { padding:1.2rem 1.5rem 4rem; } }
@media(min-width:1024px){ .content { padding:1.5rem 2.5rem 4rem; } }

/* ── TODAY BANNER ── */
.today-banner {
    background: linear-gradient(135deg,#0f4c75,#1b6ca8);
    border-radius:16px; padding:1rem 1.2rem;
    margin-bottom:1rem;
    box-shadow:0 4px 20px rgba(15,76,117,0.3);
    display:flex; flex-wrap:wrap; align-items:center;
    justify-content:space-between; gap:0.8rem;
}
.today-label {
    font-size:0.68rem; font-weight:700; text-transform:uppercase;
    letter-spacing:0.1em; color:#7dd3fc; margin-bottom:0.2rem;
}
.today-date { font-size:0.9rem; font-weight:600; color:#e0f2fe; }
.today-stats { display:flex; gap:1.2rem; flex-wrap:wrap; }
.today-stat { text-align:center; }
.today-stat-val {
    font-size:1.8rem; font-weight:700; color:white;
    font-family:'JetBrains Mono',monospace !important;
    line-height:1;
}
.today-stat-lbl { font-size:0.62rem; color:#bae6fd; margin-top:0.15rem; font-weight:500; }
@media(min-width:640px){
    .today-stat-val { font-size:2.2rem; }
    .today-banner { padding:1.2rem 1.8rem; }
}

/* ── KPI GRID ── */
.kg {
    display:grid;
    grid-template-columns:repeat(2,1fr);
    gap:0.6rem; margin-bottom:1rem;
}
@media(min-width:900px){ .kg { grid-template-columns:repeat(4,1fr); gap:0.9rem; } }
.kc {
    background:white; border-radius:14px;
    padding:0.9rem 1rem;
    border:1px solid #e2e8f0;
    box-shadow:0 1px 3px rgba(0,0,0,0.05),0 3px 12px rgba(0,0,0,0.04);
    position:relative; overflow:hidden;
}
.kc::after {
    content:''; position:absolute; top:0; left:0; right:0; height:3px;
    border-radius:14px 14px 0 0;
}
.kc-b::after{background:linear-gradient(90deg,#3b82f6,#60a5fa);}
.kc-g::after{background:linear-gradient(90deg,#10b981,#34d399);}
.kc-r::after{background:linear-gradient(90deg,#f43f5e,#fb7185);}
.kc-v::after{background:linear-gradient(90deg,#8b5cf6,#a78bfa);}
.kc-t::after{background:linear-gradient(90deg,#14b8a6,#2dd4bf);}
.kc-bg { position:absolute; bottom:-0.2rem; right:0.4rem;
          font-size:2.5rem; opacity:0.06; line-height:1; pointer-events:none; }
.kc-lbl { font-size:0.6rem; font-weight:700; text-transform:uppercase;
           letter-spacing:0.09em; color:#94a3b8; margin-bottom:0.3rem; }
.kc-val { font-size:1.75rem; font-weight:700; color:#0f172a;
          font-family:'JetBrains Mono',monospace !important; line-height:1; }
.kc-sub { font-size:0.62rem; color:#94a3b8; margin-top:0.22rem; }
@media(min-width:640px){
    .kc { padding:1.2rem 1.4rem; }
    .kc-val { font-size:2.1rem; }
    .kc-lbl { font-size:0.65rem; }
}

/* ── SECTION HEADER ── */
.sh {
    display:flex; align-items:center; gap:0.45rem;
    margin:1.2rem 0 0.7rem;
    padding-bottom:0.45rem;
    border-bottom:2px solid #e2e8f0;
}
.sh-dot { width:9px; height:9px; border-radius:50%; flex-shrink:0; }
.sh-txt { font-size:0.82rem; font-weight:700; color:#1e293b; }
@media(min-width:640px){ .sh-txt { font-size:0.92rem; } }

/* ── CHART CARD ── */
.cc {
    background:white; border-radius:14px;
    padding:1rem 0.8rem 0.5rem;
    border:1px solid #e2e8f0;
    box-shadow:0 1px 3px rgba(0,0,0,0.05),0 3px 12px rgba(0,0,0,0.04);
    margin-bottom:0.9rem;
}

/* ── LEGEND ── */
.lgd { display:flex; flex-wrap:wrap; gap:0.5rem 1.2rem;
       justify-content:center; padding:0.3rem 0 0.6rem; }
.lgd-i { display:flex; align-items:center; gap:0.35rem;
          font-size:0.75rem; color:#475569; font-weight:500; }
.lgd-dot { width:9px; height:9px; border-radius:3px; flex-shrink:0; }

/* ── REFRESH BUTTON ── */
.stButton>button {
    background:linear-gradient(135deg,#3b82f6,#2563eb) !important;
    color:white !important; border:none !important;
    border-radius:10px !important; font-weight:600 !important;
    font-size:0.85rem !important;
    box-shadow:0 3px 10px rgba(59,130,246,0.3) !important;
    padding:0.5rem 1.2rem !important;
    transition:all 0.2s !important;
}
.stButton>button:hover {
    background:linear-gradient(135deg,#60a5fa,#3b82f6) !important;
    transform:translateY(-1px) !important;
}

/* ── DATE SEARCH BOX ── */
.date-search-wrap {
    background:white; border-radius:14px;
    padding:1rem 1.1rem;
    border:1px solid #e2e8f0;
    box-shadow:0 1px 3px rgba(0,0,0,0.05),0 3px 12px rgba(0,0,0,0.04);
    margin-bottom:1rem;
}
.date-search-title {
    font-size:0.78rem; font-weight:700; color:#1e293b;
    margin-bottom:0.6rem; display:flex; align-items:center; gap:0.4rem;
}

/* ── PATIENT CARD (mobile) ── */
.pt-card {
    background:white; border-radius:12px;
    padding:0.85rem 1rem; margin-bottom:0.55rem;
    border:1px solid #e2e8f0;
    box-shadow:0 1px 4px rgba(0,0,0,0.05);
}
.pt-name { font-size:0.9rem; font-weight:700; color:#0f172a; margin-bottom:0.3rem; }
.pt-row  { display:flex; flex-wrap:wrap; gap:0.3rem 0.8rem; margin-bottom:0.25rem; }
.pt-tag  {
    font-size:0.68rem; font-weight:600; padding:0.2rem 0.55rem;
    border-radius:20px; white-space:nowrap;
}
.pt-tag-date  { background:#eff6ff; color:#1d4ed8; }
.pt-tag-spec  { background:#f0fdf4; color:#166534; }
.pt-tag-doc   { background:#fdf4ff; color:#6b21a8; }
.pt-status-att { background:#d1fae5; color:#065f46; }
.pt-status-nos { background:#fee2e2; color:#991b1b; }
.pt-status-oth { background:#f1f5f9; color:#475569; }

/* ── REPORT TABLE ── */
.rtbl-wrap { overflow-x:auto; -webkit-overflow-scrolling:touch;
             border-radius:12px; box-shadow:0 1px 4px rgba(0,0,0,0.07); }
.rtbl {
    width:100%; border-collapse:collapse;
    font-size:0.78rem; background:white; border-radius:12px;
    min-width:420px;
}
.rtbl th {
    background:#1e3a5f; color:#e0f0ff;
    padding:0.65rem 0.8rem; text-align:left;
    font-size:0.68rem; font-weight:700;
    text-transform:uppercase; letter-spacing:0.06em; white-space:nowrap;
}
.rtbl td { padding:0.55rem 0.8rem; color:#1e293b;
           border-bottom:1px solid #f1f5f9; }
.rtbl tr:nth-child(even) td { background:#f8fafc; }
.rtbl tr:hover td { background:#eff6ff; }
.rtbl .num { font-family:'JetBrains Mono',monospace; font-weight:600; color:#0f172a; }
.pct-g { color:#059669; font-weight:700; }
.pct-r { color:#dc2626; font-weight:700; }

/* ── EMPTY STATE ── */
.empty {
    text-align:center; padding:3rem 1.2rem;
    background:white; border-radius:16px;
    border:2px dashed #e2e8f0; margin-top:0.8rem;
}
.empty-ico  { font-size:2.8rem; margin-bottom:0.7rem; }
.empty-ttl  { font-size:1rem; font-weight:700; color:#1e293b; }
.empty-dsc  { font-size:0.8rem; color:#94a3b8; margin-top:0.35rem; line-height:1.6; }

/* ── STREAMLIT OVERRIDES ── */
div[data-testid="stTabs"] button {
    font-weight:600 !important; font-size:0.82rem !important;
    padding:0.5rem 0.6rem !important;
}
div[data-testid="stTabs"] button[aria-selected="true"] {
    color:#2563eb !important; border-bottom-color:#2563eb !important;
}
div[data-testid="stDataFrame"] > div {
    border-radius:10px; overflow:auto !important;
    -webkit-overflow-scrolling:touch;
}
div[data-testid="stDataFrame"] th {
    background:#1e3a5f !important; color:#e0f0ff !important;
    font-size:0.68rem !important; font-weight:700 !important;
    text-transform:uppercase !important; white-space:nowrap !important;
    padding:0.55rem 0.7rem !important;
}
div[data-testid="stDataFrame"] td {
    font-size:0.78rem !important; color:#1e293b !important;
    padding:0.45rem 0.7rem !important;
    border-bottom:1px solid #f1f5f9 !important;
}
div[data-testid="stDataFrame"] tr:nth-child(even) td { background:#f8fafc !important; }
div[data-testid="stExpander"] {
    border-radius:12px !important; border:1px solid #e2e8f0 !important;
    background:white !important;
}
div[data-testid="stExpander"] summary {
    font-weight:600 !important; font-size:0.84rem !important;
    color:#1e293b !important;
}
/* Make selectbox / date_input bigger touch targets on mobile */
div[data-testid="stSelectbox"] select,
div[data-baseweb="select"] { font-size:0.85rem !important; }
input[type="date"] { font-size:1rem !important; }
/* Download button */
div[data-testid="stDownloadButton"] button {
    background:linear-gradient(135deg,#10b981,#059669) !important;
    color:white !important; border:none !important;
    border-radius:10px !important; font-weight:600 !important;
    font-size:0.82rem !important;
}

/* ── SOURCE BADGES ── */
.src-noi  { background:#ede9fe; color:#5b21b6; border:1px solid #c4b5fd; }
.src-vl   { background:#fef3c7; color:#92400e; border:1px solid #fcd34d; }
.src-other{ background:#f1f5f9; color:#475569; border:1px solid #cbd5e1; }

/* ── 3-DAY UPCOMING TABLE ── */
.upcoming-day {
    background:white; border-radius:14px;
    border:1px solid #e2e8f0;
    box-shadow:0 1px 3px rgba(0,0,0,0.05),0 3px 12px rgba(0,0,0,0.04);
    margin-bottom:0.9rem; overflow:hidden;
}
.upcoming-day-header {
    display:flex; align-items:center; justify-content:space-between;
    padding:0.75rem 1rem;
    background:linear-gradient(135deg,#1e3a5f,#0e4d7a);
}
.upcoming-day-title { font-size:0.88rem; font-weight:700; color:#f0f9ff; }
.upcoming-day-count {
    font-size:0.72rem; font-weight:700;
    background:rgba(255,255,255,0.15); color:white;
    border-radius:20px; padding:0.18rem 0.6rem;
}
.upcoming-day-body { padding:0.5rem 0.75rem 0.75rem; }

/* ── SOURCE STATS CARDS ── */
.src-grid {
    display:grid; grid-template-columns:repeat(2,1fr);
    gap:0.6rem; margin-bottom:1rem;
}
@media(min-width:640px){ .src-grid { grid-template-columns:repeat(3,1fr); } }
.src-card {
    background:white; border-radius:14px;
    padding:0.9rem 1rem; border:1px solid #e2e8f0;
    box-shadow:0 1px 3px rgba(0,0,0,0.05),0 3px 12px rgba(0,0,0,0.04);
    text-align:center; position:relative; overflow:hidden;
}
.src-card::after {
    content:''; position:absolute; top:0; left:0; right:0; height:3px;
    border-radius:14px 14px 0 0;
}
.src-card-noi::after  { background:linear-gradient(90deg,#8b5cf6,#a78bfa); }
.src-card-vl::after   { background:linear-gradient(90deg,#f59e0b,#fbbf24); }
.src-card-total::after{ background:linear-gradient(90deg,#3b82f6,#60a5fa); }
.src-card-ico  { font-size:1.6rem; margin-bottom:0.3rem; }
.src-card-lbl  { font-size:0.62rem; font-weight:700; text-transform:uppercase;
                 letter-spacing:0.08em; color:#94a3b8; margin-bottom:0.25rem; }
.src-card-val  { font-size:1.8rem; font-weight:700; color:#0f172a;
                 font-family:'JetBrains Mono',monospace !important; line-height:1; }
.src-card-sub  { font-size:0.62rem; color:#94a3b8; margin-top:0.2rem; }
@media(min-width:640px){
    .src-card-val { font-size:2.1rem; }
}

/* Scroll hint */
.scroll-hint {
    text-align:center; font-size:0.65rem; color:#94a3b8;
    padding:0.25rem 0; font-style:italic;
}
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════
# CHART HELPERS
# ═══════════════════════════════════════════════
BASE = dict(
    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter, sans-serif", color="#64748b"),
    hoverlabel=dict(bgcolor="#0f172a", font_color="#f1f5f9", font_size=12),
)

def ch_donut(att, nos, total):
    fig = go.Figure(go.Pie(
        labels=["Đã đến khám","Vắng / Chưa khám"],
        values=[att, nos], hole=0.68,
        marker=dict(colors=[CG,CR], line=dict(color="#fff",width=3)),
        textinfo="percent", textfont=dict(size=12),
        hovertemplate="<b>%{label}</b><br>%{value} người · %{percent}<extra></extra>",
        pull=[0.02,0.02], direction="clockwise",
    ))
    fig.update_layout(**BASE, height=260, showlegend=False,
        margin=dict(t=8,b=8,l=16,r=16),
        annotations=[dict(
            text=f"<b>{total}</b><br><span style='font-size:10px;color:#94a3b8'>bệnh nhân</span>",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=20, color="#0f172a"),
        )])
    return fig

def ch_gender(gen):
    if gen is None or gen.empty: return None
    clr = {"NAM":CB,"NỮ":"#f472b6","NU":"#f472b6"}
    colors = [clr.get(g.upper(), CS) for g in gen["Giới tính"]]
    fig = go.Figure(go.Bar(
        x=gen["Giới tính"], y=gen["Số lượng"],
        marker=dict(color=colors, line=dict(color="#fff",width=2)),
        text=gen["Số lượng"], textposition="outside",
        textfont=dict(size=12, family="JetBrains Mono", color="#0f172a"),
        hovertemplate="<b>%{x}</b> — %{y} người<extra></extra>",
        width=0.5,
    ))
    fig.update_layout(**BASE, height=210, margin=dict(t=8,b=8,l=8,r=8),
        xaxis=dict(showgrid=False, tickfont=dict(size=11, color="#475569")),
        yaxis=dict(gridcolor="#f8fafc", zeroline=False))
    return fig

def ch_specialty(spec):
    if spec is None or spec.empty: return None
    spec = spec.copy().sort_values("Số lượng")
    spec["lbl"] = spec["Chuyên khoa"].apply(lambda x: (x[:28]+"…") if len(x)>28 else x)
    pal = [CT,CB,CV,CA,CG,"#06b6d4","#6366f1","#f97316"]
    fig = go.Figure(go.Bar(
        y=spec["lbl"], x=spec["Số lượng"], orientation="h",
        marker=dict(color=pal[:len(spec)], line=dict(color="#fff",width=1.5)),
        text=spec["Số lượng"], textposition="outside",
        textfont=dict(size=10, family="JetBrains Mono", color="#0f172a"),
        hovertemplate="<b>%{y}</b><br>%{x} bệnh nhân<extra></extra>",
    ))
    fig.update_layout(**BASE, height=max(220, len(spec)*42),
        margin=dict(t=8,b=8,l=8,r=40),
        xaxis=dict(gridcolor="#f8fafc", zeroline=False, tickfont=dict(size=9)),
        yaxis=dict(tickfont=dict(size=9, color="#475569"), showgrid=False))
    return fig

def ch_trend(stats, color):
    if stats is None or stats.empty: return None
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=stats["Kỳ"], y=stats["Đăng ký"], name="Lượt đặt",
        marker=dict(color=color, opacity=0.7, line=dict(color="#fff",width=1)),
        hovertemplate="<b>%{x}</b><br>Đăng ký: %{y}<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=stats["Kỳ"], y=stats["Đã khám"], mode="lines+markers",
        name="Đã khám", line=dict(color=CG, width=2.5),
        marker=dict(size=5, color=CG),
        hovertemplate="<b>%{x}</b><br>Đã khám: %{y}<extra></extra>",
    ))
    fig.update_layout(**BASE, height=260,
        margin=dict(t=10,b=10,l=8,r=8),
        xaxis=dict(tickangle=-30, tickfont=dict(size=8,color="#64748b"), showgrid=False),
        yaxis=dict(gridcolor="#f1f5f9", zeroline=False, tickfont=dict(size=8)),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                    font=dict(size=10)),
        bargap=0.3)
    return fig

def ch_source_donut(noi, vl, other):
    """Donut chart for patient source breakdown."""
    labels = ["Từ khoa / Tái khám", "Bệnh nhân vãng lai"]
    values = [noi, vl]
    colors = [CV, CA]
    if other > 0:
        labels.append("Khác")
        values.append(other)
        colors.append(CS)
    fig = go.Figure(go.Pie(
        labels=labels, values=values, hole=0.65,
        marker=dict(colors=colors, line=dict(color="#fff", width=3)),
        textinfo="percent", textfont=dict(size=12),
        hovertemplate="<b>%{label}</b><br>%{value} người · %{percent}<extra></extra>",
        pull=[0.02]*len(labels), direction="clockwise",
    ))
    total = noi + vl + other
    fig.update_layout(**BASE, height=260, showlegend=False,
        margin=dict(t=8,b=8,l=16,r=16),
        annotations=[dict(
            text=f"<b>{total}</b><br><span style='font-size:10px;color:#94a3b8'>bệnh nhân</span>",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=20, color="#0f172a"),
        )])
    return fig

def ch_source_trend(df, period):
    """Stacked bar trend by source."""
    if "_date" not in df.columns or not df["_date"].notna().any(): return None
    if COL_SOURCE not in df.columns: return None
    d = df[df["_date"].notna()].copy()
    if period=="Ngày":   d["Kỳ"] = d["_date"].dt.strftime("%d/%m/%Y")
    elif period=="Tuần": d["Kỳ"] = d["_date"].apply(
        lambda x: f"T{x.isocalendar()[1]}/{x.year}")
    elif period=="Tháng": d["Kỳ"] = d["_date"].dt.strftime("Tháng %m/%Y")
    elif period=="Quý":   d["Kỳ"] = d["_date"].apply(lambda x: f"Q{((x.month-1)//3)+1}/{x.year}")
    elif period=="Năm":   d["Kỳ"] = d["_date"].dt.strftime("Năm %Y")
    src = d[COL_SOURCE].astype(str).str.strip()
    noi_mask = src.str.contains("khoa|tái|nội trú|xuất viện", case=False, na=False)
    vl_mask  = src.str.contains("vãng lai|vang lai|ngoài|ngoai", case=False, na=False)
    d["_src_noi"] = noi_mask.astype(int)
    d["_src_vl"]  = vl_mask.astype(int)
    first = d.groupby("Kỳ")["_date"].min().reset_index()
    first.columns = ["Kỳ","_s"]
    grp = d.groupby("Kỳ")[["_src_noi","_src_vl"]].sum().reset_index()
    grp = grp.merge(first, on="Kỳ").sort_values("_s").drop(columns="_s")
    fig = go.Figure()
    fig.add_trace(go.Bar(x=grp["Kỳ"], y=grp["_src_noi"], name="Từ khoa / Tái khám",
        marker=dict(color=CV, opacity=0.85), hovertemplate="<b>%{x}</b><br>Từ khoa: %{y}<extra></extra>"))
    fig.add_trace(go.Bar(x=grp["Kỳ"], y=grp["_src_vl"], name="Bệnh nhân vãng lai",
        marker=dict(color=CA, opacity=0.85), hovertemplate="<b>%{x}</b><br>Vãng lai: %{y}<extra></extra>"))
    fig.update_layout(**BASE, height=260, barmode="stack",
        margin=dict(t=10,b=10,l=8,r=8),
        xaxis=dict(tickangle=-30, tickfont=dict(size=8,color="#64748b"), showgrid=False),
        yaxis=dict(gridcolor="#f1f5f9", zeroline=False, tickfont=dict(size=8)),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                    font=dict(size=10)),
        bargap=0.3)
    return fig

# ═══════════════════════════════════════════════
# DATA FUNCTIONS
# ═══════════════════════════════════════════════
def get_credentials():
    try:
        if "gcp_service_account" in st.secrets:
            return dict(st.secrets["gcp_service_account"])
    except Exception:
        pass
    for p in [os.path.join(os.path.dirname(os.path.abspath(__file__)),"credentials.json"),
              os.path.join(os.getcwd(),"credentials.json")]:
        if os.path.exists(p):
            with open(p,"r",encoding="utf-8-sig") as f: return json.load(f)
    return None

def authenticate(src):
    if isinstance(src, str):
        creds = Credentials.from_service_account_file(src, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_info(src, scopes=SCOPES)
    return gspread.authorize(creds)

def fetch_raw(client):
    ws = client.open_by_key(SHEET_ID).worksheet(SHEET_NAME)
    vals = ws.get_all_values()
    if not vals or len(vals) < 2:
        raise ValueError("Sheet trống hoặc không có dữ liệu.")
    hdrs, rows = vals[0], vals[1:]
    seen, clean = {}, []
    for i, h in enumerate(hdrs):
        h = h.strip()
        if h == "": h = f"_c{i}"
        elif h in seen: seen[h]+=1; h=f"{h}_{seen[h]}"
        else: seen[h]=0
        clean.append(h)
    df = pd.DataFrame(rows, columns=clean)
    df = df.loc[:, ~df.columns.str.startswith("_c")]
    return df.replace("", pd.NA).dropna(how="all").fillna("")

def process(df):
    if COL_STATUS not in df.columns:
        raise KeyError(f"Không thấy cột '{COL_STATUS}'. Hiện có: {list(df.columns)}")
    df = df.copy()
    df[COL_STATUS] = df[COL_STATUS].astype(str).str.strip()
    df = df[~df[COL_STATUS].isin(["","nan","N/A","\u200b"])]
    # parse date
    if COL_EXAM_DATE in df.columns:
        df["_date"] = pd.to_datetime(
            df[COL_EXAM_DATE].astype(str).str.strip(), format="%d/%m/%Y", errors="coerce"
        )
    else:
        df["_date"] = pd.NaT
    total = len(df)
    if total == 0:
        df["_date"] = pd.NaT
        return _mk_empty(df)
    att = int((df[COL_STATUS].str.upper()==STATUS_ATTENDED.upper()).sum())
    nos = total - att

    spec = None
    if COL_SPECIALTY in df.columns:
        s = df[df[COL_SPECIALTY].astype(str).str.strip()!=""][COL_SPECIALTY]
        if not s.empty:
            spec = s.value_counts().head(8).reset_index()
            spec.columns = ["Chuyên khoa","Số lượng"]
    gen = None
    if COL_GENDER in df.columns:
        g = df[COL_GENDER].astype(str).str.strip()
        g = g[g.str.upper().isin(["NAM","NỮ","NU"])]
        if not g.empty:
            gen = g.value_counts().reset_index()
            gen.columns = ["Giới tính","Số lượng"]
    stbl = df[COL_STATUS].value_counts().reset_index()
    stbl.columns = ["Trạng thái","Số lượng"]
    # Source stats
    src_noi = src_vl = src_other = 0
    if COL_SOURCE in df.columns:
        src_vals = df[COL_SOURCE].astype(str).str.strip()
        src_noi   = int(src_vals.str.contains("khoa|tái|nội trú|xuất viện|tai", case=False, na=False).sum())
        src_vl    = int(src_vals.str.contains("vãng lai|vang lai|ngoài|ngoai", case=False, na=False).sum())
        src_other = total - src_noi - src_vl

    return dict(total=total,att=att,nos=nos,
                att_pct=round(att/total*100,1),
                nos_pct=round(nos/total*100,1),
                spec=spec, gen=gen, stbl=stbl, df=df,
                src_noi=src_noi, src_vl=src_vl, src_other=src_other)

def _mk_empty(df):
    df = df.copy(); df["_date"]=pd.NaT
    return dict(total=0,att=0,nos=0,att_pct=0.0,nos_pct=0.0,
                spec=None,gen=None,stbl=None,df=df,
                src_noi=0,src_vl=0,src_other=0)

def today_stats(df, today_date):
    """Stats for a specific date."""
    d = df[df["_date"].dt.date == today_date] if df["_date"].notna().any() else df.iloc[0:0]
    total = len(d)
    att   = int((d[COL_STATUS].str.upper()==STATUS_ATTENDED.upper()).sum()) if total>0 else 0
    return total, att, total-att, d

def build_stats(df, period):
    if "_date" not in df.columns or not df["_date"].notna().any(): return None
    d = df[df["_date"].notna()].copy()
    if period=="Ngày":       d["Kỳ"] = d["_date"].dt.strftime("%d/%m/%Y")
    elif period=="Tuần":
        d["Kỳ"] = d["_date"].apply(
            lambda x: f"T{x.isocalendar()[1]}/{x.year} ({(x-timedelta(days=x.weekday())).strftime('%d/%m')}–{(x+timedelta(days=6-x.weekday())).strftime('%d/%m')})"
        )
    elif period=="Tháng":    d["Kỳ"] = d["_date"].dt.strftime("Tháng %m/%Y")
    elif period=="Quý":      d["Kỳ"] = d["_date"].apply(lambda x: f"Q{((x.month-1)//3)+1}/{x.year}")
    elif period=="Năm":      d["Kỳ"] = d["_date"].dt.strftime("Năm %Y")
    grp   = d.groupby("Kỳ", sort=False)
    stats = grp.size().reset_index(name="Đăng ký")
    stats["Đã khám"]    = grp.apply(lambda g: (g[COL_STATUS].str.upper()==STATUS_ATTENDED.upper()).sum()).values
    stats["Vắng / Chưa"]= stats["Đăng ký"] - stats["Đã khám"]
    stats["Tỷ lệ đến (%)"]  = (stats["Đã khám"]/stats["Đăng ký"]*100).round(1)
    stats["Tỷ lệ vắng (%)"] = (stats["Vắng / Chưa"]/stats["Đăng ký"]*100).round(1)
    first = d.groupby("Kỳ")["_date"].min().reset_index(); first.columns=["Kỳ","_s"]
    return stats.merge(first,on="Kỳ").sort_values("_s").drop(columns="_s")

def source_badge(src_val):
    """Return HTML badge for patient source."""
    s = str(src_val).strip()
    if not s or s in ("nan",""):
        return ""
    if any(k in s.lower() for k in ["khoa","tái","nội trú","xuất viện","tai"]):
        return f'<span class="pt-tag src-noi">🏥 {s[:30]}</span>'
    elif any(k in s.lower() for k in ["vãng lai","vang lai","ngoài","ngoai"]):
        return f'<span class="pt-tag src-vl">🚶 {s[:30]}</span>'
    else:
        return f'<span class="pt-tag src-other">👤 {s[:30]}</span>'

def patient_card_html(row):
    """Render a mobile-friendly patient card with source badge."""
    name   = row.get(COL_NAME,"") or "—"
    dt     = row.get(COL_EXAM_DATE,"") or "—"
    spec   = row.get(COL_SPECIALTY,"") or "—"
    doc    = row.get(COL_DOCTOR,"") or ""
    status = str(row.get(COL_STATUS,""))
    ts     = row.get(COL_TIMESTAMP,"") or ""
    src    = row.get(COL_SOURCE,"") or ""
    status_cls = ("pt-status-att" if STATUS_ATTENDED.upper() in status.upper()
                  else "pt-status-nos" if status.strip() else "pt-status-oth")
    doc_tag = f'<span class="pt-tag pt-tag-doc">👨‍⚕️ {doc[:25]}</span>' if doc.strip() else ""
    src_tag = source_badge(src)
    ts_row  = f'<div style="font-size:0.62rem;color:#94a3b8;margin-top:0.15rem">🕐 {ts}</div>' if ts else ""
    return f"""<div class="pt-card">
      <div class="pt-name">{name}</div>
      <div class="pt-row">
        <span class="pt-tag pt-tag-date">📅 {dt}</span>
        <span class="pt-tag pt-tag-spec">🩺 {spec[:28]}</span>
        {doc_tag}
      </div>
      <div class="pt-row">
        <span class="pt-tag {status_cls}">{status or "—"}</span>
        {src_tag}
      </div>
      {ts_row}
    </div>"""

# ═══════════════════════════════════════════════
# SESSION + FETCH
# ═══════════════════════════════════════════════
creds_data = get_credentials()
for k,v in [("metrics",None),("fetch_time",None),("err",None)]:
    if k not in st.session_state: st.session_state[k]=v

def do_fetch():
    if not creds_data:
        st.session_state.err="⚠️ Chưa có thông tin xác thực. Kiểm tra Streamlit Secrets."
        st.session_state.metrics=None; return
    try:
        cl  = authenticate(creds_data)
        raw = fetch_raw(cl)
        st.session_state.metrics    = process(raw)
        st.session_state.fetch_time = datetime.now().strftime("%H:%M · %d/%m/%Y")
        st.session_state.err        = None
    except Exception as e:
        st.session_state.err=f"❌ {type(e).__name__}: {e}"
        st.session_state.metrics=None

if st.session_state.metrics is None and st.session_state.err is None:
    with st.spinner("Đang tải dữ liệu…"):
        do_fetch()

# ═══════════════════════════════════════════════
# NAVBAR
# ═══════════════════════════════════════════════
fetch_ts = st.session_state.fetch_time or "Chưa tải"
ok_badge = ('<span class="badge badge-ok">● Đã kết nối</span>' if creds_data
            else '<span class="badge badge-err">● Chưa kết nối</span>')
st.markdown(f"""
<div class="nb">
  <div class="nb-brand">
    <div class="nb-ico">🏥</div>
    <div>
      <div class="nb-name">BVĐK Tâm Đức Cầu Quan</div>
      <div class="nb-sub">Hệ thống theo dõi đặt khám trực tuyến</div>
    </div>
  </div>
  <div class="nb-right">
    {ok_badge}
    <span class="badge badge-time">🕐 {fetch_ts}</span>
  </div>
</div>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════
# CONTENT
# ═══════════════════════════════════════════════
st.markdown('<div class="content">', unsafe_allow_html=True)

if st.session_state.err:
    st.error(st.session_state.err)

# Refresh button — centered, not full width
rc1, rc2, rc3 = st.columns([2,1,2])
with rc2:
    if st.button("🔄 Làm mới"):
        with st.spinner("Đang tải…"):
            do_fetch()
        st.rerun()

st.markdown("<div style='margin:0.4rem 0'></div>", unsafe_allow_html=True)

# ═══════════════════════════════════════════════
# DASHBOARD
# ═══════════════════════════════════════════════
if st.session_state.metrics:
    m  = st.session_state.metrics
    df = m["df"]
    today = datetime.now().date()

    # ── TODAY BANNER ────────────────────────────
    td_total, td_att, td_nos, td_df = today_stats(df, today)
    td_pct = round(td_att/td_total*100,1) if td_total>0 else 0.0
    st.markdown(f"""
    <div class="today-banner">
      <div>
        <div class="today-label">📅 Lượt Khám Hôm Nay</div>
        <div class="today-date">{today.strftime("%A, %d/%m/%Y").replace("Monday","Thứ Hai").replace("Tuesday","Thứ Ba").replace("Wednesday","Thứ Tư").replace("Thursday","Thứ Năm").replace("Friday","Thứ Sáu").replace("Saturday","Thứ Bảy").replace("Sunday","Chủ Nhật")}</div>
      </div>
      <div class="today-stats">
        <div class="today-stat">
          <div class="today-stat-val">{td_total}</div>
          <div class="today-stat-lbl">Tổng đặt</div>
        </div>
        <div class="today-stat">
          <div class="today-stat-val" style="color:#6ee7b7">{td_att}</div>
          <div class="today-stat-lbl">Đã khám</div>
        </div>
        <div class="today-stat">
          <div class="today-stat-val" style="color:#fca5a5">{td_nos}</div>
          <div class="today-stat-lbl">Chưa khám</div>
        </div>
        <div class="today-stat">
          <div class="today-stat-val" style="color:#fcd34d">{td_pct}%</div>
          <div class="today-stat-lbl">Tỷ lệ đến</div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── TABS ────────────────────────────────────
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📊 Tổng Quan",
        "🔍 Tìm Theo Ngày",
        "📅 3 Ngày Tới",
        "🏥 Nguồn Bệnh Nhân",
        "📈 Báo Cáo",
        "👤 Bệnh Nhân",
    ])

    # ════════════════
    # TAB 1 — TỔNG QUAN
    # ════════════════
    with tab1:
        if "_date" in df.columns and df["_date"].notna().any():
            unique_dates = int(df["_date"].dropna().dt.date.nunique())
        else:
            unique_dates = 0

        st.markdown(f"""
        <div class="kg">
          <div class="kc kc-b"><div class="kc-bg">📋</div>
            <div class="kc-lbl">Tổng Lượt Đặt Khám</div>
            <div class="kc-val">{m['total']}</div>
            <div class="kc-sub">Tổng số lịch hẹn</div>
          </div>
          <div class="kc kc-g"><div class="kc-bg">✅</div>
            <div class="kc-lbl">Đã Đến Khám</div>
            <div class="kc-val">{m['att']}</div>
            <div class="kc-sub">{m['att_pct']}% tổng lượt</div>
          </div>
          <div class="kc kc-r"><div class="kc-bg">❌</div>
            <div class="kc-lbl">Vắng / Chưa Khám</div>
            <div class="kc-val">{m['nos']}</div>
            <div class="kc-sub">{m['nos_pct']}% tổng lượt</div>
          </div>
          <div class="kc kc-v"><div class="kc-bg">📅</div>
            <div class="kc-lbl">Số Ngày Có Lịch</div>
            <div class="kc-val">{unique_dates}</div>
            <div class="kc-sub">Ngày có lịch hẹn</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        # Donut + Gender — stack on mobile
        st.markdown(f'<div class="sh"><div class="sh-dot" style="background:{CG}"></div><span class="sh-txt">Tỷ Lệ Đến Khám và Vắng Mặt</span></div>', unsafe_allow_html=True)
        c1, c2 = st.columns([3,2])
        with c1:
            st.markdown('<div class="cc">', unsafe_allow_html=True)
            st.plotly_chart(ch_donut(m["att"],m["nos"],m["total"]),
                            use_container_width=True, config={"displayModeBar":False})
            st.markdown(f"""<div class="lgd">
              <div class="lgd-i"><div class="lgd-dot" style="background:{CG}"></div>
                Đã đến khám — <b style="color:#0f172a">{m['att']}</b> ({m['att_pct']}%)</div>
              <div class="lgd-i"><div class="lgd-dot" style="background:{CR}"></div>
                Vắng / Chưa — <b style="color:#0f172a">{m['nos']}</b> ({m['nos_pct']}%)</div>
            </div>""", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="sh"><div class="sh-dot" style="background:{CB}"></div><span class="sh-txt">Phân Bố Giới Tính</span></div>', unsafe_allow_html=True)
            fg = ch_gender(m["gen"])
            if fg:
                st.markdown('<div class="cc">', unsafe_allow_html=True)
                st.plotly_chart(fg, use_container_width=True, config={"displayModeBar":False})
                st.markdown('</div>', unsafe_allow_html=True)
            else:
                st.info("Không có dữ liệu giới tính.")

        fs = ch_specialty(m["spec"])
        if fs:
            st.markdown(f'<div class="sh"><div class="sh-dot" style="background:{CV}"></div><span class="sh-txt">Chuyên Khoa Được Đặt Khám Nhiều Nhất</span></div>', unsafe_allow_html=True)
            st.markdown('<div class="cc">', unsafe_allow_html=True)
            st.plotly_chart(fs, use_container_width=True, config={"displayModeBar":False})
            st.markdown('</div>', unsafe_allow_html=True)

    # ════════════════
    # TAB 2 — TÌM THEO NGÀY
    # ════════════════
    with tab2:
        st.markdown(f'<div class="sh"><div class="sh-dot" style="background:{CA}"></div><span class="sh-txt">Tìm Kiếm Lịch Hẹn Theo Ngày Khám</span></div>', unsafe_allow_html=True)

        sc1, sc2 = st.columns([1,1])
        with sc1:
            search_date = st.date_input(
                "Chọn ngày khám",
                value=today,
                format="DD/MM/YYYY",
                label_visibility="collapsed",
            )
        with sc2:
            search_status = st.selectbox(
                "Lọc trạng thái",
                ["Tất cả trạng thái"] + list(m["stbl"]["Trạng thái"].unique()),
                label_visibility="collapsed",
            )

        # Compute results
        sd_total, sd_att, sd_nos, sd_df = today_stats(df, search_date)

        if search_status != "Tất cả trạng thái":
            sd_df = sd_df[sd_df[COL_STATUS] == search_status]

        sd_pct = round(sd_att/sd_total*100,1) if sd_total>0 else 0.0

        # Mini KPI
        st.markdown(f"""
        <div class="kg" style="grid-template-columns:repeat(2,1fr);gap:0.55rem;margin:0.7rem 0 0.9rem">
          <div class="kc kc-b"><div class="kc-bg">📋</div>
            <div class="kc-lbl">Tổng Ngày {search_date.strftime('%d/%m/%Y')}</div>
            <div class="kc-val">{sd_total}</div>
            <div class="kc-sub">Lịch hẹn trong ngày</div>
          </div>
          <div class="kc kc-g"><div class="kc-bg">✅</div>
            <div class="kc-lbl">Đã Đến Khám</div>
            <div class="kc-val">{sd_att}</div>
            <div class="kc-sub">{sd_pct}% trong ngày</div>
          </div>
          <div class="kc kc-r"><div class="kc-bg">❌</div>
            <div class="kc-lbl">Vắng / Chưa</div>
            <div class="kc-val">{sd_nos}</div>
            <div class="kc-sub">{round(100-sd_pct,1) if sd_total>0 else 0}% trong ngày</div>
          </div>
          <div class="kc kc-t"><div class="kc-bg">🔍</div>
            <div class="kc-lbl">Kết Quả Lọc</div>
            <div class="kc-val">{len(sd_df)}</div>
            <div class="kc-sub">Sau khi lọc trạng thái</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        if len(sd_df) > 0:
            st.markdown(f'<div class="sh"><div class="sh-dot" style="background:{CA}"></div><span class="sh-txt">Danh Sách Bệnh Nhân Ngày {search_date.strftime("%d/%m/%Y")} ({len(sd_df)} người)</span></div>', unsafe_allow_html=True)

            show_cols = [c for c in [COL_TIMESTAMP,COL_NAME,COL_EXAM_DATE,
                                      COL_STATUS,COL_SPECIALTY,COL_GENDER]
                         if c in sd_df.columns]

            # Mobile: card view; Desktop: table view
            # Use cards always (works on both, better on mobile)
            cards_html = "".join(patient_card_html(row) for _,row in sd_df[show_cols].iterrows())
            st.markdown(cards_html, unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            csv_day = sd_df[show_cols].to_csv(index=False, encoding="utf-8-sig")
            st.download_button(
                label=f"⬇️ Tải danh sách ngày {search_date.strftime('%d/%m/%Y')} (.csv)",
                data=csv_day.encode("utf-8-sig"),
                file_name=f"lichkham_{search_date.strftime('%d%m%Y')}.csv",
                mime="text/csv",
            )
        else:
            st.markdown(f"""
            <div class="empty">
              <div class="empty-ico">📭</div>
              <div class="empty-ttl">Không có lịch hẹn ngày {search_date.strftime('%d/%m/%Y')}</div>
              <div class="empty-dsc">Thử chọn ngày khác hoặc thay đổi bộ lọc trạng thái.</div>
            </div>
            """, unsafe_allow_html=True)

    # ════════════════
    # TAB 4 — NGUỒN BỆNH NHÂN
    # ════════════════
    with tab4:
        src_noi   = m.get("src_noi", 0)
        src_vl    = m.get("src_vl", 0)
        src_other = m.get("src_other", 0)
        src_total = src_noi + src_vl + src_other

        # ── Source KPI cards ──
        st.markdown(f"""
        <div class="src-grid">
          <div class="src-card src-card-noi">
            <div class="src-card-ico">🏥</div>
            <div class="src-card-lbl">Từ Khoa / Tái Khám</div>
            <div class="src-card-val">{src_noi}</div>
            <div class="src-card-sub">{round(src_noi/src_total*100,1) if src_total>0 else 0}% tổng đăng ký</div>
          </div>
          <div class="src-card src-card-vl">
            <div class="src-card-ico">🚶</div>
            <div class="src-card-lbl">Bệnh Nhân Vãng Lai</div>
            <div class="src-card-val">{src_vl}</div>
            <div class="src-card-sub">{round(src_vl/src_total*100,1) if src_total>0 else 0}% tổng đăng ký</div>
          </div>
          <div class="src-card src-card-total">
            <div class="src-card-ico">📋</div>
            <div class="src-card-lbl">Tổng Có Nguồn</div>
            <div class="src-card-val">{src_total}</div>
            <div class="src-card-sub">{m["total"]} tổng đăng ký</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        if src_total > 0:
            # ── Donut chart ──
            st.markdown(f'<div class="sh"><div class="sh-dot" style="background:{CV}"></div><span class="sh-txt">Phân Bố Nguồn Bệnh Nhân</span></div>', unsafe_allow_html=True)
            dc1, dc2 = st.columns([3,2])
            with dc1:
                st.markdown('<div class="cc">', unsafe_allow_html=True)
                st.plotly_chart(ch_source_donut(src_noi, src_vl, src_other),
                                use_container_width=True, config={"displayModeBar":False})
                st.markdown(f"""<div class="lgd">
                  <div class="lgd-i"><div class="lgd-dot" style="background:{CV}"></div>
                    Từ khoa / Tái khám — <b style="color:#0f172a">{src_noi}</b></div>
                  <div class="lgd-i"><div class="lgd-dot" style="background:{CA}"></div>
                    Bệnh nhân vãng lai — <b style="color:#0f172a">{src_vl}</b></div>
                </div>""", unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)

            with dc2:
                # Tái khám stats breakdown
                st.markdown(f'<div class="sh"><div class="sh-dot" style="background:{CV}"></div><span class="sh-txt">Thống Kê Tái Khám</span></div>', unsafe_allow_html=True)
                if COL_SOURCE in df.columns and src_noi > 0:
                    noi_df = df[df[COL_SOURCE].astype(str).str.contains("khoa|tái|nội trú|xuất viện|tai", case=False, na=False)]
                    noi_att = int((noi_df[COL_STATUS].str.upper()==STATUS_ATTENDED.upper()).sum())
                    noi_nos = len(noi_df) - noi_att
                    noi_pct = round(noi_att/len(noi_df)*100,1) if len(noi_df)>0 else 0
                    st.markdown(f"""
                    <div style="background:white;border-radius:12px;padding:1rem;border:1px solid #e2e8f0">
                      <div style="display:flex;flex-direction:column;gap:0.7rem">
                        <div style="display:flex;justify-content:space-between;align-items:center;padding:0.5rem 0;border-bottom:1px solid #f1f5f9">
                          <span style="font-size:0.78rem;color:#475569;font-weight:500">Tổng tái khám</span>
                          <span style="font-size:1rem;font-weight:700;color:#0f172a;font-family:JetBrains Mono,monospace">{src_noi}</span>
                        </div>
                        <div style="display:flex;justify-content:space-between;align-items:center;padding:0.5rem 0;border-bottom:1px solid #f1f5f9">
                          <span style="font-size:0.78rem;color:#059669;font-weight:500">✅ Đã đến khám</span>
                          <span style="font-size:1rem;font-weight:700;color:#059669;font-family:JetBrains Mono,monospace">{noi_att}</span>
                        </div>
                        <div style="display:flex;justify-content:space-between;align-items:center;padding:0.5rem 0;border-bottom:1px solid #f1f5f9">
                          <span style="font-size:0.78rem;color:#dc2626;font-weight:500">❌ Vắng / Chưa</span>
                          <span style="font-size:1rem;font-weight:700;color:#dc2626;font-family:JetBrains Mono,monospace">{noi_nos}</span>
                        </div>
                        <div style="display:flex;justify-content:space-between;align-items:center;padding:0.5rem 0">
                          <span style="font-size:0.78rem;color:#475569;font-weight:500">📊 Tỷ lệ đến</span>
                          <span style="font-size:1rem;font-weight:700;color:#8b5cf6;font-family:JetBrains Mono,monospace">{noi_pct}%</span>
                        </div>
                      </div>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.info("Không có dữ liệu tái khám.")

            # ── Trend by period ──
            st.markdown(f'<div class="sh"><div class="sh-dot" style="background:{CA}"></div><span class="sh-txt">Xu Hướng Nguồn Bệnh Nhân Theo Tháng</span></div>', unsafe_allow_html=True)
            src_period = st.radio("Xem theo:", ["Ngày","Tuần","Tháng","Quý","Năm"],
                horizontal=True, index=2, label_visibility="collapsed", key="src_period")
            ft_src = ch_source_trend(df, src_period)
            if ft_src:
                st.markdown('<div class="cc">', unsafe_allow_html=True)
                st.plotly_chart(ft_src, use_container_width=True, config={"displayModeBar":False})
                st.markdown('</div>', unsafe_allow_html=True)

            # ── Danh sách tái khám ──
            if COL_SOURCE in df.columns and src_noi > 0:
                st.markdown(f'<div class="sh"><div class="sh-dot" style="background:{CV}"></div><span class="sh-txt">Danh Sách Bệnh Nhân Từ Khoa / Tái Khám</span></div>', unsafe_allow_html=True)
                noi_list = df[df[COL_SOURCE].astype(str).str.contains("khoa|tái|nội trú|xuất viện|tai", case=False, na=False)]
                show_src_cols = [col for col in [COL_TIMESTAMP,COL_NAME,COL_EXAM_DATE,
                                                  COL_STATUS,COL_SPECIALTY,COL_SOURCE]
                                 if col in noi_list.columns]
                MAX_SRC = 50
                cards_src = "".join(patient_card_html(row) for _,row in noi_list[show_src_cols].head(MAX_SRC).iterrows())
                st.markdown(cards_src, unsafe_allow_html=True)
                if len(noi_list) > MAX_SRC:
                    st.info(f"Hiển thị {MAX_SRC}/{len(noi_list)}. Tải CSV để xem đầy đủ.")
                csv_src = noi_list[show_src_cols].to_csv(index=False, encoding="utf-8-sig")
                st.download_button(
                    label="⬇️ Tải danh sách tái khám (.csv)",
                    data=csv_src.encode("utf-8-sig"),
                    file_name=f"taikham_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                )
        else:
            st.markdown("""<div class="empty">
              <div class="empty-ico">📭</div>
              <div class="empty-ttl">Chưa có dữ liệu nguồn bệnh nhân</div>
              <div class="empty-dsc">Vui lòng điền cột <b>NGUỒN BỆNH NHÂN</b> trong Google Sheet.<br>
              Giá trị gợi ý: "Từ khoa / Tái khám" hoặc "Bệnh nhân vãng lai".</div>
            </div>""", unsafe_allow_html=True)

    # ════════════════
    # TAB 5 — BÁO CÁO
    # ════════════════
    with tab5:
        period_opts   = ["Ngày","Tuần","Tháng","Quý","Năm"]
        period_colors = [CB,CT,CV,CA,CG]
        period_icons  = ["📆","🗓️","🗂️","📊","🏆"]

        sel_p = st.radio("Xem theo:", period_opts,
                         horizontal=True, index=2, label_visibility="collapsed")
        p_col  = period_colors[period_opts.index(sel_p)]
        p_ico  = period_icons[period_opts.index(sel_p)]
        stats  = build_stats(df, sel_p)

        tot_p = int(stats["Đăng ký"].sum()) if stats is not None else 0
        att_p = int(stats["Đã khám"].sum()) if stats is not None else 0
        pct_p = round(att_p/tot_p*100,1) if tot_p>0 else 0.0
        peak_p= int(stats["Đăng ký"].max()) if stats is not None else 0

        st.markdown(f"""
        <div class="kg" style="grid-template-columns:repeat(2,1fr);gap:0.55rem;margin-bottom:1rem">
          <div class="kc kc-b"><div class="kc-bg">{p_ico}</div>
            <div class="kc-lbl">Tổng Lượt — {sel_p}</div>
            <div class="kc-val">{tot_p}</div>
            <div class="kc-sub">Toàn bộ kỳ</div>
          </div>
          <div class="kc kc-g"><div class="kc-bg">✅</div>
            <div class="kc-lbl">Đã Đến Khám</div>
            <div class="kc-val">{att_p}</div>
            <div class="kc-sub">{pct_p}% tổng lượt</div>
          </div>
          <div class="kc kc-r"><div class="kc-bg">📉</div>
            <div class="kc-lbl">Vắng / Chưa</div>
            <div class="kc-val">{tot_p-att_p}</div>
            <div class="kc-sub">{round(100-pct_p,1)}% tổng lượt</div>
          </div>
          <div class="kc kc-t"><div class="kc-bg">🔝</div>
            <div class="kc-lbl">Kỳ Đông Nhất</div>
            <div class="kc-val">{peak_p}</div>
            <div class="kc-sub">Lượt đặt cao nhất</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        if stats is not None and not stats.empty:
            st.markdown(f'<div class="sh"><div class="sh-dot" style="background:{p_col}"></div><span class="sh-txt">Biểu Đồ Xu Hướng Theo {sel_p}</span></div>', unsafe_allow_html=True)
            st.markdown('<div class="cc">', unsafe_allow_html=True)
            ft = ch_trend(stats[["Kỳ","Đăng ký","Đã khám"]], p_col)
            if ft:
                st.plotly_chart(ft, use_container_width=True, config={"displayModeBar":False})
            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown(f'<div class="sh"><div class="sh-dot" style="background:{p_col}"></div><span class="sh-txt">Bảng Chi Tiết Theo {sel_p}</span></div>', unsafe_allow_html=True)
            rows_html = ""
            for _, row in stats.iterrows():
                g_cls = "pct-g" if row["Tỷ lệ đến (%)"]>=50 else "pct-r"
                r_cls = "pct-r" if row["Tỷ lệ vắng (%)"]>=50 else "pct-g"
                rows_html += f"""<tr>
                  <td>{row['Kỳ']}</td>
                  <td class="num">{int(row['Đăng ký'])}</td>
                  <td class="num" style="color:#059669">{int(row['Đã khám'])}</td>
                  <td class="num" style="color:#dc2626">{int(row['Vắng / Chưa'])}</td>
                  <td class="{g_cls}">{row['Tỷ lệ đến (%)']}%</td>
                  <td class="{r_cls}">{row['Tỷ lệ vắng (%)']}%</td>
                </tr>"""
            st.markdown(f"""
            <div class="rtbl-wrap">
              <table class="rtbl"><thead><tr>
                <th>Kỳ</th><th>Tổng</th><th>Đã Khám</th>
                <th>Vắng</th><th>% Đến</th><th>% Vắng</th>
              </tr></thead><tbody>{rows_html}</tbody></table>
            </div>
            <div class="scroll-hint">← Vuốt ngang để xem thêm →</div>
            """, unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            csv_r = stats.to_csv(index=False, encoding="utf-8-sig")
            st.download_button(
                label=f"⬇️ Tải báo cáo theo {sel_p} (.csv)",
                data=csv_r.encode("utf-8-sig"),
                file_name=f"baocao_{sel_p.lower()}_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
            )
        else:
            st.markdown("""<div class="empty">
              <div class="empty-ico">📭</div>
              <div class="empty-ttl">Không có dữ liệu</div>
              <div class="empty-dsc">Dữ liệu ngày khám chưa được nhập hoặc sai định dạng dd/mm/yyyy.</div>
            </div>""", unsafe_allow_html=True)

    # ════════════════
    # TAB 6 — BỆNH NHÂN
    # ════════════════
    with tab6:
        show_cols = [c for c in [COL_TIMESTAMP,COL_NAME,COL_EXAM_DATE,
                                  COL_STATUS,COL_SPECIALTY,COL_GENDER,COL_SOURCE]
                     if c in df.columns]

        st.markdown(f'<div class="sh"><div class="sh-dot" style="background:{CA}"></div><span class="sh-txt">Lọc Danh Sách Bệnh Nhân</span></div>', unsafe_allow_html=True)
        fc1, fc2 = st.columns(2)
        with fc1:
            s_status = st.selectbox("Trạng thái",
                ["Tất cả"] + list(m["stbl"]["Trạng thái"].unique()),
                label_visibility="collapsed")
        with fc2:
            if COL_SPECIALTY in df.columns:
                specs = sorted(df[df[COL_SPECIALTY].str.strip()!=""][COL_SPECIALTY].unique().tolist())
                s_spec = st.selectbox("Chuyên khoa",
                    ["Tất cả chuyên khoa"] + specs, label_visibility="collapsed")
            else:
                s_spec = "Tất cả chuyên khoa"

        fdf = df.copy()
        if s_status != "Tất cả":
            fdf = fdf[fdf[COL_STATUS]==s_status]
        if s_spec != "Tất cả chuyên khoa" and COL_SPECIALTY in fdf.columns:
            fdf = fdf[fdf[COL_SPECIALTY]==s_spec]

        st.markdown(f"""
        <div style="margin:0.5rem 0 0.7rem;padding:0.45rem 0.85rem;
             background:white;border-radius:10px;border:1px solid #e2e8f0;
             font-size:0.78rem;color:#475569">
          Hiển thị <b style="color:#0f172a">{len(fdf)}</b> / {m['total']} bệnh nhân
        </div>""", unsafe_allow_html=True)

        # Patient cards (mobile-friendly)
        if len(fdf) > 0:
            MAX_CARDS = 50
            show_df = fdf[show_cols].head(MAX_CARDS)
            cards = "".join(patient_card_html(row) for _,row in show_df.iterrows())
            st.markdown(cards, unsafe_allow_html=True)
            if len(fdf) > MAX_CARDS:
                st.info(f"Hiển thị {MAX_CARDS}/{len(fdf)} bệnh nhân. Tải file CSV để xem đầy đủ.")
        else:
            st.markdown("""<div class="empty">
              <div class="empty-ico">🔍</div>
              <div class="empty-ttl">Không tìm thấy bệnh nhân</div>
              <div class="empty-dsc">Thử thay đổi bộ lọc.</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        csv2 = fdf[show_cols].to_csv(index=False, encoding="utf-8-sig")
        st.download_button(
            label="⬇️ Tải danh sách bệnh nhân (.csv)",
            data=csv2.encode("utf-8-sig"),
            file_name=f"benhnhan_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
        )

else:
    if not st.session_state.err:
        st.markdown("""<div class="empty">
          <div class="empty-ico">🏥</div>
          <div class="empty-ttl">Đang tải dữ liệu…</div>
          <div class="empty-dsc">Nếu không thấy sau vài giây,
            nhấn <strong>🔄 Làm mới</strong> ở trên.</div>
        </div>""", unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)
