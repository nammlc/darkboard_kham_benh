"""
BVĐK Tâm Đức Cầu Quan — Hệ Thống Theo Dõi Đặt Khám Trực Tuyến
Mobile-first · No sidebar · Today's stats · Date search
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import gspread
from google.oauth2.service_account import Credentials
import json, os, re
import openpyxl
from datetime import datetime, timedelta, date
from collections import Counter

SHEET_ID        = "1EYiRA3ar41aue8DlbWA7JTKoLL0M2tiLTcZINhdMfTs"
SHEET_NAME      = "Câu trả lời biểu mẫu 1"
COL_STATUS      = "TRẠNG THÁI"
COL_EXAM_DATE   = "NGÀY KHÁM"
COL_NAME        = "1. HỌ VÀ TÊN BỆNH NHÂN"
COL_GENDER      = "3. GIỚI TÍNH"
COL_SPECIALTY   = "CHUYÊN KHOA MONG MUỐN KHÁM"
COL_TIMESTAMP   = "Dấu thời gian"
COL_DOCTOR      = "BÁC SĨ MONG MUỐN ( nếu có)"
COL_SOURCE      = "NGUỒN BỆNH NHÂN"
COL_PHONE       = "5. SỐ ĐIÊN THOẠI"
COL_BIRTH_YEAR  = "NĂM SINH"
COL_EXAM_TIME   = "GIỜ KHÁM DỰ KIẾN"
COL_KHOA        = "KHOA KHÁM CHỮA BỆNH"
STATUS_ATTENDED     = "BỆNH NHÂN ĐÃ KHÁM"
STATUS_NOT_ATTENDED = "BỆNH NHÂN CHƯA KHÁM / BỎ KHÁM"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
SCOPES_RO = [
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

/* Breakdown chips (Khoa Khám Bệnh vs Khoa khác) trong thẻ ngày */
.upcoming-day-stats {
    display:grid; grid-template-columns:1fr 1fr;
    gap:0.55rem; padding:0.8rem 1rem 0.2rem;
}
.uds-item {
    border-radius:11px; padding:0.6rem 0.75rem;
    display:flex; flex-direction:column; gap:0.15rem;
}
.uds-kb   { background:#eff6ff; border:1px solid #bfdbfe; }
.uds-khac { background:#f5f3ff; border:1px solid #ddd6fe; }
.uds-val  { font-family:'JetBrains Mono',monospace !important; font-weight:700; font-size:1.35rem; line-height:1; }
.uds-kb   .uds-val { color:#1d4ed8 !important; }
.uds-khac .uds-val { color:#6d28d9 !important; }
.uds-lbl  { font-size:0.62rem; font-weight:600; color:#475569; line-height:1.35; margin-top:0.15rem; }
.upcoming-day-actions { padding:0.35rem 1rem 0.9rem; }

/* Nút "Xem chi tiết" trong thẻ ngày — nổi bật, dạng pill gradient */
.upcoming-day-actions .stButton>button {
    background:linear-gradient(135deg,#0f4c75,#1b6ca8) !important;
    color:white !important; border:none !important;
    border-radius:10px !important; font-weight:700 !important;
    font-size:0.8rem !important; padding:0.5rem 1rem !important;
    box-shadow:0 2px 10px rgba(15,76,117,0.25) !important;
}
.upcoming-day-actions .stButton>button:hover {
    filter:brightness(1.08);
    box-shadow:0 4px 14px rgba(15,76,117,0.35) !important;
}

/* Badge Khoa trong bảng chi tiết */
.khoa-badge {
    font-size:0.62rem; font-weight:600; padding:0.2rem 0.55rem;
    border-radius:20px; white-space:normal; display:inline-block;
}
.khoa-kb   { background:#eff6ff; color:#1d4ed8; }
.khoa-khac { background:#ede9fe; color:#5b21b6; }
.khoa-none { background:#f1f5f9; color:#64748b; }

/* Header nhóm trong dialog */
.dlg-group-hd {
    display:flex; align-items:center; justify-content:space-between;
    background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px;
    padding:0.55rem 0.85rem; margin-bottom:0.6rem;
}
.dlg-group-hd b { font-size:0.82rem; color:#1e293b; }
.dlg-group-hd span { font-size:0.72rem; color:#64748b; }

/* ── QUẢN LÝ: SỬA / XÓA BỆNH NHÂN (tab "3 Ngày Tới") ── */
.mrow-section-hd {
    display:flex; align-items:center; gap:0.4rem;
    margin:0.9rem 0 0.55rem;
    font-size:0.76rem; font-weight:700; color:#1e293b;
    background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px;
    padding:0.55rem 0.9rem;
}
div[class*="_mrow_"] {
    background:white; border:1px solid #e2e8f0; border-radius:14px;
    padding:0.8rem 0.9rem 0.7rem; margin-bottom:0.65rem;
    box-shadow:0 1px 3px rgba(0,0,0,0.05);
    transition:box-shadow 0.15s;
}
div[class*="_mrow_"]:hover { box-shadow:0 3px 10px rgba(0,0,0,0.08); }
div[class*="_mrow_"] div[data-testid="stHorizontalBlock"] { align-items:center; gap:0.5rem; }
/* Hàng info + badge trạng thái (hàng trên) */
div[class*="_mrow_"] > div > div[data-testid="stHorizontalBlock"]:first-of-type { margin-bottom:0.55rem; }
.mrow-info { display:flex; flex-direction:column; gap:0.3rem; }
.mrow-name { font-size:0.86rem; font-weight:700; color:#0f172a; }
.mrow-sub  { font-size:0.68rem; color:#64748b; font-family:'JetBrains Mono',monospace !important; }
.mrow-badge {
    display:inline-block; font-size:0.64rem; font-weight:700;
    padding:0.22rem 0.6rem; border-radius:20px; white-space:nowrap;
}
.mrow-badge-att { background:#d1fae5; color:#065f46; }
.mrow-badge-nos { background:#fee2e2; color:#991b1b; }
.mrow-status-wrap { display:flex; align-items:flex-start; justify-content:flex-end; padding:0.05rem 0; }
.mrow-status-wrap .mrow-badge { text-align:center; white-space:normal; line-height:1.3; }

/* Vạch phân cách nhẹ giữa thông tin bệnh nhân và 2 nút hành động,
   giúp mắt tách bạch rõ khu vực xem thông tin và khu vực thao tác. */
.mrow-actions {
    border-top:1px dashed #e2e8f0; padding-top:0.55rem; margin-top:0.15rem;
}

/* Nút "✏️ Sửa trạng thái" — xanh dương, full-width, rõ chữ */
div[class*="editbtn_"] .stButton>button {
    background:#eff6ff !important;
    color:#1d4ed8 !important; border:1.5px solid #bfdbfe !important;
    border-radius:9px !important; font-weight:700 !important;
    font-size:0.74rem !important; padding:0.5rem 0.4rem !important;
    transform:none !important; box-shadow:none !important;
}
div[class*="editbtn_"] .stButton>button:hover {
    background:#dbeafe !important; border-color:#93c5fd !important; transform:none !important;
}

/* Nút "🗑️ Xóa bệnh nhân" — đỏ nhạt, full-width, rõ chữ */
div[class*="delbtn_"] .stButton>button {
    background:#fef2f2 !important;
    color:#b91c1c !important; border:1.5px solid #fecaca !important;
    border-radius:9px !important; font-weight:700 !important;
    font-size:0.74rem !important; padding:0.5rem 0.4rem !important;
    transform:none !important; box-shadow:none !important;
}
div[class*="delbtn_"] .stButton>button:hover {
    background:#fee2e2 !important; border-color:#fca5a5 !important; transform:none !important;
}

/* Popup "Sửa Trạng Thái Bệnh Nhân" */
.edit-dlg-name {
    font-size:0.92rem; font-weight:700; color:#0f172a;
    background:#eff6ff; border:1px solid #bfdbfe; border-radius:10px;
    padding:0.65rem 0.9rem; margin-bottom:0.9rem; text-align:center;
}
div[data-testid="stDialog"] div[role="radiogroup"] {
    display:flex; flex-direction:column; gap:0.55rem;
}
div[data-testid="stDialog"] div[role="radiogroup"] label {
    background:#f8fafc; border:1.5px solid #e2e8f0 !important; border-radius:11px !important;
    padding:0.75rem 0.9rem !important; width:100%;
    font-size:0.82rem !important; font-weight:600 !important; color:#1e293b !important;
    transition:all 0.15s;
}
div[data-testid="stDialog"] div[role="radiogroup"] label:hover {
    background:#eff6ff; border-color:#93c5fd !important;
}
div[data-testid="stDialog"] div[role="radiogroup"] label[data-checked="true"] {
    background:#dbeafe !important; border-color:#3b82f6 !important;
}

/* Popup xác nhận xóa */
.del-dlg-warn {
    text-align:center; font-size:0.86rem; color:#334155; line-height:1.75;
    background:#fef2f2; border:1px solid #fecaca; border-radius:11px;
    padding:0.95rem 1rem; margin-bottom:0.9rem;
}
.del-dlg-warn b { color:#991b1b; }
.del-dlg-note { display:block; font-size:0.7rem; color:#dc2626; font-weight:700; margin-top:0.35rem; }

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

/* ── PAGINATION (kiểu viên thuốc bo tròn, tách biệt trang hiện tại) ── */
.pg-info {
    text-align:center; font-size:0.72rem; color:#64748b;
    margin:0.5rem 0 0.35rem; font-weight:500;
}
.pg-info b { color:#0f172a; font-family:'JetBrains Mono',monospace; }
div[data-testid="stHorizontalBlock"] div[data-testid="column"] .stButton>button {
    width:100%;
}

/* Khung bo tròn bao quanh cả dải nút phân trang.
   Dùng [class*="_pgrow"] thay vì phụ thuộc vào type="primary" của Streamlit,
   vì cách đó không phải lúc nào cũng bắt được bằng CSS ở mọi phiên bản. */
div[class*="_pgrow"] {
    background: #ffffff;
    border: 1px solid #e6eaf1;
    border-radius: 999px;
    padding: 0.22rem 0.3rem;
    box-shadow: 0 2px 10px rgba(15,23,42,0.06);
    margin: 0 auto 0.6rem;
    width: fit-content;
    max-width: 100%;
}
div[class*="_pgrow"] div[data-testid="stHorizontalBlock"] {
    gap: 0.18rem !important;
}

/* Nút phân trang mặc định: nhỏ gọn, dạng viên thuốc, trong suốt */
div[class*="_pgrow"] .stButton>button,
div[class*="_pgrow"] .stButton>button:focus {
    padding: 0 0.15rem !important;
    min-height: 1.7rem !important;
    height: 1.7rem !important;
    min-width: 1.7rem !important;
    line-height: 1.7rem !important;
    font-size: 0.7rem !important;
    font-weight: 600 !important;
    border-radius: 999px !important;
    box-shadow: none !important;
    border: none !important;
    background: transparent !important;
    color: #475569 !important;
    transform: none !important;
}
div[class*="_pgrow"] .stButton>button:hover:not(:disabled) {
    background: #f1f5f9 !important;
    color: #0f172a !important;
}
/* Nút ‹ › khi bị disabled (đã ở trang đầu/cuối) — làm mờ đi để biết không bấm được */
div[class*="_navbtn"] .stButton>button:disabled {
    opacity: 0.32 !important;
    color: #94a3b8 !important;
    background: transparent !important;
}
/* Trang HIỆN TẠI — nền đậm, chữ trắng, luôn nổi bật kể cả khi disabled.
   Đặt sau cùng + bám riêng vào container "_curbtn" nên không đụng độ nút khác. */
div[class*="_curbtn"] .stButton>button,
div[class*="_curbtn"] .stButton>button:disabled,
div[class*="_curbtn"] .stButton>button:hover {
    background: linear-gradient(135deg,#0f172a,#1e3a5f) !important;
    color: #ffffff !important;
    opacity: 1 !important;
    cursor: default !important;
}

/* == FORCE LIGHT MODE — disable dark theme == */
.stApp { background:#f0f4f8 !important; color:#1e293b !important; }
[data-testid="stAppViewContainer"] { background:#f0f4f8 !important; }
[data-testid="themeToggle"] { display:none !important; }
button[kind="header"] { display:none !important; }
.pt-card, .kc, .cc, .src-card, .upcoming-day,
.empty, div[data-testid="stExpander"] {
    background: white !important;
}
.pt-name, .kc-val, .sh-txt, .empty-ttl,
.src-card-val { color: #0f172a !important; }
.kc-lbl, .kc-sub, .empty-dsc, .src-card-lbl,
.src-card-sub { color: #94a3b8 !important; }
.rtbl td { color: #1e293b !important; background: white !important; }
.rtbl tr:nth-child(even) td { background: #f8fafc !important; }
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

def authenticate_rw(src):
    """Authenticate with WRITE scope for importing data."""
    if isinstance(src, str):
        creds = Credentials.from_service_account_file(src, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_info(src, scopes=SCOPES)
    return gspread.authorize(creds)


def _fix_phone(raw: str) -> str:
    """Chuẩn hoá số điện thoại từ Excel:
    - Loại bỏ khoảng trắng, dấu chấm, dấu gạch ngang
    - Nếu số có 9 chữ số (Excel tự bỏ số 0 đầu) → thêm lại '0'
    - Nếu đã có 10 chữ số hoặc rỗng → giữ nguyên
    """
    phone = re.sub(r"[\s.\-]", "", str(raw).strip())
    if not phone or phone.lower() in ("n/a", "none", ""):
        return "N/A"
    digits = re.sub(r"\D", "", phone)
    if len(digits) == 9:
        digits = "0" + digits
    return digits if digits else "N/A"


def parse_minh_lo_excel(uploaded_file):
    """
    Parse Minh Lo HIS Excel export using direct XML parsing.
    Robust against missing sharedStrings.xml (files with only numeric cells).
    Handles merged cells via fill-down logic.
    Returns (list_of_dicts, error_msg_or_None).
    """
    import zipfile
    import xml.etree.ElementTree as ET
    import io

    NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    def _tag(name):
        return "{" + NS + "}" + name

    try:
        # Read file bytes (works for both Streamlit UploadedFile and file path)
        if hasattr(uploaded_file, "read"):
            raw = uploaded_file.read()
        else:
            with open(uploaded_file, "rb") as f:
                raw = f.read()

        zf = zipfile.ZipFile(io.BytesIO(raw))
        # Một số phần mềm (vd. Minh Lộ HIS) nén zip với dấu '\' thay vì '/'
        # trong tên đường dẫn nội bộ (không đúng chuẩn zip nhưng vẫn mở được).
        # Chuẩn hoá về '/' để so khớp đúng "worksheets/sheet1", "sharedstrings"...
        names_lower = {n.lower().replace("\\", "/"): n for n in zf.namelist()}

        # ── Shared strings (optional — some xlsx have none) ──
        shared = []
        ss_key = next((k for k in names_lower if "sharedstrings" in k), None)
        if ss_key:
            with zf.open(names_lower[ss_key]) as f:
                tree = ET.parse(f)
            for si in tree.findall(".//" + _tag("si")):
                texts = [t.text or "" for t in si.findall(".//" + _tag("t"))]
                shared.append("".join(texts))

        # ── Worksheet ──
        ws_key = next(
            (k for k in names_lower if "worksheets/sheet1" in k or "worksheet/sheet1" in k),
            None,
        )
        if ws_key is None:
            # fallback: find any sheet
            ws_key = next((k for k in names_lower if "worksheets/sheet" in k), None)
        if ws_key is None:
            return [], "Không tìm thấy worksheet trong file xlsx."

        with zf.open(names_lower[ws_key]) as f:
            ws_tree = ET.parse(f)

        # ── Merged cells → build fill map {(row,col): (src_row,src_col)} ──
        merge_fill = {}  # (r,c) → value to fill from top-left of merge
        mc_el = ws_tree.find(".//" + _tag("mergeCells"))
        if mc_el is not None:
            for mc in mc_el.findall(_tag("mergeCell")):
                ref = mc.get("ref", "")
                if ":" not in ref:
                    continue
                p1, p2 = ref.split(":")
                def col_num(s):
                    s = "".join(ch for ch in s if ch.isalpha()).upper()
                    n = 0
                    for ch in s:
                        n = n * 26 + (ord(ch) - 64)
                    return n
                def row_num(s):
                    return int("".join(ch for ch in s if ch.isdigit()))
                r1, c1 = row_num(p1), col_num(p1)
                r2, c2 = row_num(p2), col_num(p2)
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        if r != r1 or c != c1:
                            merge_fill[(r, c)] = (r1, c1)

        # ── Read all cells into grid ──
        grid = {}  # (row, col) → raw value
        for row_el in ws_tree.findall(".//" + _tag("row")):
            r = int(row_el.get("r", 0))
            for cell_el in row_el.findall(_tag("c")):
                ref = cell_el.get("r", "")
                col_str = "".join(ch for ch in ref if ch.isalpha())
                col_n = 0
                for ch in col_str.upper():
                    col_n = col_n * 26 + (ord(ch) - 64)
                t   = cell_el.get("t", "")
                v_el = cell_el.find(_tag("v"))
                if v_el is None:
                    val = ""
                elif t == "s":
                    idx2 = int(v_el.text or 0)
                    val = shared[idx2] if idx2 < len(shared) else ""
                elif t == "inlineStr":
                    is_el = cell_el.find(_tag("is"))
                    val = is_el.findtext(_tag("t"), "") if is_el is not None else ""
                else:
                    val = v_el.text or ""
                grid[(r, col_n)] = val

        # Apply merge fill
        for (r, c), (sr, sc) in merge_fill.items():
            if (sr, sc) in grid and (r, c) not in grid:
                grid[(r, c)] = grid[(sr, sc)]

        max_row = max(r for r, _ in grid) if grid else 0
        max_col = max(c for _, c in grid) if grid else 0

        def get_row(r):
            return [grid.get((r, c), "") for c in range(1, max_col + 1)]

        # ── Find header row (row with STT and Mã y tế) ──
        header_row_idx = None
        for r in range(1, min(max_row + 1, 15)):
            row_vals = [str(v).strip() for v in get_row(r)]
            has_stt = "STT" in row_vals
            has_ma  = any("m" in v.lower() and "y t" in v.lower() for v in row_vals)
            if has_stt and has_ma:
                header_row_idx = r
                break

        if header_row_idx is None:
            return [], "Không tìm thấy hàng tiêu đề trong file. Kiểm tra đúng loại báo cáo Minh Lộ."

        headers = [str(v).strip().replace("\n", " ").lower() for v in get_row(header_row_idx)]

        def find_col(keywords):
            for i, h in enumerate(headers):
                if any(k.lower() in h for k in keywords):
                    return i
            return None

        # ── Cột "Tuổi" trong file Minh Lộ thực chất là 2 cột con "Nam" / "Nữ"
        # (mỗi bệnh nhân chỉ có 1 trong 2 cột này có giá trị, tùy giới tính).
        # Nhãn "Nam" / "Nữ" nằm ở DÒNG PHỤ ngay dưới dòng tiêu đề chính,
        # không nằm trong dòng tiêu đề chính (dòng tiêu đề chính chỉ ghi
        # "Tuổi" 1 lần cho cả 2 cột). Vì vậy phải dò trực tiếp trong dòng
        # phụ đó thay vì dò từ khóa "nam"/"nữ" trong dòng tiêu đề chính
        # (cách cũ luôn thất bại vì dòng tiêu đề chính không chứa các từ này).
        sub_row_idx = header_row_idx + 1
        sub_row_vals = (
            [str(v).strip().lower() for v in get_row(sub_row_idx)]
            if sub_row_idx <= max_row else []
        )
        tuoi_nam_col = next((i for i, v in enumerate(sub_row_vals) if v == "nam"), None)
        tuoi_nu_col  = next((i for i, v in enumerate(sub_row_vals) if v in ("nữ", "nu")), None)

        # Fallback: nếu không tìm thấy dòng phụ (một số bản export có định
        # dạng khác), thử dò từ khóa như cách cũ để không bị hỏng hoàn toàn.
        if tuoi_nam_col is None and tuoi_nu_col is None:
            tuoi_nam_col = find_col(["tuổi nam", "nam"])
            tuoi_nu_col  = find_col(["tuổi nữ", "nữ", "nu"])

        idx = {
            "ma_yt":      find_col(["mã y tế", "ma y te"]),
            "ho_ten":     find_col(["họ tên", "ho ten", "bệnh nhân"]),
            "tuoi_nam":   tuoi_nam_col,
            "tuoi_nu":    tuoi_nu_col,
            "dia_chi":    find_col(["địa chỉ", "dia chi"]),
            "bhyt":       find_col(["bhyt"]),
            "bac_sy":     find_col(["bác sỹ khám", "bac sy kham", "bác sĩ khám"]),
            "trieu_chung":find_col(["triệu chứng", "trieu chung"]),
            "chan_doan":   find_col(["chẩn đoán", "chan doan"]),
            "ngay_hen":   find_col(["ngày hẹn", "ngay hen"]),
            "ngay_lap":   find_col(["ngày lập", "ngay lap"]),
            "dt":         find_col(["điện thoại", "dien thoai"]),
            "khoa_hen":   find_col(["khoa hẹn", "khoa hen"]),
            "bac_sy_hen": find_col(["bác sỹ hẹn", "bac sy hen", "bác sĩ hẹn"]),
            "da_kham":    find_col(["đã khám", "da kham"]),
        }

        def cv(row_vals, key):
            i = idx.get(key)
            if i is None or i >= len(row_vals):
                return ""
            return str(row_vals[i]).strip()

        def to_date(val):
            val = str(val).strip()
            if not val or val in ("None", ""):
                return ""
            if re.search(r"\d{1,2}/\d{1,2}/\d{4}", val):
                return re.search(r"\d{1,2}/\d{1,2}/\d{4}", val).group()
            if re.match(r"\d{4}-\d{2}-\d{2}", val):
                return datetime.strptime(val[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
            # Excel serial
            try:
                s = float(val)
                if 40000 < s < 60000:
                    return (datetime(1899, 12, 30) + timedelta(days=int(s))).strftime("%d/%m/%Y")
            except Exception:
                pass
            return val

        # ── Locate the STT (sequential number) column ──
        stt_col = headers.index("stt") if "stt" in headers else 0

        def has_letter(s):
            return bool(re.search(r"[^\W\d_]", str(s), re.UNICODE))

        # ── Detect the start row of each patient block ──
        # Minh Lộ exports wrap long fields (Họ tên, Địa chỉ, Chỉ định điều
        # trị…) across SEVERAL physical Excel rows per patient (not just one
        # row each). A row is the START of a new patient block when its STT
        # cell is a plain number AND its "Họ tên" cell contains a real name
        # (i.e. has letters) — this also filters out junk "column index"
        # rows some exports insert right after the header (a row that just
        # reads 1, 2, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13… under each column).
        start_rows = []
        for r in range(header_row_idx + 1, max_row + 1):
            row_vals = get_row(r)
            stt_val = str(row_vals[stt_col]).strip() if stt_col < len(row_vals) else ""
            ho_ten_val = cv(row_vals, "ho_ten")
            if stt_val.isdigit() and has_letter(ho_ten_val):
                start_rows.append(r)

        # Fallback for a different report layout: if no block-start row was
        # detected this way, treat every non-empty row as its own record
        # (old behaviour) so we never silently return zero results.
        if not start_rows:
            start_rows = [
                r for r in range(header_row_idx + 1, max_row + 1)
                if any(str(v).strip() for v in get_row(r))
            ]

        def join_field(r0, r1, key):
            """Rebuild a long field spanning multiple physical rows.
            Minh Lộ splits tên bệnh nhân across rows (e.g. 'NGUYỄN' /
            ' THỊ' / ' PHƯƠNG'), join with space then collapse extras."""
            parts = [cv(get_row(r), key).strip() for r in range(r0, r1 + 1)]
            joined = " ".join(p for p in parts if p)
            return re.sub(r"\s+", " ", joined).strip()

        data_rows = []
        for i, r0 in enumerate(start_rows):
            r1 = (start_rows[i + 1] - 1) if i + 1 < len(start_rows) else max_row
            vals0 = get_row(r0)

            ma_yt   = cv(vals0, "ma_yt")
            ho_ten  = join_field(r0, r1, "ho_ten")
            dia_chi = join_field(r0, r1, "dia_chi")

            if not ma_yt and not ho_ten:
                continue

            ngay_raw = cv(vals0, "ngay_hen")
            gio_hen  = ""
            try:
                s = float(ngay_raw)
                if 40000 < s < 60000:
                    frac = s - int(s)
                    if frac > 0.0:
                        sec = int(frac * 86400)
                        gio_hen = f"{sec // 3600:02d}:{(sec % 3600) // 60:02d}"
            except Exception:
                pass

            tuoi_val = cv(vals0, "tuoi_nam") or cv(vals0, "tuoi_nu")
            gioi_tinh = ""
            if cv(vals0, "tuoi_nam"):
                gioi_tinh = "Nam"
            elif cv(vals0, "tuoi_nu"):
                gioi_tinh = "Nữ"

            nam_sinh = ""
            if tuoi_val:
                nums = re.findall(r"\d+", tuoi_val)
                if nums:
                    n = int(nums[0])
                    try:
                        if "tháng" in tuoi_val.lower():
                            # Trẻ nhỏ tính theo tháng tuổi (vd. "32 tháng tuổi")
                            if 0 <= n < 1200:
                                years_old = n // 12
                                if 0 <= years_old < 120:
                                    nam_sinh = str(datetime.now().year - years_old)
                        else:
                            # Người lớn tính theo năm tuổi (vd. "74 tuổi")
                            if 0 < n < 120:
                                nam_sinh = str(datetime.now().year - n)
                    except Exception:
                        pass

            data_rows.append({
                "MÃ Y TẾ":             ma_yt,
                "HỌ TÊN":              ho_ten,
                "NĂM SINH (ước tính)": nam_sinh,
                "GIỚI TÍNH":           gioi_tinh,
                "ĐỊA CHỈ":             dia_chi,
                "SỐ BHYT":             cv(vals0, "bhyt"),
                "BÁC SĨ KHÁM":         cv(vals0, "bac_sy"),
                "TRIỆU CHỨNG":         cv(vals0, "trieu_chung"),
                "CHẨN ĐOÁN":           cv(vals0, "chan_doan"),
                "NGÀY LẬP":            to_date(cv(vals0, "ngay_lap")),
                "NGÀY HẸN":            to_date(ngay_raw),
                "GIỜ HẸN":             gio_hen,
                "SỐ ĐIỆN THOẠI":       _fix_phone(cv(vals0, "dt")),
                "KHOA HẸN":            cv(vals0, "khoa_hen"),
                "BÁC SĨ HẸN":          cv(vals0, "bac_sy_hen"),
                "ĐÃ KHÁM":             cv(vals0, "da_kham"),
            })

        return data_rows, None

    except Exception as e:
        import traceback
        return [], f"Lỗi đọc file: {type(e).__name__}: {e}"

# Full ordered column list of the Google Sheet (khớp đúng thứ tự cột thực tế)
SHEET_COLUMNS = [
    "Dấu thời gian",                                                    # A  - import timestamp
    "NGUỒN BỆNH NHÂN",                                                  # B  - fixed value
    "TRẠNG THÁI",                                                       # C  - blank
    "NGÀY KHÁM",                                                        # D  - Ngày hẹn
    "1. HỌ VÀ TÊN BỆNH NHÂN",                                          # E  - Họ tên
    "NĂM SINH",                                                         # F  - N/A
    "5. SỐ ĐIÊN THOẠI",                                                 # G  - Điện thoại
    "2. ĐỊA CHỈ (THÔN/XÃ)",                                            # H  - Địa chỉ
    "KHOA KHÁM CHỮA BỆNH",                                              # I  - Khoa hẹn
    "3. GIỚI TÍNH",                                                     # J  - N/A
    "1. TRIỆU CHỨNG CHÍNH",                                             # K  - N/A
    "4. SỐ CĂN CƯỚC CÔNG DÂN - CHỨNG MINH THƯ",                        # L  - N/A
    "CHUYÊN KHOA MONG MUỐN KHÁM",                                       # M  - fixed value
    "BÁC SĨ MONG MUỐN ( nếu có)",                                       # N  - N/A
    "GIỜ KHÁM DỰ KIẾN",                                                 # O  - N/A
    "1. CAM KẾT CÁC THÔNG TIN LÀ THÔNG TIN ĐÚNG, CHỊU TRÁCH NHIỆM TRƯỚC PHÁP LUẬT TRƯỚC NHỮNG THÔNG TIN ĐÃ CUNG CẤP TRÊN",  # P - CÓ
    "ĐỒNG Ý CÁC ĐIỀU KHOẢN ĐẶT LỊCH KHÁM ONLINE TẠI BVĐK TÂM ĐỨC CẦU QUAN",  # Q - CÓ
]

def build_sheet_row(record, import_time_str):
    """
    Map one Minh Lo Excel record → one row matching SHEET_COLUMNS order.

    Thứ tự cột Google Sheet thực tế:
      A - Dấu thời gian            = thời gian import file
      B - NGUỒN BỆNH NHÂN          = "Bệnh nhân điều trị nội khoa tái khám"
      C - TRẠNG THÁI               = "" (trống)
      D - NGÀY KHÁM                = NGÀY HẸN từ Excel (dd/mm/yyyy)
      E - 1. HỌ VÀ TÊN BỆNH NHÂN  = HỌ TÊN từ Excel
      F - NĂM SINH                 = NĂM SINH (ước tính) từ tuổi trong Excel
      G - 5. SỐ ĐIÊN THOẠI         = SỐ ĐIỆN THOẠI từ Excel
      H - 2. ĐỊA CHỈ (THÔN/XÃ)    = ĐỊA CHỈ từ Excel
      I - KHOA KHÁM CHỮA BỆNH      = KHOA HẸN từ Excel
      J - 3. GIỚI TÍNH             = GIỚI TÍNH suy ra từ cột Tuổi Nam/Nữ trong Excel
      K - 1. TRIỆU CHỨNG CHÍNH     = N/A
      L - 4. SỐ CĂN CƯỚC...        = N/A
      M - CHUYÊN KHOA MONG MUỐN    = "Other: Bệnh nhân điều trị nội khoa tái khám"
      N - BÁC SĨ MONG MUỐN         = N/A
      O - GIỜ KHÁM DỰ KIẾN         = N/A
      P - CAM KẾT...               = "CÓ"
      Q - ĐỒNG Ý...                = "CÓ"

    Lưu ý:
      - Ngày khám: ghi dạng số serial Google Sheets (push_to_sheet xử lý convert),
        Google Sheet tự hiển thị đúng format Date, không có dấu apostrophe.
      - Số điện thoại: ghi dạng string (đã có số 0 đầu từ _fix_phone).
      - Trạng thái: mặc định "Bệnh nhân chưa khám/bỏ khám" (khớp dropdown).
    """
    # Format ngày: đảm bảo dd/mm/yyyy, nếu không có thì để N/A
    ngay_kham = record.get("NGÀY HẸN", "N/A") or "N/A"

    return [
        import_time_str,                                        # A - Dấu thời gian
        "Bệnh nhân điều trị nội khoa tái khám",                # B - NGUỒN BỆNH NHÂN
        "Bệnh nhân chưa khám/bỏ khám",                         # C - TRẠNG THÁI (mặc định dropdown)
        ngay_kham,                                              # D - NGÀY KHÁM (serial được convert trong push_to_sheet)
        record.get("HỌ TÊN", "N/A"),                           # E - HỌ VÀ TÊN BỆNH NHÂN
        record.get("NĂM SINH (ước tính)") or "N/A",            # F - NĂM SINH (đã sửa: trước đây bị hard-code "N/A")
        record.get("SỐ ĐIỆN THOẠI", "N/A"),                    # G - SỐ ĐIÊN THOẠI
        record.get("ĐỊA CHỈ", "N/A"),                          # H - ĐỊA CHỈ (THÔN/XÃ)
        record.get("KHOA HẸN", "N/A"),                         # I - KHOA KHÁM CHỮA BỆNH
        record.get("GIỚI TÍNH") or "N/A",                      # J - GIỚI TÍNH (đã sửa: trước đây bị hard-code "N/A")
        "N/A",                                                  # K - TRIỆU CHỨNG CHÍNH
        "N/A",                                                  # L - SỐ CĂN CƯỚC
        "Other: Bệnh nhân điều trị nội khoa tái khám",         # M - CHUYÊN KHOA MONG MUỐN
        "N/A",                                                  # N - BÁC SĨ MONG MUỐN
        "N/A",                                                  # O - GIỜ KHÁM DỰ KIẾN
        "CÓ",                                                   # P - CAM KẾT
        "CÓ",                                                   # Q - ĐỒNG Ý
    ]


def push_to_sheet(creds_data, sheet_id, sheet_name, records):
    """
    Append parsed Minh Lo records into the MAIN Google Sheet tab,
    mapping each field to the correct column position.
    Returns (rows_written, error_msg)
    """
    try:
        cl = authenticate_rw(creds_data)
        ss = cl.open_by_key(sheet_id)
        ws = ss.worksheet(sheet_name)

        if not records:
            return 0, "Không có dữ liệu để ghi."

        # Build timestamp once for whole import batch
        import_time_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        # Build rows in sheet column order
        rows_to_append = [build_sheet_row(r, import_time_str) for r in records]

        # ── Tìm dòng trống đầu tiên để append ──
        all_vals = ws.get_all_values()
        next_row = len(all_vals) + 1  # 1-indexed, dòng đầu tiên trống

        # ── Chuyển chuỗi ngày dd/mm/yyyy → số serial Google Sheets ──
        # Cột D (index 3) chứa NGÀY KHÁM. Cột NGÀY KHÁM có format Date trong
        # sheet nên phải ghi bằng số serial (không phải text) để tránh dấu '.
        # Google Sheets serial = số ngày kể từ 30/12/1899.
        DATE_COL_IDX = 3  # cột D, 0-indexed

        def date_str_to_serial(s):
            """Chuyển 'dd/mm/yyyy' → số serial Google Sheets (float).
            Trả về chuỗi gốc nếu không parse được."""
            try:
                d = datetime.strptime(s.strip(), "%d/%m/%Y")
                delta = d - datetime(1899, 12, 30)
                return delta.days  # số nguyên, Google Sheets tự hiểu là Date
            except Exception:
                return s  # giữ nguyên nếu không parse được

        # Chuyển đổi ngày trong từng row
        converted_rows = []
        for row in rows_to_append:
            row = list(row)
            row[DATE_COL_IDX] = date_str_to_serial(str(row[DATE_COL_IDX]))
            converted_rows.append(row)

        # ── Ghi bằng Sheets API với USER_ENTERED ──
        # USER_ENTERED: Google Sheet nhận số serial → tự hiển thị đúng định dạng Date
        # Dropdown (TRẠNG THÁI) hoạt động vì giá trị khớp với list validation
        ws.append_rows(
            converted_rows,
            value_input_option="USER_ENTERED",
            insert_data_option="INSERT_ROWS",
            table_range="A1",
        )

        return len(records), None

    except Exception as e:
        return 0, f"Lỗi ghi Sheet: {type(e).__name__}: {e}"


def update_patient_status(creds_data, sheet_id, sheet_name, sheet_row, new_status):
    """
    Cập nhật cột TRẠNG THÁI cho 1 bệnh nhân, xác định theo SỐ DÒNG THỰC TẾ
    trên Google Sheet (sheet_row, tính cả dòng tiêu đề — dòng dữ liệu đầu
    tiên là dòng 2). Chỉ ghi đúng 1 ô, không đụng tới các cột khác.
    Trả về (thành_công: bool, lỗi: str | None).
    """
    try:
        cl = authenticate_rw(creds_data)
        ss = cl.open_by_key(sheet_id)
        ws = ss.worksheet(sheet_name)

        headers = ws.row_values(1)
        if COL_STATUS not in headers:
            return False, f"Không tìm thấy cột '{COL_STATUS}' trên Google Sheet."
        col_idx = headers.index(COL_STATUS) + 1  # gspread dùng chỉ số 1-based

        ws.update_cell(sheet_row, col_idx, new_status)
        return True, None
    except Exception as e:
        return False, f"Lỗi cập nhật trạng thái: {type(e).__name__}: {e}"


def delete_patient_row(creds_data, sheet_id, sheet_name, sheet_row):
    """
    Xóa hẳn 1 dòng bệnh nhân khỏi Google Sheet, xác định theo SỐ DÒNG
    THỰC TẾ (sheet_row). Trả về (thành_công: bool, lỗi: str | None).
    """
    try:
        cl = authenticate_rw(creds_data)
        ss = cl.open_by_key(sheet_id)
        ws = ss.worksheet(sheet_name)
        ws.delete_rows(sheet_row)
        return True, None
    except Exception as e:
        return False, f"Lỗi xóa bệnh nhân: {type(e).__name__}: {e}"


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
    # Parse ngày khám trên TOÀN BỘ dữ liệu — TRƯỚC khi lọc theo trạng thái —
    # để tab "3 Ngày Tới" có thể nhắc lịch cho mọi bệnh nhân theo NGÀY KHÁM,
    # không bị bỏ sót những bệnh nhân chưa được gán cột TRẠNG THÁI.
    if COL_EXAM_DATE in df.columns:
        df["_date"] = pd.to_datetime(
            df[COL_EXAM_DATE].astype(str).str.strip(), format="%d/%m/%Y", errors="coerce"
        )
    else:
        df["_date"] = pd.NaT

    df_full = df.copy()  # bản ĐẦY ĐỦ, không lọc theo TRẠNG THÁI — dùng cho tab "3 Ngày Tới"

    df = df[~df[COL_STATUS].isin(["","nan","N/A","\u200b"])]
    total = len(df)
    if total == 0:
        df["_date"] = pd.NaT
        return _mk_empty(df, df_full)
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
                spec=spec, gen=gen, stbl=stbl, df=df, df_full=df_full,
                src_noi=src_noi, src_vl=src_vl, src_other=src_other)

def _mk_empty(df, df_full=None):
    df = df.copy(); df["_date"]=pd.NaT
    if df_full is None:
        df_full = df.copy()
    return dict(total=0,att=0,nos=0,att_pct=0.0,nos_pct=0.0,
                spec=None,gen=None,stbl=None,df=df,df_full=df_full,
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
    name   = str(row.get(COL_NAME,"")      or "—")
    dt     = str(row.get(COL_EXAM_DATE,"") or "—")
    spec   = str(row.get(COL_SPECIALTY,"") or "—")
    doc    = str(row.get(COL_DOCTOR,"")    or "")
    status = str(row.get(COL_STATUS,""))
    ts     = str(row.get(COL_TIMESTAMP,"") or "")
    src    = str(row.get(COL_SOURCE,"")    or "")

    if STATUS_ATTENDED.upper() in status.upper():
        st_cls = "pt-status-att"
    elif status.strip():
        st_cls = "pt-status-nos"
    else:
        st_cls = "pt-status-oth"

    doc_html = '<span class="pt-tag pt-tag-doc">' + doc[:25] + '</span>' if doc.strip() else ""
    src_html = source_badge(src)
    ts_html  = '<div style="font-size:0.62rem;color:#94a3b8;margin-top:0.12rem">&#128336; ' + ts + '</div>' if ts.strip() else ""

    return (
        '<div class="pt-card">'
        + '<div class="pt-name">' + name + '</div>'
        + '<div class="pt-row">'
        + '<span class="pt-tag pt-tag-date">&#128197; ' + dt + '</span>'
        + ' <span class="pt-tag pt-tag-spec">&#129658; ' + spec[:28] + '</span>'
        + (' ' + doc_html if doc_html else '')
        + '</div>'
        + '<div class="pt-row">'
        + '<span class="pt-tag ' + st_cls + '">' + (status or "—") + '</span>'
        + (' ' + src_html if src_html else '')
        + '</div>'
        + ts_html
        + '</div>'
    )


PAGE_SIZE = 10  # số bệnh nhân hiển thị mỗi trang, dùng chung cho mọi danh sách

def _smart_rerun():
    """
    Rerun thông minh cho các nút phân trang.

    Nếu đang chạy BÊN TRONG st.dialog (dialog hoạt động như 1 "fragment"),
    dùng st.rerun(scope="fragment") để CHỈ chạy lại nội dung của dialog,
    giữ nguyên cửa sổ dialog đang mở — tránh bị tắt khi bấm sang trang tiếp theo.
    Nếu đang ở ngoài dialog (trang chính), scope="fragment" sẽ báo lỗi vì
    không nằm trong fragment nào, khi đó rơi về st.rerun() bình thường
    (rerun toàn trang) như cũ.
    """
    try:
        st.rerun(scope="fragment")
    except Exception:
        st.rerun()

def _paginate_page_numbers(current, total):
    """
    Sinh danh sách số trang thông minh kiểu '1 … 4 5 [6] 7 8 … 42'.
    Luôn hiện trang đầu, trang cuối, trang hiện tại ± 1, và dùng None
    để đánh dấu chỗ cần chèn dấu '…'.
    """
    if total <= 7:
        return list(range(1, total + 1))
    pages = {1, total, current}
    for d in (-1, 0, 1):
        p = current + d
        if 1 <= p <= total:
            pages.add(p)
    pages = sorted(pages)
    out = []
    prev = None
    for p in pages:
        if prev is not None and p - prev > 1:
            out.append(None)
        out.append(p)
        prev = p
    return out

def render_paginated_cards(items_df, state_key, render_fn=None, page_size=PAGE_SIZE):
    """
    Hiển thị danh sách bệnh nhân dưới dạng thẻ (card), phân trang thông minh.

    - items_df   : DataFrame đã lọc, đúng thứ tự cần hiển thị.
    - state_key  : khoá session_state RIÊNG cho từng danh sách (vd. "pg_tab6"),
                    để mỗi danh sách nhớ trang hiện tại độc lập với nhau.
    - render_fn  : hàm nhận 1 row → trả về HTML thẻ; mặc định dùng patient_card_html.
    - page_size  : số dòng / trang (mặc định PAGE_SIZE = 10).

    Tự động:
      - Kẹp (clamp) số trang hiện tại nếu tổng số trang giảm (vd. sau khi đổi bộ lọc).
      - Reset về trang 1 nếu nội dung danh sách thay đổi (theo độ dài + index đầu/cuối).
      - Hiện thanh điều hướng « ‹ [số trang…] › » chỉ khi có nhiều hơn 1 trang.
    """
    render_fn = render_fn or patient_card_html
    total = len(items_df)

    if total == 0:
        return

    total_pages = max(1, (total + page_size - 1) // page_size)

    # Dấu vân tay đơn giản của tập dữ liệu hiện tại, để tự reset về trang 1
    # khi người dùng đổi bộ lọc / ngày / khoa (không cần callback riêng).
    fingerprint = (total, tuple(items_df.index[:1]), tuple(items_df.index[-1:]))
    fp_key = state_key + "_fp"

    if state_key not in st.session_state:
        st.session_state[state_key] = 1
    if st.session_state.get(fp_key) != fingerprint:
        st.session_state[fp_key] = fingerprint
        st.session_state[state_key] = 1

    # Kẹp trang hiện tại trong phạm vi hợp lệ
    cur = min(max(1, st.session_state[state_key]), total_pages)
    st.session_state[state_key] = cur

    start = (cur - 1) * page_size
    end = start + page_size
    page_df = items_df.iloc[start:end]

    cards_html = "".join(render_fn(row) for _, row in page_df.iterrows())
    st.markdown(cards_html, unsafe_allow_html=True)

    if total_pages > 1:
        st.markdown(
            f'<div class="pg-info">Trang <b>{cur}</b>/<b>{total_pages}</b> '
            f'&nbsp;·&nbsp; Hiển thị <b>{start+1}–{min(end,total)}</b> / <b>{total}</b> bệnh nhân</div>',
            unsafe_allow_html=True
        )

        nums = _paginate_page_numbers(cur, total_pages)
        # Bố cục gọn: [‹] [ ...số trang... ] [›]  (trang 1 và trang cuối luôn
        # có mặt trong dải số nên không cần thêm nút "về đầu / về cuối" riêng)
        with st.container(key=f"{state_key}_pgrow"):
            cols = st.columns([1] + [1] * len(nums) + [1])

            with cols[0]:
                with st.container(key=f"{state_key}_navbtn_prev"):
                    if st.button("‹", key=f"{state_key}_prev", disabled=(cur == 1),
                                 use_container_width=True):
                        st.session_state[state_key] = cur - 1
                        _smart_rerun()

            for i, p in enumerate(nums):
                with cols[1 + i]:
                    if p is None:
                        st.markdown(
                            '<div style="text-align:center;color:#94a3b8;font-size:0.68rem;height:1.7rem;line-height:1.7rem">…</div>',
                            unsafe_allow_html=True
                        )
                    else:
                        is_cur = (p == cur)
                        if is_cur:
                            # Trang hiện tại: bọc trong container riêng để CSS
                            # tô đậm chắc chắn, không phụ thuộc type="primary".
                            with st.container(key=f"{state_key}_curbtn_{p}"):
                                st.button(str(p), key=f"{state_key}_p{p}",
                                          disabled=True, use_container_width=True)
                        else:
                            if st.button(str(p), key=f"{state_key}_p{p}",
                                         use_container_width=True):
                                st.session_state[state_key] = p
                                _smart_rerun()

            with cols[1 + len(nums)]:
                with st.container(key=f"{state_key}_navbtn_next"):
                    if st.button("›", key=f"{state_key}_next", disabled=(cur == total_pages),
                                 use_container_width=True):
                        st.session_state[state_key] = cur + 1
                        _smart_rerun()


def classify_khoa_group(khoa_val):
    """
    Phân loại bệnh nhân theo cột KHOA KHÁM CHỮA BỆNH thành 2 nhóm:
      - "kb"   : Khoa Khám Bệnh + bệnh nhân chưa được gán khoa nào
      - "khac" : Các khoa điều trị nội trú khác (khác Khoa Khám Bệnh)
    """
    s = str(khoa_val).strip()
    if not s or s.lower() in ("nan", "n/a", "na", "none", "-", "—", "chưa xác định"):
        return "kb"
    if "khám bệnh" in s.lower():
        return "kb"
    return "khac"

def khoa_badge_html(khoa_val):
    s = str(khoa_val).strip()
    if not s or s.lower() in ("nan", "n/a", "na", "none", "-", "—"):
        return '<span class="khoa-badge khoa-none">Chưa phân khoa</span>'
    cls = "khoa-kb" if "khám bệnh" in s.lower() else "khoa-khac"
    label = s if len(s) <= 32 else s[:32] + "…"
    return f'<span class="khoa-badge {cls}">{label}</span>'

def patient_row_info_html(row2):
    """
    Dựng khối thông tin (tên + các thẻ chi tiết) cho MỘT bệnh nhân, dùng làm
    cột đầu tiên của mỗi "hàng" trong danh sách bệnh nhân (tab 3 Ngày Tới).
    Trạng thái + nút Sửa/Xóa được render riêng ở các cột kế tiếp bởi
    render_upcoming_table, không nằm trong khối này.
    """
    name  = str(row2.get(COL_NAME,"") or "—")
    byr   = str(row2.get(COL_BIRTH_YEAR,"") or "—")
    phone = str(row2.get(COL_PHONE,"") or "—")
    etime = str(row2.get(COL_EXAM_TIME,"") or "—")
    spec  = str(row2.get(COL_SPECIALTY,"") or "—")
    src_raw = str(row2.get(COL_SOURCE,"") or "")
    src_badge = source_badge(src_raw)
    khoa_show = khoa_badge_html(row2.get(COL_KHOA,""))

    if len(etime) >= 5 and ":" in etime:
        etime = etime[:5]
    if len(spec) > 30:
        spec = spec[:30] + "…"

    if phone not in ("—","N/A","nan",""):
        tel_digits = "".join(ch for ch in phone if ch.isdigit() or ch == "+")
        phone_show = (
            '<a href="tel:' + tel_digits + '" '
            'style="color:#1d4ed8;font-weight:700;text-decoration:none">📞 '
            + phone + '</a>'
        )
    else:
        phone_show = '<span style="color:#94a3b8">📞 ' + phone + '</span>'

    return (
        '<div class="mrow-info">'
        '<div class="mrow-name">👤 ' + name + '</div>'
        '<div class="pt-row" style="margin:0.3rem 0 0">'
        '<span class="pt-tag pt-tag-date">🕒 ' + etime + '</span>'
        '<span class="pt-tag pt-tag-spec">🎂 ' + byr + '</span>'
        '<span class="pt-tag pt-tag-spec">' + phone_show + '</span>'
        '</div>'
        '<div class="pt-row" style="margin:0.25rem 0 0">'
        '<span class="pt-tag pt-tag-doc">🩺 ' + spec + '</span>'
        + khoa_show
        + (src_badge or '')
        + '</div>'
        '</div>'
    )

def render_upcoming_table(sub_df, empty_msg, dl_prefix, dl_key, page_state_key=None,
                           day_date_iso=None, day_title=None):
    """
    Vẽ bảng chi tiết bệnh nhân + nút tải CSV cho 1 nhóm (kb / khác) trong 1 ngày.
    Phân trang thông minh 10 bệnh nhân/trang nếu page_state_key được truyền vào
    (khoá session_state riêng cho từng bảng, ví dụ 'pg_kb_2026-07-16'), để mỗi
    bảng nhớ trang hiện tại độc lập, không ảnh hưởng các bảng khác trên trang.

    day_date_iso / day_title: thông tin popup ngày đang mở (nếu bảng này được
    vẽ bên trong popup "Chi Tiết Lịch Khám Theo Ngày"), dùng để tự động MỞ LẠI
    popup danh sách này ngay sau khi người dùng đóng popup "Sửa"/"Xóa", thay vì
    phải bấm "Xem chi tiết danh sách" lại từ đầu.
    """
    if len(sub_df) == 0:
        st.markdown(
            '<div style="text-align:center;padding:1.4rem 0.5rem;color:#94a3b8;font-size:0.82rem">'
            + empty_msg + '</div>',
            unsafe_allow_html=True
        )
        return

    total = len(sub_df)
    state_key = page_state_key or (dl_key + "_pg")

    if total > PAGE_SIZE:
        total_pages = max(1, (total + PAGE_SIZE - 1) // PAGE_SIZE)
        fingerprint = (total, tuple(sub_df.index[:1]), tuple(sub_df.index[-1:]))
        fp_key = state_key + "_fp"
        if state_key not in st.session_state:
            st.session_state[state_key] = 1
        if st.session_state.get(fp_key) != fingerprint:
            st.session_state[fp_key] = fingerprint
            st.session_state[state_key] = 1
        cur = min(max(1, st.session_state[state_key]), total_pages)
        st.session_state[state_key] = cur
        start = (cur - 1) * PAGE_SIZE
        end = start + PAGE_SIZE
        page_df = sub_df.iloc[start:end]
    else:
        total_pages = 1
        cur = 1
        start, end = 0, total
        page_df = sub_df

    if total_pages > 1:
        st.markdown(
            f'<div class="pg-info">Trang <b>{cur}</b>/<b>{total_pages}</b> '
            f'&nbsp;·&nbsp; Hiển thị <b>{start+1}–{min(end,total)}</b> / <b>{total}</b> bệnh nhân</div>',
            unsafe_allow_html=True
        )
        nums = _paginate_page_numbers(cur, total_pages)
        with st.container(key=f"{state_key}_pgrow"):
            cols = st.columns([1] + [1] * len(nums) + [1])
            with cols[0]:
                with st.container(key=f"{state_key}_navbtn_prev"):
                    if st.button("‹", key=f"{state_key}_prev", disabled=(cur == 1), use_container_width=True):
                        st.session_state[state_key] = cur - 1
                        _smart_rerun()
            for i, p in enumerate(nums):
                with cols[1 + i]:
                    if p is None:
                        st.markdown(
                            '<div style="text-align:center;color:#94a3b8;font-size:0.68rem;height:1.7rem;line-height:1.7rem">…</div>',
                            unsafe_allow_html=True
                        )
                    else:
                        is_cur = (p == cur)
                        if is_cur:
                            with st.container(key=f"{state_key}_curbtn_{p}"):
                                st.button(str(p), key=f"{state_key}_p{p}",
                                          disabled=True, use_container_width=True)
                        else:
                            if st.button(str(p), key=f"{state_key}_p{p}", use_container_width=True):
                                st.session_state[state_key] = p
                                _smart_rerun()
            with cols[1 + len(nums)]:
                with st.container(key=f"{state_key}_navbtn_next"):
                    if st.button("›", key=f"{state_key}_next", disabled=(cur == total_pages), use_container_width=True):
                        st.session_state[state_key] = cur + 1
                        _smart_rerun()

    # ── DANH SÁCH BỆNH NHÂN: mỗi bệnh nhân 1 hàng, đầy đủ thông tin,
    # có nút "✏️ Sửa" / "🗑️ Xóa" ngay ở CUỐI mỗi hàng. ──
    HAS_DIALOG = hasattr(st, "dialog")
    for m_idx, m_row in page_df.iterrows():
        # m_idx là index gốc lấy từ Google Sheet (0-based, dòng dữ liệu đầu = index 0
        # = dòng 2 trên Sheet, vì dòng 1 là tiêu đề) → dòng thực tế = m_idx + 2.
        sheet_row = int(m_idx) + 2
        m_status = str(m_row.get(COL_STATUS, "") or "—")
        m_is_att = STATUS_ATTENDED.upper() in m_status.upper()
        m_badge_cls = "mrow-badge-att" if m_is_att else "mrow-badge-nos"
        info_html = patient_row_info_html(m_row)

        # return_day: nếu hàng này nằm trong popup "Chi Tiết Lịch Khám Theo Ngày",
        # ghi lại ngày đang xem để sau khi đóng popup Sửa/Xóa, hệ thống tự mở
        # lại ĐÚNG popup danh sách này — người dùng không cần bấm lại từ đầu.
        return_day = ({"date_iso": day_date_iso, "title": day_title}
                      if day_date_iso else None)

        with st.container(key=f"mrow_{dl_key}_{sheet_row}"):
            top_l, top_r = st.columns([3.3, 1.4])
            with top_l:
                st.markdown(info_html, unsafe_allow_html=True)
            with top_r:
                st.markdown(
                    f'<div class="mrow-status-wrap"><span class="mrow-badge {m_badge_cls}">{m_status}</span></div>',
                    unsafe_allow_html=True
                )

            st.markdown('<div class="mrow-actions">', unsafe_allow_html=True)
            act1, act2 = st.columns(2)
            with act1:
                with st.container(key=f"editbtn_{dl_key}_{sheet_row}"):
                    if st.button("✏️ Sửa trạng thái", key=f"btn_edit_{dl_key}_{sheet_row}", use_container_width=True):
                        if HAS_DIALOG:
                            # KHÔNG gọi dialog ngay tại đây vì hàng này đang nằm
                            # BÊN TRONG popup "Chi Tiết Lịch Khám Theo Ngày" — mở
                            # thêm 1 popup nữa lồng vào popup đang mở dễ gây lỗi/
                            # hiển thị sai ở Streamlit. Thay vào đó: đóng popup
                            # hiện tại (rerun) và ghi nhớ ý định qua session_state,
                            # popup "Sửa Trạng Thái" sẽ được mở lại ở tầng ngoài
                            # cùng (top-level) ngay sau khi rerun xong. return_day
                            # được đính kèm để popup danh sách tự mở lại sau đó.
                            st.session_state.pending_action = {
                                "type": "edit", "sheet_row": sheet_row,
                                "name": str(m_row.get(COL_NAME, "") or "—"), "status": m_status,
                                "return_day": return_day,
                            }
                            st.rerun()
                        else:
                            st.session_state[f"inline_edit_open_{sheet_row}"] = True
                            st.session_state[f"inline_del_open_{sheet_row}"] = False
                            st.rerun()
            with act2:
                with st.container(key=f"delbtn_{dl_key}_{sheet_row}"):
                    if st.button("🗑️ Xóa bệnh nhân", key=f"btn_del_{dl_key}_{sheet_row}", use_container_width=True):
                        if HAS_DIALOG:
                            st.session_state.pending_action = {
                                "type": "delete", "sheet_row": sheet_row,
                                "name": str(m_row.get(COL_NAME, "") or "—"),
                                "return_day": return_day,
                            }
                            st.rerun()
                        else:
                            st.session_state[f"inline_del_open_{sheet_row}"] = True
                            st.session_state[f"inline_edit_open_{sheet_row}"] = False
                            st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

        # Streamlit cũ không có st.dialog → hiện form Sửa/Xóa ngay dưới hàng
        if not HAS_DIALOG:
            if st.session_state.get(f"inline_edit_open_{sheet_row}"):
                render_inline_edit_form(sheet_row, str(m_row.get(COL_NAME, "") or "—"), m_status)
            if st.session_state.get(f"inline_del_open_{sheet_row}"):
                render_inline_delete_form(sheet_row, str(m_row.get(COL_NAME, "") or "—"))

    st.markdown('<div class="scroll-hint">💡 Vuốt màn hình để xem đầy đủ thông tin mỗi hàng</div>',
                unsafe_allow_html=True)

    dl_cols = [c for c in [COL_NAME, COL_BIRTH_YEAR, COL_PHONE, COL_EXAM_DATE, COL_EXAM_TIME,
                           COL_SPECIALTY, COL_KHOA, COL_DOCTOR, COL_SOURCE, COL_STATUS]
               if c in sub_df.columns]
    csv_data = sub_df[dl_cols].to_csv(index=False, encoding="utf-8-sig")
    st.download_button(
        label="⬇️ Tải danh sách (.csv)",
        data=csv_data.encode("utf-8-sig"),
        file_name=f"{dl_prefix}.csv",
        mime="text/csv",
        key=dl_key,
    )


def _do_save_status(sheet_row, new_status, close_keys=(), return_day=None):
    """Ghi trạng thái mới lên Google Sheet, làm mới cache, rồi rerun."""
    if not creds_data:
        st.error("❌ Chưa có thông tin xác thực (credentials). Kiểm tra Streamlit Secrets.")
        return
    with st.spinner("Đang cập nhật trạng thái…"):
        ok, err = update_patient_status(creds_data, SHEET_ID, SHEET_NAME, sheet_row, new_status)
    if ok:
        st.session_state.metrics = None  # buộc tải lại dữ liệu mới nhất từ Sheet
        st.session_state.pending_action = None
        if return_day:
            st.session_state.reopen_day = return_day  # tự mở lại popup danh sách ngày
        for k in close_keys:
            st.session_state[k] = False
        st.success("✅ Đã cập nhật trạng thái bệnh nhân!")
        st.rerun()
    else:
        st.error(f"❌ {err}")


def _do_delete_patient(sheet_row, close_keys=(), return_day=None):
    """Xóa bệnh nhân khỏi Google Sheet, làm mới cache, rồi rerun."""
    if not creds_data:
        st.error("❌ Chưa có thông tin xác thực (credentials). Kiểm tra Streamlit Secrets.")
        return
    with st.spinner("Đang xóa bệnh nhân…"):
        ok, err = delete_patient_row(creds_data, SHEET_ID, SHEET_NAME, sheet_row)
    if ok:
        st.session_state.metrics = None
        st.session_state.pending_action = None
        if return_day:
            st.session_state.reopen_day = return_day  # tự mở lại popup danh sách ngày
        for k in close_keys:
            st.session_state[k] = False
        st.success("✅ Đã xóa bệnh nhân khỏi danh sách!")
        st.rerun()
    else:
        st.error(f"❌ {err}")


def open_edit_status_dialog(sheet_row, patient_name, current_status, return_day=None):
    """Mở popup 'Sửa Trạng Thái Bệnh Nhân' (nằm trên popup ngày, dạng popup lồng).

    return_day: nếu được truyền, ngay khi người dùng bấm Hủy/Lưu, popup danh
    sách bệnh nhân của ngày đó sẽ tự động được mở lại — không cần bấm lại.
    """
    @st.dialog("✏️ Sửa Trạng Thái Bệnh Nhân", width="small")
    def _dlg():
        st.markdown(
            f'<div class="edit-dlg-name">👤 {patient_name}</div>',
            unsafe_allow_html=True
        )
        options = [STATUS_NOT_ATTENDED, STATUS_ATTENDED]
        default_idx = 1 if STATUS_ATTENDED.upper() in (current_status or "").upper() else 0
        new_status = st.radio(
            "Trạng thái khám",
            options=options,
            index=default_idx,
            key=f"edit_status_radio_{sheet_row}",
            label_visibility="collapsed",
        )
        st.markdown("<div style='height:0.3rem'></div>", unsafe_allow_html=True)
        bc1, bc2 = st.columns(2)
        with bc1:
            if st.button("❌ Hủy", key=f"edit_cancel_{sheet_row}", use_container_width=True):
                st.session_state.pending_action = None
                if return_day:
                    st.session_state.reopen_day = return_day
                st.rerun()
        with bc2:
            if st.button("💾 Lưu Thay Đổi", key=f"edit_save_{sheet_row}",
                         use_container_width=True, type="primary"):
                _do_save_status(sheet_row, new_status, return_day=return_day)
    _dlg()


def open_delete_dialog(sheet_row, patient_name, return_day=None):
    """Mở popup xác nhận xóa bệnh nhân khỏi Google Sheet.

    return_day: nếu được truyền, ngay khi người dùng bấm Hủy/Xác Nhận Xóa,
    popup danh sách bệnh nhân của ngày đó sẽ tự động được mở lại.
    """
    @st.dialog("🗑️ Xóa Bệnh Nhân", width="small")
    def _dlg():
        st.markdown(
            '<div class="del-dlg-warn">⚠️ Bạn có chắc chắn muốn xóa bệnh nhân<br>'
            f'<b>{patient_name}</b> khỏi danh sách trên Google Sheet?'
            '<span class="del-dlg-note">Hành động này không thể hoàn tác.</span></div>',
            unsafe_allow_html=True
        )
        bc1, bc2 = st.columns(2)
        with bc1:
            if st.button("❌ Hủy", key=f"del_cancel_{sheet_row}", use_container_width=True):
                st.session_state.pending_action = None
                if return_day:
                    st.session_state.reopen_day = return_day
                st.rerun()
        with bc2:
            if st.button("🗑️ Xác Nhận Xóa", key=f"del_confirm_{sheet_row}",
                         use_container_width=True, type="primary"):
                _do_delete_patient(sheet_row, return_day=return_day)
    _dlg()


def render_inline_edit_form(sheet_row, patient_name, current_status):
    """
    Fallback cho các phiên bản Streamlit CŨ không có st.dialog:
    hiện form Sửa trạng thái ngay bên dưới hàng bệnh nhân thay vì popup.
    """
    with st.container(border=True):
        st.markdown(f"**✏️ Sửa trạng thái — {patient_name}**")
        options = [STATUS_NOT_ATTENDED, STATUS_ATTENDED]
        default_idx = 1 if STATUS_ATTENDED.upper() in (current_status or "").upper() else 0
        new_status = st.radio(
            "Trạng thái khám", options=options, index=default_idx,
            key=f"inline_edit_radio_{sheet_row}", label_visibility="collapsed",
        )
        ic1, ic2 = st.columns(2)
        with ic1:
            if st.button("💾 Lưu", key=f"inline_edit_save_{sheet_row}", use_container_width=True):
                _do_save_status(sheet_row, new_status, close_keys=[f"inline_edit_open_{sheet_row}"])
        with ic2:
            if st.button("❌ Đóng", key=f"inline_edit_close_{sheet_row}", use_container_width=True):
                st.session_state[f"inline_edit_open_{sheet_row}"] = False
                st.rerun()


def render_inline_delete_form(sheet_row, patient_name):
    """Fallback cho Streamlit cũ: form xác nhận xóa hiện ngay dưới hàng bệnh nhân."""
    with st.container(border=True):
        st.warning(f"⚠️ Xác nhận xóa bệnh nhân **{patient_name}**? Hành động không thể hoàn tác.")
        ic1, ic2 = st.columns(2)
        with ic1:
            if st.button("🗑️ Xác nhận xóa", key=f"inline_del_confirm_{sheet_row}", use_container_width=True):
                _do_delete_patient(sheet_row, close_keys=[f"inline_del_open_{sheet_row}"])
        with ic2:
            if st.button("❌ Hủy", key=f"inline_del_close_{sheet_row}", use_container_width=True):
                st.session_state[f"inline_del_open_{sheet_row}"] = False
                st.rerun()


def split_khoa_groups(day_df):
    """Trả về (kb_df, khac_df) đã phân loại theo cột KHOA KHÁM CHỮA BỆNH."""
    if COL_KHOA in day_df.columns:
        grp = day_df[COL_KHOA].apply(classify_khoa_group)
    else:
        grp = pd.Series(["kb"] * len(day_df), index=day_df.index)
    return day_df[grp == "kb"], day_df[grp == "khac"]


# ═══════════════════════════════════════════════
# SESSION + FETCH
# ═══════════════════════════════════════════════
creds_data = get_credentials()
for k,v in [("metrics",None),("fetch_time",None),("err",None),("pending_action",None),("reopen_day",None)]:
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

# ── Mở popup "Sửa Trạng Thái" / "Xóa Bệnh Nhân" (nếu có) ──
# Luôn kiểm tra ở TẦNG NGOÀI CÙNG của script (không nằm trong bất kỳ
# popup/dialog nào khác) để tránh lỗi mở dialog lồng trong dialog của
# Streamlit. Cờ pending_action được đặt khi bấm nút Sửa/Xóa trong popup
# "Chi Tiết Lịch Khám Theo Ngày" ở tab "3 Ngày Tới".
if hasattr(st, "dialog") and st.session_state.get("pending_action"):
    _pa = st.session_state.pending_action
    if _pa.get("type") == "edit":
        open_edit_status_dialog(_pa["sheet_row"], _pa["name"], _pa.get("status", ""),
                                 return_day=_pa.get("return_day"))
    elif _pa.get("type") == "delete":
        open_delete_dialog(_pa["sheet_row"], _pa["name"], return_day=_pa.get("return_day"))

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
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "📊 Tổng Quan",
        "🔍 Tìm Theo Ngày",
        "📅 3 Ngày Tới",
        "🏥 Nguồn Bệnh Nhân",
        "📈 Báo Cáo",
        "👤 Bệnh Nhân",
        "📥 Import Từ Minh Lộ",
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
            # Use cards always (works on both, better on mobile), có phân trang
            # 10 bệnh nhân/trang. Đổi ngày/bộ lọc sẽ tự động quay về trang 1.
            render_paginated_cards(sd_df[show_cols], "pg_tab2_search")

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
    # TAB 3 — 3 NGÀY TỚI
    # ════════════════
    with tab3:
        st.markdown(
            '<div class="sh"><div class="sh-dot" style="background:#f59e0b"></div>'
            '<span class="sh-txt">Danh Sách Bệnh Nhân Chuẩn Bị Khám Trong 3 Ngày Tới</span></div>',
            unsafe_allow_html=True
        )

        # Build list of next 3 days (exclude today)
        upcoming_dates = [today + timedelta(days=i) for i in range(1, 4)]

        # Dùng dữ liệu ĐẦY ĐỦ (không lọc theo cột TRẠNG THÁI) để không bỏ
        # sót bệnh nhân chưa được gán trạng thái — phục vụ mục đích nhắc lịch khám.
        df_upcoming = m.get("df_full", df)

        HAS_DIALOG = hasattr(st, "dialog")

        def _get_day_df(udate):
            dday = df_upcoming[df_upcoming["_date"].dt.date == udate].copy()
            if COL_EXAM_TIME in dday.columns:
                dday = dday.sort_values(COL_EXAM_TIME)
            return dday

        def _render_day_detail(dday, udate_iso, day_title_=None):
            kb_df, khac_df = split_khoa_groups(dday)
            gtab1, gtab2 = st.tabs([
                f"🩺 Khoa Khám Bệnh & Chưa Phân Khoa · {len(kb_df)}",
                f"🏥 Khoa Điều Trị Nội Trú Khác · {len(khac_df)}",
            ])
            with gtab1:
                st.markdown(
                    '<div class="dlg-group-hd"><b>🩺 Khoa Khám Bệnh &amp; bệnh nhân chưa phân khoa</b>'
                    f'<span>{len(kb_df)} bệnh nhân</span></div>',
                    unsafe_allow_html=True
                )
                render_upcoming_table(
                    kb_df,
                    "Không có bệnh nhân thuộc nhóm Khoa Khám Bệnh / chưa phân khoa trong ngày này.",
                    f"kb_{udate_iso}", f"dl_kb_{udate_iso}",
                    page_state_key=f"pg_kb_{udate_iso}",
                    day_date_iso=udate_iso, day_title=day_title_,
                )
            with gtab2:
                st.markdown(
                    '<div class="dlg-group-hd"><b>🏥 Các khoa điều trị nội trú khác</b>'
                    f'<span>{len(khac_df)} bệnh nhân</span></div>',
                    unsafe_allow_html=True
                )
                render_upcoming_table(
                    khac_df,
                    "Không có bệnh nhân thuộc các khoa điều trị nội trú khác trong ngày này.",
                    f"khac_{udate_iso}", f"dl_khac_{udate_iso}",
                    page_state_key=f"pg_khac_{udate_iso}",
                    day_date_iso=udate_iso, day_title=day_title_,
                )

        if HAS_DIALOG:
            @st.dialog("📋 Chi Tiết Lịch Khám Theo Ngày", width="large")
            def _open_day_dialog(date_iso, day_title_):
                st.markdown(f"#### 📅 {day_title_}")
                dday = _get_day_df(datetime.strptime(date_iso, "%Y-%m-%d").date())
                st.caption(f"Tổng cộng {len(dday)} bệnh nhân đăng ký khám ngày này.")
                _render_day_detail(dday, date_iso, day_title_)

            # ── Tự động MỞ LẠI popup danh sách ngày sau khi popup "Sửa"/"Xóa"
            # được đóng (dù bấm Lưu/Xác Nhận Xóa hay Hủy), để người dùng không
            # phải bấm "Xem chi tiết danh sách" lại từ đầu mỗi lần sửa/xóa.
            _rd = st.session_state.get("reopen_day")
            if _rd and not st.session_state.get("pending_action"):
                st.session_state.reopen_day = None
                _open_day_dialog(_rd["date_iso"], _rd["title"])

        if "_date" in df_upcoming.columns and df_upcoming["_date"].notna().any():
            # Tổng số bệnh nhân trong cả 3 ngày tới — hiển thị rõ để dễ kiểm chứng
            # không bị thiếu, dù có bao nhiêu bệnh nhân đăng ký (100, 200,...).
            total_upcoming_all = sum(
                int((df_upcoming["_date"].dt.date == udate).sum()) for udate in upcoming_dates
            )
            st.markdown(
                '<div style="background:#eff6ff;border:1px solid #bfdbfe;border-radius:10px;'
                'padding:0.7rem 1rem;margin-bottom:0.8rem;text-align:center;color:#1d4ed8">'
                '&#128203; Tổng cộng <span style="font-size:1.2rem;font-weight:800">'
                + str(total_upcoming_all) +
                '</span> bệnh nhân đăng ký khám trong 3 ngày tới</div>',
                unsafe_allow_html=True
            )

            day_labels = {0:"Thứ Hai",1:"Thứ Ba",2:"Thứ Tư",
                           3:"Thứ Năm",4:"Thứ Sáu",5:"Thứ Bảy",6:"Chủ Nhật"}
            day_colors = ["#3b82f6","#10b981","#8b5cf6"]

            for dci, udate in enumerate(upcoming_dates):
                day_df = _get_day_df(udate)
                count  = len(day_df)
                udate_iso = udate.isoformat()

                weekday_vn = day_labels.get(udate.weekday(), "")
                day_title  = weekday_vn + " — " + udate.strftime("%d/%m/%Y")
                dc = day_colors[dci]

                kb_df, khac_df = split_khoa_groups(day_df)
                kb_count, khac_count = len(kb_df), len(khac_df)

                st.markdown(
                    '<div class="upcoming-day">'
                    '<div class="upcoming-day-header" style="background:linear-gradient(135deg,' + dc + ',#0f172a);">'
                    '<span class="upcoming-day-title">&#128197; ' + day_title + '</span>'
                    '<span class="upcoming-day-count">' + str(count) + ' bệnh nhân</span>'
                    '</div>',
                    unsafe_allow_html=True
                )

                if count > 0:
                    st.markdown(
                        '<div class="upcoming-day-stats">'
                        '<div class="uds-item uds-kb">'
                        '<span class="uds-val">' + str(kb_count) + '</span>'
                        '<span class="uds-lbl">🩺 Khoa Khám Bệnh &amp; chưa phân khoa</span>'
                        '</div>'
                        '<div class="uds-item uds-khac">'
                        '<span class="uds-val">' + str(khac_count) + '</span>'
                        '<span class="uds-lbl">🏥 Khoa điều trị nội trú khác</span>'
                        '</div>'
                        '</div>',
                        unsafe_allow_html=True
                    )

                    if HAS_DIALOG:
                        st.markdown('<div class="upcoming-day-actions">', unsafe_allow_html=True)
                        if st.button("👁️ Xem chi tiết danh sách", key="btn_open_" + udate.strftime("%Y%m%d"),
                                     use_container_width=True):
                            _open_day_dialog(udate.isoformat(), day_title)
                        st.markdown('</div>', unsafe_allow_html=True)
                    else:
                        st.markdown('<div class="upcoming-day-actions">', unsafe_allow_html=True)
                        with st.expander("👁️ Xem chi tiết danh sách bệnh nhân", expanded=False):
                            _render_day_detail(day_df, udate_iso)
                        st.markdown('</div>', unsafe_allow_html=True)
                else:
                    st.markdown(
                        '<div class="upcoming-day-body">'
                        '<p style="color:#94a3b8;font-size:0.82rem;text-align:center;padding:0.8rem 0">'
                        'Chưa có bệnh nhân đăng ký ngày này.</p>'
                        '</div>',
                        unsafe_allow_html=True
                    )

                st.markdown('</div>', unsafe_allow_html=True)  # close upcoming-day

            if total_upcoming_all == 0:
                st.markdown(
                    '<div class="empty">'
                    '<div class="empty-ico">&#128197;</div>'
                    '<div class="empty-ttl">Chưa có lịch hẹn trong 3 ngày tới</div>'
                    '<div class="empty-dsc">Bệnh nhân chưa đăng ký lịch khám cho các ngày sắp tới.</div>'
                    '</div>',
                    unsafe_allow_html=True
                )
        else:
            st.markdown(
                '<div class="empty">'
                '<div class="empty-ico">&#128197;</div>'
                '<div class="empty-ttl">Không có dữ liệu ngày khám</div>'
                '<div class="empty-dsc">Cột NGÀY KHÁM chưa có dữ liệu hoặc sai định dạng dd/mm/yyyy.</div>'
                '</div>',
                unsafe_allow_html=True
            )

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

            # ── Thống kê Tái Khám theo Ngày / Tháng / Năm ──
            if COL_SOURCE in df.columns and src_noi > 0:
                st.markdown(
                    f'<div class="sh"><div class="sh-dot" style="background:{CV}"></div>'
                    f'<span class="sh-txt">Thống Kê Tái Khám Theo Thời Gian</span></div>',
                    unsafe_allow_html=True
                )

                noi_all = df[df[COL_SOURCE].astype(str).str.contains(
                    "khoa|tái|nội trú|xuất viện|tai", case=False, na=False
                )].copy()
                noi_dated = noi_all[noi_all["_date"].notna()].copy() if "_date" in noi_all.columns else noi_all.iloc[0:0].copy()

                tk_period = st.radio(
                    "Xem theo:", ["Ngày", "Tháng", "Năm"],
                    horizontal=True, index=0, label_visibility="collapsed",
                    key="tk_period_filter"
                )

                # Danh sách các năm có dữ liệu (luôn có năm hiện tại để chọn được)
                years_avail = sorted(set(
                    noi_dated["_date"].dt.year.dropna().astype(int).tolist() + [today.year]
                ))

                if tk_period == "Ngày":
                    tk_date = st.date_input(
                        "Chọn ngày", value=today, format="DD/MM/YYYY", key="tk_date_pick"
                    )
                    tk_filtered = noi_dated[noi_dated["_date"].dt.date == tk_date]
                    tk_label = f"Ngày {tk_date.strftime('%d/%m/%Y')}"
                elif tk_period == "Tháng":
                    tcol1, tcol2 = st.columns(2)
                    with tcol1:
                        tk_month = st.selectbox(
                            "Tháng", list(range(1, 13)), index=today.month - 1,
                            format_func=lambda m: f"Tháng {m:02d}", key="tk_month_pick"
                        )
                    with tcol2:
                        tk_year = st.selectbox(
                            "Năm", years_avail,
                            index=years_avail.index(today.year), key="tk_year_pick_m"
                        )
                    tk_filtered = noi_dated[
                        (noi_dated["_date"].dt.month == tk_month) &
                        (noi_dated["_date"].dt.year == tk_year)
                    ]
                    tk_label = f"Tháng {tk_month:02d}/{tk_year}"
                else:  # Năm
                    tk_year = st.selectbox(
                        "Năm", years_avail,
                        index=years_avail.index(today.year), key="tk_year_pick_y"
                    )
                    tk_filtered = noi_dated[noi_dated["_date"].dt.year == tk_year]
                    tk_label = f"Năm {tk_year}"

                tk_total = len(tk_filtered)
                tk_att = int(
                    (tk_filtered[COL_STATUS].astype(str).str.upper() == STATUS_ATTENDED.upper()).sum()
                ) if COL_STATUS in tk_filtered.columns else 0
                tk_nos = tk_total - tk_att
                tk_att_pct = round(tk_att / tk_total * 100, 1) if tk_total > 0 else 0
                tk_nos_pct = round(100 - tk_att_pct, 1) if tk_total > 0 else 0

                st.markdown(f"""
                <div class="kg" style="grid-template-columns:repeat(3,1fr);gap:0.55rem;margin:0.7rem 0 0.9rem">
                  <div class="kc kc-v"><div class="kc-bg">🏥</div>
                    <div class="kc-lbl">Tổng Tái Khám — {tk_label}</div>
                    <div class="kc-val">{tk_total}</div>
                    <div class="kc-sub">bệnh nhân từ khoa / tái khám</div>
                  </div>
                  <div class="kc kc-g"><div class="kc-bg">✅</div>
                    <div class="kc-lbl">Đã Đến Khám</div>
                    <div class="kc-val">{tk_att}</div>
                    <div class="kc-sub">{tk_att_pct}% trong kỳ</div>
                  </div>
                  <div class="kc kc-r"><div class="kc-bg">❌</div>
                    <div class="kc-lbl">Vắng / Chưa Khám</div>
                    <div class="kc-val">{tk_nos}</div>
                    <div class="kc-sub">{tk_nos_pct}% trong kỳ</div>
                  </div>
                </div>
                """, unsafe_allow_html=True)

                if tk_total > 0:
                    st.markdown(
                        f'<div class="sh"><div class="sh-dot" style="background:{CV}"></div>'
                        f'<span class="sh-txt">Danh Sách Bệnh Nhân Tái Khám — {tk_label} '
                        f'({tk_total} người)</span></div>',
                        unsafe_allow_html=True
                    )
                    show_src_cols = [col for col in [COL_TIMESTAMP, COL_NAME, COL_EXAM_DATE,
                                                      COL_STATUS, COL_SPECIALTY, COL_SOURCE]
                                     if col in tk_filtered.columns]
                    tk_filtered_show = tk_filtered[show_src_cols].reset_index(drop=True)
                    # Khoá phân trang riêng theo kỳ đang xem, để đổi ngày/tháng/năm tự về trang 1
                    tk_pg_key = "pg_tab4_taikham_" + re.sub(r"\W+", "_", tk_period + "_" + tk_label)
                    render_paginated_cards(tk_filtered_show, tk_pg_key)

                    csv_src = tk_filtered[show_src_cols].to_csv(index=False, encoding="utf-8-sig")
                    st.download_button(
                        label=f"⬇️ Tải danh sách tái khám {tk_label} (.csv)",
                        data=csv_src.encode("utf-8-sig"),
                        file_name=f"taikham_{re.sub(r'[^0-9A-Za-z]+', '_', tk_label)}.csv",
                        mime="text/csv",
                        key="dl_taikham_period_csv",
                    )
                else:
                    st.markdown(f"""
                    <div class="empty">
                      <div class="empty-ico">📭</div>
                      <div class="empty-ttl">Không có bệnh nhân tái khám trong {tk_label}</div>
                      <div class="empty-dsc">Thử chọn ngày / tháng / năm khác.</div>
                    </div>
                    """, unsafe_allow_html=True)

            # ── Thống kê theo KHOA KHÁM CHỮA BỆNH ──
            if COL_KHOA in df.columns:
                st.markdown(
                    f'<div class="sh"><div class="sh-dot" style="background:{CT}"></div>'
                    f'<span class="sh-txt">Thống Kê Tái Khám Theo Khoa</span></div>',
                    unsafe_allow_html=True
                )

                # Chỉ lấy bệnh nhân tái khám (có nguồn từ khoa)
                noi_df_khoa = df[
                    df[COL_SOURCE].astype(str).str.contains(
                        "khoa|tái|nội trú|xuất viện|tai|nội khoa", case=False, na=False
                    )
                ].copy() if COL_SOURCE in df.columns else df.copy()

                khoa_raw = noi_df_khoa[COL_KHOA].astype(str).str.strip()
                khoa_raw = khoa_raw[khoa_raw.str.len() > 0]
                khoa_raw = khoa_raw[~khoa_raw.str.lower().isin(["", "nan", "n/a", "na"])]

                if not khoa_raw.empty:
                    khoa_counts = khoa_raw.value_counts().reset_index()
                    khoa_counts.columns = ["Khoa", "Số lượt"]
                    khoa_counts["Tỷ lệ"] = (
                        khoa_counts["Số lượt"] / khoa_counts["Số lượt"].sum() * 100
                    ).round(1)

                    # Bộ lọc khoa
                    all_khoa = ["Tất cả khoa"] + khoa_counts["Khoa"].tolist()
                    sel_khoa = st.selectbox(
                        "🔍 Lọc theo khoa:",
                        all_khoa,
                        label_visibility="visible",
                        key="sel_khoa_filter"
                    )

                    # ── Biểu đồ bar ngang ──
                    st.markdown('<div class="cc">', unsafe_allow_html=True)
                    fig_khoa = go.Figure(go.Bar(
                        x=khoa_counts["Số lượt"],
                        y=khoa_counts["Khoa"],
                        orientation="h",
                        marker=dict(
                            color=[CV if sel_khoa != "Tất cả khoa" and k == sel_khoa else CT
                                   for k in khoa_counts["Khoa"]],
                            opacity=0.85,
                        ),
                        text=[f"{v} ({p}%)" for v, p in
                              zip(khoa_counts["Số lượt"], khoa_counts["Tỷ lệ"])],
                        textposition="outside",
                        hovertemplate="<b>%{y}</b><br>Số lượt: %{x}<extra></extra>",
                    ))
                    fig_khoa.update_layout(
                        height=max(220, len(khoa_counts) * 48),
                        margin=dict(t=10, b=10, l=8, r=90),
                        xaxis=dict(
                            showgrid=True, gridcolor="#f1f5f9",
                            zeroline=False, tickfont=dict(size=9, color="#64748b"),
                        ),
                        yaxis=dict(
                            autorange="reversed",
                            tickfont=dict(size=10, color="#1e293b"),
                            showgrid=False,
                        ),
                        plot_bgcolor="white", paper_bgcolor="white",
                        font=dict(family="Inter, sans-serif"),
                    )
                    st.plotly_chart(fig_khoa, use_container_width=True,
                                    config={"displayModeBar": False})
                    st.markdown('</div>', unsafe_allow_html=True)

                    # ── Bảng chi tiết + lọc ──
                    st.markdown(
                        f'<div class="sh"><div class="sh-dot" style="background:{CT}"></div>'
                        f'<span class="sh-txt">Chi Tiết Bệnh Nhân'
                        f'{" — " + sel_khoa if sel_khoa != "Tất cả khoa" else " — Tất Cả Khoa"}'
                        f'</span></div>',
                        unsafe_allow_html=True
                    )

                    # Lọc df theo khoa
                    if sel_khoa == "Tất cả khoa":
                        filtered_khoa_df = noi_df_khoa[
                            noi_df_khoa[COL_KHOA].astype(str).str.strip()
                            .isin(khoa_counts["Khoa"].tolist())
                        ]
                    else:
                        filtered_khoa_df = noi_df_khoa[
                            noi_df_khoa[COL_KHOA].astype(str).str.strip() == sel_khoa
                        ]

                    # KPI mini cho khoa được chọn
                    fk_total = len(filtered_khoa_df)
                    fk_att = int(
                        (filtered_khoa_df[COL_STATUS].str.upper() == STATUS_ATTENDED.upper()).sum()
                    ) if COL_STATUS in filtered_khoa_df.columns else 0
                    fk_nos = fk_total - fk_att
                    fk_pct = round(fk_att / fk_total * 100, 1) if fk_total > 0 else 0

                    st.markdown(f"""
                    <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:0.5rem;margin:0.6rem 0 0.9rem">
                      <div class="kc kc-t" style="padding:0.7rem 0.9rem">
                        <div class="kc-lbl">Tổng Lượt</div>
                        <div class="kc-val" style="font-size:1.5rem">{fk_total}</div>
                      </div>
                      <div class="kc kc-g" style="padding:0.7rem 0.9rem">
                        <div class="kc-lbl">Đã Đến Khám</div>
                        <div class="kc-val" style="font-size:1.5rem;color:#059669">{fk_att}</div>
                      </div>
                      <div class="kc kc-r" style="padding:0.7rem 0.9rem">
                        <div class="kc-lbl">Vắng / Chưa</div>
                        <div class="kc-val" style="font-size:1.5rem;color:#dc2626">{fk_nos}</div>
                      </div>
                    </div>
                    """, unsafe_allow_html=True)

                    # Danh sách bệnh nhân theo khoa
                    show_khoa_cols = [col for col in [
                        COL_NAME, COL_EXAM_DATE, COL_STATUS,
                        COL_KHOA, COL_PHONE, COL_SOURCE
                    ] if col in filtered_khoa_df.columns]

                    filtered_khoa_df = filtered_khoa_df.reset_index(drop=True)
                    # Khoá phân trang riêng theo khoa đang chọn để đổi khoa tự về trang 1
                    khoa_pg_key = "pg_tab4_khoa_" + re.sub(r"\W+", "_", sel_khoa)
                    render_paginated_cards(filtered_khoa_df[show_khoa_cols], khoa_pg_key)

                    khoa_fname = sel_khoa.replace(" ", "_") if sel_khoa != "Tất cả khoa" else "tat_ca_khoa"
                    csv_khoa = filtered_khoa_df[show_khoa_cols].to_csv(index=False, encoding="utf-8-sig")
                    st.download_button(
                        label=f"⬇️ Tải danh sách {sel_khoa} (.csv)",
                        data=csv_khoa.encode("utf-8-sig"),
                        file_name=f"taikham_{khoa_fname}_{datetime.now().strftime('%Y%m%d')}.csv",
                        mime="text/csv",
                        key="dl_khoa_csv",
                    )
                else:
                    st.info("Chưa có dữ liệu cột KHOA KHÁM CHỮA BỆNH cho bệnh nhân tái khám.")
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
            show_df = fdf[show_cols].reset_index(drop=True)
            render_paginated_cards(show_df, "pg_tab6_patients")
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

    # ════════════════
    # TAB 7 — IMPORT TỪ MINH LỘ
    # ════════════════
    with tab7:
        st.markdown(
            '<div class="sh"><div class="sh-dot" style="background:#10b981"></div>'
            '<span class="sh-txt">Import Danh Sách Hẹn Khám Từ Phần Mềm Minh Lộ</span></div>',
            unsafe_allow_html=True
        )

        # Info box
        st.markdown("""
        <div style="background:#eff6ff;border:1px solid #bfdbfe;border-radius:12px;
                    padding:1rem 1.2rem;margin-bottom:1rem;font-size:0.83rem;color:#1e40af">
          <b>📋 Hướng dẫn:</b><br>
          1. Vào Minh Lộ → Báo cáo → <b>Danh sách bệnh nhân hẹn khám lại</b><br>
          2. Chọn ngày cần xuất → Export ra <b>Excel (.xlsx)</b><br>
          3. Upload file vào đây → Kiểm tra preview → Nhấn <b>Import vào Google Sheet</b><br>
          4. Dữ liệu sẽ được <b>thêm vào sheet chính</b> với đầy đủ các trường đã mapping
        </div>
        <div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:12px;
                    padding:0.8rem 1.2rem;margin-bottom:1rem;font-size:0.8rem;color:#166534">
          <b>✅ Quy tắc mapping:</b>
          Họ tên → HỌ VÀ TÊN &nbsp;|&nbsp; Địa chỉ → ĐỊA CHỈ &nbsp;|&nbsp;
          Ngày hẹn → NGÀY KHÁM &nbsp;|&nbsp; Điện thoại → SỐ ĐIÊN THOẠI &nbsp;|&nbsp;
          Khoa hẹn → KHOA KHÁM CHỮA BỆNH &nbsp;|&nbsp;
          Nguồn BN → "Bệnh nhân điều trị nội khoa tái khám" &nbsp;|&nbsp;
          Thời gian đăng ký → Thời điểm import &nbsp;|&nbsp;
          Cam kết & Đồng ý → "CÓ"
        </div>
        """, unsafe_allow_html=True)

        # Always write to main sheet tab
        target_tab = SHEET_NAME

        # File uploader
        xl_file = st.file_uploader(
            "Upload file Excel từ Minh Lộ (.xlsx)",
            type=["xlsx"],
            help="File xuất từ màn hình Danh sách bệnh nhân hẹn khám lại"
        )

        if xl_file is not None:
            with st.spinner("Đang đọc file Excel…"):
                records, err_parse = parse_minh_lo_excel(xl_file)

            if err_parse:
                st.error(f"❌ {err_parse}")
            elif not records:
                st.warning("⚠️ Không tìm thấy dữ liệu bệnh nhân trong file. Kiểm tra lại định dạng file.")
            else:
                st.success(f"✅ Đọc được **{len(records)}** bệnh nhân từ file Excel")

                # ═══════════════════════════════════════════════
                # CHỌN "NGÀY LẬP" ĐỂ LỌC BỆNH NHÂN CẦN GHI VÀO SHEET
                # (Minh Lộ liệt kê theo NGÀY VÀO VIỆN nên 1 file thường trải
                #  dài cả tháng — chỉ những bệnh nhân có "Ngày lập" trùng
                #  (các) ngày được chọn bên dưới mới được ghi vào Sheet.)
                # ═══════════════════════════════════════════════
                UNKNOWN_LAP = "__unknown__"
                lap_counts = Counter(r.get("NGÀY LẬP") or UNKNOWN_LAP for r in records)

                def _parse_dmy_safe(s):
                    try:
                        return datetime.strptime(s, "%d/%m/%Y")
                    except Exception:
                        return datetime.min

                known_lap_dates = sorted(
                    (d for d in lap_counts if d != UNKNOWN_LAP),
                    key=_parse_dmy_safe, reverse=True
                )
                lap_order = known_lap_dates + (
                    [UNKNOWN_LAP] if UNKNOWN_LAP in lap_counts else []
                )

                def _lap_label(d):
                    if d == UNKNOWN_LAP:
                        return f"❓ Không rõ ngày · {lap_counts[d]} BN"
                    return f"📅 {d} · {lap_counts[d]} BN"

                lap_pill_options = [_lap_label(d) for d in lap_order]
                lap_option_to_date = dict(zip(lap_pill_options, lap_order))

                today_str = today.strftime("%d/%m/%Y")
                today_opt = next(
                    (o for o, d in lap_option_to_date.items() if d == today_str), None
                )
                default_lap_selection = [today_opt] if today_opt else []

                st.markdown(
                    '<div class="sh"><div class="sh-dot" style="background:#f59e0b"></div>'
                    '<span class="sh-txt">🗓️ Chọn Ngày Lập Để Ghi Vào Google Sheet</span></div>',
                    unsafe_allow_html=True
                )
                st.caption(
                    "Minh Lộ liệt kê bệnh nhân theo ngày vào viện nên 1 file thường chứa "
                    "nhiều \"Ngày lập\" khác nhau. Tích chọn 1 hoặc nhiều ngày bên dưới — "
                    "ví dụ quên ghi hôm qua / hôm kia thì tích chọn thêm — chỉ những bệnh "
                    "nhân được **tạo lịch hẹn** vào (các) ngày đã chọn mới được ghi vào Sheet."
                )

                lap_key = "minhlo_lap_selected"
                if lap_key not in st.session_state:
                    st.session_state[lap_key] = default_lap_selection
                else:
                    # Loại các lựa chọn cũ không còn khớp với file vừa upload
                    # (vd. người dùng vừa đổi sang file khác có ngày khác).
                    st.session_state[lap_key] = [
                        o for o in st.session_state[lap_key] if o in lap_pill_options
                    ]

                qa1, qa2, qa3 = st.columns(3)
                with qa1:
                    if st.button("✅ Chọn Tất Cả", use_container_width=True, key="minhlo_selall"):
                        st.session_state[lap_key] = lap_pill_options.copy()
                        _smart_rerun()
                with qa2:
                    if st.button("📅 Chỉ Hôm Nay", use_container_width=True, key="minhlo_seltoday",
                                 disabled=(today_opt is None)):
                        st.session_state[lap_key] = [today_opt] if today_opt else []
                        _smart_rerun()
                with qa3:
                    if st.button("❌ Bỏ Chọn Tất Cả", use_container_width=True, key="minhlo_selnone"):
                        st.session_state[lap_key] = []
                        _smart_rerun()

                HAS_PILLS = hasattr(st, "pills")
                if HAS_PILLS:
                    selected_lap_opts = st.pills(
                        "Ngày lập", options=lap_pill_options, selection_mode="multi",
                        key=lap_key, label_visibility="collapsed",
                    ) or []
                else:
                    selected_lap_opts = st.multiselect(
                        "Ngày lập", options=lap_pill_options,
                        key=lap_key, label_visibility="collapsed",
                    )

                selected_lap_dates = set(lap_option_to_date[o] for o in selected_lap_opts)
                if selected_lap_dates:
                    filtered_records = [
                        r for r in records
                        if (r.get("NGÀY LẬP") or UNKNOWN_LAP) in selected_lap_dates
                    ]
                else:
                    filtered_records = []
                    st.warning("⚠️ Chưa chọn ngày lập nào — hãy tích chọn ít nhất 1 ngày để xem trước và ghi vào Sheet.")

                st.markdown(
                    f'<div class="pg-info" style="text-align:left;margin:0.3rem 0 1rem">'
                    f'Đã chọn <b>{len(selected_lap_dates)}</b> ngày lập · '
                    f'<b>{len(filtered_records)}</b> / {len(records)} bệnh nhân sẽ được ghi vào Sheet</div>',
                    unsafe_allow_html=True
                )

                # Preview table (chỉ hiển thị các bản ghi ĐÃ CHỌN theo Ngày Lập)
                st.markdown(
                    '<div class="sh"><div class="sh-dot" style="background:#3b82f6"></div>'
                    f'<span class="sh-txt">Xem Trước Dữ Liệu Đã Chọn '
                    f'({min(10,len(filtered_records))}/{len(filtered_records)} bệnh nhân)</span></div>',
                    unsafe_allow_html=True
                )
                if filtered_records:
                    preview_df = pd.DataFrame(filtered_records[:10])
                    st.dataframe(preview_df, use_container_width=True, hide_index=True, height=280)
                else:
                    st.info("Chưa có bệnh nhân nào được chọn để xem trước.")

                # Stats (tính trên phần ĐÃ CHỌN, phản ánh đúng những gì sắp được ghi)
                ngay_hen_list = [r["NGÀY HẸN"] for r in filtered_records if r.get("NGÀY HẸN")]
                ngay_set = set(ngay_hen_list)
                st.markdown(f"""
                <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:0.6rem;margin:0.8rem 0">
                  <div class="kc kc-b" style="padding:0.8rem 1rem">
                    <div class="kc-lbl">Tổng Trong File</div>
                    <div class="kc-val" style="font-size:1.6rem">{len(records)}</div>
                  </div>
                  <div class="kc kc-t" style="padding:0.8rem 1rem">
                    <div class="kc-lbl">Đã Chọn Để Ghi</div>
                    <div class="kc-val" style="font-size:1.6rem;color:#1d4ed8">{len(filtered_records)}</div>
                  </div>
                  <div class="kc kc-g" style="padding:0.8rem 1rem">
                    <div class="kc-lbl">Số Ngày Hẹn</div>
                    <div class="kc-val" style="font-size:1.6rem">{len(ngay_set)}</div>
                  </div>
                  <div class="kc kc-v" style="padding:0.8rem 1rem">
                    <div class="kc-lbl">Chưa Khám</div>
                    <div class="kc-val" style="font-size:1.6rem">{sum(1 for r in filtered_records if "chưa" in r.get("ĐÃ KHÁM","").lower())}</div>
                  </div>
                </div>
                """, unsafe_allow_html=True)

                col_imp1, col_imp2 = st.columns([1, 1])
                with col_imp1:
                    # Download as CSV (no Google Sheet needed) — chỉ phần đã chọn
                    csv_imp = pd.DataFrame(filtered_records).to_csv(index=False, encoding="utf-8-sig")
                    st.download_button(
                        label=f"⬇️ Tải CSV ({len(filtered_records)} BN đã chọn)",
                        data=csv_imp.encode("utf-8-sig"),
                        file_name=f"henkham_minhloc_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                        mime="text/csv",
                        disabled=(len(filtered_records) == 0),
                    )
                with col_imp2:
                    if st.button(f"📤 Import {len(filtered_records)} BN Vào Google Sheet",
                                 use_container_width=True, disabled=(len(filtered_records) == 0)):
                        if not creds_data:
                            st.error("❌ Chưa có credentials. Kiểm tra Streamlit Secrets.")
                        else:
                            with st.spinner(f"Đang ghi {len(filtered_records)} bệnh nhân vào Sheet…"):
                                rows_ok, err_push = push_to_sheet(
                                    creds_data, SHEET_ID, SHEET_NAME, filtered_records
                                )
                            if err_push:
                                st.error(f"❌ {err_push}")
                                st.info("💡 Nếu lỗi Permission: vào Google Sheet → Share → đổi Service Account từ Viewer thành Editor.")
                            else:
                                st.success(
                                    f"✅ Đã thêm thành công **{rows_ok}** dòng vào sheet chính "
                                    f"(theo {len(selected_lap_dates)} ngày lập đã chọn)!"
                                )
                                st.info("🔄 Quay lại tab **📊 Tổng Quan** và nhấn **Làm mới** để xem dữ liệu mới.")
                                st.balloons()
                                # Invalidate cache so next refresh loads new data
                                st.session_state.metrics = None

else:
    if not st.session_state.err:
        st.markdown("""<div class="empty">
          <div class="empty-ico">🏥</div>
          <div class="empty-ttl">Đang tải dữ liệu…</div>
          <div class="empty-dsc">Nếu không thấy sau vài giây,
            nhấn <strong>🔄 Làm mới</strong> ở trên.</div>
        </div>""", unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)
