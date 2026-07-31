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
COL_AGE         = "TUỔI"
COL_CCCD        = "4. SỐ CĂN CƯỚC CÔNG DÂN - CHỨNG MINH THƯ"
COL_STT         = "STT"  # khoá chính người dùng tự thêm — ổn định dù dòng bị thêm/xoá/sắp xếp lại
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
.pt-row  { display:flex; flex-wrap:wrap; align-items:center; gap:0.45rem 0.5rem; margin-bottom:0.35rem; }
.pt-tag  {
    font-size:0.74rem; font-weight:600; padding:0.28rem 0.65rem;
    border-radius:20px; white-space:nowrap; line-height:1.35;
}
.pt-tag-wrap {
    white-space:normal !important; line-height:1.35; max-width:100%;
    word-break:break-word; flex-basis:100%;
}
.pt-tag-date  { background:#eff6ff; color:#1d4ed8; }
.pt-tag-spec  { background:#f0fdf4; color:#166534; }
.pt-tag-doc   { background:#fdf4ff; color:#6b21a8; }
.src-banner {
    display:flex; align-items:center; gap:0.45rem;
    font-size:0.74rem; font-weight:600; line-height:1.35;
    padding:0.45rem 0.75rem; border-radius:8px; margin-top:0.45rem;
}
.src-banner .sb-ico { font-size:0.9rem; flex-shrink:0; }
.src-banner-noi   { background:#f5f3ff; color:#5b21b6; border:1px solid #ddd6fe; }
.src-banner-vl    { background:#fffbeb; color:#92400e; border:1px solid #fde68a; }
.src-banner-other { background:#f8fafc; color:#475569; border:1px solid #e2e8f0; }
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
.rtbl th.grp-hdr { text-align:center; font-size:0.72rem; }
.rtbl th.grp-att { background:#14532d; }
.rtbl th.grp-abs { background:#7f1d1d; }
.rtbl th.sub-att { background:#1e5c3a; text-align:center; }
.rtbl th.sub-abs { background:#8a2a2a; text-align:center; }
.rtbl th.divider-l, .rtbl td.divider-l { border-left:2px solid #cbd5e1; }
.rtbl td.col-att { background:#f0fdf4; }
.rtbl td.col-abs { background:#fef2f2; }
.rtbl tr:nth-child(even) td.col-att { background:#e6faec; }
.rtbl tr:nth-child(even) td.col-abs { background:#fde8e8; }
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
.mrow-meta-chip {
    display:inline-block; font-size:0.7rem; font-weight:600; color:#475569;
    background:#f1f5f9; padding:0.15rem 0.5rem; border-radius:10px; white-space:nowrap;
}

/* Badge Khoa trong bảng chi tiết */
.khoa-badge {
    font-size:0.72rem; font-weight:600; padding:0.28rem 0.65rem;
    border-radius:20px; white-space:normal; display:inline-block; line-height:1.35;
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
.mrow-name { font-size:0.86rem; font-weight:700; color:#0f172a; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.mrow-facts { display:flex; align-items:center; gap:1rem; margin-top:0.4rem; flex-wrap:wrap; }
.mrow-fact  { display:flex; align-items:center; gap:0.32rem; font-size:0.8rem; font-weight:500; color:#334155; white-space:nowrap; }
.mrow-fact .fi { font-size:0.85rem; }
.mrow-fact-phone { color:#1d4ed8; font-weight:700; text-decoration:none; }

/* Cột Khoa + Nguồn: xếp dọc, khoa (tím) trên, nguồn (xanh, gọn) dưới */
.khoa-src-stack { display:flex; flex-direction:column; align-items:flex-start; gap:0.4rem; }
.src-pill {
    display:inline-flex; align-items:center; gap:0.3rem;
    font-size:0.68rem; font-weight:600; padding:0.24rem 0.6rem;
    border-radius:14px; white-space:normal; line-height:1.3;
    max-width:100%; word-break:break-word;
}
.src-pill-noi   { background:#eff6ff; color:#1d4ed8; }
.src-pill-vl    { background:#fffbeb; color:#92400e; }
.src-pill-other { background:#f8fafc; color:#475569; }

/* Badge trạng thái ngắn gọn */
.mrow-status-wrap { display:flex; align-items:center; justify-content:center; }
.status-pill {
    display:inline-block; font-size:0.78rem; font-weight:700;
    padding:0.4rem 0.95rem; border-radius:20px; white-space:nowrap;
}
.status-pill-done    { background:#d1fae5; color:#065f46; }
.status-pill-absent  { background:#e2e8f0; color:#475569; }
.status-pill-pending { background:#fde68a; color:#92400e; }

/* Nút "✏️" / "🗑️" — icon trần, không khung/nền, chỉ nổi màu khi hover,
   khớp phong cách icon-button tối giản trong thiết kế mới */
div[class*="editbtn_"] .stButton>button,
div[class*="delbtn_"] .stButton>button {
    background:transparent !important; border:none !important; box-shadow:none !important;
    outline:none !important; font-size:1.05rem !important; padding:0.35rem !important;
    color:#475569 !important; transform:none !important; line-height:1 !important;
}
div[class*="editbtn_"] .stButton>button:hover,
div[class*="editbtn_"] .stButton>button:focus,
div[class*="editbtn_"] .stButton>button:active {
    background:#eff6ff !important; border:none !important; outline:none !important;
    box-shadow:none !important; border-radius:8px !important; color:#1d4ed8 !important;
}
div[class*="delbtn_"] .stButton>button:hover,
div[class*="delbtn_"] .stButton>button:focus,
div[class*="delbtn_"] .stButton>button:active {
    background:#fef2f2 !important; border:none !important; outline:none !important;
    box-shadow:none !important; border-radius:8px !important; color:#dc2626 !important;
}

/* Dropdown "✏️ Sửa trạng thái" mở ngay dưới hàng — không phải popup mới */
div[class*="editdrop_"] {
    background:#f8fbff; border:1.5px solid #bfdbfe; border-radius:11px;
    padding:0.7rem 0.8rem 0.5rem; margin:0.5rem 0 0.15rem;
}
.edit-dlg-name {
    font-size:0.84rem; font-weight:700; color:#0f172a;
    margin-bottom:0.6rem; text-align:center;
}
div[class*="editdrop_"] div[role="radiogroup"] {
    display:flex; flex-direction:row; gap:0.5rem; flex-wrap:wrap; margin-bottom:0.5rem;
}
div[class*="editdrop_"] div[role="radiogroup"] label {
    background:white; border:1.5px solid #e2e8f0 !important; border-radius:9px !important;
    padding:0.45rem 0.7rem !important;
    font-size:0.74rem !important; font-weight:600 !important; color:#1e293b !important;
    transition:all 0.15s;
}
div[class*="editdrop_"] div[role="radiogroup"] label:hover {
    background:#eff6ff; border-color:#93c5fd !important;
}
div[class*="editdrop_"] div[role="radiogroup"] label[data-checked="true"] {
    background:#dbeafe !important; border-color:#3b82f6 !important;
}

/* Dropdown xác nhận "🗑️ Xóa bệnh nhân" mở ngay dưới hàng */
div[class*="deldrop_"] {
    background:#fff8f8; border:1.5px solid #fecaca; border-radius:11px;
    padding:0.7rem 0.8rem 0.5rem; margin:0.5rem 0 0.15rem;
}
.del-dlg-warn {
    text-align:center; font-size:0.78rem; color:#334155; line-height:1.6;
    margin-bottom:0.6rem;
}
.del-dlg-warn b { color:#991b1b; }
.del-dlg-note { display:block; font-size:0.66rem; color:#dc2626; font-weight:700; margin-top:0.2rem; }

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

/* Mỗi HÀNG BỆNH NHÂN (tab "3 Ngày Tới" và các bảng dùng render_upcoming_table)
   — đóng khung riêng từng người thành 1 "thẻ" gọn, chia 4 vùng bằng gạch dọc
   nhạt: (Tên+SĐT) | (Khoa+Nguồn) | (Trạng thái) | (Sửa/Xóa). */
div[class*="mrow_"] {
    background:#ffffff;
    border:1px solid #d7dce3;
    border-radius:12px;
    padding:0.9rem 1.1rem;
    margin-bottom:0.85rem;
    box-shadow:0 2px 6px rgba(15,23,42,0.08);
    overflow-x:hidden;
}
div[class*="mrow_"]:hover { border-color:#c7d2e0; box-shadow:0 2px 8px rgba(15,23,42,0.08); }
/* 4 cột trong hàng đầu (Tên+SĐT | Khoa+Nguồn | Trạng thái | Sửa/Xóa) —
   canh giữa theo chiều dọc, có gạch dọc phân vùng giữa mỗi cột (trừ cột đầu). */
div[class*="mrow_"] div[data-testid="stHorizontalBlock"] {
    align-items:center !important; gap:0; flex-wrap:nowrap;
}
div[class*="mrow_"] div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:nth-of-type(2),
div[class*="mrow_"] div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:nth-of-type(3),
div[class*="mrow_"] div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:nth-of-type(4) {
    border-left:1px solid #eef0f4; padding-left:0.9rem !important; margin-left:0.9rem;
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
.rtbl td.col-att { background: #f0fdf4 !important; }
.rtbl td.col-abs { background: #fef2f2 !important; }
.rtbl tr:nth-child(even) td.col-att { background: #e6faec !important; }
.rtbl tr:nth-child(even) td.col-abs { background: #fde8e8 !important; }
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


def _load_xlsx_grid(uploaded_file):
    """
    Đọc 1 file .xlsx bất kỳ (kể cả file Minh Lộ xuất zip lệch chuẩn) thành
    lưới ô {(dòng, cột): giá_trị} bằng cách đọc trực tiếp XML bên trong zip —
    không phụ thuộc sharedStrings.xml (một số file chỉ có ô số) và tự fill
    giá trị cho các ô nằm trong vùng merge (gộp ô).

    Dùng CHUNG cho mọi loại báo cáo Minh Lộ (hẹn khám, ĐK KCB, …) để không
    lặp code — mỗi loại báo cáo chỉ khác nhau ở phần "hiểu cột nào là gì"
    (do header khác nhau), còn phần đọc thô xlsx thì giống hệt nhau.

    Trả về (grid, max_row, max_col, get_row_fn, error_msg_or_None).
    """
    import zipfile
    import xml.etree.ElementTree as ET
    import io

    NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    def _tag(name):
        return "{" + NS + "}" + name

    try:
        if hasattr(uploaded_file, "read"):
            raw = uploaded_file.read()
        else:
            with open(uploaded_file, "rb") as f:
                raw = f.read()

        zf = zipfile.ZipFile(io.BytesIO(raw))
        # Một số phần mềm (vd. Minh Lộ HIS) nén zip với dấu '\' thay vì '/'
        # trong tên đường dẫn nội bộ (không đúng chuẩn zip nhưng vẫn mở được).
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
            ws_key = next((k for k in names_lower if "worksheets/sheet" in k), None)
        if ws_key is None:
            return {}, 0, 0, None, "Không tìm thấy worksheet trong file xlsx."

        with zf.open(names_lower[ws_key]) as f:
            ws_tree = ET.parse(f)

        # ── Merged cells → build fill map {(row,col): (src_row,src_col)} ──
        merge_fill = {}
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
        grid = {}
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

        for (r, c), (sr, sc) in merge_fill.items():
            if (sr, sc) in grid and (r, c) not in grid:
                grid[(r, c)] = grid[(sr, sc)]

        max_row = max(r for r, _ in grid) if grid else 0
        max_col = max(c for _, c in grid) if grid else 0

        def get_row(r):
            return [grid.get((r, c), "") for c in range(1, max_col + 1)]

        return grid, max_row, max_col, get_row, None

    except Exception as e:
        return {}, 0, 0, None, f"Lỗi đọc file: {type(e).__name__}: {e}"


def _xlsx_serial_to_date_str(val):
    """Chuyển số serial ngày của Excel/Google Sheets → chuỗi 'dd/mm/yyyy'.
    Trả về None nếu không phải serial ngày hợp lệ."""
    try:
        s = float(val)
        if 20000 < s < 60000:
            return (datetime(1899, 12, 30) + timedelta(days=int(s))).strftime("%d/%m/%Y")
    except Exception:
        pass
    return None


def _parse_minhlo_date(val):
    """Đọc 1 ô ngày trong file Minh Lộ → chuỗi 'dd/mm/yyyy', chấp nhận CẢ 2
    kiểu dữ liệu Minh Lộ có thể xuất ra (tuỳ máy/tuỳ lần xuất báo cáo):

      1. Số serial Excel (ô định dạng SỐ)      → _xlsx_serial_to_date_str
      2. Chữ đã format sẵn (ô định dạng TEXT)  → thử lần lượt vài kiểu phổ
         biến: dd/mm/yyyy, d/m/yyyy, dd-mm-yyyy, yyyy-mm-dd, kể cả khi có
         giờ đính kèm (vd "03/06/2026 07:30:00" → vẫn lấy đúng phần ngày).

    QUAN TRỌNG: bản cũ CHỈ hiểu dạng số — nếu ô là chữ thì bị bỏ qua ÂM
    THẦM (trả về rỗng), khiến bệnh nhân đó biến mất khỏi toàn bộ cửa sổ so
    khớp trong reconcile_attendance (tưởng nhầm là chưa đến khám dù có
    trong file). Đây là nguyên nhân chính gây "check không thấy tên" /
    "lệch ngày" đã gặp trên dữ liệu thực tế.

    Trả về (date_str hoặc "", đã_đọc_được: bool) — cờ thứ 2 để phân biệt
    "ô trống thật sự" (không phải lỗi) với "có giá trị nhưng đọc lỗi".
    """
    raw = str(val).strip()
    if not raw:
        return "", True  # ô trống — không phải lỗi đọc

    serial = _xlsx_serial_to_date_str(raw)
    if serial:
        return serial, True

    # Cắt bỏ phần giờ nếu có (vd "03/06/2026 07:30:00" hoặc "03/06/2026 7:30")
    date_part = re.split(r"\s+", raw, maxsplit=1)[0]

    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y/%m/%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(date_part, fmt).strftime("%d/%m/%Y"), True
        except Exception:
            continue

    return "", False  # có giá trị nhưng không đọc được kiểu nào — LỖI THẬT


def parse_minh_lo_excel(uploaded_file):
    """
    Parse Minh Lo HIS Excel export ("Danh sách bệnh nhân hẹn khám lại").
    Robust against missing sharedStrings.xml (files with only numeric cells).
    Handles merged cells via fill-down logic.
    Returns (list_of_dicts, error_msg_or_None).
    """
    grid, max_row, max_col, get_row, err = _load_xlsx_grid(uploaded_file)
    if err:
        return [], err

    try:
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
                "TUỔI":                tuoi_val or "",
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


def parse_minh_lo_visit_log(uploaded_file):
    """
    Parse báo cáo "ĐK Khám Chữa Bệnh" của Minh Lộ — đây là NHẬT KÝ BỆNH NHÂN
    THỰC TẾ ĐÃ ĐẾN VIỆN (không phải danh sách hẹn), xuất theo 1 khoảng ngày
    tuỳ chọn (vd. "Từ ngày 01/06/2026 đến ngày 30/06/2026"). Dùng file này
    làm "dữ liệu thực tế" để đối chiếu xem bệnh nhân đã hẹn có thực sự đến
    khám hay chưa — bất kể họ đến sớm/muộn hơn ngày hẹn bao nhiêu, vì file
    trải dài cả khoảng thời gian nên không bị sót như so khớp theo 1 ngày.

    Mỗi dòng = 1 lượt đến khám thực tế, gồm: Mã BN, Họ tên, Ngày sinh (đầy đủ,
    CHÍNH XÁC — không phải ước tính từ tuổi), Giới tính, Địa chỉ, Số CMND,
    Ngày ĐK (ngày thực đến khám), Khoa ĐK, Điện thoại, Mã thẻ BHYT.

    Returns (list_of_dicts, error_msg_or_None, warning_msg_or_None).
    error_msg   = lỗi khiến KHÔNG đọc được file (dừng hẳn, data_rows luôn rỗng).
    warning_msg = vẫn đọc được file, nhưng có N dòng bị lỗi ngày cụ thể — vẫn
    trả về đủ data_rows, chỉ cảnh báo để biết mà kiểm tra thủ công riêng.
    """
    grid, max_row, max_col, get_row, err = _load_xlsx_grid(uploaded_file)
    if err:
        return [], err, None

    try:
        # ── Tìm dòng tiêu đề: chứa cả "Mã BN" và "Ngày ĐK" ──
        header_row_idx = None
        for r in range(1, min(max_row + 1, 15)):
            row_vals = [str(v).strip().lower() for v in get_row(r)]
            has_ma_bn = any("mã bn" in v or "ma bn" in v for v in row_vals)
            has_ngay_dk = any(v in ("ngày đk", "ngay dk") for v in row_vals)
            if has_ma_bn and has_ngay_dk:
                header_row_idx = r
                break

        if header_row_idx is None:
            return [], ("Không tìm thấy hàng tiêu đề \"Mã BN\" / \"Ngày ĐK\". "
                        "Kiểm tra đúng loại báo cáo \"ĐK Khám Chữa Bệnh\" của Minh Lộ."), None

        headers = [str(v).strip().replace("\n", " ").lower() for v in get_row(header_row_idx)]

        def find_col(keywords):
            for i, h in enumerate(headers):
                if any(k.lower() in h for k in keywords):
                    return i
            return None

        idx = {
            "ma_bn":     find_col(["mã bn", "ma bn"]),
            "ho_ten":    find_col(["họ tên", "ho ten"]),
            "ngay_sinh": find_col(["ngày tháng năm sinh", "ngay thang nam sinh"]),
            # "năm sinh" (chỉ năm) là chuỗi con của "ngày tháng năm sinh" (cả ngày
            # tháng năm) → phải so khớp CHÍNH XÁC cả ô, không dùng "in" (substring),
            # nếu không sẽ nhầm sang lấy đúng cột "ngày tháng năm sinh" ở trên.
            "nam_sinh":  next((i for i, h in enumerate(headers) if h.strip() == "năm sinh"), None),
            "tuoi":      find_col(["tuổi", "tuoi"]),
            "gioi_tinh": find_col(["giới tính", "gioi tinh"]),
            "xa":        find_col(["xã,phường", "xã", "phường", "xa,phuong"]),
            "huyen":     find_col(["huyện,tỉnh", "huyện", "tỉnh", "huyen,tinh"]),
            "cmnd":      find_col(["số cmnd", "so cmnd", "cmnd"]),
            "ngay_dk":   find_col(["ngày đk", "ngay dk"]),
            "gio_dk":    find_col(["giờ đk", "gio dk"]),
            "khoa_dk":   find_col(["khoa đk", "khoa dk"]),
            "dt":        find_col(["điện thoại", "dien thoai"]),
            "dia_chi":   find_col(["địa chỉ", "dia chi"]),
            "chan_doan": find_col(["chẩn đoán", "chan doan"]),
            "bhyt":      find_col(["mã thẻ bhyt", "ma the bhyt", "bhyt"]),
        }

        def cv(row_vals, key):
            i = idx.get(key)
            if i is None or i >= len(row_vals):
                return ""
            return str(row_vals[i]).strip()

        data_rows = []
        n_bad_date = 0
        for r in range(header_row_idx + 1, max_row + 1):
            row_vals = get_row(r)
            ho_ten = cv(row_vals, "ho_ten")
            if not ho_ten or not re.search(r"[^\W\d_]", ho_ten, re.UNICODE):
                continue  # bỏ dòng trống / dòng không phải dữ liệu bệnh nhân

            ngay_sinh_raw = cv(row_vals, "ngay_sinh")
            ngay_sinh, _ = _parse_minhlo_date(ngay_sinh_raw)
            ngay_dk_raw = cv(row_vals, "ngay_dk")
            ngay_dk, dk_ok = _parse_minhlo_date(ngay_dk_raw)
            if not dk_ok:
                n_bad_date += 1

            dia_chi = cv(row_vals, "dia_chi") or " ".join(
                p for p in [cv(row_vals, "xa"), cv(row_vals, "huyen")] if p
            )

            data_rows.append({
                "MÃ BN":        cv(row_vals, "ma_bn"),
                "HỌ TÊN":       ho_ten,
                "NGÀY SINH":    ngay_sinh,
                "NĂM SINH":     cv(row_vals, "nam_sinh"),
                "TUỔI":         cv(row_vals, "tuoi"),
                "GIỚI TÍNH":    cv(row_vals, "gioi_tinh"),
                "ĐỊA CHỈ":      dia_chi,
                "SỐ CMND":      cv(row_vals, "cmnd"),
                "NGÀY ĐK":      ngay_dk,   # ngày THỰC ĐẾN KHÁM
                "GIỜ ĐK":       cv(row_vals, "gio_dk"),
                "KHOA ĐK":      cv(row_vals, "khoa_dk"),
                "SỐ ĐIỆN THOẠI": _fix_phone(cv(row_vals, "dt")),
                "CHẨN ĐOÁN":    cv(row_vals, "chan_doan"),
                "SỐ BHYT":      cv(row_vals, "bhyt"),
            })

        warn = (f"⚠️ {n_bad_date} dòng trong file không đọc được NGÀY ĐK (định dạng lạ) — "
                f"những dòng này sẽ KHÔNG đối chiếu được, cần kiểm tra thủ công."
                if n_bad_date > 0 else None)
        return data_rows, None, warn

    except Exception as e:
        return [], f"Lỗi đọc file: {type(e).__name__}: {e}", None

# Danh sách cột "mặc định" — chỉ dùng làm PHƯƠNG ÁN DỰ PHÒNG nếu vì lý do
# nào đó không đọc được dòng tiêu đề thực tế trên Google Sheet (xem push_to_sheet).
SHEET_COLUMNS = [
    "Dấu thời gian",                                                    # A  - import timestamp
    "NGUỒN BỆNH NHÂN",                                                  # B  - fixed value
    "TRẠNG THÁI",                                                       # C  - blank
    "NGÀY KHÁM",                                                        # D  - Ngày hẹn
    "1. HỌ VÀ TÊN BỆNH NHÂN",                                          # E  - Họ tên
    "NĂM SINH",                                                         # F  - N/A
    "TUỔI",                                                             #    - Tuổi (ghi kèm để tránh lệch năm sinh)
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

def build_sheet_field_map(record, import_time_str):
    """
    Map một bản ghi Minh Lộ → dict {tên cột trên Google Sheet: giá trị}.

    Dùng dict theo TÊN CỘT (thay vì 1 list theo vị trí cố định) để
    push_to_sheet có thể ghép đúng ô bất kể cột "TUỔI" (hay cột nào khác)
    nằm ở vị trí nào trên Sheet thực tế — không phụ thuộc thứ tự cột cứng.

      Dấu thời gian                = thời gian import file
      NGUỒN BỆNH NHÂN              = "BỆNH NHÂN ĐIỀU TRỊ NỘI KHOA TÁI KHÁM" (IN HOA)
      TRẠNG THÁI                   = "BỆNH NHÂN CHƯA KHÁM/BỎ KHÁM" (IN HOA, mặc định dropdown)
      NGÀY KHÁM                    = NGÀY HẸN từ Excel (dd/mm/yyyy)
      HỌ VÀ TÊN BỆNH NHÂN          = HỌ TÊN từ Excel
      NĂM SINH                     = ước tính từ Tuổi trong Excel (có thể lệch ±1 năm
                                      vì Minh Lộ chỉ xuất Tuổi, không có ngày sinh)
      TUỔI                         = Tuổi THẬT lấy trực tiếp từ Excel Minh Lộ (chính xác,
                                      không bị lệch như Năm sinh ước tính)
      SỐ ĐIỆN THOẠI                = SỐ ĐIỆN THOẠI từ Excel
      ĐỊA CHỈ (THÔN/XÃ)           = ĐỊA CHỈ từ Excel
      KHOA KHÁM CHỮA BỆNH          = KHOA HẸN từ Excel (IN HOA)
      GIỚI TÍNH                    = suy ra từ cột Tuổi Nam/Nữ trong Excel
      TRIỆU CHỨNG CHÍNH / SỐ CĂN CƯỚC / BÁC SĨ MONG MUỐN / GIỜ KHÁM DỰ KIẾN = N/A
      CHUYÊN KHOA MONG MUỐN        = "Other: Bệnh nhân điều trị nội khoa tái khám"
      CAM KẾT / ĐỒNG Ý             = "CÓ"
    """
    ngay_kham = record.get("NGÀY HẸN", "N/A") or "N/A"
    # 3 cột NGUỒN BỆNH NHÂN, TRẠNG THÁI, KHOA KHÁM CHỮA BỆNH luôn IN HOA.
    khoa_hen = str(record.get("KHOA HẸN", "N/A") or "N/A").upper()
    tuoi_val = record.get("TUỔI") or "N/A"

    return {
        "Dấu thời gian":                                                           import_time_str,
        "NGUỒN BỆNH NHÂN":                                                         "BỆNH NHÂN ĐIỀU TRỊ NỘI KHOA TÁI KHÁM",
        "TRẠNG THÁI":                                                              "BỆNH NHÂN CHƯA KHÁM/BỎ KHÁM",
        "NGÀY KHÁM":                                                               ngay_kham,
        "1. HỌ VÀ TÊN BỆNH NHÂN":                                                 record.get("HỌ TÊN", "N/A"),
        "NĂM SINH":                                                                record.get("NĂM SINH (ước tính)") or "N/A",
        "TUỔI":                                                                    tuoi_val,
        "5. SỐ ĐIÊN THOẠI":                                                        record.get("SỐ ĐIỆN THOẠI", "N/A"),
        "2. ĐỊA CHỈ (THÔN/XÃ)":                                                   record.get("ĐỊA CHỈ", "N/A"),
        "KHOA KHÁM CHỮA BỆNH":                                                     khoa_hen,
        "3. GIỚI TÍNH":                                                            record.get("GIỚI TÍNH") or "N/A",
        "1. TRIỆU CHỨNG CHÍNH":                                                    "N/A",
        "4. SỐ CĂN CƯỚC CÔNG DÂN - CHỨNG MINH THƯ":                               "N/A",
        "CHUYÊN KHOA MONG MUỐN KHÁM":                                              "Other: Bệnh nhân điều trị nội khoa tái khám",
        "BÁC SĨ MONG MUỐN ( nếu có)":                                              "N/A",
        "GIỜ KHÁM DỰ KIẾN":                                                        "N/A",
        "1. CAM KẾT CÁC THÔNG TIN LÀ THÔNG TIN ĐÚNG, CHỊU TRÁCH NHIỆM TRƯỚC PHÁP LUẬT TRƯỚC NHỮNG THÔNG TIN ĐÃ CUNG CẤP TRÊN": "CÓ",
        "ĐỒNG Ý CÁC ĐIỀU KHOẢN ĐẶT LỊCH KHÁM ONLINE TẠI BVĐK TÂM ĐỨC CẦU QUAN":    "CÓ",
    }


def push_to_sheet(creds_data, sheet_id, sheet_name, records):
    """
    Append parsed Minh Lo records into the MAIN Google Sheet tab.

    Ghép giá trị vào ĐÚNG CỘT theo TÊN TIÊU ĐỀ đọc trực tiếp từ dòng 1 của
    Sheet thực tế (không giả định vị trí cột cố định) — nên nếu người dùng
    thêm/xóa/đổi chỗ cột nào đó (ví dụ thêm cột "TUỔI") trên Google Sheet,
    việc ghi dữ liệu vẫn tự động khớp đúng cột mà không cần sửa code.
    Cột nào trên Sheet không nằm trong danh sách trường đã biết sẽ được để trống.
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

        # ── Đọc dòng tiêu đề THỰC TẾ trên Sheet để biết đúng thứ tự cột ──
        all_vals = ws.get_all_values()
        headers = [h.strip() for h in all_vals[0]] if all_vals else []
        if not headers:
            headers = SHEET_COLUMNS  # phương án dự phòng nếu Sheet trống

        date_col_idx = headers.index(COL_EXAM_DATE) if COL_EXAM_DATE in headers else None

        def date_str_to_serial(s):
            """Chuyển 'dd/mm/yyyy' → số serial Google Sheets (float).
            Trả về chuỗi gốc nếu không parse được."""
            try:
                d = datetime.strptime(str(s).strip(), "%d/%m/%Y")
                delta = d - datetime(1899, 12, 30)
                return delta.days  # số nguyên, Google Sheets tự hiểu là Date
            except Exception:
                return s  # giữ nguyên nếu không parse được

        # Build rows khớp ĐÚNG THEO TÊN CỘT thật trên Sheet
        converted_rows = []
        for r in records:
            field_map = build_sheet_field_map(r, import_time_str)
            row = [field_map.get(h, "") for h in headers]
            if date_col_idx is not None:
                row[date_col_idx] = date_str_to_serial(row[date_col_idx])
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


def _resolve_row_by_stt(ws, headers, stt):
    """
    Tra ra SỐ DÒNG THỰC TẾ trên Google Sheet ứng với 1 giá trị STT (khoá
    chính, cột do người dùng tự thêm để đánh số ổn định cho từng bệnh nhân
    — không đổi dù dòng bị thêm/xoá/sắp xếp lại).

    Đọc TRỰC TIẾP từ Sheet ngay tại thời điểm ghi (không dùng lại vị trí
    dòng đã cache từ lúc tải dữ liệu trước đó) — đây là điểm mấu chốt để
    tránh ghi nhầm sang bệnh nhân khác nếu Sheet đã bị chỉnh sửa (thêm/xoá
    dòng, sắp xếp lại) trong khoảng thời gian giữa lúc tải dữ liệu và lúc
    bấm cập nhật.

    Trả về (sheet_row: int | None, err: str | None).
    """
    if COL_STT not in headers:
        return None, f"Không tìm thấy cột '{COL_STT}' trên Google Sheet."
    stt_str = str(stt).strip()
    if not stt_str:
        return None, "Thiếu giá trị STT để tra dòng."
    col_idx = headers.index(COL_STT) + 1
    col_vals = ws.col_values(col_idx)  # bao gồm cả dòng tiêu đề ở vị trí 0
    for i, v in enumerate(col_vals):
        if i == 0:
            continue  # bỏ qua dòng tiêu đề
        if str(v).strip() == stt_str:
            return i + 1, None  # gspread dùng chỉ số dòng 1-based
    return None, f"Không tìm thấy bệnh nhân có STT = {stt_str} trên Google Sheet (có thể dòng đã bị xoá)."


def update_patient_status(creds_data, sheet_id, sheet_name, sheet_row, new_status, stt=None):
    """
    Cập nhật cột TRẠNG THÁI cho 1 bệnh nhân. Chỉ ghi đúng 1 ô, không đụng
    tới các cột khác. Trả về (thành_công: bool, lỗi: str | None).

    - Nếu truyền `stt` (khoá chính, khuyến khích dùng): tra lại đúng số dòng
      NGAY TẠI THỜI ĐIỂM GHI theo giá trị STT — an toàn dù Sheet đã bị
      thêm/xoá/sắp xếp lại dòng từ lúc tải dữ liệu.
    - Nếu không có `stt`: dùng `sheet_row` truyền vào như cũ (tương thích
      ngược cho các chỗ chưa có cột STT).
    """
    try:
        cl = authenticate_rw(creds_data)
        ss = cl.open_by_key(sheet_id)
        ws = ss.worksheet(sheet_name)

        headers = ws.row_values(1)
        if COL_STATUS not in headers:
            return False, f"Không tìm thấy cột '{COL_STATUS}' trên Google Sheet."
        col_idx = headers.index(COL_STATUS) + 1  # gspread dùng chỉ số 1-based

        if stt is not None:
            resolved_row, err = _resolve_row_by_stt(ws, headers, stt)
            if err:
                return False, err
            sheet_row = resolved_row

        ws.update_cell(sheet_row, col_idx, new_status)
        return True, None
    except Exception as e:
        return False, f"Lỗi cập nhật trạng thái: {type(e).__name__}: {e}"


def delete_patient_row(creds_data, sheet_id, sheet_name, sheet_row, stt=None):
    """
    Xóa hẳn 1 dòng bệnh nhân khỏi Google Sheet. Trả về (thành_công: bool, lỗi: str | None).
    Truyền `stt` để tra lại đúng dòng theo khoá chính ngay tại thời điểm xoá
    (an toàn hơn dùng sheet_row đã cache) — xem update_patient_status.
    """
    try:
        cl = authenticate_rw(creds_data)
        ss = cl.open_by_key(sheet_id)
        ws = ss.worksheet(sheet_name)
        if stt is not None:
            headers = ws.row_values(1)
            resolved_row, err = _resolve_row_by_stt(ws, headers, stt)
            if err:
                return False, err
            sheet_row = resolved_row
        ws.delete_rows(sheet_row)
        return True, None
    except Exception as e:
        return False, f"Lỗi xóa bệnh nhân: {type(e).__name__}: {e}"


def update_patient_status_batch(creds_data, sheet_id, sheet_name, updates):
    """
    Cập nhật cột TRẠNG THÁI cho NHIỀU bệnh nhân cùng lúc bằng 1 lần gọi API
    (batch_update) thay vì gọi update_cell lặp lại từng dòng — nhanh hơn và
    tránh bị Google giới hạn số request khi đối chiếu hàng loạt.

    updates: list các tuple (sheet_row, new_status, stt_hoặc_None).
             Nếu phần tử thứ 3 (stt) có giá trị, dòng ghi sẽ được TRA LẠI
             theo khoá chính STT ngay tại thời điểm ghi (an toàn hơn, xem
             update_patient_status) — nếu không, dùng sheet_row như cũ.
    Trả về (số dòng đã cập nhật, lỗi | None).
    """
    if not updates:
        return 0, None
    try:
        cl = authenticate_rw(creds_data)
        ss = cl.open_by_key(sheet_id)
        ws = ss.worksheet(sheet_name)

        headers = ws.row_values(1)
        if COL_STATUS not in headers:
            return 0, f"Không tìm thấy cột '{COL_STATUS}' trên Google Sheet."
        col_letter = gspread.utils.rowcol_to_a1(1, headers.index(COL_STATUS) + 1)
        col_letter = "".join(ch for ch in col_letter if ch.isalpha())

        body = []
        for u in updates:
            sheet_row, new_status = u[0], u[1]
            stt = u[2] if len(u) > 2 else None
            if stt is not None:
                resolved_row, err = _resolve_row_by_stt(ws, headers, stt)
                if err:
                    return len(body), err
                sheet_row = resolved_row
            body.append({"range": f"{col_letter}{sheet_row}", "values": [[new_status]]})
        ws.batch_update(body, value_input_option="USER_ENTERED")
        return len(body), None
    except Exception as e:
        return 0, f"Lỗi cập nhật hàng loạt: {type(e).__name__}: {e}"


def update_patient_fields(creds_data, sheet_id, sheet_name, sheet_row, field_values, stt=None):
    """
    Cập nhật NHIỀU CỘT cùng lúc cho 1 bệnh nhân (1 dòng) — dùng cho việc sửa
    trực tiếp SĐT/Năm sinh ngay trên web ở bước đối chiếu tái khám, khi bệnh
    nhân rơi vào danh sách "cần kiểm tra thủ công" (chỉ khớp được mỗi tên).

    field_values: dict {tên_cột_trên_sheet: giá_trị_mới}, vd
      {"5. SỐ ĐIÊN THOẠI": "0987654321", "NĂM SINH": "1985"}
    Truyền `stt` để tra lại đúng dòng theo khoá chính (xem update_patient_status).
    Trả về (thành_công: bool, lỗi | None).
    """
    if not field_values:
        return True, None
    try:
        cl = authenticate_rw(creds_data)
        ss = cl.open_by_key(sheet_id)
        ws = ss.worksheet(sheet_name)

        headers = ws.row_values(1)
        if stt is not None:
            resolved_row, err = _resolve_row_by_stt(ws, headers, stt)
            if err:
                return False, err
            sheet_row = resolved_row
        body = []
        for col_name, new_val in field_values.items():
            if col_name not in headers:
                continue
            col_letter = gspread.utils.rowcol_to_a1(1, headers.index(col_name) + 1)
            col_letter = "".join(ch for ch in col_letter if ch.isalpha())
            body.append({"range": f"{col_letter}{sheet_row}", "values": [[new_val]]})
        if not body:
            return False, "Không tìm thấy cột nào khớp trên Google Sheet."
        ws.batch_update(body, value_input_option="USER_ENTERED")
        return True, None
    except Exception as e:
        return False, f"Lỗi cập nhật thông tin: {type(e).__name__}: {e}"


# ═══════════════════════════════════════════════════════════════
# ĐỐI CHIẾU TÁI KHÁM — so khớp danh sách đã hẹn (Google Sheet) với
# nhật ký bệnh nhân THỰC TẾ đến khám (file "Báo cáo ĐK KCB" Minh Lộ),
# không phụ thuộc đúng 1 ngày (bệnh nhân có thể đến sớm/muộn hơn hẹn).
# ═══════════════════════════════════════════════════════════════

def _norm_name(s):
    """Chuẩn hoá tên để so khớp:
    1. Thay chữ Đ/đ đặc thù tiếng Việt (không bỏ dấu được qua NFKD).
    2. NFKD + bỏ combining marks → ASCII base.
    3. Gộp khoảng trắng thừa, viết hoa.
    Kết quả: 'Lê Hoàng Phúc'→'LE HOANG PHUC', 'Đỗ Văn An'→'DO VAN AN'.
    """
    import unicodedata as _ud
    s2 = str(s or "").replace("Đ", "D").replace("đ", "d")
    s2 = _ud.normalize("NFKD", s2)
    s2 = "".join(c for c in s2 if not _ud.combining(c))
    return re.sub(r"\s+", " ", s2.strip()).upper()


def _norm_phone_key(s):
    digits = re.sub(r"\D", "", str(s or ""))
    if len(digits) == 9:
        digits = "0" + digits
    return digits if len(digits) >= 9 and set(digits) != {"0"} else ""


def _norm_cccd(s):
    """Chuẩn hoá số CCCD/CMND để so khớp: chỉ giữ chữ số.

    QUAN TRỌNG: CCCD (căn cước công dân) chuẩn có ĐÚNG 12 chữ số, và rất
    nhiều mã tỉnh bắt đầu bằng số 0 (vd. 001-096). Khi Excel/Google Sheets
    lưu ô này dưới dạng SỐ (number) thay vì văn bản (text) — điều này xảy
    ra rất phổ biến khi xuất báo cáo từ Minh Lộ — số 0 đầu tiên bị MẤT,
    biến 12 số thành 11 số (y hệt lỗi mất số 0 đầu ở số điện thoại đã xử lý
    ở _norm_phone_key). Trên dữ liệu thực tế đã kiểm tra, ~89% bản ghi CCCD
    trong log Minh Lộ bị lỗi này — nếu không khôi phục lại số 0, hầu hết các
    trường hợp so khớp CCCD sẽ thất bại một cách âm thầm dù đúng là cùng 1
    người, khiến bệnh nhân bị đẩy nhầm xuống tầng so khớp thấp hơn (tên+năm
    sinh) hoặc bị coi là "chưa đến khám".

    CMND cũ (9 chữ số, trước khi đổi sang CCCD 12 số) giữ nguyên, không pad.
    """
    d = re.sub(r"\D", "", str(s or ""))
    if len(d) == 11:
        d = "0" + d
    return d if len(d) >= 8 else ""


def _name_similarity(a, b):
    """Độ giống nhau giữa 2 tên, 0.0 → 1.0 (dùng difflib, không cần thư viện ngoài)."""
    import difflib
    a, b = _norm_name(a), _norm_name(b)
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()


def _name_match_ok(a, b, min_ratio):
    """
    Xét 2 tên có nên coi là CÙNG 1 NGƯỜI hay không — chặt hơn hẳn so với
    chỉ dùng độ giống ký tự toàn chuỗi (_name_similarity), vốn dễ khớp
    NHẦM giữa những bệnh nhân KHÁC NHAU khi tên tiếng Việt trùng nhiều từ
    đệm phổ biến. Ví dụ thực tế đã gặp — 4 người HOÀN TOÀN KHÁC NHAU bị
    lẫn vào nhau: "NGUYỄN THỊ ANH", "LÊ THỊ HOÀN", "NGUYỄN THỊ HOAN",
    "NGUYỄN THỊ VÂN" — vì cùng cấu trúc "X THỊ Y" nên tỉ lệ ký tự trùng
    tổng thể vẫn có thể vượt ngưỡng dù rõ ràng là 2 người khác nhau.

    QUY TẮC: coi là CÙNG 1 NGƯỜI chỉ khi:
      1. Từ ĐẦU (họ) và từ CUỐI (tên gọi — phần phân biệt rõ nhất giữa
         những người Việt khác nhau, kể cả cùng họ) khớp CHÍNH XÁC sau khi
         đã chuẩn hoá (bỏ dấu, viết hoa) — bắt buộc, không có ngoại lệ.
         KHÔNG dùng độ giống ký tự cho phần này vì "ANH"/"HOAN"/"VÂN" tuy
         ngắn nhưng là đúng chỗ khác nhau giữa các người, càng giống ký tự
         xét lẫn lộn càng dễ khớp nhầm.
      2. VÀ độ giống ký tự toàn chuỗi (đã tính cả từ đệm ở giữa, để bắt lỗi
         gõ/chính tả nhẹ trong từ đệm) đạt tối thiểu `min_ratio`.

    Trả về (đạt: bool, tỉ lệ giống toàn chuỗi: float 0.0–1.0).
    """
    na, nb = _norm_name(a), _norm_name(b)
    if not na or not nb:
        return False, 0.0
    if na == nb:
        return True, 1.0

    ta, tb = na.split(), nb.split()
    if not ta or not tb:
        return False, 0.0
    if ta[0] != tb[0] or ta[-1] != tb[-1]:
        # Khác họ hoặc khác tên gọi → chắc chắn KHÔNG cùng 1 người, bất kể
        # độ giống ký tự tổng thể cao đến đâu (thường cao giả tạo do trùng
        # từ đệm như "THỊ"/"VĂN").
        return False, _name_similarity(a, b)

    import difflib
    ratio = difflib.SequenceMatcher(None, na, nb).ratio()
    return ratio >= min_ratio, ratio


def _parse_ddmmyyyy(s):
    try:
        return datetime.strptime(str(s).strip(), "%d/%m/%Y").date()
    except Exception:
        return None


def _parse_age(s):
    """Lấy số tuổi từ chuỗi kiểu '68 tuổi', '68', 68 → 68 (int) hoặc None."""
    m = re.search(r"\d+", str(s or ""))
    return int(m.group()) if m else None


RECONCILE_LOOKBACK_DAYS = 7   # danh sách gốc: bệnh nhân có NGÀY KHÁM trong X ngày gần đây
RECONCILE_WINDOW_BEFORE = 2   # cửa sổ so khớp: chấp nhận đến SỚM hơn hẹn tối đa 2 ngày
RECONCILE_WINDOW_AFTER  = 3   # và MUỘN hơn hẹn tối đa 3 ngày — nhưng KHÔNG BAO GIỜ vượt quá HÔM NAY
                              # (xem hàm bên dưới) để không "vồ nhầm" hồ sơ của một đợt khám CŨ khác.

def reconcile_attendance(sheet_patients, visit_records,
                          window_before=RECONCILE_WINDOW_BEFORE,
                          window_after=RECONCILE_WINDOW_AFTER,
                          today=None):
    """
    Đối chiếu danh sách bệnh nhân ĐÃ HẸN (từ Google Sheet, trong
    RECONCILE_LOOKBACK_DAYS ngày gần nhất) với nhật ký bệnh nhân THỰC TẾ đến
    khám (file "Báo cáo ĐK KCB" Minh Lộ) để xác định ai đã đến khám — không bị
    sót vì đến sớm/muộn hơn hẹn.

    CỬA SỔ NGÀY — CO GIÃN THEO TỪNG BỆNH NHÂN, LUÔN CHẶN Ở HÔM NAY:
      Với 1 bệnh nhân có ngày hẹn D:
        window_start = D − window_before
        window_end   = min(D + window_after, HÔM NAY)
      Ví dụ hôm nay là 27/7: bệnh nhân hẹn ngày 20/7 → cửa sổ [18/7, 23/7];
      hẹn ngày 25/7 → cửa sổ [23/7, 27/7] (window_end bị chặn ở 27, không
      phải 28). Việc chặn ở HÔM NAY (thay vì cộng cứng window_after) là điểm
      mấu chốt để KHÔNG bị "vồ nhầm" hồ sơ của MỘT ĐỢT KHÁM CŨ khác của cùng
      bệnh nhân — ví dụ: khám ngày 20/7, hẹn tái khám ngày 27/7 (tạo lịch
      hẹn ngay trong đợt khám 20/7 đó). Nếu so theo cửa sổ rộng cố định
      (vd ±14 ngày quanh 27/7), thuật toán sẽ thấy hồ sơ ngày 20/7 nằm
      trong cửa sổ và tưởng nhầm đó là bằng chứng bệnh nhân đã tái khám
      ngày 27/7 — trong khi thực ra họ CHƯA quay lại. Cửa sổ hẹp + chặn ở
      hôm nay giải quyết đúng vấn đề này.

    THUẬT TOÁN so khớp ĐÚNG NGƯỜI — 3 TẦNG, ưu tiên theo thứ tự, dừng ở
    tầng đầu tiên tìm được ứng viên phù hợp:

      TẦNG 1 — Tên + SĐT (cả 2 cùng khớp):
        Trong số các lượt khám (đã lọc theo cửa sổ ngày), tìm bản ghi có
        SỐ ĐIỆN THOẠI khớp CHÍNH XÁC (đã chuẩn hoá) VÀ Tên đủ giống
        (≥ NAME_RATIO_MIN, sau khi bỏ dấu) — bắt buộc CẢ HAI, không chỉ
        riêng SĐT, để tránh nhận nhầm người nhà dùng chung 1 số điện thoại
        (vd. "Lê Hoàng Phúc" bị gán nhầm cho lượt khám của người thân có
        cùng SĐT nhưng tên khác hẳn). Khớp được → attended_sure.

      TẦNG 2 — Tên + Năm sinh (±1) (chỉ chạy khi Tầng 1 không ra kết quả):
        Trong số các lượt khám còn lại, tìm bản ghi có Tên đủ giống
        (≥ NAME_RATIO_MIN) VÀ Năm sinh khớp chính xác hoặc lệch đúng 1
        (do năm sinh trên Sheet có thể là ước tính từ tuổi). Khớp được
        → attended_sure.

      TẦNG 3 — CHỈ Tên (chỉ chạy khi Tầng 1 và 2 đều không ra kết quả):
        Tìm bản ghi có Tên đủ giống nhất (≥ NAME_RATIO_MIN) mà không cần
        SĐT hay năm sinh trùng khớp. Vì chỉ dựa vào mỗi cái tên — dễ trùng
        giữa nhiều người — LUÔN đưa vào danh sách "CẦN KIỂM TRA THỦ CÔNG"
        (attended_unsure), KHÔNG BAO GIỜ tự động kết luận "chắc chắn đã
        đến" dù tên khớp tuyệt đối 100%.

      Không tìm được ứng viên nào ở cả 3 tầng → not_attended (chưa đến khám).

    sheet_patients: list dict {"sheet_row","name","phone","birth_year","age","exam_date","source"}
    visit_records:  list dict từ parse_minh_lo_visit_log()

    Trả về list dict, 1 phần tử / bệnh nhân đã hẹn, gồm:
      sheet_row, name, phone, age, birth_year, exam_date, source, score,
      visit (bản ghi khớp hoặc None), match_tier (1/2/3/None), status:
        "attended_sure"   → ĐÃ ĐẾN khám chắc chắn (Tầng 1 hoặc Tầng 2)
        "attended_unsure" → CẦN KIỂM TRA THỦ CÔNG (chỉ khớp Tầng 3 — mỗi tên)
        "not_attended"    → CHƯA ĐẾN khám (không tìm thấy trong cửa sổ)
    """
    today_d = today if today is not None else datetime.now().date()

    NAME_RATIO_MIN = 0.92  # ngưỡng độ giống tên tối thiểu, dùng chung cả 3 tầng

    results = []
    for p in sheet_patients:
        p_name = p.get("name", "")
        p_phone_key = _norm_phone_key(p.get("phone"))
        p_birth = p.get("birth_year")
        try:
            p_birth = int(str(p_birth).strip()) if p_birth else None
        except Exception:
            p_birth = None

        exam_date = p.get("exam_date")

        # ── Lọc ứng viên theo cửa sổ ngày RIÊNG của bệnh nhân này ──
        if exam_date:
            w_start = exam_date - timedelta(days=window_before)
            w_end = min(exam_date + timedelta(days=window_after), today_d)
            candidates = []
            for v in visit_records:
                vd = _parse_ddmmyyyy(v.get("NGÀY ĐK"))
                if vd is not None and w_start <= vd <= w_end:
                    candidates.append(v)
        else:
            # Không xác định được ngày hẹn (hiếm) → không đủ an toàn để giới
            # hạn cửa sổ, coi như không có ứng viên (tránh khớp bừa).
            candidates = []

        entry = {
            "sheet_row": p.get("sheet_row"), "stt": p.get("stt", ""), "name": p_name,
            "phone": p.get("phone", ""), "age": p.get("age", ""),
            "birth_year": p.get("birth_year", ""),
            "exam_date": exam_date, "source": p.get("source", ""),
            "visit": None, "score": 0.0, "match_tier": None, "near_miss": None,
        }

        # Cache độ giống tên cho từng ứng viên (tính 1 lần, dùng lại cả 3 tầng).
        # Chỉ giữ lại ứng viên ĐẠT gate họ+tên gọi (xem _name_match_ok) — loại
        # ngay từ đây những tên "trông giống" nhưng thực chất khác người.
        name_sims = []
        for v in candidates:
            ok, ratio = _name_match_ok(p_name, v.get("HỌ TÊN"), NAME_RATIO_MIN)
            if ok:
                name_sims.append((ratio, v))

        # ── TẦNG 1 — Tên + SĐT (bắt buộc cả 2) ──────────────────
        best1 = None
        if p_phone_key:
            for ns, v in name_sims:
                if ns >= NAME_RATIO_MIN and _norm_phone_key(v.get("SỐ ĐIỆN THOẠI")) == p_phone_key:
                    if best1 is None or ns > best1[0]:
                        best1 = (ns, v)
        if best1 is not None:
            entry["visit"] = best1[1]
            entry["score"] = round(100 * best1[0], 1)
            entry["match_tier"] = 1
            entry["status"] = "attended_sure"
            results.append(entry)
            continue

        # ── TẦNG 2 — Tên + Năm sinh (±1) ─────────────────────────
        best2 = None
        for ns, v in name_sims:
            if ns < NAME_RATIO_MIN:
                continue
            try:
                v_birth = int(str(v.get("NĂM SINH", "")).strip())
            except Exception:
                v_birth = None
            if p_birth is None or v_birth is None:
                continue
            if abs(v_birth - p_birth) <= 1:
                if best2 is None or ns > best2[0]:
                    best2 = (ns, v)
        if best2 is not None:
            entry["visit"] = best2[1]
            entry["score"] = round(100 * best2[0], 1)
            entry["match_tier"] = 2
            entry["status"] = "attended_sure"
            results.append(entry)
            continue

        # ── TẦNG 3 — Chỉ Tên → LUÔN vào danh sách cần kiểm tra thủ công ──
        best3 = None
        for ns, v in name_sims:
            if ns >= NAME_RATIO_MIN and (best3 is None or ns > best3[0]):
                best3 = (ns, v)
        if best3 is not None:
            entry["visit"] = best3[1]
            entry["score"] = round(100 * best3[0], 1)
            entry["match_tier"] = 3
            entry["status"] = "attended_unsure"
        else:
            entry["score"] = 0.0
            entry["status"] = "not_attended"
            # ── "NGHI BỊ SÓT" — trước khi kết luận hẳn "chưa khám", thử tìm
            # tên giống trong TOÀN BỘ file (bỏ giới hạn cửa sổ ngày). Nếu có,
            # rất có thể bệnh nhân THỰC SỰ đã đến khám nhưng ngoài cửa sổ cho
            # phép (đến sớm/muộn hơn hẹn quá xa, hoặc NGÀY KHÁM trên Sheet bị
            # sai) — báo rõ ra để soát tay thay vì im lặng kết luận nhầm
            # "chưa khám" (đây chính là kiểu lỗi "bị sót" gặp trên thực tế).
            near = None
            for v in visit_records:
                ok, ratio = _name_match_ok(p_name, v.get("HỌ TÊN"), NAME_RATIO_MIN)
                if ok and (near is None or ratio > near[0]):
                    near = (ratio, v)
            if near is not None:
                vd = _parse_ddmmyyyy(near[1].get("NGÀY ĐK"))
                entry["near_miss"] = {
                    "visit": near[1], "score": round(100 * near[0], 1), "visit_date": vd,
                }

        results.append(entry)

    return results


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
        g = df[COL_GENDER].astype(str).str.strip().str.upper()
        g = g.replace({"NU": "NỮ"})
        g = g[g.isin(["NAM", "NỮ"])]
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

    d["_att"] = d[COL_STATUS].astype(str).str.upper() == STATUS_ATTENDED.upper()
    if COL_SOURCE in d.columns:
        src_vals = d[COL_SOURCE].astype(str)
        d["_tk"] = src_vals.str.contains("khoa|tái|nội trú|xuất viện|tai", case=False, na=False)
        d["_vl"] = src_vals.str.contains("vãng lai|vang lai|ngoài|ngoai", case=False, na=False)
    else:
        d["_tk"] = False
        d["_vl"] = False

    grp   = d.groupby("Kỳ", sort=False)
    stats = grp.size().reset_index(name="Đăng ký")
    # ── Đến khám / Vắng, tách riêng theo TỪNG NGUỒN trước, rồi mới cộng tổng ──
    stats["Đến - Tái Khám"] = grp.apply(lambda g: (g["_att"] & g["_tk"]).sum()).values
    stats["Đến - Vãng Lai"] = grp.apply(lambda g: (g["_att"] & g["_vl"]).sum()).values
    stats["Vắng - Tái Khám"] = grp.apply(lambda g: (g["_tk"] & ~g["_att"]).sum()).values
    stats["Vắng - Vãng Lai"] = grp.apply(lambda g: (g["_vl"] & ~g["_att"]).sum()).values
    stats["Đã khám"]    = grp.apply(lambda g: g["_att"].sum()).values
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
        return f'<span class="pt-tag pt-tag-wrap src-noi">🏥 {s}</span>'
    elif any(k in s.lower() for k in ["vãng lai","vang lai","ngoài","ngoai"]):
        return f'<span class="pt-tag pt-tag-wrap src-vl">🚶 {s}</span>'
    else:
        return f'<span class="pt-tag pt-tag-wrap src-other">👤 {s}</span>'

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

def paginate_list(items, state_key, page_size=PAGE_SIZE):
    """
    Phân trang 1 LIST bất kỳ (khác render_paginated_cards vốn chỉ nhận
    DataFrame) — dùng cho danh sách dict, vd. kết quả đối chiếu tái khám.
    Session_state riêng theo state_key, tự reset về trang 1 khi độ dài
    danh sách thay đổi (đổi bộ lọc / chạy đối chiếu lại).

    Trả về: (page_items, cur, total_pages, start, end, total)
    """
    total = len(items)
    if total == 0:
        return [], 1, 1, 0, 0, 0
    total_pages = max(1, (total + page_size - 1) // page_size)
    fp_key = state_key + "_fp"
    if state_key not in st.session_state:
        st.session_state[state_key] = 1
    if st.session_state.get(fp_key) != total:
        st.session_state[fp_key] = total
        st.session_state[state_key] = 1
    cur = min(max(1, st.session_state[state_key]), total_pages)
    st.session_state[state_key] = cur
    start = (cur - 1) * page_size
    end = start + page_size
    return items[start:end], cur, total_pages, start, end, total


def render_pagination_bar(state_key, cur, total_pages, start, end, total, label="bệnh nhân", widget_key=None):
    """Thanh điều hướng phân trang « ‹ [số trang…] › » dùng chung, khớp UI
    với render_paginated_cards / render_upcoming_table.

    state_key  : khoá session_state lưu trang hiện tại (dùng để ĐỌC/GHI trang).
    widget_key : tiền tố khoá riêng cho các nút bấm/container của LẦN GỌI này —
                 mặc định = state_key. Truyền khác nhau khi cùng 1 state_key
                 được hiển thị thanh điều hướng 2 nơi trên trang (vd. trên đầu
                 VÀ dưới cuối 1 danh sách dài) để tránh trùng khoá widget.
    """
    if total_pages <= 1:
        return
    wkey = widget_key or state_key
    st.markdown(
        f'<div class="pg-info">Trang <b>{cur}</b>/<b>{total_pages}</b> '
        f'&nbsp;·&nbsp; Hiển thị <b>{start+1}–{min(end,total)}</b> / <b>{total}</b> {label}</div>',
        unsafe_allow_html=True
    )
    nums = _paginate_page_numbers(cur, total_pages)
    with st.container(key=f"{wkey}_pgrow"):
        cols = st.columns([1] + [1] * len(nums) + [1])
        with cols[0]:
            with st.container(key=f"{wkey}_navbtn_prev"):
                if st.button("‹", key=f"{wkey}_prev", disabled=(cur == 1), use_container_width=True):
                    st.session_state[state_key] = cur - 1
                    _smart_rerun()
        for i, p in enumerate(nums):
            with cols[1 + i]:
                if p is None:
                    st.markdown(
                        '<div style="text-align:center;color:#94a3b8;font-size:0.68rem;'
                        'height:1.7rem;line-height:1.7rem">…</div>',
                        unsafe_allow_html=True
                    )
                else:
                    is_cur = (p == cur)
                    if is_cur:
                        with st.container(key=f"{wkey}_curbtn_{p}"):
                            st.button(str(p), key=f"{wkey}_p{p}", disabled=True, use_container_width=True)
                    else:
                        if st.button(str(p), key=f"{wkey}_p{p}", use_container_width=True):
                            st.session_state[state_key] = p
                            _smart_rerun()
        with cols[1 + len(nums)]:
            with st.container(key=f"{wkey}_navbtn_next"):
                if st.button("›", key=f"{wkey}_next", disabled=(cur == total_pages), use_container_width=True):
                    st.session_state[state_key] = cur + 1
                    _smart_rerun()


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

def source_pill_html(src_val):
    """Badge nguồn bệnh nhân dạng pill GỌN (khác src-banner — dùng khi đã có
    cột riêng đủ hẹp, đặt ngay dưới badge khoa, như trong layout 4-cột)."""
    s = str(src_val).strip()
    if not s or s in ("nan", ""):
        return ""
    if any(k in s.lower() for k in ["khoa", "tái", "nội trú", "xuất viện", "tai"]):
        cls, ico = "src-pill-noi", "🏥"
    elif any(k in s.lower() for k in ["vãng lai", "vang lai", "ngoài", "ngoai"]):
        cls, ico = "src-pill-vl", "🚶"
    else:
        cls, ico = "src-pill-other", "👤"
    return f'<span class="src-pill {cls}">{ico} {s}</span>'


def status_pill_html(status_str):
    """Badge trạng thái NGẮN GỌN (Đã khám / Chưa khám / Bỏ khám) thay vì in
    nguyên văn cả chuỗi trạng thái dài trên Sheet — dễ quét mắt hơn nhiều
    khi danh sách dài. Suy ra nhãn từ nội dung chuỗi gốc."""
    s = str(status_str or "").upper()
    if STATUS_ATTENDED.upper() in s:
        return '<span class="status-pill status-pill-done">Đã khám</span>'
    elif "BỎ" in s and "CHƯA" not in s:
        return '<span class="status-pill status-pill-absent">Bỏ khám</span>'
    else:
        return '<span class="status-pill status-pill-pending">Chưa khám</span>'


def patient_row_name_facts_html(row2):
    """Cột 1: tên bệnh nhân + giờ hẹn / năm sinh / SĐT dạng icon+chữ gọn
    (không phải pill nền màu) — khớp layout 4 cột có gạch dọc phân vùng."""
    name  = str(row2.get(COL_NAME,"") or "—")
    byr   = str(row2.get(COL_BIRTH_YEAR,"") or "—")
    phone = str(row2.get(COL_PHONE,"") or "—")
    etime = str(row2.get(COL_EXAM_TIME,"") or "—")
    if len(etime) >= 5 and ":" in etime:
        etime = etime[:5]

    if phone not in ("—","N/A","nan",""):
        tel_digits = "".join(ch for ch in phone if ch.isdigit() or ch == "+")
        phone_show = f'<a href="tel:{tel_digits}" class="mrow-fact-phone">{phone}</a>'
    else:
        phone_show = f'<span style="color:#94a3b8">{phone}</span>'

    return (
        '<div class="mrow-name">👤 ' + name + '</div>'
        '<div class="mrow-facts">'
        '<span class="mrow-fact"><span class="fi" style="color:#64748b">🕐</span>' + etime + '</span>'
        '<span class="mrow-fact"><span class="fi" style="color:#2563eb">📅</span>' + byr + '</span>'
        '<span class="mrow-fact"><span class="fi" style="color:#e11d48">📞</span>' + phone_show + '</span>'
        '</div>'
    )


def patient_row_khoa_source_html(row2):
    """Cột 2: badge Khoa (tím nhạt) xếp trên badge Nguồn bệnh nhân (xanh
    dương nhạt, gọn) — khớp layout 4 cột."""
    khoa_show = khoa_badge_html(row2.get(COL_KHOA,""))
    src_pill  = source_pill_html(row2.get(COL_SOURCE,""))
    return f'<div class="khoa-src-stack">{khoa_show}{src_pill}</div>'


def render_upcoming_table(sub_df, empty_msg, dl_prefix, dl_key, page_state_key=None):
    """
    Vẽ bảng chi tiết bệnh nhân + nút tải CSV cho 1 nhóm (kb / khác) trong 1 ngày.
    Phân trang thông minh 10 bệnh nhân/trang nếu page_state_key được truyền vào
    (khoá session_state riêng cho từng bảng, ví dụ 'pg_kb_2026-07-16'), để mỗi
    bảng nhớ trang hiện tại độc lập, không ảnh hưởng các bảng khác trên trang.
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

    # ── DANH SÁCH BỆNH NHÂN: mỗi bệnh nhân gọn trong 1 HÀNG DUY NHẤT
    # (thông tin — trạng thái — ✏️ — 🗑️). Bấm ✏️/🗑️ chỉ MỞ RỘNG một ô nhỏ
    # (dạng dropdown) ngay bên dưới hàng đó để sửa/xóa, dùng rerun theo
    # phạm vi "fragment" (_smart_rerun) nên KHÔNG làm đóng popup danh sách
    # đang mở — không cần bấm lại "Xem chi tiết danh sách" nữa. ──
    for m_idx, m_row in page_df.iterrows():
        # m_idx là index gốc lấy từ Google Sheet (0-based, dòng dữ liệu đầu = index 0
        # = dòng 2 trên Sheet, vì dòng 1 là tiêu đề) → dòng thực tế = m_idx + 2.
        sheet_row = int(m_idx) + 2
        m_status = str(m_row.get(COL_STATUS, "") or "—")
        m_is_att = STATUS_ATTENDED.upper() in m_status.upper()
        name_facts_html = patient_row_name_facts_html(m_row)
        khoa_source_html = patient_row_khoa_source_html(m_row)
        m_name = str(m_row.get(COL_NAME, "") or "—")
        edit_open_key = f"inline_edit_open_{dl_key}_{sheet_row}"
        del_open_key = f"inline_del_open_{dl_key}_{sheet_row}"

        with st.container(key=f"mrow_{dl_key}_{sheet_row}_{'att' if m_is_att else 'nos'}"):
            # ── 4 vùng cách nhau bằng gạch dọc nhạt: Tên+SĐT | Khoa+Nguồn |
            # Trạng thái | Sửa/Xóa (2 icon liền nhau, không có gạch ở giữa
            # chúng — xem CSS div[class*="mrow_"] ... nth-of-type(2,3,4)). ──
            mc1, mc2, mc3, mc4, mc5 = st.columns([3.2, 2.6, 1.3, 0.55, 0.55])
            with mc1:
                st.markdown(name_facts_html, unsafe_allow_html=True)
            with mc2:
                st.markdown(khoa_source_html, unsafe_allow_html=True)
            with mc3:
                st.markdown(
                    f'<div class="mrow-status-wrap">{status_pill_html(m_status)}</div>',
                    unsafe_allow_html=True
                )
            with mc4:
                with st.container(key=f"editbtn_{dl_key}_{sheet_row}"):
                    if st.button("✏️", key=f"btn_edit_{dl_key}_{sheet_row}", use_container_width=True,
                                 help="Sửa trạng thái khám"):
                        st.session_state[edit_open_key] = not st.session_state.get(edit_open_key, False)
                        st.session_state[del_open_key] = False
                        _smart_rerun()
            with mc5:
                with st.container(key=f"delbtn_{dl_key}_{sheet_row}"):
                    if st.button("🗑️", key=f"btn_del_{dl_key}_{sheet_row}", use_container_width=True,
                                 help="Xóa bệnh nhân"):
                        st.session_state[del_open_key] = not st.session_state.get(del_open_key, False)
                        st.session_state[edit_open_key] = False
                        _smart_rerun()

            # Dropdown Sửa/Xóa — mở rộng NGAY DƯỚI hàng, bên trong cùng 1 thẻ,
            # không mở popup mới nên popup danh sách bên ngoài không bị đóng.
            if st.session_state.get(edit_open_key):
                render_inline_edit_form(sheet_row, m_name, m_status, open_key=edit_open_key)
            if st.session_state.get(del_open_key):
                render_inline_delete_form(sheet_row, m_name, open_key=del_open_key)

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


def _do_save_status(sheet_row, new_status, close_keys=()):
    """Ghi trạng thái mới lên Google Sheet, làm mới cache, rồi rerun.

    Dùng _smart_rerun() (rerun theo scope "fragment" khi đang ở trong 1
    popup/dialog) để KHÔNG làm đóng popup danh sách bệnh nhân đang mở.
    """
    if not creds_data:
        st.error("❌ Chưa có thông tin xác thực (credentials). Kiểm tra Streamlit Secrets.")
        return
    with st.spinner("Đang cập nhật trạng thái…"):
        ok, err = update_patient_status(creds_data, SHEET_ID, SHEET_NAME, sheet_row, new_status)
    if ok:
        st.session_state.metrics = None  # buộc tải lại dữ liệu mới nhất từ Sheet
        for k in close_keys:
            st.session_state[k] = False
        st.success("✅ Đã cập nhật trạng thái bệnh nhân!")
        _smart_rerun()
    else:
        st.error(f"❌ {err}")


def _do_delete_patient(sheet_row, close_keys=()):
    """Xóa bệnh nhân khỏi Google Sheet, làm mới cache, rồi rerun (giữ nguyên popup đang mở)."""
    if not creds_data:
        st.error("❌ Chưa có thông tin xác thực (credentials). Kiểm tra Streamlit Secrets.")
        return
    with st.spinner("Đang xóa bệnh nhân…"):
        ok, err = delete_patient_row(creds_data, SHEET_ID, SHEET_NAME, sheet_row)
    if ok:
        st.session_state.metrics = None
        for k in close_keys:
            st.session_state[k] = False
        st.success("✅ Đã xóa bệnh nhân khỏi danh sách!")
        _smart_rerun()
    else:
        st.error(f"❌ {err}")


def render_inline_edit_form(sheet_row, patient_name, current_status, open_key):
    """
    Dropdown "Sửa trạng thái" hiện ngay bên dưới hàng bệnh nhân (trong cùng
    popup danh sách đang mở) thay vì mở popup mới — tránh làm mất popup cha.
    """
    with st.container(key=f"editdrop_{open_key}"):
        st.markdown(f'<div class="edit-dlg-name">✏️ {patient_name}</div>', unsafe_allow_html=True)
        options = [STATUS_NOT_ATTENDED, STATUS_ATTENDED]
        default_idx = 1 if STATUS_ATTENDED.upper() in (current_status or "").upper() else 0
        new_status = st.radio(
            "Trạng thái khám", options=options, index=default_idx,
            key=f"inline_edit_radio_{open_key}", label_visibility="collapsed",
        )
        ic1, ic2 = st.columns(2)
        with ic1:
            if st.button("💾 Lưu", key=f"inline_edit_save_{open_key}", use_container_width=True, type="primary"):
                _do_save_status(sheet_row, new_status, close_keys=[open_key])
        with ic2:
            if st.button("❌ Đóng", key=f"inline_edit_close_{open_key}", use_container_width=True):
                st.session_state[open_key] = False
                _smart_rerun()


def render_inline_delete_form(sheet_row, patient_name, open_key):
    """Dropdown xác nhận xóa hiện ngay dưới hàng bệnh nhân, không mở popup mới."""
    with st.container(key=f"deldrop_{open_key}"):
        st.markdown(
            f'<div class="del-dlg-warn">⚠️ Xóa bệnh nhân <b>{patient_name}</b> khỏi danh sách?'
            '<span class="del-dlg-note">Hành động này không thể hoàn tác.</span></div>',
            unsafe_allow_html=True
        )
        ic1, ic2 = st.columns(2)
        with ic1:
            if st.button("🗑️ Xác nhận xóa", key=f"inline_del_confirm_{open_key}",
                         use_container_width=True, type="primary"):
                _do_delete_patient(sheet_row, close_keys=[open_key])
        with ic2:
            if st.button("❌ Hủy", key=f"inline_del_close_{open_key}", use_container_width=True):
                st.session_state[open_key] = False
                _smart_rerun()


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
    tab1, tab2, tab3, tab3b, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        "📊 Tổng Quan",
        "🔍 Tìm Theo Ngày",
        "📅 3 Ngày Tới",
        "📞 Nhắc Lịch BN Chưa Đến",
        "🏥 Nguồn Bệnh Nhân",
        "📈 Báo Cáo",
        "👤 Bệnh Nhân",
        "📥 Import Từ Minh Lộ",
        "✅ Đối Chiếu Tái Khám",
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

        def _render_day_detail(dday, udate_iso):
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
                )

        if HAS_DIALOG:
            @st.dialog("📋 Chi Tiết Lịch Khám Theo Ngày", width="large")
            def _open_day_dialog(date_iso, day_title_):
                st.markdown(f"#### 📅 {day_title_}")
                dday = _get_day_df(datetime.strptime(date_iso, "%Y-%m-%d").date())
                st.caption(f"Tổng cộng {len(dday)} bệnh nhân đăng ký khám ngày này.")
                _render_day_detail(dday, date_iso)

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

    # ═══════════════════════════
    # TAB 3B — NHẮC LỊCH
    # Danh sách bệnh nhân 3 ngày QUA chưa đến khám / bỏ khám
    # ═══════════════════════════
    with tab3b:
        st.markdown(
            '<div class="sh"><div class="sh-dot" style="background:#ef4444"></div>'
            '<span class="sh-txt">Danh Sách Bệnh Nhân Cần Nhắc Lịch</span></div>',
            unsafe_allow_html=True,
        )

        # ── Init session state cho tracking gọi điện ──────────────────────
        if "remind_call_status" not in st.session_state:
            st.session_state["remind_call_status"] = {}   # key=row_id: "called"|"no_answer"|""
        if "remind_call_note" not in st.session_state:
            st.session_state["remind_call_note"]   = {}   # key=row_id: str

        past_dates = [today - timedelta(days=i) for i in range(1, 4)]
        df_full    = m.get("df_full", df)

        if "_date" not in df_full.columns or not df_full["_date"].notna().any():
            st.info("Chưa có dữ liệu ngày khám.")
        else:
            mask_past = df_full["_date"].dt.date.isin(past_dates)
            mask_nos  = ~(
                df_full[COL_STATUS].astype(str).str.upper()
                .str.contains(STATUS_ATTENDED.upper(), na=False)
            )
            df_remind = df_full[mask_past & mask_nos].copy()

            # Ưu tiên bệnh nhân có SĐT lên đầu
            def _has_phone(p):
                return str(p).strip() not in ["", "0", "nan", "N/A", "—"]
            df_remind["_has_phone"] = df_remind[COL_PHONE].apply(_has_phone)
            df_remind = df_remind.sort_values(
                ["_date", "_has_phone"], ascending=[False, False]
            )

            total_remind = len(df_remind)
            n_has_phone  = int(df_remind["_has_phone"].sum())
            n_no_phone   = total_remind - n_has_phone

            # Tracking stats từ session_state
            cs = st.session_state["remind_call_status"]
            n_called    = sum(1 for v in cs.values() if v == "called")
            n_no_ans    = sum(1 for v in cs.values() if v == "no_answer")
            n_remaining = total_remind - n_called - n_no_ans
            pct_called  = round(n_called / total_remind * 100) if total_remind > 0 else 0

            # ── KPI row ─────────────────────────────────────────────────────
            st.markdown(f"""
            <div style="display:grid;grid-template-columns:repeat(5,1fr);gap:0.5rem;margin:0.5rem 0 0.9rem">
              <div class="kc kc-r" style="padding:0.7rem 0.8rem">
                <div class="kc-lbl">Tổng Cần Nhắc</div>
                <div class="kc-val" style="font-size:1.6rem;color:#dc2626">{total_remind}</div>
              </div>
              <div class="kc kc-g" style="padding:0.7rem 0.8rem">
                <div class="kc-lbl">Có SĐT</div>
                <div class="kc-val" style="font-size:1.6rem;color:#059669">{n_has_phone}</div>
              </div>
              <div class="kc" style="padding:0.7rem 0.8rem;background:#fafbfc;border:1px solid #e2e8f0;border-radius:12px">
                <div class="kc-lbl">Không SĐT</div>
                <div class="kc-val" style="font-size:1.6rem;color:#f59e0b">{n_no_phone}</div>
              </div>
              <div class="kc kc-g" style="padding:0.7rem 0.8rem">
                <div class="kc-lbl">Đã Gọi</div>
                <div class="kc-val" style="font-size:1.6rem;color:#2563eb">{n_called} <span style="font-size:0.9rem;color:#64748b">({pct_called}%)</span></div>
              </div>
              <div class="kc" style="padding:0.7rem 0.8rem;background:#fff7ed;border:1px solid #fed7aa;border-radius:12px">
                <div class="kc-lbl">Chưa Gọi</div>
                <div class="kc-val" style="font-size:1.6rem;color:#ea580c">{n_remaining}</div>
              </div>
            </div>
            """, unsafe_allow_html=True)

            # Progress bar gọi điện
            if total_remind > 0:
                st.progress(n_called / total_remind,
                            text=f"Tiến độ gọi nhắc: {n_called}/{total_remind} ({pct_called}%)")

            if total_remind == 0:
                st.success("✅ Tất cả bệnh nhân trong 3 ngày qua đã đến khám!")
            else:
                # ── Bộ lọc ──────────────────────────────────────────────────
                fc1, fc2, fc3 = st.columns([1.2, 1.2, 1])
                with fc1:
                    date_opts = ["Tất cả ngày"] + [
                        d.strftime("%d/%m/%Y") for d in sorted(past_dates, reverse=True)
                        if d in df_remind["_date"].dt.date.unique()
                    ]
                    sel_date = st.selectbox("📅 Ngày:", date_opts, key="remind_date_sel")
                with fc2:
                    khoa_list = sorted(
                        df_remind[COL_KHOA].dropna().astype(str).str.strip()
                        .replace("", pd.NA).dropna().unique().tolist()
                    ) if COL_KHOA in df_remind.columns else []
                    sel_khoa_r = st.selectbox("🏥 Khoa:", ["Tất cả khoa"] + khoa_list,
                                               key="remind_khoa_sel")
                with fc3:
                    call_filter_opts = ["Tất cả", "Chưa gọi", "Đã gọi", "Không bắt máy"]
                    sel_call = st.selectbox("📞 Trạng thái:", call_filter_opts,
                                            key="remind_call_filter")

                # Áp dụng filter
                df_filtered = df_remind.copy()
                if sel_date != "Tất cả ngày":
                    sd_obj = datetime.strptime(sel_date, "%d/%m/%Y").date()
                    df_filtered = df_filtered[df_filtered["_date"].dt.date == sd_obj]
                if sel_khoa_r != "Tất cả khoa" and COL_KHOA in df_filtered.columns:
                    df_filtered = df_filtered[
                        df_filtered[COL_KHOA].astype(str).str.strip() == sel_khoa_r
                    ]
                if sel_call != "Tất cả":
                    map_call = {"Chưa gọi": "", "Đã gọi": "called", "Không bắt máy": "no_answer"}
                    target_st = map_call[sel_call]
                    df_filtered = df_filtered[
                        df_filtered.index.map(lambda i: cs.get(str(i), "") == target_st)
                    ]

                st.caption(f"Hiển thị {len(df_filtered)} / {total_remind} bệnh nhân")

                # ── Nút Reset tracking (nhỏ, nằm bên phải) ─────────────────
                _, reset_col = st.columns([5, 1])
                with reset_col:
                    if st.button("🔄 Reset", key="remind_reset", help="Xoá toàn bộ trạng thái gọi"):
                        st.session_state["remind_call_status"] = {}
                        st.session_state["remind_call_note"]   = {}
                        st.rerun()

                # ── Hàm dựng 1 thẻ bệnh nhân + nút tracking gọi + ghi chú ──────
                def _render_remind_patient_card(row_idx, row):
                    rid       = str(row_idx)
                    r_name    = str(row.get(COL_NAME,   "") or "—")
                    r_phone   = str(row.get(COL_PHONE,  "") or "")
                    r_byr     = str(row.get(COL_BIRTH_YEAR, "") or "")
                    r_khoa    = str(row.get(COL_KHOA,   "") or "")
                    r_src     = str(row.get(COL_SOURCE, "") or "")
                    r_time    = str(row.get(COL_EXAM_TIME, "") or "")
                    if len(r_time) >= 5 and ":" in r_time:
                        r_time = r_time[:5]

                    call_st   = cs.get(rid, "")
                    has_tel   = row["_has_phone"]
                    tel_clean = "".join(c for c in r_phone if c.isdigit() or c == "+")

                    border_color = (
                        "#10b981" if call_st == "called"
                        else "#f59e0b" if call_st == "no_answer"
                        else "#ef4444"
                    )

                    if has_tel:
                        phone_html = (
                            f'<a href="tel:{tel_clean}" style="'
                            f'color:#fff;background:#059669;padding:0.2rem 0.6rem;'
                            f'border-radius:6px;font-size:0.7rem;font-weight:700;'
                            f'text-decoration:none;white-space:nowrap">📞 {r_phone}</a>'
                        )
                    else:
                        phone_html = '<span style="color:#94a3b8;font-size:0.7rem;white-space:nowrap">📞 Không có SĐT</span>'

                    khoa_html = (
                        f'<span class="pt-tag pt-tag-doc" style="font-size:0.67rem">{r_khoa[:28]}</span>'
                        if r_khoa.strip() and r_khoa not in ["nan","N/A","—"] else ""
                    )
                    src_html  = source_pill_html(r_src) or ""
                    time_html = (
                        f'<span class="mrow-meta-chip">🕐 {r_time}</span>'
                        if r_time and r_time not in ["—","N/A","nan"] else ""
                    )
                    byr_html  = (
                        f'<span class="mrow-meta-chip">🎂 {r_byr}</span>'
                        if r_byr and r_byr not in ["","—","N/A","nan"] else ""
                    )

                    call_badge = {
                        "called":    '<span style="background:#d1fae5;color:#065f46;font-size:0.62rem;font-weight:700;padding:0.15rem 0.45rem;border-radius:10px">✅ Đã gọi</span>',
                        "no_answer": '<span style="background:#fef3c7;color:#92400e;font-size:0.62rem;font-weight:700;padding:0.15rem 0.45rem;border-radius:10px">📵 Không bắt máy</span>',
                        "":          '<span style="background:#fee2e2;color:#991b1b;font-size:0.62rem;font-weight:700;padding:0.15rem 0.45rem;border-radius:10px">⏳ Chưa gọi</span>',
                    }.get(call_st, "")

                    note_val = st.session_state["remind_call_note"].get(rid, "")

                    st.markdown(f"""
                    <div style="background:#fff;border:1px solid #e2e8f0;
                        border-left:4px solid {border_color};border-radius:11px;
                        padding:0.55rem 0.8rem 0.45rem;margin-bottom:0.4rem;
                        box-shadow:0 1px 4px rgba(15,23,42,0.05)">
                      <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:0.3rem;margin-bottom:0.3rem">
                        <div style="font-size:0.85rem;font-weight:700;color:#0f172a">👤 {r_name}</div>
                        <div style="display:flex;gap:0.35rem;align-items:center;flex-wrap:wrap">
                          {call_badge} {phone_html}
                        </div>
                      </div>
                      <div style="display:flex;flex-wrap:wrap;gap:0.28rem;align-items:center">
                        {time_html}{byr_html}{khoa_html}{src_html}
                      </div>
                    </div>
                    """, unsafe_allow_html=True)

                    btn1, btn2, btn3 = st.columns([1, 1, 3])
                    with btn1:
                        if st.button(
                            "✅ Đã gọi" if call_st != "called" else "↩️ Bỏ đã gọi",
                            key=f"remind_called_{rid}",
                            use_container_width=True,
                        ):
                            cs[rid] = "" if call_st == "called" else "called"
                            st.rerun()
                    with btn2:
                        if st.button(
                            "📵 Không bắt" if call_st != "no_answer" else "↩️ Bỏ",
                            key=f"remind_noans_{rid}",
                            use_container_width=True,
                        ):
                            cs[rid] = "" if call_st == "no_answer" else "no_answer"
                            st.rerun()
                    with btn3:
                        note_new = st.text_input(
                            "Ghi chú:",
                            value=note_val,
                            placeholder="Nhập ghi chú sau khi gọi...",
                            key=f"remind_note_{rid}",
                            label_visibility="collapsed",
                        )
                        if note_new != note_val:
                            st.session_state["remind_call_note"][rid] = note_new

                # ── Hàm dựng danh sách (phân trang) cho 1 nhóm khoa ────────────
                def _render_remind_group(group_df, group_key):
                    if group_df.empty:
                        st.info("Không có bệnh nhân nào trong nhóm này.")
                        return
                    pg_key = f"pg_remind_{group_key}"
                    items = list(group_df.iterrows())
                    page_items, r_cur, r_total, r_start, r_end, r_tot = paginate_list(items, pg_key)
                    render_pagination_bar(pg_key, r_cur, r_total, r_start, r_end, r_tot,
                                          widget_key=f"{pg_key}_top")
                    for row_idx, row in page_items:
                        _render_remind_patient_card(row_idx, row)
                    render_pagination_bar(pg_key, r_cur, r_total, r_start, r_end, r_tot,
                                          widget_key=f"{pg_key}_bottom")

                HAS_DIALOG_R = hasattr(st, "dialog")

                def _render_remind_day_detail(day_df_, day_iso_):
                    kb_df, khac_df = split_khoa_groups(day_df_)
                    gtab1, gtab2 = st.tabs([
                        f"🩺 Khoa Khám Bệnh & Chưa Phân Khoa · {len(kb_df)}",
                        f"🏥 Khoa Điều Trị Nội Trú Khác · {len(khac_df)}",
                    ])
                    with gtab1:
                        _render_remind_group(kb_df, f"{day_iso_}_kb")
                    with gtab2:
                        _render_remind_group(khac_df, f"{day_iso_}_khac")

                if HAS_DIALOG_R:
                    @st.dialog("📞 Chi Tiết Nhắc Lịch Theo Ngày", width="large")
                    def _open_remind_dialog(day_iso_, day_label_):
                        st.markdown(f"#### 📅 {day_label_}")
                        dday_ = df_filtered[df_filtered["_date"].dt.date ==
                                             datetime.strptime(day_iso_, "%Y-%m-%d").date()]
                        st.caption(f"Tổng cộng {len(dday_)} bệnh nhân chưa đến khám ngày này.")
                        _render_remind_day_detail(dday_, day_iso_)

                # ── Danh sách từng ngày — CHỈ hiện tóm tắt (ngày + số lượng),
                # bấm "Xem chi tiết" mới mở popup danh sách đầy đủ. ──────────
                vn_days = ["Thứ Hai","Thứ Ba","Thứ Tư","Thứ Năm","Thứ Sáu","Thứ Bảy","Chủ Nhật"]

                for past_d in sorted(past_dates, reverse=True):
                    day_df = df_filtered[df_filtered["_date"].dt.date == past_d]
                    if day_df.empty:
                        continue

                    day_label    = f"{vn_days[past_d.weekday()]} · {past_d.strftime('%d/%m/%Y')}"
                    day_iso      = past_d.isoformat()
                    n_day        = len(day_df)
                    n_day_phone  = int(day_df["_has_phone"].sum())
                    n_day_called = sum(1 for i in day_df.index if cs.get(str(i),"") == "called")
                    kb_df, khac_df = split_khoa_groups(day_df)

                    st.markdown(
                        '<div class="upcoming-day">'
                        '<div class="upcoming-day-header" style="background:linear-gradient(135deg,#ef4444,#7f1d1d);">'
                        '<span class="upcoming-day-title">&#128197; ' + day_label + '</span>'
                        '<span class="upcoming-day-count">' + str(n_day) + ' chưa đến</span>'
                        '</div>',
                        unsafe_allow_html=True
                    )
                    st.markdown(
                        '<div class="upcoming-day-stats">'
                        '<div class="uds-item uds-kb">'
                        '<span class="uds-val">' + str(len(kb_df)) + '</span>'
                        '<span class="uds-lbl">🩺 Khoa Khám Bệnh &amp; chưa phân khoa</span>'
                        '</div>'
                        '<div class="uds-item uds-khac">'
                        '<span class="uds-val">' + str(len(khac_df)) + '</span>'
                        '<span class="uds-lbl">🏥 Khoa điều trị nội trú khác</span>'
                        '</div>'
                        '</div>',
                        unsafe_allow_html=True
                    )
                    st.markdown(
                        f'<div style="text-align:center;font-size:0.76rem;color:#64748b;margin:0.3rem 0 0.6rem">'
                        f'<span style="color:#059669;font-weight:600">{n_day_phone} có SĐT</span>'
                        f' &nbsp;·&nbsp; '
                        f'<span style="color:#2563eb;font-weight:600">{n_day_called} đã gọi</span>'
                        f'</div>',
                        unsafe_allow_html=True
                    )

                    if HAS_DIALOG_R:
                        st.markdown('<div class="upcoming-day-actions">', unsafe_allow_html=True)
                        if st.button("👁️ Xem chi tiết danh sách", key="btn_open_remind_" + day_iso,
                                     use_container_width=True):
                            _open_remind_dialog(day_iso, day_label)
                        st.markdown('</div>', unsafe_allow_html=True)
                    else:
                        st.markdown('<div class="upcoming-day-actions">', unsafe_allow_html=True)
                        with st.expander("👁️ Xem chi tiết danh sách bệnh nhân", expanded=False):
                            _render_remind_day_detail(day_df, day_iso)
                        st.markdown('</div>', unsafe_allow_html=True)

                    st.markdown('</div>', unsafe_allow_html=True)  # close upcoming-day

                # ── Nút tải CSV ─────────────────────────────────────────────
                remind_cols = [col for col in [
                    COL_NAME, COL_EXAM_DATE, COL_EXAM_TIME,
                    COL_PHONE, COL_KHOA, COL_SOURCE, COL_STATUS,
                ] if col in df_filtered.columns]
                export_df = df_filtered[remind_cols].copy()
                export_df["TRẠNG THÁI GỌI"] = [
                    {"called": "Đã gọi", "no_answer": "Không bắt máy"}.get(
                        cs.get(str(i), ""), "Chưa gọi"
                    ) for i in df_filtered.index
                ]
                export_df["GHI CHÚ"] = [
                    st.session_state["remind_call_note"].get(str(i), "")
                    for i in df_filtered.index
                ]
                csv_remind = export_df.to_csv(index=False, encoding="utf-8-sig")
                fname_remind = (
                    f"nhac_lich_{sel_date.replace('/','')}.csv"
                    if sel_date != "Tất cả ngày"
                    else f"nhac_lich_3ngay_{today.strftime('%Y%m%d')}.csv"
                )
                st.download_button(
                    label="⬇️ Xuất danh sách nhắc lịch (.csv)",
                    data=csv_remind.encode("utf-8-sig"),
                    file_name=fname_remind,
                    mime="text/csv",
                    key="dl_remind_csv",
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
                  <td class="num divider-l">{int(row['Đăng ký'])}</td>
                  <td class="num col-att divider-l">{int(row['Đến - Tái Khám'])}</td>
                  <td class="num col-att">{int(row['Đến - Vãng Lai'])}</td>
                  <td class="num col-abs divider-l">{int(row['Vắng - Tái Khám'])}</td>
                  <td class="num col-abs">{int(row['Vắng - Vãng Lai'])}</td>
                  <td class="num divider-l" style="color:#059669;font-weight:700">{int(row['Đã khám'])}</td>
                  <td class="num" style="color:#dc2626;font-weight:700">{int(row['Vắng / Chưa'])}</td>
                  <td class="{g_cls} divider-l">{row['Tỷ lệ đến (%)']}%</td>
                  <td class="{r_cls}">{row['Tỷ lệ vắng (%)']}%</td>
                </tr>"""
            st.markdown(f"""
            <div class="rtbl-wrap">
              <table class="rtbl"><thead>
                <tr>
                  <th rowspan="2">Kỳ</th><th rowspan="2" class="divider-l">Tổng</th>
                  <th colspan="2" class="grp-hdr grp-att divider-l">✅ Đến Khám</th>
                  <th colspan="2" class="grp-hdr grp-abs divider-l">⏳ Vắng</th>
                  <th rowspan="2" class="divider-l">Tổng Đến</th><th rowspan="2">Tổng Vắng</th>
                  <th rowspan="2" class="divider-l">% Đến</th><th rowspan="2">% Vắng</th>
                </tr>
                <tr>
                  <th class="sub-att divider-l">🏥 Tái Khám</th><th class="sub-att">🚶 Vãng Lai</th>
                  <th class="sub-abs divider-l">🏥 Tái Khám</th><th class="sub-abs">🚶 Vãng Lai</th>
                </tr>
              </thead><tbody>{rows_html}</tbody></table>
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
          Nguồn BN → "BỆNH NHÂN ĐIỀU TRỊ NỘI KHOA TÁI KHÁM" &nbsp;|&nbsp;
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

    # ════════════════════════════════════
    # TAB 8 — ĐỐI CHIẾU TÁI KHÁM
    # ════════════════════════════════════
    with tab8:
        st.markdown(
            '<div class="sh"><div class="sh-dot" style="background:#10b981"></div>'
            '<span class="sh-txt">Đối Chiếu Bệnh Nhân Đã Hẹn Với Thực Tế Đến Khám</span></div>',
            unsafe_allow_html=True
        )
        gtab_tk, gtab_vl = st.tabs(["🧮 Tái Khám — Đối Chiếu Tự Động", "📝 Vãng Lai — Check Thủ Công"])

        with gtab_tk:
            st.markdown("""
            <div style="background:#eff6ff;border:1px solid #bfdbfe;border-radius:12px;
                        padding:1rem 1.2rem;margin-bottom:1rem;font-size:0.83rem;color:#1e40af">
              <b>📋 Hướng dẫn:</b><br>
              1. Vào Minh Lộ → Báo cáo → <b>ĐK Khám Chữa Bệnh</b><br>
              2. Chọn <b>khoảng ngày rộng</b> (vd. cả tháng, hoặc từ ngày hẹn sớm nhất tới nay) → Export Excel<br>
              3. Upload file vào đây — hệ thống sẽ tự dò từng bệnh nhân <b>chưa khám</b> trên Sheet
              xem có xuất hiện trong log này không, <b>không cần đúng 1 ngày</b> (đến sớm/muộn vẫn bắt được)
            </div>
            """, unsafe_allow_html=True)

            visit_file = st.file_uploader(
                "Upload file Excel \"Báo cáo ĐK KCB\" từ Minh Lộ (.xlsx)",
                type=["xlsx"], key="visit_log_uploader",
                help="File log bệnh nhân THỰC TẾ đến khám, xuất theo khoảng ngày rộng"
            )

            if visit_file is not None:
                with st.spinner("Đang đọc file Excel…"):
                    visit_records, err_vl, warn_vl = parse_minh_lo_visit_log(visit_file)

                if err_vl:
                    st.error(f"❌ {err_vl}")
                elif not visit_records:
                    st.warning("⚠️ Không tìm thấy dữ liệu trong file. Kiểm tra đúng loại báo cáo \"ĐK KCB\".")
                else:
                    if warn_vl:
                        st.warning(warn_vl)
                    vdates = [d for d in (_parse_ddmmyyyy(v["NGÀY ĐK"]) for v in visit_records) if d]
                    vmin = min(vdates).strftime("%d/%m/%Y") if vdates else "?"
                    vmax = max(vdates).strftime("%d/%m/%Y") if vdates else "?"
                    st.success(f"✅ Đọc được **{len(visit_records)}** lượt khám thực tế, từ **{vmin}** đến **{vmax}**")

                    st.markdown(
                        '<div class="sh"><div class="sh-dot" style="background:#f59e0b"></div>'
                        '<span class="sh-txt">⚙️ Phạm Vi Đối Chiếu</span></div>',
                        unsafe_allow_html=True
                    )

                    range_start = today - timedelta(days=RECONCILE_LOOKBACK_DAYS)
                    range_end = today

                    st.markdown(f"""
                    <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px;
                                padding:0.9rem 1.1rem;margin-bottom:0.9rem;font-size:0.82rem;color:#334155">
                      📅 Danh sách gốc: bệnh nhân <b>từ khoa / tái khám</b> có <b>NGÀY KHÁM</b> từ
                      <b>{range_start.strftime('%d/%m/%Y')}</b> đến <b>{range_end.strftime('%d/%m/%Y')}</b>
                      ({RECONCILE_LOOKBACK_DAYS} ngày gần nhất) và đang <b>CHƯA KHÁM</b>.
                      <i>(Bệnh nhân vãng lai check ở tab bên cạnh — không cần thuật toán vì không có
                      lịch hẹn cố định để đối chiếu).</i><br>
                      🔍 Với mỗi bệnh nhân, chỉ chấp nhận lượt đến khám thực tế nằm trong cửa sổ
                      <b>[Ngày hẹn − {RECONCILE_WINDOW_BEFORE}, min(Ngày hẹn + {RECONCILE_WINDOW_AFTER}, Hôm nay)]</b>
                      — cửa sổ tự thu hẹp khi ngày hẹn gần hôm nay, để không nhầm sang một đợt khám
                      <i>khác</i> (vd. đợt khám cũ) của cùng bệnh nhân.
                    </div>
                    """, unsafe_allow_html=True)

                    df_full_src = m.get("df_full", df)
                    scope_df = df_full_src[
                        df_full_src["_date"].notna()
                        & (df_full_src["_date"].dt.date >= range_start)
                        & (df_full_src["_date"].dt.date <= range_end)
                        & (~df_full_src[COL_STATUS].astype(str).str.upper()
                             .str.contains(STATUS_ATTENDED.upper(), na=False))
                        & (df_full_src[COL_SOURCE].astype(str)
                             .str.contains("khoa|tái|nội trú|xuất viện|tai", case=False, na=False))
                    ]

                    if COL_STT not in scope_df.columns:
                        st.warning(
                            f"⚠️ Chưa thấy cột '{COL_STT}' trên Google Sheet — cập nhật trạng thái vẫn "
                            f"chạy được nhưng sẽ dùng vị trí dòng đã tải (kém an toàn hơn nếu Sheet bị "
                            f"sửa/sắp xếp lại trong lúc thao tác). Thêm cột '{COL_STT}' rồi bấm 🔄 Làm Mới "
                            f"để bật chế độ tra dòng theo khoá chính, an toàn hơn."
                        )

                    sheet_patients = []
                    for idx2, row in scope_df.iterrows():
                        exam_date = row["_date"].date() if pd.notna(row.get("_date")) else None
                        sheet_patients.append({
                            "sheet_row": int(idx2) + 2,
                            "stt": row.get(COL_STT, "") if COL_STT in row.index else "",
                            "name": row.get(COL_NAME, ""),
                            "phone": row.get(COL_PHONE, ""),
                            "cccd": row.get(COL_CCCD, ""),
                            "birth_year": row.get(COL_BIRTH_YEAR, ""),
                            "age": row.get(COL_AGE, "") if COL_AGE in row.index else "",
                            "exam_date": exam_date,
                            "source": row.get(COL_SOURCE, ""),
                            "status_now": row.get(COL_STATUS, ""),
                        })

                    st.caption(f"Sẽ kiểm tra **{len(sheet_patients)}** bệnh nhân từ khoa/tái khám chưa khám.")

                    if st.button("🔍 Bắt Đầu Đối Chiếu", type="primary", use_container_width=True,
                                 disabled=(len(sheet_patients) == 0)):
                        with st.spinner("Đang đối chiếu…"):
                            results = reconcile_attendance(sheet_patients, visit_records, today=today)
                        st.session_state["rec_results"] = results
                        st.session_state["rec_sheet_patients"] = sheet_patients

                    results = st.session_state.get("rec_results")
                    if results:
                        STATUS_LABEL = {
                            "attended_sure":   "✅ Đã đến khám",
                            "attended_unsure": "❓ Có thể đã đến (cần xác nhận)",
                            "not_attended":    "⏳ Chưa đến khám",
                        }
                        counts = Counter(r["status"] for r in results)
                        da_den = counts.get("attended_sure", 0)
                        chua_den = counts.get("not_attended", 0)
                        nghi_ngo = counts.get("attended_unsure", 0)
                        sot_list = [r for r in results if r["status"] == "not_attended" and r.get("near_miss")]
                        st.success(
                            f"📊 **Kết quả đối chiếu {len(results)} bệnh nhân**: "
                            f"**{da_den}** đã đến khám (chắc chắn) · "
                            f"**{chua_den}** chưa đến khám · "
                            f"**{nghi_ngo}** ca cần xác nhận tay"
                            + (f" · **{len(sot_list)}** ca nghi bị sót (tên giống nhưng ngoài cửa sổ ngày)."
                               if sot_list else ".")
                        )

                        st.markdown(f"""
                        <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:0.6rem;margin:0.9rem 0">
                          <div class="kc kc-g" style="padding:0.8rem 1rem">
                            <div class="kc-lbl">✅ Đã Đến Khám</div>
                            <div class="kc-val" style="font-size:1.5rem">{da_den}</div>
                          </div>
                          <div class="kc kc-b" style="padding:0.8rem 1rem">
                            <div class="kc-lbl">❓ Cần Xác Nhận</div>
                            <div class="kc-val" style="font-size:1.5rem;color:#1d4ed8">{nghi_ngo}</div>
                          </div>
                          <div class="kc kc-v" style="padding:0.8rem 1rem">
                            <div class="kc-lbl">⏳ Chưa Đến</div>
                            <div class="kc-val" style="font-size:1.5rem">{chua_den}</div>
                          </div>
                        </div>
                        """, unsafe_allow_html=True)

                        filter_opt = st.selectbox(
                            "Lọc theo kết quả",
                            ["Tất cả"] + list(STATUS_LABEL.values()),
                            key="rec_filter"
                        )
                        label_to_key = {v: k for k, v in STATUS_LABEL.items()}
                        shown = results if filter_opt == "Tất cả" else [
                            r for r in results if r["status"] == label_to_key[filter_opt]
                        ]

                        def _mk_result_row(r):
                            v = r.get("visit")
                            return {
                                "STT": r.get("stt", "") or "—",
                                "Họ tên": r["name"],
                                "SĐT": r.get("phone", "") or "—",
                                "Năm sinh": r.get("birth_year", "") or "—",
                                "Tuổi": r.get("age", "") or "—",
                                "Nguồn bệnh nhân": r.get("source", "") or "—",
                                "Ngày hẹn": r["exam_date"].strftime("%d/%m/%Y") if r["exam_date"] else "—",
                                "Kết quả": STATUS_LABEL[r["status"]],
                                "Khớp qua": {1: "Tên + SĐT", 2: "Tên + Năm sinh", 3: "Chỉ Tên"}.get(r.get("match_tier"), "—"),
                                "Ngày thực đến": v["NGÀY ĐK"] if v else "—",
                                "SĐT lúc khám": v.get("SỐ ĐIỆN THOẠI", "") if v else "—",
                                "Năm sinh lúc khám": v.get("NĂM SINH", "") if v else "—",
                                "Tuổi lúc khám": v.get("TUỔI", "") if v else "—",
                                "Độ tin cậy": f"{r['score']:.0f}%",
                                "Khoa thực khám": v.get("KHOA ĐK", "") if v else "—",
                            }

                        # Bảng CSV luôn xuất TOÀN BỘ kết quả (không chỉ trang đang xem)
                        table_rows = [_mk_result_row(r) for r in shown]

                        page_shown, pg_cur, pg_total, pg_start, pg_end, pg_tot = paginate_list(
                            shown, "pg_rec_results"
                        )
                        st.dataframe(pd.DataFrame([_mk_result_row(r) for r in page_shown]),
                                     use_container_width=True, hide_index=True,
                                     height=min(420, 70 + 35 * max(1, len(page_shown))))
                        render_pagination_bar("pg_rec_results", pg_cur, pg_total, pg_start, pg_end, pg_tot)

                        # ── Cập nhật hàng loạt trạng thái "ĐÃ KHÁM" — CHỈ cho các ca
                        # khớp CHẮC CHẮN ở Tầng 1 (Tên+SĐT) hoặc Tầng 2 (Tên+Năm sinh
                        # ±1). Ca "cần kiểm tra" (Tầng 3 — chỉ khớp mỗi tên) KHÔNG
                        # tự động ghi vào Sheet — phải xác nhận tay ở mục riêng bên dưới.
                        confirmable = [r for r in results if r["status"] == "attended_sure"]
                        to_update = [
                            r for r in confirmable
                            if STATUS_ATTENDED.upper() not in str(
                                next((p["status_now"] for p in sheet_patients if p["sheet_row"] == r["sheet_row"]), "")
                            ).upper()
                        ]

                        st.markdown(
                            '<div class="sh"><div class="sh-dot" style="background:#10b981"></div>'
                            '<span class="sh-txt">📋 Bước 2 — Xem Lại Danh Sách Trước Khi Cập Nhật</span></div>',
                            unsafe_allow_html=True
                        )
                        st.markdown(
                            f'<div class="pg-info" style="text-align:left;margin:0.5rem 0 0.8rem">'
                            f'Có <b>{len(to_update)}</b> bệnh nhân khớp CHẮC CHẮN (Tên+SĐT, hoặc Tên+Năm sinh '
                            f'lệch tối đa 1 năm) và đang ở trạng thái khác "Đã khám" trên Sheet. '
                            f'Nhóm "cần kiểm tra" ({nghi_ngo} ca, chỉ khớp được mỗi tên) <b>không</b> nằm trong '
                            f'danh sách này — xem và xử lý riêng ở mục bên dưới. Kiểm tra kỹ trước khi bấm cập nhật.</div>',
                            unsafe_allow_html=True
                        )

                        if to_update:
                            def _mk_confirm_row(r):
                                return {
                                    "STT": r.get("stt", "") or "—",
                                    "Họ tên": r["name"],
                                    "SĐT": r.get("phone", "") or "—",
                                    "Năm sinh": r.get("birth_year", "") or "—",
                                    "Tuổi": r.get("age", "") or "—",
                                    "Nguồn bệnh nhân": r.get("source", "") or "—",
                                    "Ngày hẹn": r["exam_date"].strftime("%d/%m/%Y") if r["exam_date"] else "—",
                                    "Ngày thực đến": r["visit"]["NGÀY ĐK"] if r.get("visit") else "—",
                                    "Khớp qua": {1: "Tên + SĐT", 2: "Tên + Năm sinh"}.get(r.get("match_tier"), "—"),
                                }
                            # Nút "Cập Nhật" bên dưới luôn áp dụng cho TOÀN BỘ to_update
                            # (không chỉ trang đang xem) — phân trang chỉ để XEM cho gọn.
                            page_upd, upd_cur, upd_total, upd_start, upd_end, upd_tot = paginate_list(
                                to_update, "pg_rec_confirm"
                            )
                            st.dataframe(pd.DataFrame([_mk_confirm_row(r) for r in page_upd]),
                                         use_container_width=True, hide_index=True,
                                         height=min(360, 70 + 35 * max(1, len(page_upd))))
                            render_pagination_bar("pg_rec_confirm", upd_cur, upd_total, upd_start, upd_end, upd_tot)

                            confirm_check = st.checkbox(
                                f"✅ Tôi đã xem kỹ danh sách {len(to_update)} bệnh nhân ở trên và xác nhận "
                                f"đúng người trước khi ghi vào Google Sheet",
                                key="rec_confirm_check"
                            )
                        else:
                            confirm_check = False
                            st.info("Không có bệnh nhân nào đủ điều kiện cập nhật tự động ở lần đối chiếu này.")

                        bc1, bc2 = st.columns([2, 1])
                        with bc1:
                            if st.button(f"✅ Cập Nhật \"Đã Khám\" Cho {len(to_update)} Bệnh Nhân",
                                         type="primary", use_container_width=True,
                                         disabled=(len(to_update) == 0 or not confirm_check)):
                                if not creds_data:
                                    st.error("❌ Chưa có credentials. Kiểm tra Streamlit Secrets.")
                                else:
                                    with st.spinner(f"Đang cập nhật {len(to_update)} dòng…"):
                                        n_ok, err_batch = update_patient_status_batch(
                                            creds_data, SHEET_ID, SHEET_NAME,
                                            [(r["sheet_row"], STATUS_ATTENDED, r.get("stt") or None) for r in to_update]
                                        )
                                    if err_batch:
                                        st.error(f"❌ {err_batch}")
                                    else:
                                        st.success(f"✅ Đã cập nhật cột TRẠNG THÁI thành \"Đã khám\" cho {n_ok} bệnh nhân!")
                                        st.session_state.metrics = None
                                        st.session_state.pop("rec_results", None)
                                        st.session_state.pop("rec_sheet_patients", None)
                                        st.session_state.pop("rec_confirm_check", None)
                                        st.balloons()
                        with bc2:
                            csv_rec = pd.DataFrame(table_rows).to_csv(index=False, encoding="utf-8-sig")
                            st.download_button(
                                "⬇️ Tải Báo Cáo (.csv)", data=csv_rec.encode("utf-8-sig"),
                                file_name=f"doi_chieu_taikham_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                                mime="text/csv", use_container_width=True,
                            )

                        # ── Bước 3 — Danh sách CẦN KIỂM TRA THỦ CÔNG (Tầng 3: chỉ
                        # khớp được mỗi cái tên, không có SĐT/năm sinh xác nhận).
                        # Cho phép SỬA TRỰC TIẾP SĐT/Năm sinh trên Sheet (để lần đối
                        # chiếu sau tự khớp đúng tầng 1/2), hoặc XÁC NHẬN THỦ CÔNG
                        # ngay tại đây nếu nhìn info đã đủ chắc là đúng người.
                        need_review = [r for r in results if r["status"] == "attended_unsure"]
                        if need_review:
                            st.markdown(
                                '<div class="sh"><div class="sh-dot" style="background:#f59e0b"></div>'
                                '<span class="sh-txt">🔎 Bước 3 — Cần Kiểm Tra Thủ Công (chỉ khớp tên)</span></div>',
                                unsafe_allow_html=True
                            )
                            st.caption(
                                f"{len(need_review)} bệnh nhân chỉ khớp được TÊN với 1 lượt khám trong log "
                                f"(không có SĐT hoặc năm sinh để xác nhận thêm) — xem kỹ thông tin 2 bên rồi "
                                f"chọn 1 trong 2 cách xử lý cho từng người."
                            )
                            page_rev, rev_cur, rev_total, rev_start, rev_end, rev_tot = paginate_list(
                                need_review, "pg_rec_review", page_size=5
                            )
                            render_pagination_bar("pg_rec_review", rev_cur, rev_total, rev_start, rev_end, rev_tot,
                                                   widget_key="pg_rec_review_top")
                            for r in page_rev:
                                v = r.get("visit") or {}
                                with st.expander(
                                    f"👤 STT {r.get('stt', '') or '—'} · {r['name']}  ·  "
                                    f"hẹn {r['exam_date'].strftime('%d/%m/%Y') if r['exam_date'] else '—'}"
                                    f"  ·  độ giống tên {r['score']:.0f}%"
                                ):
                                    cA, cB = st.columns(2)
                                    with cA:
                                        st.markdown("**Trên Google Sheet**")
                                        st.write(f"SĐT: {r.get('phone') or '—'}")
                                        st.write(f"Năm sinh: {r.get('birth_year') or '—'}")
                                        st.write(f"Tuổi (lúc hẹn): {r.get('age') or '—'}")
                                        st.write(f"Nguồn: {r.get('source') or '—'}")
                                    with cB:
                                        st.markdown("**Ứng viên khớp trong log Minh Lộ**")
                                        st.write(f"SĐT: {v.get('SỐ ĐIỆN THOẠI') or '—'}")
                                        st.write(f"Năm sinh: {v.get('NĂM SINH') or '—'}")
                                        st.write(f"Tuổi: {v.get('TUỔI') or '—'}")
                                        st.write(f"Ngày ĐK: {v.get('NGÀY ĐK') or '—'}  ·  Khoa: {v.get('KHOA ĐK') or '—'}")

                                    st.markdown("—")
                                    ec1, ec2 = st.columns(2)
                                    with ec1:
                                        new_phone = st.text_input(
                                            "Sửa SĐT trên Sheet", value=str(v.get("SỐ ĐIỆN THOẠI") or r.get("phone") or ""),
                                            key=f"edit_phone_{r['sheet_row']}"
                                        )
                                    with ec2:
                                        new_birth = st.text_input(
                                            "Sửa Năm sinh trên Sheet", value=str(v.get("NĂM SINH") or r.get("birth_year") or ""),
                                            key=f"edit_birth_{r['sheet_row']}"
                                        )
                                    bA, bB = st.columns(2)
                                    with bA:
                                        if st.button("💾 Lưu SĐT/Năm sinh vào Sheet", key=f"save_info_{r['sheet_row']}",
                                                     use_container_width=True):
                                            if not creds_data:
                                                st.error("❌ Chưa có credentials.")
                                            else:
                                                ok, err_f = update_patient_fields(
                                                    creds_data, SHEET_ID, SHEET_NAME, r["sheet_row"],
                                                    {COL_PHONE: new_phone, COL_BIRTH_YEAR: new_birth},
                                                    stt=r.get("stt") or None
                                                )
                                                if ok:
                                                    st.success("✅ Đã lưu — lần đối chiếu sau sẽ tự khớp đúng hơn.")
                                                else:
                                                    st.error(f"❌ {err_f}")
                                    with bB:
                                        if st.button("✅ Xác nhận đây đúng — đánh dấu Đã khám", key=f"confirm_att_{r['sheet_row']}",
                                                     use_container_width=True, type="primary"):
                                            if not creds_data:
                                                st.error("❌ Chưa có credentials.")
                                            else:
                                                n_ok2, err2 = update_patient_status_batch(
                                                    creds_data, SHEET_ID, SHEET_NAME,
                                                    [(r["sheet_row"], STATUS_ATTENDED, r.get("stt") or None)]
                                                )
                                                if err2:
                                                    st.error(f"❌ {err2}")
                                                else:
                                                    st.success("✅ Đã đánh dấu Đã khám cho bệnh nhân này.")
                                                    st.session_state.metrics = None
                            render_pagination_bar("pg_rec_review", rev_cur, rev_total, rev_start, rev_end, rev_tot,
                                                   widget_key="pg_rec_review_bottom")

                        # ── Bước 4 — NGHI BỊ SÓT: bệnh nhân bị kết luận "chưa khám"
                        # nhưng tìm thấy 1 lượt khám TÊN GIỐNG ở đâu đó trong file,
                        # chỉ là NGOÀI cửa sổ ngày cho phép (đến quá sớm/quá muộn so
                        # với hẹn, hoặc NGÀY KHÁM trên Sheet ghi sai) — rất đáng ngờ
                        # là bị sót do cửa sổ quá hẹp chứ không phải thật sự chưa đến.
                        if sot_list:
                            st.markdown(
                                '<div class="sh"><div class="sh-dot" style="background:#ef4444"></div>'
                                '<span class="sh-txt">⚠️ Bước 4 — Nghi Bị Sót (tên giống, ngoài cửa sổ ngày)</span></div>',
                                unsafe_allow_html=True
                            )
                            st.caption(
                                f"{len(sot_list)} bệnh nhân bị đánh dấu \"chưa khám\" nhưng có 1 lượt khám TÊN GIỐNG "
                                f"trong file — chỉ là ngày thực đến NẰM NGOÀI cửa sổ cho phép "
                                f"[hẹn − {RECONCILE_WINDOW_BEFORE}, hẹn + {RECONCILE_WINDOW_AFTER}]. Xem kỹ rồi xác nhận "
                                f"nếu đúng là cùng 1 người."
                            )
                            page_sot, sot_cur, sot_total, sot_start, sot_end, sot_tot = paginate_list(
                                sot_list, "pg_rec_sot", page_size=5
                            )
                            render_pagination_bar("pg_rec_sot", sot_cur, sot_total, sot_start, sot_end, sot_tot,
                                                   widget_key="pg_rec_sot_top")
                            for r in page_sot:
                                nm = r["near_miss"]
                                v = nm["visit"]
                                vd_str = nm["visit_date"].strftime("%d/%m/%Y") if nm["visit_date"] else (v.get("NGÀY ĐK") or "—")
                                hen_str = r["exam_date"].strftime("%d/%m/%Y") if r["exam_date"] else "—"
                                with st.expander(
                                    f"⚠️ STT {r.get('stt', '') or '—'} · {r['name']}  ·  "
                                    f"hẹn {hen_str}  ·  thực đến {vd_str} (ngoài cửa sổ)  ·  "
                                    f"độ giống tên {nm['score']:.0f}%"
                                ):
                                    sA, sB = st.columns(2)
                                    with sA:
                                        st.markdown("**Trên Google Sheet (đã hẹn)**")
                                        st.write(f"👤 {r['name']}")
                                        st.write(f"📅 Ngày hẹn: {hen_str}")
                                        st.write(f"📞 SĐT: {r.get('phone') or '—'}")
                                        st.write(f"🎂 Năm sinh: {r.get('birth_year') or '—'}")
                                    with sB:
                                        st.markdown("**Lượt khám tìm thấy trong file (ngoài cửa sổ)**")
                                        st.write(f"👤 {v.get('HỌ TÊN', '')}")
                                        st.write(f"📅 Ngày thực đến: {vd_str}")
                                        st.write(f"📞 SĐT lúc khám: {v.get('SỐ ĐIỆN THOẠI') or '—'}")
                                        st.write(f"🎂 Năm sinh lúc khám: {v.get('NĂM SINH') or '—'}")
                                        st.write(f"🏥 Khoa thực khám: {v.get('KHOA ĐK') or '—'}")
                                    if st.button(
                                        "✅ Đúng là người này — đánh dấu Đã khám",
                                        key=f"confirm_sot_{r['sheet_row']}", use_container_width=True, type="primary"
                                    ):
                                        if not creds_data:
                                            st.error("❌ Chưa có credentials.")
                                        else:
                                            n_ok3, err3 = update_patient_status_batch(
                                                creds_data, SHEET_ID, SHEET_NAME,
                                                [(r["sheet_row"], STATUS_ATTENDED, r.get("stt") or None)]
                                            )
                                            if err3:
                                                st.error(f"❌ {err3}")
                                            else:
                                                st.success("✅ Đã đánh dấu Đã khám cho bệnh nhân này.")
                                                st.session_state.metrics = None
                            render_pagination_bar("pg_rec_sot", sot_cur, sot_total, sot_start, sot_end, sot_tot,
                                                   widget_key="pg_rec_sot_bottom")

        with gtab_vl:
            st.markdown("""
            <div style="background:#fff7ed;border:1px solid #fed7aa;border-radius:12px;
                        padding:1rem 1.2rem;margin-bottom:1rem;font-size:0.83rem;color:#9a3412">
              <b>📝 Check thủ công — Bệnh Nhân Vãng Lai</b><br>
              Bệnh nhân vãng lai không có lịch hẹn cố định nên không đối chiếu bằng thuật toán được —
              chọn khoảng ngày bên dưới rồi bấm <b>✏️</b> ở từng người để đánh dấu
              <b>Đã khám</b> / <b>Chưa khám</b> trực tiếp.
            </div>
            """, unsafe_allow_html=True)

            vl_c1, vl_c2, vl_c3 = st.columns([1, 1, 1.3])
            with vl_c1:
                vl_from = st.date_input(
                    "Từ ngày", value=today - timedelta(days=RECONCILE_LOOKBACK_DAYS), key="vl_checkin_from"
                )
            with vl_c2:
                vl_to = st.date_input("Đến ngày", value=today, key="vl_checkin_to")
            with vl_c3:
                vl_only_pending = st.checkbox(
                    "Chỉ hiện bệnh nhân CHƯA khám", value=True, key="vl_checkin_only_pending"
                )

            if vl_from > vl_to:
                st.warning('⚠️ "Từ ngày" đang sau "Đến ngày" — đổi lại giúp tao nhé.')
            else:
                vl_full_src = m.get("df_full", df)
                vl_mask = (
                    vl_full_src["_date"].notna()
                    & (vl_full_src["_date"].dt.date >= vl_from)
                    & (vl_full_src["_date"].dt.date <= vl_to)
                    & (vl_full_src[COL_SOURCE].astype(str)
                         .str.contains("vãng lai|vang lai|ngoài|ngoai", case=False, na=False))
                )
                if vl_only_pending:
                    vl_mask &= (~vl_full_src[COL_STATUS].astype(str).str.upper()
                                  .str.contains(STATUS_ATTENDED.upper(), na=False))
                vl_scope_df = vl_full_src[vl_mask]

                st.caption(
                    f"**{len(vl_scope_df)}** bệnh nhân vãng lai trong khoảng ngày đã chọn"
                    + (" (đang CHƯA khám)." if vl_only_pending else ".")
                )

                render_upcoming_table(
                    vl_scope_df,
                    empty_msg="Không có bệnh nhân vãng lai nào trong khoảng ngày này.",
                    dl_prefix="vang_lai_checkin",
                    dl_key="dl_vl_checkin",
                    page_state_key="pg_vl_checkin",
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
