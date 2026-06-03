"""
Bảng Điều Khiển Đăng Ký Khám Online — BVĐK Tâm Đức Cầu Quan
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
COL_SYMPTOM   = "1. TRIỆU CHỨNG CHÍNH"
COL_SPECIALTY = "CHUYÊN KHOA MONG MUỐN KHÁM"
COL_DOCTOR    = "BÁC SĨ MONG MUỐN ( nếu có)"
COL_EXAM_TIME = "GIỜ KHÁM DỰ KIẾN"

STATUS_ATTENDED = "BỆNH NHÂN ĐÃ KHÁM"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

# ── PAGE CONFIG ───────────────────────────────
st.set_page_config(
    page_title="Dashboard Khám Bệnh · Tâm Đức",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── DESIGN SYSTEM ─────────────────────────────
# Deep navy + teal accent — clean medical palette
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@500&display=swap');

*, html, body, [class*="css"] {
    font-family: 'Inter', sans-serif !important;
    box-sizing: border-box;
}

/* ── Background ── */
.stApp { background: #f8fafc; }
.main .block-container {
    padding: 2rem 2.5rem 3rem 2.5rem;
    max-width: 1400px;
}

/* ── Sidebar ── */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0d1b2a 0%, #1a2e44 100%) !important;
    border-right: 1px solid #1e3a5f;
}
section[data-testid="stSidebar"] * { color: #94b8d4 !important; }
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] strong { color: #e2edf7 !important; }
section[data-testid="stSidebar"] .stMarkdown p { color: #7fa8c8 !important; font-size: 0.85rem; }
section[data-testid="stSidebar"] hr { border-color: #1e3a5f !important; margin: 0.8rem 0; }

/* file uploader label */
section[data-testid="stSidebar"] label { color: #94b8d4 !important; }

/* ── Button ── */
.stButton > button {
    background: linear-gradient(135deg, #0ea5e9, #0284c7) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    font-size: 0.88rem !important;
    padding: 0.6rem 1.2rem !important;
    width: 100% !important;
    letter-spacing: 0.02em;
    box-shadow: 0 4px 14px rgba(14,165,233,0.35) !important;
    transition: all 0.2s ease !important;
}
.stButton > button:hover {
    background: linear-gradient(135deg, #38bdf8, #0ea5e9) !important;
    box-shadow: 0 6px 20px rgba(14,165,233,0.45) !important;
    transform: translateY(-1px);
}

/* ── Header banner ── */
.dash-header {
    background: linear-gradient(135deg, #0d1b2a 0%, #0f3460 60%, #164e74 100%);
    border-radius: 18px;
    padding: 1.8rem 2.2rem;
    margin-bottom: 1.8rem;
    display: flex;
    align-items: center;
    gap: 1.2rem;
    box-shadow: 0 8px 32px rgba(13,27,42,0.18);
}
.dash-header-icon {
    font-size: 2.6rem;
    background: rgba(255,255,255,0.08);
    border-radius: 14px;
    padding: 0.5rem 0.7rem;
    line-height: 1;
}
.dash-header-title {
    font-size: 1.55rem;
    font-weight: 700;
    color: #f0f9ff;
    line-height: 1.2;
    margin: 0;
}
.dash-header-sub {
    font-size: 0.82rem;
    color: #7dd3fc;
    margin-top: 0.3rem;
    letter-spacing: 0.03em;
}
.dash-header-badge {
    margin-left: auto;
    background: rgba(14,165,233,0.18);
    border: 1px solid rgba(14,165,233,0.35);
    border-radius: 20px;
    padding: 0.35rem 1rem;
    color: #7dd3fc;
    font-size: 0.78rem;
    font-weight: 500;
    white-space: nowrap;
}

/* ── KPI Cards ── */
.kpi-grid { display: flex; gap: 1rem; margin-bottom: 1.5rem; }
.kpi-card {
    flex: 1;
    background: #fff;
    border-radius: 16px;
    padding: 1.4rem 1.5rem 1.2rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06), 0 8px 24px rgba(0,0,0,0.04);
    position: relative;
    overflow: hidden;
    border: 1px solid #f1f5f9;
    transition: transform 0.15s ease, box-shadow 0.15s ease;
}
.kpi-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(0,0,0,0.08), 0 12px 32px rgba(0,0,0,0.06);
}
.kpi-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
    border-radius: 16px 16px 0 0;
}
.kpi-blue::before   { background: linear-gradient(90deg, #0ea5e9, #38bdf8); }
.kpi-green::before  { background: linear-gradient(90deg, #10b981, #34d399); }
.kpi-rose::before   { background: linear-gradient(90deg, #f43f5e, #fb7185); }
.kpi-violet::before { background: linear-gradient(90deg, #8b5cf6, #a78bfa); }

.kpi-icon {
    font-size: 1.6rem;
    margin-bottom: 0.6rem;
    display: block;
}
.kpi-label {
    font-size: 0.7rem;
    font-weight: 600;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: #94a3b8;
    margin-bottom: 0.4rem;
}
.kpi-value {
    font-size: 2.4rem;
    font-weight: 700;
    color: #0f172a;
    line-height: 1;
    font-family: 'JetBrains Mono', monospace !important;
}
.kpi-sub {
    font-size: 0.76rem;
    color: #94a3b8;
    margin-top: 0.4rem;
    font-weight: 500;
}
.kpi-trend {
    position: absolute;
    top: 1.2rem; right: 1.2rem;
    font-size: 1.5rem;
    opacity: 0.12;
}

/* ── Section header ── */
.sec-header {
    display: flex;
    align-items: center;
    gap: 0.6rem;
    margin: 1.6rem 0 0.9rem;
    padding-bottom: 0.6rem;
    border-bottom: 1.5px solid #e2e8f0;
}
.sec-header-icon {
    width: 28px; height: 28px;
    border-radius: 8px;
    display: flex; align-items: center; justify-content: center;
    font-size: 0.95rem;
}
.sec-header-blue   { background: #eff6ff; }
.sec-header-green  { background: #f0fdf4; }
.sec-header-violet { background: #f5f3ff; }
.sec-header-amber  { background: #fffbeb; }
.sec-header-text {
    font-size: 0.92rem;
    font-weight: 700;
    color: #1e293b;
    letter-spacing: 0.01em;
}

/* ── Chart card ── */
.chart-card {
    background: #fff;
    border-radius: 16px;
    padding: 1.4rem 1.4rem 0.8rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06), 0 8px 24px rgba(0,0,0,0.04);
    border: 1px solid #f1f5f9;
}

/* ── Empty state ── */
.empty-state {
    text-align: center;
    padding: 5rem 2rem;
    background: #fff;
    border-radius: 20px;
    border: 2px dashed #e2e8f0;
    margin-top: 1.5rem;
}
.empty-icon { font-size: 3.5rem; margin-bottom: 1rem; }
.empty-title { font-size: 1.2rem; font-weight: 700; color: #1e293b; margin-bottom: 0.5rem; }
.empty-desc { font-size: 0.88rem; color: #94a3b8; max-width: 380px; margin: 0 auto; line-height: 1.6; }

/* ── Sidebar logo area ── */
.sidebar-logo {
    background: rgba(14,165,233,0.08);
    border: 1px solid rgba(14,165,233,0.15);
    border-radius: 12px;
    padding: 1rem 1.2rem;
    margin-bottom: 1rem;
    text-align: center;
}
.sidebar-logo-title {
    font-size: 0.95rem !important;
    font-weight: 700 !important;
    color: #e2edf7 !important;
    line-height: 1.3;
}
.sidebar-logo-sub {
    font-size: 0.72rem !important;
    color: #5b92b5 !important;
    margin-top: 0.2rem;
}

/* Streamlit overrides */
div[data-testid="stExpander"] { background: #fff; border-radius: 12px; border: 1px solid #f1f5f9; }
.stDataFrame { border-radius: 12px; overflow: hidden; }
div[data-testid="stAlert"] { border-radius: 12px; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


# ── PALETTE for charts ─────────────────────────
C_BLUE   = "#0ea5e9"
C_TEAL   = "#14b8a6"
C_GREEN  = "#10b981"
C_ROSE   = "#f43f5e"
C_VIOLET = "#8b5cf6"
C_AMBER  = "#f59e0b"
C_SLATE  = "#94a3b8"

CHART_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter, sans-serif", color="#475569"),
    hoverlabel=dict(
        bgcolor="#0f172a",
        font_color="#f1f5f9",
        font_size=12,
        bordercolor="#1e3a5f",
    ),
)


# ── FUNCTIONS ─────────────────────────────────

def authenticate_gspread(credentials_source) -> gspread.Client:
    if isinstance(credentials_source, str):
        if not os.path.exists(credentials_source):
            raise FileNotFoundError(f"Không tìm thấy file: {credentials_source}")
        creds = Credentials.from_service_account_file(credentials_source, scopes=SCOPES)
    elif isinstance(credentials_source, dict):
        creds = Credentials.from_service_account_info(credentials_source, scopes=SCOPES)
    else:
        raise ValueError("credentials_source phải là str hoặc dict.")
    return gspread.authorize(creds)


def fetch_data(client: gspread.Client, sheet_id: str, sheet_name: str) -> pd.DataFrame:
    spreadsheet = client.open_by_key(sheet_id)
    worksheet   = spreadsheet.worksheet(sheet_name)
    all_values  = worksheet.get_all_values()
    if not all_values or len(all_values) < 2:
        raise ValueError("Sheet trống hoặc không có dữ liệu.")

    headers = all_values[0]
    rows    = all_values[1:]
    seen    = {}
    clean   = []
    for i, h in enumerate(headers):
        h = h.strip()
        if h == "":
            h = f"_col_{i}"
        elif h in seen:
            seen[h] += 1
            h = f"{h}_{seen[h]}"
        else:
            seen[h] = 0
        clean.append(h)

    df = pd.DataFrame(rows, columns=clean)
    df = df.loc[:, ~df.columns.str.startswith("_col_")]
    df = df.replace("", pd.NA).dropna(how="all").fillna("")
    return df


def process_data(df: pd.DataFrame) -> dict:
    if COL_STATUS not in df.columns:
        raise KeyError(f"Không tìm thấy cột '{COL_STATUS}'. Cột hiện có: {list(df.columns)}")

    df = df.copy()
    df[COL_STATUS] = df[COL_STATUS].astype(str).str.strip()
    df = df[~df[COL_STATUS].isin(["", "nan", "N/A", "\u200b"])]
    total = len(df)
    if total == 0:
        return _empty(df)

    attended_count = (df[COL_STATUS].str.upper() == STATUS_ATTENDED.upper()).sum()
    noshow_count   = total - attended_count

    # Specialty
    spec = None
    if COL_SPECIALTY in df.columns:
        s = df[df[COL_SPECIALTY].astype(str).str.strip() != ""][COL_SPECIALTY]
        if not s.empty:
            spec = s.value_counts().head(8).reset_index()
            spec.columns = ["Chuyên khoa", "Số lượng"]

    # Gender
    gen = None
    if COL_GENDER in df.columns:
        g = df[COL_GENDER].astype(str).str.strip()
        g = g[g.str.upper().isin(["NAM", "NỮ", "NU"])]
        if not g.empty:
            gen = g.value_counts().reset_index()
            gen.columns = ["Giới tính", "Số lượng"]

    # Daily
    daily = None
    if COL_EXAM_DATE in df.columns:
        ds = df[COL_EXAM_DATE].astype(str).str.strip()
        ds = ds[ds.str.match(r'\d{2}/\d{2}/\d{4}')]
        if not ds.empty:
            daily = ds.value_counts().sort_index().reset_index()
            daily.columns = ["Ngày khám", "Lịch hẹn"]

    # Status table
    status_tbl = df[COL_STATUS].value_counts().reset_index()
    status_tbl.columns = ["Trạng thái", "Số lượng"]

    return dict(
        total=total,
        attended_count=int(attended_count),
        noshow_count=int(noshow_count),
        attended_pct=round(attended_count / total * 100, 1),
        noshow_pct=round(noshow_count / total * 100, 1),
        spec=spec, gen=gen, daily=daily,
        status_tbl=status_tbl, df=df,
    )


def _empty(df):
    return dict(total=0, attended_count=0, noshow_count=0,
                attended_pct=0.0, noshow_pct=0.0,
                spec=None, gen=None, daily=None, status_tbl=None, df=df)


# ── CHARTS ────────────────────────────────────

def chart_donut(m):
    labels = ["Đã khám", "Chưa / Vắng"]
    values = [m["attended_count"], m["noshow_count"]]
    colors = [C_GREEN, C_ROSE]
    fig = go.Figure(go.Pie(
        labels=labels, values=values, hole=0.65,
        marker=dict(colors=colors, line=dict(color="#fff", width=3)),
        textinfo="label+percent",
        textfont=dict(size=12.5, family="Inter"),
        hovertemplate="<b>%{label}</b><br>%{value} bệnh nhân · %{percent}<extra></extra>",
        pull=[0.03, 0.03],
        direction="clockwise",
    ))
    fig.update_layout(
        **CHART_LAYOUT,
        height=320,
        showlegend=False,
        annotations=[dict(
            text=f"<b style='font-size:26px'>{m['total']}</b><br>"
                 f"<span style='font-size:11px;color:#94a3b8'>bệnh nhân</span>",
            x=0.5, y=0.5, showarrow=False,
            font=dict(family="Inter", color="#0f172a"),
        )],
        margin=dict(t=16, b=16, l=40, r=40),
    )
    return fig


def chart_gender(m):
    df = m["gen"]
    if df is None or df.empty:
        return None
    color_map = {"NAM": C_BLUE, "NỮ": "#f472b6", "NU": "#f472b6"}
    colors = [color_map.get(g.upper(), C_SLATE) for g in df["Giới tính"]]
    fig = go.Figure(go.Bar(
        x=df["Giới tính"], y=df["Số lượng"],
        marker=dict(color=colors, line=dict(color="#fff", width=2)),
        text=df["Số lượng"], textposition="outside",
        textfont=dict(size=14, family="JetBrains Mono", color="#0f172a"),
        hovertemplate="<b>%{x}</b> — %{y} người<extra></extra>",
        width=0.45,
    ))
    fig.update_layout(
        **CHART_LAYOUT,
        height=260,
        xaxis=dict(showgrid=False, tickfont=dict(size=13, color="#475569")),
        yaxis=dict(showgrid=True, gridcolor="#f1f5f9", zeroline=False,
                   tickfont=dict(size=11, color="#94a3b8")),
    )
    return fig


def chart_daily(m):
    df = m["daily"]
    if df is None or df.empty:
        return None
    # gradient blue bars
    n = len(df)
    colors = [f"rgba(14,165,233,{0.45 + 0.55 * i / max(n-1,1):.2f})" for i in range(n)]
    fig = go.Figure(go.Bar(
        x=df["Ngày khám"], y=df["Lịch hẹn"],
        marker=dict(color=colors, line=dict(color="#fff", width=1.5),
                    cornerradius=6),
        text=df["Lịch hẹn"], textposition="outside",
        textfont=dict(size=11, family="JetBrains Mono", color="#0f172a"),
        hovertemplate="<b>%{x}</b><br>%{y} lịch hẹn<extra></extra>",
    ))
    fig.update_layout(
        **CHART_LAYOUT,
        height=280,
        xaxis=dict(tickangle=-30, tickfont=dict(size=10, color="#64748b"),
                   showgrid=False),
        yaxis=dict(gridcolor="#f8fafc", zeroline=False,
                   tickfont=dict(size=10, color="#94a3b8")),
        bargap=0.28,
    )
    return fig


def chart_specialty(m):
    df = m["spec"]
    if df is None or df.empty:
        return None
    df = df.copy().sort_values("Số lượng")
    # truncate long names
    df["label"] = df["Chuyên khoa"].str.slice(0, 38).where(
        df["Chuyên khoa"].str.len() <= 38,
        df["Chuyên khoa"].str.slice(0, 38) + "…"
    )
    n = len(df)
    palette = [C_TEAL, C_BLUE, C_VIOLET, C_AMBER, C_GREEN,
               "#06b6d4", "#6366f1", "#f97316"]
    colors = palette[:n]

    fig = go.Figure(go.Bar(
        y=df["label"], x=df["Số lượng"],
        orientation="h",
        marker=dict(color=colors, line=dict(color="#fff", width=1.5)),
        text=df["Số lượng"], textposition="outside",
        textfont=dict(size=12, family="JetBrains Mono", color="#0f172a"),
        hovertemplate="<b>%{y}</b><br>%{x} bệnh nhân<extra></extra>",
    ))
    fig.update_layout(
        **CHART_LAYOUT,
        height=max(280, n * 48),
        xaxis=dict(showgrid=True, gridcolor="#f1f5f9", zeroline=False,
                   tickfont=dict(size=10, color="#94a3b8")),
        yaxis=dict(tickfont=dict(size=10.5, color="#475569"), showgrid=False),
        margin=dict(t=16, b=16, l=16, r=56),
    )
    return fig


# ── SIDEBAR ───────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div class="sidebar-logo">
      <div style="font-size:1.8rem;margin-bottom:0.3rem">🏥</div>
      <div class="sidebar-logo-title">BVĐK Tâm Đức Cầu Quan</div>
      <div class="sidebar-logo-sub">Hệ thống Dashboard Online</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("**🔑 Xác thực Google**")
    cred_mode = st.radio("", ["Upload file JSON", "Dán nội dung JSON"],
                         label_visibility="collapsed")

    creds_data = None
    if cred_mode == "Upload file JSON":
        ufile = st.file_uploader("Chọn Service Account JSON", type="json")
        if ufile:
            try:
                creds_data = json.loads(ufile.read().decode("utf-8-sig").strip())
                st.success("✓ File hợp lệ")
            except Exception as e:
                st.error(f"✗ Lỗi đọc file: {e}")
    else:
        jtext = st.text_area("Dán JSON vào đây", height=110,
                              placeholder='{"type": "service_account", ...}')
        if jtext.strip():
            try:
                creds_data = json.loads(jtext)
                st.success("✓ JSON hợp lệ")
            except:
                st.error("✗ JSON không hợp lệ")

    st.markdown("---")
    fetch_btn = st.button("🔄  Tải & Cập nhật dữ liệu")

    st.markdown("---")
    st.markdown("**📋 Sheet đang dùng**")
    st.code(f"ID: ...{SHEET_ID[-12:]}\nTab: {SHEET_NAME[:24]}", language=None)

    with st.expander("📖 Hướng dẫn cài đặt"):
        st.markdown("""
1. [console.cloud.google.com](https://console.cloud.google.com)
2. Bật **Sheets API** + **Drive API**
3. Tạo **Service Account** → Tải JSON key
4. Chia sẻ Sheet với `client_email`
5. Upload JSON → Nhấn **Tải dữ liệu**
        """)


# ── SESSION STATE ─────────────────────────────
for k, v in [("metrics", None), ("fetch_time", None), ("err", None)]:
    if k not in st.session_state:
        st.session_state[k] = v

# ── FETCH ─────────────────────────────────────
if fetch_btn:
    if not creds_data:
        st.session_state.err = "⚠️ Chưa có credentials. Vui lòng upload file JSON ở sidebar."
        st.session_state.metrics = None
    else:
        with st.spinner("Đang kết nối Google Sheets…"):
            try:
                cl  = authenticate_gspread(creds_data)
                raw = fetch_data(cl, SHEET_ID, SHEET_NAME)
                st.session_state.metrics   = process_data(raw)
                st.session_state.fetch_time = datetime.now().strftime("%H:%M — %d/%m/%Y")
                st.session_state.err       = None
            except Exception as e:
                st.session_state.err     = f"❌ {type(e).__name__}: {e}"
                st.session_state.metrics = None

# ── HEADER ────────────────────────────────────
fetch_ts = st.session_state.fetch_time or "Chưa tải"
st.markdown(f"""
<div class="dash-header">
  <div class="dash-header-icon">🏥</div>
  <div>
    <div class="dash-header-title">Dashboard Đăng Ký Khám Online</div>
    <div class="dash-header-sub">BVĐK Tâm Đức Cầu Quan · Theo dõi lịch hẹn & tình trạng bệnh nhân</div>
  </div>
  <div class="dash-header-badge">🕐 {fetch_ts}</div>
</div>
""", unsafe_allow_html=True)

if st.session_state.err:
    st.error(st.session_state.err)

# ── DASHBOARD ─────────────────────────────────
if st.session_state.metrics:
    m = st.session_state.metrics

    # ── KPI Row ──
    def kpi(icon, label, value, sub, cls):
        return f"""
        <div class="kpi-card {cls}">
          <span class="kpi-trend">{icon}</span>
          <span class="kpi-icon">{icon}</span>
          <div class="kpi-label">{label}</div>
          <div class="kpi-value">{value}</div>
          <div class="kpi-sub">{sub}</div>
        </div>"""

    unique_dates = 0
    if COL_EXAM_DATE in m["df"].columns:
        unique_dates = m["df"][COL_EXAM_DATE].astype(str).str.strip().replace("", pd.NA).dropna().nunique()

    st.markdown('<div class="kpi-grid">', unsafe_allow_html=True)
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(kpi("📋", "Tổng Đăng Ký Online", m["total"], "Toàn bộ lịch hẹn", "kpi-blue"), unsafe_allow_html=True)
    with k2:
        st.markdown(kpi("✅", "Đã Khám", m["attended_count"], f"{m['attended_pct']}% tổng đăng ký", "kpi-green"), unsafe_allow_html=True)
    with k3:
        st.markdown(kpi("❌", "Chưa / Vắng Khám", m["noshow_count"], f"{m['noshow_pct']}% tổng đăng ký", "kpi-rose"), unsafe_allow_html=True)
    with k4:
        st.markdown(kpi("📅", "Số Ngày Khám", unique_dates, "Ngày có lịch hẹn", "kpi-violet"), unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Row 1: Donut + Gender ──
    def sec(icon, label, icon_cls):
        return f"""<div class="sec-header">
          <div class="sec-header-icon {icon_cls}">{icon}</div>
          <span class="sec-header-text">{label}</span>
        </div>"""

    c1, c2 = st.columns([1.3, 0.7])
    with c1:
        st.markdown(sec("🍩", "Tỷ Lệ Đã Khám / Vắng Khám", "sec-header-green"), unsafe_allow_html=True)
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.plotly_chart(chart_donut(m), use_container_width=True, config={"displayModeBar": False})
        # Legend below donut
        st.markdown(f"""
        <div style="display:flex;gap:1.5rem;justify-content:center;padding:0.5rem 0 0.8rem;flex-wrap:wrap;">
          <div style="display:flex;align-items:center;gap:0.5rem">
            <div style="width:12px;height:12px;border-radius:3px;background:{C_GREEN}"></div>
            <span style="font-size:0.82rem;color:#475569;font-weight:500">Đã khám — <b style="color:#0f172a">{m['attended_count']}</b> ({m['attended_pct']}%)</span>
          </div>
          <div style="display:flex;align-items:center;gap:0.5rem">
            <div style="width:12px;height:12px;border-radius:3px;background:{C_ROSE}"></div>
            <span style="font-size:0.82rem;color:#475569;font-weight:500">Chưa / Vắng — <b style="color:#0f172a">{m['noshow_count']}</b> ({m['noshow_pct']}%)</span>
          </div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with c2:
        st.markdown(sec("👥", "Phân Bố Giới Tính", "sec-header-blue"), unsafe_allow_html=True)
        fig_g = chart_gender(m)
        if fig_g:
            st.markdown('<div class="chart-card">', unsafe_allow_html=True)
            st.plotly_chart(fig_g, use_container_width=True, config={"displayModeBar": False})
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.info("Không có dữ liệu giới tính.")

    # ── Row 2: Daily ──
    fig_d = chart_daily(m)
    if fig_d:
        st.markdown(sec("📅", "Lịch Hẹn Theo Ngày Khám", "sec-header-blue"), unsafe_allow_html=True)
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.plotly_chart(fig_d, use_container_width=True, config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)

    # ── Row 3: Specialty ──
    fig_s = chart_specialty(m)
    if fig_s:
        st.markdown(sec("🩺", "Chuyên Khoa Được Đăng Ký Nhiều Nhất", "sec-header-violet"), unsafe_allow_html=True)
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.plotly_chart(fig_s, use_container_width=True, config={"displayModeBar": False})
        st.markdown('</div>', unsafe_allow_html=True)

    # ── Status table + raw data ──
    st.markdown(sec("📋", "Chi Tiết Trạng Thái", "sec-header-amber"), unsafe_allow_html=True)
    if m["status_tbl"] is not None:
        st.dataframe(m["status_tbl"], use_container_width=True, hide_index=True, height=160)

    with st.expander("📄 Xem dữ liệu thô (tất cả bệnh nhân)"):
        show_cols = [c for c in [COL_TIMESTAMP, COL_NAME, COL_EXAM_DATE,
                                  COL_STATUS, COL_SPECIALTY, COL_GENDER] if c in m["df"].columns]
        st.dataframe(m["df"][show_cols].reset_index(drop=True),
                     use_container_width=True, height=380)

else:
    if not st.session_state.err:
        st.markdown("""
        <div class="empty-state">
          <div class="empty-icon">🏥</div>
          <div class="empty-title">Chưa có dữ liệu</div>
          <div class="empty-desc">
            Upload file <strong>Service Account JSON</strong> ở sidebar
            và nhấn <strong>🔄 Tải & Cập nhật dữ liệu</strong> để bắt đầu.
          </div>
        </div>
        """, unsafe_allow_html=True)
