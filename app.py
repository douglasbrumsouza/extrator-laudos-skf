import streamlit as st
import streamlit.components.v1 as components
import zipfile, re, os, base64
from io import BytesIO
from datetime import datetime, timedelta

try:
    import pdfplumber
except ImportError:
    st.error("pip install pdfplumber"); st.stop()
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    st.error("pip install openpyxl"); st.stop()
try:
    import plotly.express as px
    import plotly.graph_objects as go
    import pandas as pd
except ImportError:
    st.error("pip install plotly pandas"); st.stop()

# ── IMAGEM BASE64 ISOLADA ─────────────────────────────────────────────────────
LOGO_B64 = "data:image/webp;base64,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"

# ── Config ────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Análise de Óleo SKF — Gerdau",
    page_icon="🛢️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ── CSS global ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@300;400;500;600;700&family=Barlow+Condensed:wght@600;700&display=swap');

/* Reset Streamlit */
html, body, [class*="css"] { font-family: 'Barlow', sans-serif !important; }
#MainMenu, footer, header { display: none !important; }
[data-testid="stHeader"]    { display: none !important; }
[data-testid="stToolbar"]   { display: none !important; }
[data-testid="stDecoration"]{ display: none !important; }
[data-testid="stStatusWidget"] { display: none !important; }
section[data-testid="stSidebar"] { display: none !important; }

/* Espaçamento do container principal */
.block-container {
    padding: 2.5rem 1.5rem 1rem 1.5rem !important; 
    max-width: 100% !important;
    margin-top: 0 !important;
}
/* Fundo Enterprise Profissional (Ardósia/Chumbo Profundo) */
.stApp { background: #0B111A !important; }

/* Gap vertical padrão mais controlado e seguro */
div[data-testid="stVerticalBlock"] { gap: 0.6rem !important; }

/* ── UPLOAD SCREEN ── */
.up-wrap {
    position: fixed; top: 0; left: 0; right: 0; bottom: 0;
    display: flex; align-items: center; justify-content: center;
    padding: 20px; z-index: 999;
    background: #0B111A;
}
.up-card {
    width: 100%; max-width: 560px;
    background: #FFFFFF; border-radius: 4px; overflow: hidden;
    box-shadow: 0 10px 40px rgba(0,0,0,0.2);
}
.up-head {
    background: #004990; /* Azul Corporativo Oficial SKF */
    padding: 26px 30px;
}
.up-logo-row { display:flex; align-items:center; gap:14px; margin-bottom:12px; }
.up-logo {
    width: 46px; height: 46px; border-radius: 4px;
    background: rgba(255,255,255,0.15); border: 1px solid rgba(255,255,255,0.25);
    display: flex; align-items: center; justify-content: center;
    font-family: 'Barlow Condensed',sans-serif; font-size: 14px;
    font-weight: 700; color: white; flex-shrink: 0;
}
.up-title {
    font-family: 'Barlow Condensed',sans-serif;
    font-size: 21px; font-weight: 700; color: white; line-height: 1.1;
}
.up-sub { font-size: 11px; color: rgba(255,255,255,0.7); margin-top: 2px; }
.up-desc { font-size: 13px; color: rgba(255,255,255,0.9); line-height: 1.6; }
.up-body { padding: 24px 30px; }
.up-steps {
    background: #F4F7F9; border: 1px solid #E5E9ED; border-radius: 4px;
    padding: 14px 18px; font-size: 13px; color: #2A323C; line-height: 2.1;
}
.up-footer {
    text-align: center; font-size: 10px;
    color: rgba(255,255,255,0.3); margin-top: 14px; letter-spacing: .4px;
}

div[data-testid="stFileUploader"],
div[data-testid="stSpinner"],
div[data-testid="stAlert"],
div.stProgress {
    width: 100% !important;
    max-width: 540px !important;
    margin: 10px auto !important;
}
div[data-testid="stButton"] { display: flex; justify-content: center; }
div[data-testid="stButton"] > button { max-width: 540px; }

/* ── DASHBOARD HEADER E BARRAS ── */
.db-header {
    background: #004990; /* Azul SKF Puro */
    border-radius: 4px; padding: 10px 16px; 
    margin-bottom: 24px !important; 
    display: flex; align-items: center; justify-content: space-between;
    box-shadow: 0 2px 10px rgba(0,0,0,0.2);
}
.db-logo {
    width: 36px; height: 36px; border-radius: 4px;
    background: rgba(255,255,255,0.15); border: 1px solid rgba(255,255,255,0.2);
    display: flex; align-items: center; justify-content: center;
    font-family: 'Barlow Condensed',sans-serif; font-size: 12px;
    font-weight: 700; color: white; flex-shrink: 0;
}
.db-title {
    font-family: 'Barlow Condensed',sans-serif;
    font-size: 18px; font-weight: 700; color: white; letter-spacing: .5px;
}
.db-sub { font-size: 10px; color: rgba(255,255,255,0.7); margin-top: 1px; }
.db-badge {
    font-size: 9px; font-weight: 600; color: #FFFFFF;
    background: rgba(255,255,255,0.2); padding: 2px 10px;
    border-radius: 2px; border: 1px solid rgba(255,255,255,0.3);
}
.db-credit { font-size: 8px; color: rgba(255,255,255,0.3); margin-top: 3px; text-align:right; }

.act-bar {
    background: #151F2E; border: 1px solid #243447; border-radius: 4px;
    padding: 0 16px; display: flex; align-items: center; gap: 10px;
    height: 42px;
}
.act-title {
    font-family: 'Barlow Condensed',sans-serif; font-size: 14px;
    font-weight: 700; color: white; flex: 1;
}
.bcnt {
    font-size: 10px; font-weight: 700; padding: 2px 8px;
    border-radius: 2px; white-space: nowrap;
}

/* ── KPI CARDS (Flat Industrial) ── */
.kpi { border-radius: 4px; padding: 16px 10px; text-align: center;
       background: #151F2E !important; border: 1px solid #243447 !important; 
       position: relative; overflow: hidden; height: 100%; box-sizing: border-box; }
.kpi::before { content:''; position:absolute; top:0; left:0; right:0; height:3px; border-radius:4px 4px 0 0; }
.kpi-tot::before { background:#004990; }
.kpi-nor::before { background:#059669; }
.kpi-alt::before { background:#D97706; }
.kpi-alm::before { background:#DC2626; }

.kv { font-family:'Barlow Condensed',sans-serif; font-size:32px; font-weight:700; line-height:1; }
.kpi-tot .kv { color:#FFFFFF; }
.kpi-nor .kv { color:#10B981; }
.kpi-alt .kv { color:#F59E0B; }
.kpi-alm .kv { color:#EF4444; }
.kl { font-size:10px; font-weight:600; text-transform:uppercase; letter-spacing:1px; margin-top:6px; color:#8B9BB4 !important; }
.ks { font-size:10px; margin-top:2px; color:#5D6B82 !important; }

/* ── CAIXAS DOS GRÁFICOS (Clean UI) ── */
.sc { background:#151F2E; border:1px solid #243447; border-radius:4px; padding:12px 14px; }

.sc-top {
    background:#151F2E; border:1px solid #243447; border-bottom:none;
    border-radius:4px 4px 0 0; padding:12px 14px 20px 14px;
    position: relative; z-index: 1;
}
div[data-testid="stPlotlyChart"] {
    background: #151F2E; border: 1px solid #243447; border-top: none;
    border-radius: 0 0 4px 4px; padding: 0 10px 10px 10px;
    margin-top: -24px !important;
    position: relative; z-index: 10;
}

.st { font-size:10px; font-weight:600; text-transform:uppercase; letter-spacing:1px;
      color:#E2E8F0; display:flex; align-items:center; gap:6px; margin-bottom:8px; }
.st::before { content:''; display:inline-block; width:3px; height:12px;
              border-radius:2px; background:#004990; }

/* ── TABELA (Corporate Grid) ── */
.tbl { width:100%; border-collapse:collapse; font-size:11px; }
.tbl thead tr { background:#0F1722; }
.tbl thead th { padding:8px 10px; text-align:left; font-size:9px; font-weight:600;
                text-transform:uppercase; letter-spacing:.7px; color:#8B9BB4;
                border-bottom:2px solid #243447; }
.tbl tbody tr { border-bottom:1px solid #1C2838; }
.tbl tbody tr:hover { background:#1A2636; }
.tbl tbody td { padding:8px 10px; color:#B0BACC; }
.td-eq { color:white !important; font-weight:500 !important; }
.td-set { font-family:'Barlow Condensed',sans-serif; font-size:13px; font-weight:700; color:#60A5FA !important; }
.td-par { color:#D97706 !important; }
.td-obs { font-size:10px; color:#8B9BB4 !important; font-style:italic; }
.td-data { font-size:10px; color:#5D6B82 !important; }
.td-red { color:#EF4444 !important; font-weight:600 !important; }
.td-yel { color:#F59E0B !important; font-weight:600 !important; }
.badge { display:inline-block; font-size:9px; font-weight:700; padding:2px 8px;
         border-radius:2px; text-transform:uppercase; letter-spacing:.4px; }
.ba { background:rgba(220,38,38,.1); color:#EF4444; border:1px solid rgba(220,38,38,.3); }
.bl { background:rgba(217,119,6,.1); color:#F59E0B; border:1px solid rgba(217,119,6,.3); }
.bn { background:rgba(5,150,105,.1); color:#10B981; border:1px solid rgba(5,150,105,.3); }

/* ── SELECTBOX LABELS ── */
div[data-testid="stSelectbox"] label {
    color:#8B9BB4 !important; font-size:11px !important;
    text-transform:uppercase; letter-spacing:.8px;
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# EXTRAÇÃO
# ══════════════════════════════════════════════════════════════════════════════
PARAMETROS = [
    ("Ferro","Ferro (ppm)"),("Crômio","Crômio (ppm)"),("Níquel","Níquel (ppm)"),
    ("Alumínio","Alumínio (ppm)"),("Chumbo","Chumbo (ppm)"),("Cobre
