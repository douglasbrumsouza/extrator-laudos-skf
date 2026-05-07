import streamlit as st
import streamlit.components.v1 as components
import zipfile, re, os, base64
from io import BytesIO
from datetime import datetime

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

/* Espaçamento do container principal para não cortar as labels no topo */
.block-container {
    padding: 2.5rem 1.5rem 1rem 1.5rem !important; 
    max-width: 100% !important;
    margin-top: 0 !important;
}
.stApp { background: linear-gradient(135deg,#0D2137 0%,#1A3A5C 100%) !important; }

/* Layout Limpo: sem margens negativas. Usando o GAP nativo seguro do Streamlit */
div[data-testid="stVerticalBlock"] { gap: 1rem !important; }

/* ── UPLOAD SCREEN ── */
.up-wrap {
    position: fixed; top: 0; left: 0; right: 0; bottom: 0;
    display: flex; align-items: center; justify-content: center;
    padding: 20px; z-index: 999;
    background: linear-gradient(135deg,#0D2137 0%,#1A3A5C 100%);
}
.up-card {
    width: 100%; max-width: 560px;
    background: white; border-radius: 16px; overflow: hidden;
    box-shadow: 0 24px 80px rgba(0,0,0,0.5);
}
.up-head {
    background: linear-gradient(135deg,#1F4E79,#2E75B6);
    padding: 26px 30px;
}
.up-logo-row { display:flex; align-items:center; gap:14px; margin-bottom:12px; }
.up-logo {
    width: 46px; height: 46px; border-radius: 11px;
    background: rgba(255,255,255,0.15); border: 1px solid rgba(255,255,255,0.25);
    display: flex; align-items: center; justify-content: center;
    font-family: 'Barlow Condensed',sans-serif; font-size: 14px;
    font-weight: 700; color: white; flex-shrink: 0;
}
.up-title {
    font-family: 'Barlow Condensed',sans-serif;
    font-size: 21px; font-weight: 700; color: white; line-height: 1.1;
}
.up-sub { font-size: 11px; color: rgba(255,255,255,0.6); margin-top: 2px; }
.up-desc { font-size: 13px; color: rgba(255,255,255,0.82); line-height: 1.6; }
.up-body { padding: 24px 30px; }
.up-steps {
    background: #F0F6FF; border: 1px solid #C8DCEF; border-radius: 10px;
    padding: 14px 18px; font-size: 13px; color: #1F4E79; line-height: 2.1;
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
    background: linear-gradient(135deg,#1F4E79,#2E75B6);
    border-radius: 10px; padding: 10px 16px; margin-bottom: 2px;
    display: flex; align-items: center; justify-content: space-between;
    box-shadow: 0 4px 20px rgba(0,0,0,0.3);
}
.db-logo {
    width: 36px; height: 36px; border-radius: 8px;
    background: rgba(255,255,255,0.15); border: 1px solid rgba(255,255,255,0.2);
    display: flex; align-items: center; justify-content: center;
    font-family: 'Barlow Condensed',sans-serif; font-size: 12px;
    font-weight: 700; color: white; flex-shrink: 0;
}
.db-title {
    font-family: 'Barlow Condensed',sans-serif;
    font-size: 18px; font-weight: 700; color: white; letter-spacing: .5px;
}
.db-sub { font-size: 10px; color: rgba(255,255,255,0.6); margin-top: 1px; }
.db-badge {
    font-size: 9px; font-weight: 600; color: #BDD7EE;
    background: rgba(46,117,182,0.3); padding: 2px 10px;
    border-radius: 20px; border: 1px solid rgba(46,117,182,0.5);
}
.db-credit { font-size: 8px; color: rgba(255,255,255,0.22); margin-top: 3px; text-align:right; }

.act-bar {
    background: #0D2137; border: 1px solid #1A3A5C; border-radius: 8px;
    padding: 0 16px; display: flex; align-items: center; gap: 10px;
    height: 42px;
}
.act-title {
    font-family: 'Barlow Condensed',sans-serif; font-size: 14px;
    font-weight: 700; color: white; flex: 1;
}
.bcnt {
    font-size: 10px; font-weight: 700; padding: 2px 8px;
    border-radius: 20px; white-space: nowrap;
}

/* ── KPI CARDS ── */
.kpi { border-radius: 10px; padding: 16px 10px; text-align: center;
       border: 1px solid transparent; position: relative; overflow: hidden; 
       height: 100%; box-sizing: border-box; }
.kpi::before { content:''; position:absolute; top:0; left:0; right:0; height:3px; border-radius:10px 10px 0 0; }
.kpi-tot { background:#162E4A; border-color:#2A4A6A; }
.kpi-tot::before { background:#2E75B6; }
.kpi-nor { background:#0E2B16; border-color:#1A4D22; }
.kpi-nor::before { background:#4CAF50; }
.kpi-alt { background:#2B1E05; border-color:#4A3210; }
.kpi-alt::before { background:#F59E0B; }
.kpi-alm { background:#2B0A0A; border-color:#4A1A1A; }
.kpi-alm::before { background:#EF4444; }

.kv { font-family:'Barlow Condensed',sans-serif; font-size:32px; font-weight:700; line-height:1; }
.kpi-tot .kv { color:white; }
.kpi-nor .kv { color:#6BCB77; }
.kpi-alt .kv { color:#F59E0B; }
.kpi-alm .kv { color:#EF4444; }
.kl { font-size:10px; font-weight:600; text-transform:uppercase; letter-spacing:1px; margin-top:6px; }
.kpi-tot .kl { color:#5A7A9A; } .kpi-nor .kl { color:#3A7D44; }
.kpi-alt .kl { color:#8B6010; } .kpi-alm .kl { color:#8B2020; }
.ks { font-size:10px; margin-top:2px; }
.kpi-tot .ks { color:#3A5A7A; } .kpi-nor .ks { color:#2A5A30; }
.kpi-alt .ks { color:#6A4A08; } .kpi-alm .ks { color:#6A1010; }

/* ── CAIXAS DOS GRÁFICOS (Arquitetura Perfeita e Segura) ── */
div[data-testid="stPlotlyChart"] {
    background: #112035; 
    border: 1px solid #1A3A5C; 
    border-radius: 10px; 
    padding: 10px 10px 5px 10px;
}

/* ── TABELA E CONTAINER INFERIOR ── */
.sc { background:#112035; border:1px solid #1A3A5C; border-radius:10px; padding:16px 14px; }
.st { font-size:11px; font-weight:600; text-transform:uppercase; letter-spacing:1px;
      color:#BDD7EE; display:flex; align-items:center; gap:6px; margin-bottom:12px; }
.st::before { content:''; display:inline-block; width:3px; height:12px;
              border-radius:2px; background:#2E75B6; }

.tbl { width:100%; border-collapse:collapse; font-size:11px; }
.tbl thead tr { background:#0D2137; }
.tbl thead th { padding:8px 10px; text-align:left; font-size:9px; font-weight:600;
                text-transform:uppercase; letter-spacing:.7px; color:#3A6A9A;
                border-bottom:1px solid #1A3A5C; }
.tbl tbody tr { border-bottom:1px solid #152840; }
.tbl tbody tr:hover { background:#162E4A; }
.tbl tbody td { padding:8px 10px; color:#8AAABB; }
.td-eq { color:white !important; font-weight:500 !important; }
.td-set { font-family:'Barlow Condensed',sans-serif; font-size:13px; font-weight:700; color:#2E75B6 !important; }
.td-par { color:#C8A870 !important; }
.td-obs { font-size:10px; color:#8A6A3A !important; font-style:italic; }
.td-data { font-size:10px; color:#3A6A9A !important; }
.td-red { color:#EF4444 !important; font-weight:600 !important; }
.td-yel { color:#F59E0B !important; font-weight:600 !important; }
.badge { display:inline-block; font-size:9px; font-weight:700; padding:2px 8px;
         border-radius:20px; text-transform:uppercase; letter-spacing:.4px; }
.ba { background:rgba(239,68,68,.15); color:#EF4444; border:1px solid rgba(239,68,68,.3); }
.bl { background:rgba(245,158,11,.15); color:#F59E0B; border:1px solid rgba(245,158,11,.3); }
.bn { background:rgba(76,175,80,.15);  color:#4CAF50; border:1px solid rgba(76,175,80,.3); }

/* ── SELECTBOX LABELS ── */
div[data-testid="stSelectbox"] label {
    color:#BDD7EE !important; font-size:11px !important;
    text-transform:uppercase; letter-spacing:.8px;
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# EXTRAÇÃO
# ══════════════════════════════════════════════════════════════════════════════
PARAMETROS = [
    ("Ferro","Ferro (ppm)"),("Crômio","Crômio (ppm)"),("Níquel","Níquel (ppm)"),
    ("Alumínio","Alumínio (ppm)"),("Chumbo","Chumbo (ppm)"),("Cobre","Cobre (ppm)"),
    ("Estanho","Estanho (ppm)"),("Titânio","Titânio (ppm)"),("Prata","Prata (ppm)"),
    ("Antimônio","Antimônio (ppm)"),("Cádmio","Cádmio (ppm)"),("Manganês","Manganês (ppm)"),
    ("Bismuto","Bismuto (ppm)"),("Arsênio","Arsênio (ppm)"),("Índio","Índio (ppm)"),
    ("Cobalto","Cobalto (ppm)"),("Zircônio","Zircônio (ppm)"),("Tungstênio","Tungstênio (ppm)"),
    ("Cério","Cério (ppm)"),("Silício","Silício (ppm)"),("Sódio","Sódio (ppm)"),
    ("Vanádio","Vanádio (ppm)"),("Potássio","Potássio (ppm)"),("Lítio","Lítio (ppm)"),
    ("água","Água (ppm)"),("Bolhas","Bolhas"),("Molibdênio","Molibdênio (ppm)"),
    ("Cálcio","Cálcio (ppm)"),("Magnésio","Magnésio (ppm)"),("Fósforo","Fósforo (ppm)"),
    ("Zinco","Zinco (ppm)"),("Bário","Bário (ppm)"),("Boro","Boro (ppm)"),
    ("TAN","TAN (mg KOH/g)"),("Oxidação","Oxidação (abs/0.1mm)"),
    ("Visc 40","Visc 40 (cSt)"),("Integridade de Fluído","Integridade de Fluído"),
]

def to_num(v):
    if v is None: return None
    s = str(v).strip().replace(",", ".")
    try: return float(s)
    except: return s if s else None

def detectar_status(pdf_bytes):
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        full = "".join(p.extract_text() or "" for p in pdf.pages)
        if "Limites acionados" not in full: return "Normal"
        for img in pdf.pages[0].images:
            if img.get("name") == "Im2":
                return "Alarme" if img["height"] >= 50 else "Alerta"
    return "Alerta"

def split_ativo(raw):
    p = [x.strip() for x in raw.split("→")]
    s = p[0].strip().split("-")
    c = {f"Cod{i}": s[i-1] if len(s) >= i else "" for i in range(1, 6)}
    c["Equipamento Descrição 1"] =
