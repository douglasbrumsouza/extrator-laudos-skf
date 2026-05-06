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
.block-container {
    padding: 0.6rem 1.2rem 1rem 1.2rem !important;
    max-width: 100% !important;
    margin-top: 0 !important;
}
.stApp { background: linear-gradient(135deg,#0D2137 0%,#1A3A5C 100%) !important; }
div[data-testid="stVerticalBlock"] { gap: 0.35rem !important; }

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

/* ── CORREÇÃO DE ALINHAMENTO DA TELA DE UPLOAD ── */
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

/* ── DASHBOARD ── */
.db-header {
    background: linear-gradient(135deg,#1F4E79,#2E75B6);
    border-radius: 10px; padding: 13px 20px; margin-bottom: 6px;
    display: flex; align-items: center; justify-content: space-between;
    box-shadow: 0 4px 20px rgba(0,0,0,0.3);
}
.db-logo {
    width: 40px; height: 40px; border-radius: 9px;
    background: rgba(255,255,255,0.15); border: 1px solid rgba(255,255,255,0.2);
    display: flex; align-items: center; justify-content: center;
    font-family: 'Barlow Condensed',sans-serif; font-size: 13px;
    font-weight: 700; color: white; flex-shrink: 0;
}
.db-title {
    font-family: 'Barlow Condensed',sans-serif;
    font-size: 19px; font-weight: 700; color: white; letter-spacing: .5px;
}
.db-sub { font-size: 10px; color: rgba(255,255,255,0.6); margin-top: 1px; }
.db-badge {
    font-size: 9px; font-weight: 600; color: #BDD7EE;
    background: rgba(46,117,182,0.3); padding: 2px 10px;
    border-radius: 20px; border: 1px solid rgba(46,117,182,0.5);
}
.db-credit { font-size: 8px; color: rgba(255,255,255,0.22); margin-top: 3px; text-align:right; }

/* ── CORREÇÃO DE ALINHAMENTO DA ACTION BAR ── */
.act-bar {
    background: #0D2137; border: 1px solid #1A3A5C; border-radius: 8px;
    padding: 0 16px; display: flex; align-items: center; gap: 10px;
    height: 42px; margin-bottom: 0px; margin-top: 2px;
}
.act-title {
    font-family: 'Barlow Condensed',sans-serif; font-size: 15px;
    font-weight: 700; color: white; flex: 1;
}
.bcnt {
    font-size: 10px; font-weight: 700; padding: 2px 9px;
    border-radius: 20px; white-space: nowrap;
}

/* ── KPI CARDS ── */
.kpi { border-radius: 10px; padding: 13px 14px; text-align: center;
       border: 1px solid transparent; position: relative; overflow: hidden; }
.kpi::before { content:''; position:absolute; top:0; left:0; right:0;
               height:3px; border-radius:10px 10px 0 0; }
.kpi-tot { background:#162E4A; border-color:#2A4A6A; }
.kpi-tot::before { background:#2E75B6; }
.kpi-nor { background:#0E2B16; border-color:#1A4D22; }
.kpi-nor::before { background:#4CAF50; }
.kpi-alt { background:#2B1E05; border-color:#4A3210; }
.kpi-alt::before { background:#F59E0B; }
.kpi-alm { background:#2B0A0A; border-color:#4A1A1A; }
.kpi-alm::before { background:#EF4444; }
.kv { font-family:'Barlow Condensed',sans-serif; font-size:42px; font-weight:700; line-height:1; }
.kpi-tot .kv { color:white; }
.kpi-nor .kv { color:#6BCB77; }
.kpi-alt .kv { color:#F59E0B; }
.kpi-alm .kv { color:#EF4444; }
.kl { font-size:9px; font-weight:600; text-transform:uppercase; letter-spacing:1.2px; margin-top:3px; }
.kpi-tot .kl { color:#5A7A9A; } .kpi-nor .kl { color:#3A7D44; }
.kpi-alt .kl { color:#8B6010; } .kpi-alm .kl { color:#8B2020; }
.ks { font-size:10px; margin-top:1px; }
.kpi-tot .ks { color:#3A5A7A; } .kpi-nor .ks { color:#2A5A30; }
.kpi-alt .ks { color:#6A4A08; } .kpi-alm .ks { color:#6A1010; }

/* ── SECTION CARDS ── */
.sc { background:#112035; border:1px solid #1A3A5C; border-radius:10px;
      padding:12px 14px; }

/* ── CORREÇÃO DAS CAIXAS DOS GRÁFICOS ── */
.sc-top {
    background:#112035; border:1px solid #1A3A5C; border-bottom:none;
    border-radius:10px 10px 0 0; padding:12px 14px 4px 14px;
}
div[data-testid="stPlotlyChart"] {
    background: #112035;
    border: 1px solid #1A3A5C;
    border-top: none;
    border-radius: 0 0 10px 10px;
    padding: 0 10px 10px 10px;
    margin-top: -12px !important;
}

.st { font-size:10px; font-weight:600; text-transform:uppercase; letter-spacing:1px;
      color:#BDD7EE; display:flex; align-items:center; gap:6px; margin-bottom:8px; }
.st::before { content:''; display:inline-block; width:3px; height:12px;
              border-radius:2px; background:#2E75B6; }

/* ── TABELA ── */
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

/* ── Selectbox Labels ── */
div[data-testid="stSelectbox"] label {
    color:#BDD7EE !important; font-size:10px !important;
    text-transform:uppercase; letter-spacing:.8px;
    margin-bottom: 4px !important; /* Afasta um pouco do input para ficar mais alinhado */
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
    c["Equipamento Descrição 1"] = p[1] if len(p) > 1 else ""
    c["Equipamento Descrição 2"] = p[2] if len(p) > 2 else ""
    return c

def split_data(s):
    if not s: return None, None, None, None
    m = re.match(r"(\d{2})/(\d{2})/(\d{4}),?\s*(\d{2}:\d{2}:\d{2})?", str(s).strip())
    if not m: return None, None, None, None
    return int(m.group(1)), int(m.group(2)), int(m.group(3)), m.group(4)

def bloco(t, ini, fim):
    m = re.search(rf"{re.escape(ini)}\s*\n(.*?)(?={re.escape(fim)})", t, re.DOTALL)
    return m.group(1).strip() if m else ""

def extrair_laudo(nome, pdf_bytes):
    dados = {"Arquivo": nome}
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        pags = [p.extract_text() or "" for p in pdf.pages]
    p1 = pags[0]; texto = "\n".join(pags)
    dados["Status"] = detectar_status(pdf_bytes)
    m = re.search(r"Instalação:\s*(.+?)(?:\s{2,}|\n)", p1)
    dados["Instalação"] = m.group(1).strip() if m else None
    m = re.search(r"Localização:\s*(.+)", p1)
    loc = m.group(1).strip() if m else ""
    if loc.endswith("-"):
        m2 = re.search(r"Localização:.*\n(.+?)(?:\n|$)", p1)
        if m2: loc = loc + " " + m2.group(1).strip()
    dados["Localização"] = loc
    m = re.search(r"Ativo:\s*(.+?)(?=Número de série:|\n\n)", p1, re.DOTALL)
    raw = re.sub(r"\s+", " ", m.group(1)).strip() if m else ""
    dados["Ativo (completo)"] = raw; dados.update(split_ativo(raw))
    m = re.search(r"Tipo de componente:\s*(.+?)(?=Operação de ativo:|\n)", p1)
    dados["Tipo de componente"] = m.group(1).strip() if m else None
    m = re.search(r"Óleo:\s*(.+?)(?=Operação de óleo:|\n)", p1)
    dados["Óleo"] = m.group(1).strip() if m else None
    m = re.search(r"data de coleta:\s*([\d/,: ]+)", p1)
    ds = m.group(1).strip() if m else None
    dados["Data de coleta"] = ds
    d, me, a, h = split_data(ds)
    dados["Dia coleta"] = d; dados["Mês coleta"] = me
    dados["Ano coleta"] = a; dados["Hora coleta"] = h
    dados["Observações"]              = bloco(p1, "Observações:", "Diagnósticos:")
    dados["Diagnósticos"]             = bloco(p1, "Diagnósticos:", "Ações:")
    dados["Ações"]                    = bloco(p1, "Ações:", "Recomendações adicionais:")
    dados["Recomendações adicionais"] = bloco(p1, "Recomendações adicionais:", "Teste realizado por:")
    m = re.search(r"Teste realizado por:\s*(.+?)(?=Lançado por:|\n)", p1)
    dados["Teste realizado por"] = m.group(1).strip() if m else None
    m = re.search(r"Lançado por:\s*(.+?)(?=_{3,}|\n)", p1)
    dados["Lançado por"] = m.group(1).strip() if m else None
    m = re.search(r"ID de amostra\s+([\w-]+)", p1)
    dados["ID de amostra"] = m.group(1).strip() if m else None
    m = re.search(r"Data de teste\s+(\d{2}/\d{2}/\d{4},\s*\d{2}:\d{2}:\d{2})", p1)
    dt = m.group(1).strip() if m else None
    dados["Data de teste"] = dt
    d, me, a, _ = split_data(dt)
    dados["Dia teste"] = d; dados["Mês teste"] = me; dados["Ano teste"] = a
    m = re.search(r"Data de lançamento\s+(\d{2}/\d{2}/\d{4},\s*\d{2}:\d{2}:\d{2})", p1)
    dados["Data de lançamento"] = m.group(1).strip() if m else None
    linhas = texto.split("\n"); pm = {}
    for linha in linhas:
        ls = linha.strip()
        for raw, _ in PARAMETROS:
            if raw not in pm and (ls.startswith(raw + " ") or ls == raw):
                parts = ls.split(); nw = len(raw.split())
                if len(parts) > nw:
                    v = parts[nw]
                    if re.match(r"^[\d\-]", v) or v in ("Normal", "Desconhecido"):
                        pm[raw] = v
                break
    for raw, col in PARAMETROS:
        dados[col] = to_num(pm.get(raw))
    m = re.search(r"Limites acionados \(amostra atual\)\n(.+?)(?=TruVu|Relatório)", texto, re.DOTALL)
    if m:
        pa = []
        for l in m.group(1).strip().split("\n"):
            pm2 = re.match(r"^\d+\s+(.+?)\s+(?:Absoluto|Desvio)", l.strip())
            if pm2: pa.append(pm2.group(1).strip())
        dados["Parâmetros em Alarme"] = "; ".join(pa)
    else:
        dados["Parâmetros em Alarme"] = ""
    return dados

def gerar_excel(lista):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Laudos"
    colunas = list(lista[0].keys())
    cor_st = {"Normal": "C6EFCE", "Alerta": "FFEB9C", "Alarme": "FFC7CE"}
    hf = PatternFill("solid", start_color="1F4E79")
    hfont = Font(bold=True, color="FFFFFF", name="Arial", size=9)
    borda = Border(left=Side(style="thin"), right=Side(style="thin"),
                   top=Side(style="thin"), bottom=Side(style="thin"))
    for ci, col in enumerate(colunas, 1):
        c = ws.cell(row=1, column=ci, value=col)
        c.fill = hf; c.font = hfont; c.border = borda
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 36
    for ri, linha in enumerate(lista, 2):
        st_val = linha.get("Status", "Normal")
        fb = PatternFill("solid", start_color="EBF3FB" if ri % 2 == 0 else "FFFFFF")
        for ci, col in enumerate(colunas, 1):
            c = ws.cell(row=ri, column=ci, value=linha.get(col))
            c.font = Font(name="Arial", size=9); c.border = borda
            c.alignment = Alignment(horizontal="center", vertical="center")
            if col == "Status":
                c.fill = PatternFill("solid", start_color=cor_st.get(st_val, "FFFFFF"))
                c.font = Font(name="Arial", size=9, bold=True)
            else: c.fill = fb
    larg = {"Arquivo":35,"Status":10,"Localização":28,"Ativo (completo)":45,
            "Cod1":7,"Cod2":7,"Cod3":8,"Cod4":9,"Cod5":7,
            "Equipamento Descrição 1":26,"Equipamento Descrição 2":26,
            "Tipo de componente":28,"Óleo":24,"Data de coleta":20,
            "Dia coleta":8,"Mês coleta":8,"Ano coleta":8,"Hora coleta":10,
            "Observações":40,"Diagnósticos":50,"Ações":50,"Parâmetros em Alarme":30}
    for ci, col in enumerate(colunas, 1):
        ws.column_dimensions[get_column_letter(ci)].width = larg.get(col, 14)
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return buf.read()

# ══════════════════════════════════════════════════════════════════════════════
# TELA DE UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
def render_upload():
    components.html("""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="UTF-8">
    <link href="https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700&family=Barlow+Condensed:wght@600;700&display=swap" rel="stylesheet">
    <style>
    * { margin:0; padding:0; box-sizing:border-box; }
    body {
        font-family: 'Barlow', sans-serif;
        background: linear-gradient(135deg,#0D2137 0%,#1A3A5C 100%);
        min-height: 100vh; display: flex;
        align-items: center; justify-content: center; padding: 20px;
    }
    .card {
        width: 100%; max-width: 540px;
        background: white; border-radius: 16px; overflow: hidden;
        box-shadow: 0 24px 80px rgba(0,0,0,0.5);
    }
    .head {
        background: linear-gradient(135deg,#1F4E79,#2E75B6);
        padding: 26px 30px;
    }
    .logo-row { display:flex; align-items:center; gap:14px; margin-bottom:12px; }
    .logo {
        width:46px; height:46px; border-radius:11px;
        background:rgba(255,255,255,0.15); border:1px solid rgba(255,255,255,0.25);
        display:flex; align-items:center; justify-content:center;
        font-family:'Barlow Condensed',sans-serif; font-size:15px;
        font-weight:700; color:white;
    }
    .title { font-family:'Barlow Condensed',sans-serif; font-size:22px; font-weight:700; color:white; }
    .sub { font-size:11px; color:rgba(255,255,255,0.6); margin-top:2px; }
    .desc { font-size:13px; color:rgba(255,255,255,0.82); line-height:1.6; }
    .body { padding: 24px 30px; }
    .steps {
        background:#F0F6FF; border:1px solid #C8DCEF; border-radius:10px;
        padding:14px 18px; font-size:13px; color:#1F4E79; line-height:2.2;
    }
    .steps b { color:#0D2137; }
    .foot {
        text-align:center; font-size:10px; color:rgba(255,255,255,0.3);
        padding-bottom:20px; letter-spacing:.4px;
    }
    </style>
    </head>
    <body>
    <div class="card">
      <div class="head">
        <div class="logo-row">
          <div class="logo">SKF</div>
          <div>
            <div class="title">Extrator de Laudos SKF</div>
            <div class="sub">TruVu 360 · Gerdau Charqueadas · Eng. de Manutenção</div>
          </div>
        </div>
        <div class="desc">
          Faça o upload do ZIP com os laudos em PDF.<br>
          Gera o <strong>Excel consolidado</strong> e o <strong>painel interativo</strong> automaticamente.
        </div>
      </div>
      <div class="body">
        <div class="steps">
          📦 &nbsp;<b>1.</b> Receba o ZIP do laboratório SKF<br>
          ⬆️ &nbsp;<b>2.</b> Selecione o arquivo abaixo<br>
          ⚙️ &nbsp;<b>3.</b> Clique em Processar Laudos<br>
          📊 &nbsp;<b>4.</b> Veja o painel e baixe o Excel
        </div>
      </div>
      <div class="foot">Desenvolvido por Douglas Brum · Gerdau Charqueadas · SKF TruVu 360</div>
    </div>
    </body>
    </html>
    """, height=440)

    # Widgets Streamlit — Renderizados na raiz
    uploaded = st.file_uploader(
        "📦 Selecione o arquivo ZIP com os laudos",
        type=["zip"],
        label_visibility="collapsed"
    )
    if uploaded:
        st.success(f"✅ **{uploaded.name}** — {uploaded.size/1024/1024:.1f} MB")
        if st.button("⚙️  Processar Laudos", type="primary", use_container_width=True):
            with st.spinner("Processando laudos..."):
                zf = zipfile.ZipFile(BytesIO(uploaded.read()))
                pdfs = [(n, zf.read(n)) for n in zf.namelist()
                        if n.lower().endswith(".pdf") and not n.startswith("__MACOSX")]
            if not pdfs:
                st.error("Nenhum PDF encontrado no ZIP."); return
            prog = st.progress(0)
            lista = []
            for i, (nome, pdf_bytes) in enumerate(sorted(pdfs), 1):
                prog.progress(i / len(pdfs),
                    text=f"Processando {i}/{len(pdfs)}: {os.path.basename(nome)[:45]}")
                try: lista.append(extrair_laudo(os.path.basename(nome), pdf_bytes))
                except: pass
            prog.progress(1.0, text="Concluído!")
            st.session_state["laudos"] = lista
            st.session_state["processado"] = True
            st.session_state["excel"] = f"LAUDOS_{datetime.now().strftime('%d-%m-%Y')}.xlsx"
            st.rerun()
    else:
        st.markdown("""
        <p style="text-align:center;color:rgba(255,255,255,0.4);font-size:12px;margin-top:4px;width:100%;max-width:540px;margin-left:auto;margin-right:auto;">
        Arraste o arquivo ZIP aqui ou clique para selecionar
        </p>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
def render_dashboard(lista):
    df = pd.DataFrame(lista)
    MESES = {1:"Jan",2:"Fev",3:"Mar",4:"Abr",5:"Mai",6:"Jun",
             7:"Jul",8:"Ago",9:"Set",10:"Out",11:"Nov",12:"Dez"}
    MESES_F = {1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",6:"Junho",
               7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"}

    # ── Barra de ações ─────────────────────────────────────────────────────────
    total_g = len(df)
    norm_g  = len(df[df["Status"]=="Normal"])
    alt_g   = len(df[df["Status"]=="Alerta"])
    alm_g   = len(df[df["Status"]=="Alarme"])

    excel_bytes = gerar_excel(lista)
    c1, c2, c3, c4 = st.columns([3.5, 1.2, 1.1, 1.2])
    with c1:
        st.markdown(f"""
        <div class="act-bar">
          <span class="act-title">{total_g} laudos processados</span>
          <span class="bcnt" style="background:rgba(76,175,80,.15);color:#6BCB77;border:1px solid rgba(76,175,80,.3)">{norm_g} Normal</span>
          <span class="bcnt" style="background:rgba(245,158,11,.15);color:#F59E0B;border:1px solid rgba(245,158,11,.3)">{alt_g} Alerta</span>
          <span class="bcnt" style="background:rgba(239,68,68,.15);color:#EF4444;border:1px solid rgba(239,68,68,.3)">{alm_g} Alarme</span>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.download_button("⬇️  Baixar Excel", data=excel_bytes,
                           file_name=st.session_state["excel"],
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True, type="primary")
    with c3:
        if st.button("↩  Novo ZIP", use_container_width=True):
            for k in ["laudos","processado","excel"]:
                st.session_state.pop(k, None)
            st.rerun()
    with c4:
        components.html("""
        <style>
        * { margin:0; padding:0; box-sizing:border-box; }
        body { background:transparent; font-family:'Barlow',sans-serif; }
        button {
            width:100%; height:40px; 
            background: linear-gradient(135deg,#1F4E79,#2E75B6);
            color:white; border:none; border-radius:8px;
            font-size:13px; font-weight:600; cursor:pointer;
            display:flex; align-items:center; justify-content:center; gap:6px;
            transition:all .2s; letter-spacing:.3px;
        }
        button:hover { opacity:.85; }
        </style>
        <button onclick="entrarTelaCheia()">⛶ &nbsp;Apresentar</button>
        <script>
        function entrarTelaCheia() {
            try {
                var p = window.parent;
                var doc = p.document;
                var el  = doc.documentElement;

                // Fullscreen
                var fn = el.requestFullscreen || el.webkitRequestFullscreen
                      || el.mozRequestFullScreen || el.msRequestFullscreen;
                if (fn) fn.call(el);

                // Oculta todos os elementos do Streamlit
                function hide() {
                    var sels = [
                        'header[data-testid="stHeader"]',
                        '[data-testid="stToolbar"]',
                        '[data-testid="stDecoration"]',
                        '[data-testid="stStatusWidget"]',
                        '#MainMenu', 'footer'
                    ];
                    sels.forEach(function(s) {
                        var el = doc.querySelector(s);
                        if (el) el.style.cssText = 'display:none!important';
                    });
                    var bc = doc.querySelector('.block-container');
                    if (bc) { bc.style.paddingTop='0.3rem'; bc.style.paddingBottom='0'; }
                }
                hide(); setTimeout(hide, 500); setTimeout(hide, 1200);
            } catch(e) { console.log('Fullscreen error:', e); }
        }
        </script>
        """, height=42)

    # ── Header do dashboard ────────────────────────────────────────────────────
    ultima = df["Data de coleta"].dropna().iloc[-1] if not df["Data de coleta"].dropna().empty else "—"
    st.markdown(f"""
    <div class="db-header">
      <div style="display:flex;align-items:center;gap:12px">
        <div class="db-logo">SKF</div>
        <div>
          <div class="db-title">MONITORAMENTO DE ANÁLISE DE ÓLEO</div>
          <div class="db-sub">Gerdau Charqueadas · Engenharia de Manutenção · TruVu 360</div>
        </div>
      </div>
      <div style="text-align:right">
        <div class="db-badge">● AO VIVO</div>
        <div style="font-size:10px;color:#BDD7EE;margin-top:3px">
          Última coleta: {ultima} &nbsp;|&nbsp; {datetime.now().strftime('%d/%m/%Y %H:%M')}
        </div>
        <div class="db-credit">Desenvolvido por Douglas Brum · Gerdau Charqueadas</div>
      </div>
    </div>""", unsafe_allow_html=True)

    # ── Filtros (Malha 2:1:1) ──────────────────────────────────────────────────
    setores = ["Todos"] + sorted(df["Cod2"].dropna().unique().tolist())
    anos    = ["Todos"] + sorted(df["Ano coleta"].dropna().astype(int).unique().tolist(), reverse=True)
    meses_disp = ["Todos"] + [MESES_F[m] for m in sorted(df["Mês coleta"].dropna().astype(int).unique().tolist())]

    fc1, fc2, fc3 = st.columns([2, 1, 1])
    with fc1: setor_sel = st.selectbox("🏭  SETOR (Cod2)", setores, label_visibility="visible")
    with fc2: ano_sel   = st.selectbox("📅  ANO",          anos,    label_visibility="visible")
    with fc3: mes_sel   = st.selectbox("📆  MÊS",          meses_disp, label_visibility="visible")

    dff = df.copy()
    if setor_sel != "Todos": dff = dff[dff["Cod2"] == setor_sel]
    if ano_sel   != "Todos": dff = dff[dff["Ano coleta"] == int(ano_sel)]
    if mes_sel   != "Todos":
        mn = {v: k for k, v in MESES_F.items()}[mes_sel]
        dff = dff[dff["Mês coleta"] == mn]

    total = len(dff)
    norm  = len(dff[dff["Status"]=="Normal"])
    alt   = len(dff[dff["Status"]=="Alerta"])
    alm   = len(dff[dff["Status"]=="Alarme"])
    pn = f"{norm/total*100:.0f}%" if total else "0%"
    pa = f"{alt/total*100:.0f}%"  if total else "0%"
    pm = f"{alm/total*100:.0f}%"  if total else "0%"

    # ── KPIs (Malha 2:1:1 para Alinhamento Perfeito) ───────────────────────────
    kc1, kc2, kc3 = st.columns([2, 1, 1])
    
    with kc1:
        # Coloca o Total e o Normal dividindo o mesmo espaço (equivalente à coluna 2)
        st.markdown(f"""
        <div style="display:flex; gap:1rem; width:100%;">
          <div class="kpi-tot kpi" style="flex:1;">
            <div class="kv">{total}</div>
            <div class="kl">TOTAL DE ATIVOS</div>
            <div class="ks">laudos no período</div>
          </div>
          <div class="kpi-nor kpi" style="flex:1;">
            <div class="kv">{norm}</div>
            <div class="kl">NORMAL</div>
            <div class="ks">{pn} da frota</div>
          </div>
        </div>""", unsafe_allow_html=True)

    with kc2:
        st.markdown(f"""
        <div class="kpi-alt kpi" style="width:100%;">
          <div class="kv">{alt}</div>
          <div class="kl">ALERTA</div>
          <div class="ks">{pa} da frota</div>
        </div>""", unsafe_allow_html=True)

    with kc3:
        st.markdown(f"""
        <div class="kpi-alm kpi" style="width:100%;">
          <div class="kv">{alm}</div>
          <div class="kl">ALARME</div>
          <div class="ks">{pm} da frota</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

    # ── Gráficos (Malha 2:1:1 para Alinhamento Perfeito) ───────────────────────
    gc1, gc2, gc3 = st.columns([2, 1, 1])
    COR = {"Normal":"#639922","Alerta":"#BA7517","Alarme":"#E24B4A"}
    LAYOUT = dict(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                  font=dict(color="#8AAABB",family="Barlow"),
                  margin=dict(l=0,r=0,t=4,b=30))

    with gc1:
        st.markdown('<div class="sc-top"><div class="st">Status por Setor (Cod2)</div></div>', unsafe_allow_html=True)
        if not dff.empty:
            db = dff.groupby(["Cod2","Status"]).size().reset_index(name="n")
            db["tot"] = db.groupby("Cod2")["n"].transform("sum")
            db["pct"] = (db["n"] / db["tot"] * 100).round(1)
            fig = px.bar(db, x="pct", y="Cod2", color="Status", orientation="h",
                         barmode="stack", color_discrete_map=COR, text="pct",
                         category_orders={"Status":["Normal","Alerta","Alarme"]})
            fig.update_traces(texttemplate="%{text:.0f}%", textposition="inside",
                              textfont_size=10, textfont_color="white")
            fig.update_layout(**LAYOUT, height=200,
                xaxis=dict(showgrid=False,showticklabels=False,range=[0,100]),
                yaxis=dict(showgrid=False,tickfont=dict(color="white",size=12,family="Barlow Condensed")),
                legend=dict(orientation="h",y=-0.18,x=0,font=dict(color="#8AAABB",size=10),
                            bgcolor="rgba(0,0,0,0)"),
                bargap=0.3)
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})

    with gc2:
        st.markdown('<div class="sc-top"><div class="st">Distribuição</div></div>', unsafe_allow_html=True)
        if not dff.empty and total > 0:
            dp = dff["Status"].value_counts().reset_index()
            dp.columns = ["Status","n"]
            fig2 = go.Figure(go.Pie(
                labels=dp["Status"], values=dp["n"], hole=0.55,
                marker_colors=[COR.get(s,"#666") for s in dp["Status"]],
                textinfo="percent", textfont_size=11, textfont_color="white",
                hovertemplate="%{label}: %{value}<extra></extra>"
            ))
            fig2.add_annotation(text=f"<b>{total}</b>", x=0.5, y=0.5,
                font=dict(size=24,color="white",family="Barlow Condensed"), showarrow=False)
            fig2.update_layout(**LAYOUT, height=200,
                showlegend=True,
                legend=dict(orientation="v",x=0,y=0.5,
                            font=dict(color="#8AAABB",size=9),
                            bgcolor="rgba(0,0,0,0)"))
            st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar":False})

    with gc3:
        st.markdown('<div class="sc-top"><div class="st">Coletas por Mês</div></div>', unsafe_allow_html=True)
        if not dff.empty:
            dm = dff.groupby("Mês coleta").size().reset_index(name="n").sort_values("Mês coleta")
            dm["label"] = dm["Mês coleta"].map(lambda x: MESES.get(int(x),"") if pd.notna(x) else "")
            fig3 = go.Figure(go.Bar(
                x=dm["label"], y=dm["n"], marker_color="#2E75B6",
                text=dm["n"], textposition="outside",
                textfont=dict(color="#8AAABB",size=10),
                hovertemplate="%{x}: %{y} laudos<extra></extra>"
            ))
            fig3.update_layout(**LAYOUT, height=200,
                xaxis=dict(showgrid=False,tickfont=dict(color="#8AAABB",size=10)),
                yaxis=dict(showgrid=False,showticklabels=False))
            st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar":False})

    # ── Tabela ─────────────────────────────────────────────────────────────────
    st.markdown('<div class="sc"><div class="st">Equipamentos com Desvio — Ação Necessária</div>', unsafe_allow_html=True)

    dtab = dff[dff["Status"].isin(["Alarme","Alerta"])].sort_values("Status")
    if dtab.empty:
        st.markdown('<p style="color:#3A7D44;text-align:center;padding:16px">✅ Nenhum equipamento com desvio no período.</p>', unsafe_allow_html=True)
    else:
        rows = ""
        for _, r in dtab.iterrows():
            st_val   = r.get("Status","")
            bcls     = "ba" if st_val=="Alarme" else "bl"
            ferro    = r.get("Ferro (ppm)")
            agua     = r.get("Água (ppm)")
            visc     = r.get("Visc 40 (cSt)")
            fc = "td-red" if ferro and ferro>30 else ("td-yel" if ferro and ferro>15 else "")
            ac = "td-red" if agua  and agua>800  else ("td-yel" if agua  and agua>200  else "")
            vc = "td-yel" if visc else ""
            rows += f"""<tr>
              <td class="td-set">{r.get('Cod2','')}</td>
              <td style="color:#5A8AAA;font-size:11px">{r.get('Cod3','')}</td>
              <td class="td-eq">{str(r.get('Equipamento Descrição 1','') or '')}</td>
              <td><span class="badge {bcls}">{st_val}</span></td>
              <td class="td-par">{str(r.get('Parâmetros em Alarme','') or '—')}</td>
              <td class="td-obs">{str(r.get('Observações','') or '')[:55]}</td>
              <td class="{fc}">{ferro if ferro is not None else '—'}</td>
              <td class="{ac}">{agua  if agua  is not None else '—'}</td>
              <td class="{vc}">{visc  if visc  is not None else '—'}</td>
              <td class="td-data">{str(r.get('Data de coleta','') or '')[:10]}</td>
            </tr>"""
        st.markdown(f"""
        <div style="overflow:hidden;border-radius:8px">
        <table class="tbl"><thead><tr>
          <th>Setor</th><th>Área</th><th>Equipamento</th><th>Status</th>
          <th>Parâmetro</th><th>Observações</th>
          <th>Ferro (ppm)</th><th>Água (ppm)</th><th>Visc 40</th><th>Coleta</th>
        </tr></thead><tbody>{rows}</tbody></table></div>""", unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    # Rodapé
    st.markdown(f"""
    <div style="text-align:center;margin-top:10px;font-size:9px;color:rgba(255,255,255,0.2);letter-spacing:.4px">
      SKF TruVu 360 · Gerdau Charqueadas · Engenharia de Manutenção &nbsp;·&nbsp;
      Desenvolvido por Douglas Brum &nbsp;·&nbsp; {datetime.now().strftime('%d/%m/%Y')}
    </div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# ROTEADOR
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.get("processado"):
    render_dashboard(st.session_state["laudos"])
else:
    render_upload()
