import streamlit as st
import streamlit.components.v1 as components
import zipfile, re, os
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

st.set_page_config(
    page_title="Análise de Óleo SKF — Gerdau",
    page_icon="🛢️", layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@300;400;500;600;700&family=Barlow+Condensed:wght@600;700&display=swap');

html,body,[class*="css"]{font-family:'Barlow',sans-serif !important}
.stApp{background:linear-gradient(135deg,#0D2137 0%,#1A3A5C 100%) !important}

/* Oculta elementos nativos do Streamlit */
#MainMenu,footer,header{display:none !important}
[data-testid="stHeader"]{display:none !important}
[data-testid="stToolbar"]{display:none !important}
[data-testid="stDecoration"]{display:none !important}
[data-testid="stStatusWidget"]{display:none !important}
section[data-testid="stSidebar"]{display:none !important}

/* Container principal sem padding extra */
.block-container{
  padding:0.5rem 1.2rem 0.8rem 1.2rem !important;
  max-width:100% !important; margin-top:0 !important;
}

/* Reduz gap entre linhas do Streamlit */
div[data-testid="stVerticalBlock"]{gap:0.3rem !important}
div[data-testid="stHorizontalBlock"]{gap:0.5rem !important}

/* ── ACTION BAR ─────────────────────────────────────────── */
.act-bar{
  background:#0D2137; border:1px solid #1E3A5C;
  border-radius:8px; padding:8px 16px;
  display:flex; align-items:center; gap:10px;
}
.act-title{
  font-family:'Barlow Condensed',sans-serif;
  font-size:15px; font-weight:700; color:white; flex:1;
}
.bcnt{
  font-size:10px; font-weight:700;
  padding:2px 10px; border-radius:20px; white-space:nowrap;
}

/* ── HEADER ─────────────────────────────────────────────── */
.db-hdr{
  background:linear-gradient(135deg,#1F4E79,#2E75B6);
  border-radius:10px; padding:12px 20px;
  display:flex; align-items:center; justify-content:space-between;
  box-shadow:0 4px 20px rgba(0,0,0,0.3);
}
.db-logo{
  width:38px; height:38px; border-radius:9px;
  background:rgba(255,255,255,0.15); border:1px solid rgba(255,255,255,0.2);
  display:flex; align-items:center; justify-content:center;
  font-family:'Barlow Condensed',sans-serif;
  font-size:13px; font-weight:700; color:white; flex-shrink:0;
}
.db-ttl{
  font-family:'Barlow Condensed',sans-serif;
  font-size:18px; font-weight:700; color:white; letter-spacing:.4px;
}
.db-sub{font-size:9px; color:rgba(255,255,255,0.55); margin-top:1px}
.db-badge{
  font-size:9px; font-weight:600; color:#BDD7EE;
  background:rgba(46,117,182,0.3); padding:2px 9px;
  border-radius:20px; border:1px solid rgba(46,117,182,0.5);
}
.db-credit{font-size:8px; color:rgba(255,255,255,0.2); margin-top:3px; text-align:right}

/* ── FILTROS ─────────────────────────────────────────────── */
div[data-testid="stSelectbox"] label{
  color:#8AAABB !important; font-size:9px !important;
  text-transform:uppercase; letter-spacing:.8px;
  margin-bottom:1px !important;
}
div[data-testid="stSelectbox"] > div > div > div{
  background:#0D2137 !important; border-color:#2A4A6A !important;
  color:#C8D8E8 !important; min-height:34px !important; font-size:13px !important;
}

/* ── KPI CARDS ──────────────────────────────────────────── */
.kpi{
  border-radius:10px; padding:12px 14px; text-align:center;
  border:1px solid transparent; position:relative; overflow:hidden;
  min-height:90px;
}
.kpi::before{
  content:''; position:absolute; top:0; left:0; right:0;
  height:3px; border-radius:10px 10px 0 0;
}
.kpi-tot{background:#162E4A; border-color:#2A4A6A}
.kpi-tot::before{background:#2E75B6}
.kpi-nor{background:#0E2B16; border-color:#1A4D22}
.kpi-nor::before{background:#4CAF50}
.kpi-alt{background:#2B1E05; border-color:#4A3210}
.kpi-alt::before{background:#F59E0B}
.kpi-alm{background:#2B0A0A; border-color:#4A1A1A}
.kpi-alm::before{background:#EF4444}
.kv{font-family:'Barlow Condensed',sans-serif; font-size:40px; font-weight:700; line-height:1.1}
.kpi-tot .kv{color:white}       .kpi-nor .kv{color:#6BCB77}
.kpi-alt .kv{color:#F59E0B}     .kpi-alm .kv{color:#EF4444}
.kl{font-size:9px; font-weight:600; text-transform:uppercase; letter-spacing:1.2px; margin-top:2px}
.kpi-tot .kl{color:#5A7A9A}     .kpi-nor .kl{color:#3A7D44}
.kpi-alt .kl{color:#8B6010}     .kpi-alm .kl{color:#8B2020}
.ks{font-size:10px; margin-top:1px}
.kpi-tot .ks{color:#3A5A7A}     .kpi-nor .ks{color:#2A5A30}
.kpi-alt .ks{color:#6A4A08}     .kpi-alm .ks{color:#6A1010}

/* ── SECTION CARDS ──────────────────────────────────────── */
.sc{
  background:#112035; border:1px solid #1A3A5C;
  border-radius:10px; padding:10px 12px;
}
.sc-t{
  font-size:9px; font-weight:600; text-transform:uppercase;
  letter-spacing:1px; color:#BDD7EE;
  display:flex; align-items:center; gap:6px; margin-bottom:6px;
}
.sc-t::before{
  content:''; display:inline-block; width:3px; height:11px;
  border-radius:2px; background:#2E75B6;
}

/* ── TABELA ─────────────────────────────────────────────── */
.tbl{width:100%; border-collapse:collapse; font-size:11px}
.tbl thead tr{background:#0D2137}
.tbl thead th{
  padding:7px 10px; text-align:left; font-size:9px; font-weight:600;
  text-transform:uppercase; letter-spacing:.7px; color:#3A6A9A;
  border-bottom:1px solid #1A3A5C;
}
.tbl tbody tr{border-bottom:1px solid #152840}
.tbl tbody tr:hover{background:#162E4A}
.tbl tbody td{padding:7px 10px; color:#8AAABB}
.td-eq{color:white !important; font-weight:500 !important}
.td-set{font-family:'Barlow Condensed',sans-serif; font-size:13px; font-weight:700; color:#2E75B6 !important}
.td-par{color:#C8A870 !important}
.td-obs{font-size:10px; color:#8A6A3A !important; font-style:italic}
.td-dt{font-size:10px; color:#3A6A9A !important}
.td-r{color:#EF4444 !important; font-weight:600 !important}
.td-y{color:#F59E0B !important; font-weight:600 !important}
.badge{
  display:inline-block; font-size:9px; font-weight:700;
  padding:2px 8px; border-radius:20px;
  text-transform:uppercase; letter-spacing:.4px;
}
.ba{background:rgba(239,68,68,.15);  color:#EF4444; border:1px solid rgba(239,68,68,.3)}
.bl{background:rgba(245,158,11,.15); color:#F59E0B; border:1px solid rgba(245,158,11,.3)}
.bn{background:rgba(76,175,80,.15);  color:#4CAF50; border:1px solid rgba(76,175,80,.3)}
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
    dados["Observações"]              = bloco(p1,"Observações:","Diagnósticos:")
    dados["Diagnósticos"]             = bloco(p1,"Diagnósticos:","Ações:")
    dados["Ações"]                    = bloco(p1,"Ações:","Recomendações adicionais:")
    dados["Recomendações adicionais"] = bloco(p1,"Recomendações adicionais:","Teste realizado por:")
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
    cor_st = {"Normal":"C6EFCE","Alerta":"FFEB9C","Alarme":"FFC7CE"}
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
        sv = linha.get("Status","Normal")
        fb = PatternFill("solid", start_color="EBF3FB" if ri%2==0 else "FFFFFF")
        for ci, col in enumerate(colunas, 1):
            c = ws.cell(row=ri, column=ci, value=linha.get(col))
            c.font = Font(name="Arial", size=9); c.border = borda
            c.alignment = Alignment(horizontal="center", vertical="center")
            if col == "Status":
                c.fill = PatternFill("solid", start_color=cor_st.get(sv,"FFFFFF"))
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
    components.html("""<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700&family=Barlow+Condensed:wght@600;700&display=swap" rel="stylesheet">
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Barlow',sans-serif;background:linear-gradient(135deg,#0D2137,#1A3A5C);
  min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px}
.card{width:100%;max-width:520px;background:white;border-radius:16px;overflow:hidden;
  box-shadow:0 24px 80px rgba(0,0,0,0.5)}
.head{background:linear-gradient(135deg,#1F4E79,#2E75B6);padding:24px 28px}
.row{display:flex;align-items:center;gap:14px;margin-bottom:12px}
.logo{width:44px;height:44px;border-radius:10px;background:rgba(255,255,255,0.15);
  border:1px solid rgba(255,255,255,0.25);display:flex;align-items:center;justify-content:center;
  font-family:'Barlow Condensed',sans-serif;font-size:14px;font-weight:700;color:white}
.ttl{font-family:'Barlow Condensed',sans-serif;font-size:21px;font-weight:700;color:white}
.sub{font-size:11px;color:rgba(255,255,255,0.6);margin-top:2px}
.desc{font-size:13px;color:rgba(255,255,255,0.82);line-height:1.6}
.body{padding:22px 28px}
.steps{background:#F0F6FF;border:1px solid #C8DCEF;border-radius:10px;
  padding:14px 18px;font-size:13px;color:#1F4E79;line-height:2.2}
.foot{text-align:center;font-size:10px;color:rgba(255,255,255,0.3);
  padding:0 0 18px;letter-spacing:.4px}
</style></head>
<body><div class="card">
  <div class="head">
    <div class="row">
      <div class="logo" style="padding:0;overflow:hidden;"><img src="data:image/jpeg;base64,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" style="width:38px;height:38px;border-radius:9px;object-fit:cover;" alt="SKF"></div>
      <div><div class="ttl">Extrator de Laudos SKF</div>
        <div class="sub">TruVu 360 · Gerdau Charqueadas · Eng. de Manutenção</div>
      </div>
    </div>
    <div class="desc">Faça upload do ZIP com os laudos em PDF.<br>
      Gera <strong>Excel</strong> e <strong>painel interativo</strong> automaticamente.</div>
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
</div></body></html>""", height=420)

    _, col, _ = st.columns([1, 3, 1])
    with col:
        up = st.file_uploader("ZIP", type=["zip"], label_visibility="collapsed")
        if up:
            st.success(f"✅ **{up.name}** — {up.size/1024/1024:.1f} MB")
            if st.button("⚙️  Processar Laudos", type="primary", use_container_width=True):
                zf = zipfile.ZipFile(BytesIO(up.read()))
                pdfs = [(n, zf.read(n)) for n in zf.namelist()
                        if n.lower().endswith(".pdf") and not n.startswith("__MACOSX")]
                if not pdfs:
                    st.error("Nenhum PDF no ZIP."); return
                prog = st.progress(0)
                lista = []
                for i, (nome, pb) in enumerate(sorted(pdfs), 1):
                    prog.progress(i/len(pdfs), text=f"{i}/{len(pdfs)}: {os.path.basename(nome)[:50]}")
                    try: lista.append(extrair_laudo(os.path.basename(nome), pb))
                    except: pass
                prog.progress(1.0, text="Concluído!")
                st.session_state["laudos"] = lista
                st.session_state["processado"] = True
                st.session_state["excel"] = f"LAUDOS_{datetime.now().strftime('%d-%m-%Y')}.xlsx"
                st.rerun()
        else:
            st.markdown('<p style="text-align:center;color:rgba(255,255,255,0.35);font-size:12px">Arraste o ZIP aqui ou clique para selecionar</p>',
                        unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
def render_dashboard(lista):
    df = pd.DataFrame(lista)
    MESES  = {1:"Jan",2:"Fev",3:"Mar",4:"Abr",5:"Mai",6:"Jun",
              7:"Jul",8:"Ago",9:"Set",10:"Out",11:"Nov",12:"Dez"}
    MESES_F= {1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",6:"Junho",
              7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"}
    COR = {"Normal":"#639922","Alerta":"#BA7517","Alarme":"#E24B4A"}
    PLT = dict(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
               font=dict(color="#8AAABB",family="Barlow"),
               margin=dict(l=0,r=0,t=4,b=0))

    # totais globais
    tg = len(df)
    ng = len(df[df.Status=="Normal"])
    ag = len(df[df.Status=="Alerta"])
    mg = len(df[df.Status=="Alarme"])

    # ── LINHA 1: barra de ações ─────────────────────────────────────────────
    excel_bytes = gerar_excel(lista)
    r1a, r1b, r1c, r1d = st.columns([3.2, 1.15, 1.05, 1.2])
    with r1a:
        st.markdown(f"""<div class="act-bar">
          <span class="act-title">{tg} laudos processados</span>
          <span class="bcnt" style="background:rgba(76,175,80,.15);color:#6BCB77;border:1px solid rgba(76,175,80,.3)">{ng} Normal</span>
          <span class="bcnt" style="background:rgba(245,158,11,.15);color:#F59E0B;border:1px solid rgba(245,158,11,.3)">{ag} Alerta</span>
          <span class="bcnt" style="background:rgba(239,68,68,.15);color:#EF4444;border:1px solid rgba(239,68,68,.3)">{mg} Alarme</span>
        </div>""", unsafe_allow_html=True)
    with r1b:
        st.download_button("⬇️  Baixar Excel", data=excel_bytes,
            file_name=st.session_state["excel"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, type="primary")
    with r1c:
        if st.button("↩  Novo ZIP", use_container_width=True):
            for k in ["laudos","processado","excel"]: st.session_state.pop(k,None)
            st.rerun()
    with r1d:
        # Botão Apresentar — usa components.html para JS real via window.parent
        components.html("""
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{background:transparent}
button{
  width:100%;height:40px;
  background:linear-gradient(135deg,#1F4E79,#2E75B6);
  color:white;border:none;border-radius:8px;
  font-family:'Segoe UI',sans-serif;font-size:13px;font-weight:600;
  cursor:pointer;letter-spacing:.3px;
}
button:hover{opacity:.85}
</style>
<button onclick="go()">⛶ &nbsp;Apresentar</button>
<script>
function go(){
  var d=window.parent.document, e=d.documentElement;
  (e.requestFullscreen||e.webkitRequestFullscreen||e.mozRequestFullScreen||e.msRequestFullscreen).call(e);
  function clean(){
    ['header[data-testid="stHeader"]','[data-testid="stToolbar"]',
     '[data-testid="stDecoration"]','#MainMenu','footer'].forEach(function(s){
      var x=d.querySelector(s); if(x) x.style.display='none';
    });
    var bc=d.querySelector('.block-container');
    if(bc){bc.style.paddingTop='0.4rem';}
  }
  clean(); setTimeout(clean,400); setTimeout(clean,1000);
}
</script>""", height=44)

    # ── LINHA 2: header ─────────────────────────────────────────────────────
    ultima = df["Data de coleta"].dropna().iloc[-1] if not df["Data de coleta"].dropna().empty else "—"
    st.markdown(f"""<div class="db-hdr">
      <div style="display:flex;align-items:center;gap:11px">
        <div class="db-logo" style="padding:0;overflow:hidden;"><img src="data:image/jpeg;base64,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" style="width:38px;height:38px;border-radius:9px;object-fit:cover;" alt="SKF"></div>
        <div>
          <div class="db-ttl">MONITORAMENTO DE ANÁLISE DE ÓLEO</div>
          <div class="db-sub">Gerdau Charqueadas · Engenharia de Manutenção · SKF TruVu 360</div>
        </div>
      </div>
      <div style="text-align:right">
        <div class="db-badge">● AO VIVO</div>
        <div style="font-size:9px;color:#BDD7EE;margin-top:3px">
          Última coleta: {ultima} &nbsp;|&nbsp; {datetime.now().strftime('%d/%m/%Y %H:%M')}
        </div>
        <div class="db-credit">Desenvolvido por Douglas Brum · Gerdau Charqueadas</div>
      </div>
    </div>""", unsafe_allow_html=True)

    # ── LINHA 3: filtros ─────────────────────────────────────────────────────
    setores   = ["Todos"] + sorted(df["Cod2"].dropna().unique().tolist())
    anos      = ["Todos"] + sorted(df["Ano coleta"].dropna().astype(int).unique().tolist(), reverse=True)
    mes_opts  = ["Todos"] + [MESES_F[m] for m in sorted(df["Mês coleta"].dropna().astype(int).unique().tolist())]
    f1,f2,f3 = st.columns([2,1,1])
    with f1: setor_sel = st.selectbox("🏭 SETOR (Cod2)", setores)
    with f2: ano_sel   = st.selectbox("📅 ANO",          anos)
    with f3: mes_sel   = st.selectbox("📆 MÊS",          mes_opts)

    dff = df.copy()
    if setor_sel != "Todos": dff = dff[dff["Cod2"]==setor_sel]
    if ano_sel   != "Todos": dff = dff[dff["Ano coleta"]==int(ano_sel)]
    if mes_sel   != "Todos":
        mn = {v:k for k,v in MESES_F.items()}[mes_sel]
        dff = dff[dff["Mês coleta"]==mn]

    tot = len(dff)
    nor = len(dff[dff.Status=="Normal"])
    alt = len(dff[dff.Status=="Alerta"])
    alm = len(dff[dff.Status=="Alarme"])
    pn  = f"{nor/tot*100:.0f}%" if tot else "0%"
    pa  = f"{alt/tot*100:.0f}%" if tot else "0%"
    pm  = f"{alm/tot*100:.0f}%" if tot else "0%"

    # ── LINHA 4: KPIs — 4 colunas iguais ────────────────────────────────────
    k1,k2,k3,k4 = st.columns(4)
    for col, cls, val, lbl, sub in [
        (k1,"kpi-tot",tot,"TOTAL DE ATIVOS","laudos no período"),
        (k2,"kpi-nor",nor,"NORMAL",f"{pn} da frota"),
        (k3,"kpi-alt",alt,"ALERTA",f"{pa} da frota"),
        (k4,"kpi-alm",alm,"ALARME",f"{pm} da frota"),
    ]:
        with col:
            st.markdown(f"""<div class="kpi {cls}">
              <div class="kv">{val}</div>
              <div class="kl">{lbl}</div>
              <div class="ks">{sub}</div>
            </div>""", unsafe_allow_html=True)

    # ── LINHA 5: gráficos — proporção 3:1.2:1.2 ─────────────────────────────
    g1,g2,g3 = st.columns([3,1.2,1.2])

    with g1:
        st.markdown('<div class="sc"><div class="sc-t">Status por Setor (Cod2)</div>', unsafe_allow_html=True)
        if not dff.empty:
            db = dff.groupby(["Cod2","Status"]).size().reset_index(name="n")
            db["tot2"] = db.groupby("Cod2")["n"].transform("sum")
            db["pct"]  = (db["n"]/db["tot2"]*100).round(1)
            fig = px.bar(db, x="pct", y="Cod2", color="Status", orientation="h",
                         barmode="stack", color_discrete_map=COR, text="pct",
                         category_orders={"Status":["Normal","Alerta","Alarme"]})
            fig.update_traces(texttemplate="%{text:.0f}%", textposition="inside",
                              textfont_size=10, textfont_color="white")
            fig.update_layout(**PLT, height=180,
                xaxis=dict(showgrid=False,showticklabels=False,range=[0,100],
                           title=None),
                yaxis=dict(showgrid=False,title=None,
                           tickfont=dict(color="white",size=12,family="Barlow Condensed")),
                legend=dict(orientation="h",y=-0.15,x=0,
                            font=dict(color="#8AAABB",size=10),bgcolor="rgba(0,0,0,0)"),
                bargap=0.3)
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})
        st.markdown('</div>', unsafe_allow_html=True)

    with g2:
        st.markdown('<div class="sc"><div class="sc-t">Distribuição</div>', unsafe_allow_html=True)
        if not dff.empty and tot > 0:
            dp = dff["Status"].value_counts().reset_index()
            dp.columns = ["Status","n"]
            fig2 = go.Figure(go.Pie(
                labels=dp["Status"], values=dp["n"], hole=0.55,
                marker_colors=[COR.get(s,"#666") for s in dp["Status"]],
                textinfo="percent", textfont_size=11, textfont_color="white",
                hovertemplate="%{label}: %{value}<extra></extra>"
            ))
            fig2.add_annotation(text=f"<b>{tot}</b>", x=0.5, y=0.5,
                font=dict(size=22,color="white",family="Barlow Condensed"), showarrow=False)
            fig2.update_layout(**PLT, height=180,
                showlegend=True,
                legend=dict(orientation="v",x=-0.1,y=0.5,
                            font=dict(color="#8AAABB",size=9),bgcolor="rgba(0,0,0,0)"))
            st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar":False})
        st.markdown('</div>', unsafe_allow_html=True)

    with g3:
        st.markdown('<div class="sc"><div class="sc-t">Coletas por Mês</div>', unsafe_allow_html=True)
        if not dff.empty:
            dm = (dff.groupby("Mês coleta").size()
                    .reset_index(name="n").sort_values("Mês coleta"))
            dm["lbl"] = dm["Mês coleta"].map(
                lambda x: MESES.get(int(x),"") if pd.notna(x) else "")
            fig3 = go.Figure(go.Bar(
                x=dm["lbl"], y=dm["n"], marker_color="#2E75B6",
                text=dm["n"], textposition="outside",
                textfont=dict(color="#8AAABB",size=10),
                hovertemplate="%{x}: %{y}<extra></extra>"
            ))
            fig3.update_layout(**PLT, height=180,
                xaxis=dict(showgrid=False,tickfont=dict(color="#8AAABB",size=10)),
                yaxis=dict(showgrid=False,showticklabels=False))
            st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar":False})
        st.markdown('</div>', unsafe_allow_html=True)

    # ── LINHA 6: tabela ──────────────────────────────────────────────────────
    st.markdown('<div class="sc"><div class="sc-t">Equipamentos com Desvio — Ação Necessária</div>',
                unsafe_allow_html=True)

    dtab = dff[dff.Status.isin(["Alarme","Alerta"])].sort_values("Status")
    if dtab.empty:
        st.markdown('<p style="color:#3A7D44;text-align:center;padding:14px">✅ Nenhum desvio no período.</p>',
                    unsafe_allow_html=True)
    else:
        rows = ""
        for _, r in dtab.iterrows():
            sv   = r.get("Status","")
            bc   = "ba" if sv=="Alarme" else "bl"
            fe   = r.get("Ferro (ppm)");  fc = "td-r" if fe and fe>30 else ("td-y" if fe and fe>15 else "")
            ag   = r.get("Água (ppm)");   ac = "td-r" if ag and ag>800 else ("td-y" if ag and ag>200 else "")
            vi   = r.get("Visc 40 (cSt)"); vc = "td-y" if vi else ""
            rows += f"""<tr>
              <td class="td-set">{r.get('Cod2','')}</td>
              <td style="color:#5A8AAA;font-size:11px">{r.get('Cod3','')}</td>
              <td class="td-eq">{str(r.get('Equipamento Descrição 1','') or '')}</td>
              <td><span class="badge {bc}">{sv}</span></td>
              <td class="td-par">{str(r.get('Parâmetros em Alarme','') or '—')}</td>
              <td class="td-obs">{str(r.get('Observações','') or '')[:55]}</td>
              <td class="{fc}">{fe if fe is not None else '—'}</td>
              <td class="{ac}">{ag if ag is not None else '—'}</td>
              <td class="{vc}">{vi if vi is not None else '—'}</td>
              <td class="td-dt">{str(r.get('Data de coleta','') or '')[:10]}</td>
            </tr>"""
        st.markdown(f"""<div style="overflow:hidden;border-radius:8px">
        <table class="tbl"><thead><tr>
          <th>Setor</th><th>Área</th><th>Equipamento</th><th>Status</th>
          <th>Parâmetro</th><th>Observações</th>
          <th>Ferro(ppm)</th><th>Água(ppm)</th><th>Visc 40</th><th>Coleta</th>
        </tr></thead><tbody>{rows}</tbody></table></div>""", unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown(f"""<div style="text-align:center;margin-top:8px;font-size:9px;
      color:rgba(255,255,255,0.18);letter-spacing:.4px">
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
