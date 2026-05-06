import streamlit as st
import zipfile, re, os, base64, json
from io import BytesIO
from datetime import datetime

try:
    import pdfplumber
except ImportError:
    st.error("Instale pdfplumber"); st.stop()
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    st.error("Instale openpyxl"); st.stop()
try:
    import plotly.express as px
    import plotly.graph_objects as go
    PLOTLY_OK = True
except ImportError:
    PLOTLY_OK = False
try:
    import pandas as pd
    PANDAS_OK = True
except ImportError:
    PANDAS_OK = False

st.set_page_config(
    page_title="Análise de Óleo SKF — Gerdau",
    page_icon="🛢️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@300;400;500;600;700&family=Barlow+Condensed:wght@600;700&display=swap');
html,body,[class*="css"]{font-family:'Barlow',sans-serif!important}
.stApp{background:linear-gradient(135deg,#0D2137 0%,#1A3A5C 100%);min-height:100vh}
#MainMenu,footer,header{visibility:hidden}
.block-container{padding-top:1rem;padding-bottom:1rem;max-width:100%!important}

/* Header */
.db-header{background:linear-gradient(135deg,#1F4E79,#2E75B6);border-radius:12px;
  padding:18px 28px;display:flex;align-items:center;justify-content:space-between;
  margin-bottom:12px;box-shadow:0 8px 32px rgba(0,0,0,0.3)}
.db-logo{width:46px;height:46px;background:rgba(255,255,255,0.15);border-radius:10px;
  display:flex;align-items:center;justify-content:center;font-family:'Barlow Condensed',sans-serif;
  font-size:15px;font-weight:700;color:white;border:1px solid rgba(255,255,255,0.2);flex-shrink:0}
.db-title{font-family:'Barlow Condensed',sans-serif;font-size:22px;font-weight:700;
  color:white;letter-spacing:.5px;line-height:1.1}
.db-sub{font-size:11px;color:rgba(255,255,255,0.6);margin-top:2px}
.db-badge{font-size:10px;font-weight:600;background:rgba(46,117,182,0.3);
  color:#BDD7EE;padding:3px 12px;border-radius:20px;border:1px solid rgba(46,117,182,0.5)}
.db-credit{font-size:9px;color:rgba(255,255,255,0.3);text-align:right;margin-top:4px;letter-spacing:.3px}

/* KPI cards */
.kpi-card{border-radius:10px;padding:16px 20px;text-align:center;border:1px solid transparent;position:relative;overflow:hidden}
.kpi-card::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;border-radius:10px 10px 0 0}
.kpi-total{background:#162E4A;border-color:#2A4A6A}.kpi-total::before{background:#2E75B6}
.kpi-normal{background:#0E2B16;border-color:#1A4D22}.kpi-normal::before{background:#4CAF50}
.kpi-alerta{background:#2B1E05;border-color:#4A3210}.kpi-alerta::before{background:#F59E0B}
.kpi-alarme{background:#2B0A0A;border-color:#4A1A1A}.kpi-alarme::before{background:#EF4444}
.kpi-val{font-family:'Barlow Condensed',sans-serif;font-size:48px;font-weight:700;line-height:1}
.kpi-total .kpi-val{color:white}.kpi-normal .kpi-val{color:#6BCB77}
.kpi-alerta .kpi-val{color:#F59E0B}.kpi-alarme .kpi-val{color:#EF4444}
.kpi-lbl{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:1.2px;margin-top:4px}
.kpi-total .kpi-lbl{color:#5A7A9A}.kpi-normal .kpi-lbl{color:#3A7D44}
.kpi-alerta .kpi-lbl{color:#8B6010}.kpi-alarme .kpi-lbl{color:#8B2020}
.kpi-sub{font-size:11px;margin-top:2px}
.kpi-total .kpi-sub{color:#3A5A7A}.kpi-normal .kpi-sub{color:#2A5A30}
.kpi-alerta .kpi-sub{color:#6A4A08}.kpi-alarme .kpi-sub{color:#6A1010}

/* Cards de seção */
.section-card{background:#112035;border:1px solid #1A3A5C;border-radius:10px;padding:16px 18px;margin-bottom:10px}
.section-title{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:1px;
  color:#BDD7EE;display:flex;align-items:center;gap:6px;margin-bottom:10px}
.section-title::before{content:'';display:inline-block;width:3px;height:14px;
  border-radius:2px;background:#2E75B6}

/* Tabela */
.tbl-wrap{overflow:hidden;border-radius:8px}
.tbl{width:100%;border-collapse:collapse;font-size:11px}
.tbl thead tr{background:#0D2137}
.tbl thead th{padding:9px 12px;text-align:left;font-size:9px;font-weight:600;
  text-transform:uppercase;letter-spacing:.8px;color:#3A6A9A;border-bottom:1px solid #1A3A5C}
.tbl tbody tr{border-bottom:1px solid #152840;transition:background .15s}
.tbl tbody tr:hover{background:#162E4A}
.tbl tbody td{padding:9px 12px;color:#8AAABB}
.td-eq{color:white!important;font-weight:500!important;font-size:12px!important}
.td-set{font-family:'Barlow Condensed',sans-serif;font-size:13px;font-weight:700;color:#2E75B6!important}
.td-par{color:#C8A870!important}
.badge{display:inline-block;font-size:9px;font-weight:700;padding:3px 10px;
  border-radius:20px;text-transform:uppercase;letter-spacing:.5px}
.b-alarme{background:rgba(239,68,68,.15);color:#EF4444;border:1px solid rgba(239,68,68,.3)}
.b-alerta{background:rgba(245,158,11,.15);color:#F59E0B;border:1px solid rgba(245,158,11,.3)}
.b-normal{background:rgba(76,175,80,.15);color:#4CAF50;border:1px solid rgba(76,175,80,.3)}
.td-obs{font-size:10px;color:#8A6A3A!important;font-style:italic}
.td-data{font-size:10px;color:#3A6A9A!important}
.td-alert{color:#EF4444!important;font-weight:600!important}
.td-warn{color:#F59E0B!important;font-weight:600!important}

/* Upload card */
.upload-card{background:white;border-radius:14px;max-width:680px;margin:0 auto;
  box-shadow:0 16px 48px rgba(0,0,0,0.25);overflow:hidden}
.upload-header{background:linear-gradient(135deg,#1F4E79,#2E75B6);padding:28px 32px;color:white}
.upload-body{padding:28px 32px}
.footer-note{text-align:center;font-size:10px;color:rgba(255,255,255,0.3);margin-top:12px}

/* Filtros */
div[data-testid="stSelectbox"] label{color:#BDD7EE!important;font-size:10px!important;
  text-transform:uppercase;letter-spacing:.8px}
div[data-testid="stMultiSelect"] label{color:#BDD7EE!important;font-size:10px!important;
  text-transform:uppercase;letter-spacing:.8px}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# LÓGICA DE EXTRAÇÃO
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
    s=str(v).strip().replace(",",".")
    try: return float(s)
    except: return s if s else None

def detectar_status(pdf_bytes):
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        full="".join(p.extract_text() or "" for p in pdf.pages)
        if "Limites acionados" not in full: return "Normal"
        for img in pdf.pages[0].images:
            if img.get("name")=="Im2":
                return "Alarme" if img["height"]>=50 else "Alerta"
    return "Alerta"

def split_ativo(raw):
    p=[x.strip() for x in raw.split("→")]
    s=p[0].strip().split("-")
    c={f"Cod{i}":s[i-1] if len(s)>=i else "" for i in range(1,6)}
    c["Equipamento Descrição 1"]=p[1] if len(p)>1 else ""
    c["Equipamento Descrição 2"]=p[2] if len(p)>2 else ""
    return c

def split_data(s):
    if not s: return None,None,None,None
    m=re.match(r"(\d{2})/(\d{2})/(\d{4}),?\s*(\d{2}:\d{2}:\d{2})?",str(s).strip())
    if not m: return None,None,None,None
    return int(m.group(1)),int(m.group(2)),int(m.group(3)),m.group(4)

def bloco(t,ini,fim):
    m=re.search(rf"{re.escape(ini)}\s*\n(.*?)(?={re.escape(fim)})",t,re.DOTALL)
    return m.group(1).strip() if m else ""

def extrair_laudo(nome,pdf_bytes):
    dados={"Arquivo":nome}
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        pags=[p.extract_text() or "" for p in pdf.pages]
    p1=pags[0]; texto="\n".join(pags)
    dados["Status"]=detectar_status(pdf_bytes)
    m=re.search(r"Instalação:\s*(.+?)(?:\s{2,}|\n)",p1)
    dados["Instalação"]=m.group(1).strip() if m else None
    m=re.search(r"Localização:\s*(.+)",p1)
    loc=m.group(1).strip() if m else ""
    if loc.endswith("-"):
        m2=re.search(r"Localização:.*\n(.+?)(?:\n|$)",p1)
        if m2: loc=loc+" "+m2.group(1).strip()
    dados["Localização"]=loc
    m=re.search(r"Ativo:\s*(.+?)(?=Número de série:|\n\n)",p1,re.DOTALL)
    raw=re.sub(r"\s+"," ",m.group(1)).strip() if m else ""
    dados["Ativo (completo)"]=raw; dados.update(split_ativo(raw))
    m=re.search(r"Tipo de componente:\s*(.+?)(?=Operação de ativo:|\n)",p1)
    dados["Tipo de componente"]=m.group(1).strip() if m else None
    m=re.search(r"Óleo:\s*(.+?)(?=Operação de óleo:|\n)",p1)
    dados["Óleo"]=m.group(1).strip() if m else None
    m=re.search(r"data de coleta:\s*([\d/,: ]+)",p1)
    ds=m.group(1).strip() if m else None
    dados["Data de coleta"]=ds
    d,me,a,h=split_data(ds)
    dados["Dia coleta"]=d;dados["Mês coleta"]=me;dados["Ano coleta"]=a;dados["Hora coleta"]=h
    dados["Observações"]=bloco(p1,"Observações:","Diagnósticos:")
    dados["Diagnósticos"]=bloco(p1,"Diagnósticos:","Ações:")
    dados["Ações"]=bloco(p1,"Ações:","Recomendações adicionais:")
    dados["Recomendações adicionais"]=bloco(p1,"Recomendações adicionais:","Teste realizado por:")
    m=re.search(r"Teste realizado por:\s*(.+?)(?=Lançado por:|\n)",p1)
    dados["Teste realizado por"]=m.group(1).strip() if m else None
    m=re.search(r"Lançado por:\s*(.+?)(?=_{3,}|\n)",p1)
    dados["Lançado por"]=m.group(1).strip() if m else None
    m=re.search(r"ID de amostra\s+([\w-]+)",p1)
    dados["ID de amostra"]=m.group(1).strip() if m else None
    m=re.search(r"Data de teste\s+(\d{2}/\d{2}/\d{4},\s*\d{2}:\d{2}:\d{2})",p1)
    dt=m.group(1).strip() if m else None
    dados["Data de teste"]=dt
    d,me,a,_=split_data(dt)
    dados["Dia teste"]=d;dados["Mês teste"]=me;dados["Ano teste"]=a
    m=re.search(r"Data de lançamento\s+(\d{2}/\d{2}/\d{4},\s*\d{2}:\d{2}:\d{2})",p1)
    dados["Data de lançamento"]=m.group(1).strip() if m else None
    linhas=texto.split("\n"); pm={}
    for linha in linhas:
        ls=linha.strip()
        for raw,_ in PARAMETROS:
            if raw not in pm and (ls.startswith(raw+" ") or ls==raw):
                parts=ls.split(); nw=len(raw.split())
                if len(parts)>nw:
                    v=parts[nw]
                    if re.match(r"^[\d\-]",v) or v in ("Normal","Desconhecido"):
                        pm[raw]=v
                break
    for raw,col in PARAMETROS:
        dados[col]=to_num(pm.get(raw))
    m=re.search(r"Limites acionados \(amostra atual\)\n(.+?)(?=TruVu|Relatório)",texto,re.DOTALL)
    if m:
        pa=[]
        for l in m.group(1).strip().split("\n"):
            pm2=re.match(r"^\d+\s+(.+?)\s+(?:Absoluto|Desvio)",l.strip())
            if pm2: pa.append(pm2.group(1).strip())
        dados["Parâmetros em Alarme"]="; ".join(pa)
    else:
        dados["Parâmetros em Alarme"]=""
    return dados

def gerar_excel(lista):
    wb=openpyxl.Workbook(); ws=wb.active; ws.title="Laudos"
    colunas=list(lista[0].keys())
    cor_st={"Normal":"C6EFCE","Alerta":"FFEB9C","Alarme":"FFC7CE"}
    hf=PatternFill("solid",start_color="1F4E79")
    hfont=Font(bold=True,color="FFFFFF",name="Arial",size=9)
    borda=Border(left=Side(style="thin"),right=Side(style="thin"),
                 top=Side(style="thin"),bottom=Side(style="thin"))
    for ci,col in enumerate(colunas,1):
        c=ws.cell(row=1,column=ci,value=col)
        c.fill=hf;c.font=hfont;c.border=borda
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    ws.row_dimensions[1].height=36
    for ri,linha in enumerate(lista,2):
        st_val=linha.get("Status","Normal")
        fb=PatternFill("solid",start_color="EBF3FB" if ri%2==0 else "FFFFFF")
        for ci,col in enumerate(colunas,1):
            c=ws.cell(row=ri,column=ci,value=linha.get(col))
            c.font=Font(name="Arial",size=9);c.border=borda
            c.alignment=Alignment(horizontal="center",vertical="center")
            if col=="Status":
                c.fill=PatternFill("solid",start_color=cor_st.get(st_val,"FFFFFF"))
                c.font=Font(name="Arial",size=9,bold=True)
            else: c.fill=fb
    larg={"Arquivo":35,"Status":10,"Localização":28,"Ativo (completo)":45,
          "Cod1":7,"Cod2":7,"Cod3":8,"Cod4":9,"Cod5":7,
          "Equipamento Descrição 1":26,"Equipamento Descrição 2":26,
          "Tipo de componente":28,"Óleo":24,"Data de coleta":20,
          "Dia coleta":8,"Mês coleta":8,"Ano coleta":8,"Hora coleta":10,
          "Observações":40,"Diagnósticos":50,"Ações":50,"Parâmetros em Alarme":30}
    for ci,col in enumerate(colunas,1):
        ws.column_dimensions[get_column_letter(ci)].width=larg.get(col,14)
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
    buf=BytesIO(); wb.save(buf); buf.seek(0)
    return buf.read()

# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD INTERATIVO
# ══════════════════════════════════════════════════════════════════════════════
def render_dashboard(lista):
    if not PANDAS_OK or not PLOTLY_OK:
        st.error("Instale pandas e plotly: pip install pandas plotly")
        return

    df = pd.DataFrame(lista)

    # ── Header ────────────────────────────────────────────────────────────────
    st.markdown(f"""
    <div class="db-header">
      <div style="display:flex;align-items:center;gap:14px">
        <div class="db-logo">SKF</div>
        <div>
          <div class="db-title">MONITORAMENTO DE ANÁLISE DE ÓLEO</div>
          <div class="db-sub">Gerdau Charqueadas · Engenharia de Manutenção · TruVu 360</div>
        </div>
      </div>
      <div style="text-align:right">
        <div class="db-badge">● AO VIVO</div>
        <div style="font-size:11px;color:#BDD7EE;margin-top:4px">
          Última coleta: {df['Data de coleta'].dropna().iloc[-1] if not df['Data de coleta'].dropna().empty else '—'} &nbsp;|&nbsp; {datetime.now().strftime('%d/%m/%Y %H:%M')}
        </div>
        <div class="db-credit">Desenvolvido por Douglas Brum · Gerdau Charqueadas</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Filtros ────────────────────────────────────────────────────────────────
    col1, col2, col3 = st.columns([2, 1, 1])
    setores = ["Todos"] + sorted(df["Cod2"].dropna().unique().tolist())
    anos    = ["Todos"] + sorted(df["Ano coleta"].dropna().astype(int).unique().tolist(), reverse=True)
    meses_map = {1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",6:"Junho",
                 7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"}
    meses_disp = ["Todos"] + [meses_map[m] for m in sorted(df["Mês coleta"].dropna().astype(int).unique().tolist())]

    with col1:
        setor_sel = st.selectbox("🏭  SETOR (Cod2)", setores)
    with col2:
        ano_sel = st.selectbox("📅  ANO", anos)
    with col3:
        mes_sel = st.selectbox("📆  MÊS", meses_disp)

    # Aplicar filtros
    dff = df.copy()
    if setor_sel != "Todos":
        dff = dff[dff["Cod2"] == setor_sel]
    if ano_sel != "Todos":
        dff = dff[dff["Ano coleta"] == int(ano_sel)]
    if mes_sel != "Todos":
        mes_num = {v:k for k,v in meses_map.items()}[mes_sel]
        dff = dff[dff["Mês coleta"] == mes_num]

    total  = len(dff)
    normal = len(dff[dff["Status"]=="Normal"])
    alerta = len(dff[dff["Status"]=="Alerta"])
    alarme = len(dff[dff["Status"]=="Alarme"])
    pct_n  = f"{normal/total*100:.0f}%" if total else "0%"
    pct_al = f"{alerta/total*100:.0f}%" if total else "0%"
    pct_ar = f"{alarme/total*100:.0f}%" if total else "0%"

    # ── KPIs ──────────────────────────────────────────────────────────────────
    k1,k2,k3,k4 = st.columns(4)
    with k1:
        st.markdown(f"""<div class="kpi-card kpi-total">
          <div class="kpi-val">{total}</div>
          <div class="kpi-lbl">Total de Ativos</div>
          <div class="kpi-sub">laudos no período</div>
        </div>""", unsafe_allow_html=True)
    with k2:
        st.markdown(f"""<div class="kpi-card kpi-normal">
          <div class="kpi-val">{normal}</div>
          <div class="kpi-lbl">Normal</div>
          <div class="kpi-sub">{pct_n} da frota</div>
        </div>""", unsafe_allow_html=True)
    with k3:
        st.markdown(f"""<div class="kpi-card kpi-alerta">
          <div class="kpi-val">{alerta}</div>
          <div class="kpi-lbl">Alerta</div>
          <div class="kpi-sub">{pct_al} da frota</div>
        </div>""", unsafe_allow_html=True)
    with k4:
        st.markdown(f"""<div class="kpi-card kpi-alarme">
          <div class="kpi-val">{alarme}</div>
          <div class="kpi-lbl">Alarme</div>
          <div class="kpi-sub">{pct_ar} da frota</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    # ── Gráficos principais ────────────────────────────────────────────────────
    g1, g2, g3 = st.columns([3, 1.2, 1.2])

    with g1:
        st.markdown('<div class="section-card"><div class="section-title">Status por Setor (Cod2)</div>', unsafe_allow_html=True)
        if not dff.empty:
            df_bar = dff.groupby(["Cod2","Status"]).size().reset_index(name="count")
            df_bar["total"] = df_bar.groupby("Cod2")["count"].transform("sum")
            df_bar["pct"] = (df_bar["count"] / df_bar["total"] * 100).round(1)
            cor_map = {"Normal":"#639922","Alerta":"#BA7517","Alarme":"#E24B4A"}
            ordem_status = ["Normal","Alerta","Alarme"]
            fig = px.bar(df_bar, x="pct", y="Cod2", color="Status",
                         orientation="h", barmode="stack",
                         color_discrete_map=cor_map,
                         category_orders={"Status":ordem_status},
                         text="pct",
                         labels={"pct":"%","Cod2":"Setor","count":"Qtd"})
            fig.update_traces(texttemplate="%{text:.0f}%", textposition="inside",
                              textfont_size=10, textfont_color="white")
            fig.update_layout(
                plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#8AAABB", family="Barlow"),
                xaxis=dict(showgrid=False, showticklabels=False, range=[0,100]),
                yaxis=dict(showgrid=False, tickfont=dict(color="white", size=12, family="Barlow Condensed")),
                legend=dict(orientation="h", y=-0.15, x=0,
                            font=dict(color="#8AAABB", size=10),
                            bgcolor="rgba(0,0,0,0)"),
                margin=dict(l=0, r=10, t=0, b=30),
                height=220, showlegend=True,
                bargap=0.25
            )
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})
        st.markdown('</div>', unsafe_allow_html=True)

    with g2:
        st.markdown('<div class="section-card"><div class="section-title">Distribuição</div>', unsafe_allow_html=True)
        if not dff.empty and total > 0:
            df_pizza = dff["Status"].value_counts().reset_index()
            df_pizza.columns = ["Status","count"]
            cor_map = {"Normal":"#639922","Alerta":"#BA7517","Alarme":"#E24B4A"}
            fig2 = go.Figure(go.Pie(
                labels=df_pizza["Status"], values=df_pizza["count"],
                hole=0.55,
                marker_colors=[cor_map.get(s,"#666") for s in df_pizza["Status"]],
                textinfo="percent", textfont_size=11,
                textfont_color="white",
                hovertemplate="%{label}: %{value} (%{percent})<extra></extra>"
            ))
            fig2.add_annotation(text=f"<b>{total}</b>", x=0.5, y=0.5,
                                font=dict(size=26, color="white", family="Barlow Condensed"),
                                showarrow=False)
            fig2.update_layout(
                plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                showlegend=True,
                legend=dict(orientation="v", x=0, y=0.5,
                            font=dict(color="#8AAABB", size=9),
                            bgcolor="rgba(0,0,0,0)"),
                margin=dict(l=0, r=0, t=0, b=0), height=220
            )
            st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar":False})
        st.markdown('</div>', unsafe_allow_html=True)

    with g3:
        st.markdown('<div class="section-card"><div class="section-title">Coletas por Mês</div>', unsafe_allow_html=True)
        if not dff.empty:
            df_mes = dff.groupby("Mês coleta").size().reset_index(name="count")
            df_mes = df_mes.sort_values("Mês coleta")
            df_mes["mes_nome"] = df_mes["Mês coleta"].map(
                lambda x: meses_map.get(int(x),"")[:3] if pd.notna(x) else "")
            fig3 = go.Figure(go.Bar(
                x=df_mes["mes_nome"], y=df_mes["count"],
                marker_color="#2E75B6",
                text=df_mes["count"], textposition="outside",
                textfont=dict(color="#8AAABB", size=10),
                hovertemplate="%{x}: %{y} laudos<extra></extra>"
            ))
            fig3.update_layout(
                plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#8AAABB", family="Barlow"),
                xaxis=dict(showgrid=False, tickfont=dict(color="#8AAABB", size=9)),
                yaxis=dict(showgrid=False, showticklabels=False),
                margin=dict(l=0, r=0, t=10, b=0), height=220
            )
            st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar":False})
        st.markdown('</div>', unsafe_allow_html=True)

    # ── Tabela de desvios ──────────────────────────────────────────────────────
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Equipamentos com Desvio — Ação Necessária</div>', unsafe_allow_html=True)

    df_tab = dff[dff["Status"].isin(["Alarme","Alerta"])].copy()
    df_tab = df_tab.sort_values("Status", ascending=True)

    if df_tab.empty:
        st.markdown('<p style="color:#3A7D44;text-align:center;padding:20px">✅ Nenhum equipamento com desvio no período selecionado.</p>', unsafe_allow_html=True)
    else:
        rows_html = ""
        for _, row in df_tab.iterrows():
            st_val   = row.get("Status","")
            badge_cls= "b-alarme" if st_val=="Alarme" else "b-alerta"
            ferro    = row.get("Ferro (ppm)")
            agua     = row.get("Água (ppm)")
            visc     = row.get("Visc 40 (cSt)")
            ferro_cls= "td-alert" if ferro and ferro>30 else ("td-warn" if ferro and ferro>15 else "")
            agua_cls = "td-alert" if agua and agua>800 else ("td-warn" if agua and agua>200 else "")
            visc_cls = "td-warn" if visc else ""
            obs      = str(row.get("Observações","") or "")[:60]
            param    = str(row.get("Parâmetros em Alarme","") or "—")
            data     = str(row.get("Data de coleta","") or "")[:10]
            eq       = str(row.get("Equipamento Descrição 1","") or "")

            rows_html += f"""<tr>
              <td class="td-set">{row.get('Cod2','')}</td>
              <td style="color:#5A8AAA;font-size:11px">{row.get('Cod3','')}</td>
              <td class="td-eq">{eq}</td>
              <td><span class="badge {badge_cls}">{st_val}</span></td>
              <td class="td-par">{param}</td>
              <td class="td-obs">{obs}</td>
              <td class="{ferro_cls}">{ferro if ferro is not None else '—'}</td>
              <td class="{agua_cls}">{agua if agua is not None else '—'}</td>
              <td class="{visc_cls}">{visc if visc is not None else '—'}</td>
              <td class="td-data">{data}</td>
            </tr>"""

        st.markdown(f"""
        <div class="tbl-wrap">
          <table class="tbl">
            <thead><tr>
              <th>Setor</th><th>Área</th><th>Equipamento</th><th>Status</th>
              <th>Parâmetro</th><th>Observações</th>
              <th>Ferro (ppm)</th><th>Água (ppm)</th><th>Visc 40</th><th>Coleta</th>
            </tr></thead>
            <tbody>{rows_html}</tbody>
          </table>
        </div>""", unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    # ── Rodapé ─────────────────────────────────────────────────────────────────
    st.markdown(f"""
    <div style="text-align:center;margin-top:16px;font-size:9px;color:rgba(255,255,255,0.2);
                letter-spacing:.5px;padding-bottom:8px">
      SKF TruVu 360 · Gerdau Charqueadas · Engenharia de Manutenção
      &nbsp;·&nbsp; Desenvolvido por Douglas Brum
      &nbsp;·&nbsp; {datetime.now().strftime('%d/%m/%Y')}
    </div>
    """, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TELA DE UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
def render_upload():
    st.markdown("""
    <div style="display:flex;justify-content:center;align-items:center;min-height:80vh">
    <div class="upload-card">
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="upload-header">
      <div style="display:flex;align-items:center;gap:14px;margin-bottom:10px">
        <div class="db-logo">SKF</div>
        <div>
          <div class="db-title" style="font-size:20px">Extrator de Laudos SKF</div>
          <div class="db-sub">TruVu 360 · Gerdau Charqueadas · Eng. de Manutenção</div>
        </div>
      </div>
      <div style="font-size:13px;color:rgba(255,255,255,0.78);line-height:1.6">
        Faça o upload do ZIP com os laudos em PDF.<br>
        Gera Excel consolidado <b>e</b> painel interativo automaticamente.
      </div>
    </div>
    """, unsafe_allow_html=True)

    with st.container():
        st.markdown('<div class="upload-body">', unsafe_allow_html=True)
        uploaded = st.file_uploader("📦 Selecione o arquivo ZIP com os laudos",
                                     type=["zip"],
                                     help="Arquivo .zip enviado pelo laboratório SKF")

        if uploaded:
            st.success(f"✅ **{uploaded.name}** — {uploaded.size/1024/1024:.1f} MB")

            if st.button("⚙️ Processar Laudos", type="primary", use_container_width=True):
                with st.spinner("Processando laudos..."):
                    zf = zipfile.ZipFile(BytesIO(uploaded.read()))
                    pdfs = [(n, zf.read(n)) for n in zf.namelist()
                            if n.lower().endswith(".pdf") and not n.startswith("__MACOSX")]

                if not pdfs:
                    st.error("Nenhum PDF encontrado no ZIP.")
                    return

                progress = st.progress(0)
                lista, erros = [], []
                for i,(nome,pdf_bytes) in enumerate(sorted(pdfs),1):
                    progress.progress(i/len(pdfs), text=f"Processando {i}/{len(pdfs)}: {os.path.basename(nome)[:45]}")
                    try:
                        lista.append(extrair_laudo(os.path.basename(nome), pdf_bytes))
                    except Exception as e:
                        erros.append(str(e))

                progress.progress(1.0, text="Concluído!")

                # Salva na sessão
                st.session_state["laudos"] = lista
                st.session_state["processado"] = True
                st.session_state["nome_excel"] = f"LAUDOS_{datetime.now().strftime('%d-%m-%Y')}.xlsx"
                st.rerun()

        else:
            st.markdown("""
            <div style="background:rgba(31,78,121,0.15);border:1px solid rgba(46,117,182,0.2);
                        border-radius:10px;padding:20px;font-size:13px;color:rgba(0,0,0,0.65);line-height:1.8">
              📦 &nbsp;<b>1.</b> Receba o ZIP do laboratório<br>
              ⬆️ &nbsp;<b>2.</b> Faça o upload acima<br>
              ⚙️ &nbsp;<b>3.</b> Clique em Processar Laudos<br>
              📊 &nbsp;<b>4.</b> Veja o painel interativo e baixe o Excel
            </div>
            """, unsafe_allow_html=True)

        st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div></div>', unsafe_allow_html=True)
    st.markdown('<div class="footer-note">Desenvolvido por Douglas Brum · Gerdau Charqueadas · SKF TruVu 360</div>',
                unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TELA DE RESULTADOS (após processar)
# ══════════════════════════════════════════════════════════════════════════════
def render_resultados():
    lista = st.session_state["laudos"]
    nome_excel = st.session_state["nome_excel"]

    total  = len(lista)
    normal = sum(1 for d in lista if d.get("Status")=="Normal")
    alerta = sum(1 for d in lista if d.get("Status")=="Alerta")
    alarme = sum(1 for d in lista if d.get("Status")=="Alarme")

    # Barra superior com ações
    c1, c2, c3 = st.columns([2, 1, 1])
    with c1:
        st.markdown(f"""
        <div style="background:#112035;border:1px solid #1A3A5C;border-radius:10px;
                    padding:12px 18px;display:flex;align-items:center;gap:16px">
          <div style="font-family:'Barlow Condensed',sans-serif;font-size:18px;
                      font-weight:700;color:white">{total} laudos processados</div>
          <span style="background:rgba(76,175,80,.15);color:#6BCB77;border:1px solid rgba(76,175,80,.3);
                       font-size:10px;font-weight:700;padding:3px 10px;border-radius:20px">{normal} Normal</span>
          <span style="background:rgba(245,158,11,.15);color:#F59E0B;border:1px solid rgba(245,158,11,.3);
                       font-size:10px;font-weight:700;padding:3px 10px;border-radius:20px">{alerta} Alerta</span>
          <span style="background:rgba(239,68,68,.15);color:#EF4444;border:1px solid rgba(239,68,68,.3);
                       font-size:10px;font-weight:700;padding:3px 10px;border-radius:20px">{alarme} Alarme</span>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        excel_bytes = gerar_excel(lista)
        st.download_button("⬇️  Baixar Excel", data=excel_bytes, file_name=nome_excel,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True, type="primary")
    with c3:
        if st.button("↩  Novo ZIP", use_container_width=True):
            del st.session_state["laudos"]
            del st.session_state["processado"]
            st.rerun()

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    # Dashboard
    render_dashboard(lista)

# ══════════════════════════════════════════════════════════════════════════════
# ROTEADOR PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.get("processado"):
    render_resultados()
else:
    render_upload()
