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
[data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"] { display: none !important; }

.block-container {
    padding: 1rem 1.5rem !important; 
    max-width: 100% !important;
}
.stApp { background: linear-gradient(135deg,#0D2137 0%,#1A3A5C 100%) !important; }

/* ── KPI CARDS ── */
.kpi { 
    border-radius: 10px; padding: 15px 10px; text-align: center;
    border: 1px solid transparent; position: relative; 
    height: 110px; display: flex; flex-direction: column; justify-content: center;
}
.kpi::before { content:''; position:absolute; top:0; left:0; right:0; height:4px; border-radius:10px 10px 0 0; }
.kpi-tot { background:#162E4A; border-color:#2A4A6A; } .kpi-tot::before { background:#2E75B6; }
.kpi-nor { background:#0E2B16; border-color:#1A4D22; } .kpi-nor::before { background:#4CAF50; }
.kpi-alt { background:#2B1E05; border-color:#4A3210; } .kpi-alt::before { background:#F59E0B; }
.kpi-alm { background:#2B0A0A; border-color:#4A1A1A; } .kpi-alm::before { background:#EF4444; }

.kv { font-family:'Barlow Condensed',sans-serif; font-size:38px; font-weight:700; line-height:1; margin-bottom: 5px;}
.kpi-tot .kv { color:white; } .kpi-nor .kv { color:#6BCB77; } .kpi-alt .kv { color:#F59E0B; } .kpi-alm .kv { color:#EF4444; }
.kl { font-size:10px; font-weight:700; text-transform:uppercase; letter-spacing:1px; }
.ks { font-size:11px; margin-top:3px; opacity: 0.8; }

/* ── DASHBOARD HEADER ── */
.db-header {
    background: linear-gradient(135deg,#1F4E79,#2E75B6);
    border-radius: 10px; padding: 12px 20px; margin-bottom: 15px;
    display: flex; align-items: center; justify-content: space-between;
}
.db-title { font-family:'Barlow Condensed',sans-serif; font-size:20px; font-weight:700; color:white; }

/* ── GRAFICOS CONTAINERS ── */
.sc-box { 
    background:#112035; border:1px solid #1A3A5C; border-radius:10px; 
    padding: 0px; overflow: hidden; height: 260px;
}
.sc-title {
    background: rgba(255,255,255,0.03);
    padding: 10px 15px; border-bottom: 1px solid #1A3A5C;
    font-size:11px; font-weight:600; text-transform:uppercase; color:#BDD7EE;
    display:flex; align-items:center; gap:8px;
}
.sc-title::before { content:''; width:3px; height:12px; background:#2E75B6; border-radius:2px; }

/* Ajuste fino Plotly */
.js-plotly-plot { margin-top: 0 !important; }

/* ── TABELA ── */
.sc-table { background:#112035; border:1px solid #1A3A5C; border-radius:10px; padding:15px; margin-top: 15px; }
.tbl { width:100%; border-collapse:collapse; font-size:12px; }
.tbl th { background:#0D2137; padding:10px; text-align:left; color:#3A6A9A; font-size:10px; text-transform:uppercase; }
.tbl td { padding:10px; border-bottom:1px solid #152840; color:#8AAABB; }
.td-eq { color:white; font-weight:600; }
.badge { padding:3px 10px; border-radius:20px; font-size:9px; font-weight:700; text-transform:uppercase; }
.ba { background:rgba(239,68,68,.15); color:#EF4444; border:1px solid #EF4444; }
.bl { background:rgba(245,158,11,.15); color:#F59E0B; border:1px solid #F59E0B; }

/* Selectbox labels */
div[data-testid="stSelectbox"] label { color:#BDD7EE !important; font-size:12px !important; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

# [Funções de extração permanecem as mesmas que as suas para manter a lógica de negócio]
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

def extrair_laudo(nome, pdf_bytes):
    dados = {"Arquivo": nome}
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        pags = [p.extract_text() or "" for p in pdf.pages]
    p1 = pags[0]; texto = "\n".join(pags)
    dados["Status"] = detectar_status(pdf_bytes)
    m = re.search(r"Ativo:\s*(.+?)(?=Número de série:|\n\n)", p1, re.DOTALL)
    raw = re.sub(r"\s+", " ", m.group(1)).strip() if m else ""
    dados.update(split_ativo(raw))
    m = re.search(r"data de coleta:\s*([\d/,: ]+)", p1)
    ds = m.group(1).strip() if m else None
    dados["Data de coleta"] = ds
    if ds:
        partes = ds.split('/')
        if len(partes) >= 3:
            dados["Mês coleta"] = int(partes[1])
            dados["Ano coleta"] = int(partes[2][:4])
    
    dados["Observações"] = re.search(r"Observações:\s*\n(.*?)(?=Diagnósticos:)", p1, re.DOTALL).group(1).strip() if "Observações:" in p1 else ""
    
    # Extração simples de parâmetros para exemplo da tabela
    for p_name, p_col in [("Ferro","Ferro (ppm)"), ("água","Água (ppm)"), ("Visc 40","Visc 40 (cSt)")]:
        m = re.search(rf"{p_name}\s+([\d\.,]+)", texto)
        dados[p_col] = to_num(m.group(1)) if m else None

    m = re.search(r"Limites acionados \(amostra atual\)\n(.+?)(?=TruVu|Relatório)", texto, re.DOTALL)
    dados["Parâmetros em Alarme"] = m.group(1).split('\n')[0].strip() if m else ""
    return dados

def render_dashboard(lista):
    df = pd.DataFrame(lista)
    total_g, norm_g, alt_g, alm_g = len(df), len(df[df["Status"]=="Normal"]), len(df[df["Status"]=="Alerta"]), len(df[df["Status"]=="Alarme"])

    # Cabeçalho
    st.markdown(f"""
    <div class="db-header">
        <div>
            <div class="db-title">MONITORAMENTO DE ANÁLISE DE ÓLEO</div>
            <div style="font-size:11px; color:#BDD7EE; opacity:0.8">Gerdau Charqueadas • Engenharia de Manutenção</div>
        </div>
        <div style="text-align:right; color:white;">
            <div style="font-size:10px; font-weight:700; background:rgba(255,255,255,0.1); padding:4px 12px; border-radius:5px;">{datetime.now().strftime('%d/%m/%Y %H:%M')}</div>
        </div>
    </div>""", unsafe_allow_html=True)

    # Filtros
    fc1, fc2, fc3 = st.columns([2, 1, 1])
    with fc1: setor_sel = st.selectbox("🏭 SETOR (Cod2)", ["Todos"] + sorted(df["Cod2"].unique().tolist()))
    with fc2: ano_sel = st.selectbox("📅 ANO", ["Todos"] + sorted(df["Ano coleta"].dropna().unique().tolist(), reverse=True))
    with fc3: st.button("↩ Novo ZIP", on_click=lambda: st.session_state.clear(), use_container_width=True)

    dff = df.copy()
    if setor_sel != "Todos": dff = dff[dff["Cod2"] == setor_sel]
    
    # KPIs
    total, norm, alt, alm = len(dff), len(dff[dff["Status"]=="Normal"]), len(dff[dff["Status"]=="Alerta"]), len(dff[dff["Status"]=="Alarme"])
    
    k1, k2, k3, k4 = st.columns(4)
    with k1: st.markdown(f'<div class="kpi kpi-tot"><div class="kv">{total}</div><div class="kl">Total Ativos</div><div class="ks">Laudos Processados</div></div>', unsafe_allow_html=True)
    with k2: st.markdown(f'<div class="kpi kpi-nor"><div class="kv">{norm}</div><div class="kl">Normal</div><div class="ks">{norm/total*100:.0f}% da Frota</div></div>', unsafe_allow_html=True)
    with k3: st.markdown(f'<div class="kpi kpi-alt"><div class="kv">{alt}</div><div class="kl">Alerta</div><div class="ks">{alt/total*100:.0f}% da Frota</div></div>', unsafe_allow_html=True)
    with k4: st.markdown(f'<div class="kpi kpi-alm"><div class="kv">{alm}</div><div class="kl">Alarme</div><div class="ks">{alm/total*100:.0f}% da Frota</div></div>', unsafe_allow_html=True)

    # Gráficos
    st.write("") 
    g1, g2, g3 = st.columns([2, 1, 1])
    COR = {"Normal":"#639922","Alerta":"#BA7517","Alarme":"#E24B4A"}
    
    with g1:
        st.markdown('<div class="sc-box"><div class="sc-title">Status por Setor (Cod2)</div>', unsafe_allow_html=True)
        db = dff.groupby(["Cod2","Status"]).size().reset_index(name="n")
        fig = px.bar(db, x="n", y="Cod2", color="Status", orientation="h", color_discrete_map=COR, height=210)
        fig.update_layout(margin=dict(l=10,r=10,t=10,b=10), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", 
                          font=dict(color="#8AAABB"), showlegend=False, xaxis_title="", yaxis_title="")
        st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
        st.markdown('</div>', unsafe_allow_html=True)

    with g2:
        st.markdown('<div class="sc-box"><div class="sc-title">Distribuição</div>', unsafe_allow_html=True)
        fig2 = go.Figure(go.Pie(labels=dff["Status"], hole=.6, marker_colors=[COR.get(s) for s in dff["Status"].unique()]))
        fig2.update_layout(margin=dict(l=10,r=10,t=10,b=10), height=210, paper_bgcolor="rgba(0,0,0,0)", 
                           showlegend=True, legend=dict(font=dict(size=10, color="#8AAABB"), orientation="h", y=-0.2))
        st.plotly_chart(fig2, use_container_width=True, config={'displayModeBar': False})
        st.markdown('</div>', unsafe_allow_html=True)

    with g3:
        st.markdown('<div class="sc-box"><div class="sc-title">Coletas / Mês</div>', unsafe_allow_html=True)
        dm = dff.groupby("Mês coleta").size().reset_index(name="n")
        fig3 = px.bar(dm, x="Mês coleta", y="n", height=210)
        fig3.update_traces(marker_color="#2E75B6")
        fig3.update_layout(margin=dict(l=10,r=10,t=10,b=10), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", 
                           font=dict(color="#8AAABB"), xaxis_title="", yaxis_title="")
        st.plotly_chart(fig3, use_container_width=True, config={'displayModeBar': False})
        st.markdown('</div>', unsafe_allow_html=True)

    # Tabela
    st.markdown('<div class="sc-table"><div class="sc-title" style="background:none; border:none; padding:0 0 15px 0;">Equipamentos com Desvio — Ação Necessária</div>', unsafe_allow_html=True)
    dtab = dff[dff["Status"] != "Normal"].sort_values("Status")
    rows = ""
    for _, r in dtab.iterrows():
        bcls = "ba" if r['Status']=="Alarme" else "bl"
        rows += f"""<tr>
            <td style="color:#2E75B6; font-weight:700;">{r.get('Cod2','')}</td>
            <td class="td-eq">{r.get('Equipamento Descrição 1','')}</td>
            <td><span class="badge {bcls}">{r['Status']}</span></td>
            <td style="color:#C8A870">{r.get('Parâmetros em Alarme','')}</td>
            <td style="font-size:10px; font-style:italic;">{str(r.get('Observações',''))[:60]}...</td>
            <td style="color:#EF4444; font-weight:700;">{r.get('Ferro (ppm)','—')}</td>
            <td style="color:#EF4444;">{r.get('Água (ppm)','—')}</td>
            <td>{r.get('Data de coleta','')}</td>
        </tr>"""
    st.markdown(f'<table class="tbl"><thead><tr><th>Setor</th><th>Equipamento</th><th>Status</th><th>Parâmetro</th><th>Observação</th><th>Ferro</th><th>Água</th><th>Coleta</th></tr></thead><tbody>{rows}</tbody></table></div>', unsafe_allow_html=True)

# ── ROTEADOR ──
if "laudos" not in st.session_state:
    # Simulação de processamento para exemplo ou tela de upload
    uploaded = st.file_uploader("ZIP", type="zip")
    if uploaded:
        with zipfile.ZipFile(uploaded) as z:
            pdfs = [n for n in z.namelist() if n.lower().endswith(".pdf")]
            lista = [extrair_laudo(n, z.read(n)) for n in pdfs]
            st.session_state["laudos"] = lista
            st.rerun()
else:
    render_dashboard(st.session_state["laudos"])
