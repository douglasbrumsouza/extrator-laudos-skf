import streamlit as st
import streamlit.components.v1 as components
import zipfile
import re
import os
from io import BytesIO
from datetime import datetime

try:
    import pdfplumber
except ImportError:
    st.error("pip install pdfplumber")
    st.stop()

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    st.error("pip install openpyxl")
    st.stop()

try:
    import pandas as pd
    import plotly.express as px
    import plotly.graph_objects as go
except ImportError:
    st.error("pip install pandas plotly")
    st.stop()

# =========================================================
# CONFIG
# =========================================================

st.set_page_config(
    page_title="Análise de Óleo SKF — Gerdau",
    page_icon="🛢️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =========================================================
# CSS GLOBAL
# =========================================================

st.markdown("""
<style>

@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@300;400;500;600;700&family=Barlow+Condensed:wght@600;700&display=swap');

/* RESET */

html, body, [class*="css"] {
    font-family: 'Barlow', sans-serif !important;
}

#MainMenu,
footer,
header,
[data-testid="stHeader"],
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"],
section[data-testid="stSidebar"] {
    display: none !important;
}

/* APP */

.stApp {
    background: linear-gradient(135deg,#0D2137 0%,#1A3A5C 100%) !important;
}

.block-container {
    padding: 1rem 1rem 1rem 1rem !important;
    max-width: 100% !important;
}

/* ESPAÇAMENTOS */

div[data-testid="stVerticalBlock"] {
    gap: 0.55rem !important;
}

div[data-testid="stHorizontalBlock"] {
    gap: 0.8rem !important;
    align-items: stretch !important;
}

div[data-testid="column"] {
    padding-top: 0 !important;
    padding-bottom: 0 !important;
}

/* SELECTBOX */

div[data-testid="stSelectbox"] label {
    color:#BDD7EE !important;
    font-size:11px !important;
    text-transform:uppercase;
    letter-spacing:.8px;
    margin-bottom:4px !important;
}

div[data-baseweb="select"] {
    margin-top: 0 !important;
}

/* HEADER */

.db-header {
    background: linear-gradient(135deg,#1F4E79,#2E75B6);
    border-radius: 10px;
    padding: 12px 16px;
    margin-bottom: 4px;
    display:flex;
    align-items:center;
    justify-content:space-between;
    box-shadow: 0 4px 20px rgba(0,0,0,0.3);
}

.db-logo {
    width:36px;
    height:36px;
    border-radius:8px;
    background:rgba(255,255,255,.15);
    border:1px solid rgba(255,255,255,.2);
    display:flex;
    align-items:center;
    justify-content:center;
    font-family:'Barlow Condensed',sans-serif;
    font-size:12px;
    font-weight:700;
    color:white;
}

.db-title {
    font-family:'Barlow Condensed',sans-serif;
    font-size:18px;
    font-weight:700;
    color:white;
}

.db-sub {
    font-size:10px;
    color:rgba(255,255,255,.65);
}

.db-badge {
    font-size:9px;
    font-weight:700;
    color:#BDD7EE;
    background:rgba(46,117,182,.3);
    padding:2px 10px;
    border-radius:20px;
}

.db-credit {
    font-size:8px;
    color:rgba(255,255,255,.22);
    margin-top:3px;
    text-align:right;
}

/* ACTION BAR */

.act-bar {
    background:#0D2137;
    border:1px solid #1A3A5C;
    border-radius:8px;
    padding:8px 14px;
    display:flex;
    align-items:center;
    gap:10px;
    min-height:42px;
}

.act-title {
    font-family:'Barlow Condensed',sans-serif;
    font-size:14px;
    font-weight:700;
    color:white;
    flex:1;
}

.bcnt {
    font-size:10px;
    font-weight:700;
    padding:2px 8px;
    border-radius:20px;
    white-space:nowrap;
}

/* KPI */

.kpi {
    border-radius:10px;
    padding:12px 10px;
    border:1px solid transparent;
    position:relative;
    overflow:hidden;
    min-height:118px;
    display:flex;
    flex-direction:column;
    justify-content:center;
    align-items:center;
    text-align:center;
    box-sizing:border-box;
}

.kpi::before {
    content:'';
    position:absolute;
    top:0;
    left:0;
    right:0;
    height:3px;
}

.kpi-tot {
    background:#162E4A;
    border-color:#2A4A6A;
}

.kpi-tot::before {
    background:#2E75B6;
}

.kpi-nor {
    background:#0E2B16;
    border-color:#1A4D22;
}

.kpi-nor::before {
    background:#4CAF50;
}

.kpi-alt {
    background:#2B1E05;
    border-color:#4A3210;
}

.kpi-alt::before {
    background:#F59E0B;
}

.kpi-alm {
    background:#2B0A0A;
    border-color:#4A1A1A;
}

.kpi-alm::before {
    background:#EF4444;
}

.kv {
    font-family:'Barlow Condensed',sans-serif;
    font-size:34px;
    font-weight:700;
    line-height:1;
}

.kpi-tot .kv { color:white; }
.kpi-nor .kv { color:#6BCB77; }
.kpi-alt .kv { color:#F59E0B; }
.kpi-alm .kv { color:#EF4444; }

.kl {
    font-size:9px;
    font-weight:700;
    text-transform:uppercase;
    letter-spacing:1px;
    margin-top:6px;
}

.ks {
    font-size:10px;
    margin-top:3px;
}

/* CONTAINERS */

.sc {
    background:#112035;
    border:1px solid #1A3A5C;
    border-radius:10px;
    padding:12px 14px;
    overflow:hidden;
}

.sc-top {
    background:#112035;
    border:1px solid #1A3A5C;
    border-bottom:none;
    border-radius:10px 10px 0 0;
    padding:10px 14px 6px 14px;
    margin-bottom:0 !important;
}

.st {
    font-size:10px;
    font-weight:700;
    text-transform:uppercase;
    letter-spacing:1px;
    color:#BDD7EE;
    display:flex;
    align-items:center;
    gap:6px;
}

.st::before {
    content:'';
    width:3px;
    height:12px;
    border-radius:2px;
    background:#2E75B6;
}

/* PLOTLY */

div[data-testid="stPlotlyChart"] {
    background:#112035;
    border:1px solid #1A3A5C;
    border-top:none;
    border-radius:0 0 10px 10px;
    padding:0 10px 10px 10px;
    margin-top:0 !important;
    overflow:hidden;
}

div[data-testid="stPlotlyChart"] > div {
    margin-top:0 !important;
    padding-top:0 !important;
}

/* TABLE */

.tbl {
    width:100%;
    border-collapse:collapse;
    font-size:11px;
}

.tbl thead tr {
    background:#0D2137;
}

.tbl thead th {
    padding:8px 10px;
    text-align:left;
    font-size:9px;
    font-weight:700;
    text-transform:uppercase;
    letter-spacing:.7px;
    color:#3A6A9A;
    border-bottom:1px solid #1A3A5C;
}

.tbl tbody tr {
    border-bottom:1px solid #152840;
}

.tbl tbody tr:hover {
    background:#162E4A;
}

.tbl tbody td {
    padding:8px 10px;
    color:#8AAABB;
}

.td-eq {
    color:white !important;
    font-weight:500 !important;
}

.td-set {
    font-family:'Barlow Condensed',sans-serif;
    font-size:13px;
    font-weight:700;
    color:#2E75B6 !important;
}

.td-par {
    color:#C8A870 !important;
}

.td-obs {
    font-size:10px;
    color:#8A6A3A !important;
    font-style:italic;
}

.td-data {
    font-size:10px;
    color:#3A6A9A !important;
}

.td-red {
    color:#EF4444 !important;
    font-weight:700 !important;
}

.td-yel {
    color:#F59E0B !important;
    font-weight:700 !important;
}

.badge {
    display:inline-block;
    font-size:9px;
    font-weight:700;
    padding:2px 8px;
    border-radius:20px;
    text-transform:uppercase;
}

.ba {
    background:rgba(239,68,68,.15);
    color:#EF4444;
    border:1px solid rgba(239,68,68,.3);
}

.bl {
    background:rgba(245,158,11,.15);
    color:#F59E0B;
    border:1px solid rgba(245,158,11,.3);
}

.bn {
    background:rgba(76,175,80,.15);
    color:#4CAF50;
    border:1px solid rgba(76,175,80,.3);
}

div[data-testid="stButton"] button,
div[data-testid="stDownloadButton"] button {
    min-height:42px !important;
    border-radius:8px !important;
}

</style>
""", unsafe_allow_html=True)

# =========================================================
# FUNÇÕES
# =========================================================

def to_num(v):
    if v is None:
        return None

    s = str(v).strip().replace(",", ".")

    try:
        return float(s)
    except:
        return s if s else None


def split_data(s):

    if not s:
        return None, None, None, None

    m = re.match(
        r"(\\d{2})/(\\d{2})/(\\d{4}),?\\s*(\\d{2}:\\d{2}:\\d{2})?",
        str(s).strip()
    )

    if not m:
        return None, None, None, None

    return (
        int(m.group(1)),
        int(m.group(2)),
        int(m.group(3)),
        m.group(4)
    )


def detectar_status(pdf_bytes):

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:

        full = "".join(p.extract_text() or "" for p in pdf.pages)

        if "Limites acionados" not in full:
            return "Normal"

        for img in pdf.pages[0].images:

            if img.get("name") == "Im2":

                return "Alarme" if img["height"] >= 50 else "Alerta"

    return "Alerta"


def extrair_laudo(nome, pdf_bytes):

    dados = {"Arquivo": nome}

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        pags = [p.extract_text() or "" for p in pdf.pages]

    texto = "\\n".join(pags)

    dados["Status"] = detectar_status(pdf_bytes)

    m = re.search(r"Localização:\\s*(.+)", texto)

    dados["Localização"] = m.group(1).strip() if m else ""

    m = re.search(r"Ativo:\\s*(.+)", texto)

    dados["Equipamento"] = m.group(1).strip() if m else ""

    m = re.search(r"data de coleta:\\s*([\\d/,: ]+)", texto)

    ds = m.group(1).strip() if m else None

    dados["Data de coleta"] = ds

    d, me, a, h = split_data(ds)

    dados["Dia coleta"] = d
    dados["Mês coleta"] = me
    dados["Ano coleta"] = a
    dados["Hora coleta"] = h

    return dados


def gerar_excel(lista):

    wb = openpyxl.Workbook()

    ws = wb.active

    ws.title = "Laudos"

    cols = list(lista[0].keys())

    for c, nome in enumerate(cols, 1):

        ws.cell(1, c, nome)

    for r, linha in enumerate(lista, 2):

        for c, nome in enumerate(cols, 1):

            ws.cell(r, c, linha.get(nome))

    for i, col in enumerate(cols, 1):

        ws.column_dimensions[get_column_letter(i)].width = 24

    buf = BytesIO()

    wb.save(buf)

    buf.seek(0)

    return buf.read()

# =========================================================
# UPLOAD
# =========================================================

uploaded = st.file_uploader(
    "Selecione o ZIP",
    type=["zip"]
)

if uploaded:

    with st.spinner("Processando PDFs..."):

        zf = zipfile.ZipFile(BytesIO(uploaded.read()))

        pdfs = [
            (n, zf.read(n))
            for n in zf.namelist()
            if n.lower().endswith(".pdf")
        ]

        lista = []

        for nome, pdf_bytes in pdfs:

            try:
                lista.append(extrair_laudo(nome, pdf_bytes))
            except:
                pass

    if not lista:
        st.error("Nenhum PDF válido encontrado.")
        st.stop()

    df = pd.DataFrame(lista)

    excel_bytes = gerar_excel(lista)

    total = len(df)
    norm = len(df[df["Status"] == "Normal"])
    alt = len(df[df["Status"] == "Alerta"])
    alm = len(df[df["Status"] == "Alarme"])

    pn = f"{(norm/total)*100:.0f}%"
    pa = f"{(alt/total)*100:.0f}%"
    pm = f"{(alm/total)*100:.0f}%"

    # =====================================================
    # ACTION BAR
    # =====================================================

    c1, c2 = st.columns([4, 1])

    with c1:

        st.markdown(f"""
        <div class="act-bar">
            <span class="act-title">{total} laudos processados</span>

            <span class="bcnt"
            style="background:rgba(76,175,80,.15);
            color:#6BCB77;
            border:1px solid rgba(76,175,80,.3)">
            {norm} Normal
            </span>

            <span class="bcnt"
            style="background:rgba(245,158,11,.15);
            color:#F59E0B;
            border:1px solid rgba(245,158,11,.3)">
            {alt} Alerta
            </span>

            <span class="bcnt"
            style="background:rgba(239,68,68,.15);
            color:#EF4444;
            border:1px solid rgba(239,68,68,.3)">
            {alm} Alarme
            </span>
        </div>
        """, unsafe_allow_html=True)

    with c2:

        st.download_button(
            "⬇️ Excel",
            data=excel_bytes,
            file_name="laudos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # =====================================================
    # HEADER
    # =====================================================

    st.markdown(f"""
    <div class="db-header">

        <div style="display:flex;align-items:center;gap:12px">

            <div class="db-logo">SKF</div>

            <div>
                <div class="db-title">
                    MONITORAMENTO DE ANÁLISE DE ÓLEO
                </div>

                <div class="db-sub">
                    Gerdau Charqueadas · Engenharia de Manutenção
                </div>
            </div>

        </div>

        <div style="text-align:right">

            <div class="db-badge">● AO VIVO</div>

            <div style="font-size:10px;color:#BDD7EE;margin-top:3px">
                {datetime.now().strftime('%d/%m/%Y %H:%M')}
            </div>

            <div class="db-credit">
                Desenvolvido por Douglas Brum
            </div>

        </div>

    </div>
    """, unsafe_allow_html=True)

    # =====================================================
    # KPIS
    # =====================================================

    kc1, kc2, kc3 = st.columns([2, 1, 1])

    with kc1:

        st.markdown(f"""
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:1rem;width:100%;">

            <div class="kpi kpi-tot">
                <div class="kv">{total}</div>
                <div class="kl">TOTAL DE ATIVOS</div>
                <div class="ks">laudos no período</div>
            </div>

            <div class="kpi kpi-nor">
                <div class="kv">{norm}</div>
                <div class="kl">NORMAL</div>
                <div class="ks">{pn} da frota</div>
            </div>

        </div>
        """, unsafe_allow_html=True)

    with kc2:

        st.markdown(f"""
        <div class="kpi kpi-alt">
            <div class="kv">{alt}</div>
            <div class="kl">ALERTA</div>
            <div class="ks">{pa} da frota</div>
        </div>
        """, unsafe_allow_html=True)

    with kc3:

        st.markdown(f"""
        <div class="kpi kpi-alm">
            <div class="kv">{alm}</div>
            <div class="kl">ALARME</div>
            <div class="ks">{pm} da frota</div>
        </div>
        """, unsafe_allow_html=True)

    # =====================================================
    # GRAFICOS
    # =====================================================

    gc1, gc2, gc3 = st.columns([2, 1, 1])

    COR = {
        "Normal":"#639922",
        "Alerta":"#BA7517",
        "Alarme":"#E24B4A"
    }

    LAYOUT = dict(
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#8AAABB", family="Barlow"),
        margin=dict(l=0,r=0,t=10,b=25)
    )

    with gc1:

        st.markdown(
            '<div class="sc-top"><div class="st">Status por Setor</div></div>',
            unsafe_allow_html=True
        )

        db = df.groupby(["Localização","Status"]).size().reset_index(name="n")

        db["tot"] = db.groupby("Localização")["n"].transform("sum")

        db["pct"] = (db["n"] / db["tot"] * 100).round(1)

        fig = px.bar(
            db,
            x="pct",
            y="Localização",
            color="Status",
            orientation="h",
            barmode="stack",
            color_discrete_map=COR,
            text="pct"
        )

        fig.update_layout(
            **LAYOUT,
            height=230,
            xaxis=dict(showgrid=False,showticklabels=False),
            yaxis=dict(showgrid=False)
        )

        st.plotly_chart(
            fig,
            use_container_width=True,
            config={"displayModeBar":False}
        )

    with gc2:

        st.markdown(
            '<div class="sc-top"><div class="st">Distribuição</div></div>',
            unsafe_allow_html=True
        )

        dp = df["Status"].value_counts().reset_index()

        dp.columns = ["Status","n"]

        fig2 = go.Figure(go.Pie(
            labels=dp["Status"],
            values=dp["n"],
            hole=0.55,
            marker_colors=[COR.get(s,"#666") for s in dp["Status"]]
        ))

        fig2.update_layout(
            **LAYOUT,
            height=230
        )

        st.plotly_chart(
            fig2,
            use_container_width=True,
            config={"displayModeBar":False}
        )

    with gc3:

        st.markdown(
            '<div class="sc-top"><div class="st">Coletas por Mês</div></div>',
            unsafe_allow_html=True
        )

        dm = df.groupby("Mês coleta").size().reset_index(name="n")

        fig3 = go.Figure(go.Bar(
            x=dm["Mês coleta"],
            y=dm["n"]
        ))

        fig3.update_layout(
            **LAYOUT,
            height=230
        )

        st.plotly_chart(
            fig3,
            use_container_width=True,
            config={"displayModeBar":False}
        )

    # =====================================================
    # TABELA
    # =====================================================

    st.markdown(
        '<div class="sc"><div class="st">Laudos Processados</div>',
        unsafe_allow_html=True
    )

    rows = ""

    for _, r in df.iterrows():

        stt = r.get("Status", "")

        badge = "bn"

        if stt == "Alerta":
            badge = "bl"

        if stt == "Alarme":
            badge = "ba"

        rows += f"""
        <tr>

            <td class="td-eq">
                {r.get('Equipamento','')}
            </td>

            <td>
                <span class="badge {badge}">
                    {stt}
                </span>
            </td>

            <td class="td-data">
                {r.get('Data de coleta','')}
            </td>

        </tr>
        """

    st.markdown(f"""
    <table class="tbl">

        <thead>
            <tr>
                <th>Equipamento</th>
                <th>Status</th>
                <th>Data Coleta</th>
            </tr>
        </thead>

        <tbody>
            {rows}
        </tbody>

    </table>
    """, unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

else:

    st.info("Faça upload do ZIP contendo os laudos PDF.")
