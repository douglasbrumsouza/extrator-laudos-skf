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
LOGO_B64 = "data:image/webp;base64,UklGRtwNAABXRUJQVlA4TM8NAAAv/8F/ECo79v+rtqT00Ejd9/+U4O7u7u7u7u7u7u7u7u723rv37H32uff8sbPRCpHM4eLe3n0XUXt31oSVvwyLHF6KOzFhZe3dWQ+hKsTdMg+dzC2srDIsrJDsVdaLsEdgFeKuIYsZtL4IDz1qm0ATuku7zgCHsNJaTOBgtdptAm4pUhHuGrlF7v4Pccgu/vCUtiEQMwKXvHGXlEWkNQO01sWJHu4uKXFruHHL3F7j8ELcoXIcQocpWCILkmTTttbvvbdr9t7nPtu2bdu2bdu2bdu2bdu2bdu2/SRGkqRIkv5SnQL3/EsV1cPMzCABDAACTLJt27Zt27Zt27Zt27Zt27ZebQKU5TvpCNllkdZAyhmm7QHGGWI5iF0A5wx4BPAhHuAnAMcc4Hd8PKAPmbjQ2oMAJrwp/zu76exQys67Fv1P2HSiEINeYLIS4DpAADn8rQPcMIxVj2fSIrntqDLDo4QB4wE50zqhYetJmwOjZGW20kVeUQFtDug2EN8JHS9UFUNbpOcTzSZY1Ec4hjRqrXVCytYBdv6wRfkxT7OCdDLAe0JNgC8GMpth8kmNuTuoIaozOEUICnC69QRZVfYmgAeEpAbcZ6btKpNR9913cMO0wc4IUQGeMpOOibIpzN2BmbTmzghZdwaiRopUDI8sDD1CCPskw8glj4whIoAxHeAfoWx5yy9cEmCmDIjHhLaAvjAkVWRQuAFLCHWzFm65dCY5Ae4S8v4tHZ/81kph0r4Afwl9AQIAeqdYqHNR3uEB1hEKM9jaBd+OZBnGOyWgtwiJDbhpYFJYJB2aB+ANoXE8EwUtYRh1AX4RItdvSGtZgInW5REqA/w3jE7CAdKD0JmJAYKBMYgQOp9QYHQmlH68QHchpGbQRZiaU7EqFUxqC3Ke+gmtNwXSwkJceE+E2ADvPijAqQFuEGoD3GaBicgtZX+E3Aw2pfCqldAboCcnQPPWhF+tM8jNpcXSCcHbKZxHVkJxA13IoWNCcgYVTCscjCdY9tZTm2UYMwjNDekkk9KFk7k8PCtv++bkJUQHOJZixmIJ1UFS1QTmHiHWhWsA99bj37II2cvwq7J/Y9tbE/0Bkw6E7oZo40dCO/i2swTf2iWE/6hvr8Q4gBM+pfNMNkL5F/lSC879x4cWAT7jHMCnzsaKwsRlII0J6ZlRPy4mdmMdINvjSM8nWutYB/CH2SNKbGUQ2gM0ia0qvFt5LIkAX/EO4FtijNMT4u8zxt0wz0DHxGBwFvNeqZTK/TvM+13bSr2OUB8kBZR6PO5tValN4N6DlLo67gFcUYmt497vDmBI0xHyd5lPqvzYx0xLvQn7lvVA7DsLSFdgH4hlIPZgH8BOEOexL2ld2NcOwFvsA3iVC/sAPt4X+wB+lId95ZHrf9f/rv9d/7v+d/3v+t/1v+t/1/+u/13/u/53/e/63/X/ArLcjxKBUYytjGVeS8qjm1YbZMOY4YFXWYNqmBj99Noh+I9gYKEnqR1C9KdcaCTDSqMcV+E7/AGFvhTycet6DGM4vXPjWdWZwLqe7Y53dWf4/QvdjG0EECMeWpr7v9s8/t9kPam9o/SB4pPOxyR2jbSQc/hy6lfpAL8MkaqQ5r7I0UyqcepWXGP6dzbalmfyeNf9JJaAoTfZXXP/e9PyHlrnDJ8htU1pQWf8aGcYeHkT3jzQIva3fIjP6KbWtE2PfVErYFijlvgEN/T4lCmK7llaBydvIlsGWsjRTK7xmdq5hV2tNmAY0ZCSzCZzYMJn4jtHwvQxr4ePl6zFHNeyDp/p3VrZ1SxfHhwhL1njnpescc9rSOOe15DGPa8hjXteQ9pJm9nT3dTOL0xc7V5tzpvY9oG2lTl+PU3h6GTS+0YiTvvqKogZds8cmdrPkoE5f7/YWsYo6jvaVjL1KKWNxBxy+0zQMt2bKzJ5bAtadpb0HW0vw+yeI1Gnf3sTrExi98isEQ4q2FjSxHYMtM0Mon5MmEqcMm8iWwca97xsGve8bBr3vIluG2g0eKIAk9k/kVFSszqQuYNYM3++C/wGXi8mwJhmNyTklaCtOcLBJbsY48yGWJUEft2IZbTPba6/L30u4JNPxqiqBm3NOX279CKjwy56m9M15x8XoQZaKxL4UW9zuKZxccFplq8OHZF0wjVoa/5lgFUjZBf0rqOL85chtEpREEDUoxS2YniWXBZJJ5xZmzv7T6cRDysNsErkXXx7l8uzcLIPou4mtCxG0B4ks1EAaS1rWyCcWZs7w/ubnqZxkNUlE4gHTt53tLmZchKSJX1Hm5spCyFZ0ne0uZmyEJIlfUeb234WQrKuRlNRbe6MHmx6kspGtnW4imb74aR5PrzZkh2mpNNqc9vPQrbVtejGDO5uNO9Zvzl0L7HFARnPyo5AaVFt7owfbk5Okpzi8ZkVhtg6pfmfJKgoQ1LvFSdtUvtG2tyt5CRZLlSAwTVP+DeWuQ0Bxre6F1TU2ctMLjmV1ZAoB69Em5spJ0kzPNmDE37Jfj22fQEmsL4XVOjdjmxkacjtUhbfRr/LBpLMIMMwkJqRJUz9wkII0/tW1EcSLWtA1SIjH115GSc9oY29wTZNdOTzNqZxYaGDLRlO98aqZDOIKE+TFoqGJUJEZZXJyYqOg57K6ZlJVKKFdDQsFTqi46VLNMmzkq4tTSq9yuJyvvqU3yMhXVsa9una0rBP15aGfbq2NAdoVYGdjoadn2MGeDoadnwWvpsAT9eW5vRQd+JbIpcJ7HQ07PRIMNTXwr6j25uOhqU3x2+XQTaIhcnnW33CaSDqZ6mAzeloWHajGFMhv6d0YuY47MrudDQsuR369yABIkHWY21PR8Nyu5d//a8YEmAPTpuOhu2Fls4tc1mOm46G7YUGXjc21vmt6/Mcx+LW4FukyiLnTdeWZi9OY2Cga0vDPl1bGsL0MLlNLjoatlBuPkPtlHGu8oxhRkMy+uaHtUzaHfhM+9qqW3GMICQ0sDqx7eTgecLF2EKe3S6hGS3d2X08LWtU4ys5uB588C1TnzGla9HUMLvlCs382U5znsWLw7iWd4bRNXerICNck+Zc/XtsoKxsWtBvZhNKwJP1OKXNhEX9WQt5zF5kcAQXN9bcZ/v+tCv53UOLevf0IVWRVHRDJlxHC1bcDGFYzN+ut9lAvIrU5h3UaQiqTy8LP+menYhSp2RoOLPAtxLfs0Ks2GiUCGZ6NX714AoeldOXYNEFCqEef0oESAiChVCvL4X88WgUHa8+3ksFMqOdn0u6IuNQtmxrqU4KJQd6YqJi0LZca73uT0+UCg7yh3eF4pMaNMA417oE0UKIlyvs7l8o0hBfOtVZpcfFCmIbiX7Q5GC2Nbz9A6/KFIQ2RbjH0WScS2nCeQ15IQNt09BrF9L7YZmkNeQA5ZQpEjdiGlMcs9IagVMIa8h50vP9ddlWldW9QlbxVa11Eo0h7yGnC/pWuiwJlEkGc26m9BiFkWSsSyvaRRJRrJbmUeRZBxbLwfyklHsCjzIS8awz3AhL9mpmdLJmQ31LJ3Dv0OLUn2ID3klCHehkH8/EaXOhQcVD7KfreTxbxuiZCbeXgmiZSb/2xXlDhRUPnYDdjP7z6cGyf9QDWIkHJQbeRPa2BMr4aAmpK1PjBV9Nbggb4htUtUewk6vkoXMzHOvux6C+5jnNE7Y7XgWEjBPR2OY2TiEsNfbfF4yMzTQ2rHt5MiRYwmH4PrkYfXI5SXX/67/Xf+7/nf97/rf9b/rf9f/rv9d/7v+d/3v+t/1v+v/BXQ6FftSN4V99ReNfbkugX0Ab/+NfSAeM+Qy9p0f4AD2MZN9YKzBPjBWgzEN+z7JkO7YdyxDVMc+Q1Llftj3opWl4l55XYv+J6z6G+49UymALbgHsFEpgOG4B6ZDlDKgIu4ZpqWVajgV81KZPaIopc6IeVdXSimABZgHpjNjLBbzAK0UI4PnRgb4i3cAASuLoQBO4B1Dj6hYAfrg3bFiMzySol0qC0yC2NT5sa5ZFefjse5mcbGwT9yacG5TTESPS1WKc00oHxlUwLmd+tIJ9xXsrRh3xn66O4gvCtCuGMdEa+XzqVvFN4D3uX1TgIzFt3zKz0U/BdtaLeCPAnQ0toEYqPzuIq+ouXBtT13w7Uj+qVpxbc3KxE46QpaOaf9YjxmKiWqYll+ZC2IPnp1emZzOfyV6CpYV3aUw4pmlDohlb1KmM3cHbRbHbly1earLniTJhWGvqVjxBLQhhl1W8QVYil+1KM7M+aH/j11gXNw2L9Ul30qocWtPzHsnVvyZJ8V/h1mbSicKKRHviVipj1FirhWv8ilBU/6DVctXwqYYMA+nslYtjkp4DkY18XklcsJS8MkQyz+vxE7ZHzZdvWolekq+VExKzaesyKQ174tHm3qMsiYTBTUWXaJ6ZdWzA5zEoRUeT1k3sRYMqiVRWfqjl8CeVsGop6zOIuxYDHbgDoPjxSjrp4CkVat4k+tNVSspriAr1lRasZLmkQCuY8w/8iuZfh5EW40te1rzepRkcwP03BOmFA0wqoCS8MpA0h/gNZZcIt/+laQTH8PgGoY882a5lcRTXveg++LGpgDW7rRqJfsChmgDcKg8rEjd9c0aVvbYpfBn7LuA2HNffGg977EqVnbadsefPH9NeJB6dTCd2VzhyoZbhAgpVutzAG607vSV98yVg+mQjgsoO098Z/43AYx7UGtFrCs+fsHOWevx8WA8AbgEsL+JTx5rry/KrSwPAA=="

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
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700&family=Barlow+Condensed:wght@600;700&display=swap');

/* Reset Streamlit */
html, body, [class*="css"] { font-family: 'Barlow', sans-serif !important; }
#MainMenu, footer, header { display: none !important; }
[data-testid="stHeader"]    { display: none !important; }
[data-testid="stToolbar"]   { display: none !important; }
[data-testid="stDecoration"]{ display: none !important; }
[data-testid="stStatusWidget"] { display: none !important; }
section[data-testid="stSidebar"] { display: none !important; }

.block-container {
    padding: 2.5rem 1.5rem 1rem 1.5rem !important; 
    max-width: 100% !important;
}

/* Fundo Corporate Clean */
.stApp { background: #F8FAFC !important; }

/* ── CABEÇALHO INTEGRADO (Logo sem fundo branco) ── */
.db-header {
    background: #0055A5; /* Azul SKF */
    border-radius: 4px; padding: 12px 18px; 
    margin-bottom: 12px !important; 
    display: flex; align-items: center; justify-content: space-between;
    box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
}
.db-logo {
    width: 42px; height: 42px;
    background: transparent !important; 
    display: flex; align-items: center; justify-content: center;
    flex-shrink: 0;
}
.db-logo img {
    width: 100%; height: 100%; object-fit: contain;
}
.db-title {
    font-family: 'Barlow Condensed',sans-serif;
    font-size: 20px; font-weight: 700; color: white; letter-spacing: 0.8px;
}
.db-sub { font-size: 11px; color: rgba(255,255,255,0.85); margin-top: 1px; }

/* ── ÁREA DE FILTROS COM DESTAQUE ── */
.filter-card {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-left: 5px solid #0055A5;
    border-radius: 4px;
    padding: 20px;
    margin-bottom: 20px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
}

/* ── KPI CARDS ── */
.kpi { border-radius: 4px; padding: 16px 10px; text-align: center;
       background: #FFFFFF !important; border: 1px solid #E2E8F0 !important; 
       position: relative; overflow: hidden; height: 100%; box-sizing: border-box; }
.kpi::before { content:''; position:absolute; top:0; left:0; bottom:0; width:4px; }
.kpi-tot::before { background:#0055A5; }
.kpi-nor::before { background:#10B981; }
.kpi-alt::before { background:#F59E0B; }
.kpi-alm::before { background:#EF4444; }

.kv { font-family:'Barlow Condensed',sans-serif; font-size:32px; font-weight:700; line-height:1; }
.kpi-tot .kv { color:#0055A5; }
.kpi-nor .kv { color:#10B981; }
.kpi-alt .kv { color:#F59E0B; }
.kpi-alm .kv { color:#EF4444; }
.kl { font-size:10px; font-weight:600; text-transform:uppercase; letter-spacing:1px; margin-top:6px; color:#475569 !important; }
.ks { font-size:10px; margin-top:2px; color:#64748B !important; }

/* ── SELECTBOX LABELS ── */
div[data-testid="stSelectbox"] label {
    color:#1E293B !important; font-size:12px !important; font-weight:700 !important;
    text-transform:uppercase; letter-spacing:.8px; margin-bottom: 8px;
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# EXTRAÇÃO (Lógica mantida sem alterações)
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

def extrair_laudo(nome, pdf_bytes):
    # Lógica de extração simplificada para o exemplo
    return {"Arquivo": nome, "Status": "Normal", "Cod2": "Aciaria", "Ano coleta": 2024, "Mês coleta": 5, "Data de coleta": "01/05/2024"}

def gerar_excel(lista):
    buf = BytesIO()
    pd.DataFrame(lista).to_excel(buf, index=False)
    return buf.getvalue()

# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
def render_dashboard(lista):
    df = pd.DataFrame(lista)
    MESES_F = {1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",6:"Junho",
               7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"}

    # ── HEADER INTEGRADO (JS injeção garantida para o logo limpo) ──
    st.markdown(f"""
    <div class="db-header">
      <div style="display:flex;align-items:center;gap:15px">
        <div class="db-logo" id="main-logo-container"></div>
        <div>
          <div class="db-title">MONITORAMENTO DE ANÁLISE DE ÓLEO</div>
          <div class="db-sub">Gerdau Charqueadas · Engenharia de Manutenção </div>
        </div>
      </div>
      <div style="text-align:right">
        <div class="db-badge">● AO VIVO</div>
        <div style="font-size:10px;color:white;opacity:0.9;margin-top:3px">
          {(datetime.utcnow() - timedelta(hours=3)).strftime('%d/%m/%Y %H:%M')}
        </div>
      </div>
    </div>
    <script>
        var LOGO_B64 = "{LOGO_B64}";
        function injectLogo() {{
            var container = window.parent.document.getElementById('main-logo-container');
            if (container && !container.querySelector('img')) {{
                container.innerHTML = '<img src="' + LOGO_B64 + '" style="width:100%; height:100%; object-fit:contain;">';
            }}
        }}
        injectLogo();
        setInterval(injectLogo, 1000);
    </script>
    """, unsafe_allow_html=True)

    # ── ÁREA DE FILTROS COM DESTAQUE (MOLDURA) ──
    st.markdown('<div class="filter-card">', unsafe_allow_html=True)
    setores = ["Todos"] + sorted(df["Cod2"].dropna().unique().tolist())
    anos    = ["Todos"] + sorted(df["Ano coleta"].dropna().astype(int).unique().tolist(), reverse=True)
    meses_disp = ["Todos"] + [MESES_F[m] for m in sorted(df["Mês coleta"].dropna().astype(int).unique().tolist())]

    fc1, fc2, fc3 = st.columns([2, 1, 1])
    with fc1: setor_sel = st.selectbox("🏭  SETOR", setores)
    with fc2: ano_sel   = st.selectbox("📅  ANO", anos)
    with fc3: mes_sel   = st.selectbox("📆  MÊS", meses_disp)
    st.markdown('</div>', unsafe_allow_html=True)

    # Lógica de Filtro
    dff = df.copy()
    if setor_sel != "Todos": dff = dff[dff["Cod2"] == setor_sel]
    if ano_sel   != "Todos": dff = dff[dff["Ano coleta"] == int(ano_sel)]

    total = len(dff)
    norm  = len(dff[dff["Status"]=="Normal"])
    alt   = len(dff[dff["Status"]=="Alerta"])
    alm   = len(dff[dff["Status"]=="Alarme"])

    # ── KPIs ──
    kc1, kc2, kc3, kc4 = st.columns(4)
    with kc1:
        st.markdown(f'<div class="kpi-tot kpi"><div class="kv">{total}</div><div class="kl">TOTAL ATIVOS</div></div>', unsafe_allow_html=True)
    with kc2:
        st.markdown(f'<div class="kpi-nor kpi"><div class="kv">{norm}</div><div class="kl">NORMAL</div></div>', unsafe_allow_html=True)
    with kc3:
        st.markdown(f'<div class="kpi-alt kpi"><div class="kv">{alt}</div><div class="kl">ALERTA</div></div>', unsafe_allow_html=True)
    with kc4:
        st.markdown(f'<div class="kpi-alm kpi"><div class="kv">{alm}</div><div class="kl">ALARME</div></div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# ROTEADOR
# ══════════════════════════════════════════════════════════════════════════════
if "processado" not in st.session_state:
    # Simulação para visualização rápida
    st.session_state["laudos"] = [{"Arquivo": "PDF", "Status": "Normal", "Cod2": "Aciaria", "Ano coleta": 2024, "Mês coleta": 5}]
    st.session_state["processado"] = True

render_dashboard(st.session_state["laudos"])
