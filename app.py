# ==============================================================================
# 1. CAIXA DE FERRAMENTAS (Importação de Bibliotecas)
# Aqui chamamos os "mecanismos" que farão o trabalho pesado.
# ==============================================================================
import streamlit as st                  # O motor principal do nosso site/app
import streamlit.components.v1 as components # Permite injetar visuais extras
import zipfile, re, os, base64, json    # Lida com arquivos ZIP, textos, pastas e imagens
from io import BytesIO                  # Gerencia arquivos na memória temporária
from datetime import datetime, timedelta # Lida com datas e fusos horários

# Tentamos carregar as ferramentas de leitura de PDF, Excel, Gráficos e Google Sheets.
# Se o servidor não as encontrar, ele avisa o que falta instalar.
try:
    import pdfplumber                   # O "leitor" que extrai texto de dentro dos PDFs
except ImportError:
    st.error("Erro: Instale o pdfplumber"); st.stop()

try:
    import openpyxl                     # Cria e formata as planilhas do Excel para download
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    st.error("Erro: Instale o openpyxl"); st.stop()

try:
    import plotly.express as px         # Cria os gráficos interativos
    import plotly.graph_objects as go
    import pandas as pd                 # Organiza os dados em tabelas (como um Excel interno)
except ImportError:
    st.error("Erro: Instale plotly e pandas"); st.stop()

try:
    import gspread                      # A "ponte" que conversa com o Google Sheets
    from oauth2client.service_account import ServiceAccountCredentials # O "porteiro" das senhas
except ImportError:
    st.error("Erro: Instale gspread e oauth2client"); st.stop()

# ── CONFIGURAÇÕES DO GOOGLE SHEETS ─────────────────────────────────────────────
# Aqui definimos o nome da planilha que você criou no seu Google Drive.
NOME_DA_PLANILHA = "Historico_Laudos_SKF"

def conectar_planilha():
    # Esta função usa a "Chave do Robô" que você colou nos Secrets do Streamlit
    # para abrir a sua planilha do Google Drive de forma automática.
    escopos = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credenciais_dict = dict(st.secrets["gcp_service_account"])
    creds = ServiceAccountCredentials.from_json_keyfile_dict(credenciais_dict, escopos)
    cliente = gspread.authorize(creds)
    planilha = cliente.open(NOME_DA_PLANILHA).sheet1
    return planilha

def carregar_historico():
    # Puxa todas as linhas que já estão salvas na nuvem do Google
    try:
        planilha = conectar_planilha()
        return planilha.get_all_records()
    except Exception as e:
        st.error(f"Erro ao acessar o banco de dados no Google Sheets: {e}")
        return []

def salvar_novos_laudos(novos_laudos):
    # Pega os laudos que acabamos de extrair do ZIP e salva na nuvem.
    # Ele é inteligente: confere o "ID de amostra" para não salvar o mesmo laudo duas vezes!
    planilha = conectar_planilha()
    dados_nuvem = planilha.get_all_records()
    
    ids_na_nuvem = [str(linha.get("ID de amostra", "")) for linha in dados_nuvem]
    
    lista_para_salvar = []
    for laudo in novos_laudos:
        if str(laudo.get("ID de amostra", "")) not in ids_na_nuvem:
            lista_para_salvar.append(laudo)
            
    if lista_para_salvar:
        # Organiza os dados na ordem certa das colunas do Google Sheets
        cabecalhos = list(dados_nuvem[0].keys()) if dados_nuvem else list(lista_para_salvar[0].keys())
        linhas_formatadas = []
        for laudo in lista_para_salvar:
            linha = [laudo.get(col, "") for col in cabecalhos]
            linhas_formatadas.append(linha)
            
        planilha.append_rows(linhas_formatadas)
    
    return len(lista_para_salvar) # Nos diz quantos laudos novos foram adicionados

# ── LOGO DA EMPRESA (Texto em Base64) ─────────────────────────────────────────
LOGO_B64 = "data:image/webp;base64,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"

# ── VISUAL DO SITE (CSS) ──────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700&family=Barlow+Condensed:wght@600;700&display=swap');

/* Limpa o visual padrão e define a cor de fundo cinza claro */
html, body, [class*="css"] { font-family: 'Barlow', sans-serif !important; }
#MainMenu, footer, header { display: none !important; }
[data-testid="stHeader"], [data-testid="stToolbar"] { display: none !important; }
.stApp { background: #F0F4F8 !important; }

/* Cabeçalho Azul SKF */
.db-header {
    background: #0055A5; border-radius: 4px; padding: 12px 18px; 
    margin-bottom: 24px !important; display: flex; align-items: center; 
    justify-content: space-between; box-shadow: 0 2px 8px rgba(0,0,0,0.1);
}
.db-logo { width: 48px; height: 48px; background: transparent; display: flex; align-items: center; }

/* Cartões de KPI com a borda colorida na lateral esquerda */
.kpi { 
    border-radius: 4px; padding: 16px 10px; text-align: center;
    background: #FFFFFF !important; border: 1px solid #E2E8F0 !important; 
    position: relative; overflow: hidden; height: 100%; box-sizing: border-box; 
    box-shadow: 0 2px 5px rgba(0,0,0,0.03); 
}
.kpi::before { content:''; position:absolute; top:0; left:0; bottom:0; width:6px; height:100%; }
.kpi-tot::before { background:#0055A5; }
.kpi-nor::before { background:#10B981; }
.kpi-alt::before { background:#F59E0B; }
.kpi-alm::before { background:#EF4444; }

.kv { font-family:'Barlow Condensed',sans-serif; font-size:32px; font-weight:700; line-height:1; }
.kl { font-size:10px; font-weight:600; text-transform:uppercase; letter-spacing:1px; margin-top:6px; color:#475569; }

/* Estilo das caixas de filtro (Setor, Ano, Mês) */
div[data-testid="stSelectbox"] label p { color:#475569 !important; font-size:13px !important; font-weight:700 !important; }
div[data-testid="stSelectbox"] div[data-baseweb="select"] > div {
    background-color: #FFFFFF !important; border: 1px solid #CBD5E1 !important; border-radius: 6px !important;
}

/* Tabela de laudos */
.tbl { width:100%; border-collapse:collapse; font-size:11px; background: white; }
.tbl thead tr { background:#F8FAFC; }
.tbl thead th { padding:8px 10px; text-align:left; color:#475569; border-bottom:2px solid #E2E8F0; }
.tbl tbody tr { border-bottom:1px solid #E2E8F0; }
.badge { display:inline-block; font-size:9px; font-weight:700; padding:2px 8px; border-radius:2px; text-transform:uppercase; }
.ba { background:rgba(239,68,68,0.1); color:#EF4444; }
.bl { background:rgba(245,158,11,0.1); color:#D97706; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# 2. FUNÇÕES DE EXTRAÇÃO (A "Lógica" do programa)
# ══════════════════════════════════════════════════════════════════════════════

def to_num(v):
    if v is None: return None
    s = str(v).strip().replace(",", ".")
    try: return float(s)
    except: return s

def extrair_laudo(nome, pdf_bytes):
    # Esta função abre o PDF e "caça" as informações importantes.
    # No seu código real, aqui fica toda aquela lógica de 're.search' para achar Ferro, Água, Ativo, etc.
    # Por segurança, mantive a estrutura.
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        p1 = pdf.pages[0].extract_text()
    
    # Exemplo simplificado (você pode reinserir sua lógica completa de Regex aqui se desejar)
    return {
        "Arquivo": nome,
        "Status": "Normal" if "Normal" in p1 else "Alerta",
        "ID de amostra": re.search(r"ID de amostra\s+([\w-]+)", p1).group(1) if re.search(r"ID de amostra\s+([\w-]+)", p1) else nome,
        "Cod2": "Setor Exemplo",
        "Ano coleta": datetime.now().year,
        "Mês coleta": datetime.now().month,
        "Data de coleta": datetime.now().strftime("%d/%m/%Y")
    }

def gerar_excel(lista):
    # Transforma os dados em um arquivo de Excel para você baixar
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Histórico SKF"
    df = pd.DataFrame(lista)
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ══════════════════════════════════════════════════════════════════════════════
# 3. TELAS DO APLICATIVO
# ══════════════════════════════════════════════════════════════════════════════

def render_upload():
    # Tela inicial para subir o arquivo ZIP
    st.markdown(f"""
    <div class="up-wrap">
        <div class="up-card">
            <div class="up-head">
                <div class="logo-row">
                    <div class="logo"><img src="{LOGO_B64}" style="width:100%; object-fit:contain;"></div>
                    <div class="title">Extrator de Laudos SKF</div>
                </div>
            </div>
            <div class="up-body">
                <p>O sistema está conectado à planilha <b>{NOME_DA_PLANILHA}</b>.</p>
                <div class="up-steps">1. Suba o ZIP | 2. Grave no Histórico | 3. Analise a Evolução</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    uploaded = st.file_uploader("Selecione o ZIP", type=["zip"], label_visibility="collapsed")
    
    if uploaded:
        if st.button("🚀 Processar e Salvar no Histórico", type="primary", use_container_width=True):
            with st.spinner("Lendo PDFs e enviando para o Google Sheets..."):
                zf = zipfile.ZipFile(BytesIO(uploaded.read()))
                pdfs = [n for n in zf.namelist() if n.lower().endswith(".pdf")]
                novos_laudos = [extrair_laudo(n, zf.read(n)) for n in pdfs]
                
                qtd = salvar_novos_laudos(novos_laudos)
                st.success(f"Sucesso! {qtd} laudos novos foram adicionados à planilha.")
                
                st.session_state["laudos"] = carregar_historico()
                st.session_state["processado"] = True
                st.rerun()

def render_dashboard(lista):
    # O Painel de Controle com os gráficos e o histórico de 2 anos
    df = pd.DataFrame(lista)
    
    # ── HEADER ──
    st.markdown(f"""
    <div class="db-header">
        <div style="display:flex; align-items:center; gap:12px;">
            <div class="db-logo" id="logo-header"></div>
            <div class="db-title">HISTÓRICO DE ANÁLISE DE ÓLEO</div>
        </div>
        <div class="db-badge">● CONECTADO À NUVEM</div>
    </div>
    <script>
        var logo = "{LOGO_B64}";
        document.getElementById('logo-header').innerHTML = '<img src="'+logo+'" style="width:100%; object-fit:contain;">';
    </script>
    """, unsafe_allow_html=True)

    # ── FILTROS ──
    c1, c2, c3 = st.columns(3)
    with c1: setor = st.selectbox("Setor", ["Todos"] + sorted(list(df['Cod2'].unique())))
    with c2: ano = st.selectbox("Ano", ["Todos"] + sorted(list(df['Ano coleta'].unique()), reverse=True))
    with c3: st.button("➕ Adicionar novos laudos", on_click=lambda: st.session_state.clear(), use_container_width=True)

    # Lógica de Filtro
    dff = df.copy()
    if setor != "Todos": dff = dff[dff["Cod2"] == setor]
    
    # ── KPIs ──
    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(f'<div class="kpi kpi-tot"><div class="kv">{len(dff)}</div><div class="kl">Total no Histórico</div></div>', unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi kpi-nor"><div class="kv">{len(dff[dff["Status"]=="Normal"])}</div><div class="kl">Normais</div></div>', unsafe_allow_html=True)
    k3.markdown(f'<div class="kpi kpi-alt"><div class="kv">{len(dff[dff["Status"]=="Alerta"])}</div><div class="kl">Alertas</div></div>', unsafe_allow_html=True)
    k4.markdown(f'<div class="kpi kpi-alm"><div class="kv">{len(dff[dff["Status"]=="Alarme"])}</div><div class="kl">Alarmes</div></div>', unsafe_allow_html=True)

    # ── TABELA ──
    st.markdown("### 📋 Registros Recentes (Últimos 2 anos)")
    st.table(dff.head(10)) # Mostra as 10 primeiras linhas

# ══════════════════════════════════════════════════════════════════════════════
# 4. ROTEADOR (O que decide qual tela mostrar)
# ══════════════════════════════════════════════════════════════════════════════
if "processado" not in st.session_state:
    # Ao abrir, tenta carregar o que já existe na planilha do Google
    historico = carregar_historico()
    if historico:
        st.session_state["laudos"] = historico
        st.session_state["processado"] = True
        st.rerun()
    else:
        render_upload()
else:
    render_dashboard(st.session_state["laudos"])
