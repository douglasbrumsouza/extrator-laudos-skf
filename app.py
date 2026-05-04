import streamlit as st
import zipfile, re, os, base64
from io import BytesIO
from datetime import datetime

# ── Dependências ──────────────────────────────────────────────────────────────
try:
    import pdfplumber
except ImportError:
    st.error("Instale pdfplumber: pip install pdfplumber"); st.stop()
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    st.error("Instale openpyxl: pip install openpyxl"); st.stop()

# ── Config página ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Extrator de Laudos SKF",
    page_icon="🛢️",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ── CSS personalizado ─────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@300;400;500;600;700&family=Barlow+Condensed:wght@600;700&display=swap');

html, body, [class*="css"] { font-family: 'Barlow', sans-serif !important; }

/* Fundo */
.stApp { background: linear-gradient(135deg, #0D2137 0%, #1A3A5C 100%); min-height: 100vh; }

/* Header customizado */
.app-header {
    background: linear-gradient(135deg, #1F4E79, #2E75B6);
    border-radius: 14px; padding: 28px 32px; margin-bottom: 20px;
    box-shadow: 0 8px 32px rgba(0,0,0,0.3);
}
.app-header-row { display:flex; align-items:center; gap:14px; margin-bottom:10px; }
.app-logo {
    width:50px; height:50px; background:rgba(255,255,255,0.15);
    border-radius:12px; display:flex; align-items:center; justify-content:center;
    font-family:'Barlow Condensed',sans-serif; font-size:16px; font-weight:700;
    color:white; border:1px solid rgba(255,255,255,0.2); flex-shrink:0;
}
.app-title { font-family:'Barlow Condensed',sans-serif; font-size:26px; font-weight:700; color:white; letter-spacing:.5px; line-height:1.1; }
.app-sub { font-size:12px; color:rgba(255,255,255,0.6); margin-top:2px; }
.app-desc { font-size:13px; color:rgba(255,255,255,0.78); line-height:1.6; }

/* Card principal */
.main-card {
    background: white; border-radius: 14px; padding: 28px 32px;
    box-shadow: 0 16px 48px rgba(0,0,0,0.25); margin-bottom: 16px;
}

/* KPI cards */
.kpi-row { display:grid; grid-template-columns:repeat(4,1fr); gap:10px; margin:16px 0; }
.kpi { border-radius:10px; padding:14px 16px; text-align:center; border:1px solid transparent; }
.kpi-total { background:#EBF3FB; border-color:#BDD7EE; }
.kpi-normal { background:#C6EFCE; border-color:#A8D5B0; }
.kpi-alerta { background:#FFEB9C; border-color:#F0D060; }
.kpi-alarme { background:#FFC7CE; border-color:#F0A0A8; }
.kpi-val { font-family:'Barlow Condensed',sans-serif; font-size:36px; font-weight:700; line-height:1; }
.kpi-total .kpi-val  { color:#1F4E79; }
.kpi-normal .kpi-val { color:#1A6B2E; }
.kpi-alerta .kpi-val { color:#7D4F00; }
.kpi-alarme .kpi-val { color:#8B0A0A; }
.kpi-lbl { font-size:10px; font-weight:600; text-transform:uppercase; letter-spacing:.8px; margin-top:3px; }
.kpi-total .kpi-lbl  { color:#2E75B6; }
.kpi-normal .kpi-lbl { color:#2A7A3A; }
.kpi-alerta .kpi-lbl { color:#8B6010; }
.kpi-alarme .kpi-lbl { color:#9C0006; }

/* Status badge na tabela */
.badge { display:inline-block; font-size:10px; font-weight:700; padding:3px 10px;
         border-radius:20px; text-transform:uppercase; letter-spacing:.5px; }
.b-normal { background:#C6EFCE; color:#1A6B2E; }
.b-alerta { background:#FFEB9C; color:#7D4F00; }
.b-alarme { background:#FFC7CE; color:#9C0006; }

/* Tabela de log */
.log-box {
    background:#F8FAFC; border:1px solid #D0DCE8; border-radius:8px;
    padding:12px 14px; max-height:200px; overflow-y:auto;
    font-family:'Courier New',monospace; font-size:11px;
}
.log-ok   { color:#1A6B2E; }
.log-err  { color:#8B0A0A; }
.log-info { color:#1F4E79; }

/* Botão de download */
.dl-btn {
    display:block; width:100%; background:#1A6B2E; color:white;
    border:none; border-radius:10px; padding:14px;
    font-family:'Barlow',sans-serif; font-size:15px; font-weight:600;
    text-align:center; cursor:pointer; text-decoration:none;
    box-shadow:0 4px 12px rgba(26,107,46,0.3); transition:all .2s;
}

/* Ocultar elementos padrão do Streamlit */
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding-top: 2rem; padding-bottom: 2rem; max-width: 720px; }
</style>
""", unsafe_allow_html=True)

# ── Lógica de extração ────────────────────────────────────────────────────────
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
    s = str(v).strip().replace(",",".")
    try: return float(s)
    except: return s if s else None

def detectar_status(pdf_bytes):
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        full = "".join(p.extract_text() or "" for p in pdf.pages)
        if "Limites acionados" not in full:
            return "Normal"
        imgs = pdf.pages[0].images
        for img in imgs:
            if img.get("name") == "Im2":
                return "Alarme" if img["height"] >= 50 else "Alerta"
    return "Alerta"

def split_ativo(raw):
    partes = [p.strip() for p in raw.split("→")]
    segs = partes[0].strip().split("-")
    c = {f"Cod{i}": segs[i-1] if len(segs) >= i else "" for i in range(1, 6)}
    c["Equipamento Descrição 1"] = partes[1] if len(partes) > 1 else ""
    c["Equipamento Descrição 2"] = partes[2] if len(partes) > 2 else ""
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
    dados["Ativo (completo)"] = raw
    dados.update(split_ativo(raw))

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
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "Laudos"
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
            else:
                c.fill = fb
    larg = {"Arquivo": 35, "Status": 10, "Localização": 28, "Ativo (completo)": 45,
            "Cod1": 7, "Cod2": 7, "Cod3": 8, "Cod4": 9, "Cod5": 7,
            "Equipamento Descrição 1": 26, "Equipamento Descrição 2": 26,
            "Tipo de componente": 28, "Óleo": 24, "Data de coleta": 20,
            "Dia coleta": 8, "Mês coleta": 8, "Ano coleta": 8, "Hora coleta": 10,
            "Observações": 40, "Diagnósticos": 50, "Ações": 50, "Parâmetros em Alarme": 30}
    for ci, col in enumerate(colunas, 1):
        ws.column_dimensions[get_column_letter(ci)].width = larg.get(col, 14)
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return buf.read()

# ══════════════════════════════════════════════════════════════════════════════
# INTERFACE
# ══════════════════════════════════════════════════════════════════════════════

# Header
st.markdown("""
<div class="app-header">
  <div class="app-header-row">
    <div class="app-logo">SKF</div>
    <div>
      <div class="app-title">Extrator de Laudos SKF</div>
      <div class="app-sub">TruVu 360 · Gerdau Charqueadas · Engenharia de Manutenção</div>
    </div>
  </div>
  <div class="app-desc">
    Faça o upload do arquivo ZIP com os laudos em PDF.<br>
    O Excel consolidado será gerado automaticamente para importação no Power BI.
  </div>
</div>
""", unsafe_allow_html=True)

# Upload
uploaded = st.file_uploader(
    "📦 Selecione o arquivo ZIP com os laudos",
    type=["zip"],
    help="Arquivo .zip enviado pelo laboratório SKF contendo os PDFs dos laudos"
)

if uploaded:
    st.success(f"✅ **{uploaded.name}** selecionado — {uploaded.size/1024/1024:.1f} MB")

    if st.button("⚙️ Processar Laudos", type="primary", use_container_width=True):
        with st.spinner("Abrindo arquivo ZIP..."):
            zip_data = uploaded.read()
            zf = zipfile.ZipFile(BytesIO(zip_data))
            pdfs = [(n, zf.read(n)) for n in zf.namelist()
                    if n.lower().endswith(".pdf") and not n.startswith("__MACOSX")]

        if not pdfs:
            st.error("Nenhum PDF encontrado no ZIP. Verifique o arquivo.")
            st.stop()

        st.info(f"📄 **{len(pdfs)} laudos** encontrados no ZIP")

        # Progresso
        progress = st.progress(0, text="Iniciando processamento...")
        log_area = st.empty()
        log_lines = []

        lista, erros = [], []
        for i, (nome, pdf_bytes) in enumerate(sorted(pdfs), 1):
            nome_curto = os.path.basename(nome)
            pct = int(i / len(pdfs) * 85)
            progress.progress(pct / 100, text=f"Processando {i}/{len(pdfs)}: {nome_curto[:50]}")

            try:
                d = extrair_laudo(nome_curto, pdf_bytes)
                lista.append(d)
                cor = {"Normal": "🟢", "Alerta": "🟡", "Alarme": "🔴"}.get(d["Status"], "⚪")
                log_lines.append(f"{cor} [{d['Status']:7s}] {d['Cod2']:5s} | {nome_curto[:55]}")
            except Exception as e:
                erros.append(nome_curto)
                log_lines.append(f"❌ ERRO: {nome_curto[:45]} — {e}")

            log_area.code("\n".join(log_lines[-12:]), language=None)

        progress.progress(0.95, text="Gerando Excel...")

        excel_bytes = gerar_excel(lista)
        nome_excel = f"LAUDOS_{datetime.now().strftime('%d-%m-%Y')}.xlsx"

        progress.progress(1.0, text="Concluído!")

        # ── Estatísticas ──────────────────────────────────────────────────
        total  = len(lista)
        normal = sum(1 for d in lista if d.get("Status") == "Normal")
        alerta = sum(1 for d in lista if d.get("Status") == "Alerta")
        alarme = sum(1 for d in lista if d.get("Status") == "Alarme")

        st.markdown(f"""
        <div class="kpi-row">
          <div class="kpi kpi-total">
            <div class="kpi-val">{total}</div>
            <div class="kpi-lbl">Total</div>
          </div>
          <div class="kpi kpi-normal">
            <div class="kpi-val">{normal}</div>
            <div class="kpi-lbl">Normal</div>
          </div>
          <div class="kpi kpi-alerta">
            <div class="kpi-val">{alerta}</div>
            <div class="kpi-lbl">Alerta</div>
          </div>
          <div class="kpi kpi-alarme">
            <div class="kpi-val">{alarme}</div>
            <div class="kpi-lbl">Alarme</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        if erros:
            st.warning(f"⚠️ {len(erros)} laudo(s) com erro: {', '.join(erros)}")

        # ── Download ──────────────────────────────────────────────────────
        st.download_button(
            label=f"⬇️  Baixar Excel — {nome_excel}",
            data=excel_bytes,
            file_name=nome_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )

        st.caption("💡 Salve o Excel na pasta do Power BI e clique em **Atualizar** para sincronizar o dashboard.")

else:
    # Estado inicial — instruções
    st.markdown("""
    <div style="background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.1);
                border-radius:12px;padding:24px 28px;color:rgba(255,255,255,0.8);">
      <div style="font-size:14px;font-weight:600;color:white;margin-bottom:14px;">Como usar</div>
      <div style="display:flex;flex-direction:column;gap:10px;font-size:13px;line-height:1.5">
        <div>📦 &nbsp;<b style="color:white">1.</b> Receba o ZIP com os laudos do laboratório</div>
        <div>⬆️ &nbsp;<b style="color:white">2.</b> Faça o upload do arquivo ZIP acima</div>
        <div>⚙️ &nbsp;<b style="color:white">3.</b> Clique em <b style="color:white">Processar Laudos</b></div>
        <div>⬇️ &nbsp;<b style="color:white">4.</b> Baixe o Excel consolidado</div>
        <div>📊 &nbsp;<b style="color:white">5.</b> Atualize o Power BI com o novo arquivo</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div style="text-align:center;margin-top:16px;font-size:11px;color:rgba(255,255,255,0.35);">
      Nenhum dado é armazenado · Processamento seguro · Gerdau Charqueadas
    </div>
    """, unsafe_allow_html=True)
