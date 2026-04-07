# DOM safe version
import os
import runpy
from pathlib import Path
import streamlit as st
import pandas as pd

BASE_DIR = Path(__file__).parent
LEGACY_DIR = BASE_DIR / "legacy_apps"

st.set_page_config(
    page_title="NEW ORDER CONTROL CENTER",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

def run_legacy_app(path: Path):
    original_cwd = Path.cwd()
    original_fn = st.set_page_config
    try:
        os.chdir(path.parent)
        st.set_page_config = lambda *args, **kwargs: None
        runpy.run_path(str(path), run_name="__main__")
    finally:
        st.set_page_config = original_fn
        os.chdir(original_cwd)

def inject_theme():
    st.markdown("""
    <style>
    .stApp {
        background:
            radial-gradient(circle at 12% 10%, rgba(0,195,255,0.08), transparent 18%),
            radial-gradient(circle at 88% 14%, rgba(255,120,0,0.08), transparent 18%),
            linear-gradient(135deg, #071018 0%, #0b131d 40%, #0b1118 100%);
        color: #edf3f9;
    }
    section[data-testid="stSidebar"] {
        background:
            linear-gradient(180deg, rgba(255,255,255,0.04), rgba(255,255,255,0.015)),
            linear-gradient(180deg, #09111a 0%, #111a25 100%);
        border-right: 1px solid rgba(255,255,255,0.08);
    }
    .block-container {
        padding-top: 1.1rem;
        padding-bottom: 2rem;
        max-width: 1450px;
    }
    .hero {
        position: relative;
        overflow: hidden;
        border-radius: 28px;
        padding: 34px 36px;
        margin-bottom: 22px;
        background:
            linear-gradient(135deg, rgba(255,255,255,0.07), rgba(255,255,255,0.02)),
            linear-gradient(120deg, rgba(0,195,255,0.06), rgba(255,120,0,0.05));
        border: 1px solid rgba(255,255,255,0.08);
        box-shadow:
            0 22px 46px rgba(0,0,0,0.34),
            inset 0 1px 0 rgba(255,255,255,0.10);
    }
    .hero:before {
        content: "";
        position: absolute;
        inset: 0;
        background:
            repeating-linear-gradient(
                120deg,
                rgba(255,255,255,0.018) 0px,
                rgba(255,255,255,0.018) 2px,
                transparent 2px,
                transparent 14px
            );
        pointer-events: none;
    }
    .hero-title {
        font-size: 50px;
        font-weight: 950;
        line-height: 0.98;
        letter-spacing: 1px;
        color: #ffffff;
        margin-bottom: 12px;
        text-shadow: 0 0 18px rgba(0,195,255,0.08);
    }
    .hero-subtitle {
        color: #b9c5d2;
        font-size: 14px;
        text-transform: uppercase;
        letter-spacing: 1.6px;
        margin-bottom: 16px;
    }
    .hero-line {
        width: 220px;
        height: 3px;
        border-radius: 999px;
        background: linear-gradient(90deg, #00c3ff, #ff7800);
        margin-bottom: 18px;
        box-shadow: 0 0 16px rgba(0,195,255,0.25);
    }
    .hero-text {
        color: #d8e1ea;
        font-size: 15px;
        line-height: 1.65;
        max-width: 880px;
        margin-bottom: 18px;
    }
    .hero-badge {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        padding: 8px 12px;
        border-radius: 999px;
        background: rgba(255,255,255,0.06);
        border: 1px solid rgba(255,255,255,0.08);
        color: #e8f0f7;
        font-size: 12px;
        font-weight: 700;
        margin: 6px 8px 0 0;
    }
    .section-title {
        color: #ffffff;
        font-size: 24px;
        font-weight: 900;
        margin: 10px 0 14px 0;
    }
    .module-card {
        position: relative;
        overflow: hidden;
        border-radius: 24px;
        padding: 22px;
        min-height: 210px;
        background:
            linear-gradient(180deg, rgba(255,255,255,0.06), rgba(255,255,255,0.02)),
            linear-gradient(135deg, rgba(0,195,255,0.03), rgba(255,120,0,0.03));
        border: 1px solid rgba(255,255,255,0.08);
        box-shadow:
            0 14px 34px rgba(0,0,0,0.25),
            inset 0 1px 0 rgba(255,255,255,0.06);
    }
    .module-card:before {
        content: "";
        position: absolute;
        left: 0;
        top: 0;
        width: 100%;
        height: 4px;
        background: linear-gradient(90deg, #00c3ff, #ff7800);
        opacity: 0.95;
    }
    .module-icon { font-size: 34px; margin-bottom: 8px; }
    .module-tag {
        display: inline-block; padding: 6px 10px; border-radius: 999px; font-size: 11px;
        font-weight: 800; text-transform: uppercase; letter-spacing: 1px; color: #dce7f2;
        background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.08); margin-bottom: 14px;
    }
    .module-title { color: #ffffff; font-size: 25px; font-weight: 900; margin-bottom: 10px; }
    .module-desc { color: #a9b8c8; font-size: 13px; line-height: 1.55; }
    div.stButton > button {
        width: 100%; min-height: 78px !important; border-radius: 18px !important;
        font-size: 18px !important; font-weight: 900 !important;
        border: 1px solid rgba(255,255,255,0.10) !important;
        background: linear-gradient(135deg, rgba(0,195,255,0.14), rgba(255,120,0,0.10)) !important;
        color: #ffffff !important; box-shadow: 0 14px 28px rgba(0,0,0,0.22);
    }
    div.stButton > button:hover {
        border-color: rgba(255,255,255,0.18) !important;
        box-shadow: 0 18px 34px rgba(0,0,0,0.28), 0 0 16px rgba(0,195,255,0.10);
    }
    .kpi-card {
        border-radius: 22px; padding: 18px; min-height: 160px;
        background: linear-gradient(180deg, rgba(255,255,255,0.05), rgba(255,255,255,0.02));
        border: 1px solid rgba(255,255,255,0.08); box-shadow: 0 12px 28px rgba(0,0,0,0.22);
    }
    .kpi-label {
        color: #9fb0c2; font-size: 12px; text-transform: uppercase; font-weight: 800;
        letter-spacing: 1px; margin-bottom: 8px;
    }
    .kpi-value {
        color: #ffffff; font-size: 34px; font-weight: 950; line-height: 1.0; margin-bottom: 8px;
    }
    .kpi-sub { color: #9aabbd; font-size: 12px; line-height: 1.45; }
    .sem-pill {
        display:inline-flex; align-items:center; gap:8px; padding:6px 10px; border-radius:999px;
        margin-top:10px; font-size:11px; font-weight:800; text-transform:uppercase; letter-spacing:1px;
        border:1px solid rgba(255,255,255,0.10); background: rgba(255,255,255,0.05); color:#edf3f9;
    }
    .sem-dot {
        width:10px; height:10px; border-radius:50%; display:inline-block;
        box-shadow: 0 0 10px currentColor;
    }
    .green { background:#21d07a; color:#21d07a; }
    .orange { background:#ffb020; color:#ffb020; }
    .red { background:#ff5a5f; color:#ff5a5f; }
    .info-card {
        border-radius: 24px; padding: 20px; min-height: 230px;
        background: linear-gradient(180deg, rgba(255,255,255,0.05), rgba(255,255,255,0.02));
        border: 1px solid rgba(255,255,255,0.08); box-shadow: 0 12px 28px rgba(0,0,0,0.22);
    }
    .info-title { color: #ffffff; font-size: 20px; font-weight: 900; margin-bottom: 8px; }
    .info-item { border-left: 3px solid rgba(0,195,255,0.55); padding-left: 14px; margin: 14px 0; }
    .info-head { color: #ffffff; font-size: 15px; font-weight: 800; }
    .info-text { color: #9aabbd; font-size: 12px; margin-top: 4px; line-height: 1.45; }
    </style>
    """, unsafe_allow_html=True)

def read_excel_safe(path: Path, **kwargs):
    try:
        return pd.read_excel(path, engine="openpyxl", **kwargs)
    except Exception:
        return pd.DataFrame()

def find_col(df: pd.DataFrame, keywords):
    if df is None or df.empty:
        return None
    cols = [str(c).strip() for c in df.columns]
    lower_map = {str(c).strip().lower(): c for c in df.columns}
    for key in keywords:
        for c in cols:
            if key.lower() == c.lower():
                return lower_map[c.lower()]
    for key in keywords:
        for c in cols:
            if key.lower() in c.lower():
                return lower_map[c.lower()]
    return None

def parse_date_series(s):
    try:
        return pd.to_datetime(s, errors="coerce", dayfirst=True)
    except Exception:
        return pd.to_datetime(pd.Series([], dtype="object"))

def semaforo(status: str):
    status = (status or "").lower()
    if status == "bom":
        return '<span class="sem-pill"><span class="sem-dot green"></span>Bom</span>'
    if status == "atenção" or status == "atencao":
        return '<span class="sem-pill"><span class="sem-dot orange"></span>Atenção</span>'
    return '<span class="sem-pill"><span class="sem-dot red"></span>Crítico</span>'

def get_producao_kpis():
    path = LEGACY_DIR / "producao" / "PROD-PRODT.xlsx"
    if not path.exists():
        return {"produzido_mes": "-", "faturado_mes": "-", "forecast_mes": "-", "status": "crítico", "sub": "Base não encontrada"}
    xl = pd.ExcelFile(path, engine="openpyxl")
    frames = []
    for sheet in xl.sheet_names[:5]:
        try:
            df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
            frames.append(df)
        except Exception:
            pass
    if not frames:
        return {"produzido_mes": "-", "faturado_mes": "-", "forecast_mes": "-", "status": "crítico", "sub": "Sem leitura"}
    df = pd.concat(frames, ignore_index=True)

    prod_col = find_col(df, ["produzido", "produção", "qtde produzida", "qtd produzida"])
    fat_col = find_col(df, ["faturado", "faturamento"])
    fore_col = find_col(df, ["forecast", "previsão", "previsao"])
    date_col = find_col(df, ["data", "dia", "mês", "mes"])

    if date_col:
        dt = parse_date_series(df[date_col])
        hoje = pd.Timestamp.today()
        if dt.notna().any():
            df = df[(dt.dt.month == hoje.month) & (dt.dt.year == hoje.year)]

    def soma(col):
        if not col or col not in df.columns:
            return 0
        return pd.to_numeric(df[col], errors="coerce").fillna(0).sum()

    produzido = soma(prod_col)
    faturado = soma(fat_col)
    forecast = soma(fore_col)

    if forecast <= 0:
        status = "atenção"
    else:
        ratio = produzido / forecast
        status = "bom" if ratio >= 1 else ("atenção" if ratio >= 0.85 else "crítico")

    return {
        "produzido_mes": f"{produzido:,.0f}".replace(",", "."),
        "faturado_mes": f"{faturado:,.0f}".replace(",", "."),
        "forecast_mes": f"{forecast:,.0f}".replace(",", "."),
        "status": status,
        "sub": "Produção do mês com base Excel",
    }

def get_carga_kpis():
    path = LEGACY_DIR / "carga_maquina" / "CG BOT PY.xlsx"
    if not path.exists():
        return {"gargalo": "-", "utilizacao": "-", "linhas": "-", "status": "crítico", "sub": "Base não encontrada"}
    df = read_excel_safe(path)
    if df.empty:
        return {"gargalo": "-", "utilizacao": "-", "linhas": "-", "status": "crítico", "sub": "Sem leitura"}
    tempo_col = find_col(df, ["tempo individual", "tempo", "min"])
    desc_col = df.columns[5] if df.shape[1] > 5 else None
    cr_col = find_col(df, ["cr"])

    tempos = pd.to_numeric(df[tempo_col], errors="coerce").fillna(0) if tempo_col else pd.Series([0] * len(df))
    if desc_col:
        agg = df.assign(_tempo=tempos).groupby(desc_col, dropna=False)["_tempo"].sum().sort_values(ascending=False)
        gargalo = str(agg.index[0]) if len(agg) else "-"
        util = float(agg.iloc[0]) if len(agg) else 0
    else:
        gargalo = "-"
        util = 0
    linhas = int(df[cr_col].nunique()) if cr_col and cr_col in df.columns else int(len(df))

    status = "bom" if util < 5000 else ("atenção" if util < 12000 else "crítico")
    return {
        "gargalo": gargalo,
        "utilizacao": f"{util:,.0f}".replace(",", "."),
        "linhas": str(linhas),
        "status": status,
        "sub": "Leitura inicial da carga máquina",
    }

def get_rh_kpis():
    path = LEGACY_DIR / "rh" / "QUADRO COLABORADORES.xlsx"
    if not path.exists():
        return {"efetivo": "-", "abs_mes": "-", "turn_mes": "-", "status": "crítico", "sub": "Base não encontrada"}
    base = read_excel_safe(path, sheet_name="RELAÇÃO DE COLABORADORES 19")
    abs_df = read_excel_safe(path, sheet_name="ABSENTEISMO")
    turn_df = read_excel_safe(path, sheet_name="TURNOVER")

    efetivo = len(base) if not base.empty else 0
    hoje = pd.Timestamp.today()

    def conta_mes(df, col_name):
        if df.empty:
            return 0
        col = find_col(df, [col_name, "data"])
        if not col:
            return 0
        dt = parse_date_series(df[col])
        return int(((dt.dt.month == hoje.month) & (dt.dt.year == hoje.year)).sum())

    abs_mes = conta_mes(abs_df, "data")
    turn_mes = conta_mes(turn_df, "data")

    abs_rate = (abs_mes / efetivo * 100) if efetivo else 0
    turn_rate = (turn_mes / efetivo * 100) if efetivo else 0
    worst = max(abs_rate, turn_rate)
    status = "bom" if worst <= 2 else ("atenção" if worst <= 5 else "crítico")

    return {
        "efetivo": str(efetivo),
        "abs_mes": str(abs_mes),
        "turn_mes": str(turn_mes),
        "status": status,
        "sub": "Indicadores de pessoas via Excel",
    }

def get_pa_kpis():
    path = LEGACY_DIR / "pa" / "data" / "pa.xlsx"
    if not path.exists():
        path = LEGACY_DIR / "pa" / "pa.xlsx"
    if not path.exists():
        return {"abertas": "-", "atrasadas": "-", "responsavel": "-", "status": "crítico", "sub": "Base não encontrada"}

    try:
        raw = pd.read_excel(path, sheet_name="PA", engine="openpyxl", header=None).dropna(how="all")
    except Exception:
        return {"abertas": "-", "atrasadas": "-", "responsavel": "-", "status": "crítico", "sub": "Sem leitura"}

    header_row = None
    for i, row in raw.iterrows():
        vals = row.astype(str)
        if vals.str.contains("Ação", case=False, na=False).any() and vals.str.contains("Status", case=False, na=False).any():
            header_row = i
            break
    if header_row is None:
        return {"abertas": "-", "atrasadas": "-", "responsavel": "-", "status": "crítico", "sub": "Cabeçalho não localizado"}

    df = pd.read_excel(path, sheet_name="PA", engine="openpyxl", header=header_row).dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]

    status_col = find_col(df, ["status"])
    prazo_col = find_col(df, ["prazo"])
    resp_col = find_col(df, ["responsável", "responsavel"])

    if status_col and status_col in df.columns:
        status = df[status_col].astype(str).str.strip().str.lower()
        abertas = int(status.isin(["aberta", "em execução", "em execucao", "atrasada", "open"]).sum())
    else:
        abertas = len(df)

    atrasadas = 0
    if prazo_col and prazo_col in df.columns:
        dt = parse_date_series(df[prazo_col])
        if status_col and status_col in df.columns:
            closed = df[status_col].astype(str).str.lower().isin(["executado", "cancelada"])
            atrasadas = int(((dt < pd.Timestamp.today().normalize()) & (~closed)).sum())
        else:
            atrasadas = int((dt < pd.Timestamp.today().normalize()).sum())

    responsavel = "-"
    if resp_col and resp_col in df.columns:
        top = df[resp_col].astype(str).str.strip()
        top = top[top.ne("") & top.ne("nan")]
        if not top.empty:
            responsavel = top.value_counts().index[0]

    status = "bom" if atrasadas == 0 else ("atenção" if atrasadas <= 3 else "crítico")
    return {
        "abertas": str(abertas),
        "atrasadas": str(atrasadas),
        "responsavel": responsavel,
        "status": status,
        "sub": "Indicadores do plano de ação",
    }

def kpi_card(label: str, value: str, sub: str, status: str):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-sub">{sub}</div>
            {semaforo(status)}
        </div>
        """,
        unsafe_allow_html=True,
    )

def module_card(icon: str, tag: str, title: str, desc: str):
    st.markdown(
        f"""
        <div class="module-card">
            <div class="module-icon">{icon}</div>
            <div class="module-tag">{tag}</div>
            <div class="module-title">{title}</div>
            <div class="module-desc">{desc}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

def home():
    logo_path = LEGACY_DIR / "producao" / "logo.png"
    if logo_path.exists():
        st.image(str(logo_path), width=210)

    prod = get_producao_kpis()
    carga = get_carga_kpis()
    rh = get_rh_kpis()
    pa = get_pa_kpis()

    st.markdown(
        """
        <div class="hero">
            <div class="hero-title">NEW ORDER<br>CONTROL CENTER</div>
            <div class="hero-subtitle">Central industrial integrada • visão executiva de fábrica</div>
            <div class="hero-line"></div>
            <div class="hero-text">
                Plataforma corporativa para gestão de produção, plano de ação, carga máquina e pessoas,
                consolidando leitura operacional, indicadores e navegação entre módulos em um único ambiente.
            </div>
            <span class="hero-badge">🔵 Produção</span>
            <span class="hero-badge">🟠 Planejamento</span>
            <span class="hero-badge">🟢 Pessoas</span>
            <span class="hero-badge">🔴 Alertas operacionais</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-title">Módulos Principais</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        if st.button("🏭 PRODUÇÃO", use_container_width=True):
            st.session_state["menu_target"] = "Produção"
            st.rerun()
    with c2:
        if st.button("📋 PLANO DE AÇÃO", use_container_width=True):
            st.session_state["menu_target"] = "Plano de Ação Fábrica"
            st.rerun()
    with c3:
        if st.button("⚙️ CARGA MÁQUINA", use_container_width=True):
            st.session_state["menu_target"] = "Carga Máquina"
            st.rerun()
    with c4:
        if st.button("👥 RH • PESSOAS", use_container_width=True):
            st.session_state["menu_target"] = "RH • Pessoas"
            st.rerun()
    st.markdown('<div class="section-title">KPIs Executivos da Home</div>', unsafe_allow_html=True)
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        kpi_card("Produzido mês", prod["produzido_mes"], f'Forecast: {prod["forecast_mes"]} • Faturado: {prod["faturado_mes"]}', prod["status"])
    with k2:
        kpi_card("Gargalo atual", carga["gargalo"], f'Carga crítica: {carga["utilizacao"]} • Linhas/CR: {carga["linhas"]}', carga["status"])
    with k3:
        kpi_card("Plano de ação", pa["abertas"], f'Atrasadas: {pa["atrasadas"]} • Resp.: {pa["responsavel"]}', pa["status"])
    with k4:
        kpi_card("Efetivo RH", rh["efetivo"], f'Absenteísmo mês: {rh["abs_mes"]} • Turnover mês: {rh["turn_mes"]}', rh["status"])

    st.markdown('<div class="section-title">Visão dos Módulos</div>', unsafe_allow_html=True)
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        module_card("🏭", "Operação", "Produção", "Dashboard produtivo com produtividade, tendência mensal, forecast x produzido x faturado, rankings e leitura direta da base operacional.")
    with m2:
        module_card("📋", "Gestão à vista", "Plano de Ação Fábrica", "Gemba Board ANDON com KPIs, Pareto, Kanban, saúde do plano de ação, atrasos críticos e acompanhamento por responsável e setor.")
    with m3:
        module_card("⚙️", "Capacidade", "Carga Máquina", "Simulação industrial com gargalo matemático real, mão de obra, indiretos, TAKT, previsão de cenário e análise executiva de utilização.")
    with m4:
        module_card("👥", "Pessoas", "RH • Pessoas", "Controle de absenteísmo, turnover e novas contratações com lançamento diário, indicadores mensais e análise por setor.")

    i1, i2 = st.columns([1.15, 0.85], gap="large")
    with i1:
        st.markdown(
            """
            <div class="info-card">
                <div class="info-title">Centro de Comando</div>
                <div class="info-item">
                    <div class="info-head">Semáforos automáticos</div>
                    <div class="info-text">A Home agora classifica os indicadores em bom, atenção ou crítico conforme leitura dos Excel.</div>
                </div>
                <div class="info-item">
                    <div class="info-head">Leitura imediata</div>
                    <div class="info-text">A gestão visualiza rapidamente quais áreas precisam de foco antes de entrar nos módulos detalhados.</div>
                </div>
                <div class="info-item">
                    <div class="info-head">Base pronta para metas</div>
                    <div class="info-text">A próxima evolução pode usar metas configuráveis por módulo para refinar os alertas.</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with i2:
        st.markdown(
            """
            <div class="info-card">
                <div class="info-title">Leitura Executiva</div>
                <div class="info-item">
                    <div class="info-head">Produção</div>
                    <div class="info-text">Compara produzido x forecast e já classifica o mês.</div>
                </div>
                <div class="info-item">
                    <div class="info-head">Carga e Ações</div>
                    <div class="info-text">Mostra gargalo atual, criticidade da carga e situação do plano de ação.</div>
                </div>
                <div class="info-item">
                    <div class="info-head">Pessoas</div>
                    <div class="info-text">Consolida efetivo, absenteísmo e turnover com leitura simples para gestão.</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

inject_theme()
logo_sidebar = LEGACY_DIR / "producao" / "logo.png"
if logo_sidebar.exists():
    st.sidebar.image(str(logo_sidebar), width=150)
st.sidebar.markdown("## GNO • Control Center")

options = ["Home", "Produção", "Plano de Ação Fábrica", "Carga Máquina", "RH • Pessoas"]
target = st.session_state.pop("menu_target", None)
default_idx = options.index(target) if target in options else 0
menu = st.sidebar.radio("Navegação", options, index=default_idx)

if menu == "Home":
    home()
elif menu == "Produção":
    run_legacy_app(LEGACY_DIR / "producao" / "app.py")
elif menu == "Plano de Ação Fábrica":
    run_legacy_app(LEGACY_DIR / "pa" / "app.py")
elif menu == "Carga Máquina":
    run_legacy_app(LEGACY_DIR / "carga_maquina" / "app.py")
elif menu == "RH • Pessoas":
    run_legacy_app(LEGACY_DIR / "rh" / "app.py")
