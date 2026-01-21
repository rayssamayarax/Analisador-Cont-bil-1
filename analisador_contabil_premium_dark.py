import streamlit as st
import pandas as pd
import re
from datetime import datetime
import io
import base64

# ==================== MOTORES DE PROCESSAMENTO ====================

def parse_num(v):
    """Converte string numérica para float"""
    if pd.isna(v):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        s = str(v).strip()
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", ".")
        return float(s)
    except Exception:
        return 0.0

# ==================== MOTOR 1: RAZÃO DE FORNECEDORES (ORIGINAL) ====================

def eh_conta(v):
    """Verifica se é uma conta no formato"""
    return isinstance(v, str) and re.match(r"^\d+\s*-\s*", v.strip())

def texto_tem_saldo_anterior(v):
    return isinstance(v, str) and ("saldo anterior" in v.lower())

def processar_fornecedores(df):
    """Motor ORIGINAL para razão de fornecedores"""
    conta_atual = None
    data_atual = None
    saldo_inicial = {}
    movimentos_por_dia = {}

    for _, row in df.iterrows():
        hist = row["Histórico"]
        chave = row["Chave"]
        contra = row["Contra"]
        valor = row["Valor"]
        saldo = row["Saldo"]

        if eh_conta(hist):
            conta_atual = str(hist).strip()
            data_atual = None
            saldo_inicial.setdefault(conta_atual, 0.0)
            movimentos_por_dia.setdefault(conta_atual, {})

            if texto_tem_saldo_anterior(valor):
                saldo_inicial[conta_atual] = parse_num(saldo)
            continue

        if not conta_atual:
            continue

        if texto_tem_saldo_anterior(valor):
            saldo_inicial[conta_atual] = parse_num(saldo)
            continue

        if isinstance(hist, datetime):
            data_atual = hist.date()
            continue

        if not data_atual:
            continue

        if isinstance(hist, str) and "total dia" in hist.lower():
            deb = parse_num(chave)
            cred = parse_num(valor)

            movimentos_por_dia[conta_atual].setdefault(data_atual, {"deb": 0.0, "cred": 0.0})
            movimentos_por_dia[conta_atual][data_atual]["deb"] += deb
            movimentos_por_dia[conta_atual][data_atual]["cred"] += cred

    resumo = {}
    detalhado = []

    for conta, dias_dict in movimentos_por_dia.items():
        saldo_atual = float(saldo_inicial.get(conta, 0.0))

        for data in sorted(dias_dict.keys()):
            deb = float(dias_dict[data]["deb"])
            cred = float(dias_dict[data]["cred"])

            saldo_antes_do_dia = saldo_atual
            saldo_atual = saldo_atual - deb + cred

            if saldo_atual < 0:
                resumo.setdefault(conta, {"Qtd dias com erro": 0, "Primeira data": None})
                resumo[conta]["Qtd dias com erro"] += 1
                if resumo[conta]["Primeira data"] is None:
                    resumo[conta]["Primeira data"] = data.strftime("%d/%m/%Y")

                detalhado.append({
                    "Conta": conta,
                    "Data": data.strftime("%d/%m/%Y"),
                    "Saldo Anterior": saldo_antes_do_dia,
                    "Débito do Dia": deb,
                    "Crédito do Dia": cred,
                    "Saldo Após o Dia": saldo_atual
                })

    resumo_df = pd.DataFrame([
        {
            "Conta": c,
            "Qtd dias com saldo negativo": d["Qtd dias com erro"],
            "Primeira data com erro": d["Primeira data"]
        }
        for c, d in resumo.items()
    ])

    detalhado_df = pd.DataFrame(detalhado) if detalhado else pd.DataFrame()

    return resumo_df, detalhado_df

# ==================== MOTOR 2: RAZÃO DE CONTAS COMUNS (NOVO) ====================

def processar_contas_comuns(df):
    """Motor NOVO para razão de contas comuns"""
    contas_saldo_inicial_negativo = {}

    conta_atual = None
    data_atual = None
    saldo_inicial = {}
    movimentos_por_dia = {}

    for _, row in df.iterrows():
        hist = row["Histórico"]
        chave = row["Chave"]
        contra = row["Contra"]
        valor = row["Valor"]
        saldo = row["Saldo"]

        if eh_conta(hist):
            conta_atual = str(hist).strip()
            data_atual = None

            if conta_atual not in saldo_inicial:
                saldo_inicial[conta_atual] = 0.0
                movimentos_por_dia[conta_atual] = {}

            if texto_tem_saldo_anterior(valor):
                saldo_valor = parse_num(saldo)
                saldo_inicial[conta_atual] = saldo_valor
                if saldo_valor < 0:
                    contas_saldo_inicial_negativo[conta_atual] = saldo_valor
            continue

        if not conta_atual:
            continue

        if texto_tem_saldo_anterior(valor):
            saldo_valor = parse_num(saldo)
            saldo_inicial[conta_atual] = saldo_valor
            if saldo_valor < 0:
                contas_saldo_inicial_negativo[conta_atual] = saldo_valor
            continue

        if isinstance(hist, datetime):
            data_atual = hist.date()
            continue

        if not data_atual:
            continue

        if isinstance(hist, str) and "total dia" in hist.lower():
            deb = parse_num(chave)
            cred = parse_num(valor)

            if data_atual not in movimentos_por_dia[conta_atual]:
                movimentos_por_dia[conta_atual][data_atual] = {"deb": 0.0, "cred": 0.0}

            movimentos_por_dia[conta_atual][data_atual]["deb"] += deb
            movimentos_por_dia[conta_atual][data_atual]["cred"] += cred

    resumo = {}
    detalhado = []

    for conta, saldo_valor in contas_saldo_inicial_negativo.items():
        if conta not in resumo:
            resumo[conta] = {
                "Qtd dias com erro": 1,
                "Primeira data": "Saldo inicial",
                "Tem saldo inicial negativo": True
            }
        detalhado.append({
            "Conta": conta,
            "Data": "Saldo inicial",
            "Tipo de erro": "Saldo inicial negativo",
            "Saldo Anterior": 0.0,
            "Débito do Dia": 0.0,
            "Crédito do Dia": 0.0,
            "Saldo Após o Dia": saldo_valor
        })

    for conta, dias_dict in movimentos_por_dia.items():
        saldo_atual = float(saldo_inicial.get(conta, 0.0))

        for data in sorted(dias_dict.keys()):
            deb = float(dias_dict[data]["deb"])
            cred = float(dias_dict[data]["cred"])

            saldo_antes = saldo_atual
            saldo_atual = saldo_atual - deb + cred

            if saldo_atual < 0:
                if conta not in resumo:
                    resumo[conta] = {
                        "Qtd dias com erro": 0,
                        "Primeira data": None,
                        "Tem saldo inicial negativo": False
                    }

                resumo[conta]["Qtd dias com erro"] += 1
                if resumo[conta]["Primeira data"] is None:
                    resumo[conta]["Primeira data"] = data.strftime("%d/%m/%Y")

                detalhado.append({
                    "Conta": conta,
                    "Data": data.strftime("%d/%m/%Y"),
                    "Tipo de erro": "Saldo negativo no período",
                    "Saldo Anterior": saldo_antes,
                    "Débito do Dia": deb,
                    "Crédito do Dia": cred,
                    "Saldo Após o Dia": saldo_atual
                })

    resumo_data = []
    for conta, info in resumo.items():
        resumo_data.append({
            "Conta": conta,
            "Qtd dias com saldo negativo": info["Qtd dias com erro"],
            "Primeira data com erro": info["Primeira data"],
            "Tem saldo inicial negativo": "✓" if info.get("Tem saldo inicial negativo", False) else ""
        })

    resumo_df = pd.DataFrame(resumo_data) if resumo_data else pd.DataFrame()
    detalhado_df = pd.DataFrame(detalhado) if detalhado else pd.DataFrame()

    return resumo_df, detalhado_df

# ==================== FUNÇÕES AUXILIARES ====================

def convert_df_to_excel(resumo_df, detalhado_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if not resumo_df.empty:
            resumo_df.to_excel(writer, sheet_name='Resumo', index=False)
        if not detalhado_df.empty:
            detalhado_df.to_excel(writer, sheet_name='Detalhado', index=False)
    output.seek(0)
    return output.getvalue()

# ==================== INTERFACE STREAMLIT ====================

_AC_SVG_B64 = "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSI2NCIgaGVpZ2h0PSI2NCIgdmlld0JveD0iMCAwIDY0IDY0Ij4KICA8ZGVmcz4KICAgIDxsaW5lYXJHcmFkaWVudCBpZD0iZyIgeDE9IjAiIHkxPSIwIiB4Mj0iMSIgeTI9IjEiPgogICAgICA8c3RvcCBvZmZzZXQ9IjAiIHN0b3AtY29sb3I9IiMyMmM1NWUiLz4KICAgICAgPHN0b3Agb2Zmc2V0PSIwLjU1IiBzdG9wLWNvbG9yPSIjM2I4MmY2Ii8+CiAgICAgIDxzdG9wIG9mZnNldD0iMSIgc3RvcC1jb2xvcj0iIzI1NjNlYiIvPgogICAgPC9saW5lYXJHcmFkaWVudD4KICAgIDxmaWx0ZXIgaWQ9InMiIHg9Ii0yMCUiIHk9Ii0yMCUiIHdpZHRoPSIxNDAlIiBoZWlnaHQ9IjE0MCUiPgogICAgICA8ZmVEcm9wU2hhZG93IGR4PSIwIiBkeT0iNCIgc3RkRGV2aWF0aW9uPSI0IiBmbG9vZC1jb2xvcj0iIzAwMCIgZmxvb2Qtb3BhY2l0eT0iMC4zNSIvPgogICAgPC9maWx0ZXI+CiAgPC9kZWZzPgogIDxyZWN0IHg9IjYiIHk9IjYiIHdpZHRoPSI1MiIgaGVpZ2h0PSI1MiIgcng9IjE2IiBmaWxsPSJyZ2JhKDI1NSwyNTUsMjU1LDAuMDYpIiBzdHJva2U9InJnYmEoMjU1LDI1NSwyNTUsMC4xOCkiLz4KICA8Y2lyY2xlIGN4PSIzMiIgY3k9IjMyIiByPSIxNiIgZmlsbD0idXJsKCNnKSIgZmlsdGVyPSJ1cmwoI3MpIi8+CiAgPHRleHQgeD0iMzIiIHk9IjM4IiB0ZXh0LWFuY2hvcj0ibWlkZGxlIiBmb250LWZhbWlseT0iSW50ZXIsQXJpYWwiIGZvbnQtc2l6ZT0iMTYiIGZvbnQtd2VpZ2h0PSI4MDAiIGZpbGw9IiMwYjEwMjAiPkFDPC90ZXh0Pgo8L3N2Zz4="

def main():
    st.set_page_config(
        page_title="Analisador Contábil",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    st.markdown(r"""
    <style>
    :root{
        --bg: #070a12;
        --panel: #0b1020;
        --stroke: rgba(255,255,255,0.09);
        --stroke2: rgba(255,255,255,0.06);
        --text: #e5e7eb;
        --muted: rgba(229,231,235,0.68);
        --muted2: rgba(229,231,235,0.52);
        --blue: #3b82f6;
        --blue2: #2563eb;
    }

    html, body, [data-testid="stAppViewContainer"]{
        background: radial-gradient(1200px 600px at 25% 0%, rgba(59,130,246,0.14), transparent 60%),
                    radial-gradient(900px 500px at 70% 20%, rgba(34,197,94,0.10), transparent 55%),
                    var(--bg);
        color: var(--text);
        font-family: Inter, system-ui, -apple-system, Segoe UI, Roboto, Ubuntu, Cantarell, "Helvetica Neue", Arial;
    }

    [data-testid="stHeader"]{ background: transparent; }
    .main{ padding: 2.0rem 2.2rem 2rem 2.2rem; }

    [data-testid="stSidebar"]{
        background: linear-gradient(180deg, rgba(15,23,42,0.92), rgba(2,6,23,0.94));
        border-right: 1px solid var(--stroke2);
    }

    .custom-card{
        background: radial-gradient(140% 120% at 50% 0%, rgba(59,130,246,0.10), rgba(2,6,23,0.88));
        border: 1px solid var(--stroke);
        border-radius: 16px;
        padding: 1.6rem 1.7rem;
        box-shadow: 0 25px 50px rgba(0,0,0,0.35);
        margin-bottom: 1.3rem;
    }
    .subtle{ color: var(--muted); font-size: 14px; margin-top: 0.25rem; }

    .step-container{ display:flex; align-items:center; gap: 10px; margin: 1.1rem 0 1.2rem 0; flex-wrap: wrap; }
    .step-number{
        width: 36px; height: 36px; border-radius: 12px; display:flex; align-items:center; justify-content:center;
        font-weight: 750; font-size: 14px; background: rgba(255,255,255,0.05); border: 1px solid var(--stroke2);
        color: var(--muted);
    }
    .step-active{
        color: var(--text);
        background: rgba(59,130,246,0.16);
        border-color: rgba(59,130,246,0.35);
        box-shadow: 0 0 0 1px rgba(59,130,246,0.18);
    }
    .step-label{ font-weight: 520; color: var(--muted); font-size: 12.5px; margin-right: 14px; }

    .stButton>button{
        border-radius: 12px !important;
        height: 42px;
        font-weight: 650;
        border: 1px solid var(--stroke2);
        background: rgba(255,255,255,0.04);
        color: var(--text);
    }
    .stButton>button:hover{
        border-color: rgba(255,255,255,0.14);
        background: rgba(255,255,255,0.06);
    }
    .stButton>button[kind="primary"]{
        background: linear-gradient(90deg, var(--blue), var(--blue2)) !important;
        border: none !important;
        box-shadow: 0 16px 30px rgba(37,99,235,0.25);
    }

    [data-testid="stFileUploaderDropzone"]{
        background: rgba(255,255,255,0.03);
        border: 1px dashed rgba(255,255,255,0.14);
        border-radius: 16px;
    }

    [data-testid="stDataFrame"]{
        border-radius: 14px;
        border: 1px solid var(--stroke);
        overflow: hidden;
        background: rgba(2,6,23,0.75);
    }

    .stInfo, .stWarning, .stSuccess, .stError{
        border-radius: 14px !important;
        border: 1px solid var(--stroke2) !important;
        background: rgba(255,255,255,0.03) !important;
    }

    .tipo-badges{ display:flex; gap: 0.5rem; flex-wrap: wrap; }
    .badge{
        font-size: 12px; padding: 0.25rem 0.55rem; border-radius: 999px;
        border: 1px solid rgba(255,255,255,0.10); background: rgba(255,255,255,0.04); color: var(--muted);
    }
    .badge.primary{
        border-color: rgba(59,130,246,0.35);
        background: rgba(59,130,246,0.10);
        color: rgba(229,231,235,0.92);
    }

    .metric-line{ display:flex; gap: 1rem; margin: 1.2rem 0 0.2rem 0; align-items: stretch; flex-wrap: wrap; }
    .metric-item{
        background: rgba(255,255,255,0.03);
        border: 1px solid var(--stroke2);
        border-radius: 16px;
        padding: 0.95rem 1.05rem;
        min-width: 170px;
    }
    .metric-label{ font-size: 12px; color: var(--muted2); font-weight: 520; margin-bottom: 0.2rem; }
    .metric-value{ font-size: 26px; font-weight: 780; color: var(--text); letter-spacing: -0.02em; }

    /* Pills for st.radio */
    div[role="radiogroup"]{
        background: rgba(255,255,255,0.03);
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 999px;
        padding: 6px;
        display: inline-flex;
        gap: 6px;
    }
    div[role="radiogroup"] label{ margin: 0 !important; padding: 0 !important; }
    div[role="radiogroup"] label > div{
        border-radius: 999px !important;
        padding: 10px 14px !important;
        border: 1px solid transparent !important;
        background: transparent !important;
    }
    div[role="radiogroup"] input:checked + div{
        background: rgba(59,130,246,0.14) !important;
        border-color: rgba(59,130,246,0.35) !important;
        box-shadow: 0 0 0 1px rgba(59,130,246,0.14);
    }
    </style>
    """, unsafe_allow_html=True)

    if 'current_step' not in st.session_state:
        st.session_state.current_step = 1
    if 'df' not in st.session_state:
        st.session_state.df = None
    if 'resumo_df' not in st.session_state:
        st.session_state.resumo_df = None
    if 'detalhado_df' not in st.session_state:
        st.session_state.detalhado_df = None
    if 'tipo_analise' not in st.session_state:
        st.session_state.tipo_analise = None
    if 'arquivo_processado' not in st.session_state:
        st.session_state.arquivo_processado = False

    with st.sidebar:
        st.markdown(f"""
        <div style="display:flex; align-items:center; gap:10px; padding: 0.4rem 0.3rem 0.2rem 0.3rem;">
            <img src="data:image/svg+xml;base64,{_AC_SVG_B64}" style="width:46px; height:46px; border-radius:14px;" />
            <div>
                <div style="font-size:16px; font-weight:780; line-height:1.1;">Analisador Contábil</div>
                <div style="margin-top:3px; color: rgba(229,231,235,0.62); font-size: 12px;">Fluxo de análise</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<div class="step-container">', unsafe_allow_html=True)

        step1_active = st.session_state.current_step == 1
        st.markdown('<div class="step-number ' + ('step-active' if step1_active else '') + '">1</div>', unsafe_allow_html=True)
        st.markdown('<div class="step-label">Carregar Arquivo</div>', unsafe_allow_html=True)

        step2_active = st.session_state.current_step == 2
        st.markdown('<div class="step-number ' + ('step-active' if step2_active else '') + '">2</div>', unsafe_allow_html=True)
        st.markdown('<div class="step-label">Resumo</div>', unsafe_allow_html=True)

        step3_active = st.session_state.current_step == 3
        st.markdown('<div class="step-number ' + ('step-active' if step3_active else '') + '">3</div>', unsafe_allow_html=True)
        st.markdown('<div class="step-label">Detalhamento</div>', unsafe_allow_html=True)

        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("### 📁 Status")
        if st.session_state.arquivo_processado:
            st.success("✅ Processado")
            if st.session_state.tipo_analise:
                tipo_nome = "Fornecedores" if st.session_state.tipo_analise == "fornecedores" else "Contas Comuns"
                st.info(f"**Tipo:** {tipo_nome}")
            if st.session_state.resumo_df is not None and not st.session_state.resumo_df.empty:
                st.info(f"**Contas:** {len(st.session_state.resumo_df)}")
        else:
            st.info("⏳ Aguardando")

        st.markdown("---")
        if st.button("🗑️ Limpar Tudo", use_container_width=True):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

    if st.session_state.current_step == 1:
        etapa_carregar_arquivo()
    elif st.session_state.current_step == 2:
        etapa_escolher_tipo()
    elif st.session_state.current_step == 3:
        etapa_resultados()

def etapa_carregar_arquivo():
    st.markdown('<div class="custom-card">', unsafe_allow_html=True)
    st.markdown("## 📤 Carregar Arquivo Excel")
    st.markdown('<div class="subtle">Faça o upload do arquivo Excel com os dados contábeis para análise.</div>', unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "Selecione o arquivo",
        type=["xlsx"],
        help="Arraste e solte ou clique para selecionar",
        label_visibility="collapsed"
    )

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file, engine="openpyxl")

            colunas_necessarias = {"Histórico", "Chave", "Contra", "Valor", "Saldo"}
            if not colunas_necessarias.issubset(df.columns):
                st.error("❌ Colunas necessárias não encontradas.")
                st.info(f"Encontradas: {list(df.columns)}")
                st.markdown("</div>", unsafe_allow_html=True)
                return

            st.session_state.df = df
            st.success(f"✅ **Arquivo carregado:** {uploaded_file.name} ({uploaded_file.size/1024:.1f}KB)")

            st.markdown('<div class="metric-line">', unsafe_allow_html=True)
            col1, col2, col3 = st.columns(3)

            with col1:
                st.markdown(f'<div class="metric-item"><div class="metric-label">Linhas</div><div class="metric-value">{len(df):,}</div></div>', unsafe_allow_html=True)

            with col2:
                contas = df["Histórico"].apply(lambda x: 1 if eh_conta(str(x)) else 0).sum()
                st.markdown(f'<div class="metric-item"><div class="metric-label">Contas</div><div class="metric-value">{contas}</div></div>', unsafe_allow_html=True)

            with col3:
                datas = df[df["Histórico"].apply(lambda x: isinstance(x, datetime))]
                if not datas.empty:
                    periodo = f"{datas['Histórico'].min().date()} a {datas['Histórico'].max().date()}"
                    periodo_display = periodo.replace("2025-", "")
                else:
                    periodo_display = "Não identificado"
                st.markdown(f'<div class="metric-item"><div class="metric-label">Período</div><div class="metric-value">{periodo_display}</div></div>', unsafe_allow_html=True)

            st.markdown('</div>', unsafe_allow_html=True)

            if st.button("👁️ Visualizar primeiras linhas"):
                st.dataframe(df.head(10), use_container_width=True)

            st.markdown("---")
            st.markdown("#### 📝 Requisitos do arquivo")
            st.markdown("- Formato: **.xlsx** (Excel)")
            st.markdown("- Colunas obrigatórias: **Histórico, Chave, Contra, Valor, Saldo**")

            st.markdown("---")
            if st.button("➡️ Escolher Tipo de Análise", type="primary", use_container_width=True):
                st.session_state.current_step = 2
                st.rerun()

        except Exception as e:
            st.error(f"❌ Erro ao processar: {str(e)}")
    else:
        st.info("👆 **Aguardando upload do arquivo**")

    st.markdown("</div>", unsafe_allow_html=True)

def etapa_escolher_tipo():
    if st.session_state.df is None:
        st.warning("⚠️ Nenhum arquivo carregado.")
        if st.button("⬅️ Voltar"):
            st.session_state.current_step = 1
            st.rerun()
        return

    st.markdown('<div class="custom-card">', unsafe_allow_html=True)
    st.markdown("## 🧭 Tipo de Análise")
    st.markdown('<div class="subtle">Selecione a forma de interpretação do seu razão.</div>', unsafe_allow_html=True)

    options = ["Razão de Fornecedores", "Razão de Contas Comuns"]
    idx = 0 if st.session_state.tipo_analise != "contas_comuns" else 1
    choice = st.radio("", options, index=idx, horizontal=True, label_visibility="collapsed")
    st.session_state.tipo_analise = "fornecedores" if choice == options[0] else "contas_comuns"

    if st.session_state.tipo_analise == "fornecedores":
        st.markdown("""
        <div style="margin-top:12px; color: rgba(229,231,235,0.70); font-size: 13px;">
          <b>Para arquivos com participantes.</b> Calcula saldos por dia.
        </div>
        <div class="tipo-badges" style="margin-top:10px;">
          <span class="badge primary">Participantes</span>
          <span class="badge">Saldo por dia</span>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div style="margin-top:12px; color: rgba(229,231,235,0.70); font-size: 13px;">
          <b>Para contas comuns (ex: 277 - ICMS).</b> Considera saldo inicial e movimentações.
        </div>
        <div class="tipo-badges" style="margin-top:10px;">
          <span class="badge primary">Contábil/Fiscal</span>
          <span class="badge">Saldo inicial</span>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    col_prev, col_next = st.columns(2)

    with col_prev:
        if st.button("⬅️ Voltar", use_container_width=True):
            st.session_state.current_step = 1
            st.rerun()

    with col_next:
        if st.button("📊 Processar", type="primary", use_container_width=True):
            with st.spinner("Processando..."):
                if st.session_state.tipo_analise == "fornecedores":
                    resumo, detalhado = processar_fornecedores(st.session_state.df)
                else:
                    resumo, detalhado = processar_contas_comuns(st.session_state.df)

                st.session_state.resumo_df = resumo
                st.session_state.detalhado_df = detalhado
                st.session_state.arquivo_processado = True
                st.session_state.current_step = 3
                st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

def etapa_resultados():
    if not st.session_state.arquivo_processado:
        st.warning("⚠️ Nenhum arquivo processado.")
        if st.button("⬅️ Voltar"):
            st.session_state.current_step = 2
            st.rerun()
        return

    tipo_nome = "Fornecedores" if st.session_state.tipo_analise == "fornecedores" else "Contas Comuns"
    st.markdown(f"## 📊 Resultados — {tipo_nome}")
    st.markdown('<div class="subtle">Visão consolidada dos saldos negativos identificados.</div>', unsafe_allow_html=True)

    if st.button("⬅️ Voltar", use_container_width=False):
        st.session_state.current_step = 2
        st.rerun()

    if st.session_state.resumo_df is not None:
        if not st.session_state.resumo_df.empty:
            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.markdown("### 📈 Resumo Consolidado")

            if st.session_state.tipo_analise == "contas_comuns" and "Tem saldo inicial negativo" in st.session_state.resumo_df.columns:
                st.markdown('<div class="metric-line">', unsafe_allow_html=True)
                c1, c2, c3, c4 = st.columns(4)

                with c1:
                    st.markdown(f'<div class="metric-item"><div class="metric-label">Contas</div><div class="metric-value">{len(st.session_state.resumo_df)}</div></div>', unsafe_allow_html=True)
                with c2:
                    total_dias = st.session_state.resumo_df["Qtd dias com saldo negativo"].sum()
                    st.markdown(f'<div class="metric-item"><div class="metric-label">Dias com erro</div><div class="metric-value">{int(total_dias)}</div></div>', unsafe_allow_html=True)
                with c3:
                    primeira_data = st.session_state.resumo_df["Primeira data com erro"].min()
                    display_data = primeira_data if pd.notna(primeira_data) else "N/A"
                    st.markdown(f'<div class="metric-item"><div class="metric-label">Primeira ocorrência</div><div class="metric-value">{display_data}</div></div>', unsafe_allow_html=True)
                with c4:
                    saldos_iniciais = (st.session_state.resumo_df["Tem saldo inicial negativo"] == "✓").sum()
                    st.markdown(f'<div class="metric-item"><div class="metric-label">Saldos iniciais</div><div class="metric-value">{saldos_iniciais}</div></div>', unsafe_allow_html=True)

                st.markdown('</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="metric-line">', unsafe_allow_html=True)
                c1, c2, c3 = st.columns(3)

                with c1:
                    st.markdown(f'<div class="metric-item"><div class="metric-label">Contas</div><div class="metric-value">{len(st.session_state.resumo_df)}</div></div>', unsafe_allow_html=True)
                with c2:
                    total_dias = st.session_state.resumo_df["Qtd dias com saldo negativo"].sum()
                    st.markdown(f'<div class="metric-item"><div class="metric-label">Dias com erro</div><div class="metric-value">{int(total_dias)}</div></div>', unsafe_allow_html=True)
                with c3:
                    primeira_data = st.session_state.resumo_df["Primeira data com erro"].min()
                    display_data = primeira_data if pd.notna(primeira_data) else "N/A"
                    st.markdown(f'<div class="metric-item"><div class="metric-label">Primeira ocorrência</div><div class="metric-value">{display_data}</div></div>', unsafe_allow_html=True)

                st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("---")
            st.markdown("### 📋 Resultados Resumidos")
            st.dataframe(st.session_state.resumo_df, use_container_width=True, height=400)

            if st.session_state.detalhado_df is not None and not st.session_state.detalhado_df.empty:
                st.markdown("---")
                if st.button("🔍 Ver Detalhamento Completo", use_container_width=True):
                    st.session_state.show_detalhamento = not st.session_state.get('show_detalhamento', False)
                    st.rerun()

            st.markdown("</div>", unsafe_allow_html=True)

            if st.session_state.get('show_detalhamento', False) and st.session_state.detalhado_df is not None and not st.session_state.detalhado_df.empty:
                st.markdown('<div class="custom-card">', unsafe_allow_html=True)
                st.markdown("### 🔍 Detalhamento Completo")

                detalhado_display = st.session_state.detalhado_df.copy()
                colunas_monetarias = ["Saldo Anterior", "Débito do Dia", "Crédito do Dia", "Saldo Após o Dia"]

                for col in colunas_monetarias:
                    if col in detalhado_display.columns:
                        detalhado_display[col] = detalhado_display[col].apply(
                            lambda x: f"R$ {float(x):,.2f}" if isinstance(x, (int, float)) and not pd.isna(x) else "R$ 0,00"
                        )

                st.dataframe(detalhado_display, use_container_width=True, height=400)
                st.markdown("</div>", unsafe_allow_html=True)

            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.markdown("### 💾 Exportar Dados")

            excel_data = convert_df_to_excel(
                st.session_state.resumo_df,
                st.session_state.detalhado_df if st.session_state.detalhado_df is not None else pd.DataFrame()
            )

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="📥 Baixar Excel Completo",
                    data=excel_data,
                    file_name=f"relatorio_{st.session_state.tipo_analise}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            with col2:
                if st.button("🔄 Analisar Novo Arquivo", use_container_width=True):
                    tipo = st.session_state.tipo_analise
                    for key in list(st.session_state.keys()):
                        del st.session_state[key]
                    st.session_state.tipo_analise = tipo
                    st.session_state.current_step = 1
                    st.rerun()

            st.markdown("</div>", unsafe_allow_html=True)
        else:
            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.success("✅ **Excelente!** Nenhum saldo negativo encontrado.")
            st.balloons()

            st.markdown("---")
            if st.button("🔄 Analisar Novo Arquivo", use_container_width=True):
                tipo = st.session_state.tipo_analise
                for key in list(st.session_state.keys()):
                    del st.session_state[key]
                st.session_state.tipo_analise = tipo
                st.session_state.current_step = 1
                st.rerun()

            st.markdown("</div>", unsafe_allow_html=True)
    else:
        st.markdown('<div class="custom-card">', unsafe_allow_html=True)
        st.warning("⚠️ **Nenhum resultado disponível**")
        st.markdown("</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
