import streamlit as st
import pandas as pd
import re
from datetime import datetime
import io

# [Todas as funções anteriores permanecem EXATAMENTE as mesmas]
# Funções auxiliares do motor original
def eh_conta(v):
    return isinstance(v, str) and re.match(r"^\d+\s*-\s*", v.strip())

def texto_tem_saldo_anterior(v):
    return isinstance(v, str) and ("saldo anterior" in v.lower())

def parse_num(v):
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

def processar_arquivo(df):
    """Motor de processamento original"""
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
            conta_atual = hist.strip()
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

def convert_df_to_excel(resumo_df, detalhado_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        resumo_df.to_excel(writer, sheet_name='Resumo', index=False)
        detalhado_df.to_excel(writer, sheet_name='Detalhado', index=False)
    output.seek(0)
    return output.getvalue()

# Função principal do app Streamlit
def main():
    st.set_page_config(
        page_title="Analisador Contábil",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # Estilos CSS
    st.markdown("""
    <style>
    .main {
        padding: 1rem;
    }

    [data-testid="stSidebar"] {
        background: #1e3a5f;
        padding-top: 2rem;
    }

    .step-container {
        display: flex;
        flex-direction: column;
        align-items: center;
        margin: 2rem 0;
        gap: 10px;
    }

    .step-indicator {
        width: 40px;
        height: 40px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: bold;
        font-size: 16px;
        background: #4a5568;
        color: white;
        border: 3px solid #718096;
    }

    .step-active {
        background: #22c55e;
        border-color: #16a34a;
    }

    .step-label {
        font-weight: 500;
        color: #e2e8f0;
        font-size: 14px;
        text-align: center;
    }

    .custom-card {
        background: white;
        border-radius: 10px;
        padding: 1.5rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        border: 1px solid #e2e8f0;
        margin-bottom: 1rem;
    }

    /* Estilos para as caixas de info */
    .stInfo {
        border-left: 4px solid #22c55e !important;
        background-color: #f0f9ff !important;
    }

    .stWarning {
        border-left: 4px solid #f59e0b !important;
        background-color: #fef3c7 !important;
    }

    /* Nomes das colunas em negrito */
    .column-name {
        font-weight: 600;
        color: #2d3748;
    }
    </style>
    """, unsafe_allow_html=True)

    # Inicializar estado da sessão
    if 'current_step' not in st.session_state:
        st.session_state.current_step = 1
    if 'df' not in st.session_state:
        st.session_state.df = None
    if 'resumo_df' not in st.session_state:
        st.session_state.resumo_df = None
    if 'detalhado_df' not in st.session_state:
        st.session_state.detalhado_df = None
    if 'arquivo_processado' not in st.session_state:
        st.session_state.arquivo_processado = False

    # Sidebar com indicador de etapas
    with st.sidebar:
        st.markdown("""
        <div style="text-align: center;">
            <h2 style="color: white; margin-bottom: 30px;">📊 Analisador Contábil</h2>
        </div>
        """, unsafe_allow_html=True)

        # Indicador de etapas
        st.markdown('<div class="step-container">', unsafe_allow_html=True)

        # Etapa 1 - Carregar Arquivo
        step1_active = st.session_state.current_step == 1
        st.markdown(f'''
        <div class="step-indicator {'step-active' if step1_active else ''}">
            1
        </div>
        <div class="step-label">
            Carregar Arquivo
        </div>
        ''', unsafe_allow_html=True)

        # Etapa 2 - Resumo
        step2_active = st.session_state.current_step == 2
        st.markdown(f'''
        <div class="step-indicator {'step-active' if step2_active else ''}">
            2
        </div>
        <div class="step-label">
            Resumo
        </div>
        ''', unsafe_allow_html=True)

        # Etapa 3 - Detalhamento
        step3_active = st.session_state.current_step == 3
        st.markdown(f'''
        <div class="step-indicator {'step-active' if step3_active else ''}">
            3
        </div>
        <div class="step-label">
            Detalhamento
        </div>
        ''', unsafe_allow_html=True)

        st.markdown('</div>', unsafe_allow_html=True)

        # Informações do status
        if st.session_state.arquivo_processado:
            st.markdown("---")
            st.markdown("### 📁 Status")
            st.success("✅ Arquivo processado")

            if st.session_state.resumo_df is not None and not st.session_state.resumo_df.empty:
                total_contas = len(st.session_state.resumo_df)
                st.info(f"**Contas com erro:** {total_contas}")

    # Conteúdo principal baseado na etapa atual
    if st.session_state.current_step == 1:
        carregar_arquivo()
    elif st.session_state.current_step == 2:
        exibir_resumo()
    elif st.session_state.current_step == 3:
        exibir_detalhamento()

def carregar_arquivo():
    st.markdown('<div class="custom-card">', unsafe_allow_html=True)
    st.markdown("## 📤 Carregar Arquivo Excel")
    st.markdown("Faça o upload do arquivo Excel com os dados contábeis para análise.")
    st.markdown("</div>", unsafe_allow_html=True)

    # Container principal
    st.markdown('<div class="custom-card">', unsafe_allow_html=True)

    # Layout em duas colunas
    col1, col2 = st.columns([3, 2])

    with col1:
        st.markdown("### Selecione o arquivo")
        uploaded_file = st.file_uploader(
            "Escolha um arquivo .xlsx",
            type=["xlsx"],
            help="Arraste e solte ou clique para selecionar",
            label_visibility="collapsed"
        )

        if uploaded_file is not None:
            try:
                # Mostrar nome do arquivo
                st.markdown(f"**Arquivo selecionado:** `{uploaded_file.name}`")

                # Ler o arquivo
                df = pd.read_excel(uploaded_file, engine="openpyxl")

                # Verificar colunas necessárias
                cols_necessarias = {"Histórico", "Chave", "Contra", "Valor", "Saldo"}
                if not cols_necessarias.issubset(df.columns):
                    st.error(f"**Erro:** Colunas necessárias não encontradas.")
                    st.error(f"Colunas encontradas: {list(df.columns)}")
                    st.error(f"Colunas necessárias: {list(cols_necessarias)}")
                    return

                # Salvar no estado da sessão
                st.session_state.df = df

                # Mostrar prévia dos dados
                st.markdown("---")
                st.markdown("### 📋 Pré-visualização (5 primeiras linhas)")
                st.dataframe(df.head(), use_container_width=True)

                # Botões de ação
                st.markdown("---")
                st.markdown("### ⚙️ Processar arquivo")

                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    if st.button("📊 Ir para Resumo", use_container_width=True):
                        with st.spinner("Processando arquivo..."):
                            # Processar arquivo
                            resumo_df, detalhado_df = processar_arquivo(df)
                            st.session_state.resumo_df = resumo_df
                            st.session_state.detalhado_df = detalhado_df
                            st.session_state.arquivo_processado = True
                            st.session_state.current_step = 2
                        st.rerun()

                with col_btn2:
                    if st.button("🔍 Ir para Detalhamento", use_container_width=True):
                        with st.spinner("Processando arquivo..."):
                            # Processar arquivo
                            resumo_df, detalhado_df = processar_arquivo(df)
                            st.session_state.resumo_df = resumo_df
                            st.session_state.detalhado_df = detalhado_df
                            st.session_state.arquivo_processado = True
                            st.session_state.current_step = 3
                        st.rerun()

            except Exception as e:
                st.error(f"**Erro ao processar o arquivo:** {str(e)}")
                st.info("Verifique se o arquivo está no formato correto (.xlsx) e não está corrompido.")
        else:
            st.info("👆 Aguardando upload do arquivo...")

    with col2:
        st.markdown("### ℹ️ Instruções")

        # Caixa azul de informações - USANDO st.info()
        with st.container():
            st.markdown("#### Requisitos do arquivo")

            # Conteúdo simples e limpo
            info_content = """
            **Formato:** .xlsx (Excel)

            **Colunas obrigatórias:**

            • **Histórico**
            • **Chave**
            • **Contra**
            • **Valor**
            • **Saldo**

            **⚠️ Importante:** O arquivo deve seguir o layout padrão do sistema contábil.
            """

            # Usamos st.info() que já tem fundo azul
            st.info(info_content, icon="📋")

        # Espaço
        st.markdown("")

        # Caixa amarela de dica - USANDO st.warning()
        with st.container():
            st.markdown("#### 💡 Dica")
            st.warning(
                "Após fazer o upload, clique em um dos botões abaixo para processar o arquivo e visualizar os resultados."
            )

    st.markdown("</div>", unsafe_allow_html=True)

# [As funções exibir_resumo() e exibir_detalhamento() permanecem EXATAMENTE as mesmas da versão anterior]
def exibir_resumo():
    st.markdown('<div class="custom-card">', unsafe_allow_html=True)

    col_title, col_nav = st.columns([3, 1])
    with col_title:
        st.markdown("## 📊 Resumo dos Dados")
        st.markdown("Visão consolidada dos saldos negativos identificados")

    with col_nav:
        if st.button("⬅️ Voltar", use_container_width=True):
            st.session_state.current_step = 1
            st.rerun()
        if st.button("➡️ Detalhamento", use_container_width=True):
            st.session_state.current_step = 3
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

    if st.session_state.resumo_df is not None:
        if not st.session_state.resumo_df.empty:
            st.markdown('<div class="custom-card">', unsafe_allow_html=True)

            # Métricas rápidas
            st.markdown("### 📈 Métricas")
            col1, col2, col3 = st.columns(3)

            with col1:
                total_contas = len(st.session_state.resumo_df)
                st.metric("Total de Contas", f"{total_contas}")

            with col2:
                total_dias = st.session_state.resumo_df["Qtd dias com saldo negativo"].sum()
                st.metric("Dias com Erros", f"{total_dias}")

            with col3:
                primeira_data = st.session_state.resumo_df["Primeira data com erro"].min()
                display_data = primeira_data if pd.notna(primeira_data) else "Nenhuma"
                st.metric("Primeira Ocorrência", display_data)

            st.markdown("---")

            # Tabela de resumo
            st.markdown("### 📋 Resultados Detalhados")
            st.dataframe(
                st.session_state.resumo_df,
                use_container_width=True,
                height=400
            )

            # Botões de ação
            st.markdown("---")
            st.markdown("### 💾 Exportar Dados")
            col_export, col_space = st.columns([1, 3])
            with col_export:
                excel_data = convert_df_to_excel(
                    st.session_state.resumo_df,
                    st.session_state.detalhado_df if st.session_state.detalhado_df is not None else pd.DataFrame()
                )
                st.download_button(
                    label="📥 Baixar Excel",
                    data=excel_data,
                    file_name="relatorio_contabil.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            st.markdown("</div>", unsafe_allow_html=True)
        else:
            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.success("✅ **Excelente!** Nenhum saldo negativo foi identificado nos dados processados.")
            st.markdown("""
            <div style="background: #f0f9ff; padding: 1rem; border-radius: 8px; margin-top: 1rem;">
                <p style="color: #1e3a5f; margin: 0;">
                Todos os saldos estão positivos ou zerados. Não há inconsistências a serem reportadas.
                </p>
            </div>
            """, unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)
    else:
        st.markdown('<div class="custom-card">', unsafe_allow_html=True)
        st.warning("⚠️ **Nenhum arquivo processado**")
        st.markdown("""
        <div style="background: #fef3c7; padding: 1rem; border-radius: 8px; margin-top: 1rem;">
            <p style="color: #92400e; margin: 0;">
            Você precisa carregar e processar um arquivo primeiro na etapa de **"Carregar Arquivo"**.
            </p>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

def exibir_detalhamento():
    st.markdown('<div class="custom-card">', unsafe_allow_html=True)

    col_title, col_nav = st.columns([3, 1])
    with col_title:
        st.markdown("## 🔍 Detalhamento dos Movimentos")
        st.markdown("Movimentações detalhadas com saldos negativos")

    with col_nav:
        if st.button("⬅️ Voltar ao Resumo", use_container_width=True):
            st.session_state.current_step = 2
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

    if st.session_state.detalhado_df is not None:
        if not st.session_state.detalhado_df.empty:
            st.markdown('<div class="custom-card">', unsafe_allow_html=True)

            # Estatísticas
            total_registros = len(st.session_state.detalhado_df)
            st.markdown(f"**Total de registros encontrados:** `{total_registros}`")

            # Formatar valores monetários
            detalhado_display = st.session_state.detalhado_df.copy()
            colunas_monetarias = ["Saldo Anterior", "Débito do Dia", "Crédito do Dia", "Saldo Após o Dia"]

            for col in colunas_monetarias:
                if col in detalhado_display.columns:
                    detalhado_display[col] = detalhado_display[col].apply(
                        lambda x: f"R$ {float(x):,.2f}" if isinstance(x, (int, float)) and not pd.isna(x) else "R$ 0,00"
                    )

            st.markdown("---")
            st.markdown("### 📋 Detalhes das Movimentações")

            # Tabela detalhada
            st.dataframe(
                detalhado_display,
                use_container_width=True,
                height=500
            )

            # Exportação
            st.markdown("---")
            st.markdown("### 💾 Exportar Dados")
            col_export, col_space = st.columns([1, 3])
            with col_export:
                excel_data = convert_df_to_excel(
                    st.session_state.resumo_df,
                    st.session_state.detalhado_df
                )
                st.download_button(
                    label="📥 Baixar Excel",
                    data=excel_data,
                    file_name="detalhamento_contabil.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            st.markdown("</div>", unsafe_allow_html=True)
        else:
            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.success("✅ **Excelente!** Não há movimentações com saldos negativos.")
            st.markdown("""
            <div style="background: #f0f9ff; padding: 1rem; border-radius: 8px; margin-top: 1rem;">
                <p style="color: #1e3a5f; margin: 0;">
                Todas as movimentações resultaram em saldos positivos. Não há inconsistências a serem reportadas.
                </p>
            </div>
            """, unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)
    else:
        st.markdown('<div class="custom-card">', unsafe_allow_html=True)
        st.warning("⚠️ **Nenhum arquivo processado**")
        st.markdown("""
        <div style="background: #fef3c7; padding: 1rem; border-radius: 8px; margin-top: 1rem;">
            <p style="color: #92400e; margin: 0;">
            Você precisa carregar e processar um arquivo primeiro na etapa de **"Carregar Arquivo"**.
            </p>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
