from __future__ import annotations

import base64
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

from core import analyze_balances, dataframe_to_excel, read_csv_semicolon


APP_DIR = Path(__file__).parent
LOGO_PATH = APP_DIR / "logo_analisador_contabil.svg"


def load_logo_data_uri() -> str:
    svg = LOGO_PATH.read_text(encoding="utf-8")
    encoded = base64.b64encode(svg.encode("utf-8")).decode("ascii")
    return f"data:image/svg+xml;base64,{encoded}"


def inject_css() -> None:
    st.markdown(
        """
        <style>
        :root {
            --bg: #0B1220;
            --panel: #131D32;
            --panel-soft: #17243B;
            --line: #24324A;
            --text: #F8FAFC;
            --muted: #94A3B8;
            --blue: #3B82F6;
            --cyan: #06B6D4;
            --green: #22C55E;
            --red: #EF4444;
            --purple: #8B5CF6;
            --yellow: #F6B73C;
        }

        .stApp {
            background:
              radial-gradient(900px 520px at 18% 0%, rgba(47, 140, 255, .13), transparent 60%),
              radial-gradient(850px 520px at 86% 6%, rgba(16, 232, 138, .10), transparent 58%),
              var(--bg);
            color: var(--text);
        }

        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #07101E 0%, #091221 100%);
            border-right: 1px solid var(--line);
        }

        [data-testid="stSidebar"] > div:first-child {
            padding: 28px 18px 22px;
        }

        .block-container {
            max-width: 1480px;
            padding: 36px 42px 48px;
        }

        h1, h2, h3, p { letter-spacing: 0; }

        .app-title {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 20px;
            margin-bottom: 26px;
        }

        .app-title h1 {
            margin: 0;
            font-size: 38px;
            line-height: 1.12;
            color: var(--text);
        }

        .app-title p {
            margin: 10px 0 0;
            color: var(--muted);
            font-size: 16px;
        }

        .primary-action {
            display: inline-flex;
            align-items: center;
            gap: 10px;
            padding: 14px 22px;
            background: linear-gradient(135deg, #0ABF6B, #10E88A);
            color: #03120B;
            border-radius: 8px;
            font-weight: 800;
            box-shadow: 0 16px 34px rgba(16, 232, 138, .18);
            white-space: nowrap;
        }

        .sidebar-logo {
            display: flex;
            gap: 12px;
            align-items: center;
            margin: 4px 0 28px;
        }

        .sidebar-logo img { width: 54px; height: 54px; }
        .sidebar-logo strong {
            display: block;
            font-size: 19px;
            line-height: 1.1;
            color: var(--text);
        }
        .sidebar-logo span {
            display: block;
            margin-top: 5px;
            font-size: 12px;
            color: var(--muted);
        }

        .side-item {
            display: flex;
            align-items: center;
            gap: 12px;
            padding: 14px 14px;
            margin: 7px 0;
            border-radius: 8px;
            color: #C8D2E5;
            font-weight: 650;
        }

        .side-item.active {
            background: rgba(47, 140, 255, .13);
            color: var(--text);
            box-shadow: inset 0 0 0 1px rgba(148, 163, 184, .13);
        }

        .side-file-card {
            margin-top: 34px;
            padding: 16px;
            border: 1px solid var(--line);
            border-radius: 8px;
            background: linear-gradient(180deg, rgba(17, 29, 48, .86), rgba(10, 18, 31, .92));
        }
        .side-file-card small { color: var(--muted); display: block; margin-bottom: 10px; }
        .side-file-card strong { color: var(--text); font-size: 14px; word-break: break-word; }
        .side-status {
            display: inline-flex;
            margin-top: 12px;
            padding: 7px 10px;
            border-radius: 7px;
            color: var(--green);
            background: rgba(16, 232, 138, .12);
            font-weight: 800;
            font-size: 12px;
        }

        .upload-panel, .section-panel {
            border: 1px solid var(--line);
            background: linear-gradient(180deg, rgba(17, 29, 48, .88), rgba(9, 17, 30, .94));
            border-radius: 8px;
            padding: 24px;
            box-shadow: 0 18px 54px rgba(0, 0, 0, .20);
        }

        .upload-title {
            font-size: 26px;
            font-weight: 850;
            color: var(--text);
            margin-bottom: 8px;
        }

        .upload-subtitle {
            color: var(--muted);
            margin-bottom: 20px;
        }

        .metric-card {
            border: 1px solid var(--line);
            border-radius: 8px;
            padding: 22px;
            min-height: 148px;
            background: linear-gradient(145deg, rgba(17, 29, 48, .96), rgba(11, 20, 34, .96));
            box-shadow: inset 0 1px 0 rgba(255,255,255,.03), 0 16px 46px rgba(0,0,0,.16);
        }

        .metric-card .label {
            color: #CBD5E1;
            font-size: 15px;
            margin-bottom: 14px;
        }
        .metric-card .value {
            font-size: 36px;
            font-weight: 900;
            line-height: 1;
        }
        .metric-card .hint {
            color: var(--muted);
            font-size: 13px;
            margin-top: 14px;
        }

        .section-head {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 16px;
            margin-bottom: 20px;
        }
        .section-head h2 {
            font-size: 26px;
            margin: 0 0 6px;
        }
        .section-head p {
            margin: 0;
            color: var(--muted);
        }

        .badge {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            padding: 6px 10px;
            border-radius: 7px;
            font-size: 12px;
            font-weight: 800;
            white-space: nowrap;
        }
        .badge-fornecedor { color: #E9D5FF; background: rgba(181, 108, 255, .18); }
        .badge-cliente { color: #BDFCEE; background: rgba(32, 217, 194, .16); }
        .badge-conta { color: #BFDBFE; background: rgba(47, 140, 255, .16); }
        .badge-credor { color: var(--green); background: rgba(16, 232, 138, .10); }
        .badge-devedor { color: var(--red); background: rgba(255, 92, 102, .10); }

        .table-header, .table-row {
            display: grid;
            grid-template-columns: 115px 130px minmax(250px, 1.8fr) 130px 120px 120px 130px 105px 86px;
            gap: 12px;
            align-items: center;
        }
        .table-header {
            padding: 16px 0;
            color: #E8EEF9;
            font-weight: 800;
            font-size: 13px;
            border-bottom: 1px solid var(--line);
        }
        .table-row {
            padding: 14px 0;
            border-bottom: 1px solid rgba(148, 163, 184, .12);
            color: #DDE6F5;
            font-size: 13px;
        }
        .table-row .desc {
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }

        .empty-state {
            border: 1px dashed rgba(148, 163, 184, .28);
            border-radius: 8px;
            padding: 36px;
            text-align: center;
            color: var(--muted);
        }

        .detail-panel {
            border: 1px solid var(--line);
            border-radius: 8px;
            padding: 16px;
            background: rgba(6, 11, 20, .66);
            color: #DDE6F5;
            margin-top: 12px;
        }

        div[data-testid="stFileUploader"] {
            border: 1px solid var(--line);
            border-radius: 8px;
            padding: 12px;
            background: rgba(6, 11, 20, .42);
        }

        .stButton > button, .stDownloadButton > button {
            border-radius: 8px !important;
            border: 1px solid rgba(148, 163, 184, .22) !important;
            background: rgba(17, 29, 48, .9) !important;
            color: #F5F7FB !important;
            font-weight: 800 !important;
            min-height: 42px;
        }
        .stButton > button:hover, .stDownloadButton > button:hover {
            border-color: rgba(47, 140, 255, .68) !important;
            color: #FFFFFF !important;
        }

        @media (max-width: 1100px) {
            .table-header { display: none; }
            .table-row {
                grid-template-columns: 1fr 1fr;
                gap: 10px 16px;
            }
            .app-title {
                align-items: flex-start;
                flex-direction: column;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def ensure_state() -> None:
    defaults: dict[str, Any] = {
        "analysis_done": False,
        "result": pd.DataFrame(),
        "issues": pd.DataFrame(),
        "file_name": "",
        "analysis_time": "",
        "selected_row": None,
        "search": "",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def render_status_badge(text: str, kind: str) -> str:
    cls = {
        "fornecedor": "badge-fornecedor",
        "cliente": "badge-cliente",
        "conta": "badge-conta",
        "credor": "badge-credor",
        "devedor": "badge-devedor",
    }.get(kind.lower(), "badge-conta")
    return f'<span class="badge {cls}">{text}</span>'


def render_sidebar() -> None:
    logo_uri = load_logo_data_uri()
    st.sidebar.markdown(
        f"""
        <div class="sidebar-logo">
            <img src="{logo_uri}" alt="Analisador Contabil">
            <div>
                <strong>Analisador<br>Contábil</strong>
                <span>Conferência automática</span>
            </div>
        </div>
        <div class="side-item active">▦ Resumo</div>
        <div class="side-item">△ Inconsistências</div>
        <div class="side-item">▤ Contas</div>
        <div class="side-item">♙ Participantes</div>
        <div class="side-item">⚙ Configuracoes</div>
        <div class="side-item">ⓘ Sobre</div>
        """,
        unsafe_allow_html=True,
    )

    file_name = st.session_state.get("file_name") or "Nenhum arquivo analisado"
    analysis_time = st.session_state.get("analysis_time") or "-"
    status = "Análise concluída" if st.session_state.get("analysis_done") else "Aguardando arquivos"
    st.sidebar.markdown(
        f"""
        <div class="side-file-card">
            <small>Arquivo analisado</small>
            <strong>{file_name}</strong>
            <small style="margin-top:10px">Atualizado em {analysis_time}</small>
            <span class="side-status">{status}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.sidebar.button("Nova análise", use_container_width=True):
        for key in ["analysis_done", "result", "issues", "file_name", "analysis_time", "selected_row", "search"]:
            st.session_state.pop(key, None)
        ensure_state()
        st.rerun()


def render_metric_card(label: str, value: str, hint: str, color: str) -> None:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="label">{label}</div>
            <div class="value" style="color:{color}">{value}</div>
            <div class="hint">{hint}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def br_money(value: Any) -> str:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return ""
    formatted = f"{abs(number):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"({formatted})" if number < 0 else formatted


def classify_row(row: pd.Series) -> str:
    text = " ".join(
        str(row.get(column, ""))
        for column in ["Codigo da conta", "Conta analisada", "Nome da conta no razao", "Nome no plano de contas", "Grupo"]
    ).lower()
    if "fornecedor" in text or str(row.get("Codigo da conta", "")) == "148":
        return "Fornecedor"
    if "cliente" in text:
        return "Cliente"
    return "Conta"


def participant_code(row: pd.Series) -> str:
    codigo = str(row.get("Codigo da conta", "")).strip()
    name = str(row.get("Nome da conta no razao", "")).strip()
    first = name.split(" ", 1)[0].strip()
    if first.isdigit():
        return f"{codigo} - {first}"
    return codigo


def description(row: pd.Series) -> str:
    name = str(row.get("Nome da conta no razao", "")).strip()
    pieces = name.split(" ", 1)
    if pieces and pieces[0].isdigit() and len(pieces) > 1:
        return pieces[1].lstrip("- ").strip()
    return name or str(row.get("Conta analisada", ""))


def expected_label(row: pd.Series) -> str:
    nature = str(row.get("Natureza esperada", "")).lower()
    return "Credor" if nature == "credora" else "Devedor" if nature == "devedora" else "Revisao"


def current_label(row: pd.Series) -> str:
    nature = str(row.get("Natureza esperada", "")).lower()
    return "Devedor" if nature == "credora" else "Credor" if nature == "devedora" else "Revisao"


def metric_summary(result: pd.DataFrame, issues: pd.DataFrame) -> dict[str, str]:
    if result.empty:
        return {
            "issues": "0",
            "accounts": "0",
            "participants": "0",
            "period": "-",
            "period_hint": "Nenhum arquivo analisado",
        }

    dates = pd.to_datetime(result["Data"], format="%d/%m/%Y", errors="coerce")
    min_date = dates.min()
    max_date = dates.max()
    period = "-"
    days = "0 dias analisados"
    if pd.notna(min_date) and pd.notna(max_date):
        period = f"{min_date.strftime('%d/%m/%Y')} a {max_date.strftime('%d/%m/%Y')}"
        days = f"{dates.dropna().nunique()} dias analisados"

    issue_types = issues.apply(classify_row, axis=1) if not issues.empty else pd.Series(dtype=str)
    participants = int(issue_types.isin(["Fornecedor", "Cliente"]).sum())

    return {
        "issues": str(len(issues)),
        "accounts": str(issues["Codigo da conta"].nunique() if not issues.empty else 0),
        "participants": str(participants),
        "period": period,
        "period_hint": days,
    }


def render_upload_area() -> None:
    st.markdown(
        """
        <div class="upload-panel">
            <div class="upload-title">Enviar arquivos para análise</div>
            <div class="upload-subtitle">Importe o plano de contas e o razão SCI em CSV separado por ponto e vírgula.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.container():
        col1, col2 = st.columns(2)
        with col1:
            plan_file = st.file_uploader("Plano de contas CSV", type=["csv"], key="plan_file")
        with col2:
            ledger_file = st.file_uploader("Razao diario CSV", type=["csv"], key="ledger_file")

        analyze = st.button("Analisar saldos diarios", type="primary", use_container_width=True)
        if analyze:
            if not plan_file or not ledger_file:
                st.warning("Envie o plano de contas e o razão diário para iniciar a análise.")
                return
            try:
                plan_df = read_csv_semicolon(plan_file)
                ledger_df = read_csv_semicolon(ledger_file)
                result, issues = analyze_balances(ledger_df, plan_df)
            except Exception as exc:
                st.error(f"Nao foi possivel analisar os arquivos: {exc}")
                return

            st.session_state.result = result
            st.session_state.issues = issues
            st.session_state.analysis_done = True
            st.session_state.file_name = ledger_file.name
            st.session_state.analysis_time = datetime.now().strftime("%d/%m/%Y %H:%M")
            st.session_state.selected_row = None
            st.rerun()


def render_table_row(row: pd.Series, index: int) -> None:
    row_type = classify_row(row)
    expected = expected_label(row)
    current = current_label(row)
    cols = st.columns([1.15, 1.3, 3.0, 1.25, 1.2, 1.15, 1.25, 1.0, .75])
    cols[0].markdown(render_status_badge(row_type, row_type.lower()), unsafe_allow_html=True)
    cols[1].markdown(f"<div class='table-row-cell'>{participant_code(row)}</div>", unsafe_allow_html=True)
    cols[2].markdown(f"<div class='table-row-cell desc'>{description(row)}</div>", unsafe_allow_html=True)
    cols[3].markdown(render_status_badge(expected, expected.lower()), unsafe_allow_html=True)
    cols[4].markdown(render_status_badge(current, current.lower()), unsafe_allow_html=True)
    cols[5].markdown(br_money(row.get("Saldo final do dia", 0)))
    cols[6].markdown(str(row.get("Data", "")))
    cols[7].markdown(str(row.get("Dias impactados", "") or "1"))
    if cols[8].button("Ver", key=f"details_{index}", use_container_width=True):
        selected = str(row.get("Conta analisada", ""))
        st.session_state.selected_row = selected if st.session_state.selected_row != selected else None


def render_inconsistencias_table(issues: pd.DataFrame) -> None:
    st.markdown(
        """
        <div class="section-panel">
          <div class="section-head">
            <div>
              <h2>Principais Inconsistências</h2>
              <p>Lista dos casos que precisam de verificação</p>
            </div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    search_col, filter_col, export_col = st.columns([3.4, 1.1, 1.2])
    with search_col:
        search = st.text_input("Buscar conta ou participante", key="search", placeholder="Buscar conta ou participante...")
    with filter_col:
        tipo = st.selectbox("Filtros", ["Todos", "Fornecedor", "Cliente", "Conta"], label_visibility="collapsed")
    with export_col:
        st.download_button(
            "Exportar Excel",
            data=dataframe_to_excel(issues),
            file_name="inconsistencias_analisador_contabil.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    filtered = issues.copy()
    if search:
        mask = filtered.apply(lambda row: search.lower() in " ".join(map(str, row.values)).lower(), axis=1)
        filtered = filtered[mask]
    if tipo != "Todos":
        filtered = filtered[filtered.apply(classify_row, axis=1).eq(tipo)]

    if filtered.empty:
        st.markdown('<div class="empty-state">Nenhum caso encontrado para os filtros atuais.</div>', unsafe_allow_html=True)
        return

    st.markdown(
        """
        <div class="table-header">
          <div>Tipo</div><div>Código</div><div>Descrição</div><div>Saldo Esperado</div>
          <div>Saldo Atual</div><div>Valor</div><div>1a Ocorrencia</div><div>Dias Afetados</div><div>Acoes</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    for index, row in filtered.reset_index(drop=True).iterrows():
        st.markdown('<div class="table-row">', unsafe_allow_html=True)
        render_table_row(row, index)
        st.markdown("</div>", unsafe_allow_html=True)
    st.caption(f"Mostrando {len(filtered)} de {len(issues)} resultados")


def render_detalhes() -> None:
    selected = st.session_state.get("selected_row")
    if not selected:
        return

    issues: pd.DataFrame = st.session_state.issues
    result: pd.DataFrame = st.session_state.result
    selected_rows = issues[issues["Conta analisada"].astype(str).eq(str(selected))]
    if selected_rows.empty:
        return

    row = selected_rows.iloc[0]
    history = result[result["Conta analisada"].astype(str).eq(str(selected))].copy()
    history["_data_dt"] = pd.to_datetime(history["Data"], format="%d/%m/%Y", errors="coerce")
    history = history.sort_values("_data_dt")

    st.markdown(
        f"""
        <div class="detail-panel">
            <strong>{row.get('Conta analisada', '')}</strong><br>
            <span>Tipo de inconsistência: {row.get('Tipo de inconsistencia', '')}</span><br>
            <span>Natureza esperada: {row.get('Natureza esperada', '')}</span><br>
            <span>Primeira ocorrência: {row.get('Data', '')}</span><br>
            <span>Última ocorrência: {row.get('Data final da sequencia', row.get('Data', ''))}</span><br>
            <span>Dias afetados: {row.get('Dias impactados', '1')}</span><br>
            <span>Saldo encontrado: {br_money(row.get('Saldo final do dia', 0))}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not history.empty:
        chart_df = history[["_data_dt", "Saldo final do dia"]].dropna().set_index("_data_dt").tail(90)
        st.line_chart(chart_df, height=220)
        st.dataframe(
            history[["Data", "Saldo final do dia"]].tail(40),
            use_container_width=True,
            hide_index=True,
        )


def render_dashboard() -> None:
    result: pd.DataFrame = st.session_state.result
    issues: pd.DataFrame = st.session_state.issues
    summary = metric_summary(result, issues)

    st.markdown(
        """
        <div class="app-title">
            <div>
                <h1>Resumo da Análise</h1>
                <p>Visão geral das inconsistências encontradas no razão contábil.</p>
            </div>
            <div class="primary-action">⇧ Enviar novo arquivo</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.button("Enviar novo arquivo", key="new_upload_top"):
        for key in ["analysis_done", "result", "issues", "file_name", "analysis_time", "selected_row", "search"]:
            st.session_state.pop(key, None)
        ensure_state()
        st.rerun()

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        render_metric_card("Total de Inconsistências", summary["issues"], "Casos que precisam de revisão", "var(--red)")
    with col2:
        render_metric_card("Contas Afetadas", summary["accounts"], "Contas do plano de contas", "var(--yellow)")
    with col3:
        render_metric_card("Participantes Afetados", summary["participants"], "Fornecedores e clientes", "var(--blue)")
    with col4:
        render_metric_card("Periodo Analisado", summary["period"], summary["period_hint"], "var(--green)")

    st.write("")
    render_inconsistencias_table(issues)
    render_detalhes()


def main() -> None:
    st.set_page_config(
        page_title="Analisador Contábil",
        page_icon="✓",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    ensure_state()
    inject_css()
    render_sidebar()

    if st.session_state.analysis_done:
        render_dashboard()
    else:
        st.markdown(
            """
            <div class="app-title">
                <div>
                    <h1>Analisador Contábil</h1>
                    <p>Conferência automática de saldos no razão SCI.</p>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        render_upload_area()


if __name__ == "__main__":
    main()
