# Streamlit app — Análise Full (VBA → Python)
# ---------------------------------------------------------------
# • Converte a macro VBA enviada em um app visual, com upload de planilhas,
#   KPIs, tabelas com cores, simulação de reposição e exportação para Excel.
# • Mantém as mesmas regras de negócio da macro (filtros, sugestões, custos, alertas).
# • Suporta consolidar várias empresas na mesma sessão (guarda no session_state).
# ---------------------------------------------------------------

import io
import math
import time
import zipfile
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

# --------------------------
# Config e estilo
# --------------------------
st.set_page_config(
    page_title="Análise Full — Dashboard",
    page_icon="📦",
    layout="wide",
)

# Pequeno CSS para cabeçalhos e chips
st.markdown(
    """
    <style>
    .kpi-card {
        border-radius: 14px; padding: 14px 16px; box-shadow: 0 2px 12px rgba(0,0,0,0.06);
        background: var(--background-color);
        border: 1px solid rgba(0,0,0,0.06);
    }
    .tag { display:inline-block; padding: 2px 8px; border-radius: 999px; font-size: 12px; font-weight: 600; }
    .tag.red { background:#ffe5e5; color:#7a0613; }
    .tag.yellow { background:#fff7db; color:#8a6a00; }
    .tag.green { background:#e9f7ef; color:#1e6b3a; }
    .tag.gray { background:#efefef; color:#444; }
    .section-title { font-weight: 700; margin-top: 0.6rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

# --------------------------
# Regras de negócio (idênticas à macro)
# --------------------------
EMPRESAS = ["VALE RACE", "VANPARTS", "MOTOILBR", "LUB EXPRESS"]

# Mapas de peso (usados no consolidado)
ACAO_PESO = {
    "Repor imediatamente": 6,
    "Corrigir anúncio e repor": 5,
    "Campanha de giro agressiva": 4,
    "Campanha de giro / reduzir estoque": 3,
    "Avaliar retirada / sem reposição": 2,
    "Evitar reposição / promoção": 1,
    "Evitar reposição / criar promoção": 1,
    "Sem ação definida": 0,
}

ALERTA_PESO = {
    "Alerta Vermelho": 3,
    "Avaliar giro": 2,
    "Sem urgência": 1,
    "Sem custo": 0,
}

# --------------------------
# Utilitários
# --------------------------

def to_int(x):
    try:
        if pd.isna(x):
            return 0
        if isinstance(x, str):
            x = x.strip().replace(".", "").replace(",", ".")
            if x == "":
                return 0
            return int(float(x))
        return int(float(x))
    except Exception:
        return 0


def to_float(x):
    try:
        if pd.isna(x):
            return 0.0
        if isinstance(x, str):
            x = x.strip().replace(".", "").replace(",", ".")
            if x == "":
                return 0.0
            return float(x)
        return float(x)
    except Exception:
        return 0.0


def normalize_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def color_alert(val):
    base = str(val).strip()
    if base == "Alerta Vermelho":
        return "background-color: #FFC7CE; color: #7a0613;"
    if base == "Avaliar giro":
        return "background-color: #FFEB9C; color: #5a4b00;"
    if base == "Sem urgência":
        return "background-color: #C6EFCE; color: #1e6b3a;"
    if base == "Sem custo":
        return "background-color: #F2F2F2; color: #333;"
    return ""


def human_int(n):
    return f"{int(n):,}".replace(",", ".")


# --------------------------
# Parsing do Relatório FULL (aba "Resumo") — colunas equivalentes à macro
# --------------------------
# Colunas na macro (por letra):
# D SKU | E #Anúncio | F Produto | I Status | K Vendas30d | L Afeta métrica | M Entrada pendente |
# P Aptas | Q Não aptas | U Estoque Full | W Boa Qualidade (Qtd monitorar) | X Impulsionar |
# Y Corrigir | Z Descarte | AA Tempo até esgotar

FULL_MAP = {
    "SKU": "D",
    "# Anúncio": "E",
    "Produto": "F",
    "Status": "I",
    "Vendas últimos 30 dias": "K",
    "Afeta métrica estoque": "L",
    "Entrada pendente": "M",
    "Aptas venda": "P",
    "Não aptas": "Q",
    "Estoque Full": "U",
    "Boa Qualidade": "W",  # qtd monitorar
    "Qtd. Impulsionar": "X",
    "Qtd. Corrigir": "Y",
    "Qtd. Risco Descarte": "Z",
    "Tempo até esgotar": "AA",
}


def excel_letter_to_index(letter: str) -> int:
    # 1-based
    letter = letter.upper().strip()
    result = 0
    for ch in letter:
        result = result * 26 + (ord(ch) - 64)
    return result


def read_full_resumo(xls, start_row=12):
    """Lê a planilha 'Resumo' do arquivo FULL e retorna DataFrame com colunas já mapeadas."""
    df = pd.read_excel(xls, sheet_name="Resumo", header=None)
    out = {}
    for name, col in FULL_MAP.items():
        idx = excel_letter_to_index(col) - 1
        out[name] = df.iloc[:, idx]
    tmp = pd.DataFrame(out)
    # manter a partir da linha 13 da planilha (índice 12 zero-based)
    tmp = tmp.iloc[start_row:, :].copy()
    # limpeza básica
    for c in [
        "Vendas últimos 30 dias",
        "Aptas venda",
        "Não aptas",
        "Estoque Full",
        "Boa Qualidade",
        "Qtd. Impulsionar",
        "Qtd. Corrigir",
        "Qtd. Risco Descarte",
    ]:
        tmp[c] = tmp[c].map(to_int)
    tmp["Tempo até esgotar"] = tmp["Tempo até esgotar"].map(normalize_str)
    tmp["Status"] = tmp["Status"].map(lambda x: normalize_str(x).lower())
    tmp["SKU"] = tmp["SKU"].map(normalize_str)
    tmp["Produto"] = tmp["Produto"].map(normalize_str)
    return tmp


# --------------------------
# Regras de filtro e sugestão (iguais à macro)
# --------------------------

def filtrar_e_sugerir(df):
    rows = []
    for _, r in df.iterrows():
        status = (r.get("Status", "") or "").strip().lower()
        estoque_full = int(r.get("Estoque Full", 0))
        # filtra: ativo OU (n/a E estoque_full>0)
        if status == "ativo" or (status == "n/a" and estoque_full > 0):
            vendas = int(r.get("Vendas últimos 30 dias", 0))
            qtd_mon = int(r.get("Boa Qualidade", 0))
            qtd_imp = int(r.get("Qtd. Impulsionar", 0))
            qtd_cor = int(r.get("Qtd. Corrigir", 0))
            qtd_desc = int(r.get("Qtd. Risco Descarte", 0))

            # Lógica de sugestão (igual à macro)
            if vendas == 0 and qtd_desc > 0:
                sugestao = "Avaliar retirada / sem reposição"
            elif estoque_full < 5 and vendas >= 10:
                sugestao = "Repor imediatamente"
            elif qtd_imp > 100:
                sugestao = "Campanha de giro agressiva"
            elif qtd_imp > 0 and vendas >= 3:
                sugestao = "Campanha de giro / reduzir estoque"
            elif qtd_cor > 0 and vendas > 5:
                sugestao = "Corrigir anúncio e repor"
            elif vendas < 5 and estoque_full > 10:
                sugestao = "Evitar reposição / criar promoção"
            else:
                sugestao = "Sem ação definida"

            rows.append({
                "SKU": r.get("SKU"),
                "# Anúncio": r.get("# Anúncio"),
                "Produto": r.get("Produto"),
                "Vendas últimos 30 dias": vendas,
                "Afeta métrica estoque": r.get("Afeta métrica estoque"),
                "Entrada pendente": r.get("Entrada pendente"),
                "Unid. aptas p/ venda": int(r.get("Aptas venda", 0)),
                "Não aptas": int(r.get("Não aptas", 0)),
                "Estoque Full": estoque_full,
                "Boa Qualidade": qtd_mon,
                "Qtd. Impulsionar": qtd_imp,
                "Qtd. Corrigir": qtd_cor,
                "Qtd. Risco Descarte": qtd_desc,
                "Tempo até esgotar": r.get("Tempo até esgotar"),
                "Comentário estoque": sugestao,
            })
    return pd.DataFrame(rows)


# --------------------------
# Custos (sheet: "Custos por estoque antigo")
# C: SKU, F: unidades (estoque antigo), I: dias estocado, K: custo total, L: aptas
# --------------------------
CUSTOS_MAP = {
    "SKU": "C",
    "Estoque com custo antigo": "F",
    "Dias estocado (média) [sum]": "I",
    "Custo total": "K",
    "Unid. aptas p/ venda (custo)": "L",
}


def read_custos(xls, sheet="Custos por estoque antigo", start_row=2):
    df = pd.read_excel(xls, sheet_name=sheet, header=None)
    out = {}
    for name, col in CUSTOS_MAP.items():
        idx = excel_letter_to_index(col) - 1
        out[name] = df.iloc[:, idx]
    tmp = pd.DataFrame(out)
    tmp = tmp.iloc[start_row:, :].copy()
    tmp["SKU"] = tmp["SKU"].map(normalize_str)
    for c in ["Estoque com custo antigo", "Unid. aptas p/ venda (custo)"]:
        tmp[c] = tmp[c].map(to_float)
    for c in ["Dias estocado (média) [sum]", "Custo total"]:
        tmp[c] = tmp[c].map(to_float)
    # agrega por SKU (soma e média de dias)
    agg = tmp.groupby("SKU").agg({
        "Estoque com custo antigo": "sum",
        "Dias estocado (média) [sum]": ["sum", "count"],
        "Custo total": "sum",
        "Unid. aptas p/ venda (custo)": "sum",
    })
    agg.columns = ["Estoque antigo", "dias_sum", "dias_n", "Custo total", "Aptas custo"]
    # média de dias = dias_sum / dias_n
    agg["Dias estocado (média)"] = (agg["dias_sum"] / agg["dias_n"]).round(0)
    return agg.reset_index()[["SKU", "Estoque antigo", "Dias estocado (média)", "Custo total", "Aptas custo"]]


def aplicar_custos(df_emp, df_custos):
    out = df_emp.merge(df_custos, on="SKU", how="left")
    out["Custo total"] = out["Custo total"].fillna(0).round(2)
    out["Estoque antigo"] = out["Estoque antigo"].fillna(0).round(0)
    out["Dias estocado (média)"] = out["Dias estocado (média)"].fillna(0).round(0)
    out["Aptas custo"] = out["Aptas custo"].fillna(0).round(0)

    # Alerta de custo — mesmo critério da macro
    def alerta(v):
        v = float(v or 0)
        if v > 150:
            return "Alerta Vermelho"
        elif v >= 101:
            return "Avaliar giro"
        elif v == 0:
            return "Sem custo"
        else:
            return "Sem urgência"

    out["Alerta de custo"] = out["Custo total"].map(alerta)
    return out


# --------------------------
# Consolidado (junta várias empresas carregadas)
# --------------------------

def consolidar_empresas(objs: dict):
    # objs: {empresa: df_empresarial_com_custos}
    # Prepara estrutura por SKU
    buckets = {}
    for emp, df in objs.items():
        for _, r in df.iterrows():
            sku = r["SKU"]
            if not sku:
                continue
            b = buckets.get(sku, {
                "SKU": sku,
                "Produto": r.get("Produto", ""),
                "Vendas VALE RACE": 0,
                "Estoque VALE RACE": 0,
                "Vendas VANPARTS": 0,
                "Estoque VANPARTS": 0,
                "Vendas MOTOILBR": 0,
                "Estoque MOTOILBR": 0,
                "Vendas LUB EXPRESS": 0,
                "Estoque LUB EXPRESS": 0,
                "Total Vendas 30d": 0,
                "Total Estoque": 0,
                "Custo Total": 0.0,
                "Empresas Envolvidas": "",
                "Ação Recomendada": "Sem ação definida",
                "Alerta de Custo": "Sem custo",
            })
            vendas = int(r.get("Vendas últimos 30 dias", 0))
            estoque = int(r.get("Estoque Full", 0))
            custo = float(r.get("Custo total", 0.0))
            acao = r.get("Comentário estoque", "Sem ação definida")
            alerta = r.get("Alerta de custo", "Sem custo")

            if emp == "VALE RACE":
                b["Vendas VALE RACE"] += vendas
                b["Estoque VALE RACE"] += estoque
            elif emp == "VANPARTS":
                b["Vendas VANPARTS"] += vendas
                b["Estoque VANPARTS"] += estoque
            elif emp == "MOTOILBR":
                b["Vendas MOTOILBR"] += vendas
                b["Estoque MOTOILBR"] += estoque
            elif emp == "LUB EXPRESS":
                b["Vendas LUB EXPRESS"] += vendas
                b["Estoque LUB EXPRESS"] += estoque

            b["Total Vendas 30d"] = (
                b["Vendas VALE RACE"] + b["Vendas VANPARTS"] + b["Vendas MOTOILBR"] + b["Vendas LUB EXPRESS"]
            )
            b["Total Estoque"] = (
                b["Estoque VALE RACE"] + b["Estoque VANPARTS"] + b["Estoque MOTOILBR"] + b["Estoque LUB EXPRESS"]
            )
            b["Custo Total"] += custo

            if emp not in b["Empresas Envolvidas"]:
                b["Empresas Envolvidas"] = (b["Empresas Envolvidas"] + ", " + emp).strip(", ")

            # Maior prioridade
            if ACAO_PESO.get(acao, -1) > ACAO_PESO.get(b["Ação Recomendada"], -1):
                b["Ação Recomendada"] = acao
            if ALERTA_PESO.get(alerta, -1) > ALERTA_PESO.get(b["Alerta de Custo"], -1):
                b["Alerta de Custo"] = alerta

            buckets[sku] = b

    if not buckets:
        return pd.DataFrame()
    out = pd.DataFrame(list(buckets.values()))

    # Margem % (placeholder  — na macro: dados(16) = (Total Vendas * 1) / Custo_total)
    out["Margem %"] = np.where(out["Custo Total"] > 0, (out["Total Vendas 30d"] * 1.0) / out["Custo Total"], 0.0)
    return out


# --------------------------
# Reposição (DBM) — mesma lógica da macro
# --------------------------

def simular_reposicao(df_consol):
    if df_consol.empty:
        return df_consol
    df = df_consol.copy()
    media_diaria = df["Total Vendas 30d"].astype(float) / 30.0

    def classificar(md):
        if md > 1:
            return ("Alta", 1.3, 2)
        elif md >= 0.3:
            return ("Média", 1.2, 1)
        return ("Baixa", 1.1, 0)

    cats = media_diaria.map(classificar)
    df["Categoria"] = cats.map(lambda x: x[0])
    df["Fator"] = cats.map(lambda x: x[1])
    df["Extra"] = cats.map(lambda x: x[2])
    df["Qtd. Sugerida"] = (media_diaria * 15 * df["Fator"] + df["Extra"]).round(0).astype(int)

    def crit(estoque, sug):
        if estoque == 0:
            return "Ruptura total"
        if estoque < sug * 0.5:
            return "Reposição urgente"
        if estoque < sug:
            return "Reposição recomendada"
        return "OK"

    df["Criticidade"] = [crit(e, s) for e, s in zip(df["Total Estoque"].astype(int), df["Qtd. Sugerida"].astype(int))]
    df["Cálculo Usado"] = [
        f"Média {md:.2f} × 15 × {f} + {ex} = {int(s)}" for md, f, ex, s in zip(media_diaria, df["Fator"], df["Extra"], df["Qtd. Sugerida"])
    ]
    return df[[
        "SKU", "Produto", "Total Vendas 30d", "Total Estoque", "Qtd. Sugerida", "Criticidade", "Categoria", "Cálculo Usado"
    ]]


# --------------------------
# Export helpers
# --------------------------

def to_excel_bytes(dfs: dict):
    # dfs: {sheet_name: df}
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        for name, df in dfs.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
    bio.seek(0)
    return bio.getvalue()


# --------------------------
# UI — Sidebar
# --------------------------
with st.sidebar:
    st.header("⚙️ Entrada de Dados")
    empresa = st.selectbox("Empresa", EMPRESAS, index=0)

    st.caption("Relatório de estoque FULL (aba 'Resumo'):")
    full_file = st.file_uploader("Relatório FULL (.xlsx)", type=["xlsx"], key="full")

    st.caption("Planilha de custos (aba 'Custos por estoque antigo'):")
    custos_file = st.file_uploader("Planilha de Custos (.xlsx)", type=["xlsx"], key="custos")

    run = st.button("▶️ Processar")

    if "empresas_data" not in st.session_state:
        st.session_state.empresas_data = {}

st.title("📦 Análise Full — Dashboard")

# --------------------------
# Execução
# --------------------------
if run:
    if not full_file:
        st.error("Envie o Relatório FULL.")
    else:
        with st.spinner("Lendo e preparando dados…"):
            # 1) FULL
            df_full = read_full_resumo(full_file)
            df_emp = filtrar_e_sugerir(df_full)

            # 2) Custos (opcional)
            if custos_file is not None:
                df_custos = read_custos(custos_file)
                df_emp = aplicar_custos(df_emp, df_custos)
            else:
                # Preencher colunas de custo vazias p/ manter o layout
                for c in ["Estoque antigo", "Dias estocado (média)", "Custo total", "Aptas custo", "Alerta de custo"]:
                    if c not in df_emp.columns:
                        df_emp[c] = 0 if c != "Alerta de custo" else "Sem custo"

            # Reordena e renomeia colunas para espelhar a macro
            cols_order = [
                "SKU", "# Anúncio", "Produto", "Vendas últimos 30 dias", "Afeta métrica estoque",
                "Entrada pendente", "Unid. aptas p/ venda", "Não aptas", "Estoque Full", "Boa Qualidade",
                "Qtd. Impulsionar", "Qtd. Corrigir", "Qtd. Risco Descarte", "Tempo até esgotar",
                "Comentário estoque", "Estoque antigo", "Dias estocado (média)", "Custo total",
                "Aptas custo", "Alerta de custo"
            ]
            df_emp = df_emp.reindex(columns=cols_order)

            # Salva na sessão
            st.session_state.empresas_data[empresa] = df_emp.copy()

        st.success(f"Processado para {empresa} — {len(df_emp):,} SKUs")

# --------------------------
# Abas principais
# --------------------------
aba = st.tabs(["📋 Empresa", "📊 Painel Consolidado", "🚚 Reposição Full", "⬇️ Exportar"])

# --- 1) Empresa
with aba[0]:
    st.subheader("Visão da Empresa (dados atuais)")
    df_emp = st.session_state.empresas_data.get(empresa)
    if df_emp is None or df_emp.empty:
        st.info("Envie e processe os arquivos para esta empresa na barra lateral.")
    else:
        # KPIs
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            st.markdown('<div class="kpi-card"><div class="section-title">SKUs</div>'
                        f'<h3>{human_int(len(df_emp))}</h3></div>', unsafe_allow_html=True)
        with c2:
            st.markdown('<div class="kpi-card"><div class="section-title">Vendas 30d</div>'
                        f'<h3>{human_int(df_emp["Vendas últimos 30 dias"].sum())}</h3></div>', unsafe_allow_html=True)
        with c3:
            st.markdown('<div class="kpi-card"><div class="section-title">Estoque Full</div>'
                        f'<h3>{human_int(df_emp["Estoque Full"].sum())}</h3></div>', unsafe_allow_html=True)
        with c4:
            total_custo = float(df_emp["Custo total"].fillna(0).sum())
            st.markdown('<div class="kpi-card"><div class="section-title">Custo Total</div>'
                        f'<h3>R$ {total_custo:,.2f}</h3></div>'.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)
        with c5:
            alertas = df_emp["Alerta de custo"].value_counts().to_dict()
            a_red = alertas.get("Alerta Vermelho", 0)
            a_yel = alertas.get("Avaliar giro", 0)
            a_grn = alertas.get("Sem urgência", 0)
            a_gray = alertas.get("Sem custo", 0)
            st.markdown(
                '<div class="kpi-card"><div class="section-title">Alertas</div>'
                f'<div class="tag red">Vermelho: {a_red}</div> '
                f'<div class="tag yellow">Avaliar: {a_yel}</div> '
                f'<div class="tag green">OK: {a_grn}</div> '
                f'<div class="tag gray">Sem custo: {a_gray}</div>'
                '</div>',
                unsafe_allow_html=True,
            )

        st.divider()

        # Filtros básicos
        colf1, colf2, colf3 = st.columns(3)
        with colf1:
            filtro_alerta = st.multiselect(
                "Filtrar por Alerta de custo",
                options=["Alerta Vermelho", "Avaliar giro", "Sem urgência", "Sem custo"],
                default=["Alerta Vermelho", "Avaliar giro", "Sem urgência", "Sem custo"],
            )
        with colf2:
            filtro_acao = st.multiselect(
                "Filtrar por Comentário estoque",
                options=list(ACAO_PESO.keys()),
                default=list(ACAO_PESO.keys()),
            )
        with colf3:
            termo = st.text_input("Busca por SKU/Produto")

        df_view = df_emp.copy()
        if filtro_alerta:
            df_view = df_view[df_view["Alerta de custo"].isin(filtro_alerta)]
        if filtro_acao:
            df_view = df_view[df_view["Comentário estoque"].isin(filtro_acao)]
        if termo:
            termo_l = termo.lower().strip()
            df_view = df_view[
                df_view["SKU"].str.lower().str.contains(termo_l, na=False)
                | df_view["Produto"].str.lower().str.contains(termo_l, na=False)
            ]

        # Estilo para a coluna "Alerta de custo"
        def styler(sdf: pd.DataFrame):
            sty = sdf.style.applymap(color_alert, subset=["Alerta de custo"])  # cores
            return sty

        st.dataframe(styler(df_view), use_container_width=True, hide_index=True)

# --- 2) Consolidado
with aba[1]:
    st.subheader("Painel Consolidado (todas empresas nesta sessão)")
    dados = st.session_state.empresas_data.copy()
    if not dados:
        st.info("Carregue pelo menos uma empresa na aba anterior.")
    else:
        df_con = consolidar_empresas(dados)
        if df_con.empty:
            st.info("Sem dados consolidados.")
        else:
            # KPIs
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.markdown('<div class="kpi-card"><div class="section-title">SKUs</div>'
                            f'<h3>{human_int(len(df_con))}</h3></div>', unsafe_allow_html=True)
            with c2:
                st.markdown('<div class="kpi-card"><div class="section-title">Vendas 30d (Total)</div>'
                            f'<h3>{human_int(df_con["Total Vendas 30d"].sum())}</h3></div>', unsafe_allow_html=True)
            with c3:
                st.markdown('<div class="kpi-card"><div class="section-title">Estoque (Total)</div>'
                            f'<h3>{human_int(df_con["Total Estoque"].sum())}</h3></div>', unsafe_allow_html=True)
            with c4:
                st.markdown('<div class="kpi-card"><div class="section-title">Custo Total</div>'
                            f'<h3>R$ {df_con["Custo Total"].sum():,.2f}</h3></div>'.replace(",", "X").replace(".", ",").replace("X", "."), unsafe_allow_html=True)

            st.divider()

            # Tabela com cor de Alerta
            st.dataframe(df_con.style.applymap(color_alert, subset=["Alerta de Custo"]), use_container_width=True, hide_index=True)

# --- 3) Reposição
with aba[2]:
    st.subheader("Simulação de Reposição (DBM)")
    dados = st.session_state.empresas_data.copy()
    if not dados:
        st.info("Carregue pelo menos uma empresa na aba 'Empresa'.")
    else:
        df_con = consolidar_empresas(dados)
        if df_con.empty:
            st.info("Sem dados consolidados.")
        else:
            df_rep = simular_reposicao(df_con)
            # Ordenar por criticidade
            ord_map = {"Ruptura total": 0, "Reposição urgente": 1, "Reposição recomendada": 2, "OK": 3}
            df_rep["_ord"] = df_rep["Criticidade"].map(ord_map)
            df_rep = df_rep.sort_values(["_ord", "Qtd. Sugerida"], ascending=[True, False]).drop(columns=["_ord"])            
            st.dataframe(df_rep, use_container_width=True, hide_index=True)

# --- 4) Export
with aba[3]:
    st.subheader("Exportar Excel")
    if not st.session_state.empresas_data:
        st.info("Não há dados para exportar.")
    else:
        # Monta pacotes por empresa e um consolidado
        dfs_xlsx = {}
        for emp, dfe in st.session_state.empresas_data.items():
            dfs_xlsx[f"{emp}"] = dfe
        # Consolidado geral
        df_con = consolidar_empresas(st.session_state.empresas_data)
        if not df_con.empty:
            dfs_xlsx["Painel Consolidado"] = df_con
            df_rep = simular_reposicao(df_con)
            if not df_rep.empty:
                dfs_xlsx["Reposição Full"] = df_rep

        blob = to_excel_bytes(dfs_xlsx)
        st.download_button(
            label="💾 Baixar Excel (todas as abas)",
            data=blob,
            file_name=f"AnaliseFull_{datetime.now().strftime('%Y-%m-%d_%Hh%Mm')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

st.caption("Feito com ❤️ em Streamlit • Regras espelhadas da macro VBA • Suporta várias empresas na mesma sessão")
