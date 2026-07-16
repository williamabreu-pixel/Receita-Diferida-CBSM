# app.py
"""Dashboard Streamlit. Rodar com:  streamlit run app.py"""

import io
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
import streamlit as st

import config
import data_loader
import data_pipeline
import gerador_lancamentos

# theme=None é usado em todos os st.plotly_chart (ver mais abaixo) para que o
# modebar nativo do Plotly (zoom / reset axes / autoscale) funcione — o tema
# "streamlit" embutido interfere nesses botões. Sem esse tema, o Plotly usa
# seu template padrão (claro); fixamos "plotly_dark" para casar com o fundo
# escuro do app.
pio.templates.default = "plotly_dark"

st.set_page_config(page_title="Passivo & Receita Diferida — IFRS 15",
                   page_icon="📊", layout="wide")

# --- Identidade visual Dotz -------------------------------------------------
# Coloque os arquivos na MESMA pasta do app.py (data_raw/). Se não existirem,
# o app roda normalmente sem o logo (fallback silencioso).
CAMINHO_LOGO = "logo_dotz.png"          # logo horizontal (cabeçalho + canto)
CAMINHO_ICONE_LOGO = "icone_dotz.png"   # opcional: versão compacta/quadrada


def _aplicar_logo_canto():
    """Marca Dotz persistente no canto superior esquerdo (st.logo)."""
    if not Path(CAMINHO_LOGO).exists():
        return
    try:
        if Path(CAMINHO_ICONE_LOGO).exists():
            st.logo(CAMINHO_LOGO, size="large", icon_image=CAMINHO_ICONE_LOGO)
        else:
            st.logo(CAMINHO_LOGO, size="large")
    except Exception:
        pass  # Streamlit antigo sem st.logo — o cabeçalho abaixo já cobre


def _render_cabecalho():
    """Cabeçalho com o logo Dotz ao lado do título."""
    if Path(CAMINHO_LOGO).exists():
        col_logo, col_titulo = st.columns([1, 6])
        with col_logo:
            st.image(CAMINHO_LOGO, width=150)
        with col_titulo:
            st.title("Passivo & Receita Diferida — IFRS 15 / CPC 47")
            st.caption("Conciliação de pontos: emissão, resgate e breakage vs balancete SAP.")
    else:
        st.title("📊 Passivo & Receita Diferida — IFRS 15 / CPC 47")
        st.caption("Conciliação de pontos: emissão, resgate e breakage vs balancete SAP.")


CORES_MOTOR = {"Breakage": "#FF4F0D", "Trocas": "#FEC114",
               "Spread": "#009F3C", "Promodotz": "#000000",
               "Receita Projetos Especiais": "#D82598"}

# Paleta Dotz para as categorias de faturamento (aba Faturamento de Pontos) —
# as 7 categorias da fonte batem exatamente com as 7 cores oficiais da marca.
CORES_FATURAMENTO = {
    "NSF": "#FF4F0D",
    "SF": "#FEC114",
    "BV": "#009F3C",
    "Viaja Dotz": "#D82598",
    "Comissão com Conversão de Dotz": "#AF1010",
    "MKT": "#000000",
    "Prestação de Serviços Comissão -  Banco do Brasil": "#E3D2C8",
}

# Rampa sequencial (laranja Dotz, #FF4F0D) para dimensões temporais por
# safra/vintage (2022-2026) — paleta separada da de motor, já que aqui a
# ordem cronológica importa mais do que distinguir categorias de negócio.
PALETA_SAFRA = {"2022": "#FFD9C7", "2023": "#FFAD8A", "2024": "#FF7A47",
                "2025": "#FF4F0D", "2026": "#C93D0A"}

TIPO_DIFERIMENTO = {
    "Breakage": "Diferimento 48 meses (linear por competência de emissão)",
    "Spread": "Diferimento 48 meses (linear por competência de emissão)",
    "Trocas": "Fato gerador (resgate)",
    "Promodotz": "Fato gerador (evento)",
    "Receita Projetos Especiais": "Fato gerador (resgate)",
}

MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def fmt_brl(valor):
    if valor is None or pd.isna(valor):
        return "—"
    s = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def to_excel_bytes(df, sheet_name="Dados"):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    return buffer.getvalue()


@st.cache_data(show_spinner="Processando pipeline contábil...")
def carregar_resultado(competencia):
    return data_pipeline.executar_pipeline(competencia=competencia)


@st.cache_data(show_spinner="Carregando classificação de receita (DRE 2026)...")
def carregar_dre_categoria():
    return data_loader.load_dre_classificacao_receita()


@st.cache_data(show_spinner="Carregando aba...")
def carregar_aba_bruta(nome_aba):
    return data_loader.load_aba_bruta(nome_aba)


@st.cache_data(show_spinner="Carregando grupos da U1 - Movimentação...")
def carregar_u1_grupos():
    return data_loader.load_controle_u1_grupos()


@st.cache_data(show_spinner="Carregando série atuarial...")
def carregar_serie_atuarial(nome_aba, col_ancora):
    return data_loader.load_serie_atuarial(nome_aba, col_ancora)


@st.cache_data(show_spinner="Carregando U1.6 - Emissão e Resgates...")
def carregar_u1_6():
    return data_loader.load_u1_6_resgates()


@st.cache_data(show_spinner="Carregando Dados_Gráfico...")
def carregar_dados_grafico():
    return data_loader.load_dados_grafico_resumo()


@st.cache_data(show_spinner="Calculando split CP/LP (fonte: U1 - Movimentação)...")
def carregar_passivo_cp_lp():
    return data_loader.calcular_passivo_cp_lp()


@st.cache_data(show_spinner="Carregando expectativa de realização...")
def carregar_expectativa_realizacao():
    return data_loader.load_dados_grafico_expectativa_realizacao()


@st.cache_data(show_spinner="Carregando faturamento de vendas...")
def carregar_faturamento_vendas():
    return data_loader.load_faturamento_vendas()


EPS_CHECK_LINHA = 0.01


def fmt_numero_pt(valor, decimais=0):
    if valor is None or pd.isna(valor):
        return "—"
    s = f"{valor:,.{decimais}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


MESES_ABREV_PT = {"01": "Jan", "02": "Fev", "03": "Mar", "04": "Abr", "05": "Mai", "06": "Jun",
                  "07": "Jul", "08": "Ago", "09": "Set", "10": "Out", "11": "Nov", "12": "Dez"}


def fmt_competencia_pt(competencia):
    if not isinstance(competencia, str) or "-" not in competencia:
        return competencia
    ano, mes = competencia.split("-")
    return f"{MESES_ABREV_PT.get(mes, mes)}/{ano}"


def competencia_para_pt(serie_competencia):
    """Converte 'YYYY-MM' para rótulo em português ('Jan/2026'), como
    Categorical ordenado cronologicamente — evita que o Plotly interprete a
    coluna como data (mostrando meses em inglês) ou a reordene
    alfabeticamente depois do rótulo. Padrão para todo o app."""
    ordem_original = sorted(serie_competencia.dropna().unique())
    rotulos_ordenados = [fmt_competencia_pt(c) for c in ordem_original]
    mapa = dict(zip(ordem_original, rotulos_ordenados))
    return pd.Categorical(serie_competencia.map(mapa), categories=rotulos_ordenados, ordered=True)


def fmt_percentual(valor, decimais=2):
    if valor is None or pd.isna(valor):
        return "—"
    return f"{fmt_numero_pt(valor * 100, decimais)}%"


def fmt_valor_coluna(nome_col, valor):
    if not isinstance(valor, (int, float)) or pd.isna(valor):
        return "—" if (valor is None or pd.isna(valor)) else valor
    nome = nome_col.lower()
    if nome in ("ano", "mês", "mes"):
        return str(int(valor))
    if "preço" in nome:
        return f"R$ {fmt_numero_pt(valor, 6)}"
    if nome.startswith("(sem rótulo"):
        return fmt_percentual(valor) if abs(valor) <= 1 else fmt_numero_pt(valor, 2)
    if "pontos" in nome:
        return fmt_numero_pt(valor, 0)
    return fmt_brl(valor)


def render_cartao_titulo(linhas):
    if not linhas:
        return
    empresa = linhas[0] if len(linhas) > 0 else ""
    titulo = linhas[1] if len(linhas) > 1 else ""
    nota = linhas[2] if len(linhas) > 2 else ""
    parametros = linhas[3:]
    badges = "".join(
        f'<span style="background:rgba(255,255,255,.12);border-radius:6px;'
        f'padding:3px 10px;margin-right:8px;font-size:.78rem;">'
        f'{parametros[i]}: <b>{parametros[i + 1]}</b></span>'
        for i in range(0, len(parametros) - 1, 2)
    )
    st.markdown(f"""
<div style="background:linear-gradient(135deg,#0f172a,#1e293b);color:white;
            padding:16px 22px;border-radius:12px;margin-bottom:14px;">
  <div style="font-size:.7rem;opacity:.6;letter-spacing:.06em;text-transform:uppercase;">{empresa}</div>
  <div style="font-size:1.15rem;font-weight:700;margin-top:2px;">{titulo}</div>
  <div style="font-size:.78rem;opacity:.7;margin-top:4px;">{nota}</div>
  {f'<div style="margin-top:8px;">{badges}</div>' if badges else ''}
</div>
""", unsafe_allow_html=True)


def render_alerta(res):
    if res.conciliacao_ok:
        st.success("🟢 Conciliação 100% OK — todos os blocos dentro da tolerância.")
    else:
        st.error("🔴 Conciliação com divergência(s) — verifique os blocos abaixo.")
    for alerta in res.alertas:
        st.warning(alerta)


def render_cards(res):
    delta = res.passivo_total_controle - res.passivo_total_contabil
    c1, c2, c3 = st.columns(3)
    c1.metric("Passivo Contábil (Balancete)", fmt_brl(res.passivo_total_contabil))
    c2.metric("Passivo Controle (U1)", fmt_brl(res.passivo_total_controle))
    c3.metric("Delta (Controle − Contábil)", fmt_brl(delta),
              delta=fmt_brl(delta),
              delta_color="normal" if abs(delta) < 1 else "inverse")


RENOMEIA_QUADRO = {"A": "Total Passivo", "G": "Passivo Recalculado"}


def render_conciliacao(res):
    st.subheader("Conciliação")
    df = res.conciliacao.copy()
    df["quadro"] = df["quadro"].map(lambda q: RENOMEIA_QUADRO.get(q, q))
    for col in ("controle", "contabil", "delta"):
        df[col] = df[col].map(fmt_brl)

    def cor(v):
        return "color:#009F3C;font-weight:600;" if v == "OK" else "color:#AF1010;font-weight:600;"
    st.dataframe(df.style.map(cor, subset=["status"]),
                 width="stretch", hide_index=True)


def _formatar_hover_barras(fig):
    """Corrige o hover das barras empilhadas (px.bar) para usar a formatação
    BRL (ponto de milhar, sem decimais) em vez do padrão do Plotly — padrão
    usado em todos os gráficos empilhados do app."""
    for trace in fig.data:
        if trace.type == "bar":
            trace.customdata = [fmt_brl(v) for v in trace.y]
            trace.hovertemplate = f"%{{x}}<br>{trace.name}: %{{customdata}}<extra></extra>"


def _adicionar_linha_total(fig, df, col_x, col_y, nome="Total"):
    """Sobrepõe uma linha com o total (soma de todas as categorias/séries)
    por col_x, em qualquer gráfico de barras empilhadas — padrão usado em
    todos os gráficos empilhados do app. Os rótulos ficam deslocados acima
    da linha (trace de texto à parte) para não sobrepor os marcadores, e o
    hover usa a mesma formatação BRL dos rótulos (não o padrão do Plotly)."""
    total = df.groupby(col_x, as_index=False, observed=True)[col_y].sum().sort_values(col_x)
    hover = [fmt_brl(v) for v in total[col_y]]
    fig.add_trace(go.Scatter(x=total[col_x], y=total[col_y], mode="lines+markers", name=nome,
                             line=dict(color="#FFFFFF", width=3, dash="dot"),
                             marker=dict(size=7, color="#FFFFFF"),
                             customdata=hover,
                             hovertemplate=f"%{{x}}<br>{nome}: %{{customdata}}<extra></extra>"))
    deslocamento = total[col_y].max() * 0.08
    fig.add_trace(go.Scatter(x=total[col_x], y=total[col_y] + deslocamento, mode="text",
                             text=[f"<b>R$ {fmt_numero_pt(v, 0)}</b>" for v in total[col_y]],
                             textfont=dict(color="#FFFFFF", size=11),
                             showlegend=False, hoverinfo="skip"))


def render_receita_barras():
    st.subheader("Receita por Categoria")
    st.caption("Fonte: aba DRE 2026 (coluna Classificação Receita) — receita reconhecida no resultado.")
    detalhe = carregar_dre_categoria()
    if detalhe.empty:
        st.info("Sem dados de receita.")
        return
    agg = detalhe.groupby(["competencia", "categoria"], as_index=False)["valor"].sum()
    agg["competencia"] = competencia_para_pt(agg["competencia"])
    fig = px.bar(agg, x="competencia", y="valor", color="categoria",
                 color_discrete_map=CORES_MOTOR, barmode="stack",
                 labels={"valor": "Valor (R$)", "competencia": "Competência", "categoria": "Categoria"})
    _formatar_hover_barras(fig)
    _adicionar_linha_total(fig, agg, "competencia", "valor", "Receita Bruta")
    fig.update_layout(height=430, legend=dict(orientation="h", y=1.12), margin=dict(t=40, b=60))
    st.plotly_chart(fig, width="stretch", theme=None, key="chart_receita_barras_visao")


def render_passivo_rosca(res):
    st.subheader("Composição do Passivo Diferido")
    blocos = res.blocos_controle
    if blocos.empty:
        st.info("Sem blocos de controle.")
        return
    fig = go.Figure(go.Pie(labels=blocos["bloco"], values=blocos["passivo"], hole=0.55,
                           textinfo="label+percent", pull=[0, 0, 0, 0.08],
                           customdata=[fmt_brl(v) for v in blocos["passivo"]],
                           hovertemplate="%{label}<br>%{customdata} (%{percent})<extra></extra>",
                           marker=dict(colors=["#FF4F0D", "#FEC114", "#009F3C", "#000000"])))
    fig.update_layout(height=430, margin=dict(t=30, b=20),
                      annotations=[dict(text="Passivo Diferido", x=0.5, y=0.5, showarrow=False, font_size=14)])
    st.plotly_chart(fig, width="stretch", theme=None, key="chart_rosca_visao")


def render_dre(res):
    st.subheader("DRE — Receita Reconhecida por Competência")
    dre = res.resumo_dre
    if dre.empty:
        st.info("Sem linhas de DRE.")
        return
    col1, col2 = st.columns([2, 1])
    with col1:
        dre_chart = dre.copy()
        dre_chart["competencia"] = competencia_para_pt(dre_chart["competencia"])
        fig = px.area(dre_chart, x="competencia", y="receita_reconhecida",
                      labels={"receita_reconhecida": "Receita (R$)", "competencia": "Competência"})
        fig.update_traces(line_color="#FF4F0D", fillcolor="rgba(255,79,13,0.15)")
        fig.update_layout(height=360, margin=dict(t=30, b=60))
        st.plotly_chart(fig, width="stretch", theme=None, key="chart_dre")
    with col2:
        tab = dre.copy()
        tab["receita_reconhecida"] = tab["receita_reconhecida"].map(fmt_brl)
        st.dataframe(tab, width="stretch", hide_index=True)
        st.metric("Receita Total Reconhecida", fmt_brl(dre["receita_reconhecida"].sum()))


def render_lancamentos(res):
    st.subheader("Lançamentos de Ajuste de Conciliação")
    diario = gerador_lancamentos.gerar_diario(res)
    c1, c2 = st.columns(2)
    with c1:
        if diario.partidas_ok:
            st.success("🟢 Partidas dobradas OK (débitos = créditos)")
        else:
            st.error("🔴 Partidas dobradas NÃO fecham")
    with c2:
        if diario.cobertura_ok:
            st.success("🟢 Radar de cobertura: diário cobre 100% dos desvios")
        else:
            st.warning("⚠️ Cobertura incompleta — há desvio não lançado")
    for alerta in diario.alertas:
        st.warning(alerta)
    df = diario.lancamentos.copy()
    for col in ("debito", "credito", "J_desvio"):
        df[col] = df[col].map(fmt_brl)
    st.dataframe(df, width="stretch", hide_index=True)
    st.caption(f"Totais — Débito: {fmt_brl(diario.total_debito)}  |  Crédito: {fmt_brl(diario.total_credito)}")
    st.download_button("📥 Exportar diário (Excel)", data=to_excel_bytes(diario.lancamentos, "Lancamentos"),
                       file_name="lancamentos_ajuste.xlsx", mime=MIME_XLSX)


def render_aba_visao_geral(res):
    render_alerta(res)
    st.divider()
    render_cards(res)
    st.divider()
    c_esq, c_dir = st.columns(2)
    with c_esq:
        render_receita_barras()
    with c_dir:
        render_passivo_rosca(res)
    st.divider()
    render_conciliacao(res)
    st.divider()
    render_aba_dados_grafico()


def render_aba_receita_motor():
    st.subheader("Receita Diferida por Categoria")
    st.caption("Fonte: aba DRE 2026 (coluna Classificação Receita) — receita reconhecida no "
               "resultado. \"Trocas\" aqui é um valor diferente do \"Custo do Produto\" usado na "
               "conciliação (Visão Geral/Controle U1), pois vem de uma base contábil distinta.")
    df = carregar_dre_categoria()
    if df.empty:
        st.info("Sem dados de receita.")
        return

    competencias = sorted(df["competencia"].dropna().unique())
    categorias = sorted(df["categoria"].dropna().unique())
    c1, c2 = st.columns(2)
    with c1:
        sel_comp = st.multiselect("Competência", competencias, default=competencias, key="rec_comp")
    with c2:
        sel_cat = st.multiselect("Categoria", categorias, default=categorias, key="rec_motor")

    filtrado = df[df["competencia"].isin(sel_comp) & df["categoria"].isin(sel_cat)]
    if filtrado.empty:
        st.warning("Nenhuma linha para os filtros selecionados.")
        return

    agg = (filtrado.groupby(["competencia", "categoria"], as_index=False)
                   .agg(valor=("valor", "sum"), linhas=("valor", "size"))
                   .sort_values(["competencia", "categoria"]))

    agg_grafico = agg.copy()
    agg_grafico["competencia"] = competencia_para_pt(agg_grafico["competencia"])
    fig = px.bar(agg_grafico, x="competencia", y="valor", color="categoria",
                 color_discrete_map=CORES_MOTOR, barmode="stack",
                 labels={"valor": "Valor (R$)", "competencia": "Competência", "categoria": "Categoria"})
    _formatar_hover_barras(fig)
    _adicionar_linha_total(fig, agg_grafico, "competencia", "valor", "Receita Bruta")
    fig.update_layout(height=400, legend=dict(orientation="h", y=1.12), margin=dict(t=40, b=60))
    st.plotly_chart(fig, width="stretch", theme=None, key="chart_receita_categoria")

    st.caption("Resumo por competência × categoria")
    resumo = agg.copy()
    resumo["competencia"] = resumo["competencia"].map(fmt_competencia_pt)
    resumo["valor"] = resumo["valor"].map(fmt_brl)
    st.dataframe(resumo, width="stretch", hide_index=True)

    st.markdown("**🔍 Detalhe por competência × categoria** — clique para expandir e ver as contas que compõem cada total.")
    for _, linha in agg.iterrows():
        tipo = TIPO_DIFERIMENTO.get(linha["categoria"], "A definir")
        titulo = (f"{fmt_competencia_pt(linha['competencia'])} — {linha['categoria']} — {fmt_brl(linha['valor'])} "
                  f"({int(linha['linhas'])} contas)")
        with st.expander(titulo):
            st.caption(f"Tipo de diferimento: **{tipo}**")
            detalhe = filtrado[(filtrado["competencia"] == linha["competencia"])
                               & (filtrado["categoria"] == linha["categoria"])].copy()
            detalhe["valor"] = detalhe["valor"].map(fmt_brl)
            st.dataframe(detalhe[["conta", "descricao", "valor"]], width="stretch", hide_index=True)


def render_aba_passivo_diferido(res):
    st.subheader("Composição do Passivo Diferido")
    blocos = res.blocos_controle
    if blocos.empty:
        st.info("Sem blocos de controle.")
        return

    c1, c2 = st.columns([1, 1])
    with c1:
        fig = go.Figure(go.Pie(labels=blocos["bloco"], values=blocos["passivo"], hole=0.55,
                               textinfo="label+percent", pull=[0, 0, 0, 0.08],
                               customdata=[fmt_brl(v) for v in blocos["passivo"]],
                               hovertemplate="%{label}<br>%{customdata} (%{percent})<extra></extra>",
                               marker=dict(colors=["#FF4F0D", "#FEC114", "#009F3C", "#000000"])))
        fig.update_layout(height=420, margin=dict(t=20, b=20),
                          annotations=[dict(text="Passivo Diferido", x=0.5, y=0.5, showarrow=False, font_size=14)])
        st.plotly_chart(fig, width="stretch", theme=None, key="chart_rosca_passivo_diferido")
    with c2:
        tabela = blocos[["bloco", "passivo", "receita"]].copy()
        tabela["passivo"] = tabela["passivo"].map(fmt_brl)
        tabela["receita"] = tabela["receita"].map(fmt_brl)
        st.dataframe(tabela, width="stretch", hide_index=True)
        st.metric("Passivo Total (Controle U1)", fmt_brl(res.passivo_total_controle))

    st.divider()
    st.markdown("**Split CP/LP** — calculado direto da U1 - Movimentação "
               "(réplica da fórmula real do Excel: G60/I60), sem depender da aba Dados_Gráfico")
    cp, lp = carregar_passivo_cp_lp()
    soma = cp + lp
    delta_cp_lp = soma - res.passivo_total_controle

    c3, c4, c5 = st.columns(3)
    c3.metric("Circulante (CP)", fmt_brl(cp))
    c4.metric("Não Circulante (LP)", fmt_brl(lp))
    c5.metric("Soma CP+LP vs Passivo Controle", fmt_brl(delta_cp_lp), delta=fmt_brl(delta_cp_lp),
              delta_color="normal" if abs(delta_cp_lp) < 1 else "inverse")
    st.caption("Soma CP + LP conferida contra o Passivo Total do Controle U1. Calculado somando "
               "passivo_cp/passivo_lp dos cabeçalhos de bloco (Subtotal Split Fee, Breakage, Custo "
               "do Produto, Spread) — a mesma fonte que a aba Dados_Gráfico usa (coluna 'CBSM'), "
               "mas sem os ajustes de Netpoints/Dotz Pay de lá, que ficam fora do escopo deste sistema.")


def render_aba_faturamento_pontos():
    render_cartao_titulo(["CBSM — Companhia Brasileira de Soluções de Marketing S/A",
                          "Faturamento de Pontos — 2026",
                          "Resumo do faturamento do ano atual por categoria e competência "
                          "(fonte: aba Fat. Análise de Vendas 2026)"])

    df = carregar_faturamento_vendas()
    if df.empty:
        st.info("Sem lançamentos de faturamento.")
        return

    resumo_comp_cat = (df.groupby(["competencia", "categoria"], as_index=False)
                         .agg(valor_vendas=("valor_vendas", "sum"), quantidade=("quantidade", "sum"))
                         .sort_values(["competencia", "categoria"]))
    resumo_comp_cat["competencia"] = competencia_para_pt(resumo_comp_cat["competencia"])

    resumo_categoria = (df.groupby("categoria", as_index=False)
                          .agg(valor_vendas=("valor_vendas", "sum"), quantidade=("quantidade", "sum"))
                          .sort_values("valor_vendas", ascending=False))

    c1, c2, c3 = st.columns(3)
    c1.metric("Faturamento Total (Valor Vendas)", fmt_brl(resumo_categoria["valor_vendas"].sum()))
    c2.metric("Quantidade de Pontos Total", fmt_numero_pt(resumo_categoria["quantidade"].sum()))
    c3.metric("Competências com lançamento", df["competencia"].nunique())

    fig = px.bar(resumo_comp_cat, x="competencia", y="valor_vendas", color="categoria", barmode="stack",
                 color_discrete_map=CORES_FATURAMENTO,
                 labels={"valor_vendas": "Valor de Vendas (R$)", "competencia": "Competência",
                         "categoria": "Categoria"})
    _formatar_hover_barras(fig)
    _adicionar_linha_total(fig, resumo_comp_cat, "competencia", "valor_vendas", "Faturamento Total")
    fig.update_layout(height=400, legend=dict(orientation="h", y=1.15, title="Categoria"), margin=dict(t=20, b=60))
    st.plotly_chart(fig, width="stretch", theme=None, key="chart_faturamento")

    st.caption("Resumo por categoria (todo o período)")
    tab_categoria = resumo_categoria.copy()
    tab_categoria["valor_vendas"] = tab_categoria["valor_vendas"].map(fmt_brl)
    tab_categoria["quantidade"] = tab_categoria["quantidade"].map(fmt_numero_pt)
    tab_categoria.columns = ["Categoria", "Valor Vendas", "Quantidade de Pontos"]
    st.dataframe(tab_categoria, width="stretch", hide_index=True)

    with st.expander("📅 Detalhe por competência × categoria"):
        tab_comp_cat = resumo_comp_cat.copy()
        tab_comp_cat["valor_vendas"] = tab_comp_cat["valor_vendas"].map(fmt_brl)
        tab_comp_cat["quantidade"] = tab_comp_cat["quantidade"].map(fmt_numero_pt)
        tab_comp_cat.columns = ["Competência", "Categoria", "Valor Vendas", "Quantidade de Pontos"]
        st.dataframe(tab_comp_cat, width="stretch", hide_index=True)
        st.download_button("📥 Exportar detalhe por competência (Excel)",
                           data=to_excel_bytes(resumo_comp_cat, "Faturamento"),
                           file_name="faturamento_por_competencia.xlsx", mime=MIME_XLSX)

    with st.expander("🔍 Ver lançamentos detalhados"):
        detalhe = df.copy()
        detalhe["valor_vendas"] = detalhe["valor_vendas"].map(fmt_brl)
        detalhe["quantidade"] = detalhe["quantidade"].map(fmt_numero_pt)
        st.dataframe(detalhe, width="stretch", hide_index=True)
        st.download_button("📥 Exportar lançamentos detalhados (Excel)",
                           data=to_excel_bytes(df, "Lancamentos"),
                           file_name="faturamento_lancamentos.xlsx", mime=MIME_XLSX)


def render_aba_placeholder(nome):
    st.info(f"🚧 Aba **{nome}** ainda em construção — será preenchida na próxima etapa.")


def render_cartao_grupo_u1(g):
    divergente = abs(g["check"]) > EPS_CHECK_LINHA
    status = "🔴" if divergente else "🟢"
    with st.container(border=True):
        st.markdown(f"##### {status} {g['nome']}")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Passivo Total (U1)", fmt_brl(g["passivo_total"]))
        c2.metric("Passivo Recalculado", fmt_brl(g["passivo_recalculado"]))
        c3.metric("Check (Δ)", fmt_brl(g["check"]),
                  delta=fmt_brl(g["check"]), delta_color="normal" if not divergente else "inverse")
        c4.metric("Receita", fmt_brl(g["receita"]))

        with st.expander(f"🔍 Expandir a nível conta contábil ({len(g['detalhe'])} contas)"):
            detalhe = g["detalhe"].copy()
            for col in ("passivo_cp", "passivo_lp", "passivo_total", "check", "passivo_recalculado", "receita"):
                detalhe[col] = detalhe[col].map(fmt_brl)
            detalhe.columns = ["Conta", "Descrição", "Passivo CP", "Passivo LP", "Passivo Total",
                               "Check", "Passivo Recalculado", "Receita"]
            st.dataframe(detalhe, width="stretch", hide_index=True)


def render_aba_controle_u1():
    render_cartao_titulo(["CBSM — Companhia Brasileira de Soluções de Marketing S/A",
                          "U1 — Movimentação: Conciliação por Grupo de Negócio",
                          "Cruzamento do Passivo Total (U1) contra o Passivo Recalculado, "
                          "grupo a grupo — os 4 grupos da planilha (nomes literais da coluna Descrição)"])

    grupos = carregar_u1_grupos()
    if not grupos:
        st.info("Sem grupos de controle U1.")
        return

    n_diverg = sum(1 for g in grupos if abs(g["check"]) > EPS_CHECK_LINHA)
    if n_diverg:
        st.warning(f"⚠️ {n_diverg} de {len(grupos)} grupo(s) com |check| > R$ {EPS_CHECK_LINHA:.2f}.")
    else:
        st.success(f"🟢 Todos os {len(grupos)} grupos conciliados dentro da tolerância.")

    for g in grupos:
        render_cartao_grupo_u1(g)


def _ordenar_serie_desc(dados):
    """Ordena decrescente: safra mais recente (2026) primeiro, mês a mês,
    com o subtotal 'Total' de cada ano ao final do seu próprio bloco."""
    ano_grupo = pd.to_numeric(dados.iloc[:, 0], errors="coerce").ffill()
    mes_num = pd.to_numeric(dados.iloc[:, 1], errors="coerce").fillna(0)
    ordem = (pd.DataFrame({"ano_grupo": ano_grupo, "mes_num": mes_num}, index=dados.index)
             .sort_values(["ano_grupo", "mes_num"], ascending=[False, False]).index)
    return dados.loc[ordem].reset_index(drop=True)


def render_tabela_serie_atuarial(dados):
    dados = _ordenar_serie_desc(dados)
    exibicao = pd.DataFrame(index=dados.index)
    for col in dados.columns:
        exibicao[col] = [fmt_valor_coluna(str(col), v) for v in dados[col]]
    exibicao.columns = [c if str(c).strip() else "—" for c in dados.columns]

    def estilo_total(row):
        e_total = str(dados.iloc[row.name, 0]).strip() == "Total"
        cor = "background-color:#eef2ff;color:#0f172a;font-weight:600;" if e_total else ""
        return [cor] * len(row)

    st.dataframe(exibicao.style.apply(estilo_total, axis=1), width="stretch", height=460, hide_index=True)


EPS_CHECK_DRE = 1000.0  # tolerância generosa: ruído de arredondamento na planilha é ~R$33


def render_check_dre(resumo, categoria_dre):
    """Replica o check da própria planilha (U1.5, linhas 143-152: linha
    'Grand total' vs. soma das contas reconhecidas na DRE, linha 'Check >>>')
    — compara a linha 'Total' (motor atuarial) com o valor efetivamente
    reconhecido na DRE 2026 para a mesma categoria, mês a mês, e também no
    acumulado até o mês de apuração corrente (sempre lido da própria DRE —
    hoje, junho/2026)."""
    linha_total = resumo[resumo["Safra"] == "Total"]
    if linha_total.empty:
        return
    linha_total = linha_total.iloc[0]

    dre = carregar_dre_categoria()
    mes_apuracao = dre["competencia"].max()
    reconhecido = dre[dre["categoria"] == categoria_dre].groupby("competencia")["valor"].sum()

    linhas = []
    for mes_nome, mes_num in data_loader.MESES_DRE.items():
        if mes_nome not in resumo.columns:
            continue
        competencia = f"2026-{mes_num}"
        if competencia > mes_apuracao or competencia not in reconhecido.index:
            continue
        calculado = float(linha_total[mes_nome])
        recon = float(reconhecido[competencia])
        linhas.append({"Mês": mes_nome, "Calculado (U1)": calculado,
                       "Reconhecido (DRE)": recon, "Check": calculado - recon})
    if not linhas:
        return
    tabela = pd.DataFrame(linhas)
    linha_acumulada = pd.DataFrame([{
        "Mês": "Total acumulado", "Calculado (U1)": tabela["Calculado (U1)"].sum(),
        "Reconhecido (DRE)": tabela["Reconhecido (DRE)"].sum(),
        "Check": tabela["Calculado (U1)"].sum() - tabela["Reconhecido (DRE)"].sum(),
    }])
    tabela_completa = pd.concat([tabela, linha_acumulada], ignore_index=True)

    st.markdown(f"**✅ Check — Conciliação com a DRE 2026** (categoria: {categoria_dre})")
    st.caption(f"Compara a linha 'Total' (motor atuarial) com o valor reconhecido na DRE 2026 "
               f"para a mesma categoria, mês a mês e no acumulado — replica o check já existente "
               f"na planilha (aba U1.5, linhas 143-152: 'Grand total' vs. 'Check >>>'). Mês de "
               f"apuração corrente (lido da DRE): **{fmt_competencia_pt(mes_apuracao)}**.")
    exibicao = tabela_completa.copy()
    for col in ("Calculado (U1)", "Reconhecido (DRE)", "Check"):
        exibicao[col] = exibicao[col].map(fmt_brl)

    def estilo_total(row):
        cor = "background-color:#eef2ff;color:#0f172a;font-weight:600;" if row["Mês"] == "Total acumulado" else ""
        return [cor] * len(row)

    st.dataframe(exibicao.style.apply(estilo_total, axis=1), width="stretch", hide_index=True)

    diverg_total = abs(float(linha_acumulada["Check"].iloc[0]))
    if diverg_total > EPS_CHECK_DRE:
        st.warning(f"⚠️ Divergência no acumulado: {fmt_brl(diverg_total)} "
                   f"(tolerância {fmt_brl(EPS_CHECK_DRE)}).")
    else:
        st.success(f"🟢 Conciliado com a DRE no acumulado — divergência {fmt_brl(diverg_total)}.")
    st.divider()


def render_quadro_receita_atual(dados, chave, categoria_dre=None):
    resumo = data_loader.resumo_safra_por_competencia(dados)
    st.markdown("**📅 Receita a Reconhecer por Safra × Competência (dentro do próprio ano de emissão)**")
    st.caption("Soma de todas as safras emitidas em cada ano (2022–2026), mês a mês — não inclui o "
               "residual de safras de anos anteriores ainda em amortização nesses meses.")

    meses_presentes = [m for m in data_loader.MESES_ORDEM if m in resumo.columns]
    exibicao = resumo.copy()
    for col in meses_presentes + (["Total"] if "Total" in resumo.columns else []):
        exibicao[col] = exibicao[col].map(fmt_brl)

    def estilo_total(row):
        cor = "background-color:#eef2ff;color:#0f172a;font-weight:600;" if row["Safra"] == "Total" else ""
        return [cor] * len(row)

    st.dataframe(exibicao.style.apply(estilo_total, axis=1), width="stretch", hide_index=True)

    if "Total" in resumo.columns:
        total_geral = float(resumo.loc[resumo["Safra"] != "Total", "Total"].sum())
        st.metric("Receita Total a Reconhecer (soma das safras 2022–2026)", fmt_brl(total_geral))

    if categoria_dre:
        render_check_dre(resumo, categoria_dre)

    longo = resumo[resumo["Safra"] != "Total"].melt(
        id_vars="Safra", value_vars=meses_presentes, var_name="Mês", value_name="Valor")
    longo["Mês"] = pd.Categorical(longo["Mês"], categories=data_loader.MESES_ORDEM, ordered=True)
    fig = px.bar(longo.sort_values("Mês"), x="Mês", y="Valor", color="Safra", barmode="stack",
                 color_discrete_map=PALETA_SAFRA,
                 labels={"Valor": "Receita a Reconhecer (R$)"})
    _formatar_hover_barras(fig)
    _adicionar_linha_total(fig, longo, "Mês", "Valor", "Total")
    fig.update_layout(height=380, margin=dict(t=20, b=60), legend=dict(orientation="h", y=1.15, title="Safra"))
    st.plotly_chart(fig, width="stretch", theme=None, key=f"chart_receita_atual_{chave}")
    st.divider()


def render_aba_serie_atuarial(nome_aba, col_ancora=1, categoria_dre=None):
    titulo, dados = carregar_serie_atuarial(nome_aba, col_ancora)
    render_cartao_titulo(titulo)
    render_quadro_receita_atual(dados, chave=nome_aba, categoria_dre=categoria_dre)
    st.caption(f"{len(dados)} linhas (safras Ano × Mês + subtotais anuais em destaque) — "
               f"ordem decrescente, 2026 primeiro.")
    render_tabela_serie_atuarial(dados)
    with st.expander("🔍 Ver grade completa (fiel ao Excel, sem tratamento)"):
        st.dataframe(carregar_aba_bruta(nome_aba), width="stretch", hide_index=True)


def render_aba_u1_6():
    meta, mensal = carregar_u1_6()
    render_cartao_titulo(["CBSM — Companhia Brasileira de Soluções de Marketing S/A",
                          "U1.6 — Emissão e Resgates",
                          "II - Receita com Resgate — todos os parceiros exceto Banco do Brasil"])

    exibicao = mensal.copy()
    exibicao["ano"] = exibicao["ano"].map(lambda v: fmt_numero_pt(v, 0))
    exibicao["mes"] = exibicao["mes"].map(lambda v: fmt_numero_pt(v, 0))
    exibicao["pontos_resgatados"] = mensal["pontos_resgatados"].map(lambda v: fmt_numero_pt(v, 0))
    exibicao["preco_negociado"] = mensal["preco_negociado"].map(lambda v: f"R$ {fmt_numero_pt(v, 4)}")
    exibicao["valor_faturado"] = mensal["valor_faturado"].map(fmt_brl)
    exibicao.columns = ["Ano", "Mês", "Pontos Resgatados", "Preço Negociado", "Valor Faturado"]
    st.dataframe(exibicao, width="stretch", hide_index=True)
    if meta["total_resgate"] is not None:
        st.metric("Total resgatado no período (todos os parceiros exceto BB)", fmt_brl(meta["total_resgate"]))

    if meta["tie_in"]:
        st.markdown("**Tie-in Contábil × Analyser**")
        tie_in_df = pd.DataFrame(
            [{"Item": k, "Valor": fmt_brl(v)} for k, v in meta["tie_in"].items()])
        st.dataframe(tie_in_df, width="stretch", hide_index=True)

    with st.expander("🔍 Ver grade completa (fiel ao Excel, sem tratamento)"):
        st.dataframe(carregar_aba_bruta(config.ABA_U1_6_RESGATES), width="stretch", hide_index=True)


def render_quadro_expectativa_realizacao():
    st.markdown("**📅 Expectativa de Realização da Receita** (por vintage × ano de realização)")
    st.caption("Fonte: aba Dados_Gráfico. Valores na planilha em milhares de reais — "
               "convertidos aqui para reais para consistência com o resto do app.")

    tabela = carregar_expectativa_realizacao()
    anos = [c for c in tabela.columns if c not in ("vintage", "total")]

    exibicao = tabela.copy()
    for col in ["total"] + anos:
        exibicao[col] = exibicao[col].map(fmt_brl)
    exibicao.columns = ["Vintage", "Total"] + anos

    def estilo_total(row):
        cor = "background-color:#eef2ff;color:#0f172a;font-weight:600;" if row["Vintage"] == "Total" else ""
        return [cor] * len(row)

    st.dataframe(exibicao.style.apply(estilo_total, axis=1), width="stretch", hide_index=True)

    longo = tabela[tabela["vintage"] != "Total"].melt(
        id_vars="vintage", value_vars=anos, var_name="Ano de Realização", value_name="Valor")
    fig = px.bar(longo, x="Ano de Realização", y="Valor", color="vintage", barmode="stack",
                 color_discrete_map=PALETA_SAFRA,
                 labels={"Valor": "Receita a Realizar (R$)", "vintage": "Vintage"})
    _formatar_hover_barras(fig)
    _adicionar_linha_total(fig, longo, "Ano de Realização", "Valor", "Total")
    fig.update_layout(height=380, margin=dict(t=20, b=60), legend=dict(orientation="h", y=1.15, title="Vintage"))
    st.plotly_chart(fig, width="stretch", theme=None, key="chart_expectativa_realizacao")


def render_aba_dados_grafico():
    render_cartao_titulo(["CBSM — Companhia Brasileira de Soluções de Marketing S/A",
                          "Dados_Gráfico — Consolidado",
                          "Base para os gráficos do dashboard original"])
    render_quadro_expectativa_realizacao()
    st.divider()

    consolidado, cp_lp = carregar_dados_grafico()
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Receitas diferidas e prêmios a distribuir**")
        tab = consolidado.copy()
        tab["em_30_06_2026"] = tab["em_30_06_2026"].map(fmt_brl)
        tab["em_31_12_2025"] = tab["em_31_12_2025"].map(fmt_brl)
        tab.columns = ["Item", "Em 30/06/2026", "Em 31/12/2025"]
        st.dataframe(tab, width="stretch", hide_index=True)
    with c2:
        st.markdown("**Receita Diferida CP/LP** (bate com o Passivo Total Controle)")
        tab2 = cp_lp.copy()
        tab2["categoria"] = tab2["categoria"].fillna("Total (CP + LP)")
        for col in ("total", "cbsm", "netpoints", "dotz_pay", "total_ajustado"):
            tab2[col] = tab2[col].map(fmt_brl)
        tab2.columns = ["Categoria", "Total", "CBSM", "Netpoints", "Dotz Pay", "Total Ajustado"]
        st.dataframe(tab2, width="stretch", hide_index=True)

    with st.expander("🔍 Ver grade completa (fiel ao Excel, sem tratamento — inclui dados de outros gráficos)"):
        st.dataframe(carregar_aba_bruta(config.ABA_DADOS_GRAFICO), width="stretch", hide_index=True)


def main():
    # Evita que o navegador traduza automaticamente termos de negócio em
    # inglês (ex.: "Breakage", "Spread") para português, o que corrompe os
    # rótulos das categorias (ex.: "Breakage" -> "Quebra"). st.markdown com
    # unsafe_allow_html remove <script>; st.html com
    # unsafe_allow_javascript=True permite o script marcar a página inteira
    # como "translate=no", igual a <meta name="google" content="notranslate">
    # no <head> real.
    st.html("""
    <script>
    document.documentElement.setAttribute('translate', 'no');
    document.documentElement.classList.add('notranslate');
    if (!document.querySelector('meta[name="google"]')) {
        var meta = document.createElement('meta');
        meta.name = 'google';
        meta.content = 'notranslate';
        document.head.appendChild(meta);
    }
    </script>
    """, unsafe_allow_javascript=True)
    _aplicar_logo_canto()
    _render_cabecalho()
    try:
        res = carregar_resultado(None)
    except (FileNotFoundError, ValueError, KeyError) as e:
        st.error(f"Falha ao processar o pipeline: {e}")
        st.stop()

    (aba_visao, aba_passivo, aba_receita, aba_faturamento, aba_u1, aba_u14,
     aba_u15, aba_u16, aba_lancamentos) = st.tabs([
        "📊 Visão Geral / Conciliação",
        "📉 Passivo Diferido",
        "💰 Receita Diferida por Categoria",
        "🧾 Faturamento de Pontos",
        "🔧 Controle U1",
        "U1.4_Par._Emissão Expiração",
        "U1.5_Par._Emissão Margem",
        "U1.6_Emissão_Resgates",
        "🧾 Lançamentos de Ajuste",
    ])

    with aba_visao:
        render_aba_visao_geral(res)
    with aba_passivo:
        render_aba_passivo_diferido(res)
    with aba_receita:
        render_aba_receita_motor()
    with aba_faturamento:
        render_aba_faturamento_pontos()
    with aba_u1:
        render_aba_controle_u1()
    with aba_u14:
        render_aba_serie_atuarial(config.ABA_U1_4_EXPIRACAO, col_ancora=1, categoria_dre="Breakage")
    with aba_u15:
        render_aba_serie_atuarial(config.ABA_U1_5_MARGEM, col_ancora=1, categoria_dre="Spread")
    with aba_u16:
        render_aba_u1_6()
    with aba_lancamentos:
        render_lancamentos(res)


if __name__ == "__main__":
    main()