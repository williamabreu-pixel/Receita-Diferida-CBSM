# export_html.py
"""Gera um snapshot estático em HTML (arquivo único, sem servidor) com o
mesmo conteúdo e identidade visual do dashboard Streamlit (app.py): todas as
9 abas, cores da Dotz, receita por categoria (base DRE), split CP/LP, etc.
Não inclui os expanders "ver grade completa (fiel ao Excel)" de cada aba
(dumps brutos da planilha) nem filtros interativos — é um retrato estático
para compartilhar por e-mail/rede, não substitui o app rodando.

Rodar com:  python export_html.py   (gera ../dashboard.html)
"""

from __future__ import annotations

import base64
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio

import config
import data_loader
import data_pipeline
import gerador_lancamentos

pio.templates.default = "plotly_dark"

ARQUIVO_SAIDA = "../dashboard.html"

CORES_MOTOR = {"Breakage": "#FF4F0D", "Trocas": "#FEC114",
               "Spread": "#009F3C", "Promodotz": "#000000",
               "Receita Projetos Especiais": "#D82598"}

CORES_FATURAMENTO = {
    "NSF": "#FF4F0D", "SF": "#FEC114", "BV": "#009F3C",
    "Viaja Dotz": "#D82598", "Comissão com Conversão de Dotz": "#AF1010",
    "MKT": "#000000", "Prestação de Serviços Comissão -  Banco do Brasil": "#E3D2C8",
}

PALETA_SAFRA = {"2022": "#FFD9C7", "2023": "#FFAD8A", "2024": "#FF7A47",
                "2025": "#FF4F0D", "2026": "#C93D0A"}

TIPO_DIFERIMENTO = {
    "Breakage": "Diferimento 48 meses (linear por competência de emissão)",
    "Spread": "Diferimento 48 meses (linear por competência de emissão)",
    "Trocas": "Fato gerador (resgate)",
    "Promodotz": "Fato gerador (evento)",
    "Receita Projetos Especiais": "Fato gerador (resgate)",
}

RENOMEIA_QUADRO = {"A": "Total Passivo", "G": "Passivo Recalculado"}

MESES_ABREV_PT = {"01": "Jan", "02": "Fev", "03": "Mar", "04": "Abr", "05": "Mai", "06": "Jun",
                  "07": "Jul", "08": "Ago", "09": "Set", "10": "Out", "11": "Nov", "12": "Dez"}

EPS_CHECK_LINHA = 0.01
EPS_CHECK_DRE = 1000.0

CAMINHO_LOGO = "logo_dotz.png"


# --- Formatação (mesmas regras do app.py) -----------------------------------

def fmt_brl(valor):
    if valor is None or pd.isna(valor):
        return "—"
    s = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def fmt_numero_pt(valor, decimais=0):
    if valor is None or pd.isna(valor):
        return "—"
    s = f"{valor:,.{decimais}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_competencia_pt(competencia):
    if not isinstance(competencia, str) or "-" not in competencia:
        return competencia
    ano, mes = competencia.split("-")
    return f"{MESES_ABREV_PT.get(mes, mes)}/{ano}"


def competencia_para_pt(serie_competencia):
    ordem_original = sorted(serie_competencia.dropna().unique())
    rotulos_ordenados = [fmt_competencia_pt(c) for c in ordem_original]
    mapa = dict(zip(ordem_original, rotulos_ordenados))
    return pd.Categorical(serie_competencia.map(mapa), categories=rotulos_ordenados, ordered=True)


def fmt_valor_coluna(nome_col, valor):
    if not isinstance(valor, (int, float)) or pd.isna(valor):
        return "—" if (valor is None or pd.isna(valor)) else valor
    nome = nome_col.lower()
    if nome in ("ano", "mês", "mes"):
        return str(int(valor))
    if "preço" in nome:
        return f"R$ {fmt_numero_pt(valor, 6)}"
    if nome.startswith("(sem rótulo"):
        return fmt_numero_pt(valor, 2)
    if "pontos" in nome:
        return fmt_numero_pt(valor, 0)
    return fmt_brl(valor)


def _formatar_hover_barras(fig):
    for trace in fig.data:
        if trace.type == "bar":
            trace.customdata = [fmt_brl(v) for v in trace.y]
            trace.hovertemplate = f"%{{x}}<br>{trace.name}: %{{customdata}}<extra></extra>"


def _adicionar_linha_total(fig, df, col_x, col_y, nome="Total"):
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


_PLOTLY_JA_EMBUTIDO = False


def fig_html(fig):
    """Embute plotly.js inline só na 1ª figura da página (offline, sem CDN);
    as demais reutilizam a mesma cópia já carregada."""
    global _PLOTLY_JA_EMBUTIDO
    incluir = not _PLOTLY_JA_EMBUTIDO
    _PLOTLY_JA_EMBUTIDO = True
    return pio.to_html(fig, include_plotlyjs=incluir, full_html=False)


def tabela_html(df, moeda_cols=(), destaque_col=None, destaque_valor="Total", status_col=None):
    linhas_html = []
    header = "".join(f"<th>{c}</th>" for c in df.columns)
    for _, row in df.iterrows():
        destaque = destaque_col is not None and str(row[destaque_col]).strip() == destaque_valor
        cells = []
        for c in df.columns:
            v = row[c]
            if pd.isna(v):
                v = "—"
            elif c in moeda_cols and isinstance(v, (int, float)):
                v = fmt_brl(v)
            if c == status_col:
                cor = "#009F3C" if str(v) == "OK" else "#AF1010"
                cells.append(f'<td style="color:{cor};font-weight:600;">{v}</td>')
            else:
                cells.append(f"<td>{v}</td>")
        cls = ' class="destaque"' if destaque else ""
        linhas_html.append(f"<tr{cls}>{''.join(cells)}</tr>")
    return (f'<div class="tabela-wrap"><table class="tabela"><thead><tr>{header}</tr></thead>'
            f'<tbody>{"".join(linhas_html)}</tbody></table></div>')


def cartao_titulo_html(linhas):
    if not linhas:
        return ""
    empresa = linhas[0] if len(linhas) > 0 else ""
    titulo = linhas[1] if len(linhas) > 1 else ""
    nota = linhas[2] if len(linhas) > 2 else ""
    parametros = linhas[3:]
    badges = "".join(
        f'<span class="badge">{parametros[i]}: <b>{parametros[i + 1]}</b></span>'
        for i in range(0, len(parametros) - 1, 2)
    )
    return f"""<div class="cartao-titulo">
  <div class="empresa">{empresa}</div>
  <div class="titulo">{titulo}</div>
  <div class="nota">{nota}</div>
  {f'<div class="badges">{badges}</div>' if badges else ''}
</div>"""


def metric_html(label, valor, delta=None, ok=None):
    cor = ""
    if ok is True:
        cor = "color:#009F3C;"
    elif ok is False:
        cor = "color:#AF1010;"
    delta_html = f'<div class="metric-delta" style="{cor}">{delta}</div>' if delta is not None else ""
    return f'<div class="metric"><div class="metric-label">{label}</div><div class="metric-value">{valor}</div>{delta_html}</div>'


def alerta_html(texto, tipo="warning"):
    return f'<div class="alerta alerta-{tipo}">{texto}</div>'


def status_html(ok, texto_ok, texto_falha):
    return alerta_html(texto_ok, "success") if ok else alerta_html(texto_falha, "error")


# --- Aba 1: Visão Geral / Conciliação ---------------------------------------

def secao_receita_barras(dre_categoria):
    detalhe = dre_categoria
    if detalhe.empty:
        return "<p>Sem dados de receita.</p>"
    agg = detalhe.groupby(["competencia", "categoria"], as_index=False)["valor"].sum()
    agg["competencia"] = competencia_para_pt(agg["competencia"])
    fig = px.bar(agg, x="competencia", y="valor", color="categoria",
                 color_discrete_map=CORES_MOTOR, barmode="stack",
                 labels={"valor": "Valor (R$)", "competencia": "Competência", "categoria": "Categoria"})
    _formatar_hover_barras(fig)
    _adicionar_linha_total(fig, agg, "competencia", "valor", "Receita Bruta")
    fig.update_layout(height=430, legend=dict(orientation="h", y=1.12), margin=dict(t=40, b=60))
    return f"<h3>Receita por Categoria</h3><p class='caption'>Fonte: aba DRE 2026 (coluna Classificação Receita).</p>{fig_html(fig)}"


def secao_passivo_rosca(res):
    blocos = res.blocos_controle
    if blocos.empty:
        return "<p>Sem blocos de controle.</p>"
    fig = go.Figure(go.Pie(labels=blocos["bloco"], values=blocos["passivo"], hole=0.55,
                           textinfo="label+percent", pull=[0, 0, 0, 0.08],
                           customdata=[fmt_brl(v) for v in blocos["passivo"]],
                           hovertemplate="%{label}<br>%{customdata} (%{percent})<extra></extra>",
                           marker=dict(colors=["#FF4F0D", "#FEC114", "#009F3C", "#000000"])))
    fig.update_layout(height=430, margin=dict(t=30, b=20),
                      annotations=[dict(text="Passivo Diferido", x=0.5, y=0.5, showarrow=False, font_size=14)])
    return f"<h3>Composição do Passivo Diferido</h3>{fig_html(fig)}"


def secao_conciliacao(res):
    df = res.conciliacao.copy()
    df["quadro"] = df["quadro"].map(lambda q: RENOMEIA_QUADRO.get(q, q))
    return f"<h3>Conciliação</h3>{tabela_html(df, moeda_cols=('controle', 'contabil', 'delta'), status_col='status')}"


def secao_expectativa_realizacao():
    tabela = data_loader.load_dados_grafico_expectativa_realizacao()
    anos = [c for c in tabela.columns if c not in ("vintage", "total")]
    exibicao = tabela.copy()
    for col in ["total"] + anos:
        exibicao[col] = exibicao[col].map(fmt_brl)
    exibicao.columns = ["Vintage", "Total"] + anos

    longo = tabela[tabela["vintage"] != "Total"].melt(
        id_vars="vintage", value_vars=anos, var_name="Ano de Realização", value_name="Valor")
    fig = px.bar(longo, x="Ano de Realização", y="Valor", color="vintage", barmode="stack",
                 color_discrete_map=PALETA_SAFRA,
                 labels={"Valor": "Receita a Realizar (R$)", "vintage": "Vintage"})
    _formatar_hover_barras(fig)
    _adicionar_linha_total(fig, longo, "Ano de Realização", "Valor", "Total")
    fig.update_layout(height=380, margin=dict(t=20, b=60), legend=dict(orientation="h", y=1.15, title="Vintage"))

    return (f"<h4>📅 Expectativa de Realização da Receita</h4>"
            f"<p class='caption'>Fonte: aba Dados_Gráfico (valores convertidos de milhares para reais).</p>"
            f"{tabela_html(exibicao, destaque_col='Vintage')}{fig_html(fig)}")


def secao_dados_grafico_consolidado():
    consolidado, cp_lp = data_loader.load_dados_grafico_resumo()
    tab = consolidado.copy()
    tab["em_30_06_2026"] = tab["em_30_06_2026"].map(fmt_brl)
    tab["em_31_12_2025"] = tab["em_31_12_2025"].map(fmt_brl)
    tab.columns = ["Item", "Em 30/06/2026", "Em 31/12/2025"]

    tab2 = cp_lp.copy()
    tab2["categoria"] = tab2["categoria"].fillna("Total (CP + LP)")
    for col in ("total", "cbsm", "netpoints", "dotz_pay", "total_ajustado"):
        tab2[col] = tab2[col].map(fmt_brl)
    tab2.columns = ["Categoria", "Total", "CBSM", "Netpoints", "Dotz Pay", "Total Ajustado"]

    return f"""<div class="col2">
  <div><h4>Receitas diferidas e prêmios a distribuir</h4>{tabela_html(tab)}</div>
  <div><h4>Receita Diferida CP/LP</h4>{tabela_html(tab2)}</div>
</div>"""


def aba_visao_geral_html(res, dre_categoria):
    partes = [status_html(res.conciliacao_ok,
                          "🟢 Conciliação 100% OK — todos os blocos dentro da tolerância.",
                          "🔴 Conciliação com divergência(s) — verifique os blocos abaixo.")]
    for a in res.alertas:
        partes.append(alerta_html(a, "warning"))

    delta = res.passivo_total_controle - res.passivo_total_contabil
    partes.append('<div class="cards">' +
                  metric_html("Passivo Contábil (Balancete)", fmt_brl(res.passivo_total_contabil)) +
                  metric_html("Passivo Controle (U1)", fmt_brl(res.passivo_total_controle)) +
                  metric_html("Delta (Controle − Contábil)", fmt_brl(delta), fmt_brl(delta), abs(delta) < 1) +
                  '</div>')

    partes.append('<div class="col2">'
                  f'<div>{secao_receita_barras(dre_categoria)}</div>'
                  f'<div>{secao_passivo_rosca(res)}</div>'
                  '</div>')
    partes.append(secao_conciliacao(res))
    partes.append("<h3>Dados_Gráfico — Consolidado</h3>")
    partes.append(secao_expectativa_realizacao())
    partes.append(secao_dados_grafico_consolidado())
    return "".join(partes)


# --- Aba 2: Passivo Diferido -------------------------------------------------

def aba_passivo_diferido_html(res):
    blocos = res.blocos_controle
    if blocos.empty:
        return "<p>Sem blocos de controle.</p>"
    tabela = blocos[["bloco", "passivo", "receita"]].copy()
    tabela["passivo"] = tabela["passivo"].map(fmt_brl)
    tabela["receita"] = tabela["receita"].map(fmt_brl)

    cp, lp = data_loader.calcular_passivo_cp_lp()
    soma = cp + lp
    delta_cp_lp = soma - res.passivo_total_controle

    html = ['<div class="col2">']
    html.append(f"<div>{secao_passivo_rosca(res)}</div>")
    html.append(f"<div>{tabela_html(tabela)}{metric_html('Passivo Total (Controle U1)', fmt_brl(res.passivo_total_controle))}</div>")
    html.append("</div><hr>")
    html.append('<h4>Split CP/LP</h4><p class="caption">Calculado direto da U1 - Movimentação '
               '(réplica da fórmula real do Excel: G60/I60), sem depender da aba Dados_Gráfico.</p>')
    html.append('<div class="cards">' +
                metric_html("Circulante (CP)", fmt_brl(cp)) +
                metric_html("Não Circulante (LP)", fmt_brl(lp)) +
                metric_html("Soma CP+LP vs Passivo Controle", fmt_brl(delta_cp_lp), fmt_brl(delta_cp_lp), abs(delta_cp_lp) < 1) +
                '</div>')
    html.append('<p class="caption">Soma CP + LP conferida contra o Passivo Total do Controle U1. Calculado somando '
                'passivo_cp/passivo_lp dos cabeçalhos de bloco (Subtotal Split Fee, Breakage, Custo do Produto, '
                'Spread) — a mesma fonte que a aba Dados_Gráfico usa (coluna "CBSM"), mas sem os ajustes de '
                'Netpoints/Dotz Pay de lá, que ficam fora do escopo deste sistema.</p>')
    return "".join(html)


# --- Aba 3: Receita Diferida por Categoria -----------------------------------

def aba_receita_categoria_html(dre_categoria):
    df = dre_categoria
    if df.empty:
        return "<p>Sem dados de receita.</p>"
    partes = ['<p class="caption">Fonte: aba DRE 2026 (coluna Classificação Receita) — receita reconhecida no '
             'resultado. "Trocas" aqui é um valor diferente do "Custo do Produto" usado na conciliação (Visão '
             'Geral/Controle U1), pois vem de uma base contábil distinta.</p>']

    agg = (df.groupby(["competencia", "categoria"], as_index=False)
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
    partes.append(fig_html(fig))

    partes.append("<h4>Resumo por competência × categoria</h4>")
    resumo = agg.copy()
    resumo["competencia"] = resumo["competencia"].map(fmt_competencia_pt)
    resumo_exib = resumo.copy()
    resumo_exib["valor"] = resumo_exib["valor"].map(fmt_brl)
    partes.append(tabela_html(resumo_exib))

    partes.append("<h4>🔍 Detalhe por competência × categoria</h4>")
    for _, linha in agg.iterrows():
        tipo = TIPO_DIFERIMENTO.get(linha["categoria"], "A definir")
        titulo = (f"{fmt_competencia_pt(linha['competencia'])} — {linha['categoria']} — {fmt_brl(linha['valor'])} "
                  f"({int(linha['linhas'])} contas)")
        detalhe = df[(df["competencia"] == linha["competencia"]) & (df["categoria"] == linha["categoria"])].copy()
        detalhe["valor"] = detalhe["valor"].map(fmt_brl)
        partes.append(f"""<details><summary>{titulo}</summary>
<p class="caption">Tipo de diferimento: <b>{tipo}</b></p>
{tabela_html(detalhe[["conta", "descricao", "valor"]])}
</details>""")
    return "".join(partes)


# --- Aba 4: Faturamento de Pontos --------------------------------------------

def aba_faturamento_html():
    df = data_loader.load_faturamento_vendas()
    if df.empty:
        return "<p>Sem lançamentos de faturamento.</p>"

    resumo_comp_cat = (df.groupby(["competencia", "categoria"], as_index=False)
                         .agg(valor_vendas=("valor_vendas", "sum"), quantidade=("quantidade", "sum"))
                         .sort_values(["competencia", "categoria"]))
    resumo_comp_cat_grafico = resumo_comp_cat.copy()
    resumo_comp_cat_grafico["competencia"] = competencia_para_pt(resumo_comp_cat_grafico["competencia"])

    resumo_categoria = (df.groupby("categoria", as_index=False)
                          .agg(valor_vendas=("valor_vendas", "sum"), quantidade=("quantidade", "sum"))
                          .sort_values("valor_vendas", ascending=False))

    partes = [cartao_titulo_html(["CBSM — Companhia Brasileira de Soluções de Marketing S/A",
                                  "Faturamento de Pontos — 2026",
                                  "Resumo do faturamento do ano atual por categoria e competência "
                                  "(fonte: aba Fat. Análise de Vendas 2026)"])]

    partes.append('<div class="cards">' +
                  metric_html("Faturamento Total (Valor Vendas)", fmt_brl(resumo_categoria["valor_vendas"].sum())) +
                  metric_html("Quantidade de Pontos Total", fmt_numero_pt(resumo_categoria["quantidade"].sum())) +
                  metric_html("Competências com lançamento", str(df["competencia"].nunique())) +
                  '</div>')

    fig = px.bar(resumo_comp_cat_grafico, x="competencia", y="valor_vendas", color="categoria", barmode="stack",
                 color_discrete_map=CORES_FATURAMENTO,
                 labels={"valor_vendas": "Valor de Vendas (R$)", "competencia": "Competência", "categoria": "Categoria"})
    _formatar_hover_barras(fig)
    _adicionar_linha_total(fig, resumo_comp_cat_grafico, "competencia", "valor_vendas", "Faturamento Total")
    fig.update_layout(height=400, legend=dict(orientation="h", y=1.15, title="Categoria"), margin=dict(t=20, b=60))
    partes.append(fig_html(fig))

    partes.append("<h4>Resumo por categoria (todo o período)</h4>")
    tab_categoria = resumo_categoria.copy()
    tab_categoria["valor_vendas"] = tab_categoria["valor_vendas"].map(fmt_brl)
    tab_categoria["quantidade"] = tab_categoria["quantidade"].map(fmt_numero_pt)
    tab_categoria.columns = ["Categoria", "Valor Vendas", "Quantidade de Pontos"]
    partes.append(tabela_html(tab_categoria))

    tab_comp_cat = resumo_comp_cat.copy()
    tab_comp_cat["competencia"] = tab_comp_cat["competencia"].map(fmt_competencia_pt)
    tab_comp_cat["valor_vendas"] = tab_comp_cat["valor_vendas"].map(fmt_brl)
    tab_comp_cat["quantidade"] = tab_comp_cat["quantidade"].map(fmt_numero_pt)
    tab_comp_cat.columns = ["Competência", "Categoria", "Valor Vendas", "Quantidade de Pontos"]
    partes.append(f"<details><summary>📅 Detalhe por competência × categoria</summary>{tabela_html(tab_comp_cat)}</details>")

    detalhe = df.copy()
    detalhe["valor_vendas"] = detalhe["valor_vendas"].map(fmt_brl)
    detalhe["quantidade"] = detalhe["quantidade"].map(fmt_numero_pt)
    detalhe["data_emissao"] = detalhe["data_emissao"].dt.strftime("%d/%m/%Y")
    partes.append(f"<details><summary>🔍 Ver lançamentos detalhados ({len(detalhe)})</summary>{tabela_html(detalhe)}</details>")
    return "".join(partes)


# --- Aba 5: Controle U1 ------------------------------------------------------

def cartao_grupo_u1_html(g):
    divergente = abs(g["check"]) > EPS_CHECK_LINHA
    status = "🔴" if divergente else "🟢"
    detalhe = g["detalhe"].copy()
    for col in ("passivo_cp", "passivo_lp", "passivo_total", "check", "passivo_recalculado", "receita"):
        detalhe[col] = detalhe[col].map(fmt_brl)
    detalhe.columns = ["Conta", "Descrição", "Passivo CP", "Passivo LP", "Passivo Total",
                       "Check", "Passivo Recalculado", "Receita"]
    return f"""<div class="container-borda">
  <h5>{status} {g['nome']}</h5>
  <div class="cards">
    {metric_html("Passivo Total (U1)", fmt_brl(g['passivo_total']))}
    {metric_html("Passivo Recalculado", fmt_brl(g['passivo_recalculado']))}
    {metric_html("Check (Δ)", fmt_brl(g['check']), fmt_brl(g['check']), not divergente)}
    {metric_html("Receita", fmt_brl(g['receita']))}
  </div>
  <details><summary>🔍 Expandir a nível conta contábil ({len(g['detalhe'])} contas)</summary>
    {tabela_html(detalhe)}
  </details>
</div>"""


def aba_controle_u1_html():
    grupos = data_loader.load_controle_u1_grupos()
    if not grupos:
        return "<p>Sem grupos de controle U1.</p>"

    partes = [cartao_titulo_html(["CBSM — Companhia Brasileira de Soluções de Marketing S/A",
                                  "U1 — Movimentação: Conciliação por Grupo de Negócio",
                                  "Cruzamento do Passivo Total (U1) contra o Passivo Recalculado, "
                                  "grupo a grupo — os 4 grupos da planilha (nomes literais da coluna Descrição)"])]

    n_diverg = sum(1 for g in grupos if abs(g["check"]) > EPS_CHECK_LINHA)
    partes.append(status_html(n_diverg == 0,
                              f"🟢 Todos os {len(grupos)} grupos conciliados dentro da tolerância.",
                              f"⚠️ {n_diverg} de {len(grupos)} grupo(s) com |check| > R$ {EPS_CHECK_LINHA:.2f}."))
    for g in grupos:
        partes.append(cartao_grupo_u1_html(g))
    return "".join(partes)


# --- Abas 6/7: U1.4 (Breakage) / U1.5 (Spread) -------------------------------

def _ordenar_serie_desc(dados):
    ano_grupo = pd.to_numeric(dados.iloc[:, 0], errors="coerce").ffill()
    mes_num = pd.to_numeric(dados.iloc[:, 1], errors="coerce").fillna(0)
    ordem = (pd.DataFrame({"ano_grupo": ano_grupo, "mes_num": mes_num}, index=dados.index)
             .sort_values(["ano_grupo", "mes_num"], ascending=[False, False]).index)
    return dados.loc[ordem].reset_index(drop=True)


def tabela_serie_atuarial_html(dados):
    dados = _ordenar_serie_desc(dados)
    exibicao = pd.DataFrame(index=dados.index)
    for col in dados.columns:
        exibicao[col] = [fmt_valor_coluna(str(col), v) for v in dados[col]]
    exibicao.columns = [c if str(c).strip() else "—" for c in dados.columns]
    return tabela_html(exibicao, destaque_col=exibicao.columns[0], destaque_valor="Total")


def check_dre_html(resumo, categoria_dre, dre_categoria):
    linha_total = resumo[resumo["Safra"] == "Total"]
    if linha_total.empty:
        return ""
    linha_total = linha_total.iloc[0]

    dre = dre_categoria
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
        return ""
    tabela = pd.DataFrame(linhas)
    linha_acumulada = pd.DataFrame([{
        "Mês": "Total acumulado", "Calculado (U1)": tabela["Calculado (U1)"].sum(),
        "Reconhecido (DRE)": tabela["Reconhecido (DRE)"].sum(),
        "Check": tabela["Calculado (U1)"].sum() - tabela["Reconhecido (DRE)"].sum(),
    }])
    tabela_completa = pd.concat([tabela, linha_acumulada], ignore_index=True)
    exibicao = tabela_completa.copy()
    for col in ("Calculado (U1)", "Reconhecido (DRE)", "Check"):
        exibicao[col] = exibicao[col].map(fmt_brl)

    diverg_total = abs(float(linha_acumulada["Check"].iloc[0]))
    status = status_html(diverg_total <= EPS_CHECK_DRE,
                         f"🟢 Conciliado com a DRE no acumulado — divergência {fmt_brl(diverg_total)}.",
                         f"⚠️ Divergência no acumulado: {fmt_brl(diverg_total)} (tolerância {fmt_brl(EPS_CHECK_DRE)}).")

    return f"""<h4>✅ Check — Conciliação com a DRE 2026 (categoria: {categoria_dre})</h4>
<p class="caption">Compara a linha "Total" (motor atuarial) com o valor reconhecido na DRE 2026 para a mesma
categoria, mês a mês e no acumulado. Mês de apuração corrente (lido da DRE): <b>{fmt_competencia_pt(mes_apuracao)}</b>.</p>
{tabela_html(exibicao, destaque_col="Mês", destaque_valor="Total acumulado")}
{status}"""


def aba_serie_atuarial_html(nome_aba, dre_categoria, col_ancora=1, categoria_dre=None):
    titulo, dados = data_loader.load_serie_atuarial(nome_aba, col_ancora)
    partes = [cartao_titulo_html(titulo)]

    resumo = data_loader.resumo_safra_por_competencia(dados)
    partes.append("<h4>📅 Receita a Reconhecer por Safra × Competência (dentro do próprio ano de emissão)</h4>")
    partes.append('<p class="caption">Soma de todas as safras emitidas em cada ano (2022–2026), mês a mês — não '
                  'inclui o residual de safras de anos anteriores ainda em amortização nesses meses.</p>')
    exibicao = resumo.copy()
    meses_presentes = [m for m in data_loader.MESES_ORDEM if m in resumo.columns]
    for col in meses_presentes + (["Total"] if "Total" in resumo.columns else []):
        exibicao[col] = exibicao[col].map(fmt_brl)
    partes.append(tabela_html(exibicao, destaque_col="Safra"))

    if "Total" in resumo.columns:
        total_geral = float(resumo.loc[resumo["Safra"] != "Total", "Total"].sum())
        partes.append(metric_html("Receita Total a Reconhecer (soma das safras 2022–2026)", fmt_brl(total_geral)))

    if categoria_dre:
        partes.append(check_dre_html(resumo, categoria_dre, dre_categoria))

    longo = resumo[resumo["Safra"] != "Total"].melt(
        id_vars="Safra", value_vars=meses_presentes, var_name="Mês", value_name="Valor")
    longo["Mês"] = pd.Categorical(longo["Mês"], categories=data_loader.MESES_ORDEM, ordered=True)
    fig = px.bar(longo.sort_values("Mês"), x="Mês", y="Valor", color="Safra", barmode="stack",
                 color_discrete_map=PALETA_SAFRA, labels={"Valor": "Receita a Reconhecer (R$)"})
    _formatar_hover_barras(fig)
    _adicionar_linha_total(fig, longo, "Mês", "Valor", "Total")
    fig.update_layout(height=380, margin=dict(t=20, b=60), legend=dict(orientation="h", y=1.15, title="Safra"))
    partes.append(fig_html(fig))

    partes.append(f"<p class='caption'>{len(dados)} linhas (safras Ano × Mês + subtotais anuais em destaque) — "
                  f"ordem decrescente, 2026 primeiro.</p>")
    partes.append(tabela_serie_atuarial_html(dados))
    return "".join(partes)


# --- Aba 8: U1.6 --------------------------------------------------------------

def aba_u1_6_html():
    meta, mensal = data_loader.load_u1_6_resgates()
    partes = [cartao_titulo_html(["CBSM — Companhia Brasileira de Soluções de Marketing S/A",
                                  "U1.6 — Emissão e Resgates",
                                  "II - Receita com Resgate — todos os parceiros exceto Banco do Brasil"])]

    exibicao = mensal.copy()
    exibicao["ano"] = exibicao["ano"].map(lambda v: fmt_numero_pt(v, 0))
    exibicao["mes"] = exibicao["mes"].map(lambda v: fmt_numero_pt(v, 0))
    exibicao["pontos_resgatados"] = mensal["pontos_resgatados"].map(lambda v: fmt_numero_pt(v, 0))
    exibicao["preco_negociado"] = mensal["preco_negociado"].map(lambda v: f"R$ {fmt_numero_pt(v, 4)}")
    exibicao["valor_faturado"] = mensal["valor_faturado"].map(fmt_brl)
    exibicao.columns = ["Ano", "Mês", "Pontos Resgatados", "Preço Negociado", "Valor Faturado"]
    partes.append(tabela_html(exibicao))

    if meta["total_resgate"] is not None:
        partes.append(metric_html("Total resgatado no período (todos os parceiros exceto BB)", fmt_brl(meta["total_resgate"])))

    if meta["tie_in"]:
        tie_in_df = pd.DataFrame([{"Item": k, "Valor": fmt_brl(v)} for k, v in meta["tie_in"].items()])
        partes.append("<h4>Tie-in Contábil × Analyser</h4>")
        partes.append(tabela_html(tie_in_df))
    return "".join(partes)


# --- Aba 9: Lançamentos de Ajuste --------------------------------------------

def aba_lancamentos_html(res):
    diario = gerador_lancamentos.gerar_diario(res)
    partes = ['<div class="col2">']
    partes.append(status_html(diario.partidas_ok, "🟢 Partidas dobradas OK (débitos = créditos)",
                              "🔴 Partidas dobradas NÃO fecham"))
    partes.append(status_html(diario.cobertura_ok, "🟢 Radar de cobertura: diário cobre 100% dos desvios",
                              "⚠️ Cobertura incompleta — há desvio não lançado"))
    partes.append("</div>")
    for a in diario.alertas:
        partes.append(alerta_html(a, "warning"))
    df = diario.lancamentos.copy()
    for col in ("debito", "credito", "J_desvio"):
        df[col] = df[col].map(fmt_brl)
    partes.append(tabela_html(df))
    partes.append(f"<p class='caption'>Totais — Débito: {fmt_brl(diario.total_debito)} | "
                  f"Crédito: {fmt_brl(diario.total_credito)}</p>")
    return "".join(partes)


# --- Montagem final -----------------------------------------------------------

def _logo_data_uri(caminho):
    p = Path(caminho)
    if not p.exists():
        return None
    b64 = base64.b64encode(p.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{b64}"


CSS = """
:root { color-scheme: dark; }
* { box-sizing: border-box; }
body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 0 0 60px;
       background: #0f172a; color: #e2e8f0; }
header { display: flex; align-items: center; gap: 20px; padding: 24px 40px 12px; }
header img { height: 56px; }
header h1 { margin: 0; font-size: 1.5rem; }
header p { margin: 2px 0 0; color: #94a3b8; font-size: .9rem; }
.gerado-em { color: #64748b; font-size: .78rem; padding: 0 40px 12px; }
.tabs-wrapper { padding: 0 40px; }
.tabs-wrapper input[type=radio] { display: none; }
.tab-bar { display: flex; flex-wrap: wrap; gap: 6px; border-bottom: 2px solid #1e293b; margin-bottom: 24px; }
.tab-bar label { cursor: pointer; padding: 10px 16px; border-radius: 8px 8px 0 0; font-size: .88rem;
                 background: #1e293b; color: #94a3b8; }
.tab-panel { display: none; }
#t1:checked ~ .tabs-wrapper #p1, #t2:checked ~ .tabs-wrapper #p2, #t3:checked ~ .tabs-wrapper #p3,
#t4:checked ~ .tabs-wrapper #p4, #t5:checked ~ .tabs-wrapper #p5, #t6:checked ~ .tabs-wrapper #p6,
#t7:checked ~ .tabs-wrapper #p7, #t8:checked ~ .tabs-wrapper #p8, #t9:checked ~ .tabs-wrapper #p9 { display: block; }
#t1:checked ~ .tabs-wrapper .tab-bar label[for=t1], #t2:checked ~ .tabs-wrapper .tab-bar label[for=t2],
#t3:checked ~ .tabs-wrapper .tab-bar label[for=t3], #t4:checked ~ .tabs-wrapper .tab-bar label[for=t4],
#t5:checked ~ .tabs-wrapper .tab-bar label[for=t5], #t6:checked ~ .tabs-wrapper .tab-bar label[for=t6],
#t7:checked ~ .tabs-wrapper .tab-bar label[for=t7], #t8:checked ~ .tabs-wrapper .tab-bar label[for=t8],
#t9:checked ~ .tabs-wrapper .tab-bar label[for=t9] {
  background: #FF4F0D; color: white; font-weight: 600;
}
h3 { color: #f8fafc; }
h4 { color: #f1f5f9; margin-top: 28px; }
h5 { margin: 0 0 12px; }
.caption { color: #94a3b8; font-size: .85rem; }
.cards { display: flex; gap: 16px; margin: 16px 0 24px; flex-wrap: wrap; }
.metric { background: #1e293b; border-radius: 10px; padding: 14px 18px; flex: 1; min-width: 200px; }
.metric-label { color: #94a3b8; font-size: .8rem; }
.metric-value { font-size: 1.3rem; font-weight: 700; margin-top: 4px; }
.metric-delta { font-size: .85rem; margin-top: 2px; }
.col2 { display: flex; gap: 24px; margin-bottom: 24px; flex-wrap: wrap; }
.col2 > div { flex: 1; min-width: 320px; }
.alerta { padding: 12px 16px; border-radius: 8px; margin-bottom: 10px; font-weight: 600; }
.alerta-success { background: #052e1a; color: #4ade80; }
.alerta-error { background: #3f0d0d; color: #f87171; }
.alerta-warning { background: #3f2d0d; color: #fbbf24; font-weight: 400; }
.tabela-wrap { overflow-x: auto; margin-bottom: 12px; }
table.tabela { width: 100%; border-collapse: collapse; font-size: .85rem; }
table.tabela th, table.tabela td { padding: 7px 12px; border-bottom: 1px solid #1e293b; text-align: left; white-space: nowrap; }
table.tabela th { background: #1e293b; color: #f1f5f9; position: sticky; top: 0; }
table.tabela tr.destaque td { background: #1e293b; color: #f8fafc; font-weight: 600; }
details { margin-bottom: 8px; background: #111827; border-radius: 8px; padding: 6px 14px; }
details summary { cursor: pointer; padding: 8px 0; font-size: .9rem; }
.cartao-titulo { background: linear-gradient(135deg,#0f172a,#1e293b); padding: 16px 22px; border-radius: 12px; margin-bottom: 14px; }
.cartao-titulo .empresa { font-size: .7rem; opacity: .6; letter-spacing: .06em; text-transform: uppercase; }
.cartao-titulo .titulo { font-size: 1.15rem; font-weight: 700; margin-top: 2px; }
.cartao-titulo .nota { font-size: .78rem; opacity: .7; margin-top: 4px; }
.badge { background: rgba(255,255,255,.12); border-radius: 6px; padding: 3px 10px; margin-right: 8px; font-size: .78rem; display: inline-block; margin-top: 8px; }
.container-borda { border: 1px solid #1e293b; border-radius: 10px; padding: 16px 20px; margin-bottom: 16px; }
"""


def montar_html(res, dre_categoria):
    logo_uri = _logo_data_uri(CAMINHO_LOGO)
    logo_img = f'<img src="{logo_uri}" alt="Dotz">' if logo_uri else ""

    abas = [
        ("t1", "📊 Visão Geral / Conciliação", aba_visao_geral_html(res, dre_categoria)),
        ("t2", "📉 Passivo Diferido", aba_passivo_diferido_html(res)),
        ("t3", "💰 Receita Diferida por Categoria", aba_receita_categoria_html(dre_categoria)),
        ("t4", "🧾 Faturamento de Pontos", aba_faturamento_html()),
        ("t5", "🔧 Controle U1", aba_controle_u1_html()),
        ("t6", "U1.4 Expiração", aba_serie_atuarial_html(config.ABA_U1_4_EXPIRACAO, dre_categoria, col_ancora=1, categoria_dre="Breakage")),
        ("t7", "U1.5 Margem", aba_serie_atuarial_html(config.ABA_U1_5_MARGEM, dre_categoria, col_ancora=1, categoria_dre="Spread")),
        ("t8", "U1.6 Emissão/Resgates", aba_u1_6_html()),
        ("t9", "🧾 Lançamentos de Ajuste", aba_lancamentos_html(res)),
    ]

    radios = "".join(f'<input type="radio" name="tabs" id="{tid}"{" checked" if tid == "t1" else ""}>' for tid, _, _ in abas)
    tab_bar = "".join(f'<label for="{tid}">{titulo}</label>' for tid, titulo, _ in abas)
    paineis = "".join(f'<div class="tab-panel" id="p{tid[1:]}">{conteudo}</div>' for tid, _, conteudo in abas)

    agora = datetime.now().strftime("%d/%m/%Y %H:%M")

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>Passivo & Receita Diferida — IFRS 15 / CPC 47</title>
<style>{CSS}</style>
</head>
<body>
{radios}
<header>
  {logo_img}
  <div>
    <h1>Passivo & Receita Diferida — IFRS 15 / CPC 47</h1>
    <p>Conciliação de pontos: emissão, resgate e breakage vs balancete SAP.</p>
  </div>
</header>
<div class="gerado-em">Snapshot estático gerado em {agora} a partir de dados.xlsx — não interativo, sem filtros.</div>
<div class="tabs-wrapper">
  <div class="tab-bar">{tab_bar}</div>
  <div class="tabs-wrapper2">{paineis}</div>
</div>
</body>
</html>"""


def main():
    res = data_pipeline.executar_pipeline()
    dre_categoria = data_loader.load_dre_classificacao_receita()
    html = montar_html(res, dre_categoria)
    with open(ARQUIVO_SAIDA, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Arquivo gerado: {ARQUIVO_SAIDA}")


if __name__ == "__main__":
    main()
