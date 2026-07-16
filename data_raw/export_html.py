# export_html.py
"""Gera um snapshot estático em HTML do dashboard (gráficos + tabelas),
para abrir direto no navegador sem precisar do Streamlit rodando."""

from __future__ import annotations

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio

import data_pipeline
import gerador_lancamentos

ARQUIVO_SAIDA = "../dashboard.html"

CORES_MOTOR = {"Trocas": "#059669", "Breakage": "#d97706",
               "Spread": "#2563eb", "Promodotz": "#7c3aed"}


def fmt_brl(valor):
    if valor is None or pd.isna(valor):
        return "—"
    s = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def fig_receita_barras(res):
    agg = res.receita_por_motor
    fig = px.bar(agg, x="competencia", y="valor", color="motor",
                 color_discrete_map=CORES_MOTOR, barmode="stack",
                 labels={"valor": "Valor (R$)", "competencia": "Competência", "motor": "Motor"},
                 title="Receita por Motor")
    fig.update_layout(height=430, legend=dict(orientation="h", y=1.12), margin=dict(t=60, b=20))
    return fig


def fig_passivo_rosca(res):
    blocos = res.blocos_controle
    fig = go.Figure(go.Pie(labels=blocos["bloco"], values=blocos["passivo"], hole=0.55,
                           textinfo="label+percent",
                           marker=dict(colors=["#d97706", "#0891b2", "#2563eb", "#7c3aed"])))
    fig.update_layout(height=430, margin=dict(t=60, b=20), title="Composição do Passivo",
                      annotations=[dict(text="Passivo", x=0.5, y=0.5, showarrow=False, font_size=16)])
    return fig


def fig_dre(res):
    dre = res.resumo_dre
    fig = px.area(dre, x="competencia", y="receita_reconhecida",
                  labels={"receita_reconhecida": "Receita (R$)", "competencia": "Competência"},
                  title="DRE — Receita Reconhecida por Competência")
    fig.update_traces(line_color="#059669", fillcolor="rgba(5,150,105,0.15)")
    fig.update_layout(height=360, margin=dict(t=60, b=20))
    return fig


def tabela_html(df, moeda_cols=()):
    df = df.copy()
    for col in moeda_cols:
        df[col] = df[col].map(fmt_brl)
    return df.to_html(index=False, border=0, classes="tabela", escape=False)


def montar_html(res, diario):
    delta = res.passivo_total_controle - res.passivo_total_contabil
    status_geral = "🟢 Conciliação 100% OK" if res.conciliacao_ok else "🔴 Conciliação com divergência(s)"
    alertas_html = "".join(f"<li>{a}</li>" for a in res.alertas) or "<li>Nenhum alerta.</li>"
    alertas_diario_html = "".join(f"<li>{a}</li>" for a in diario.alertas) or "<li>Nenhum alerta.</li>"

    conc_html = tabela_html(res.conciliacao, moeda_cols=("controle", "contabil", "delta"))
    dre_html = tabela_html(res.resumo_dre, moeda_cols=("receita_reconhecida",))
    lanc_html = tabela_html(diario.lancamentos, moeda_cols=("debito", "credito", "J_desvio"))

    fig_receita = pio.to_html(fig_receita_barras(res), include_plotlyjs="cdn", full_html=False)
    fig_passivo = pio.to_html(fig_passivo_rosca(res), include_plotlyjs=False, full_html=False)
    fig_dre_html = pio.to_html(fig_dre(res), include_plotlyjs=False, full_html=False)

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>Passivo & Receita Diferida — IFRS 15 / CPC 47</title>
<style>
  body {{ font-family: Segoe UI, Arial, sans-serif; margin: 0; padding: 24px 40px; background: #f8fafc; color: #0f172a; }}
  h1 {{ margin-bottom: 4px; }}
  .caption {{ color: #64748b; margin-bottom: 24px; }}
  .status {{ font-size: 1.1rem; font-weight: 600; padding: 12px 16px; border-radius: 8px;
             background: {"#dcfce7" if res.conciliacao_ok else "#fee2e2"};
             color: {"#166534" if res.conciliacao_ok else "#991b1b"}; margin-bottom: 20px; }}
  .cards {{ display: flex; gap: 16px; margin-bottom: 28px; }}
  .card {{ background: white; border-radius: 10px; padding: 16px 20px; flex: 1;
           box-shadow: 0 1px 3px rgba(0,0,0,.08); }}
  .card .label {{ color: #64748b; font-size: .85rem; }}
  .card .value {{ font-size: 1.4rem; font-weight: 700; margin-top: 4px; }}
  .row {{ display: flex; gap: 24px; margin-bottom: 28px; }}
  .row > div {{ flex: 1; background: white; border-radius: 10px; padding: 8px; box-shadow: 0 1px 3px rgba(0,0,0,.08); }}
  section {{ background: white; border-radius: 10px; padding: 20px; margin-bottom: 28px; box-shadow: 0 1px 3px rgba(0,0,0,.08); }}
  h2 {{ margin-top: 0; }}
  table.tabela {{ width: 100%; border-collapse: collapse; }}
  table.tabela th, table.tabela td {{ padding: 8px 12px; border-bottom: 1px solid #e2e8f0; text-align: left; font-size: .92rem; }}
  table.tabela th {{ background: #f1f5f9; }}
  ul.alertas {{ margin: 0; padding-left: 20px; }}
  .footer {{ color: #94a3b8; font-size: .8rem; margin-top: 40px; }}
</style>
</head>
<body>
  <h1>📊 Passivo & Receita Diferida — IFRS 15 / CPC 47</h1>
  <p class="caption">Conciliação de pontos: emissão, resgate e breakage vs balancete SAP. Snapshot estático gerado a partir de dados.xlsx.</p>

  <div class="status">{status_geral}</div>

  <div class="cards">
    <div class="card"><div class="label">Passivo Contábil (Balancete)</div><div class="value">{fmt_brl(res.passivo_total_contabil)}</div></div>
    <div class="card"><div class="label">Passivo Controle (U1)</div><div class="value">{fmt_brl(res.passivo_total_controle)}</div></div>
    <div class="card"><div class="label">Delta (Controle − Contábil)</div><div class="value">{fmt_brl(delta)}</div></div>
  </div>

  <div class="row">
    <div>{fig_receita}</div>
    <div>{fig_passivo}</div>
  </div>

  <section>
    <h2>Conciliação — Quadros A / G / Receita por bloco</h2>
    {conc_html}
    <h3>Alertas</h3>
    <ul class="alertas">{alertas_html}</ul>
  </section>

  <section>
    {fig_dre_html}
  </section>

  <section>
    <h2>Lançamentos de Ajuste de Conciliação</h2>
    <p>Partidas dobradas: {"🟢 OK (débitos = créditos)" if diario.partidas_ok else "🔴 NÃO fecham"} —
       Cobertura: {"🟢 100% dos desvios" if diario.cobertura_ok else "⚠️ incompleta"}</p>
    <ul class="alertas">{alertas_diario_html}</ul>
    {lanc_html}
    <p>Totais — Débito: {fmt_brl(diario.total_debito)} | Crédito: {fmt_brl(diario.total_credito)}</p>
  </section>

  <div class="footer">Gerado por export_html.py</div>
</body>
</html>"""


def main():
    res = data_pipeline.executar_pipeline()
    diario = gerador_lancamentos.gerar_diario(res)
    html = montar_html(res, diario)
    with open(ARQUIVO_SAIDA, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Arquivo gerado: {ARQUIVO_SAIDA}")


if __name__ == "__main__":
    main()
