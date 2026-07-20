# data_pipeline.py
"""
Camada de Negócio. Consome o data_loader e monta a conciliação por blocos
(Quadros A–G), a receita por motor e a DRE. Não altera o RAW.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field

import pandas as pd

import data_loader

logger = logging.getLogger(__name__)

PREFIXOS_PASSIVO = ("219", "231")
_COMPETENCIA_RE = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")
TOLERANCIA = 1.0  # Passivo Diferido Total (quadros A/G) e check conta-a-conta da U1 — deve fechar exato.
TOLERANCIA_CONTA = 999.0  # Receita por conta/motor (quadros "Receita/*") — folga para ruído de arredondamento.


@dataclass
class ResultadoPipeline:
    receita_por_motor: pd.DataFrame
    blocos_controle: pd.DataFrame
    conciliacao: pd.DataFrame
    resumo_dre: pd.DataFrame
    passivo_total_contabil: float = 0.0
    passivo_total_controle: float = 0.0
    check_maximo: float = 0.0
    alertas: list = field(default_factory=list)
    conciliacao_ok: bool = True


def filtrar_competencias_validas(df):
    comp = df["competencia"].astype("string").str.strip()
    mask = comp.str.match(_COMPETENCIA_RE, na=False)
    out = df.loc[mask].copy()
    out["competencia"] = comp[mask]
    return out.reset_index(drop=True)


def agregar_receita_por_motor(df_receita):
    df = filtrar_competencias_validas(df_receita)
    return (df.groupby(["competencia", "motor"], as_index=False)
              .agg(valor=("valor", "sum"), linhas=("valor", "size"))
              .sort_values(["competencia", "motor"]).reset_index(drop=True))


def montar_dre(agg):
    return (agg.groupby("competencia", as_index=False)["valor"].sum()
              .rename(columns={"valor": "receita_reconhecida"})
              .sort_values("competencia").reset_index(drop=True))


def _passivo_contabil(df_balancete):
    s = df_balancete["conta"].astype("string").str.strip()
    m = s.str.startswith(PREFIXOS_PASSIVO, na=False) & (s.str.len() == 7)
    return -float(df_balancete.loc[m, "saldo_atual"].sum())


def conciliar(blocos, passivo_contabil, receita_agg):
    passivo_controle = float(blocos.attrs["total_passivo_controle"])
    principais = blocos[blocos["bloco"].isin(["Breakage", "Custo do Produto", "Spread"])]
    soma_decomposicao = float(principais["passivo"].sum())
    rec_motor = receita_agg.groupby("motor")["valor"].sum()
    mapa_motor = {"Breakage": "Breakage", "Spread": "Spread",
                  "Custo do Produto": "Trocas", "Promodotz": "Promodotz"}
    linhas = [
        {"quadro": "A", "item": "Passivo Diferido Total",
         "controle": passivo_controle, "contabil": passivo_contabil},
        {"quadro": "G", "item": "Decomposição (Breakage+Custo+Spread)",
         "controle": soma_decomposicao, "contabil": passivo_contabil},
    ]
    for _, b in blocos.iterrows():
        motor = mapa_motor.get(b["bloco"])
        rec_base = float(rec_motor.get(motor, 0.0)) if motor else 0.0
        linhas.append({
            "quadro": f"Receita/{b['bloco']}", "item": f"Receita {b['bloco']}",
            "controle": float(b["receita"]), "contabil": rec_base,
        })
    conc = pd.DataFrame(linhas)
    conc["delta"] = conc["controle"] - conc["contabil"]
    tolerancia_linha = conc["quadro"].str.startswith("Receita/").map(
        {True: TOLERANCIA_CONTA, False: TOLERANCIA})
    conc["status"] = (conc["delta"].abs() <= tolerancia_linha).map({True: "OK", False: "DIVERGENTE"})
    return conc


def executar_pipeline(competencia=None):
    df_bal = data_loader.load_balancete(competencia)
    df_rec = data_loader.load_receita(competencia)
    blocos, linhas_u1 = data_loader.load_controle_u1(competencia)

    agg = agregar_receita_por_motor(df_rec)
    dre = montar_dre(agg)
    passivo_contabil = _passivo_contabil(df_bal)
    passivo_controle = float(blocos.attrs["total_passivo_controle"])
    conc = conciliar(blocos, passivo_contabil, agg)
    check_max = float(linhas_u1["check"].abs().max())

    alertas = []
    divergentes = conc[conc["status"] == "DIVERGENTE"]
    for _, r in divergentes.iterrows():
        alertas.append(f"❌ Quadro {r['quadro']} — {r['item']}: "
                       f"controle R$ {r['controle']:,.2f} vs contábil "
                       f"R$ {r['contabil']:,.2f} (Δ R$ {r['delta']:,.2f}).")
    if check_max > TOLERANCIA_CONTA:
        alertas.append(f"⚠️ Check conta-a-conta da U1 excede tolerância: R$ {check_max:,.4f}.")

    conciliacao_ok = divergentes.empty and check_max <= TOLERANCIA_CONTA

    return ResultadoPipeline(
        receita_por_motor=agg, blocos_controle=blocos, conciliacao=conc,
        resumo_dre=dre, passivo_total_contabil=passivo_contabil,
        passivo_total_controle=passivo_controle, check_maximo=check_max,
        alertas=alertas, conciliacao_ok=conciliacao_ok,
    )


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s | %(message)s")
    res = executar_pipeline()
    print("\nPassivo contábil:", f"R$ {res.passivo_total_contabil:,.2f}")
    print("Passivo controle:", f"R$ {res.passivo_total_controle:,.2f}")
    print("\nConciliação:")
    print(res.conciliacao.to_string(index=False))
    print("\nStatus:", "OK" if res.conciliacao_ok else "DIVERGENTE")