# gerador_lancamentos.py
"""
Gerador de Lançamentos de Ajuste de Conciliação. Consome o ResultadoPipeline
e monta partidas dobradas pela regra: Débito=J se J>0; Crédito=-J se J<0.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

TOLERANCIA = 1.0


@dataclass
class LinhaTemplate:
    lancamento: int
    conta: str
    descricao_conta: str
    historico: str
    bloco_fonte: str


TEMPLATE_LANCAMENTOS = [
    LinhaTemplate(1, "2191013", "Passivo Expiração de Pontos - Non Split Fee",
                  "Reconhecimento de receita de expiração (breakage) do período", "Breakage"),
    LinhaTemplate(1, "3111450", "Receita de Expiração de Dotz - Non Split Fee",
                  "Contrapartida — receita de expiração de Dotz", "Breakage"),
    LinhaTemplate(3, "3316100", "Promo Dotz",
                  "Ajuste Promodotz — contrapartida no resultado", "Promodotz"),
    LinhaTemplate(3, "2311020", "Dotz Promocionais - LP",
                  "Ajuste Promodotz — recomposição do passivo", "Promodotz"),
]


@dataclass
class DiarioAjuste:
    lancamentos: pd.DataFrame
    total_debito: float = 0.0
    total_credito: float = 0.0
    partidas_ok: bool = True
    desvio_motor: float = 0.0
    desvio_coberto: float = 0.0
    cobertura_ok: bool = True
    alertas: list = field(default_factory=list)


def _desvio_por_bloco(res):
    conc = res.conciliacao
    desvios = {}
    for _, row in conc.iterrows():
        quadro = str(row["quadro"])
        if quadro.startswith("Receita/"):
            desvios[quadro.split("/", 1)[1]] = float(row["delta"])
    desvios["PassivoTotal"] = float(conc.loc[conc["quadro"] == "A", "delta"].sum())
    return desvios


def gerar_diario(res, competencia_fechamento=None):
    desvios = _desvio_por_bloco(res)
    linhas = []
    for t in TEMPLATE_LANCAMENTOS:
        j = desvios.get(t.bloco_fonte, 0.0)
        if abs(j) < TOLERANCIA:
            j = 0.0
        debito = j if j > 0 else 0.0
        credito = -j if j < 0 else 0.0
        linhas.append({
            "lancamento": t.lancamento, "conta": t.conta,
            "descricao_conta": t.descricao_conta, "historico": t.historico,
            "debito": round(debito, 2), "credito": round(credito, 2),
            "J_desvio": round(j, 2), "fonte": t.bloco_fonte,
        })
    df = pd.DataFrame(linhas)

    total_deb = float(df["debito"].sum())
    total_cred = float(df["credito"].sum())
    partidas_ok = abs(total_deb - total_cred) < TOLERANCIA
    desvio_motor = sum(abs(v) for k, v in desvios.items()
                       if k != "PassivoTotal" and abs(v) >= TOLERANCIA)
    desvio_coberto = float(df["J_desvio"].abs().sum())
    cobertura_ok = abs(desvio_motor - desvio_coberto) < TOLERANCIA

    alertas = []
    if not partidas_ok:
        alertas.append(f"❌ Partidas não fecham: débitos R$ {total_deb:,.2f} ≠ créditos R$ {total_cred:,.2f}.")
    if not cobertura_ok:
        alertas.append(f"⚠️ Cobertura incompleta: desvio motor R$ {desvio_motor:,.2f} vs coberto R$ {desvio_coberto:,.2f}.")

    return DiarioAjuste(
        lancamentos=df, total_debito=total_deb, total_credito=total_cred,
        partidas_ok=partidas_ok, desvio_motor=desvio_motor,
        desvio_coberto=desvio_coberto, cobertura_ok=cobertura_ok, alertas=alertas,
    )