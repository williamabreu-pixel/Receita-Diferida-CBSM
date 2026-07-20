# gerador_lancamentos.py
"""
Gerador de Lançamentos de Ajuste de Conciliação. Fonte do valor: o "check" de
cada grupo de negócio da aba Controle U1 (Passivo Total x Passivo
Recalculado) — o lançamento existe para que essa aba sempre feche, ao mesmo
tempo em que reconhece a receita do período. Monta partidas dobradas pela
regra: Débito=J se J>0; Crédito=-J se J<0, sempre pelo valor exato (sem
zerar diferenças pequenas).
"""

from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

import data_loader

TOLERANCIA = 1.0  # só para o arredondamento de ponto flutuante nos totais, não filtra valores.

# Nome literal do grupo (aba U1 - Movimentação, ver data_loader.U1_GRUPOS_FAIXAS) -> motor.
GRUPO_PARA_MOTOR = {
    "Receita De Expiração De Pontos - Non Split Fee": "Breakage",
    "Receita De Resgate De Dotz - Non Split Fee": "Custo do Produto",
    "Provisão Para Dotz Promocionais - LP": "Promodotz",
    "Receita Venda Dotz - Non Split Fee": "Spread",
}


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
    LinhaTemplate(2, "2311017", "Passivo Custo do Produto - Resgate",
                  "Reconhecimento de receita de resgate (custo do produto) do período", "Custo do Produto"),
    LinhaTemplate(2, "3111199", "Provisão de Receita",
                  "Contrapartida — provisão de receita (custo do produto)", "Custo do Produto"),
    LinhaTemplate(3, "3316100", "Promo Dotz",
                  "Ajuste Promodotz — contrapartida no resultado", "Promodotz"),
    LinhaTemplate(3, "2311020", "Dotz Promocionais - LP",
                  "Ajuste Promodotz — recomposição do passivo", "Promodotz"),
    LinhaTemplate(4, "2311011", "Passivo Margem dos Produtos",
                  "Reconhecimento de receita de margem (spread) do período", "Spread"),
    LinhaTemplate(4, "3111150", "Provisão de Receita",
                  "Contrapartida — provisão de receita (spread)", "Spread"),
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


def _desvio_por_motor():
    """Check de cada grupo de negócio da aba Controle U1 (Passivo Total menos
    Passivo Recalculado), reclassificado pelo nome do motor correspondente."""
    grupos = data_loader.load_controle_u1_grupos()
    desvios = {}
    for g in grupos:
        motor = GRUPO_PARA_MOTOR.get(g["nome"])
        if motor:
            desvios[motor] = float(g["check"])
    return desvios


def gerar_diario(res, competencia_fechamento=None):
    desvios = _desvio_por_motor()
    linhas = []
    ordem_no_lancamento = {}
    for t in TEMPLATE_LANCAMENTOS:
        j = desvios.get(t.bloco_fonte, 0.0)
        # Cada lançamento tem 2 linhas (conta principal + contrapartida). A 2ª
        # linha do par espelha o sinal da 1ª, para a partida dobrada fechar
        # (débito = crédito) — sem isso, as duas linhas caíam sempre do mesmo
        # lado e o total de débitos nunca fechava com o de créditos.
        ordem = ordem_no_lancamento.get(t.lancamento, 0)
        ordem_no_lancamento[t.lancamento] = ordem + 1
        j_linha = j if ordem == 0 else -j
        debito = j_linha if j_linha > 0 else 0.0
        credito = -j_linha if j_linha < 0 else 0.0
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

    motores_cobertos = set(t.bloco_fonte for t in TEMPLATE_LANCAMENTOS)
    desvio_motor = sum(abs(v) for v in desvios.values())
    desvio_coberto = sum(abs(v) for k, v in desvios.items() if k in motores_cobertos)
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
