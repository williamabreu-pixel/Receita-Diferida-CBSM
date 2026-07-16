# data_loader.py
"""
Camada de Ingestão (RAW). Lê balancete, Base Receita Motor e controle U1
do MESMO arquivo Excel (a golden), sem depender de exports separados.
Mantém tipagem robusta: contas como texto, valores como float.
"""

from __future__ import annotations

import logging
import re

import pandas as pd

import config

logger = logging.getLogger(__name__)

BALANCETE_COLMAP = {
    "Cta.contáb./cód.PN": "conta",
    "Nome": "conta_nome",
    "Moeda corrente - Real - SI": "saldo_inicial",
    "Moeda corrente - Real - Saldo": "saldo_atual",
}
RECEITA_COLMAP = {
    "Competência": "competencia", "Motor": "motor", "Conta/Origem": "conta",
    "Descrição (TB)": "descricao_tb", "Canal": "canal", "Vintage": "vintage",
    "Valor": "valor", "Visível": "visivel",
}

U1_COLS = {
    "conta_cp": 1, "conta_lp": 2, "conta_resultado": 3, "processo": 4,
    "descricao": 5, "passivo_cp": 6, "passivo_lp": 8, "passivo_total": 10,
    "check": 11, "passivo_recalculado": 12, "receita": 13,
}
U1_BLOCOS = {
    "Expiração Parceiros": "Breakage",
    "Custo do Produto - Resgate": "Custo do Produto",
    "Margem dos Produtos": "Spread",
    "Promo Dotz": "Promodotz",
}
U1_LINHA_TOTAL = "Subtotal Non Split fee"

# Faixas de linha (planilha U1 - Movimentação, fixa e estável): grupo de
# negócio -> (linha de cabeçalho com o total do grupo; faixa de linhas de
# detalhe por conta). Nomes = texto literal da coluna 'Descrição' do
# cabeçalho de cada grupo na planilha (confirmado com o usuário).
U1_GRUPOS_FAIXAS = [
    ("Receita De Expiração De Pontos - Non Split Fee", 9, range(10, 17)),
    ("Receita De Resgate De Dotz - Non Split Fee", 18, range(19, 25)),
    ("Provisão Para Dotz Promocionais - LP", 31, range(31, 35)),
    ("Receita Venda Dotz - Non Split Fee", 36, range(37, 51)),
]

# Linhas (0-indexed) somadas pela própria fórmula da U1 para o split CP/LP
# (Excel: G60=G8+G10+G19+G37, I60=I8+I10+I19+I37 -> aqui, linhas 7,9,18,36).
# Confirmado lendo a fórmula real via openpyxl — bate ao centavo com
# 'U1 - Movimentação'!G60/I60 (a mesma fonte que a aba Dados_Gráfico usa,
# antes dos ajustes de Netpoints/Dotz Pay, que são de fora do escopo CBSM).
# Promodotz fica de fora dessa soma — mesmo comportamento da fórmula original.
U1_CP_LP_LINHAS = {"Subtotal Split Fee": 7, "Breakage": 9,
                   "Custo do Produto": 18, "Spread": 36}


def _clean_account_code(series):
    def _fmt(x):
        if pd.isna(x):
            return pd.NA
        if isinstance(x, float) and x.is_integer():
            return str(int(x))
        return str(x).strip()
    return series.map(_fmt).astype("string")


def _to_float(series):
    if pd.api.types.is_numeric_dtype(series):
        return series.astype("float64")
    s = series.astype("string").str.strip().str.replace(r"[^\d,.\-]", "", regex=True)
    tem_virgula = s.str.contains(",", na=False)
    br = s.where(tem_virgula).str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    us = s.where(~tem_virgula)
    return pd.to_numeric(br.fillna(us), errors="coerce").astype("float64")


def _to_bool(series):
    mapa = {"sim": True, "s": True, "true": True, "1": True, "verdadeiro": True,
            "não": False, "nao": False, "n": False, "false": False, "0": False, "falso": False}
    return series.astype("string").str.strip().str.lower().map(mapa).astype("boolean")


def _normalize_headers(df):
    df.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in df.columns]
    return df


def _require_columns(df, expected, source):
    faltando = [c for c in expected if c not in df.columns]
    if faltando:
        raise ValueError(f"[{source}] Colunas ausentes: {faltando}")


def load_balancete(competencia=None):
    with pd.ExcelFile(config.ARQUIVO_EXCEL) as xls:
        raw = pd.read_excel(xls, sheet_name=config.ABA_BALANCETE, header=0)
    raw = _normalize_headers(raw)
    _require_columns(raw, BALANCETE_COLMAP, "balancete")
    df = raw[list(BALANCETE_COLMAP)].rename(columns=BALANCETE_COLMAP).copy()
    df["conta"] = _clean_account_code(df["conta"])
    df["conta_nome"] = df["conta_nome"].astype("string").str.strip()
    df["saldo_inicial"] = _to_float(df["saldo_inicial"])
    df["saldo_atual"] = _to_float(df["saldo_atual"])
    df = df.dropna(subset=["conta"]).reset_index(drop=True)
    return df


def load_receita(competencia=None):
    with pd.ExcelFile(config.ARQUIVO_EXCEL) as xls:
        raw = pd.read_excel(xls, sheet_name=config.ABA_RECEITA, header=2)
    raw = _normalize_headers(raw)
    _require_columns(raw, RECEITA_COLMAP, "receita")
    df = raw[list(RECEITA_COLMAP)].rename(columns=RECEITA_COLMAP).copy()
    df["competencia"] = df["competencia"].astype("string").str.strip()
    df["motor"] = df["motor"].astype("string").str.strip()
    df["conta"] = _clean_account_code(df["conta"])
    df["descricao_tb"] = df["descricao_tb"].astype("string").str.strip()
    df["canal"] = df["canal"].astype("string").str.strip()
    df["vintage"] = df["vintage"].astype("string").str.strip()
    df["valor"] = _to_float(df["valor"])
    df["visivel"] = _to_bool(df["visivel"])
    df = df.dropna(subset=["motor", "valor"]).reset_index(drop=True)
    return df


def load_controle_u1(competencia=None):
    with pd.ExcelFile(config.ARQUIVO_EXCEL) as xls:
        raw = pd.read_excel(xls, sheet_name=config.ABA_CONTROLE, header=None)

    def col(nome):
        return raw.iloc[:, U1_COLS[nome]]

    detalhe = pd.DataFrame({
        "conta_resultado": _clean_account_code(col("conta_resultado")),
        "processo": col("processo").astype("string").str.strip(),
        "descricao": col("descricao").astype("string").str.strip(),
        "passivo_cp": _to_float(col("passivo_cp")),
        "passivo_lp": _to_float(col("passivo_lp")),
        "passivo_total": _to_float(col("passivo_total")),
        "check": _to_float(col("check")),
        "receita": _to_float(col("receita")),
    })

    blocos_rows = []
    for proc, rotulo in U1_BLOCOS.items():
        sel = detalhe[detalhe["processo"] == proc]
        if sel.empty:
            logger.warning("Bloco U1 não encontrado: %s", proc)
            continue
        r = sel.iloc[0]
        blocos_rows.append({
            "bloco": rotulo, "processo": proc,
            "passivo": float(r["passivo_total"]),
            "receita": float(r["receita"]) if pd.notna(r["receita"]) else 0.0,
        })
    blocos = pd.DataFrame(blocos_rows)

    tot = detalhe[detalhe["descricao"].str.contains(U1_LINHA_TOTAL, na=False)]
    total_passivo = float(tot["passivo_total"].iloc[0]) if not tot.empty else float("nan")
    blocos.attrs["total_passivo_controle"] = total_passivo

    linhas = detalhe[detalhe["conta_resultado"].notna()].reset_index(drop=True)
    return blocos, linhas


def load_controle_u1_grupos():
    """Reorganiza a U1 - Movimentação nos 4 grupos de negócio da planilha
    (nomeados com o texto literal da coluna 'Descrição' de cada cabeçalho).
    Cada grupo traz o total conciliado (Passivo Total × Passivo Recalculado
    × Check × Receita) e as linhas de detalhe por conta contábil, para
    expansão. Faixas de linha fixas — planilha pequena e estável (ver
    U1_GRUPOS_FAIXAS)."""
    with pd.ExcelFile(config.ARQUIVO_EXCEL) as xls:
        raw = pd.read_excel(xls, sheet_name=config.ABA_CONTROLE, header=None)

    def col(nome):
        return raw.iloc[:, U1_COLS[nome]]

    detalhe = pd.DataFrame({
        "conta": _clean_account_code(col("conta_resultado")),
        "descricao": col("descricao").astype("string").str.strip(),
        "descricao_processo": col("processo").astype("string").str.strip(),
        "passivo_cp": _to_float(col("passivo_cp")),
        "passivo_lp": _to_float(col("passivo_lp")),
        "passivo_total": _to_float(col("passivo_total")),
        "check": _to_float(col("check")),
        "passivo_recalculado": _to_float(col("passivo_recalculado")),
        "receita": _to_float(col("receita")),
    })
    detalhe["descricao"] = detalhe["descricao"].fillna(detalhe["descricao_processo"])

    grupos = []
    for nome, linha_cabecalho, linhas_detalhe in U1_GRUPOS_FAIXAS:
        detalhe_grupo = (detalhe.iloc[list(linhas_detalhe)]
                          .dropna(subset=["descricao"])
                          .drop(columns=["descricao_processo"])
                          .reset_index(drop=True))
        cab = detalhe.iloc[linha_cabecalho]
        passivo_total = float(cab["passivo_total"]) if pd.notna(cab["passivo_total"]) else 0.0
        passivo_recalculado = float(cab["passivo_recalculado"]) if pd.notna(cab["passivo_recalculado"]) else 0.0
        receita = float(cab["receita"]) if pd.notna(cab["receita"]) else 0.0
        check = float(cab["check"]) if pd.notna(cab["check"]) else 0.0
        grupos.append({
            "nome": nome,
            "passivo_total": passivo_total, "passivo_recalculado": passivo_recalculado,
            "check": check, "receita": receita, "detalhe": detalhe_grupo,
        })
    return grupos


def calcular_passivo_cp_lp():
    """Split Circulante (CP) × Não Circulante (LP) do Passivo Diferido,
    replicando a fórmula real da própria U1 - Movimentação (G60=G8+G10+G19
    +G37 e I60=I8+I10+I19+I37 — confirmado lendo a fórmula via openpyxl):
    soma passivo_cp/passivo_lp das linhas de cabeçalho de bloco (Subtotal
    Split Fee + Breakage + Custo do Produto + Spread). Bate ao centavo com
    'U1 - Movimentação'!G60/I60 — a MESMA fonte que a aba Dados_Gráfico lê
    (G4/G5, coluna 'CBSM'), sem os ajustes de Netpoints/Dotz Pay de lá
    (que são de fora do escopo deste sistema). Não depende da aba
    Dados_Gráfico nem de nenhuma projeção temporal."""
    with pd.ExcelFile(config.ARQUIVO_EXCEL) as xls:
        raw = pd.read_excel(xls, sheet_name=config.ABA_CONTROLE, header=None)
    cp_col, lp_col = U1_COLS["passivo_cp"], U1_COLS["passivo_lp"]
    cp_total, lp_total = 0.0, 0.0
    for idx in U1_CP_LP_LINHAS.values():
        cp_val = _to_float(pd.Series([raw.iloc[idx, cp_col]])).iloc[0]
        lp_val = _to_float(pd.Series([raw.iloc[idx, lp_col]])).iloc[0]
        cp_total += 0.0 if pd.isna(cp_val) else float(cp_val)
        lp_total += 0.0 if pd.isna(lp_val) else float(lp_val)
    return cp_total, lp_total


def load_aba_bruta(nome_aba, linhas_removidas_topo=0):
    """Lê uma aba do Excel sem remodelar: view crua para exibição estilo planilha.
    Remove apenas linhas e colunas 100% vazias, para facilitar a navegação."""
    with pd.ExcelFile(config.ARQUIVO_EXCEL) as xls:
        raw = pd.read_excel(xls, sheet_name=nome_aba, header=None,
                             skiprows=linhas_removidas_topo)
    raw = raw.dropna(axis=0, how="all").dropna(axis=1, how="all")
    # Colunas de tipo misto (texto + número na mesma coluna, comum em planilhas
    # com rótulos e dados na mesma posição) quebram a serialização Arrow do
    # Streamlit; convertendo para texto evitamos o erro sem alterar o conteúdo.
    for col in raw.columns:
        if raw[col].dtype == object and raw[col].map(type).nunique() > 1:
            raw[col] = raw[col].where(raw[col].isna(), raw[col].astype(str))
    return raw.reset_index(drop=True)


def _dividir_titulo_e_dados(raw, col_ancora, rotulo_ancora="Ano"):
    """Separa o bloco de título (linhas acima do 1º cabeçalho) do bloco de
    dados, removendo as repetições do cabeçalho que aparecem a cada bloco
    anual dentro da planilha original (paginação de impressão do Excel)."""
    col = raw.iloc[:, col_ancora].astype("string").str.strip().str.lower()
    idx_header = col[col == rotulo_ancora.lower()].index
    if idx_header.empty:
        return raw.iloc[:0].reset_index(drop=True), None, raw.reset_index(drop=True)
    primeiro = idx_header[0]
    titulo_bloco = raw.iloc[:primeiro].reset_index(drop=True)
    cabecalho = raw.iloc[primeiro]
    dados = raw.iloc[primeiro:].drop(index=idx_header).reset_index(drop=True)
    return titulo_bloco, cabecalho, dados


def load_serie_atuarial(nome_aba, col_ancora=1):
    """Lê as abas U1.5 (Margem) e U1.4 (Expiração): séries de apropriação
    por safra (Ano/Mês de emissão) com cabeçalho repetido a cada bloco anual.
    Devolve as linhas de título/parâmetros (texto puro, sem interpretar) e o
    bloco de dados já com os rótulos de coluna originais da planilha."""
    with pd.ExcelFile(config.ARQUIVO_EXCEL) as xls:
        raw = pd.read_excel(xls, sheet_name=nome_aba, header=None)
    titulo_bloco, cabecalho, dados = _dividir_titulo_e_dados(raw, col_ancora)
    if cabecalho is not None:
        rotulos, sem_rotulo_n = [], 0
        for c in cabecalho:
            if pd.isna(c):
                sem_rotulo_n += 1
                rotulos.append("(sem rótulo)" if sem_rotulo_n == 1 else f"(sem rótulo {sem_rotulo_n})")
            else:
                rotulos.append(str(c).replace("\n", " ").strip())
        dados.columns = rotulos

    # Mantém só linhas da série (Ano numérico) e os subtotais anuais ("Total");
    # descarta blocos de conferência soltos no fim da aba (fora da série Ano×Mês).
    coluna_ano = dados.iloc[:, col_ancora]
    ano_valido = (pd.to_numeric(coluna_ano, errors="coerce").notna()
                  | (coluna_ano.astype("string") == "Total"))
    dados = dados[ano_valido].reset_index(drop=True)
    dados = dados.dropna(axis=1, how="all").reset_index(drop=True)

    titulo_linhas = [str(v) for v in titulo_bloco.values.flatten() if pd.notna(v)]
    return titulo_linhas, dados


MESES_ORDEM = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
               "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]


def resumo_safra_por_competencia(dados, anos=range(2022, 2027)):
    """Pivô Safra (ano de emissão) × mês: usa as linhas 'Total' anuais já
    presentes na série (soma de todas as safras emitidas naquele ano, mês a
    mês, dentro do próprio ano de emissão) e soma uma linha 'Total' geral.
    Não inclui o residual de safras de anos anteriores ainda em amortização —
    cada linha reflete só a apropriação das safras DAQUELE ano de emissão."""
    ano_num = pd.to_numeric(dados.iloc[:, 0], errors="coerce")
    ano_grupo = ano_num.ffill()
    mask_total = dados.iloc[:, 0].astype("string") == "Total"

    totais = dados[mask_total].copy()
    totais["Safra"] = ano_grupo[mask_total].values
    totais = totais[totais["Safra"].isin(list(anos))].sort_values("Safra")
    totais["Safra"] = totais["Safra"].astype(int).astype(str)

    colunas = [m for m in MESES_ORDEM if m in dados.columns]
    if "Total" in dados.columns:
        colunas = colunas + ["Total"]
    resumo = totais[["Safra"] + colunas].reset_index(drop=True)

    linha_total = {"Safra": "Total"}
    for col in colunas:
        linha_total[col] = resumo[col].sum()
    return pd.concat([resumo, pd.DataFrame([linha_total])], ignore_index=True)


def load_u1_6_resgates():
    """Lê a aba U1.6 - Emissão_Resgates: série mensal de resgate de pontos
    (seção 'II - Receita com Resgate') e o quadro de conferência
    'Tie-in Contábil x Analyser'. Posições fixas — planilha pequena e estável."""
    with pd.ExcelFile(config.ARQUIVO_EXCEL) as xls:
        raw = pd.read_excel(xls, sheet_name=config.ABA_U1_6_RESGATES, header=None)

    mensal = pd.DataFrame({
        "ano": raw.iloc[11:23, 2].astype("Int64"),
        "mes": raw.iloc[11:23, 3].astype("Int64"),
        "pontos_resgatados": _to_float(raw.iloc[11:23, 4]),
        "preco_negociado": _to_float(raw.iloc[11:23, 5]),
        "valor_faturado": _to_float(raw.iloc[11:23, 6]),
    }).reset_index(drop=True)

    total_resgate = float(raw.iloc[24, 6]) if pd.notna(raw.iloc[24, 6]) else None

    tie_in = {}
    for r in range(28, 33):
        rotulo, valor = raw.iloc[r, 4], raw.iloc[r, 5]
        if isinstance(rotulo, str) and pd.notna(valor):
            tie_in[rotulo.strip()] = float(valor)

    return {"total_resgate": total_resgate, "tie_in": tie_in}, mensal


def load_dados_grafico_resumo():
    """Lê os dois quadros compactos e legíveis da aba Dados_Gráfico:
    'Consolidado' (prêmios/breakage a distribuir) e 'Receita Diferida CP/LP'.
    Valores na fonte em milhares de reais — convertidos aqui para reais
    (× 1000), mesmo tratamento de load_dados_grafico_expectativa_realizacao.
    O restante da aba (dados-fonte de outros gráficos) fica disponível só
    na visão bruta (load_aba_bruta), pois não tem estrutura tabular clara."""
    with pd.ExcelFile(config.ARQUIVO_EXCEL) as xls:
        raw = pd.read_excel(xls, sheet_name=config.ABA_DADOS_GRAFICO, header=None)

    consolidado = pd.DataFrame({
        "item": [raw.iloc[3, 0], raw.iloc[4, 0], raw.iloc[5, 0], raw.iloc[7, 0]],
        "em_30_06_2026": _to_float(pd.Series([raw.iloc[3, 1], raw.iloc[4, 1], raw.iloc[5, 1], raw.iloc[7, 1]])) * 1000,
        "em_31_12_2025": _to_float(pd.Series([raw.iloc[3, 2], raw.iloc[4, 2], raw.iloc[5, 2], raw.iloc[7, 2]])) * 1000,
    })

    cp_lp = pd.DataFrame({
        "categoria": [raw.iloc[3, 4], raw.iloc[4, 4], raw.iloc[5, 4]],
        "total": _to_float(pd.Series([raw.iloc[3, 5], raw.iloc[4, 5], raw.iloc[5, 5]])) * 1000,
        "cbsm": _to_float(pd.Series([raw.iloc[3, 6], raw.iloc[4, 6], raw.iloc[5, 6]])) * 1000,
        "netpoints": _to_float(pd.Series([raw.iloc[3, 7], raw.iloc[4, 7], raw.iloc[5, 7]])) * 1000,
        "dotz_pay": _to_float(pd.Series([raw.iloc[3, 8], raw.iloc[4, 8], raw.iloc[5, 8]])) * 1000,
        "total_ajustado": _to_float(pd.Series([raw.iloc[3, 9], raw.iloc[4, 9], raw.iloc[5, 9]])) * 1000,
    })

    return consolidado, cp_lp


def load_faturamento_vendas():
    """Lê a aba 'Fat. Análise de Vendas 2026': lançamentos de faturamento por
    mês × categoria (MKT, NSF, SF, Viaja Dotz, BV, Comissão etc.). O cabeçalho
    'Mês/Ano/Data Emissão/...' se repete a cada bloco mensal — removido aqui,
    mantendo só linhas de lançamento (Mês numérico)."""
    with pd.ExcelFile(config.ARQUIVO_EXCEL) as xls:
        raw = pd.read_excel(xls, sheet_name=config.ABA_FATURAMENTO_VENDAS,
                             header=None, usecols=range(9))

    col_mes = raw.iloc[:, 0].astype("string").str.strip()
    dados = raw[col_mes != "Mês"].copy()
    dados = dados[pd.to_numeric(dados.iloc[:, 0], errors="coerce").notna()].reset_index(drop=True)

    df = pd.DataFrame({
        "mes": dados.iloc[:, 0].astype(int),
        "ano": dados.iloc[:, 1].astype(int),
        "data_emissao": pd.to_datetime(dados.iloc[:, 2], errors="coerce"),
        "categoria": dados.iloc[:, 3].astype("string").str.strip(),
        "item": dados.iloc[:, 4].astype("string").str.strip(),
        "descricao": dados.iloc[:, 5].astype("string").str.strip(),
        "quantidade": _to_float(dados.iloc[:, 6]),
        "valor_vendas": _to_float(dados.iloc[:, 7]),
        "lucro_bruto": _to_float(dados.iloc[:, 8]),
    })
    df["competencia"] = df["ano"].astype(str) + "-" + df["mes"].astype(str).str.zfill(2)
    df = df.dropna(subset=["valor_vendas"]).reset_index(drop=True)
    return df


MESES_DRE = {"Janeiro": "01", "Fevereiro": "02", "Março": "03", "Abril": "04",
             "Maio": "05", "Junho": "06", "Julho": "07", "Agosto": "08",
             "Setembro": "09", "Outubro": "10", "Novembro": "11", "Dezembro": "12"}
CATEGORIA_DRE = {"Breakage": "Breakage", "Spread": "Spread", "Receita de Trocas": "Trocas",
                  "Projetos Especiais": "Receita Projetos Especiais"}
CONTA_PROJETOS_ESPECIAIS_DRE = "3111160"


def load_dre_classificacao_receita():
    """Lê a aba DRE 2026 linha a linha (por conta × competência), usando a
    coluna 'Classificação Receita' (fonte oficial para receita reconhecida
    no resultado, por decisão do usuário — substitui a Base Receita Motor
    nos gráficos de 'Receita por Categoria'). Formato "detalhe", análogo ao
    load_receita(): quem for agregar decide o agrupamento.

    Regras de categorização (confirmadas com o usuário):
    - Breakage / Spread / Receita de Trocas: usa a própria classificação
      (rótulo 'Receita de Trocas' encurtado para 'Trocas', mesmo nome já
      usado no resto do app — mas é um valor MENOR e DIFERENTE do "Trocas"
      da conciliação/Custo do Produto, pois a base é outra).
    - Conta 3111160 "Receita de Resgate de Dotz - Split Fee e Projetos
      Especiais": separada como 'Receita Projetos Especiais' — o restante
      de 'Receita de Trocas' continua classificado como 'Trocas'.
    - Promodotz (conta 3211010, "(-) Custo de Troca Promodotz") é redutor de
      custo, não receita — excluído destes gráficos de resultado.
    - Demais linhas de 'Receita de Trocas Cost' e todas as de 'Não Receita':
      excluídas — são custo/dedução ou não são receita, não entram no
      gráfico de receita reconhecida.
    Sinal: a fonte vem em convenção SAP (crédito negativo); invertido aqui
    para sentido de negócio (positivo), igual ao resto do app."""
    with pd.ExcelFile(config.ARQUIVO_EXCEL) as xls:
        raw = pd.read_excel(xls, sheet_name=config.ABA_DRE, header=0)
    raw = _normalize_headers(raw)

    conta = raw["Name"].astype("string").str.extract(r"^(\d{6,7})\s*-")[0]
    descricao = raw["Name"].astype("string").str.replace(r"^\d{6,7}\s*-\s*", "", regex=True)
    eh_projetos_especiais = (conta == CONTA_PROJETOS_ESPECIAIS_DRE).fillna(False)
    categoria = raw["Classificação Receita"].astype("string").map(CATEGORIA_DRE)
    categoria = categoria.mask(eh_projetos_especiais, "Receita Projetos Especiais")
    raw = raw.assign(categoria=categoria, conta=conta, descricao=descricao)
    dados = raw[raw["categoria"].notna()].copy()

    partes = []
    for mes_nome, mes_num in MESES_DRE.items():
        if mes_nome not in dados.columns:
            continue
        valores = _to_float(dados[mes_nome]) * -1.0
        partes.append(pd.DataFrame({
            "competencia": f"2026-{mes_num}",
            "categoria": dados["categoria"].to_numpy(),
            "conta": dados["conta"].to_numpy(),
            "descricao": dados["descricao"].to_numpy(),
            "valor": valores.to_numpy(),
        }))
    longo = pd.concat(partes, ignore_index=True)
    return longo.dropna(subset=["valor"]).reset_index(drop=True)


def load_dados_grafico_expectativa_realizacao():
    """Lê o quadro 'Expectativa de Realização da Receita' da aba
    Dados_Gráfico: passivo por vintage (ano de emissão) × ano de realização
    esperada. Valores na fonte em milhares de reais — convertidos para reais
    (× 1000) para consistência com o resto do app. Posições fixas (bloco
    pequeno e estável, com linha de total já pronta na planilha)."""
    with pd.ExcelFile(config.ARQUIVO_EXCEL) as xls:
        raw = pd.read_excel(xls, sheet_name=config.ABA_DADOS_GRAFICO, header=None)

    anos_realizacao = [int(raw.iloc[15, c]) for c in range(5, 10)]
    linhas = []
    for r in range(16, 21):
        linha = {"vintage": str(int(raw.iloc[r, 0])), "total": float(raw.iloc[r, 2]) * 1000}
        for i, ano in enumerate(anos_realizacao):
            linha[str(ano)] = float(raw.iloc[r, 5 + i]) * 1000
        linhas.append(linha)
    tabela = pd.DataFrame(linhas)

    linha_total = {"vintage": "Total", "total": float(raw.iloc[22, 2]) * 1000}
    for i, ano in enumerate(anos_realizacao):
        linha_total[str(ano)] = float(raw.iloc[22, 5 + i]) * 1000
    return pd.concat([tabela, pd.DataFrame([linha_total])], ignore_index=True)