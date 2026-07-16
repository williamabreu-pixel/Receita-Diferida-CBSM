# Sistema CBSM — Passivo & Receita Diferida (IFRS 15 / CPC 47)

Dashboard Streamlit para conciliação de pontos Dotz (emissão, resgate, breakage) entre o
balancete SAP e a planilha de controle interno (U1), com geração de lançamentos de ajuste.

Ver `CONTEXTO.md` para o briefing completo de arquitetura e regras de negócio.

## ⚠️ Dados

Este repositório **não inclui `dados.xlsx`** (a planilha-fonte, com números financeiros
reais da CBSM). Para rodar o app localmente, coloque seu próprio arquivo em
`data_raw/dados.xlsx` — ele precisa ter as abas listadas em `data_raw/config.py`.

## Como rodar localmente

```powershell
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
cd data_raw
..\.venv\Scripts\streamlit run app.py
```

> No Windows, se a pasta do projeto estiver num caminho muito longo (ex.: dentro do
> OneDrive), crie o venv num caminho curto (ex.: `%USERPROFILE%\venvs\sistema_cbsm`) —
> veja a seção "Ambiente técnico" do `CONTEXTO.md`.

## Estrutura

- `data_raw/config.py` — caminho do Excel e nomes das abas.
- `data_raw/data_loader.py` — ingestão (RAW) do Excel.
- `data_raw/data_pipeline.py` — camada de negócio: conciliação, DRE, blocos.
- `data_raw/gerador_lancamentos.py` — lançamentos de ajuste (partida dobrada).
- `data_raw/app.py` — dashboard Streamlit.
