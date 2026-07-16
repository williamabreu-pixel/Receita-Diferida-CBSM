# gerar_documentacao.py
"""Gera a documentação viva do projeto em Word (.docx) na raiz do projeto.
Rodar com:  python gerar_documentacao.py
Atualize o conteúdo abaixo conforme o projeto evoluir e rode de novo —
o arquivo é sobrescrito, preservando o histórico na seção 'Atualizações'."""

from __future__ import annotations

import datetime

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor

ARQUIVO_SAIDA = "../Documentacao_Sistema_CBSM.docx"

AZUL_CBSM = RGBColor(0x0F, 0x17, 0x2A)
CINZA = RGBColor(0x47, 0x55, 0x69)


def _titulo(doc, texto, nivel=1):
    h = doc.add_heading(texto, level=nivel)
    for run in h.runs:
        run.font.color.rgb = AZUL_CBSM
    return h


def _paragrafo(doc, texto, negrito=False, italico=False, tamanho=11, cor=None):
    p = doc.add_paragraph()
    run = p.add_run(texto)
    run.bold = negrito
    run.italic = italico
    run.font.size = Pt(tamanho)
    if cor:
        run.font.color.rgb = cor
    return p


def _bullet(doc, texto, negrito_prefixo=None):
    p = doc.add_paragraph(style="List Bullet")
    if negrito_prefixo:
        r = p.add_run(negrito_prefixo)
        r.bold = True
    p.add_run(texto)
    return p


def _tabela(doc, cabecalho, linhas):
    t = doc.add_table(rows=1, cols=len(cabecalho))
    t.style = "Light Grid Accent 1"
    for i, titulo_col in enumerate(cabecalho):
        cel = t.rows[0].cells[i]
        cel.text = titulo_col
        for p in cel.paragraphs:
            for r in p.runs:
                r.bold = True
    for linha in linhas:
        row = t.add_row()
        for i, valor in enumerate(linha):
            row.cells[i].text = str(valor)
    doc.add_paragraph("")
    return t


def gerar():
    doc = Document()
    estilo = doc.styles["Normal"]
    estilo.font.name = "Calibri"
    estilo.font.size = Pt(11)

    # Capa
    titulo = doc.add_heading("Sistema CBSM — Passivo & Receita Diferida", level=0)
    for run in titulo.runs:
        run.font.color.rgb = AZUL_CBSM
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = sub.add_run("Conciliação de pontos Dotz (IFRS 15 / CPC 47) — Documentação Viva do Projeto")
    r.italic = True
    r.font.size = Pt(13)
    r.font.color.rgb = CINZA
    _paragrafo(doc, f"Última atualização: {datetime.date.today().strftime('%d/%m/%Y')}",
               italico=True, tamanho=10, cor=CINZA)
    doc.add_paragraph("")

    # 1. Visão geral
    _titulo(doc, "1. O que o sistema faz")
    _paragrafo(doc, "Concilia o passivo e a receita diferida de pontos de fidelidade (Dotz), conforme "
                     "IFRS 15 / CPC 47, cruzando duas fontes:")
    _bullet(doc, "balancete do SAP (aba Trial Balance) — lado contábil.", "Lado contábil: ")
    _bullet(doc, "planilha interna U1 - Movimentação — lado controle, com passivo recalculado por conta.",
            "Lado controle: ")
    _paragrafo(doc, "A partir dessa conciliação, o sistema gera lançamentos contábeis de ajuste para "
                     "cobrir os desvios encontrados, e oferece um dashboard Streamlit navegável por abas "
                     "para explorar cada parte da planilha original.")

    # 2. Motores de receita
    _titulo(doc, "2. Motores de receita (4)")
    _tabela(doc, ["Motor", "O que é", "Tipo de diferimento"], [
        ["Breakage", "Expiração estimada de pontos que se tornam receita.",
         "Diferimento 48 meses (linear por competência de emissão)"],
        ["Spread", "Margem sobre a venda de pontos.",
         "Diferimento 48 meses (linear por competência de emissão)"],
        ["Trocas (Custo do Produto)", "Resgate de pontos + provisão manual \"a faturar\" (conta 3111199).",
         "Fato gerador (resgate)"],
        ["Promodotz", "Pontos promocionais.", "Fato gerador (evento)"],
    ])

    # 3. Arquitetura
    _titulo(doc, "3. Arquitetura")
    _paragrafo(doc, "Camadas do sistema (todas em data_raw/):", negrito=True)
    _tabela(doc, ["Arquivo", "Camada", "Responsabilidade"], [
        ["config.py", "Config", "Caminho do Excel (dados.xlsx) e nomes das abas usadas."],
        ["data_loader.py", "Ingestão (RAW)", "Lê balancete, Base Receita Motor, U1 e as abas de apoio "
         "(U1.4, U1.5, U1.6, Dados_Gráfico, Fat. Análise de Vendas). Tipagem robusta, sem alterar valores."],
        ["data_pipeline.py", "Negócio", "Agrega receita, monta a DRE, concilia por blocos "
         "(Quadros A/G), calcula passivo contábil × controle. Não altera o RAW."],
        ["gerador_lancamentos.py", "Negócio", "Gera lançamentos de ajuste (partida dobrada: "
         "Débito=J se J>0, Crédito=-J se J<0) para os desvios encontrados."],
        ["app.py", "Apresentação", "Dashboard Streamlit interativo — 10 abas (ver seção 5)."],
        ["export_html.py", "Apresentação (opcional)", "Snapshot HTML estático — só quando pedido "
         "explicitamente; o app Streamlit é o alvo real."],
        ["gerar_documentacao.py", "Documentação", "Gera este documento. Rodar de novo após "
         "mudanças relevantes no projeto."],
        ["gerar_logo_placeholder.py", "Identidade visual (opcional)", "Gera um ícone Dotz "
         "placeholder (círculo laranja + \"DZ\") via Pillow, enquanto o arquivo oficial da "
         "marca não é exportado — ver seção 5.1."],
    ])
    _paragrafo(doc, "Arquitetura híbrida: os números ATUAIS/de conciliação vêm do balancete e da U1 "
                     "(via data_loader). A PROJEÇÃO usa um motor atuarial de safras "
                     "(apropriação linear em 48 meses), com dados-fonte nas abas U1.4/U1.5.")

    # 4. Regras contábeis
    _titulo(doc, "4. Regras contábeis importantes")
    _bullet(doc, "No balancete SAP, passivo (grupo 2) e receita (grupo 3) vêm com crédito NEGATIVO. "
                  "O RAW preserva o sinal do SAP; a inversão para \"sentido de negócio\" (positivo) só "
                  "ocorre na camada de negócio (data_pipeline.py).")
    _bullet(doc, "A Base Receita Motor já vem positiva (sentido de negócio) — nunca inverter o sinal dela.")
    _bullet(doc, "Contas de passivo: grupos 219 (circulante) e 231 (não circulante). Usar só contas "
                  "ANALÍTICAS de 7 dígitos, para evitar dupla contagem pai/filho da hierarquia SAP.")
    _bullet(doc, "Split CP/LP não é derivável por prefixo de conta — vem da aba Dados_Gráfico "
                  "(quadro Receita Diferida CP/LP), já incorporado à aba Passivo Diferido do dashboard.")
    _bullet(doc, "_to_float tolera formato BR (1.234,56) e US/Excel (1234.56) — decide pela presença "
                  "de vírgula. Nunca remover o ponto decimal cegamente (já inflou valores 100x no passado).")
    _bullet(doc, "Contas contábeis são sempre TEXTO, nunca número (ex.: '3111160', não 3111160.0).")
    _bullet(doc, "Premissas contábeis (régua de breakage, CPD, provisão manual) são inputs explícitos "
                  "no Excel — nunca adivinhados pelo código.")

    # 5. Abas do dashboard
    _titulo(doc, "5. As 10 abas do dashboard (app.py)")
    abas = [
        ("📊 Visão Geral / Conciliação", "Alerta verde/vermelho, cards de passivo (contábil/controle/delta), "
         "gráfico de receita por categoria, rosca de Composição do Passivo Diferido, e a tabela de "
         "Conciliação (Total Passivo / Passivo Recalculado / Receita por bloco)."),
        ("💰 Receita Diferida por Categoria", "Filtros de competência e motor, gráfico de barras "
         "empilhadas, resumo por competência × motor, e detalhe expansível linha a linha com o "
         "tipo de diferimento de cada categoria."),
        ("📉 Passivo Diferido", "Composição do passivo (rosca + tabela por bloco) e o split "
         "Circulante (CP) × Não Circulante (LP), fonte: aba Dados_Gráfico, conferido contra o "
         "Passivo Total do Controle U1."),
        ("U1.5_Par._Emissão Margem", "Motor atuarial do Spread: quadro \"Receita a Reconhecer por "
         "Safra × Competência\" (2022-2026, com total) + gráfico de barras empilhadas, e a série "
         "completa por safra em ordem decrescente (2026 primeiro)."),
        ("U1.4_Par._Emissão Expiração", "Mesma estrutura da U1.5, para o motor de Breakage."),
        ("U1.6_Emissão_Resgates", "Série mensal de resgate de pontos e o quadro de conferência "
         "\"Tie-in Contábil × Analyser\"."),
        ("Dados_Gráfico", "Quadro principal \"Expectativa de Realização da Receita\" (vintage × "
         "ano de realização, 2022-2030) com gráfico de barras empilhadas, mais os quadros "
         "Consolidado e Receita Diferida CP/LP."),
        ("🧾 Faturamento de Pontos", "Resumo do faturamento do ano por categoria (valor de vendas "
         "e quantidade de pontos), com detalhe por competência × categoria recolhido e exportável "
         "em Excel."),
        ("🔧 Controle U1", "Conciliação por grupo de negócio: 4 cartões (nomeados com o texto "
         "literal da planilha) cruzando Passivo Total × Passivo Recalculado × Check × Receita, "
         "cada um expansível a nível de conta contábil."),
        ("🧾 Lançamentos de Ajuste", "Diário de ajuste gerado a partir dos desvios de conciliação, "
         "com checks de partida dobrada e cobertura, e exportação em Excel."),
    ]
    for nome, desc in abas:
        _bullet(doc, desc, f"{nome}: ")

    # 6. Identidade visual
    _titulo(doc, "6. Identidade visual (Design System Dotz)")
    _paragrafo(doc, "Paleta oficial da marca (fornecida pelo usuário), aplicada em app.py:", negrito=True)
    _tabela(doc, ["Cor", "Hex", "Uso no dashboard"], [
        ["Laranja (primária)", "#FF4F0D", "Breakage, status OK-adjacente em gráficos, DRE, "
         "rampa de safra/vintage, categoria NSF do faturamento."],
        ["Amarelo", "#FEC114", "Trocas (Custo do Produto), categoria SF do faturamento."],
        ["Verde", "#009F3C", "Spread, status \"OK\" da tabela de Conciliação, categoria BV."],
        ["Preto", "#000000", "Promodotz, categoria MKT do faturamento."],
        ["Magenta", "#D82598", "Categoria Viaja Dotz do faturamento."],
        ["Vermelho escuro", "#AF1010", "Status \"DIVERGENTE\", categoria Comissão com "
         "Conversão de Dotz."],
        ["Bege", "#E3D2C8", "Categoria Prestação de Serviços Comissão - Banco do Brasil."],
    ])
    _paragrafo(doc, "Composição do Passivo Diferido (rosca, abas Visão Geral e Passivo Diferido): "
                     "Promodotz permanece na rosca (não é removido) e recebe destaque visual "
                     "(fatia levemente separada — \"pull\") a pedido do usuário, mesmo sabendo que "
                     "Breakage + Custo do Produto + Spread já somam 100% do Passivo Diferido Total "
                     "sozinhos.")
    _paragrafo(doc, "Logo: o header (app.py) já está preparado para exibir logo_dotz.png / "
                     "icone_dotz.png (pasta data_raw) via st.logo + st.image, com fallback "
                     "silencioso caso os arquivos não existam. Como o arquivo oficial ainda não "
                     "foi exportado, gerar_logo_placeholder.py cria uma aproximação (círculo "
                     "laranja #FF4F0D + \"DZ\" em negrito) — trocar pelo arquivo oficial assim "
                     "que disponível, para fidelidade total à tipografia da marca.")

    # 7. Números validados
    _titulo(doc, "7. Números de referência validados (jun/2026 YTD)")
    _tabela(doc, ["Item", "Valor"], [
        ["Passivo Diferido Total", "R$ 196.321.706,82"],
        ["Receita Breakage", "R$ 21.610.402,09"],
        ["Receita Custo do Produto / Trocas", "R$ 22.665.859,03"],
        ["Receita Spread", "R$ 16.095.728,24"],
        ["Receita Promodotz", "R$ 1.075.793,82"],
        ["Passivo Circulante (CP)", "R$ 135.027.553,55"],
        ["Passivo Não Circulante (LP)", "R$ 61.294.153,27"],
        ["Tolerância de conciliação", "R$ 1,00"],
    ])

    # 8. Ambiente
    _titulo(doc, "8. Ambiente técnico")
    _bullet(doc, "A pasta do projeto está no OneDrive com caminho longo, o que estoura o limite de "
                  "~260 caracteres do Windows e quebra a instalação do Streamlit num .venv local.")
    _bullet(doc, "Solução: ambiente virtual em caminho curto — %USERPROFILE%\\venvs\\sistema_cbsm "
                  "— com pandas, openpyxl, streamlit, plotly, python-docx e Pillow instalados.")
    _bullet(doc, "Rodar o dashboard: abrir PowerShell na pasta data_raw e executar "
                  "\"%USERPROFILE%\\venvs\\sistema_cbsm\\Scripts\\streamlit.exe run app.py\".")
    _bullet(doc, "O arquivo de dados deve se chamar dados.xlsx e estar na mesma pasta dos .py "
                  "(data_raw). Fechar o Excel antes de eu (ou o app) precisar ler o arquivo.")

    # 9. Itens em aberto
    _titulo(doc, "9. Itens em aberto")
    _bullet(doc, "Segundo tipo de lançamento de ajuste (reclassificação CP × LP) em "
                  "gerador_lancamentos.py — agora que o split CP/LP está disponível na aba "
                  "Passivo Diferido, falta decidir se/como gerar o lançamento contábil correspondente.")
    _bullet(doc, "Substituir o logo placeholder (gerar_logo_placeholder.py) pelo arquivo oficial "
                  "da marca (logo_dotz.png / icone_dotz.png) assim que exportado — hoje é só uma "
                  "aproximação, sem a tipografia customizada do \"Z\" do logo real.")

    # 10. Histórico de atualizações
    _titulo(doc, "10. Histórico de atualizações")
    _paragrafo(doc, "Documentação viva — cada sessão de trabalho relevante deve somar uma entrada "
                     "aqui (nova entrada no topo, com data fixa) antes de regerar o arquivo. As datas "
                     "abaixo são fixas no texto — não usar datetime.date.today() para elas, senão "
                     "o histórico troca de data toda vez que o documento é regerado.",
               italico=True, cor=CINZA)
    _bullet(doc, "Identidade visual Dotz aplicada em app.py com a paleta OFICIAL da marca "
                  "(#FF4F0D laranja, #FEC114 amarelo, #009F3C verde, #000000 preto, #D82598 "
                  "magenta, #AF1010 vermelho, #E3D2C8 bege) — motores, roscas (Promodotz mantido "
                  "e destacado, não removido), DRE, status da Conciliação, rampa de safra/vintage "
                  "(U1.4/U1.5/Dados_Gráfico) e as 7 categorias do Faturamento de Pontos. Header "
                  "com logo Dotz preparado (fallback silencioso) e ícone placeholder gerado via "
                  "Pillow enquanto o arquivo oficial não é exportado.", "16/07/2026 — ")
    _bullet(doc, "Documento inicial gerado, consolidando todo o estado do projeto até aqui: "
                  "10 abas do dashboard, split CP/LP resolvido, aba Controle U1 redesenhada em "
                  "4 grupos de negócio, aba Faturamento de Pontos criada, exportações padronizadas "
                  "em Excel (XLSX).", "15/07/2026 — ")

    doc.save(ARQUIVO_SAIDA)
    print(f"Documento gerado: {ARQUIVO_SAIDA}")


if __name__ == "__main__":
    gerar()
