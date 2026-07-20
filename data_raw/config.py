# config.py
"""Configuração central: caminho do Excel e nomes das abas."""

# Nome do arquivo Excel na mesma pasta. A partir de 20/07/2026 a base passou a
# ser atualizada por macros dentro do próprio arquivo — por isso a fonte
# oficial agora é o .xlsm (macro-enabled), não mais o .xlsx.
ARQUIVO_EXCEL = "dados.xlsm"

# Abas dentro do Excel (não mude, a menos que os nomes das abas mudem).
ABA_BALANCETE = "Trial Balance 2026"
ABA_RECEITA = "Base Receita Motor"
ABA_CONTROLE = "U1 - Movimentação"

# Abas extras exibidas como grade crua (estilo planilha) nas abas do dashboard.
ABA_U1_5_MARGEM = "U1.5_Par._Emissão Margem"
ABA_U1_4_EXPIRACAO = "U1.4_Par._Emissão Expiração"
ABA_U1_6_RESGATES = "U1.6_Emissão_Resgates"
ABA_DADOS_GRAFICO = "Dados_Gráfico"
ABA_FATURAMENTO_VENDAS = "Fat. Análise de Vendas 2026"
ABA_DRE = "DRE 2026"