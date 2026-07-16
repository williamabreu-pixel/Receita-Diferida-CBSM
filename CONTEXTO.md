# CONTEXTO DO PROJETO — Sistema de Conciliação CBSM (IFRS 15 / CPC 47)

> Este documento é o briefing completo do projeto. Leia-o inteiro antes de agir.
> Ele carrega decisões, validações e regras de negócio já acordadas. NÃO
> reinvente arquitetura nem tome atalhos que contrariem o que está aqui.

## 0. Regras de ouro para quem for mexer no projeto
1. **O alvo é o app Streamlit interativo (`streamlit run app.py`)**, NÃO um HTML
   estático. Não gere "snapshots" HTML como substituto do dashboard — isso perde
   interatividade, filtros e exportação. HTML estático só se for explicitamente pedido.
2. **Não altere a lógica de negócio já validada** (data_loader, data_pipeline)
   sem avisar. Os números batem ao centavo com o gabarito; qualquer mudança que
   afete valores deve ser justificada e revalidada.
3. **Explique antes de apagar, mover ou instalar.** Nunca rode comandos destrutivos
   (Remove-Item, del, rm) sem explicar o porquê e ter confirmação.
4. **Premissas contábeis são inputs explícitos, nunca "adivinhados" pelo código.**

## 1. O que o sistema faz
Concilia o passivo e a receita diferida de pontos de fidelidade (Dotz), conforme
IFRS 15 / CPC 47, cruzando:
- **Lado contábil**: balancete do SAP (aba Trial Balance).
- **Lado controle**: planilha interna U1 (aba U1 - Movimentação).
E gera lançamentos contábeis de ajuste para os desvios.

## 2. Motores de receita (4)
- **Trocas** — resgate de pontos (receita de resgate + provisão manual "a faturar" 3111199).
- **Breakage** — expiração estimada de pontos que viram receita.
- **Spread** — margem sobre a venda de pontos.
- **Promodotz** — pontos promocionais.

## 3. Arquitetura (definida e validada)
**HÍBRIDA:**
- **Atuais / conciliação** → dados vêm do balancete e da U1 (via data_loader).
- **Projeção / forecast** → motor atuarial de safras (apropriação linear em 48 meses).

### Fontes de dados
| Insumo | Origem | Tipo |
|---|---|---|
| Saldos contábeis (passivo/receita) | Balancete SAP (Trial Balance) | dado |
| Passivo de controle + receita por bloco | Planilha U1 - Movimentação | dado (controle interno) |
| Receita por motor | Base Receita Motor | dado |
| Régua de taxas breakage (25/28/32%) | Config | PREMISSA contábil |
| CPD (0,0135) | Config | PREMISSA contábil |
| Provisão "Receita a Faturar" (3111199) | Lançamento manual no SAP | PREMISSA/julgamento |

## 4. Regras de conciliação (Quadros A–G do Dashboard original)
- **Quadro A**: Passivo Diferido Total contábil (contas grupo 219 + 231, analíticas
  de 7 dígitos, sinal invertido) DEVE bater com o passivo total da U1.
  Valor validado: **R$ 196.321.706,82**.
- **Quadro G**: Passivo Total = Breakage + Custo do Produto + Spread (decomposição = 0).
- **Receita por bloco** (validado ao centavo, jun/2026 YTD):
  - Breakage: 21.610.402,09
  - Custo do Produto/Trocas: 22.665.859,03
  - Spread: 16.095.728,24
  - Promodotz: 1.075.793,82
- Tolerância de conciliação: R$ 1,00 (cobre o plug de arredondamento ~0,005).

## 5. Regras contábeis importantes (NÃO esquecer)
- **Sinais**: no balancete SAP, passivo (grupo 2) e receita (grupo 3) vêm com
  crédito NEGATIVO. O RAW preserva o sinal do SAP; a inversão para "sentido de
  negócio" (positivo) é feita na camada de negócio.
- **Base Receita Motor JÁ vem positiva** (sentido de negócio). NÃO inverter o sinal dela.
- **Contas de passivo**: grupos 219 (circulante) e 231 (não circulante). Usar SOMENTE
  contas ANALÍTICAS de 7 dígitos (startswith + len==7) para evitar dupla contagem
  pai/filho da hierarquia SAP.
- **Split CP/LP**: NÃO é derivável por prefixo. Vem da reclassificação da U1
  (coluna de passivo recalculado). Item ainda EM ABERTO no motor atuarial.
- **_to_float**: deve tolerar formato BR (1.234,56) e US/Excel (1234.56). A decisão
  é pela presença de vírgula. NÃO remover o ponto decimal cegamente (inflava valores 100x).
- **Contas contábeis são TEXTO**, nunca número (ex: '3111160', não 3111160.0).

## 6. Motor atuarial (para projeção)
- Cada safra (mês de emissão) tem um Valor Faturado a reconhecer.
- Apropriação LINEAR em 48 meses, começando no mês SEGUINTE ao da emissão
  (janela [emissão+1, emissão+48]).
- Breakage por safra = pontos emitidos × taxa_breakage(ano) × preço negociado.
  - Régua de taxas: 2018–2024 = 25%; 2025 = 28%; 2026+ = 32%.
- Spread/Margem por safra = pontos emitidos × taxa_margem × (preço − CPD).
- Validado ao centavo contra o gabarito (receita) e ~0,0005% (passivo).

## 7. Arquivos do sistema
- `config.py` — caminho do Excel (dados.xlsx) e nomes das 3 abas.
- `data_loader.py` — camada RAW: lê balancete, Base Receita Motor e U1. Tipagem robusta.
- `data_pipeline.py` — camada de negócio: agrega receita, monta DRE, concilia por blocos.
- `gerador_lancamentos.py` — gera lançamentos de ajuste (partida dobrada: Débito=J se J>0,
  Crédito=-J se J<0). Checks: partidas dobradas fecham + radar de cobertura.
- `app.py` — dashboard Streamlit (interativo): alerta verde/vermelho, cards de passivo
  (contábil/controle/delta), barras empilhadas de receita, rosca de composição do passivo,
  tabela de conciliação, DRE e exportação CSV dos lançamentos.

## 8. Itens EM ABERTO (finish line)
1. **Split CP/LP** no motor atuarial (reclassificação da U1: CP=135,0mi / LP=61,3mi
   sobre total 196,3mi — não sai de prefixo).
2. **Lançamento tipo 2 (reclassificação CP×LP)** no gerador_lancamentos — depende do item 1.
3. Rodar o app Streamlit de verdade (não HTML estático).

## 9. Ambiente (Windows) — cuidado conhecido
- A pasta do projeto está no OneDrive com caminho longo, o que estoura o limite de
  ~260 caracteres do Windows e QUEBRA a instalação do streamlit dentro de .venv local.
- Solução adotada: ambiente virtual em caminho curto (ex: %USERPROFILE%\venvs\sistema_cbsm).
- Recomendação futura: mover o projeto para caminho curto fora do OneDrive (ex: C:\projetos\cbsm).
- Rodar SEMPRE com o ambiente completo (o de caminho curto). Comando: `streamlit run app.py`.

## 10. Status da validação
Tudo que envolve os ATUAIS está validado e batendo ao centavo com o gabarito Excel.
O que falta é o split CP/LP (projeção) e rodar o dashboard interativo de forma estável.
