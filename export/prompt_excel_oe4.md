# Prompt — Criação do Excel OE4: Plano de Financiamento do Investimento

## Objetivo

Cria um workbook Excel (.xlsx) intitulado **OE04_G18.xlsx** para a Grestel, Produtos Cerâmicos S.A., documentando o Objetivo Estratégico 4 (OE4) do Plano Estratégico Financeiro: **Plano de Financiamento do Investimento no Hub Logístico 4.0**.

Usa Python com a biblioteca **openpyxl** para gerar o ficheiro. Todos os valores abaixo foram computados por um motor financeiro (fonte de verdade); são definitivos e não devem ser alterados.

---

## Arquitetura do Workbook

O workbook tem **10 folhas**, na ordem:

| # | Nome da Folha | Conteúdo |
|---|--------------|----------|
| 0 | `00_PRESSUPOSTOS` | Todos os parâmetros de entrada — fonte de verdade do Excel |
| 1 | `01_Pre_Projeto` | Balanço e rácios financeiros pré-projeto (2024) |
| 2 | `02_Mapa_Investimento` | CAPEX Hub + Core + Fundo de Maneio |
| 3 | `03_Estrutura_Capital` | Critérios + rácios projetados 2024–2029 |
| 4 | `04_Plano_Financ` | Resumo das fontes de financiamento |
| 5 | `05_SD_BancoHub` | Serviço da dívida — Banco Hub |
| 6 | `06_SD_BEI` | Serviço da dívida — Linha BEI |
| 7 | `07_PT2030` | Mapa do subsídio PT2030 |
| 8 | `08_SD_Existente` | Serviço da dívida pré-existente |
| 9 | `09_Pos_Projeto` | Equilíbrio financeiro pós-projeto (2025–2029) |

**Princípio de ligação entre folhas:** A folha `00_PRESSUPOSTOS` é a âncora. As folhas 04, 05, 06 e 07 usam **fórmulas Excel** que referenciam células de `00_PRESSUPOSTOS` (ex.: `='00_PRESSUPOSTOS'!B5`). Assim o Excel tem lógica própria verificável.

---

## Paleta de Cores e Estilos

```python
BLUE_DARK  = "1F4E79"   # cabeçalhos de tabela (texto branco)
BLUE_MID   = "2E75B6"   # subtítulos (texto branco)
BLUE_LIGHT = "D6E4F0"   # linha de totais
GREY_LIGHT = "F2F2F2"   # linhas alternadas
GREEN_SOFT = "E2EFDA"   # condição cumprida / positivo
ORANGE     = "FCE4D6"   # condição não cumprida / alerta

# Fontes
TITLE_FONT = Font(bold=True, name="Arial", size=11, color="1F4E79")
HDR_FONT   = Font(bold=True, color="FFFFFF", name="Arial", size=9)
LABEL_FONT = Font(name="Arial", size=9)
BOLD_FONT  = Font(bold=True, name="Arial", size=9)
META_FONT  = Font(italic=True, name="Arial", size=8, color="808080")

# Formatos numéricos
EUR0 = '#,##0 €'
EUR2 = '#,##0.00 €'
PCT1 = '0.0%'
PCT2 = '0.00%'
DEC2 = '#,##0.00'
INT  = '#,##0'
```

Rodapé em todas as folhas (última linha, coluna A):
`"Gerado em 2026-05-24 | GrestelPy v0.9.5 | Cenário Base + Hub"`

---

## Folha 00_PRESSUPOSTOS

Título: `"00 | Pressupostos do Modelo — Hub Logístico 4.0"`

Esta folha centraliza todos os parâmetros. As outras folhas referenciam aqui via fórmulas. Escreve cada parâmetro com nome do intervalo definido (usando `ws.defined_names` em openpyxl ou simplesmente documenta os endereços).

**Layout: duas colunas (A = Label, B = Valor), com fundo amarelo nas células de valor (`FFFF00`).**

### Secção A — CAPEX e Cronograma (linhas 4–12)

| Linha | Ref. | Label | Valor | Formato |
|-------|------|-------|-------|---------|
| 4 | B4 | CAPEX Base Total (€) | 6 000 000 | EUR0 |
| 5 | B5 | CAPEX 2025 (€) | 2 850 000 | EUR0 |
| 6 | B6 | CAPEX 2026 (€) | 3 150 000 | EUR0 |
| 7 | B7 | Terreno — Custo de Oportunidade (€) | 150 000 | EUR0 |
| 8 | B8 | Gastos Pré-Operacionais 2025 (€) | 105 000 | EUR0 |
| 9 | B9 | Depreciação Anual 2026–2028 (€) | 705 035 | EUR0 |
| 10 | B10 | Depreciação Anual 2029 (€) | 640 035 | EUR0 |
| 11 | B11 | AFT Hub Líquido Fim 2029 (€) | 3 550 110 | EUR0 |
| 12 | B12 | Ano Início Benefícios | 2026 | INT |

### Secção B — Banco Hub (linhas 14–20)

| Linha | Ref. | Label | Valor | Formato |
|-------|------|-------|-------|---------|
| 14 | B14 | Banco Hub — Capital (€) | 3 000 000 | EUR0 |
| 15 | B15 | Banco Hub — Taxa de Juro (a.a.) | 0,0415 | PCT2 |
| 16 | B16 | Banco Hub — Amort. Anual (€) | 300 000 | EUR0 |
| 17 | B17 | Banco Hub — Início Amortização | 2028 | INT |
| 18 | B18 | Banco Hub — Ano Desembolso | 2025 | INT |
| 19 | B19 | Banco Hub — Período de Carência (anos) | 3 | INT |
| 20 | B20 | Banco Hub — Juros Capitalizados 2025 (€) | 124 500 | EUR0 |
| 21 | B21 | Banco Hub — Juros Capitalizados 2026 (€) | 180 750 | EUR0 |

### Secção C — Linha BEI (linhas 23–29)

| Linha | Ref. | Label | Valor | Formato |
|-------|------|-------|-------|---------|
| 23 | B23 | BEI — Capital (€) | 1 500 000 | EUR0 |
| 24 | B24 | BEI — Taxa de Juro (a.a.) | 0,0375 | PCT2 |
| 25 | B25 | BEI — Amort. Anual (€) | 150 000 | EUR0 |
| 26 | B26 | BEI — Início Amortização | 2028 | INT |
| 27 | B27 | BEI — Ano Desembolso | 2026 | INT |
| 28 | B28 | BEI — Período de Carência (anos) | 2 | INT |
| 29 | B29 | Total Capital Alheio Hub (€) | 4 500 000 | EUR0 |

_(nota: B29 pode ser fórmula `=B14+B23`)_

### Secção D — PT2030 (linhas 31–35)

| Linha | Ref. | Label | Valor | Formato |
|-------|------|-------|-------|---------|
| 31 | B31 | PT2030 — Montante (€) | 2 700 000 | EUR0 |
| 32 | B32 | PT2030 — % do CAPEX | 0,45 | PCT2 |
| 33 | B33 | PT2030 — Ano de Recebimento | 2027 | INT |
| 34 | B34 | PT2030 — Natureza | "Subsídio não reembolsável" | — |

### Secção E — Parâmetros de Viabilidade (linhas 37–45)

| Linha | Ref. | Label | Valor | Formato |
|-------|------|-------|-------|---------|
| 37 | B37 | IRC Taxa (%) | 0,245 | PCT2 |
| 38 | B38 | WACC Hub (%) | 0,073 | PCT2 |
| 39 | B39 | Capital Próprio Investido (€) | 1 500 000 | EUR0 |
| 40 | B40 | Autonomia Financeira Mínima (OE4) | 0,30 | PCT2 |
| 41 | B41 | DSCR Mínimo (covenant bancário) | 1,20 | DEC2 |
| 42 | B42 | RFAI — Taxa (%) | 0,10 | PCT2 |
| 43 | B43 | RFAI — CAPEX Elegível (€) | 6 000 000 | EUR0 |
| 44 | B44 | RFAI — Crédito Total Gerado (€) | 600 000 | EUR0 |
| 45 | B45 | RFAI — Limite IRC (%) | 0,50 | PCT2 |

### Secção F — KPIs de Viabilidade (linhas 47–55)

| Linha | Ref. | Label | Valor | Formato |
|-------|------|-------|-------|---------|
| 47 | B47 | VAL (€) | 3 619 509 | EUR0 |
| 48 | B48 | TIR (%) | 0,2584 | PCT2 |
| 49 | B49 | Payback Simples (anos) | 2,74 | DEC2 |
| 50 | B50 | Payback Atualizado (anos) | 3,68 | DEC2 |
| 51 | B51 | Índice de Rendibilidade (x) | 1,603 | DEC2 |
| 52 | B52 | Valor Residual dos Ativos (€) | 2 043 060 | EUR0 |
| 53 | B53 | Recuperação NFM Terminal (€) | 90 375 | EUR0 |
| 54 | B54 | Capital Bancário Vivo no Ano 10 (€) | 1 350 000 | EUR0 |
| 55 | B55 | Valor Terminal Total (€) | 2 133 435 | EUR0 |

---

## Folha 01_Pre_Projeto

Título: `"01 | Situação de Equilíbrio Financeiro Pré-Projeto (2024)"`
Subtítulo: `"Fonte: R&C 2024 auditado — Grestel, Produtos Cerâmicos S.A."`

### Tabela 1 — Balanço Resumido 31/12/2024

Formato: chave-valor (col A = label, col B = valor EUR0). **Todos os valores são hardcoded (não dependem de PRESSUPOSTOS).**

| Label | Valor (€) |
|-------|-----------|
| **Ativo Não Corrente** | **17 812 056** |
| — AFT líquido | 12 466 455 |
| — Goodwill | 1 701 104 |
| — Investimentos em Subsidiárias | 3 062 681 |
| — Outros Ativos NC | 581 817 |
| **Ativo Corrente** | **22 446 710** |
| — Inventários | 13 061 556 |
| — Clientes | 4 962 136 |
| — Caixa e equivalentes | 542 391 |
| — Outros AC | 3 880 444 |
| **Ativo Total Líquido** | **40 258 766** |
| *(linha em branco)* | |
| **Capital Próprio** | **12 199 666** |
| — Capital social | 526 318 |
| — Outros IC Próprio | 1 233 333 |
| — Reservas e RL transitados | 8 425 122 |
| — Resultado líquido 2024 | 1 390 209 |
| **Passivo Total** | **28 059 099** |
| — Empréstimos MLP (NC) | 12 203 269 |
| — Empréstimos CP | 5 530 545 |
| — Fornecedores | 3 914 140 |
| — Outros passivos | 6 411 145 |
| **Total CP + Passivo** | **40 258 766** |

### Tabela 2 — Rácios de Equilíbrio Financeiro Pré-Projeto

| Rácio | Fórmula / Valor | Formato |
|-------|----------------|---------|
| Autonomia Financeira [CP/ATL] | 30,3% | PCT1 |
| Solvabilidade [CP/Passivo] | 43,5% | PCT1 |
| Endividamento [Passivo/ATL] | 69,7% | PCT1 |
| Dívida Financeira Total (€) | 17 733 814 | EUR0 |
| Dívida Líquida (€) | 17 191 423 | EUR0 |
| EBITDA 2024 (€) | 4 149 679 | EUR0 |
| Dívida Líquida / EBITDA [x] | 4,14 | DEC2 |
| Cobertura de Juros [EBIT/Juros] | 3,75 | DEC2 |
| *(linha em branco)* | | |
| Condição AF mínima exigida (OE4) | 30,0% | PCT1 |
| Condição satisfeita pré-projeto? | "SIM ✓" | — |

Célula "SIM ✓" com fundo verde `E2EFDA`.

---

## Folha 02_Mapa_Investimento

Título: `"02 | Mapa de Investimento do Projeto"`
Subtítulo: `"Inclui: CAPEX Hub (AFT), CAPEX Core Grestel, Investimento em FM e Custo de Oportunidade do Terreno"`

### 2.1 — CAPEX Hub Logístico por Ano

Tabela com cabeçalho azul escuro:

| Ano | CAPEX Hub AFT (€) | Juros Cap. AFT (€) | Depreciação (€) | AFT Hub Líquido Fim Ano (€) |
|-----|------------------|--------------------|-----------------|----------------------------|
| 2025 | 2 850 000 | 124 500 | 0 | 2 974 500 |
| 2026 | 3 150 000 | 180 750 | 705 035 | 5 600 215 |
| 2027 | 0 | 0 | 705 035 | 4 895 180 |
| 2028 | 0 | 0 | 705 035 | 4 190 145 |
| 2029 | 0 | 0 | 640 035 | 3 550 110 |

### 2.2 — CAPEX Core Grestel por Ano (valores ilustrativos — usar motor para atualização)

Tabela com as colunas: Ano | CAPEX AFT Core (€) | CAPEX Intangíveis Core (€) | Total CAPEX Core (€).
(Preencher com zeros se não disponível; o motor calcula estes valores separadamente.)

### 2.3 — Investimento em Fundo de Maneio (Hub)

| Ano | ΔNFM Hub (€) | NFM Acumulado (€) |
|-----|-------------|-------------------|
| 2026 | 43 500 | 43 500 |
| 2027 | 0 | 43 500 |
| 2028 | 37 500 | 81 000 |
| 2029 | 9 375 | 90 375 |

### 2.4 — Resumo do Investimento Total (chave-valor)

| Rubrica | Valor (€) |
|---------|-----------|
| CAPEX Hub — Construção e Equipamento | 6 000 000 |
| Custo de Oportunidade do Terreno | 150 000 |
| Gastos Pré-Operacionais (formação 2025) | 105 000 |
| Juros Capitalizados (NCRF 10) | 305 250 |
| Investimento em Fundo de Maneio | 90 375 |
| **TOTAL INVESTIMENTO DO PROJETO** | **6 650 625** |
| dos quais financiados por PT2030 (subsídio) | 2 700 000 |
| dos quais financiados por Capital Alheio Bancário | 4 500 000 |
| dos quais financiados por Capitais Próprios | 1 500 000 |

---

## Folha 03_Estrutura_Capital

Título: `"03 | Critérios e Estrutura de Capital do Projeto"`

### 3.1 — Critérios para a Estrutura de Capital

Tabela 2 colunas (Critério | Orientação/Decisão), sem fórmulas:

| Critério | Orientação / Decisão |
|----------|---------------------|
| Manutenção da Autonomia Financeira | AF ≥ 30% em todos os anos (covenant bancário BPI e condição OE4) |
| Alavancagem moderada | Dívida Líq./EBITDA ≤ 4,0x — limita o risco de refinanciamento |
| Custo ponderado de capital (WACC) | Minimizar custo com diversificação: banco 4,15%, BEI 3,75%, PT2030 0% |
| Preservação de CP para autofinanciamento | RL retido cobre CAPEX Core anual; evitar aumentos de capital |
| Alívio fiscal — RFAI (CFI art. 22.º-23.º) | Crédito 10% × €6M = €600k deduzido diretamente à coleta IRC |
| Elegibilidade PT2030 (COMPETE 2030 — SI Inovação) | Subsídio não reembolsável 45% CAPEX; reduz CA e melhora AF |
| Financiamento BEI via banco parceiro | Taxa fixa 3,75% vs. mercado 4,15%; diversificação sem risco cambial |

### 3.2 — Projeção dos Rácios de Equilíbrio Financeiro (2024–2029)

Tabela com cabeçalho azul escuro. **Valores hardcoded (motor).** Todos os anos com AF ≥ 30% → fundo verde; abaixo → fundo laranja.

| Ano | Ativo Total (€) | Capital Próprio (€) | AF [%] | Solvab. [%] | Endivid. [%] | DívLíq/EBITDA [x] | Cob. Juros [x] | DSCR [x] |
|-----|----------------|--------------------|---------|-----------|-----------|--------------------|-----------------|----------|
| 2024 | 40 258 766 | 12 199 666 | 30,3% | 43,5% | 69,7% | 4,14 | 3,75 | 0,77 |
| 2025 | 42 118 220 | 14 445 570 | 34,3% | 52,2% | 65,7% | 3,27 | 5,95 | 0,50 |
| 2026 | 44 878 080 | 17 846 050 | 39,8% | 66,0% | 60,2% | 1,58 | 12,77 | 1,56 |
| 2027 | 47 913 240 | 21 417 540 | 44,7% | 80,8% | 55,3% | 0,84 | 13,78 | 1,92 |
| 2028 | 50 437 860 | 25 742 130 | 51,0% | 104,2% | 49,0% | 0,51 | 18,52 | 2,08 |
| 2029 | 53 378 190 | 30 381 580 | 56,9% | 132,1% | 43,1% | 0,26 | 24,43 | 2,43 |

Formatação condicional na coluna AF: ≥ 30% → fundo verde `E2EFDA`; < 30% → fundo laranja `FCE4D6`.

### 3.3 — Referências e Covenants (chave-valor)

| Indicador | Valor |
|-----------|-------|
| AF mínima exigida (OE4 e covenant BPI) | 30,0% |
| AF mínima recomendada (orientação prudencial) | 35,0% |
| DSCR mínimo aceitável | 1,20x |
| Dívida Líquida / EBITDA máx. (orientação) | 4,0x |

---

## Folha 04_Plano_Financ

Título: `"04 | Plano de Financiamento do Investimento"`
Subtítulo: `"Mínimo 2 fontes de capital alheio distintas + subsídio PT2030"`

### Tabela Principal — Fontes de Financiamento

Cabeçalho: `Fonte | Tipo | Montante (€) | Taxa Juro (a.a.) | Prazo (anos) | Início Amort. | Amort. Anual (€) | Desembolso | Reembolsável?`

**Usar fórmulas para referenciar 00_PRESSUPOSTOS nas colunas Montante, Taxa, Início Amort. e Amort. Anual:**

| Fonte | Tipo | Montante | Taxa Juro | Prazo | Início Amort. | Amort. Anual | Desembolso | Reembolsável? |
|-------|------|----------|-----------|-------|--------------|-------------|-----------|--------------|
| Banco Hub (BPI/CA) | Empréstimo Bancário | `='00_PRESSUPOSTOS'!B14` | `='00_PRESSUPOSTOS'!B15` | 10 | `='00_PRESSUPOSTOS'!B17` | `='00_PRESSUPOSTOS'!B16` | `='00_PRESSUPOSTOS'!B18` | Sim |
| Linha BEI (via banco parceiro) | Empréstimo BEI | `='00_PRESSUPOSTOS'!B23` | `='00_PRESSUPOSTOS'!B24` | 10 | `='00_PRESSUPOSTOS'!B26` | `='00_PRESSUPOSTOS'!B25` | `='00_PRESSUPOSTOS'!B27` | Sim |
| PT2030 — COMPETE 2030 SI Inovação | Subsídio Não Reembolsável | `='00_PRESSUPOSTOS'!B31` | 0,00% | 0 | `='00_PRESSUPOSTOS'!B33` | 0 | `='00_PRESSUPOSTOS'!B33` | Não |

Linha de totais (negrito, fundo `D6E4F0`):
- TOTAL FINANCIAMENTO PROJETO: `=B_BancoHub + B_BEI + B_PT2030` (deve resultar em 7 200 000 €)

### Nota sobre Capitais Próprios

Texto em itálico tamanho 8:
> "A estrutura de capital não requer reforço de CP no cenário base: AF pré-projeto (30,3%) cumpre o covenant mínimo de 30% e o plano mantém AF ≥ 30% em todos os anos. Capital Próprio investido: €1.500.000 (25% CAPEX — autofinanciamento via resultados retidos e reservas livres)."

---

## Folha 05_SD_BancoHub

Título: `"05 | Serviço da Dívida — Banco Hub (Empréstimo Bancário)"`

### Pressupostos (chave-valor, referenciando 00_PRESSUPOSTOS)

| Label | Fórmula / Valor |
|-------|----------------|
| Fonte / Designação | "Banco Hub (BPI / Banco Comercial Parceiro)" |
| Tipo | "Empréstimo Bancário (Euribor 3M + Spread)" |
| Capital (€) | `='00_PRESSUPOSTOS'!B14` → 3 000 000 |
| Taxa de Juro (a.a.) | `='00_PRESSUPOSTOS'!B15` → 4,15% |
| Amortização Anual (€) | `='00_PRESSUPOSTOS'!B16` → 300 000 |
| Início Amortização | `='00_PRESSUPOSTOS'!B17` → 2028 |
| Desembolso | `='00_PRESSUPOSTOS'!B18` → 2025 |
| Período de Carência | `='00_PRESSUPOSTOS'!B19` → 3 anos |

Nota em itálico: `"Taxa indexada: Euribor 3M (2,90%) + Spread 1,25% = 4,15%. Revisão anual. Amortização linear. Carência 3 anos (2025-2027: obra + ramp-up)."`

### Plano de Amortização

Cabeçalho: `Ano | Saldo Inicial (€) | Desembolso (€) | Juros Totais (€) | Juros Capitalizados (€) | Juros Expensed P&L (€) | Amortização (€) | Saldo Final (€)`

**Valores hardcoded (do motor) — mas as colunas de juros e saldo devem ter fórmulas internas para consistência:**

| Ano | Saldo Ini. | Desembolso | Juros Totais | Juros Cap. | Juros Expensed | Amortização | Saldo Final |
|-----|-----------|-----------|-------------|-----------|---------------|------------|------------|
| 2025 | 0 | 3 000 000 | 124 500 | 124 500 | 0 | 0 | 3 000 000 |
| 2026 | 3 000 000 | 0 | 124 500 | 124 500 | 0 | 0 | 3 000 000 |
| 2027 | 3 000 000 | 0 | 124 500 | 0 | 124 500 | 0 | 3 000 000 |
| 2028 | 3 000 000 | 0 | 124 500 | 0 | 124 500 | 300 000 | 2 700 000 |
| 2029 | 2 700 000 | 0 | 112 050 | 0 | 112 050 | 300 000 | 2 400 000 |

_(Nota: Juros 2025 = 3 000 000 × 4,15% = 124 500; Juros 2029 = 2 700 000 × 4,15% = 112 050)_

Após a tabela, adicionar **Totais do Período** (chave-valor):

| Rubrica | Valor |
|---------|-------|
| Total Juros Pagos (€) | 609 550 |
| Total Amortizações Capital (€) | 600 000 |
| Total Serviço da Dívida (€) | 1 209 550 |

---

## Folha 06_SD_BEI

Título: `"06 | Serviço da Dívida — Linha BEI (European Investment Bank)"`

### Pressupostos (referenciando 00_PRESSUPOSTOS)

| Label | Fórmula / Valor |
|-------|----------------|
| Capital (€) | `='00_PRESSUPOSTOS'!B23` → 1 500 000 |
| Taxa de Juro (a.a.) | `='00_PRESSUPOSTOS'!B24` → 3,75% |
| Amortização Anual (€) | `='00_PRESSUPOSTOS'!B25` → 150 000 |
| Início Amortização | `='00_PRESSUPOSTOS'!B26` → 2028 |
| Desembolso | `='00_PRESSUPOSTOS'!B27` → 2026 |

Nota: `"Taxa fixa 3,75% (benefício BEI vs. mercado ~4,15%). Desembolso 2026 após conclusão da 1.ª fase. Elegível por investimento em infraestrutura logística e digitalização (InvestEU)."`

### Plano de Amortização

| Ano | Saldo Ini. | Desembolso | Juros Totais | Juros Cap. | Juros Expensed | Amortização | Saldo Final |
|-----|-----------|-----------|-------------|-----------|---------------|------------|------------|
| 2025 | 0 | 0 | 0 | 0 | 0 | 0 | 0 |
| 2026 | 0 | 1 500 000 | 56 250 | 56 250 | 0 | 0 | 1 500 000 |
| 2027 | 1 500 000 | 0 | 56 250 | 0 | 56 250 | 0 | 1 500 000 |
| 2028 | 1 500 000 | 0 | 56 250 | 0 | 56 250 | 150 000 | 1 350 000 |
| 2029 | 1 350 000 | 0 | 50 625 | 0 | 50 625 | 150 000 | 1 200 000 |

_(Juros 2026: 1 500 000 × 3,75% = 56 250; Juros 2029: 1 350 000 × 3,75% = 50 625)_

Totais do Período:

| Rubrica | Valor |
|---------|-------|
| Total Juros Pagos (€) | 219 375 |
| Total Amortizações Capital (€) | 300 000 |
| Total Serviço da Dívida (€) | 519 375 |

---

## Folha 07_PT2030

Título: `"07 | Subsídio PT2030 — COMPETE 2030 SI Inovação"`
Subtítulo: `"Instrumento não reembolsável — sem serviço de dívida associado"`

### Caracterização do Instrumento (chave-valor, referencias 00_PRESSUPOSTOS para Montante e Ano)

| Label | Fórmula / Valor |
|-------|----------------|
| Programa | "Portugal 2030 — COMPETE 2030" |
| Tipologia | "SI Inovação Produtiva" |
| Natureza | "Subsídio não reembolsável (fundo perdido)" |
| Montante Previsto (€) | `='00_PRESSUPOSTOS'!B31` → 2 700 000 |
| % do CAPEX Elegível | `='00_PRESSUPOSTOS'!B32` → 45,00% |
| Ano de Recebimento Previsto | `='00_PRESSUPOSTOS'!B33` → 2027 |
| Critérios de Elegibilidade | "PME/Mid-Cap, CAE 23 (cerâmica), região Centro, investimento produtivo" |
| Tratamento Contabilístico | "NCRF 22 — rendimento diferido; amortizado proporcionalmente à dep. dos ativos" |
| Impacto na AF | "Reduz capital alheio líquido; melhora AF no ano de receção" |

### Mapa de Reconhecimento NCRF 22 (tabela)

Cabeçalho: `Ano | Depreciação Pools (€) | CAPEX Base Elegível (€) | Rácio Reconhecimento | PT2030 Accrual (€) | PT2030 Recebimento Caixa (€)`

Coluna "Rácio" = Dep_Pools / CAPEX_Base (fórmula interna na folha).
Coluna "CAPEX Base Elegível" pode referenciar `='00_PRESSUPOSTOS'!B4`.
Coluna "PT2030 Recebimento Caixa" pode referenciar `='00_PRESSUPOSTOS'!B31` na linha de 2027.

| Ano | Dep. Pools (€) | CAPEX Elegível (€) | Rácio | Accrual PT2030 (€) | Recebimento Caixa (€) |
|-----|---------------|-------------------|---------|--------------------|----------------------|
| 2025 | 0 | 6 000 000 | 0,00% | 0 | 0 |
| 2026 | 692 825 | 6 000 000 | 11,55% | 311 771 | 0 |
| 2027 | 692 825 | 6 000 000 | 11,55% | 311 771 | 2 700 000 |
| 2028 | 692 825 | 6 000 000 | 11,55% | 311 771 | 0 |
| 2029 | 627 825 | 6 000 000 | 10,46% | 282 521 | 0 |
| **TOTAL** | **2 706 300** | — | **45,11%** | **1 217 834** | **2 700 000** |

Linha de totais com negrito e fundo `D6E4F0`.

### Nota sobre Saldo Diferido (conta 282)

Texto explicativo: o subsídio de €2.700.000 é reconhecido como rendimento diferido (SNC: conta 282) quando recebido em 2027. É amortizado como outros rendimentos ao longo da vida útil dos ativos, na proporção da depreciação anual. Total reconhecido no horizonte 2026–2029: €1.217.834. Saldo diferido a reconhecer após 2029: €1.482.166.

---

## Folha 08_SD_Existente

Título: `"08 | Serviço da Dívida Pré-Existente (2024–2029)"`
Subtítulo: `"Fontes: BPI, Santander, CGD, Abanca, IAPMEI, Locações Financeiras"`

### 8.1 — Serviço da Dívida Agregado

Cabeçalho: `Ano | Capital em Dívida FdA (€) | Amortiz. Capital (€) | Juros (€) | Serviço Total (€) | Emp. NC (€) | Emp. CP (€)`

| Ano | Cap. em Dívida FdA | Amortiz. Capital | Juros | Serviço Total | Emp. NC | Emp. CP |
|-----|-------------------|-----------------|-------|--------------|---------|---------|
| 2024 | 18 867 135 | 300 000 | 528 161 | 828 161 | 12 203 269 | 5 530 545 |
| 2025 | 16 336 588 | 5 530 545 | 419 000 | 5 949 545 | 15 248 604 | 1 087 984 |
| 2026 | 13 372 852 | 2 043 094 | 318 061 | 2 361 155 | 11 049 319 | 2 323 533 |
| 2027 | 11 499 319 | 1 873 533 | 259 708 | 2 133 241 | 9 325 407 | 2 173 912 |
| 2028 | 9 325 407 | 2 173 912 | 207 567 | 2 381 479 | 7 312 500 | 2 012 907 |
| 2029 | 7 312 500 | 2 012 907 | 157 398 | 2 170 305 | 5 525 000 | 1 787 500 |

### 8.2 — Nota Explicativa

Texto: "A dívida pré-existente inclui linhas BPI (taxa 2,81%), linhas IAPMEI (taxa 0% — COVID), e Locações Financeiras (taxa 3,5%). O montante elevado de amortizações em 2025 reflete o vencimento das tranches IAPMEI e a reestruturação do passivo corrente. Após 2025, o serviço estabiliza em ~€2,0–2,4M/ano."

---

## Folha 09_Pos_Projeto

Título: `"09 | Situação de Equilíbrio Financeiro Pós-Projeto (2025–2029)"`
Subtítulo: `"Condição OE4: AF ≥ 30% em todos os anos do horizonte"`

### Tabela Principal — Rácios Projetados (com Hub)

**Referenciar os dados da folha 03_Estrutura_Capital** para os anos 2025–2029. Em openpyxl, usar fórmulas do tipo `='03_Estrutura_Capital'!B_linha` para cada célula.

Cabeçalho: `Ano | Ativo Total (€) | Capital Próprio (€) | AF [%] | Solvab. [%] | Endivid. [%] | Dív. Líquida (€) | EBITDA (€) | DívLíq/EBITDA [x] | Cob. Juros [x] | DSCR [x]`

Os valores são os mesmos da tabela 3.2 (anos 2025–2029). Referenciá-los com fórmulas cross-sheet é o ponto chave:

Por exemplo, se na folha 03 os dados estão nas linhas 9–13 (anos 2025–2029), colunas B–J:
```
='03_Estrutura_Capital'!B9   (Ativo Total 2025)
='03_Estrutura_Capital'!C9   (CP 2025)
='03_Estrutura_Capital'!D9   (AF 2025)
... etc.
```

### Verificação AF ≥ 30% — Semáforo

Tabela após a principal, cabeçalho: `Ano | AF Projetada | Condição AF ≥ 30% | Margem (p.p.)`

| Ano | AF Projetada | Condição | Margem |
|-----|-------------|----------|--------|
| 2025 | 34,3% | ✓ CUMPRE | +4,3 pp |
| 2026 | 39,8% | ✓ CUMPRE | +9,8 pp |
| 2027 | 44,7% | ✓ CUMPRE | +14,7 pp |
| 2028 | 51,0% | ✓ CUMPRE | +21,0 pp |
| 2029 | 56,9% | ✓ CUMPRE | +26,9 pp |

Fórmulas para "Condição": `=IF(AF>=0.30,"✓ CUMPRE","✗ NÃO CUMPRE")`
Fórmulas para "Margem": `=AF - 0.30`

Fundo verde (`E2EFDA`) em todas as linhas (todas cumprem). Se AF < 30%, fundo laranja (`FCE4D6`).

### Nota de Conclusão

Texto (itálico, tamanho 9):
> "O Hub Logístico 4.0 não deteriora a solidez financeira da Grestel: a Autonomia Financeira melhora de 30,3% (2024) para 56,9% (2029). O DSCR Hub mínimo é de 1,58x (2028), acima do covenant bancário de 1,20x. O VAL do projeto é de €3.620k (WACC 7,3%), TIR de 25,8%, com payback atualizado de 3,7 anos."

---

## Instruções de Implementação

1. **Geração:** Usar `openpyxl` em Python. Criar todas as folhas na ordem indicada.

2. **Folha 00_PRESSUPOSTOS:** Escrever os valores com `ws.cell(row=..., col=..., value=...)`. Aplicar fundo amarelo (`FFFF00`) nas células de valor (coluna B).

3. **Fórmulas cross-sheet:** Usar strings de fórmula, ex.:
   ```python
   ws.cell(row=5, column=2, value="='00_PRESSUPOSTOS'!B14")
   ```
   O openpyxl escreve a string como fórmula; o Excel calcula quando abre o ficheiro.

4. **Freeze panes:** Em todas as folhas, `ws.freeze_panes = "A4"` (ou `"B5"` para tabelas com cabeçalho na linha 4).

5. **Larguras de coluna:** Col A ≥ 36 char, colunas de valor ≥ 18 char.

6. **Bordas:** Usar `Border(left=thin, right=thin, top=thin, bottom=thin)` em todas as células de dados.

7. **Rodapé:** Última linha de cada folha: `"Gerado em 2026-05-24 | GrestelPy v0.9.5 | Cenário Base + Hub"` em fonte itálica cinza.

8. **Output:** Guardar como `OE04_G18.xlsx` no diretório de trabalho.

---

*Prompt gerado em 2026-05-24 — GrestelPy v0.9.5 | Todos os valores auditados.*
