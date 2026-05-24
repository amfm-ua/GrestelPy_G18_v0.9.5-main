# Guia: Construção do Excel Hub Logístico — Coerente com o Motor Python

> **Princípio único:** o motor Python é a fonte de verdade.
> O Excel replica a lógica, não recalcula de forma independente.
> Fluxo: correr Python → extrair valores → colar no Excel → validar com fórmulas.

---

## 0. Pré-requisitos: extrair todos os valores do motor

Execute uma vez. Guarda os DataFrames/dicts em variáveis — são eles que alimentam cada célula do Excel.

```python
import sys, pathlib
sys.path.insert(0, str(pathlib.Path("src")))

from engine.projetos.hub_logistico import (
    load, hub_capex, hub_financing, hub_fcf,
    hub_dr_impact, hub_dfc_impact,
    pt2030_reconhecimento, hub_nfm, hub_rfai,
    mapa_servico_divida, viabilidade_hub,
)

# ── parâmetros base (iguais ao motor) ────────────────────────────────
hub      = load()
IRC      = 0.21          # irc_taxa
WACC     = 0.073         # WACC Hub (7,3 %)
ANOS     = list(range(2025, 2030))   # horizonte FCF base
ANOS_EXT = list(range(2025, 2035))   # se for necessário 10 anos no Excel

# ── módulos ──────────────────────────────────────────────────────────
df_capex   = hub_capex(hub)                    # CAPEX + dep + AFT rolling
df_fin     = hub_financing(hub)                # plano amortização empréstimos
df_fcf     = hub_fcf(hub, IRC)                 # FCF livre unlevered
dr         = hub_dr_impact(hub)                # P&L incremental por ano
dfc        = hub_dfc_impact(hub)               # DFC incremental por ano
pt2030_map = pt2030_reconhecimento(hub)        # accrual NCRF 22 por ano
rfai_map   = hub_rfai(hub, IRC)                # crédito RFAI por ano
nfm_map    = hub_nfm(hub)                      # ΔNFM por ano
df_dscr    = mapa_servico_divida(hub)          # DSCR anual
via        = viabilidade_hub(hub, IRC, WACC)   # KPIs de viabilidade

# ── imprimir para confirmar antes de colar no Excel ──────────────────
print(df_fcf.to_string())
print(df_dscr.to_string())
print(df_capex.to_string())
print(df_fin.to_string())
for y in ANOS:
    print(y, dr[y])
    print(y, "PT2030:", pt2030_map[y], "RFAI:", rfai_map[y])
print(via["val"], via["tir"], via["payback_simples"], via["payback_atualizado"])
```

---

## 1. Estrutura do Workbook

| Sheet | Conteúdo | Fonte Python |
|-------|----------|-------------|
| `PRESSUPOSTOS` | Tabela de inputs auditável | `hub["projeto_hub"]` (YAML) |
| `CAPEX_DEP` | Schedule CAPEX + depreciação + AFT | `hub_capex(hub)` |
| `FINANCIAMENTO` | Plano amortização BancoHub + BEI | `hub_financing(hub)` |
| `DR_INCR` | P&L incremental 2025-2034 | `hub_dr_impact(hub)` |
| `PT2030_RFAI` | Reconhecimento subsídio + RFAI | `pt2030_reconhecimento()` + `hub_rfai()` |
| `FCF` | Fluxo de caixa livre (FCFF) | `hub_fcf(hub, IRC)` |
| `DSCR` | Serviço dívida e cobertura | `mapa_servico_divida(hub)` |
| `KPIS` | VAL, TIR, Payback, IR, tornado | `viabilidade_hub(hub, IRC, WACC)` |

---

## 2. Sheet `PRESSUPOSTOS`

Esta sheet é a única onde entram dados fixos. Todas as outras sheets referenciam daqui.

| Linha | Label | Valor | Referência Python |
|-------|-------|-------|-------------------|
| 3 | CAPEX Base (€) | 6 000 000 | `proj["capex"]["base"]` |
| 4 | PT2030 Montante (€) | 2 700 000 | `proj["financiamento"]["PT2030"]["montante"]` |
| 5 | PT2030 Ano Recebimento | 2027 | `proj["financiamento"]["PT2030"]["ano_recebimento"]` |
| 6 | Banco Hub — Capital (€) | 3 000 000 | `proj["financiamento"]["Banco_Hub"]["montante"]` |
| 7 | Banco Hub — Taxa Juro | 4,15% | `proj["financiamento"]["Banco_Hub"]["taxa_juro"]` |
| 8 | Banco Hub — Início Amort. | 2028 | `proj["financiamento"]["Banco_Hub"]["inicio_amortizacao"]` |
| 9 | Banco Hub — Amort. Anual (€) | 300 000 | `proj["financiamento"]["Banco_Hub"]["amortizacao_anual"]` |
| 10 | BEI — Capital (€) | 1 500 000 | `proj["financiamento"]["Linha_BEI"]["montante"]` |
| 11 | BEI — Taxa Juro | 3,75% | `proj["financiamento"]["Linha_BEI"]["taxa_juro"]` |
| 12 | BEI — Início Amort. | 2028 | `proj["financiamento"]["Linha_BEI"]["inicio_amortizacao"]` |
| 13 | BEI — Amort. Anual (€) | 150 000 | `proj["financiamento"]["Linha_BEI"]["amortizacao_anual"]` |
| 14 | IRC Taxa | 21% | `proj["viabilidade"]["irc_taxa"]` |
| 15 | WACC Hub | 7,30% | `proj["viabilidade"]["wacc"]` |
| 16 | RFAI Taxa | 10% | `proj["rfai"]["taxa"]` |
| 17 | RFAI CAPEX Elegível (€) | 6 000 000 | `proj["rfai"]["capex_elegivel"]` |
| 18 | RFAI Limite IRC | 50% | `proj["rfai"]["limite_irc_pct"]` |
| 19 | Poupança Operacional Base (€) | ver YAML | `proj["beneficios_anuais"]["poupanca_operacional"]` |
| 20 | Ano Início Benefícios | 2026 | `proj["ano_inicio_beneficios"]` |

> **Convenção:** nomeie os intervalos (`Ctrl+Shift+F3`) para usar como `CAPEX_BASE`, `PT2030_MONTANTE`, `IRC_TAXA`, `WACC`, etc. nas fórmulas das outras sheets.

---

## 3. Sheet `CAPEX_DEP`

**Fonte:** `df_capex = hub_capex(hub)` — um DataFrame com linhas por ano.

### Layout (colunas A–H, linhas começam na 5)

| Col | Label | Python (df_capex) | Formato |
|-----|-------|-------------------|---------|
| A | Ano | `df_capex["ano"]` | Inteiro |
| B | CAPEX Caixa (€) | `df_capex["capex"]` | EUR0 |
| C | Juros Capitalizados AFT (€) | `df_capex["juros_capitalizados_aft"]` | EUR0 |
| D | Depreciação Total (€) | `df_capex["depreciacao"]` | EUR0 |
| E | Dep. Pools Base (€) | `df_capex["dep_pools"]` | EUR0 |
| F | Dep. Juros Capitalizados (€) | `df_capex["dep_juros_cap"]` | EUR0 |
| G | AFT Líquido Fim Ano (€) | `df_capex["aft_liquido_fim"]` | EUR0 |

### Fórmulas de validação (coluna H — só para conferência)

Na linha do ano 2027 (ex. linha 7), célula H7:
```excel
= G6 + B7 + C7 - D7
```
(AFT_fim = AFT_ano_anterior + CAPEX_caixa + Juros_cap - Dep_total)

Deve coincidir com `G7`. Se não coincidir, há inconsistência.

### Linha de totais (última linha do bloco)

```excel
= SUM(B5:B9)   ' total CAPEX (deve = −6 000 000 se só 2025-2026 têm CAPEX)
= SUM(D5:D9)   ' total depreciação no horizonte base
```

---

## 4. Sheet `FINANCIAMENTO`

**Fonte:** `df_fin = hub_financing(hub)`

### Layout — um bloco por tranche (separados por linha em branco)

#### Bloco 4a — Banco Hub (linhas 5–14, colunas A–H)

| Col | Label | Python (df_fin) | Fórmula alternativa Excel |
|-----|-------|----------------|--------------------------|
| A | Ano | `df_fin["ano"]` | — |
| B | Saldo Início Ano (€) | saldo do ano anterior | `=H[linha-1]` |
| C | Desembolso (€) | `df_fin["desembolso"]` (só 2025) | `=IF(A5=PRES!B6, PRES!B6, 0)` |
| D | Juros Totais (€) | `df_fin["juros"]` | `=B5 * PRES!B7` |
| E | Juros Capitalizados (€) | `df_fin["juros_capitalizados"]` | ver nota abaixo |
| F | Juros Expensed P&L (€) | `df_fin["juros_expensed"]` | `=D5-E5` |
| G | Amortização Capital (€) | `df_fin["amortizacao"]` | `=IF(A5>=PRES!B8, PRES!B9, 0)` |
| H | Saldo Fim Ano (€) | `df_fin["saldo_fim"]` | `=B5+C5-G5` |

> **Nota Juros Capitalizados (col E):** o motor capitaliza juros durante o período de construção (2025-2026) para a tranche Banco Hub. A lógica está em `base._juros_capitalizados_map()`. Na prática: se o ano é período de carência E o ativo ainda está em construção, `juros_cap = juros_totais`. A partir de 2027 (ativo em operação): `juros_cap = 0`.

#### Bloco 4b — Linha BEI (mesmo layout, logo abaixo)

Mesmas colunas mas com as taxas e montantes da BEI (`PRES!B10..B13`).

#### Bloco 4c — Consolidado (soma BancoHub + BEI)

| Col | Label | Fórmula |
|-----|-------|---------|
| D | Juros Totais Consolidados | `= Banco_D + BEI_D` |
| E | Juros Capitalizados Total | `= Banco_E + BEI_E` |
| F | Juros P&L Total | `= D - E` |
| G | Amortização Total | `= Banco_G + BEI_G` |
| H | Serviço Total Dívida | `= D + G` (exceto juros cap que não saem a caixa) |

---

## 5. Sheet `DR_INCR`

**Fonte:** `dr = hub_dr_impact(hub)` — dict `{ano: {campo: valor}}`.

Esta é a Demonstração de Resultados **incremental** do Hub (só os efeitos adicionais face ao cenário sem Hub). Cobre 2025-2034.

### Layout — linhas = rubricas, colunas = anos

```
       Col B    Col C    Col D    Col E    Col F    Col G    Col H    Col I    Col J    Col K
Linha  2025     2026     2027     2028     2029     2030     2031     2032     2033     2034
```

#### Bloco Receitas (linhas 5–10)

| Linha | Rubrica | Python key | Sinal |
|-------|---------|------------|-------|
| 5 | Poupança Pessoal (€) | `dr[y]["pessoal_reducao"]` | Positivo (redução de custo) |
| 6 | Poupança FSE (€) | `dr[y]["fse_reducao"]` | Positivo |
| 7 | Redução CMVMC / Quebras (€) | `dr[y]["cmvmc_reducao"]` | Positivo |
| 8 | VN Incremental B2C (€) | `dr[y]["vn_incremental"]` | Positivo |
| 9 | Outros Rendimentos — PT2030 (€) | `dr[y]["outros_rend_subsidio"]` | Positivo (não-caixa) |
| 10 | Libertação Inventário (€) | `dr[y]["inventario_libertado"]` | Positivo (one-time 2026) |

#### Bloco Gastos (linhas 12–16)

| Linha | Rubrica | Python key | Sinal |
|-------|---------|------------|-------|
| 12 | OPEX Hub — FSE (€) | `dr[y]["fse_opex_hub"]` | Negativo |
| 13 | CMVMC Incremental B2C (€) | `dr[y]["cmvmc_incremental"]` | Negativo |
| 14 | Gastos Pré-Operacionais (€) | `dr[y].get("gastos_preop", 0)` | Negativo (só 2025) |
| 15 | Depreciação Hub (€) | `dr[y]["depreciacao_hub"]` | Negativo |

#### Bloco Resultados (linhas 18–24)

| Linha | Rubrica | Fórmula Excel | Python equiv |
|-------|---------|---------------|-------------|
| 18 | **EBITDA Incremental** | `=SUM(5:10)+SUM(12:13)` (excluindo dep) | `dr[y]["ebitda_impact"]` |
| 19 | Depreciação (já em L15) | referência | `dr[y]["depreciacao_hub"]` |
| 20 | **EBIT Incremental** | `=L18-L19` | `dr[y]["ebit_impact"]` |
| 21 | Juros Expensed Hub (€) | ref. FINANCIAMENTO!F | `df_fin["juros_expensed"]` |
| 22 | **EBT Incremental** | `=L20-L21` | — |
| 23 | IRC Incremental (€) | `=MAX(L22,0)*IRC_TAXA` | — |
| 24 | **Resultado Líquido Incremental** | `=L22-L23` | — |

> **Atenção:** as linhas 2030-2034 (cols H-K) replicam a mesma lógica mas os valores para esses anos **não saem do `hub_fcf()` directamente** (o FCF só vai a 2029 no cenário base). Para 2030-2034, use `hub_dr_impact(hub)` que itera sobre `YEARS` definido em `engine/inputs/constants.py`. Verifique se `YEARS` inclui 2030-2034 — se não incluir, os valores de extensão têm de ser calculados manualmente ou estendendo o YAML.

#### Validação cruzada obrigatória

Na linha de EBITDA (L18), para cada ano:
```excel
= valor da célula deve coincidir com df_fcf["ebitda_impact"] para esse ano
```
Se divergir → há uma discrepância entre `hub_dr_impact` e `hub_fcf`.

---

## 6. Sheet `PT2030_RFAI`

**Fonte:** `pt2030_reconhecimento(hub)` + `hub_rfai(hub, IRC)` + `hub_capex(hub)`

Esta sheet explica o mecanismo NCRF 22 (reconhecimento subsídio) e CFI art. 22-23 (RFAI).

### Secção 6.1 — Pool de Depreciação (base do reconhecimento PT2030)

| Col | Label | Fonte | Fórmula validação |
|-----|-------|-------|-------------------|
| A | Ano | 2025–2034 | — |
| B | Dep. Pools Base (€) | `df_capex["dep_pools"]` | da sheet `CAPEX_DEP` col E |
| C | Dep. Juros Cap. (€) | `df_capex["dep_juros_cap"]` | da sheet `CAPEX_DEP` col F |
| D | Dep. Pools Líq. (€) | `=B-C` | dep sem juros cap (base elegível PT2030) |
| E | CAPEX Base Elegível (€) | 6 000 000 | `PRESSUPOSTOS!B3` |
| F | **Rácio Reconhecimento** | `=D/E` | `dep_pools_ano / capex_base` |
| G | **PT2030 Accrual (€)** | `=PT2030_MONTANTE * F` | `pt2030_reconhecimento(hub)[y]` |

**Validação célula G:** deve coincidir exactamente com `pt2030_map[y]`.

> **Porquê D (pools sem juros cap):** o motor usa `_dep_por_ano()` (pools base) e não `df_capex["depreciacao"]` (total). Os juros capitalizados são custo NCRF 10, não CAPEX elegível ao PT2030. A fórmula no motor é explícita:
> ```python
> return { y: montante * _dep_por_ano(proj, y) / capex_base for y in YEARS }
> ```

### Secção 6.2 — Balanço do Subsídio Diferido (conta 282)

| Col | Label | Fórmula |
|-----|-------|---------|
| H | Saldo Início Ano (€) | `=H_ano_anterior - G_ano_anterior` |
| I | Recebimento PT2030 (€) | só em 2027: `=PT2030_MONTANTE` |
| J | Reconhecido no Período (€) | `= G` (accrual do período) |
| K | **Saldo Fim Ano (€)** | `= H + I - J` |

Arranque:
- 2025: Saldo_início = 0, Recebimento = 0
- 2026: Saldo_início = 0, Recebimento = 0 (ainda não recebido)
- 2027: Saldo_início = valor_acumulado_não_reconhecido + **2 700 000 recebidos**
- 2028 em diante: sem novos recebimentos

### Secção 6.3 — RFAI (Crédito Fiscal CFI art. 22-23)

| Col | Label | Fonte | Fórmula |
|-----|-------|-------|---------|
| A | Ano | 2025–2029 | — |
| B | EBIT Hub Incremental (€) | `dr[y]["ebit_impact"]` | ref. `DR_INCR` |
| C | IRC Bruto (€) | `=MAX(B,0)*IRC_TAXA` | IRC incremental do hub |
| D | Limite 50% IRC (€) | `=C*0.50` | art. 23.º §6 CFI |
| E | Crédito Restante Ano Início | carry-forward acumulado | — |
| F | **RFAI Aplicado (€)** | `=MIN(E, D)` | `rfai_map[y]` |
| G | Crédito Restante Fim Ano | `=E-F` | — |

**Validação col F:** deve coincidir com `rfai_map[y]` para todos os anos.

Inicialização:
- E para 2025: `= RFAI_TAXA * RFAI_CAPEX_ELEGIVEL` → 10% × 6 000 000 = 600 000 €
- E para 2026: `= E_2025 - F_2025`
- etc.

---

## 7. Sheet `FCF`

**Fonte:** `df_fcf = hub_fcf(hub, IRC)` — DataFrame, uma linha por ano 2025-2029.

Esta é a sheet central para o VAL/TIR.

### Layout — linhas = rubricas, colunas B-F = anos 2025-2029

| Linha | Rubrica | Python (df_fcf) | Fórmula Excel | Sinal |
|-------|---------|----------------|---------------|-------|
| 5 | EBITDA Incremental Hub (€) | `"ebitda_impact"` | ref. `DR_INCR` L18 | + |
| 6 | (−) PT2030 Accrual — não-caixa (€) | `"pt2030_accrual"` | ref. `PT2030_RFAI` G | − |
| 7 | **EBIT para FCF** | `"ebit_fcf"` | `=L5-L6-depreciacao` | = EBIT sem subsídio |
| 8 | IRC s/ EBIT para FCF (€) | `=MAX(L7,0)*IRC_TAXA` | — | − |
| 9 | RFAI Crédito (€) | `"rfai_credito"` | ref. `PT2030_RFAI` F | + |
| 10 | **NOPAT Efectivo** | `"nopat"` | `=L7-L8+L9` | = |
| 11 | (+) Depreciação (€) | `"depreciacao"` | ref. `CAPEX_DEP` D | + |
| 12 | (−) CAPEX (€) | `"capex"` | ref. `CAPEX_DEP` B | − |
| 13 | (−) ΔNFM (€) | `"delta_nfm"` | nfm_map[y] | − se saída |
| 14 | (+) Libertação Inventário (€) | `"inventario_libertado"` | só 2026 | + |
| 15 | (−) Terreno Oportunidade (€) | `"terreno_oportunidade"` | só 2025 | − |
| 16 | (+) PT2030 Recebimento Caixa (€) | `"pt2030_cash"` | só 2027 | + |
| 17 | **FCF Livre (FCFF)** | `"fcf_livre"` | `=SUM(10:16)` | = |

**Validação obrigatória:** célula `FCF!L17` deve ser igual a `df_fcf.loc[df_fcf.ano==ano, "fcf_livre"].values[0]`.

### Linha de VAL, TIR, Payback (abaixo do FCF, ex. linhas 20-24)

| Linha | Métrica | Fórmula Excel | Python (via) |
|-------|---------|---------------|-------------|
| 20 | **VAL (€)** | `=NPV(WACC, FCF_2026:FCF_2034) + FCF_2025` | `via["val"]` |
| 21 | **TIR (%)** | `=IRR(FCF_2025:FCF_2034)` | `via["tir"]` |
| 22 | Payback Simples (anos) | calculado manualmente | `via["payback_simples"]` |
| 23 | Payback Atualizado (anos) | sobre FCFs descontados | `via["payback_atualizado"]` |
| 24 | Índice de Rendibilidade | `=1 + VAL/ABS(FCF_2025)` | `via["indice_rendibilidade"]` |

> **Nota Excel NPV:** em Excel, `NPV(rate, v1:vn)` desconta v1 em t=1 e vn em t=n.
> O FCF de 2025 ocorre em t=0 (desembolso inicial), por isso não entra no NPV() — é somado directamente.
> Fórmula correcta: `= NPV(WACC, FCF_2026, FCF_2027, …, FCF_2029) + FCF_2025`

---

## 8. Sheet `DSCR`

**Fonte:** `df_dscr = mapa_servico_divida(hub)`

### Layout — linhas = rubricas, colunas = anos 2025-2029

| Linha | Rubrica | Python (df_dscr) | Sinal/Formato |
|-------|---------|-----------------|---------------|
| 5 | EBITDA Hub Incremental (€) | `"ebitda_hub_incremental"` | + |
| 6 | Juros Pagos Total (€) | `"juros_pagos_total"` | − |
| 7 | Amortização Capital (€) | `"amortizacao_capital"` | − |
| 8 | **Serviço Total Dívida (€)** | `"servico_total_divida"` | = |
| 9 | **DSCR [x]** | `"dscr_hub"` | `=L5/L8` |
| 10 | Período Carência? | `"periodo_carencia"` | Texto |
| 11 | Saldo em Dívida Início (€) | `"saldo_em_divida"` | — |
| 12 | Saldo em Dívida Fim (€) | `"saldo_fim"` | — |

### Validação e sinalização

Após a tabela, adicione linha de semáforo para DSCR:
```excel
= IF(L9 >= 1.2, "✓ CUMPRE", IF(L9 >= 1.0, "⚠ MARGINAL", "✗ INCUMPRE"))
```

O DSCR mínimo aceitável é **1,20x** (covenant bancário).

Durante o período de carência (2025–2027), DSCR não é calculado — colocar "N/A (carência)".

---

## 9. Sheet `KPIS`

### Secção 9.1 — Resumo de Viabilidade

| Linha | Métrica | Valor Python | Fórmula alternativa |
|-------|---------|-------------|---------------------|
| 5 | VAL (€) | `via["val"]` | ref. `FCF!L20` |
| 6 | TIR (%) | `via["tir"]` | ref. `FCF!L21` |
| 7 | Payback Simples (anos) | `via["payback_simples"]` | — |
| 8 | Payback Atualizado (anos) | `via["payback_atualizado"]` | — |
| 9 | Índice Rendibilidade (x) | `via["indice_rendibilidade"]` | ref. `FCF!L24` |
| 10 | Valor Terminal — Ativos (€) | `via["valor_residual_ativos"]` | — |
| 11 | Valor Terminal — NFM (€) | `via["nfm_recovery_terminal"]` | — |
| 12 | Capital Vivo T10 (€) | `via["capital_vivo_t10"]` | — |
| 13 | WACC Usado (%) | `WACC` | `PRESSUPOSTOS!B15` |
| 14 | IRC Usado (%) | `IRC` | `PRESSUPOSTOS!B14` |

### Secção 9.2 — Análise de Sensibilidade (via["fcf_df"] não usado aqui)

Reproduz os resultados de `sensibilidade_hub()` para os 5 drivers do Tornado:

| Driver | Variação (−15%) | Base | Variação (+15%) |
|--------|----------------|------|----------------|
| Poupança Pessoal | ... | VAL_base | ... |
| CAPEX | ... | VAL_base | ... |
| WACC | ... | VAL_base | ... |
| Libertação Inventário | ... | VAL_base | ... |
| PT2030 Taxa | ... | VAL_base | ... |

> Estes valores saem de `via["tornado"]` se disponível, ou de chamadas manuais a `sensibilidade_hub(driver, [v], hub, IRC)`.

---

## 10. Checklist de Validação Final

Antes de dar o Excel por concluído, verificar **célula a célula** as 5 identidades abaixo:

| # | Check | Excel | Python |
|---|-------|-------|--------|
| 1 | PT2030 accrual 2026 | `PT2030_RFAI!G2` | `pt2030_map[2026]` |
| 2 | EBITDA Hub 2028 | `DR_INCR!E18` | `df_fcf.loc[df_fcf.ano==2028,"ebitda_impact"].values[0]` |
| 3 | FCF Livre 2027 | `FCF!D17` | `df_fcf.loc[df_fcf.ano==2027,"fcf_livre"].values[0]` |
| 4 | DSCR 2028 | `DSCR!E9` | `df_dscr.loc[df_dscr.ano==2028,"dscr_hub"].values[0]` |
| 5 | VAL base | `FCF!B20` | `via["val"]` |

Se todos os 5 coincidirem (±1 €, diferença por arredondamento), o Excel está coerente com o motor.

---

## 11. Extensão 2030-2034 (se necessário no Excel)

O `hub_fcf()` e `hub_dr_impact()` usam `YEARS` de `engine/inputs/constants.py`.

Para verificar se o motor já calcula 2030-2034:
```python
from engine.inputs.constants import YEARS
print(YEARS)  # se incluir 2030-2034, os dicts dr[] e fcf têm essas linhas
```

Se `YEARS = range(2025, 2030)` (só 5 anos), os valores 2030-2034 **não existem no motor** e o Excel não os pode derivar diretamente. Nesse caso:

- **Opção A (preferida):** alargar `YEARS` no `constants.py` para `range(2025, 2035)` e revalidar o motor.
- **Opção B (manual):** calcular 2030-2034 no Excel com as mesmas fórmulas, usando os parâmetros da sheet `PRESSUPOSTOS` — mas assinalar claramente que são valores calculados no Excel, não confirmados pelo motor.

Se usar Opção B, as depreciações 2030-2034 derivam dos pools:
- Construção (25 a): `=106_200` anual até 2050
- Honorários (25 a): incluídos na construção
- VLM (8 a, inicio 2026): termina 2033
- Robótica AMR (5 a, inicio 2026): termina 2030
- WMS (4 a, inicio 2026): termina 2029
- Box-on-Demand (5 a, inicio 2026): termina 2030
- Solar (20 a, inicio 2026): `=13_500` até 2045
- Software Integração (3 a, inicio 2026): termina 2028

Em 2031-2033, `dep_pools` = Construção + VLM + Solar = verificar com `df_capex`.

---

## 12. Convenções de Formatação

- Células de input (hardcoded): fundo **amarelo** `#FFFF00`
- Células derivadas do motor (coladas do Python): fundo **azul claro** `#D6E4F0`
- Células de fórmula Excel: fundo **branco**
- Células de totais/subtotais: negrito + fundo cinza `#F2F2F2`
- Valores negativos: formato `#,##0 €;[Red](#,##0 €)`
- DSCR < 1,20: fill **laranja** `#FCE4D6` via formatação condicional
- DSCR ≥ 1,20: fill **verde** `#E2EFDA` via formatação condicional

---

*Guia gerado em 2026-05-24 | GrestelPy v0.9.5 | Cenário Base + Hub*
