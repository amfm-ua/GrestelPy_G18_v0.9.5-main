Análise Crítica — GrestelPy M6



2. FSE cai -13% em 2025 com vendas a crescer +20%
FSE 2024 = 7,463 k€ → FSE 2025 = 6,489 k€ (-13,0%), enquanto as vendas crescem +20%. Em termos de proporcionalidade, uma empresa a crescer 20% teria normalmente mais subcontratos, energia, transportes — não menos. A racionalização do FSE precisa de um racional específico (p.ex.: internalização de um subcontrato, renegociação de contratos de energia).



4. Taxa IRC 24,5% vs. taxa efetiva histórica de 8,4%
O modelo usa IRC = 20% + 1,5% derrama municipal + 3% derrama estadual = 24,5% no hub. A derrama estadual de 3% só se aplica quando o lucro tributável supera 35M€. Mas a IRC pago pela Grestel em 2024 foi apenas 127 k€ sobre RAI de 1,517 M€ → taxa efetiva real = 8,4%. O modelo projetado usa 24,5% (taxa marginal máxima), o que é muito mais conservador, mas cria inconsistência: o RFAI (crédito de 595 k€) demora muitos anos a absorver porque o IRC base é alto. Se a taxa efetiva real for mais baixa, o RFAI absorve-se mais rapidamente e o VAL é diferente.

5. Rendimentos Financeiros Estáticos com Aplicações Crescentes
As aplicacoes_fin_cp crescem de 0 em 2025 para 15,806 k€ em 2029 (tesouraria excedente do modelo). Mas os rend_financeiros ficam congelados em 60 k€/ano em todo o período (valor de 2024). Com 15M€ aplicados a qualquer taxa de mercado, os rendimentos financeiros deviam crescer para 600–900 k€/ano. Esta inconsistência subtrai valor ao modelo e é empiricamente detetável.

6. PT2030 de 45% — Moda degenerada no Monte Carlo
No Monte Carlo, a distribuição do PT2030 é Triangular(min=20%, mode=45%, max=45%) — degenerada com mode = max. Isto significa que o cenário mais provável assumido é a aprovação máxima. Para um primeiro projeto com PT2030, a probabilidade-moda deveria ser mais baixa (30–35%). A assimetria atual favorece sistematicamente o cenário otimista.

7. Crescimento dos Benefícios do Hub a 3,5%/ano sem CAPEX adicional
O beneficio_liquido_anual (280 k€) cresce à taxa nominal de 3,5%/ano indefinidamente. Mas os benefícios derivam da automação (VLMs, AMRs) já instalada — um ativo físico que não cresce. A poupança de pessoal não cresce automaticamente só porque há inflação; as poupanças de automação tendem a estabilizar ou até reduzir-se com o envelhecimento do sistema. Esta premissa implica que em 2035 o benefício líquido anual seria ~400 k€, uma extrapolação otimista sem base física.

8. Libertação de Inventário como Benefício "do Projeto"
O modelo inclui €1,900 k€ de libertação de inventário como FCF positivo do Hub em 2026–2027. Esta libertação é um benefício de gestão de tesouraria, não de retorno do ativo Hub em si. Mais importante: numa perspetiva de FCFF incremental, a libertação de inventário é um benefício one-time de variação de NFM. Incluí-la no numerador do VAL é metodologicamente correto (variação de NFM é FCF), mas o benchmark de 14,5% de redução do stock está no limite superior do citado "10–15% para PMEs industriais" — o mínimo seria mais conservador.

9. FSE 2028–2029: Salto Abrupto de Custos
Spread real FSE: 2026: 1,0% → 2027: 1,2% → 2028: 3,2% → 2029: 3,3%. Um salto de 1,2% para 3,2% de spread real em 2028 não tem explicação explícita no YAML. O mesmo padrão ocorre no pessoal em 2029 (1,3% → 3,3%). Estes saltos criam um efeito de "cliff" nos custos em 2028–2029 que deteriora o EBITDA e deveria ser fundamentado.

10. Ausência de Dividendos — Acumulação de Capital
Os Resultados Transitados acumulam de 5,770 k€ (2024) para 28,033 k€ em 2029 sem qualquer distribuição de dividendos, com RL entre 5–9 M€/ano. Isto é financeiramente improvável para uma empresa desta dimensão e tradição. A política de dividendos deveria estar explicitada; a sua ausência sobrestima o capital próprio no balanço e distorce os rácios de rentabilidade.

11. Estimativa de Gás Natural (ESG) sem Fonte do Preço
O objetivo SMART de emissões usa: 447,486 € / 0,15 €/kWh = 2,983,000 kWh. O preço de 0,15 €/kWh para gás natural industrial em Portugal 2024 é questionável — as fontes do ERSE e Eurostat indicam ~0,07–0,10 €/kWh para uso industrial. Se o preço real for 0,08 €/kWh, o consumo seria 5,6M kWh (quase o dobro), alterando significativamente a intensidade energética calculada e a credibilidade dos objetivos ESG.

12. Euribor 3M corrigido de 2,65% para 2,90% — Justificação Pontual
O YAML de base histórica regista: Euribor_3M_base_2025: 0.0290 # BCE mar-2025; corrigido de 0.0265 (M3). Isto indica que os pressupostos foram ajustados pós-entrega de M3. Uma correção destas tem impacto no custo da dívida e nas demonstrações financeiras — o professor pode questionar se outros pressupostos de M3 foram igualmente revistos ou apenas este.




