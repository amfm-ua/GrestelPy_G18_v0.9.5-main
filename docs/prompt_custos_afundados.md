# Prompt — Implementação: Tratamento de Custos Afundados e Honorários de Exploração

## Contexto

O modelo M6 capitaliza honorários e licenças (85–165 k€) dentro do pool `construcao_civil`
em `m6_hub_assumptions.yaml`. Parte desses honorários pode corresponder a custos de
exploração/estudo já incorridos antes da decisão de investimento (sunk costs), que devem ser
excluídos da análise incremental por princípio financeiro — não pertencem ao CAPEX elegível
nem ao cálculo de NPV/VAL.

Atualmente não existe separação explícita entre:
- honorários futuros (comprometidos com a obra — CAPEX elegível, depreciáveis)
- honorários já pagos (afundados — irrelevantes para a decisão, não depreciáveis)

---

## Prompt de Implementação

Implementa o seguinte no engine GrestelPy:

### 1. YAML — `m6_hub_assumptions.yaml`

Separa o campo de honorários em dois sub-campos dentro de `projeto_hub.capex.pools.construcao_civil`
(ou como pool dedicado a par do `construcao_civil`):

```yaml
honorarios:
  futuros:
    descricao: "Honorários arquitetura, eng. de detalhe, licenças — ainda a incorrer"
    montante: 120000          # valor a calibrar — parte dos 85–165 k€ estimados
    ano_inicio: 2025
    taxa_depreciacao: 0.04    # mesmo pool que construção civil (NCRF 7)
    vida_util_anos: 25
    capex_elegivel_rfai: true
  afundados:
    descricao: "Estudos de viabilidade, consultoria técnica pré-decisão — já incorridos"
    montante: 45000           # valor a calibrar
    capex_elegivel_rfai: false
    excluir_analise_incremental: true
    nota: "Sunk cost — não entra nos pools de depreciação nem no FCF projetado"
```

### 2. Engine — leitura de pools (`dr.py` ou módulo de investimento)

Ao iterar `pools.values()` para calcular depreciações e CAPEX incremental, filtra pools ou
sub-campos com `excluir_analise_incremental: true`. Estes valores:

- **não** contribuem para a base depreciável
- **não** aparecem no mapa de cash outflows do hub
- **podem** ser lidos para nota de rodapé/auditoria nos outputs (ex: DR, notas ao balanço)

```python
# Exemplo de filtro no loader de pools
pools_ativos = {
    k: v for k, v in pools.items()
    if not v.get("excluir_analise_incremental", False)
}
```

### 3. Cenário "sem hub" (`incluir_hub: false`)

Confirmar que, quando `incluir_hub: false`, **nenhum** custo proveniente de
`projeto_hub` é carregado — incluindo honorários futuros. O cenário base não deve
carregar qualquer custo específico do hub, afundado ou não.

Adicionar assertion/guard no engine:

```python
if not incluir_hub:
    assert hub_capex_total == 0, "Cenário sem hub não deve conter CAPEX do hub"
```

### 4. Output — Nota de auditoria

No relatório de viabilidade do hub (outputs DR/NPV), adicionar linha de rodapé:

> "Custos afundados de exploração (X k€) excluídos da análise incremental por
> corresponderem a honorários incorridos antes da decisão de investimento (princípio
> dos cash flows incrementais — Brealey, Myers & Allen, Cap. 6)."

---

## Referências

- Princípio dos sunk costs: Brealey, Myers & Allen — *Principles of Corporate Finance*, Cap. 6
- NCRF 7 §7: reconhecimento de ativos fixos tangíveis (apenas custos futuros são capitalizáveis)
- CFI art. 22.º–23.º: RFAI incide sobre investimento realizado, não sobre estudos pré-decisão
- `m6_hub_assumptions.yaml` — pools atuais, linha 52–94
- `pressupostos/investimento.yaml` — estrutura de CAPEX global

---

## Critério de Aceitação

- [ ] `honorarios.afundados` não aparece na base depreciável nem no outflow de FCF
- [ ] `honorarios.futuros` é tratado como CAPEX normal (depreciado, elegível RFAI)
- [ ] Cenário `incluir_hub: false` não carrega qualquer custo do hub
- [ ] Nota de auditoria gerada no output de viabilidade quando `afundados.montante > 0`
- [ ] Soma `futuros + afundados` reconcilia com o intervalo 85–165 k€ documentado no YAML
