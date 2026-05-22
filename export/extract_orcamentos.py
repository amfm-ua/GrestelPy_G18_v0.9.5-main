"""Exporta Secção 6 — Orçamentos Operacionais e Financeiros para orcamentos_secc6.md.

Cobre as subsecções 6.1 a 6.15 conforme o índice do relatório M6/G18.
Executa run_model() e formata cada orçamento como tabela Markdown.
"""

import sys
import io
from pathlib import Path
from datetime import date

_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT))           # project root → src.engine.* works
sys.path.insert(0, str(_ROOT / "src"))   # src/ → engine.* works (existing scripts)

import pandas as pd
import numpy as np

pd.set_option("display.float_format", "{:,.0f}".format)
pd.set_option("display.max_columns", None)
pd.set_option("display.width", 300)


# ── helpers ──────────────────────────────────────────────────────────────────

def _fmt(v, decimais=0):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "—"
    try:
        if decimais == 0:
            return f"{float(v):,.0f}"
        return f"{float(v):,.{decimais}f}"
    except (TypeError, ValueError):
        return str(v)


def _pct(v, decimais=1):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "—"
    try:
        return f"{float(v)*100:.{decimais}f}%"
    except (TypeError, ValueError):
        return str(v)


def _df_to_md(df: pd.DataFrame, float_cols: list[str] | None = None) -> str:
    """Converte DataFrame para tabela Markdown simples."""
    if df is None or df.empty:
        return "_Sem dados disponíveis._\n"

    df = df.copy()
    skip = {"ano", "year", "mes"}
    if float_cols:
        for c in float_cols:
            if c in df.columns:
                df[c] = df[c].apply(lambda x: _fmt(x) if pd.notna(x) else "—")
    else:
        for c in df.select_dtypes(include="number").columns:
            if c in skip:
                df[c] = df[c].apply(lambda x: str(int(x)) if pd.notna(x) else "—")
            else:
                df[c] = df[c].apply(lambda x: _fmt(x) if pd.notna(x) else "—")

    header = "| " + " | ".join(str(c) for c in df.columns) + " |"
    sep = "| " + " | ".join(["---"] * len(df.columns)) + " |"
    rows = []
    for _, row in df.iterrows():
        rows.append("| " + " | ".join(str(v) for v in row.values) + " |")
    return "\n".join([header, sep] + rows) + "\n"


def _pivot_mensal(df: pd.DataFrame, index_col: str, value_col: str,
                  mes_col: str = "mes") -> pd.DataFrame:
    """Pivot mensal: linhas = rubrica, colunas = 12 meses + Total."""
    from engine.inputs import MESES
    piv = df.pivot_table(index=index_col, columns=mes_col,
                         values=value_col, aggfunc="sum", fill_value=0)
    # garantir ordem e presença de todos os meses
    for m in MESES:
        if m not in piv.columns:
            piv[m] = 0.0
    piv = piv[MESES]
    piv["Total"] = piv.sum(axis=1)
    return piv.reset_index()


def _section(num: str, title: str, body: str) -> str:
    return f"\n## {num} {title}\n\n{body}\n"


# ── modelo ───────────────────────────────────────────────────────────────────

def _run():
    from engine.modelo.model import run_model
    from engine.inputs import load, MESES

    print("A executar o modelo…")
    dfs = run_model(cenario="Base", hub_on=False, ecogres_on=False)
    a, base, sched = load(cenario="Base")
    print("Modelo concluído.")
    return dfs, a, base, sched, MESES


# ── secções ──────────────────────────────────────────────────────────────────

def s61_vendas(dfs, a, base, sched, MESES):
    lines = []

    lines.append("**Pressupostos:** ver Secção 4.3.\n")

    # Vendas anuais por produto
    df_p = dfs.get("vendas_produto_anual")
    if df_p is not None and not df_p.empty:
        lines.append("### Vendas por Produto (anual, €)\n")
        cols_show = [c for c in ["ano", "produto", "qty_vendida", "pvu", "vn_prod"]
                     if c in df_p.columns]
        lines.append(_df_to_md(df_p[cols_show]))

    # Vendas anuais por mercado
    df_m = dfs.get("vendas_mercado_anual")
    if df_m is not None and not df_m.empty:
        lines.append("### Vendas por Mercado (anual, €)\n")
        cols_show = [c for c in ["ano", "mercado", "vn"] if c in df_m.columns]
        lines.append(_df_to_md(df_m[cols_show]))

    # Mapa mensal 2025
    df_men = dfs.get("vendas_mensal_2025")
    if df_men is not None and not df_men.empty:
        lines.append("### Mapa Orçamental Mensal 2025 — Vendas por Produto (€)\n")
        try:
            piv = _pivot_mensal(df_men, index_col="produto", value_col="vn")
            lines.append(_df_to_md(piv))
        except Exception as e:
            lines.append(f"_Erro ao pivotar: {e}_\n")

    # PMR dos KPIs
    kpis = dfs.get("kpis")
    if kpis is not None and not kpis.empty and "pmr_dias" in kpis.columns:
        lines.append("### PMR — Prazo Médio de Recebimento (dias)\n")
        cols_show = [c for c in ["ano", "pmr_dias"] if c in kpis.columns]
        lines.append(_df_to_md(kpis[cols_show]))

    # DR mensal — VN linha
    dr_men = dfs.get("dr_mensal_2025")
    if dr_men is not None and not dr_men.empty and "vn" in dr_men.columns:
        lines.append("### Vendas Mensais 2025 — DR (€)\n")
        row = dr_men[["mes", "vn"]].copy() if "mes" in dr_men.columns else dr_men[["vn"]].copy()
        lines.append(_df_to_md(row))

    return "".join(lines)


def s62_producao(dfs, a, base, sched, MESES):
    lines = []

    df = dfs.get("producao_anual")
    if df is not None and not df.empty:
        lines.append("### Orçamento de Produção — Anual (€ e unidades)\n")
        cols_show = [c for c in [
            "ano", "produto", "qty_vendida", "qty_produzida",
            "cup", "cip_unitario", "cmvmc_vendas", "cmvmc_prod",
            "pa_stock_ei", "pa_stock_ef", "var_pa",
        ] if c in df.columns]
        lines.append(_df_to_md(df[cols_show]))

    df_men = dfs.get("producao_mensal_2025")
    if df_men is not None and not df_men.empty:
        lines.append("### Produção Mensal 2025 — Custo Industrial por Produto (€)\n")
        try:
            piv = _pivot_mensal(df_men, index_col="produto", value_col="cmvmc_prod")
            lines.append(_df_to_md(piv))
        except Exception as e:
            lines.append(f"_Erro ao pivotar: {e}_\n")

    return "".join(lines)


def s63_pessoal(dfs, a, base, sched, MESES):
    lines = []

    df_anual = dfs.get("pessoal_anual")
    if df_anual is not None and not df_anual.empty:
        lines.append("### Gastos com Pessoal — Totais Anuais (€)\n")
        cols_show = [c for c in [
            "ano", "gastos_pessoal", "headcount", "custo_medio", "peso_vn_pct",
        ] if c in df_anual.columns]
        lines.append(_df_to_md(df_anual[cols_show]))

    df_depart = dfs.get("pessoal_depart_anual")
    if df_depart is not None and not df_depart.empty:
        lines.append("### Gastos com Pessoal por Departamento (€)\n")
        lines.append(_df_to_md(df_depart))

    df_contab = dfs.get("pessoal_contab_anual")
    if df_contab is not None and not df_contab.empty:
        lines.append("### Decomposição Contabilística (Remunerações, TSU, Outros) — €\n")
        lines.append(_df_to_md(df_contab))

    df_men = dfs.get("pessoal_mensal_2025")
    if df_men is not None and not df_men.empty:
        lines.append("### Gastos com Pessoal — Mensal 2025 (€)\n")
        lines.append(_df_to_md(df_men))

    return "".join(lines)


def s64_depreciações(dfs, a, base, sched, MESES):
    lines = []

    # D&A a partir da DR anual
    dr = dfs.get("dr")
    if dr is not None and not dr.empty and "depreciacoes" in dr.columns:
        lines.append("### Depreciações e Amortizações — Totais Anuais (€)\n")
        cols_show = [c for c in ["ano", "depreciacoes"] if c in dr.columns]
        df_da = dr[cols_show].copy()
        df_da["depreciacoes"] = df_da["depreciacoes"].abs()
        lines.append(_df_to_md(df_da))

    # D&A mensal 2025 a partir da DR mensal
    dr_men = dfs.get("dr_mensal_2025")
    if dr_men is not None and not dr_men.empty and "depreciacoes" in dr_men.columns:
        lines.append("### D&A Mensal 2025 (€)\n")
        df_m = dr_men[["mes", "depreciacoes"]].copy() if "mes" in dr_men.columns else dr_men[["depreciacoes"]].copy()
        df_m["depreciacoes"] = df_m["depreciacoes"].abs()
        lines.append(_df_to_md(df_m))

    lines.append(
        "\n> **Método:** Linear. Taxas definidas em `schedules.yaml` "
        "(taxa_dep_aft, taxa_amort_intang). Novos ativos CAPEX 2025 "
        "identificados na Secção 6.8.\n"
    )

    return "".join(lines)


def s65_fse(dfs, a, base, sched, MESES):
    lines = []

    df_anual = dfs.get("fse_detalhe_anual")
    if df_anual is not None and not df_anual.empty:
        lines.append("### FSE por Natureza — Anual (€)\n")
        lines.append(_df_to_md(df_anual))

    fse_men = dfs.get("fse_detalhe_mensal_2025")
    if isinstance(fse_men, dict) and fse_men:
        lines.append("### FSE Mensal 2025 por Rubrica (€)\n")
        rows = []
        for rubrica, mapa in fse_men.items():
            row = {"Rubrica": rubrica}
            total = 0.0
            for m in MESES:
                v = float(mapa.get(m, 0.0))
                row[m] = _fmt(v)
                total += v
            row["Total"] = _fmt(total)
            rows.append(row)
        if rows:
            df_fse_md = pd.DataFrame(rows)
            lines.append(_df_to_md(df_fse_md, float_cols=[]))

    # PMP dos KPIs
    kpis = dfs.get("kpis")
    if kpis is not None and not kpis.empty and "pmp_dias" in kpis.columns:
        lines.append("### PMP — Prazo Médio de Pagamento (dias)\n")
        cols_show = [c for c in ["ano", "pmp_dias"] if c in kpis.columns]
        lines.append(_df_to_md(kpis[cols_show]))

    return "".join(lines)


def s66_cmvmc(dfs, a, base, sched, MESES):
    lines = []

    df = dfs.get("producao_anual")
    if df is not None and not df.empty:
        lines.append("### CMVMC por Produto — Anual (€)\n")
        cols_show = [c for c in [
            "ano", "produto", "qty_produzida", "cup", "cmvmc_prod",
            "pa_stock_ei", "pa_stock_ef",
        ] if c in df.columns]
        lines.append(_df_to_md(df[cols_show]))

    df_men = dfs.get("cmvmc_mensal_2025")
    if df_men is not None and not df_men.empty:
        lines.append("### CMVMC Mensal 2025 (€)\n")
        lines.append(_df_to_md(df_men))

    return "".join(lines)


def s67_outros(dfs, a, base, sched, MESES):
    lines = []

    dr = dfs.get("dr")
    if dr is not None and not dr.empty:
        cands = [c for c in [
            "ano", "outros_rendimentos", "subsidios_exploracao",
            "outros_gastos", "rendimentos_financeiros",
        ] if c in dr.columns]
        if len(cands) > 1:
            lines.append("### Outros Rendimentos e Gastos Operacionais — Anuais (€)\n")
            lines.append(_df_to_md(dr[cands]))

    return "".join(lines)


def s68_investimento(dfs, a, base, sched, MESES):
    lines = []

    try:
        from engine.investimento.investimento import investimento_anual
        df_inv = investimento_anual(a, base, sched)
        if df_inv is not None and not df_inv.empty:
            lines.append("### Plano de Investimentos Anual (€)\n")
            lines.append(_df_to_md(df_inv))
    except Exception as e:
        lines.append(f"_Dados de investimento não disponíveis: {e}_\n")

    balanco = dfs.get("balanco")
    if balanco is not None and not balanco.empty:
        af_cols = [c for c in ["ano", "aft_liquido", "intangiveis", "total_ativo_nao_corrente"]
                   if c in balanco.columns]
        if af_cols:
            lines.append("### Ativos Fixos no Balanço — Anual (€)\n")
            lines.append(_df_to_md(balanco[af_cols]))

    lines.append(
        "\n> **CAPEX 2025:** eficiência energética (+200 kW solar), "
        "otimização de fornos, digitalização de processos.\n"
    )

    return "".join(lines)


def s69_nfm(dfs, a, base, _sched, MESES):
    lines = []

    try:
        from engine.demonstracoes.nfm import nfm_anual
        df_nfm = nfm_anual(a, base, dfs.get("balanco"), dfs.get("dr"))
        if df_nfm is not None and not df_nfm.empty:
            lines.append("### NFM Operacional — Anual (€)\n")
            lines.append(_df_to_md(df_nfm))
    except Exception as e:
        lines.append(f"_NFM calculado a partir do balanço: {e}_\n")

        balanco = dfs.get("balanco")
        if balanco is not None and not balanco.empty:
            nfm_cols = [c for c in ["ano", "clientes", "inventarios", "fornecedores"]
                        if c in balanco.columns]
            if nfm_cols:
                df_b = balanco[nfm_cols].copy()
                if all(c in df_b.columns for c in ["clientes", "inventarios", "fornecedores"]):
                    df_b["NFM"] = df_b["clientes"] + df_b["inventarios"] - df_b["fornecedores"]
                lines.append(_df_to_md(df_b))

    kpis = dfs.get("kpis")
    if kpis is not None and not kpis.empty:
        cycle_cols = [c for c in ["ano", "pmr_dias", "dmi_dias", "pmp_dias", "ciclo_caixa"]
                      if c in kpis.columns]
        if cycle_cols:
            lines.append("### Rácios do Ciclo Operacional (dias)\n")
            lines.append(_df_to_md(kpis[cycle_cols]))

    lines.append(
        "\n> **PMR-alvo:** 60 dias (Objetivo SMART 6).  "
        "**PMP-alvo:** 55 dias em dezembro (Objetivo SMART 6).\n"
    )

    return "".join(lines)


def s610_aplicacao_resultados(dfs, a, base, sched, MESES):
    lines = []

    dr = dfs.get("dr")
    if dr is not None and not dr.empty and "rl" in dr.columns:
        df_2024 = dr[dr.ano == 2024][["ano", "rl"]] if "ano" in dr.columns else dr[["rl"]].head(1)
        if not df_2024.empty:
            rl_2024 = float(df_2024["rl"].iloc[0])
            reserva_legal = max(0.0, rl_2024 * 0.05)
            lines.append("### Resultado Líquido 2024 e Proposta de Aplicação (€)\n\n")
            lines.append(f"| Rubrica | Valor (€) |\n")
            lines.append(f"|---|---|\n")
            lines.append(f"| Resultado Líquido 2024 | {_fmt(rl_2024)} |\n")
            lines.append(f"| Reserva Legal (5%) | {_fmt(reserva_legal)} |\n")
            lines.append(f"| Reservas Livres / Dividendos | {_fmt(rl_2024 - reserva_legal)} |\n\n")

    lines.append(
        "> **Fundamentação:** política de retenção prudente dado o baixo EBITDA "
        "registado em 2024 (2,3%) e a necessidade de reforço da autonomia financeira.\n"
    )

    return "".join(lines)


def s611_tesouraria(dfs, a, base, sched, MESES):
    lines = []

    df_teso = dfs.get("tesouraria_mensal_2025")
    if df_teso is not None and not df_teso.empty:
        lines.append("### Orçamento de Tesouraria Mensal 2025 (€)\n")
        lines.append(_df_to_md(df_teso))

    dr_men = dfs.get("dr_mensal_2025")
    if dr_men is not None and not dr_men.empty:
        lines.append("### DR Mensal 2025 — Resumo (€)\n")
        lines.append(_df_to_md(dr_men))

    return "".join(lines)


def s612_financiamento(dfs, a, base, sched, MESES):
    lines = []

    dr = dfs.get("dr")
    balanco = dfs.get("balanco")

    if dr is not None and not dr.empty:
        fin_cols = [c for c in ["ano", "juros", "resultado_financeiro", "rl"]
                    if c in dr.columns]
        if fin_cols:
            lines.append("### Carga Financeira — Anuais (€)\n")
            df_f = dr[fin_cols].copy()
            for c in ["juros", "resultado_financeiro"]:
                if c in df_f.columns:
                    df_f[c] = df_f[c].abs()
            lines.append(_df_to_md(df_f))

    if balanco is not None and not balanco.empty:
        div_cols = [c for c in ["ano", "divida_ml", "divida_cp", "total_passivo_nao_corrente"]
                    if c in balanco.columns]
        if div_cols:
            lines.append("### Dívida Financeira no Balanço (€)\n")
            lines.append(_df_to_md(balanco[div_cols]))

    return "".join(lines)


def s613_servico_divida(dfs, a, base, sched, MESES):
    lines = []

    try:
        from engine.financiamento.financiamento import financiamento_anual
        df_sd = financiamento_anual(sched, a)
        if df_sd is not None and not df_sd.empty:
            lines.append("### Mapa de Serviço da Dívida (€)\n")
            lines.append(_df_to_md(df_sd))
    except Exception as e:
        lines.append(f"_Função servico_divida_anual não disponível: {e}_\n")

        # Fallback: reconstruir a partir da DR e Balanço
        dr = dfs.get("dr")
        balanco = dfs.get("balanco")
        if dr is not None and balanco is not None and not dr.empty and not balanco.empty:
            juros_col = [c for c in ["juros"] if c in dr.columns]
            div_cols = [c for c in ["divida_ml", "divida_cp"] if c in balanco.columns]
            if juros_col and div_cols:
                df_sd_fb = dr[["ano"] + juros_col].merge(
                    balanco[["ano"] + div_cols], on="ano", how="left"
                )
                df_sd_fb["juros"] = df_sd_fb["juros"].abs()
                lines.append("### Serviço da Dívida — Estimativa (DR + Balanço) (€)\n")
                lines.append(_df_to_md(df_sd_fb))

    kpis = dfs.get("kpis")
    if kpis is not None and not kpis.empty:
        debt_cols = [c for c in ["ano", "dscr", "debt_ebitda", "nd_ebitda", "cobertura_juros"]
                     if c in kpis.columns]
        if debt_cols:
            lines.append("### Rácios de Cobertura da Dívida\n")
            df_ratios = kpis[debt_cols].copy()
            for c in ["dscr", "debt_ebitda", "nd_ebitda", "cobertura_juros"]:
                if c in df_ratios.columns:
                    df_ratios[c] = df_ratios[c].apply(
                        lambda x: f"{x:.2f}x" if pd.notna(x) else "—"
                    )
            lines.append(_df_to_md(df_ratios, float_cols=[]))

    return "".join(lines)


def s614_eoep(dfs, a, base, sched, MESES):
    lines = []

    df_eoep = dfs.get("eoep_mensal_2025")
    if df_eoep is not None and not df_eoep.empty:
        lines.append("### EOEP — Calendário Fiscal Mensal 2025 (€)\n")
        lines.append(_df_to_md(df_eoep))
    else:
        lines.append("_Mapa EOEP mensal não disponível; ver `engine/modelo/eoep.py`._\n")

    lines.append(
        "\n> **IRC:** pagamentos por conta (Jul, Set, Dez — 76,5% × IRC 2024 ÷ 3).  \n"
        "> **IVA:** regime normal mensal (CIVA art. 27, liquidado M+2).  \n"
        "> **Segurança Social:** TSU patronal, paga no mês M+1.\n"
    )

    return "".join(lines)


def s615_outros_orcamentos(dfs, a, base, sched, MESES):
    lines = []

    dr = dfs.get("dr")
    if dr is not None and not dr.empty:
        outros_cols = [c for c in [
            "ano", "outros_gastos", "outros_rendimentos",
        ] if c in dr.columns]
        if outros_cols:
            lines.append("### Outros Gastos / Rendimentos na DR — Anuais (€)\n")
            lines.append(_df_to_md(dr[["ano"] + [c for c in outros_cols if c != "ano"]]))

    lines.append(
        "\n> Inclui:\n"
        "> - **I&D:** design de novas coleções e tecnologia laser (cf. M1).\n"
        "> - **Qualidade & Certificações:** ISO 14001:2015.\n"
        "> - **Marketing & Feiras:** Ambiente Frankfurt, Maison & Objet.\n"
    )

    return "".join(lines)


# ── main ─────────────────────────────────────────────────────────────────────

def main():
    dfs, a, base, sched, MESES = _run()

    secoes = [
        ("6.1", "Orçamento de Vendas", s61_vendas),
        ("6.2", "Orçamento de Produção", s62_producao),
        ("6.3", "Orçamento de Gastos com Pessoal", s63_pessoal),
        ("6.4", "Mapa de Depreciações e Amortizações", s64_depreciações),
        ("6.5", "Orçamento de Fornecimentos e Serviços Externos (FSE)", s65_fse),
        ("6.6", "Orçamento de CMVMC", s66_cmvmc),
        ("6.7", "Outros Rendimentos e Gastos Operacionais", s67_outros),
        ("6.8", "Plano de Investimentos", s68_investimento),
        ("6.9", "Necessidades de Fundo de Maneio (NFM)", s69_nfm),
        ("6.10", "Proposta de Aplicação de Resultados", s610_aplicacao_resultados),
        ("6.11", "Orçamento de Tesouraria", s611_tesouraria),
        ("6.12", "Orçamento Financeiro", s612_financiamento),
        ("6.13", "Mapas de Serviço da Dívida", s613_servico_divida),
        ("6.14", "Orçamento e Calendarização Fiscal (EOEP)", s614_eoep),
        ("6.15", "Outros Orçamentos Relevantes", s615_outros_orcamentos),
    ]

    doc = io.StringIO()
    doc.write(f"# Secção 6 — Orçamentos Operacionais e Financeiros\n\n")
    doc.write(f"**Empresa:** Costa Nova — Grestel, S.A.  \n")
    doc.write(f"**Cenário:** Base  \n")
    doc.write(f"**Gerado em:** {date.today().isoformat()}  \n")
    doc.write(f"**Período de projeção:** 2024–2029 (2025: Jan–Set, 9 meses)  \n\n")
    doc.write("---\n")

    for num, title, fn in secoes:
        print(f"  Secção {num}…", end=" ", flush=True)
        try:
            body = fn(dfs, a, base, sched, MESES)
        except Exception as exc:
            import traceback
            body = f"_Erro ao gerar secção: {exc}_\n\n```\n{traceback.format_exc()}\n```\n"
        doc.write(_section(num, title, body))
        print("ok")

    out_path = Path(__file__).parent / "orcamentos_secc6.md"
    out_path.write_text(doc.getvalue(), encoding="utf-8")
    print(f"\nFicheiro guardado em: {out_path}")


if __name__ == "__main__":
    main()
