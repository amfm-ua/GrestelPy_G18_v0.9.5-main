"""Exporta os dados do Hub Logístico (M6/OE4) para hub_dados.md.

Cobre viabilidade, FCF unlevered, mapa de serviço da dívida, impacto anual
no DR, análise tornado e o modelo completo com hub (KPIs, DR, Balanço).
"""

import sys
from pathlib import Path
from datetime import date

_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT))           # raiz do projeto → src.engine.* funciona
sys.path.insert(0, str(_ROOT / "src"))   # src/ → engine.* funciona

import pandas as pd
import numpy as np

from engine.projetos.hub_logistico import (
    load, viabilidade_hub, mapa_servico_divida, hub_dr_impact, tornado_hub
)
from engine.modelo.model import run_model


# ── helpers ──────────────────────────────────────────────────────────────────

def _fmt(v, decimais=0):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "—"
    try:
        return f"{float(v):,.{decimais}f}"
    except (TypeError, ValueError):
        return str(v)


def _pct(v, decimais=2):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "—"
    try:
        return f"{float(v)*100:.{decimais}f}%"
    except (TypeError, ValueError):
        return str(v)


# Classificação de colunas para formatação automática
_PCT_COLS = {"autonomia_financeira"}
_RATIO_COLS = {"liquidez_geral", "debt_ebitda", "nd_ebitda",
               "cobertura_juros", "dscr", "dscr_hub"}
_DIAS_COLS = {"dmi_dias", "pmr_dias", "pmp_dias", "ciclo_caixa"}

# Subconjunto de KPIs apresentado (evita despejar rácios sem formatação própria)
_KPI_COLS = ["ano", "vn", "ebitda", "ebit", "rl", "autonomia_financeira",
             "liquidez_geral", "dmi_dias", "pmr_dias", "pmp_dias",
             "ciclo_caixa", "debt_ebitda", "nd_ebitda", "cobertura_juros",
             "dscr", "total_ativo", "cp"]


def _df_to_md(df: pd.DataFrame) -> str:
    """Converte DataFrame para tabela Markdown, formatando por tipo de coluna."""
    if df is None or getattr(df, "empty", True):
        return "_Sem dados disponíveis._\n"

    df = df.copy()
    for c in df.columns:
        if c in {"ano", "year"}:
            df[c] = df[c].apply(lambda x: str(int(x)) if pd.notna(x) else "—")
        elif pd.api.types.is_bool_dtype(df[c]):
            df[c] = df[c].apply(lambda x: "Sim" if bool(x) else "Não")
        elif c in _PCT_COLS:
            df[c] = df[c].apply(lambda x: _pct(x, 1))
        elif c in _RATIO_COLS:
            df[c] = df[c].apply(lambda x: _fmt(x, 2))
        elif c in _DIAS_COLS:
            df[c] = df[c].apply(lambda x: _fmt(x, 0))
        elif pd.api.types.is_numeric_dtype(df[c]):
            df[c] = df[c].apply(lambda x: _fmt(x) if pd.notna(x) else "—")

    header = "| " + " | ".join(str(c) for c in df.columns) + " |"
    sep = "| " + " | ".join(["---"] * len(df.columns)) + " |"
    rows = ["| " + " | ".join(str(v) for v in r.values) + " |"
            for _, r in df.iterrows()]
    return "\n".join([header, sep] + rows) + "\n"


def _kv_table(pares: list[tuple[str, str]]) -> str:
    """Tabela Markdown de duas colunas (Indicador | Valor)."""
    out = ["| Indicador | Valor |", "| --- | --- |"]
    out += [f"| {k} | {v} |" for k, v in pares]
    return "\n".join(out) + "\n"


# ── secções ──────────────────────────────────────────────────────────────────

def _viabilidade(res: dict) -> str:
    pares = [
        ("VAL (WACC 8%, 10 anos)", _fmt(res.get("val")) + " €"),
        ("TIR", _pct(res.get("tir"))),
        ("Payback simples", _fmt(res.get("payback_simples"), 2) + " anos"),
        ("Payback atualizado", _fmt(res.get("payback_atualizado"), 2) + " anos"),
        ("Índice de Rendibilidade", _fmt(res.get("indice_rendibilidade"), 4)),
        ("Valor terminal", _fmt(res.get("valor_terminal")) + " €"),
        ("Valor residual de ativos", _fmt(res.get("valor_residual_ativos")) + " €"),
    ]
    body = _kv_table(pares)

    params = res.get("parametros", {})
    if isinstance(params, dict) and params:
        body += "\n**Parâmetros:**\n\n"
        body += "".join(f"- `{k}`: {v}\n" for k, v in params.items())
    return body


def _dr_impact_df(dr_imp) -> pd.DataFrame:
    """Converte o impacto anual no DR (dict[int, dict]) num DataFrame."""
    if isinstance(dr_imp, dict):
        df = pd.DataFrame.from_dict(dr_imp, orient="index").reset_index()
        return df.rename(columns={"index": "ano"})
    return dr_imp


# ── main ─────────────────────────────────────────────────────────────────────

def main():
    hub = load()

    doc = []
    doc.append("# Hub Logístico 4.0 — Dados do Projeto\n\n")
    doc.append("**Empresa:** Costa Nova — Grestel, S.A.  \n")
    doc.append(f"**Gerado em:** {date.today().isoformat()}  \n\n")
    doc.append("---\n")

    print("A calcular a viabilidade do hub…")
    doc.append("\n## Viabilidade do Hub\n\n")
    res = viabilidade_hub(hub)
    doc.append(_viabilidade(res))

    doc.append("\n## FCF Unlevered (FCFF)\n\n")
    doc.append(_df_to_md(res.get("fcf_df")))

    print("A construir o mapa de serviço da dívida…")
    doc.append("\n## Mapa de Serviço da Dívida\n\n")
    try:
        doc.append(_df_to_md(mapa_servico_divida(hub)))
    except Exception as exc:
        doc.append(f"_Não disponível: {exc}_\n")

    print("A calcular o impacto anual no DR…")
    doc.append("\n## Impacto Anual no DR\n\n")
    try:
        doc.append(_df_to_md(_dr_impact_df(hub_dr_impact(hub))))
    except Exception as exc:
        doc.append(f"_Não disponível: {exc}_\n")

    print("A calcular o tornado…")
    doc.append("\n## Análise Tornado (Top 5)\n\n")
    try:
        tornado = tornado_hub()
        cols = ["label", "driver", "val_low", "val_base", "val_high",
                "impacto_total"]
        avail = [c for c in cols if c in tornado.columns]
        doc.append(_df_to_md(tornado[avail]))
    except Exception as exc:
        doc.append(f"_Não disponível: {exc}_\n")

    print("A executar o modelo completo (hub_on=True)…")
    doc.append("\n## Modelo Completo (hub_on=True)\n\n")
    try:
        result = run_model(hub_on=True)
        if isinstance(result, dict):
            doc.append("### KPIs 2024–2029\n\n")
            kpis = result.get("kpis")
            if kpis is not None and not kpis.empty:
                kpis = kpis[[c for c in _KPI_COLS if c in kpis.columns]]
            doc.append(_df_to_md(kpis))
            doc.append("\n### Demonstração de Resultados\n\n")
            doc.append(_df_to_md(result.get("dr")))
            doc.append("\n### Balanço\n\n")
            doc.append(_df_to_md(result.get("balanco")))
    except Exception as exc:
        import traceback
        doc.append(f"_Erro ao executar o modelo: {exc}_\n\n"
                   f"```\n{traceback.format_exc()}\n```\n")

    out_path = Path(__file__).parent / "hub_dados.md"
    out_path.write_text("".join(doc), encoding="utf-8")
    print(f"\nFicheiro guardado em: {out_path}")


if __name__ == "__main__":
    main()
