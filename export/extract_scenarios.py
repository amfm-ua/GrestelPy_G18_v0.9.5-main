"""Exporta os 4 cenários do Hub Logístico e os KPIs consolidados para cenarios.md.

Cobre:
  - Viabilidade do hub nos cenários Base / Upside / Downside / Stress
    (VAL, TIR, Payback atualizado, Índice de Rendibilidade);
  - KPIs anuais do modelo completo (hub_on=True);
  - Mapa de serviço da dívida do hub (extendido).
"""

import sys
import copy
from pathlib import Path
from datetime import date

_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT))           # raiz do projeto → src.engine.* funciona
sys.path.insert(0, str(_ROOT / "src"))   # src/ → engine.* funciona

import pandas as pd
import numpy as np

from engine.projetos.hub_logistico import load, viabilidade_hub, mapa_servico_divida
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


def _df_to_md(df: pd.DataFrame) -> str:
    """Converte DataFrame para tabela Markdown, formatando por tipo de coluna."""
    if df is None or df.empty:
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


# ── cenários ─────────────────────────────────────────────────────────────────

SCENARIOS = [
    ("Base", {}),
    ("Upside", {"mult_op": 1.2, "mult_vn": 1.2}),
    ("Downside", {"mult_op": 0.8, "mult_vn": 0.8}),
    ("Stress", {"mult_capex": 1.15, "mult_op": 0.7, "mult_vn": 0.6}),
]


def _viabilidade_cenarios(hub) -> pd.DataFrame:
    """Aplica os multiplicadores de cada cenário e calcula a viabilidade do hub."""
    rows = []
    for name, ov in SCENARIOS:
        h = copy.deepcopy(hub)
        b = h["projeto_hub"]["beneficios_anuais"]
        c = h["projeto_hub"]["beneficios_comerciais"]
        cp = h["projeto_hub"]["capex"]

        mo = ov.get("mult_op", 1.0)
        mv = ov.get("mult_vn", 1.0)
        mc = ov.get("mult_capex", 1.0)

        if mo != 1.0:
            b["poupanca_operacional"] = 480000 * mo
            b["reducao_quebras"] = 80000 * mo
            b["beneficio_liquido_anual"] = 350000 * mo
        if mv != 1.0:
            vi = c.get("vn_incremental", {})
            c["vn_incremental"] = {k: v * mv for k, v in vi.items()}
        if mc != 1.0:
            cp["base"] = 6000000 * mc
            cp["cronograma"] = {2025: 3000000 * mc, 2026: 3000000 * mc}

        res = viabilidade_hub(h)
        rows.append({
            "Cenário": name,
            "VAL (€)": _fmt(res.get("val")),
            "TIR": _pct(res.get("tir")),
            "Payback Atualizado (anos)": _fmt(res.get("payback_atualizado"), 2),
            "Índice de Rendibilidade": _fmt(res.get("indice_rendibilidade"), 4),
        })
    return pd.DataFrame(rows)


# ── main ─────────────────────────────────────────────────────────────────────

def main():
    hub = load()

    doc = []
    doc.append("# Cenários — Hub Logístico e KPIs Consolidados\n\n")
    doc.append("**Empresa:** Costa Nova — Grestel, S.A.  \n")
    doc.append(f"**Gerado em:** {date.today().isoformat()}  \n\n")
    doc.append("---\n")

    print("A avaliar os 4 cenários…")
    doc.append("\n## 4 Cenários — Viabilidade do Hub\n\n")
    doc.append(_df_to_md(_viabilidade_cenarios(hub)))
    doc.append("\n> **Base** sem multiplicadores · **Upside** +20% op/VN · "
               "**Downside** −20% op/VN · **Stress** +15% CAPEX, −30% op, −40% VN.\n")

    print("A executar o modelo completo (hub_on=True)…")
    doc.append("\n## Modelo Completo (hub_on=True) — KPIs 2024–2029\n\n")
    try:
        result = run_model(hub_on=True)
        kpis = result.get("kpis") if isinstance(result, dict) else None
        if kpis is not None and not kpis.empty:
            cols = ["ano", "vn", "ebitda", "ebit", "rl", "autonomia_financeira",
                    "liquidez_geral", "dmi_dias", "pmr_dias", "pmp_dias",
                    "ciclo_caixa", "debt_ebitda", "nd_ebitda", "cobertura_juros",
                    "dscr", "total_ativo", "cp"]
            avail = [c for c in cols if c in kpis.columns]
            doc.append(_df_to_md(kpis[avail]))
        else:
            doc.append("_KPIs não disponíveis._\n")
    except Exception as exc:
        import traceback
        doc.append(f"_Erro ao executar o modelo: {exc}_\n\n"
                   f"```\n{traceback.format_exc()}\n```\n")

    print("A construir o mapa de serviço da dívida…")
    doc.append("\n## Mapa de Serviço da Dívida (extendido)\n\n")
    try:
        ds = mapa_servico_divida(hub)
        doc.append(_df_to_md(ds))
    except Exception as exc:
        doc.append(f"_Mapa de serviço da dívida não disponível: {exc}_\n")

    out_path = Path(__file__).parent / "cenarios.md"
    out_path.write_text("".join(doc), encoding="utf-8")
    print(f"\nFicheiro guardado em: {out_path}")


if __name__ == "__main__":
    main()
