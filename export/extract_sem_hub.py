"""Exporta os dados do modelo SEM hub para todos os cenários → modelo_sem_hub.md.

Executa run_model(cenario=c, hub_on=False) para cada cenário do modelo
(Base / Upside / Downside), formata os KPIs anuais como tabelas Markdown
e acrescenta um quadro comparativo do ano-horizonte (2029).
"""

import sys
from pathlib import Path
from datetime import date

_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT))           # raiz do projeto → src.engine.* funciona
sys.path.insert(0, str(_ROOT / "src"))   # src/ → engine.* funciona

import pandas as pd
import numpy as np

from engine.modelo.model import run_model
from engine.inputs.loader import CENARIOS


# ── helpers ──────────────────────────────────────────────────────────────────

def _fmt(v, decimais=0):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "—"
    try:
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


# Classificação de colunas para formatação automática
_PCT_COLS = {"autonomia_financeira", "ebitda_margin"}
_RATIO_COLS = {"liquidez_geral", "debt_ebitda", "nd_ebitda",
               "cobertura_juros", "dscr"}
_DIAS_COLS = {"dmi_dias", "pmr_dias", "pmp_dias", "ciclo_caixa"}

# Colunas de KPIs apresentadas (pela ordem)
_KPI_COLS = ["ano", "vn", "ebitda", "ebit", "rl", "autonomia_financeira",
             "liquidez_geral", "dmi_dias", "pmr_dias", "pmp_dias",
             "ciclo_caixa", "debt_ebitda", "nd_ebitda", "cobertura_juros",
             "dscr", "total_ativo", "cp"]


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
            df[c] = df[c].apply(_pct)
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


def _kpis(result) -> pd.DataFrame | None:
    """Extrai a tabela de KPIs anuais de um resultado de run_model()."""
    kpis = result.get("kpis") if isinstance(result, dict) else None
    if kpis is None or kpis.empty:
        return None
    return kpis[[c for c in _KPI_COLS if c in kpis.columns]].copy()


def _comparativo(kpis_por_cenario: dict, ano: int) -> pd.DataFrame:
    """Quadro comparativo de um ano-alvo entre cenários."""
    cols = ["vn", "ebitda", "ebit", "rl", "autonomia_financeira",
            "total_ativo", "cp"]
    rows = []
    for nome, kpis in kpis_por_cenario.items():
        if kpis is None:
            continue
        sel = kpis[kpis["ano"] == ano]
        if sel.empty:
            continue
        r = sel.iloc[0]
        row = {"Cenário": nome}
        row.update({c: r[c] for c in cols if c in kpis.columns})
        rows.append(row)
    return pd.DataFrame(rows)


# ── main ─────────────────────────────────────────────────────────────────────

def main():
    kpis_por_cenario: dict = {}
    for nome in CENARIOS:
        print(f"A executar o modelo SEM hub — cenário {nome}…")
        try:
            result = run_model(cenario=nome, hub_on=False)
            kpis_por_cenario[nome] = _kpis(result)
        except Exception as exc:
            print(f"  erro: {exc}")
            kpis_por_cenario[nome] = None
    print("Modelos concluídos.")

    doc = []
    doc.append("# Modelo SEM Hub — Dados por Cenário\n\n")
    doc.append("**Empresa:** Costa Nova — Grestel, S.A.  \n")
    doc.append(f"**Cenários:** {', '.join(CENARIOS)}  \n")
    doc.append("**Configuração:** `hub_on=False` (modelo standalone, sem Hub Logístico)  \n")
    doc.append(f"**Gerado em:** {date.today().isoformat()}  \n")
    doc.append("**Período de projeção:** 2024–2029  \n\n")
    doc.append("---\n")

    for nome in CENARIOS:
        doc.append(f"\n## Cenário {nome} — KPIs Anuais (sem hub)\n\n")
        doc.append(_df_to_md(kpis_por_cenario.get(nome)))

    doc.append("\n## Comparação entre Cenários — Ano-Horizonte 2029 (sem hub)\n\n")
    comp = _comparativo(kpis_por_cenario, 2029)
    doc.append(_df_to_md(comp))
    doc.append("\n> Valores em euros; **autonomia_financeira** = CP / Ativo Total.\n")

    out_path = Path(__file__).parent / "modelo_sem_hub.md"
    out_path.write_text("".join(doc), encoding="utf-8")
    print(f"\nFicheiro guardado em: {out_path}")


if __name__ == "__main__":
    main()
