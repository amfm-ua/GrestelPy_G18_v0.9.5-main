from __future__ import annotations

from typing import Optional, Dict, Any

import pandas as pd

from ...inputs import Assumptions, Base2024, Schedules
from ...financiamento import tesouraria as teso_mod
from .integrado import _build_integrated_monthly
from .mensais import build_nfm_mensal, build_tesouraria_completa
from .reconciliacao import build_reconciliacao_mensal_anual, _overlay_dez_mensal_no_anual


# ──────────────────────────────────────────────────────────────────────────────
# Ponto de Entrada
# ──────────────────────────────────────────────────────────────────────────────

def build_rolling_forecast(
    a: Assumptions,
    base: Base2024,
    sched: Schedules,
    ov: Optional[Dict[str, Any]] = None,
) -> dict:
    """Constrói todas as demonstrações mensais articuladas (M3) e propaga para M6/OE4.

    Fluxo Option B — M3 → M6 → OE4:
      1. Modelo mensal DFC-first — fonte de verdade para 2025
      2. Modelo anual DR + Balanço + DFC construídos independentemente
      3. Linha 2025 do Balanço anual substituída pelo fecho de Dezembro mensal
      4. DFC anual reconstruída com o Balanço híbrido — 2026-2029 usam o
         fecho mensal de 2025 como abertura, propagando M3 → M6 → OE4
    """
    from ..dr import build_dr as _build_dr_anual
    from ..balanco import build_balanco as _build_balanco_anual
    from ..dfc import build_dfc as _build_dfc_anual

    # ── 1. Modelo mensal DFC-first ────────────────────────────────────────────
    df_dr = teso_mod.build_dr_mensal(a, base, sched)
    df_t  = teso_mod.build_tesouraria(a, base, sched)
    df_bs, df_dfc = _build_integrated_monthly(a, base, sched, df_dr, df_t)

    df_nfm = build_nfm_mensal(df_bs, df_dr)
    df_tc  = build_tesouraria_completa(a, base, sched, df_bs=df_bs)

    # ── 2. Modelo anual independente (DR → Balanço → DFC) ────────────────────
    df_dr_anual     = _build_dr_anual(a, base, sched)
    df_bs_anual_raw = _build_balanco_anual(a, base, sched, df_dr_anual)

    # ── 3. Opção B: fecho Dez mensal ancora Balanço anual 2025 ───────────────
    df_bs_hibrido = _overlay_dez_mensal_no_anual(df_bs, df_bs_anual_raw)

    # ── 4. DFC anual reconstruída — 2026-2029 partem do fecho mensal 2025 ────
    df_dfc_anual = _build_dfc_anual(a, df_dr_anual, df_bs_hibrido, sched, base)

    stmt_m6 = {
        "dr":      df_dr_anual,
        "balanco": df_bs_hibrido,
        "dfc":     df_dfc_anual,
    }

    # ── 5. Reconciliação DR/DFC (Balanço Dez = anual 2025 por construção) ────
    reconciliacao = build_reconciliacao_mensal_anual(
        a, base, df_bs, df_dr, df_dfc, sched, stmt=stmt_m6
    )

    return {
        "dr_mensal":           df_dr,
        "balanco_mensal":      df_bs,
        "dfc_mensal":          df_dfc,
        "nfm_mensal":          df_nfm,
        "tesouraria_completa": df_tc,
        "reconciliacao_anual": reconciliacao,
        "stmt_m6":             stmt_m6,
    }
