from __future__ import annotations

import pandas as pd

from ...inputs import Assumptions, Base2024, Schedules, MESES
from ...financiamento import tesouraria as teso_mod
from .auxiliares import _financiamento_mensal, _capex_mensal
from .integrado import _build_integrated_monthly


# ──────────────────────────────────────────────────────────────────────────────
# Balanço Mensal
# ──────────────────────────────────────────────────────────────────────────────

def build_balanco_mensal(
    a: Assumptions,
    base: Base2024,
    sched: Schedules,
    _df_dr_m: pd.DataFrame | None = None,
    _df_t_m: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """Balanço mensal 2025, com Caixa derivada dos fluxos DFC."""
    if _df_dr_m is None:
        _df_dr_m = teso_mod.build_dr_mensal(a, base, sched)

    if _df_t_m is None:
        _df_t_m = teso_mod.build_tesouraria(a, base, sched)

    df_bs, _ = _build_integrated_monthly(a, base, sched, _df_dr_m, _df_t_m)
    return df_bs


# ──────────────────────────────────────────────────────────────────────────────
# DFC Mensal
# ──────────────────────────────────────────────────────────────────────────────

def build_dfc_mensal(
    a: Assumptions,
    base: Base2024,
    sched: Schedules,
    df_bs: pd.DataFrame | None = None,
    df_dr_m: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """DFC mensal 2025 pelo método indireto, reconciliada com o Balanço."""
    if df_dr_m is None:
        df_dr_m = teso_mod.build_dr_mensal(a, base, sched)

    df_t_m = teso_mod.build_tesouraria(a, base, sched)

    _, df_dfc = _build_integrated_monthly(a, base, sched, df_dr_m, df_t_m)
    return df_dfc


# ──────────────────────────────────────────────────────────────────────────────
# NFM Mensal
# ──────────────────────────────────────────────────────────────────────────────

def build_nfm_mensal(
    df_bs: pd.DataFrame,
    df_dr_m: pd.DataFrame,
) -> pd.DataFrame:
    """NFM e Ciclo de Conversão de Caixa mensais, derivados do Balanço."""
    bs_map = df_bs.set_index("mes").to_dict("index")
    dr_map = df_dr_m.set_index("mes").to_dict("index")

    rows: list[dict] = []

    for m in MESES:
        bs = bs_map[m]
        dr = dr_map[m]

        vn_m = max(dr["vn"], 1)
        cmvmc_m = max(dr["cmvmc"], 1)
        fse_m = max(dr["fse"], 1)

        ac_cicl = bs["inventarios"] + bs["clientes"]
        pc_cicl = bs["fornecedores"] + bs["eoep_credor"]
        nfm_m = ac_cicl - pc_cicl

        pmr_eff = bs["clientes"] / vn_m * 30
        dmi_eff = bs["inventarios"] / cmvmc_m * 30
        pmp_eff = bs["fornecedores"] / (cmvmc_m + fse_m) * 30
        ccc_eff = pmr_eff + dmi_eff - pmp_eff

        rows.append(
            {
                "mes": m,
                "ativo_ciclico": round(ac_cicl),
                "inventarios": round(bs["inventarios"]),
                "clientes": round(bs["clientes"]),
                "passivo_ciclico": round(pc_cicl),
                "fornecedores": round(bs["fornecedores"]),
                "eoep_credor": round(bs["eoep_credor"]),
                "nfm": round(nfm_m),
                "pmr_eff": round(pmr_eff, 1),
                "dmi_eff": round(dmi_eff, 1),
                "pmp_eff": round(pmp_eff, 1),
                "ccc_eff": round(ccc_eff, 1),
            }
        )

    df = pd.DataFrame(rows)
    df["delta_nfm"] = df["nfm"].diff().fillna(0).round().astype(int)

    return df


# ──────────────────────────────────────────────────────────────────────────────
# Tesouraria Completa
# ──────────────────────────────────────────────────────────────────────────────

def build_tesouraria_completa(
    a: Assumptions,
    base: Base2024,
    sched: Schedules,
    df_bs: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """Tesouraria mensal completa: operacional + serviço dívida + CAPEX."""
    if df_bs is None:
        df_bs = build_balanco_mensal(a, base, sched)

    df_teso = teso_mod.build_tesouraria(a, base, sched)
    fin_m = _financiamento_mensal(sched)
    cap_m = _capex_mensal(sched)

    t_map = df_teso.set_index("mes").to_dict("index")
    bs_map = df_bs.set_index("mes").to_dict("index")

    caixa_prev = base.balanco["ativo_corrente"]["Caixa"]

    rows: list[dict] = []

    for m in MESES:
        t = t_map[m]
        bs = bs_map[m]
        fin = fin_m[m]
        cap = cap_m[m]

        rec = t["recebimentos_clientes"]
        pag_forn = t["pagamentos_fornecedores"]
        pag_pess = t["pagamentos_pessoal"]
        fluxo_fisc = t["fluxo_fiscal"]
        fluxo_op_b = t["fluxo_operacional_bruto"]
        fluxo_op_l = t["fluxo_liquido"]

        capex_pag = cap["capex_aft"] + cap["capex_int"]
        amort_pag = fin["amortizacao"]
        juros_pag = fin["juros"]

        fluxo_fin = -(amort_pag + juros_pag)

        var_total = fluxo_op_l - capex_pag + fluxo_fin
        caixa_bruta = caixa_prev + var_total

        linha_cp = bs["linha_credito_cp"]

        rows.append(
            {
                "mes": m,
                "recebimentos_clientes": round(rec),
                "pagamentos_fornecedores": round(pag_forn),
                "pagamentos_pessoal": round(pag_pess),
                "fluxo_operacional_bruto": round(fluxo_op_b),
                "iva_pago_estado": round(t["iva_pagamento_estado"]),
                "ss_pagamento": round(t["ss_pagamento"]),
                "irc_ppc": round(t["irc_ppc"]),
                "fluxo_fiscal": round(fluxo_fisc),
                "fluxo_operacional_liquido": round(fluxo_op_l),
                "capex_pagamento": round(-capex_pag),
                "amortizacoes": round(-amort_pag),
                "juros_pagos": round(-juros_pag),
                "fluxo_financiamento": round(fluxo_fin),
                "variacao_caixa_total": round(var_total),
                "caixa_abertura": round(caixa_prev),
                "caixa_antes_credito": round(caixa_bruta),
                "linha_credito_utilizada": round(linha_cp),
                "caixa_fecho": round(bs["caixa"]),
            }
        )

        caixa_prev = bs["caixa"]

    return pd.DataFrame(rows)
