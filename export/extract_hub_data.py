"""Script temporário para extrair dados do hub para os relatórios M6/OE4."""
import sys
sys.path.insert(0, 'src')

from engine.projetos.hub_logistico import (
    load, viabilidade_hub, mapa_servico_divida,
    hub_capex, hub_dr_impact, hub_fcf, tornado_hub
)
from engine.modelo.model import run_model
import pandas as pd

hub = load()

print("=" * 60)
print("=== VIABILIDADE HUB ===")
print("=" * 60)
res = viabilidade_hub(hub)
print(f"VAL (WACC 8%, 10 anos): {res.get('val'):,.0f} €")
print(f"TIR: {res.get('tir'):.2%}")
print(f"Payback simples: {res.get('payback_simples'):.2f} anos")
print(f"Payback atualizado: {res.get('payback_atualizado'):.2f} anos")
print(f"Índice de Rendibilidade: {res.get('indice_rendibilidade'):.4f}")
print(f"Valor terminal: {res.get('valor_terminal'):,.0f} €")
print(f"Valor residual ativos: {res.get('valor_residual_ativos'):,.0f} €")
print(f"Parâmetros: {res.get('parametros', {})}")

print()
print("=" * 60)
print("=== FCF UNLEVERED (FCFF) ===")
print("=" * 60)
fcf_df = res.get('fcf_df')
if hasattr(fcf_df, 'to_string'):
    pd.set_option('display.float_format', '{:,.0f}'.format)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', 200)
    print(fcf_df.to_string())

print()
print("=" * 60)
print("=== MAPA SERVIÇO DA DÍVIDA ===")
print("=" * 60)
ds = mapa_servico_divida(hub)
if hasattr(ds, 'to_string'):
    print(ds.to_string())

print()
print("=" * 60)
print("=== IMPACTO ANUAL NO DR ===")
print("=" * 60)
dr = hub_dr_impact(hub)
if hasattr(dr, 'to_string'):
    print(dr.to_string())

print()
print("=" * 60)
print("=== TORNADO (TOP 5) ===")
print("=" * 60)
try:
    tornado = tornado_hub()
    if hasattr(tornado, 'to_string'):
        print(tornado[['label', 'driver', 'val_low', 'val_base', 'val_high', 'impacto_total']].to_string())
except Exception as e:
    print(f"Erro tornado: {e}")

print()
print("=" * 60)
print("=== MODELO COMPLETO (KPIs 2024-2029 com hub) ===")
print("=" * 60)
try:
    model_result = run_model(hub_on=True)
    if isinstance(model_result, dict):
        kpis = model_result.get('kpis')
        if kpis is not None and hasattr(kpis, 'to_string'):
            print("KPIs:")
            print(kpis.to_string())
        dr_full = model_result.get('dr')
        if dr_full is not None and hasattr(dr_full, 'to_string'):
            print("\nDR completa:")
            print(dr_full.to_string())
        bal = model_result.get('balanco')
        if bal is not None and hasattr(bal, 'to_string'):
            print("\nBalanço:")
            print(bal.to_string())
except Exception as e:
    print(f"Erro modelo completo: {e}")
    import traceback
    traceback.print_exc()
