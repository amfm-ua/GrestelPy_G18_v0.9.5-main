"""Extrai dados dos 4 cenários e KPIs consolidados com hub."""
import sys, copy
sys.path.insert(0, 'src')

from engine.projetos.hub_logistico import load, viabilidade_hub, mapa_servico_divida
from engine.modelo.model import run_model

hub = load()

scenarios = [
    ('Base', {}),
    ('Upside', {'mult_op': 1.2, 'mult_vn': 1.2}),
    ('Downside', {'mult_op': 0.8, 'mult_vn': 0.8}),
    ('Stress', {'mult_capex': 1.15, 'mult_op': 0.7, 'mult_vn': 0.6}),
]

print("=" * 70)
print("=== 4 CENÁRIOS ===")
print("=" * 70)
for name, ov in scenarios:
    h = copy.deepcopy(hub)
    b = h['projeto_hub']['beneficios_anuais']
    c = h['projeto_hub']['beneficios_comerciais']
    cp = h['projeto_hub']['capex']
    mo = ov.get('mult_op', 1.0)
    mv = ov.get('mult_vn', 1.0)
    mc = ov.get('mult_capex', 1.0)
    if mo != 1.0:
        b['poupanca_operacional'] = 480000 * mo
        b['reducao_quebras'] = 80000 * mo
        b['beneficio_liquido_anual'] = 350000 * mo
    if mv != 1.0:
        vi = c.get('vn_incremental', {})
        c['vn_incremental'] = {k: v * mv for k, v in vi.items()}
    if mc != 1.0:
        cp['base'] = 6000000 * mc
        cp['cronograma'] = {2025: 3000000 * mc, 2026: 3000000 * mc}
    res = viabilidade_hub(h)
    val = res.get('val')
    tir = res.get('tir')
    pb = res.get('payback_atualizado')
    ir = res.get('indice_rendibilidade')
    print(f"{name:10s}: VAL={val:>12,.0f} EUR  TIR={tir:.2%}  Payback={pb:.2f}a  IR={ir:.4f}")

print()
print("=" * 70)
print("=== MODELO COMPLETO (hub_on=True) - KPIs ===")
print("=" * 70)
try:
    result = run_model(hub_on=True)
    if isinstance(result, dict):
        kpis = result.get('kpis')
        if kpis is not None:
            import pandas as pd
            pd.set_option('display.max_columns', None)
            pd.set_option('display.width', 300)
            pd.set_option('display.float_format', '{:,.0f}'.format)
            cols = ['ano', 'vn', 'ebitda', 'ebit', 'rl', 'autonomia_financeira',
                    'liquidez_geral', 'dmi_dias', 'pmr_dias', 'pmp_dias', 'ciclo_caixa',
                    'debt_ebitda', 'nd_ebitda', 'cobertura_juros', 'dscr', 'total_ativo', 'cp']
            avail = [c for c in cols if c in kpis.columns]
            print(kpis[avail].to_string(index=False))
except Exception as e:
    import traceback
    print(f"Erro: {e}")
    traceback.print_exc()

print()
print("=" * 70)
print("=== MAPA SERVIÇO DÍVIDA (extendido) ===")
print("=" * 70)
ds = mapa_servico_divida(hub)
if hasattr(ds, 'to_string'):
    import pandas as pd
    pd.set_option('display.float_format', '{:,.0f}'.format)
    print(ds.to_string(index=False))
