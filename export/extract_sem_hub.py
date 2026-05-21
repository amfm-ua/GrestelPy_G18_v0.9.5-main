"""Extrai AF sem hub para comparação."""
import sys
sys.path.insert(0, 'src')
from engine.modelo.model import run_model
import pandas as pd

pd.set_option('display.float_format', '{:,.0f}'.format)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 300)

print("=== SEM HUB ===")
result_off = run_model(hub_on=False)
if isinstance(result_off, dict):
    kpis = result_off.get('kpis')
    if kpis is not None:
        cols = ['ano', 'vn', 'ebitda', 'rl', 'total_ativo', 'cp', 'autonomia_financeira']
        avail = [c for c in cols if c in kpis.columns]
        print(kpis[avail].to_string(index=False))

        # Calculate AF manually
        for _, row in kpis.iterrows():
            ano = int(row['ano'])
            atl = float(row['total_ativo'])
            cp = float(row['cp'])
            af = cp / atl * 100 if atl > 0 else 0
            print(f"  {ano}: AF = {af:.1f}%  CP={cp:,.0f}  ATL={atl:,.0f}")

print()
print("=== COM HUB ===")
result_on = run_model(hub_on=True)
if isinstance(result_on, dict):
    kpis = result_on.get('kpis')
    if kpis is not None:
        for _, row in kpis.iterrows():
            ano = int(row['ano'])
            atl = float(row['total_ativo'])
            cp = float(row['cp'])
            af = cp / atl * 100 if atl > 0 else 0
            vn = float(row['vn'])
            ebitda = float(row['ebitda'])
            rl = float(row['rl'])
            print(f"  {ano}: VN={vn/1e3:,.0f}k  EBITDA={ebitda/1e3:,.0f}k ({ebitda/vn*100:.1f}%)  RL={rl/1e3:,.0f}k  AF={af:.1f}%")
