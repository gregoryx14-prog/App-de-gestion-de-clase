"""Ejecutor de las pruebas de consolidación V5.1 Web."""
import subprocess, sys
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
env=dict(__import__('os').environ)
env['PYTHONPATH']=str(ROOT/'backend')
for test in ('test_stage20.py','test_e2e_stabilization.py'):
    print(f'==> {test}')
    r=subprocess.run([sys.executable,str(Path(__file__).with_name(test))],cwd=ROOT,env=env)
    if r.returncode:
        raise SystemExit(r.returncode)
print('ALL_STABILIZATION_TESTS_OK')
