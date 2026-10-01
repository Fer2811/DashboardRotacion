from pathlib import Path
import os,subprocess,sys
ROOT=Path(__file__).resolve().parent
if not (ROOT/'.env').exists():raise SystemExit('Falta .env. Copia .env.example a .env.')
if not (ROOT/'base_dashboard_rotacion.xlsx').exists():raise SystemExit('Falta base_dashboard_rotacion.xlsx. Ejecuta 01 y 02 o coloca la base vigente.')
print('[1/2] Generando dashboard + acciones Odoo...')
if subprocess.run([sys.executable,str(ROOT/'03_generar_dashboard.py')],cwd=ROOT).returncode:raise SystemExit('03 terminó con error')
print('[2/2] Iniciando servidor...')
os.execv(sys.executable,[sys.executable,str(ROOT/'dashboard_server.py')])
