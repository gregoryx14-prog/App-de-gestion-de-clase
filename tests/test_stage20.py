import os, tempfile
from pathlib import Path
os.environ['GESTION_DB_PATH']=str(Path(tempfile.gettempdir())/'gestion_clase_stage20_test.db')
try: Path(os.environ['GESTION_DB_PATH']).unlink()
except FileNotFoundError: pass
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]/'backend'))
from app.main import app, init_db
from fastapi.testclient import TestClient
init_db()
c=TestClient(app)
r=c.get('/api/health'); assert r.status_code==200 and r.json()['estado']=='ok'
r=c.post('/api/auth/login',json={'usuario':'admin','password':'admin123'}); assert r.status_code==200
t=r.json()['access_token']
r=c.get('/api/sistema/estado',headers={'Authorization':'Bearer '+t}); assert r.status_code==200 and r.json()['base_datos']['integridad']=='ok'
print('STAGE20_TEST_OK')
