import os, tempfile, importlib, sys
from datetime import date
from pathlib import Path

DB=tempfile.NamedTemporaryFile(suffix='.db', delete=False).name
os.environ['GESTION_DB_PATH']=DB
os.environ['GESTION_SECRET_KEY']='TEST-STABILIZATION-SECRET-32-CHARACTERS-LONG'
sys.path.insert(0, str(Path(__file__).resolve().parents[1]/'backend'))
import app.main as m
from fastapi.testclient import TestClient
client=TestClient(m.app)

def ok(r, code=200):
    assert r.status_code==code, f'{r.request.method} {r.request.url} => {r.status_code}: {r.text}'
    return r.json()

def auth(user='admin', password='admin123'):
    return ok(client.post('/api/auth/login', json={'usuario':user,'password':password}))['access_token']

def H(t): return {'Authorization':f'Bearer {t}'}

# 1-2 auth/system
assert ok(client.get('/api/health'))['estado']=='ok'
t_admin=auth()
me=ok(client.get('/api/auth/me',headers=H(t_admin)))
assert me['rol']=='admin'
state=ok(client.get('/api/sistema/estado',headers=H(t_admin)))
assert state['base_datos']['integridad']=='ok'

# 3 teachers
teach1=ok(client.post('/api/docentes',headers=H(t_admin),json={'nombre':'Ana Docente','especialidad':'Matemática'}))
teach2=ok(client.post('/api/docentes',headers=H(t_admin),json={'nombre':'Luis Docente','especialidad':'Lengua'}))
assert ok(client.get('/api/docentes',headers=H(t_admin)))

# 4 teacher users/linking
u1=ok(client.post('/api/usuarios',headers=H(t_admin),json={'usuario':'ana','nombre':'Ana Docente','rol':'docente','password':'ana12345'}))
ok(client.put(f"/api/usuarios/{u1['id']}/docente",headers=H(t_admin),json={'docente_id':teach1['id']}))
t_ana=auth('ana','ana12345')
assert ok(client.get('/api/auth/me',headers=H(t_ana)))['docente_id']==teach1['id']
assert client.post('/api/docentes',headers=H(t_ana),json={'nombre':'No Debe'}).status_code==403

# 5-6 school year + periods
an=ok(client.post('/api/anios-escolares',headers=H(t_admin),json={'nombre':'2027-2028','fecha_inicio':'2027-08-01','fecha_fin':'2028-06-30','activo':False,'cerrado':False}))
ok(client.post(f"/api/anios-escolares/{an['id']}/activar",headers=H(t_admin)))
periods=ok(client.get('/api/periodos-academicos',headers=H(t_admin),params={'anio_escolar_id':an['id']}))
assert len(periods)==4 and all(p['anio_escolar_id']==an['id'] for p in periods)

# 7-9 class, subjects, teacher assignment, students
course=ok(client.post('/api/cursos',headers=H(t_admin),json={'nombre':'4to A','nivel':'Secundaria','grado':'4to','seccion':'A','anio_escolar':'2027-2028','docente_id':teach1['id']}))
math=ok(client.post('/api/asignaturas',headers=H(t_admin),json={'nombre':'Matemática','codigo':'MAT','grado':'4to'}))
lang=ok(client.post('/api/asignaturas',headers=H(t_admin),json={'nombre':'Lengua Española','codigo':'LEN','grado':'4to'}))
ok(client.post('/api/cursos/asignaturas',headers=H(t_admin),json={'curso_id':course['id'],'asignatura_id':math['id'],'docente_id':teach1['id']}))
ok(client.post('/api/cursos/asignaturas',headers=H(t_admin),json={'curso_id':course['id'],'asignatura_id':lang['id'],'docente_id':teach2['id']}))
st1=ok(client.post('/api/estudiantes',headers=H(t_admin),json={'nombres':'Juan Pedro','apellidos':'Gomez Diaz','fecha_nacimiento':'2010-05-10','curso_id':course['id']}))
st2=ok(client.post('/api/estudiantes',headers=H(t_admin),json={'nombres':'Maria','apellidos':'Lopez','fecha_nacimiento':'2010-02-03','curso_id':course['id']}))
assert st1['matricula'].endswith('-1') and st2['matricula'].endswith('-2')
assert len(ok(client.get('/api/estudiantes',headers=H(t_admin),params={'curso_id':course['id']})))==2

# 10 competencies
c1=ok(client.post('/api/competencias',headers=H(t_admin),json={'asignatura_id':math['id'],'codigo':'C1','nombre':'Resolución','descripcion':'Resuelve problemas','orden':1}))
cf=ok(client.post('/api/competencias-fundamentales',headers=H(t_admin),json={'codigo':'CF1','nombre':'Pensamiento','descripcion':'Pensamiento crítico','orden':1}))
ok(client.post('/api/competencias-fundamentales/mapear',headers=H(t_admin),json={'fundamental_id':cf['id'],'especifica_id':c1['id']}))
assert len(ok(client.get(f"/api/competencias-fundamentales/{cf['id']}/especificas",headers=H(t_admin))))==1

# 11 planning
plan=ok(client.post('/api/planificaciones',headers=H(t_ana),json={'curso_id':course['id'],'asignatura_id':math['id'],'anio_escolar_id':an['id'],'periodo':1,'competencia_id':c1['id'],'titulo':'Unidad 1','fecha_inicio':'2027-08-10','fecha_fin':'2027-08-20','tipo_actividad':'Proyecto','instrumento':'Rúbrica','descripcion':'Plan','recursos':'Libro','crear_evento':True}))
assert plan['evento_calendario_id'] is not None

# 12 activity
act=ok(client.post('/api/actividades',headers=H(t_ana),json={'estudiante_id':st1['id'],'asignatura_id':math['id'],'competencia_id':c1['id'],'periodo':1,'fecha':'2026-09-12','titulo':'Tarea 1','tipo_evaluacion':'Tarea','instrumento':'Lista de cotejo','calificacion':85}))
ok(client.put(f"/api/actividades/{act['id']}",headers=H(t_ana),json={'fecha':'2026-09-13','titulo':'Tarea 1 editada','tipo_evaluacion':'Tarea','instrumento':'Lista de cotejo','calificacion':90}))

# 13 attendance
ok(client.post('/api/asistencia',headers=H(t_ana),json={'estudiante_id':st1['id'],'fecha':'2026-09-12','estado':'Presente','hora_entrada':'07:55'}))
ok(client.post('/api/asistencia',headers=H(t_ana),json={'estudiante_id':st2['id'],'fecha':'2026-09-12','estado':'Ausente'}))
res=ok(client.get('/api/asistencia/resumen',headers=H(t_ana),params={'curso_id':course['id']}))

# 14 incidents
inc=ok(client.post('/api/incidencias',headers=H(t_ana),json={'estudiante_id':st2['id'],'fecha':'2026-09-12','tipo':'Disciplina','gravedad':'Leve','descripcion':'Prueba'}))
assert inc['id']

# 15 permission
per=ok(client.post('/api/permisos',headers=H(t_ana),json={'estudiante_id':st2['id'],'fecha':'2026-09-13','tipo':'Salida','hora_desde':'10:00','hora_hasta':'11:00','motivo':'Cita'}))
ok(client.put(f"/api/permisos/{per['id']}",headers=H(t_admin),json={'estado':'Aprobado','observacion':'OK'}))

# 16-17 grades/gradebook
ok(client.post('/api/calificaciones/lote',headers=H(t_ana),json={'curso_id':course['id'],'asignatura_id':math['id'],'periodo':1,'tipo':'P','registros':[{'estudiante_id':st1['id'],'competencia_id':c1['id'],'calificacion':90}]}))
gb=ok(client.get('/api/libro-calificaciones',headers=H(t_ana),params={'curso_id':course['id'],'asignatura_id':math['id'],'periodo':1}))
assert len(gb['filas'])==2

# 18 closing
close=ok(client.get('/api/cierre/asignatura',headers=H(t_ana),params={'curso_id':course['id'],'asignatura_id':math['id']}))
assert close

# 19 reports
for path in [
    f'/api/reportes/cierre.xlsx?curso_id={course["id"]}&asignatura_id={math["id"]}',
    f'/api/reportes/cierre.pdf?curso_id={course["id"]}&asignatura_id={math["id"]}',
    f'/api/reportes/boletin/{st1["id"]}.xlsx',
    f'/api/reportes/boletin/{st1["id"]}.pdf',
    f'/api/reportes/asistencia.xlsx?curso_id={course["id"]}',
    f'/api/reportes/asistencia.pdf?curso_id={course["id"]}',
    f'/api/reportes/incidencias.xlsx?curso_id={course["id"]}',
    f'/api/reportes/permisos.xlsx?curso_id={course["id"]}',
]:
    rr=client.get(path,headers=H(t_admin)); assert rr.status_code==200 and len(rr.content)>20, (path,rr.status_code,rr.text[:200])

# 20 risk + intervention
risk=ok(client.get(f'/api/estudiante/{st2["id"]}/riesgo',headers=H(t_ana)))
alerts=ok(client.get('/api/alertas-avanzadas',headers=H(t_ana)))
# if an alert exists, exercise management
if alerts:
    aid=alerts[0]['id']
    ok(client.put(f'/api/alertas-avanzadas/{aid}',headers=H(t_ana),json={'estado':'En seguimiento','responsable':'Ana','observacion':'Revisada'}))
plan2=ok(client.post('/api/planes-intervencion',headers=H(t_ana),json={'estudiante_id':st2['id'],'titulo':'Plan prueba','objetivo':'Mejorar','estrategia':'Acompañamiento','responsable':'Ana','fecha_inicio':'2027-08-15'}))
action=ok(client.post('/api/planes-intervencion/acciones',headers=H(t_ana),json={'plan_id':plan2['id'],'descripcion':'Acción 1','responsable':'Ana','fecha_limite':'2026-09-20'}))

# 21 expediente
ok(client.put(f'/api/estudiantes/{st1["id"]}/datos',headers=H(t_ana),json={'direccion':'Calle 1','contacto_emergencia':'Pedro','telefono_emergencia':'809'}))
fam=ok(client.post(f'/api/estudiantes/{st1["id"]}/familiares',headers=H(t_ana),json={'nombre':'Pedro Gomez','parentesco':'Padre','es_contacto_principal':True}))
exp=ok(client.get(f'/api/estudiantes/{st1["id"]}/expediente',headers=H(t_ana)))
assert exp['datos']['direccion']=='Calle 1' and len(exp['familiares'])==1

# 22 communications
com=ok(client.post('/api/comunicados',headers=H(t_admin),json={'titulo':'Aviso','mensaje':'Mensaje','tipo':'General','prioridad':'Normal','destinatarios':[{'tipo_destinatario':'todos'}]}))
incbox=ok(client.get('/api/comunicados',headers=H(t_ana)))
ok(client.post(f'/api/comunicados/{com["id"]}/leer',headers=H(t_ana)))

# 23 agenda
ag=ok(client.post('/api/agenda',headers=H(t_ana),json={'titulo':'Examen','tipo':'Evaluación','fecha':'2026-09-18','hora_inicio':'08:00','hora_fin':'09:00','curso_id':course['id'],'asignatura_id':math['id']}))
assert len(ok(client.get('/api/agenda/proximos',headers=H(t_ana))))>=1
ok(client.put(f'/api/agenda/{ag["id"]}',headers=H(t_ana),json={'titulo':'Examen editado','tipo':'Evaluación','fecha':'2026-09-18','hora_inicio':'08:30','hora_fin':'09:30','curso_id':course['id'],'asignatura_id':math['id']}))

# 24 schedule + conflict
hor=ok(client.post('/api/horario',headers=H(t_ana),json={'dia_semana':1,'hora_inicio':'08:00','hora_fin':'09:00','curso_id':course['id'],'asignatura_id':math['id'],'aula':'A1','bloque':'1'}))
conf=client.post('/api/horario',headers=H(t_ana),json={'dia_semana':1,'hora_inicio':'08:30','hora_fin':'09:30','curso_id':course['id'],'asignatura_id':math['id'],'aula':'A2','bloque':'2'})
assert conf.status_code==409, conf.text
ok(client.put(f'/api/horario/{hor["id"]}',headers=H(t_ana),json={'dia_semana':1,'hora_inicio':'09:00','hora_fin':'10:00','curso_id':course['id'],'asignatura_id':math['id'],'aula':'A1','bloque':'1'}))

# 25 import teachers/students admin
csv_t='nombre,cedula,correo,telefono,especialidad,estado\nImportado,001,a@b.com,809,Historia,Activo\n'
imp=ok(client.post('/api/admin/importar/docentes',headers=H(t_admin),json={'csv_text':csv_t}))
assert imp['creados']>=1
csv_s='nombres,apellidos,fecha_nacimiento,curso_id,correo,telefono,estado\nCarlos,Importado,2010-01-01,%d,c@d.com,809,Activo\n' % course['id']
imp2=ok(client.post('/api/admin/importar/estudiantes',headers=H(t_admin),json={'csv_text':csv_s}))
assert imp2['creados']>=1

# 26 export
rr=client.get('/api/admin/exportar/general.xlsx',headers=H(t_admin)); assert rr.status_code==200 and rr.content[:2]==b'PK'

# 27 backup/restore
rr=client.get('/api/admin/backup',headers=H(t_admin)); assert rr.status_code==200 and len(rr.content)>100
rr2=client.post('/api/admin/restore',headers={**H(t_admin),'Content-Type':'application/octet-stream'},content=rr.content); assert rr2.status_code==200, rr2.text
assert ok(client.get('/api/health'))['estado']=='ok'

# 28 role/access controls
assert client.get('/api/usuarios',headers=H(t_ana)).status_code==403
# Ana has math assignment, should not write lang grades
bad=client.post('/api/calificaciones/lote',headers=H(t_ana),json={'curso_id':course['id'],'asignatura_id':lang['id'],'periodo':1,'tipo':'P','registros':[{'estudiante_id':st1['id'],'competencia_id':c1['id'],'calificacion':70}]})
assert bad.status_code in (400,403), bad.text

# 29 final integrity + audit evidence
state=ok(client.get('/api/sistema/estado',headers=H(t_admin)))
assert state['base_datos']['integridad']=='ok'
assert state['conteos']['estudiantes']>=3
assert state['conteos']['auditoria']>0

print('E2E_STABILIZATION_OK')
