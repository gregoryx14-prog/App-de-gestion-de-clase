
# ============================================================
# SISTEMA DE GESTIÓN DE CLASE - VERSIÓN 5.1 — DOCENTES POR ASIGNATURA
# Google Colab + Gradio + SQLite + Google Drive
#
# V5 incluye:
# - Todo lo de V3: usuarios, docentes, cursos, asignaturas, estudiantes, asistencia,
#   incidencias, permisos, calendario, estadísticas, alertas, Excel, PDF y respaldos.
# - Competencias fundamentales y específicas administrables.
# - Relación independiente entre competencias fundamentales y específicas (una CE puede
#   estar asociada a una o varias CF sin duplicar su descripción).
# - Crear, editar, eliminar, activar/desactivar, copiar, pegar y codificar competencias.
# - Actividades evaluativas por período, estudiante, competencia, tipo e instrumento.
# - P1-P4 calculados automáticamente desde el promedio de actividades de cada competencia.
# - RP1-RP4 manuales y reemplazo configurable.
# - Vista general del promedio de todas las actividades para verificación.
# - Cálculo anual y módulo de calificaciones finales con Completiva, Extraordinaria,
#   Especial y Situación Final (A/R), siguiendo la estructura del documento suministrado.
# - Boletín final individual en PDF y exportación Excel.
#
# CRITERIO DE ACTIVIDADES:
# Cada actividad pertenece a UNA competencia específica y a UN período. Para cada
# estudiante y competencia, la nota del período es el promedio de las actividades
# registradas para esa competencia en ese período. La vista general de actividades
# sirve únicamente como verificación y no sustituye el promedio por competencia.
#
# CRITERIO ANUAL:
# La Calificación Final (C.F.) se obtiene del promedio de las competencias específicas
# con datos. En la hoja final se aplican, de forma configurable, las ponderaciones que
# aparecen en el documento: Completiva = 50% C.F. + 50% C.E.C.; Extraordinaria =
# 30% C.F. + 70% C.E.EX. Las calificaciones especiales pueden registrarse manualmente.
# La situación final A/R se determina usando el umbral de aprobación configurado.
# Estas reglas son configurables y no se presentan como una fórmula oficial del MINERD.
# ============================================================

import os
import re
import sqlite3
import shutil
import hashlib
import secrets
import math
from pathlib import Path
from datetime import datetime, date, time, timedelta

# ---------- Instalación ----------
os.system("pip -q install 'gradio>=6.0,<7.0' pandas openpyxl matplotlib reportlab")

import pandas as pd
import matplotlib.pyplot as plt
import gradio as gr

from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.enums import TA_CENTER


# ============================================================
# 1. GOOGLE DRIVE Y RUTAS
# ============================================================

IN_COLAB = os.path.exists("/content")
if IN_COLAB:
    try:
        from google.colab import drive
        drive.mount("/content/drive")
    except Exception:
        pass

if os.path.exists("/content/drive/MyDrive"):
    BASE_DIR = Path("/content/drive/MyDrive/Gestion_Clase_V5")
else:
    BASE_DIR = Path("/content/Gestion_Clase_V5")

BASE_DIR.mkdir(parents=True, exist_ok=True)
REPORTES_DIR = BASE_DIR / "reportes"
BACKUP_DIR = BASE_DIR / "respaldos"
REPORTES_DIR.mkdir(exist_ok=True)
BACKUP_DIR.mkdir(exist_ok=True)

DB_PATH = BASE_DIR / "gestion_clase_v5.db"


# ============================================================
# 2. BASE DE DATOS
# ============================================================

def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def today_str():
    return date.today().isoformat()

def hash_password(password):
    salt = secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), salt.encode("utf-8"), 120000
    ).hex()
    return f"{salt}${digest}"

def verify_password(password, stored):
    try:
        salt, digest = stored.split("$", 1)
        test = hashlib.pbkdf2_hmac(
            "sha256", password.encode("utf-8"), salt.encode("utf-8"), 120000
        ).hex()
        return secrets.compare_digest(test, digest)
    except Exception:
        return False

def init_db():
    conn = get_conn()
    cur = conn.cursor()

    cur.executescript("""
    CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT UNIQUE NOT NULL,
        nombre TEXT NOT NULL,
        rol TEXT NOT NULL DEFAULT 'docente',
        password_hash TEXT NOT NULL,
        activo INTEGER DEFAULT 1,
        fecha_registro TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS docentes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL,
        cedula TEXT,
        correo TEXT,
        telefono TEXT,
        especialidad TEXT,
        estado TEXT DEFAULT 'Activo',
        fecha_registro TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS cursos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL,
        nivel TEXT DEFAULT 'Secundaria',
        grado TEXT,
        seccion TEXT,
        anio_escolar TEXT,
        docente_id INTEGER,
        estado TEXT DEFAULT 'Activo',
        FOREIGN KEY(docente_id) REFERENCES docentes(id) ON DELETE SET NULL
    );

    CREATE TABLE IF NOT EXISTS asignaturas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL,
        codigo TEXT,
        grado TEXT,
        descripcion TEXT,
        activa INTEGER DEFAULT 1,
        UNIQUE(nombre, grado)
    );

    CREATE TABLE IF NOT EXISTS estudiantes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        matricula TEXT UNIQUE NOT NULL,
        nombre TEXT NOT NULL,
        curso TEXT,
        seccion TEXT,
        correo TEXT,
        telefono TEXT,
        estado TEXT DEFAULT 'Activo',
        fecha_registro TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS estudiante_curso (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        estudiante_id INTEGER NOT NULL,
        curso_id INTEGER NOT NULL,
        fecha_inicio TEXT NOT NULL,
        fecha_fin TEXT,
        UNIQUE(estudiante_id, curso_id),
        FOREIGN KEY(estudiante_id) REFERENCES estudiantes(id) ON DELETE CASCADE,
        FOREIGN KEY(curso_id) REFERENCES cursos(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS curso_asignatura (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        curso_id INTEGER NOT NULL,
        asignatura_id INTEGER NOT NULL,
        docente_id INTEGER,
        activa INTEGER DEFAULT 1,
        UNIQUE(curso_id,asignatura_id),
        FOREIGN KEY(curso_id) REFERENCES cursos(id) ON DELETE CASCADE,
        FOREIGN KEY(asignatura_id) REFERENCES asignaturas(id) ON DELETE CASCADE,
        FOREIGN KEY(docente_id) REFERENCES docentes(id) ON DELETE SET NULL
    );

    CREATE TABLE IF NOT EXISTS asistencia (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        estudiante_id INTEGER NOT NULL,
        fecha TEXT NOT NULL,
        estado TEXT NOT NULL,
        hora_entrada TEXT,
        observacion TEXT,
        UNIQUE(estudiante_id, fecha),
        FOREIGN KEY(estudiante_id) REFERENCES estudiantes(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS incidencias (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        estudiante_id INTEGER NOT NULL,
        fecha TEXT NOT NULL,
        hora TEXT,
        tipo TEXT NOT NULL,
        gravedad TEXT NOT NULL,
        descripcion TEXT,
        accion_tomada TEXT,
        observacion TEXT,
        FOREIGN KEY(estudiante_id) REFERENCES estudiantes(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS permisos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        estudiante_id INTEGER NOT NULL,
        fecha TEXT NOT NULL,
        tipo TEXT NOT NULL,
        hora_desde TEXT,
        hora_hasta TEXT,
        motivo TEXT,
        estado TEXT DEFAULT 'Pendiente',
        observacion TEXT,
        FOREIGN KEY(estudiante_id) REFERENCES estudiantes(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS competencias (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        asignatura_id INTEGER NOT NULL,
        codigo TEXT NOT NULL,
        nombre TEXT NOT NULL,
        descripcion TEXT NOT NULL,
        orden INTEGER DEFAULT 1,
        activa INTEGER DEFAULT 1,
        UNIQUE(asignatura_id, codigo),
        FOREIGN KEY(asignatura_id) REFERENCES asignaturas(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS calificaciones (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        estudiante_id INTEGER NOT NULL,
        asignatura_id INTEGER NOT NULL,
        competencia_id INTEGER NOT NULL,
        periodo INTEGER NOT NULL,
        tipo TEXT NOT NULL DEFAULT 'P',
        calificacion REAL,
        observacion TEXT,
        fecha_actualizacion TEXT NOT NULL,
        UNIQUE(estudiante_id, asignatura_id, competencia_id, periodo, tipo),
        FOREIGN KEY(estudiante_id) REFERENCES estudiantes(id) ON DELETE CASCADE,
        FOREIGN KEY(asignatura_id) REFERENCES asignaturas(id) ON DELETE CASCADE,
        FOREIGN KEY(competencia_id) REFERENCES competencias(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS calendario (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha TEXT NOT NULL,
        hora_inicio TEXT,
        hora_fin TEXT,
        titulo TEXT NOT NULL,
        tipo TEXT,
        curso TEXT,
        seccion TEXT,
        descripcion TEXT
    );

    CREATE TABLE IF NOT EXISTS competencias_fundamentales (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        codigo TEXT UNIQUE NOT NULL,
        nombre TEXT NOT NULL,
        descripcion TEXT NOT NULL,
        orden INTEGER DEFAULT 1,
        activa INTEGER DEFAULT 1
    );

    CREATE TABLE IF NOT EXISTS competencia_fundamental_especifica (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fundamental_id INTEGER NOT NULL,
        especifica_id INTEGER NOT NULL,
        UNIQUE(fundamental_id, especifica_id),
        FOREIGN KEY(fundamental_id) REFERENCES competencias_fundamentales(id) ON DELETE CASCADE,
        FOREIGN KEY(especifica_id) REFERENCES competencias(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS actividades (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        estudiante_id INTEGER NOT NULL,
        asignatura_id INTEGER NOT NULL,
        competencia_id INTEGER NOT NULL,
        periodo INTEGER NOT NULL,
        fecha TEXT NOT NULL,
        titulo TEXT NOT NULL,
        descripcion TEXT,
        tipo_evaluacion TEXT NOT NULL,
        instrumento TEXT NOT NULL,
        calificacion REAL NOT NULL,
        observacion TEXT,
        fecha_registro TEXT NOT NULL,
        FOREIGN KEY(estudiante_id) REFERENCES estudiantes(id) ON DELETE CASCADE,
        FOREIGN KEY(asignatura_id) REFERENCES asignaturas(id) ON DELETE CASCADE,
        FOREIGN KEY(competencia_id) REFERENCES competencias(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS calificaciones_finales (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        estudiante_id INTEGER NOT NULL,
        asignatura_id INTEGER NOT NULL,
        completiva REAL,
        extraordinaria REAL,
        especial_cf REAL,
        especial_ce REAL,
        situacion TEXT,
        observacion TEXT,
        fecha_actualizacion TEXT NOT NULL,
        UNIQUE(estudiante_id, asignatura_id),
        FOREIGN KEY(estudiante_id) REFERENCES estudiantes(id) ON DELETE CASCADE,
        FOREIGN KEY(asignatura_id) REFERENCES asignaturas(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS configuracion (
        clave TEXT PRIMARY KEY,
        valor TEXT
    );
    """)

    # Migración V5: separar estudiante de su clase y añadir fecha de nacimiento.
    # Se conservan los datos de V4 y se amplía la estructura sin borrarlos.
    def add_column_if_missing(table, column, definition):
        cols = [r["name"] for r in cur.execute(f"PRAGMA table_info({table})").fetchall()]
        if column not in cols:
            cur.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")

    add_column_if_missing("estudiantes", "fecha_nacimiento", "TEXT")
    add_column_if_missing("estudiantes", "edad_registro", "INTEGER")
    add_column_if_missing("estudiantes", "curso_id", "INTEGER")
    add_column_if_missing("estudiantes", "orden_lista", "INTEGER")

    # V5.1 Docentes: una asignatura dentro de una clase puede tener su propio docente.
    cols_ca = [r["name"] for r in cur.execute("PRAGMA table_info(curso_asignatura)").fetchall()]
    if "docente_id" not in cols_ca:
        cur.execute("ALTER TABLE curso_asignatura ADD COLUMN docente_id INTEGER")

    # Si existe la BD V4 y la V5 es nueva, se intentará usarla como base de migración.
    # (El archivo V4 se copia antes de init_db; aquí solo completamos columnas.)

    # Migración sencilla desde nombres usados en V1/V2 si la DB vieja existe
    # en la misma carpeta (no sobrescribe la información).
    defaults = {
        "anio_escolar": "2026-2027",
        "max_calificacion": "100",
        "rp_reemplaza_p": "1",
        "final_promedio_competencias": "1",
        "nombre_centro": "Centro Educativo",
        "umbral_aprobacion": "70",
        "peso_completiva_cf": "50",
        "peso_completiva_cec": "50",
        "peso_extra_cf": "30",
        "peso_extra_ceex": "70",
    }
    for k, v in defaults.items():
        cur.execute(
            "INSERT OR IGNORE INTO configuracion(clave, valor) VALUES (?, ?)",
            (k, v)
        )

    # Usuario inicial
    cur.execute("SELECT COUNT(*) AS n FROM usuarios")
    if cur.fetchone()["n"] == 0:
        cur.execute(
            """INSERT INTO usuarios(usuario,nombre,rol,password_hash,fecha_registro)
               VALUES (?,?,?,?,?)""",
            ("admin", "Administrador", "admin", hash_password("admin123"), now_str())
        )

    # Asignatura de ejemplo alineada con el documento suministrado.
    cur.execute(
        "SELECT id FROM asignaturas WHERE nombre=? AND grado=?",
        ("Matemática", "4to. Grado")
    )
    row = cur.fetchone()
    if not row:
        cur.execute(
            """INSERT INTO asignaturas(nombre,codigo,grado,descripcion)
               VALUES (?,?,?,?)""",
            (
                "Matemática", "MAT",
                "4to. Grado",
                "Asignatura de Matemática para 4to. grado de secundaria."
            )
        )
        asignatura_id = cur.lastrowid

        competencias = [
            ("CE-MAT1", "Competencia 1",
             "Comprende la matemática leída en textos apropiados a su nivel de desarrollo, formulando preguntas de aclaración y ampliación, para clasificar conceptos y relaciones matemáticas."),
            ("CE-MAT2", "Competencia 2",
             "Construye argumentos sencillos haciendo uso de sus conocimientos y de reglas de razonamiento lógico para dar validez a las propias ideas matemáticas."),
            ("CE-MAT3", "Competencia 3",
             "Aplica estrategias integradas de resolución de problemas para investigar y entender conceptos matemáticos."),
            ("CE-MAT4", "Competencia 4",
             "Transmite asertivamente los propios puntos de vista para generar opciones creativas frente a la interpretación de situaciones matemáticas respetando las diferentes propuestas de los demás."),
            ("CE-MAT5", "Competencia 5",
             "Utiliza herramientas tecnológicas para la resolución de problemas y situaciones matemáticas diversas."),
            ("CE-MAT6", "Competencia 6",
             "Emplea sus conocimientos matemáticos conectándolos con otras áreas y disciplinas para comprender los fenómenos naturales que afectan el medioambiente."),
            ("CE-MAT7", "Competencia 7",
             "Muestra relaciones positivas en el trabajo matemático, en equipo, aportando soluciones frente a situaciones problemáticas relativas a la comunidad."),
        ]
        for i, (codigo, nombre, descripcion) in enumerate(competencias, 1):
            cur.execute(
                """INSERT INTO competencias
                   (asignatura_id,codigo,nombre,descripcion,orden)
                   VALUES (?,?,?,?,?)""",
                (asignatura_id, codigo, nombre, descripcion, i)
            )

    # Competencias fundamentales editables. Se cargan como catálogo inicial;
    # el docente puede modificarlas desde el módulo de competencias.
    fundamentales = [
        ("CF-EC", "Ética y Ciudadana", "Actúa con responsabilidad, respeto, convivencia y compromiso ciudadano."),
        ("CF-COM", "Comunicativa", "Comprende, produce e intercambia mensajes de manera efectiva en diferentes contextos."),
        ("CF-PCC", "Pensamiento Lógico, Creativo y Crítico", "Construye razonamientos, analiza información y genera soluciones creativas y críticas."),
        ("CF-RP", "Resolución de Problemas", "Identifica situaciones problemáticas, formula estrategias y evalúa soluciones."),
        ("CF-CT", "Científica y Tecnológica", "Utiliza conocimientos, procedimientos y herramientas científicas y tecnológicas para comprender y actuar."),
        ("CF-AAS", "Ambiental y de la Salud", "Promueve el cuidado de la salud, el ambiente y el uso responsable de los recursos."),
        ("CF-DPE", "Desarrollo Personal y Espiritual", "Fortalece identidad, autonomía, bienestar, valores y desarrollo integral."),
    ]
    for i,(codigo,nombre,descripcion) in enumerate(fundamentales,1):
        cur.execute(
            "INSERT OR IGNORE INTO competencias_fundamentales(codigo,nombre,descripcion,orden) VALUES (?,?,?,?)",
            (codigo,nombre,descripcion,i)
        )

    # Mapa inicial: cada CE de Matemática se puede asociar sin duplicar su texto.
    # Si la asignatura/competencias ya existían, se intenta completar el mapa.
    cur.execute("SELECT id FROM asignaturas WHERE nombre=? AND grado=?", ("Matemática", "4to. Grado"))
    math_row=cur.fetchone()
    if math_row:
        math_id=math_row["id"]
        cur.execute("SELECT id,codigo FROM competencias WHERE asignatura_id=?", (math_id,))
        ce_rows=cur.fetchall()
        cur.execute("SELECT id FROM competencias_fundamentales ORDER BY orden")
        cf_rows=cur.fetchall()
        for idx,ce in enumerate(ce_rows):
            if cf_rows:
                cf_id=cf_rows[idx % len(cf_rows)]["id"]
                cur.execute("INSERT OR IGNORE INTO competencia_fundamental_especifica(fundamental_id,especifica_id) VALUES (?,?)", (cf_id,ce["id"]))

    # Migrar estudiantes V4: si curso/sección coinciden con una clase, se crea la inscripción.
    try:
        cur.execute("""
            UPDATE estudiantes
            SET curso_id=(
                SELECT c.id FROM cursos c
                WHERE c.grado=estudiantes.curso AND c.seccion=estudiantes.seccion
                ORDER BY CASE WHEN c.anio_escolar=(SELECT valor FROM configuracion WHERE clave='anio_escolar') THEN 0 ELSE 1 END, c.id DESC
                LIMIT 1
            )
            WHERE curso_id IS NULL AND curso IS NOT NULL AND seccion IS NOT NULL
        """)
        cur.execute("""
            INSERT OR IGNORE INTO estudiante_curso(estudiante_id,curso_id,fecha_inicio,fecha_fin)
            SELECT e.id,e.curso_id,substr(e.fecha_registro,1,10),NULL
            FROM estudiantes e WHERE e.curso_id IS NOT NULL
        """)
        rows=cur.execute("SELECT id,curso_id FROM estudiantes WHERE curso_id IS NOT NULL AND (orden_lista IS NULL OR orden_lista=0) ORDER BY curso_id,nombre COLLATE NOCASE,id").fetchall()
        counters={}
        for r in rows:
            counters[r["curso_id"]]=counters.get(r["curso_id"],0)+1
            cur.execute("UPDATE estudiantes SET orden_lista=? WHERE id=?",(counters[r["curso_id"]],r["id"]))
    except Exception:
        pass

    conn.commit()
    conn.close()

# Si hay una base V4 y todavía no existe la V5, migramos la base completa.
# Esto permite abrir el proyecto anterior sin perder estudiantes, cursos ni notas.
V4_DB_CANDIDATES = [
    Path("/content/drive/MyDrive/Gestion_Clase_V4/gestion_clase_v4.db"),
    Path("/content/Gestion_Clase_V4/gestion_clase_v4.db"),
]
if not DB_PATH.exists():
    for old_db in V4_DB_CANDIDATES:
        if old_db.exists():
            shutil.copy2(old_db, DB_PATH)
            break

init_db()


# ============================================================
# 3. UTILIDADES
# ============================================================

def db_df(query, params=()):
    conn = get_conn()
    try:
        return pd.read_sql_query(query, conn, params=params)
    finally:
        conn.close()

def execute(query, params=()):
    conn = get_conn()
    try:
        cur = conn.execute(query, params)
        conn.commit()
        return cur.lastrowid
    finally:
        conn.close()

def get_config(key, default=None):
    conn = get_conn()
    row = conn.execute(
        "SELECT valor FROM configuracion WHERE clave=?", (key,)
    ).fetchone()
    conn.close()
    return row["valor"] if row else default

def set_config(key, value):
    execute(
        """INSERT INTO configuracion(clave,valor) VALUES (?,?)
           ON CONFLICT(clave) DO UPDATE SET valor=excluded.valor""",
        (key, str(value))
    )

def clean_number(x):
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return None
    s = str(x).strip().replace(",", ".")
    if not s:
        return None
    try:
        return float(s)
    except Exception:
        return None

def valid_grade(x):
    n = clean_number(x)
    if n is None:
        return None
    mx = float(get_config("max_calificacion", "100"))
    if n < 0 or n > mx:
        raise ValueError(f"La calificación debe estar entre 0 y {mx}.")
    return n

def student_choices(course_id=None):
    params=[]; where="e.estado='Activo'"
    if course_id:
        where += " AND ec.curso_id=? AND ec.fecha_fin IS NULL"
        params.append(int(course_id))
    df=db_df(f"""
        SELECT e.id, e.matricula || ' - ' || e.nombre AS etiqueta
        FROM estudiantes e
        {"JOIN estudiante_curso ec ON ec.estudiante_id=e.id" if course_id else ""}
        WHERE {where}
        ORDER BY e.nombre COLLATE NOCASE, e.id
    """, params)
    return [(r["etiqueta"], int(r["id"])) for _,r in df.iterrows()]

def subject_choices(course_id=None):
    cid=course_id_from_choice(course_id) if 'course_id_from_choice' in globals() else None
    if cid:
        mapped=db_df("SELECT COUNT(*) n FROM curso_asignatura WHERE curso_id=? AND activa=1",(cid,)).iloc[0]["n"]
        if int(mapped)>0:
            df=db_df("""
                SELECT a.id,a.nombre || ' - ' || COALESCE(a.grado,'') AS etiqueta
                FROM asignaturas a JOIN curso_asignatura ca ON ca.asignatura_id=a.id
                WHERE a.activa=1 AND ca.curso_id=? AND ca.activa=1
                ORDER BY a.grado,a.nombre
            """,(cid,))
        else:
            df=db_df("""SELECT id,nombre || ' - ' || COALESCE(grado,'') AS etiqueta FROM asignaturas WHERE activa=1 ORDER BY grado,nombre""")
    else:
        df = db_df("""
            SELECT id, nombre || ' - ' || COALESCE(grado,'') AS etiqueta
            FROM asignaturas WHERE activa=1 ORDER BY grado, nombre
        """)
    return [(r["etiqueta"], int(r["id"])) for _, r in df.iterrows()]

def course_choices():
    df = db_df("""
        SELECT id, nombre || ' - ' || COALESCE(grado,'') || ' ' ||
               COALESCE(seccion,'') || ' | ' || COALESCE(anio_escolar,'') AS etiqueta
        FROM cursos WHERE estado='Activo'
        ORDER BY anio_escolar DESC, grado, seccion, nombre
    """)
    return [(r["etiqueta"], int(r["id"])) for _, r in df.iterrows()]

def course_id_from_choice(choice):
    if choice is None:
        return None
    if isinstance(choice, int):
        return choice
    try:
        return int(str(choice).split(" - ", 1)[0])
    except Exception:
        return None

def student_id_from_choice(choice):
    if choice is None:
        return None
    if isinstance(choice, int):
        return choice
    try:
        return int(str(choice).split(" - ", 1)[0])
    except Exception:
        return None

def subject_id_from_choice(choice):
    if choice is None:
        return None
    if isinstance(choice, int):
        return choice
    try:
        return int(str(choice).split(" - ", 1)[0])
    except Exception:
        return None

def refresh_all(course_id=None):
    students=gr.update(choices=student_choices(course_id))
    subjects=gr.update(choices=subject_choices(course_id))
    return (students,students,students,subjects,subjects,subjects,students,subjects,subjects)


# ============================================================
# 4. AUTENTICACIÓN
# ============================================================

SESSION = {"user_id": None, "usuario": None, "rol": None, "nombre": None}

def login(usuario, password):
    conn = get_conn()
    row = conn.execute(
        "SELECT * FROM usuarios WHERE usuario=? AND activo=1",
        (str(usuario).strip(),)
    ).fetchone()
    conn.close()
    if row and verify_password(password, row["password_hash"]):
        SESSION.update({
            "user_id": row["id"],
            "usuario": row["usuario"],
            "rol": row["rol"],
            "nombre": row["nombre"]
        })
        return (
            gr.update(visible=False),
            gr.update(visible=True),
            f"Sesión iniciada: **{row['nombre']}** ({row['rol']})."
        )
    return (
        gr.update(visible=True),
        gr.update(visible=False),
        "❌ Usuario o contraseña incorrectos."
    )

def logout():
    SESSION.update({"user_id": None, "usuario": None, "rol": None, "nombre": None})
    return gr.update(visible=True), gr.update(visible=False), "Sesión cerrada."


# ============================================================
# 5. USUARIOS
# ============================================================

def crear_usuario(usuario, nombre, rol, password):
    if not usuario or not nombre or not password:
        return "Completa usuario, nombre y contraseña."
    try:
        execute(
            """INSERT INTO usuarios(usuario,nombre,rol,password_hash,fecha_registro)
               VALUES (?,?,?,?,?)""",
            (usuario.strip(), nombre.strip(), rol, hash_password(password), now_str())
        )
        return "✅ Usuario creado."
    except sqlite3.IntegrityError:
        return "❌ Ese usuario ya existe."

def listar_usuarios():
    return db_df("""
        SELECT id, usuario, nombre, rol,
               CASE activo WHEN 1 THEN 'Activo' ELSE 'Inactivo' END AS estado,
               fecha_registro
        FROM usuarios ORDER BY id
    """)

def cambiar_estado_usuario(usuario_id, estado):
    execute("UPDATE usuarios SET activo=? WHERE id=?", (1 if estado=="Activo" else 0, usuario_id))
    return listar_usuarios()


# ============================================================
# 6. DOCENTES / CURSOS / ASIGNATURAS
# ============================================================

def es_admin():
    return SESSION.get("rol") == "admin"

def crear_docente(nombre, cedula, correo, telefono, especialidad):
    if not es_admin():
        return "❌ Solo un usuario administrador puede registrar docentes."
    if not nombre:
        return "Escribe el nombre del docente."
    execute(
        """INSERT INTO docentes(nombre,cedula,correo,telefono,especialidad,fecha_registro)
           VALUES (?,?,?,?,?,?)""",
        (nombre.strip(), cedula, correo, telefono, especialidad, now_str())
    )
    return "✅ Docente registrado."

def listar_docentes():
    return db_df("""
        SELECT id,nombre,cedula,correo,telefono,especialidad,estado
        FROM docentes ORDER BY nombre COLLATE NOCASE
    """)

def cargar_docente(docente_id):
    if not docente_id:
        return "", "", "", "", "", "Activo", "Selecciona un docente."
    df=db_df("SELECT nombre,cedula,correo,telefono,especialidad,estado FROM docentes WHERE id=?",(int(docente_id),))
    if df.empty:
        return "", "", "", "", "", "Activo", "Docente no encontrado."
    r=df.iloc[0]
    return r["nombre"],r["cedula"],r["correo"],r["telefono"],r["especialidad"],r["estado"],"✅ Docente cargado."

def actualizar_docente(docente_id, nombre, cedula, correo, telefono, especialidad, estado):
    if not es_admin():
        return "❌ Solo un usuario administrador puede modificar docentes."
    if not docente_id or not nombre:
        return "Selecciona un docente y escribe el nombre."
    execute(
        """UPDATE docentes SET nombre=?,cedula=?,correo=?,telefono=?,especialidad=?,estado=? WHERE id=?""",
        (nombre.strip(),cedula,correo,telefono,especialidad,estado or "Activo",int(docente_id))
    )
    return "✅ Docente actualizado."

def eliminar_docente(docente_id):
    if not es_admin():
        return "❌ Solo un usuario administrador puede desactivar docentes."
    if not docente_id:
        return "Selecciona un docente."
    # No se borra físicamente: se conserva el historial y las asignaciones quedan sin docente.
    execute("UPDATE docentes SET estado='Inactivo' WHERE id=?",(int(docente_id),))
    return "✅ Docente desactivado. Se conserva su historial."

def docente_choices():
    df=listar_docentes()
    if df.empty:
        return []
    return [f"{int(r.id)} — {r.nombre} ({r.especialidad or 'Sin especialidad'})" for r in df.itertuples()]

def docente_id_from_choice(value):
    if value is None or value == "": return None
    try: return int(str(value).split("—",1)[0].strip())
    except Exception:
        try: return int(value)
        except Exception: return None


def crear_curso(nombre, nivel, grado, seccion, anio, docente_id):
    if not nombre or not grado or not seccion or not anio:
        return "Completa nombre, grado, sección y año escolar."
    execute(
        """INSERT INTO cursos(nombre,nivel,grado,seccion,anio_escolar,docente_id)
           VALUES (?,?,?,?,?,?)""",
        (str(nombre).strip(), nivel or "Secundaria", str(grado).strip(), str(seccion).strip(), str(anio).strip(), docente_id or None)
    )
    return "✅ Clase creada."

def actualizar_curso(curso_id, nombre, nivel, grado, seccion, anio, docente_id, estado="Activo"):
    if not curso_id:
        return "Selecciona una clase para actualizar."
    execute(
        """UPDATE cursos SET nombre=?,nivel=?,grado=?,seccion=?,anio_escolar=?,docente_id=?,estado=? WHERE id=?""",
        (str(nombre).strip(), nivel or "Secundaria", str(grado).strip(), str(seccion).strip(), str(anio).strip(), docente_id or None, estado or "Activo", int(curso_id))
    )
    return "✅ Clase actualizada."

def eliminar_curso(curso_id):
    if not curso_id:
        return "Selecciona una clase."
    # No se borran estudiantes: se cierra su asignación para conservar historial.
    execute("UPDATE estudiante_curso SET fecha_fin=? WHERE curso_id=? AND fecha_fin IS NULL", (today_str(), int(curso_id)))
    execute("UPDATE cursos SET estado='Inactivo' WHERE id=?", (int(curso_id),))
    return "✅ Clase desactivada. Los estudiantes y su historial se conservaron."

def eliminar_curso_definitivo(curso_id, confirmar=False):
    cid=course_id_from_choice(curso_id)
    if not cid:
        return "❌ Selecciona una clase."
    if not confirmar:
        return "⚠️ Marca la casilla de confirmación para eliminar definitivamente la clase."
    # La eliminación definitiva borra solo la clase y sus relaciones de matrícula/asignaturas.
    # Los estudiantes, asistencia, incidencias y calificaciones permanecen en la base de datos.
    try:
        conn=get_conn()
        conn.execute("UPDATE estudiantes SET curso_id=NULL, orden_lista=NULL WHERE curso_id=?", (cid,))
        conn.execute("DELETE FROM estudiante_curso WHERE curso_id=?", (cid,))
        conn.execute("DELETE FROM curso_asignatura WHERE curso_id=?", (cid,))
        conn.execute("DELETE FROM cursos WHERE id=?", (cid,))
        conn.commit()
        conn.close()
        return "🗑️ Clase eliminada definitivamente. Los estudiantes y sus registros personales se conservaron."
    except Exception as e:
        try: conn.close()
        except Exception: pass
        return f"❌ No se pudo eliminar la clase: {e}"

def listar_cursos():
    return db_df("""
        SELECT c.id,c.nombre,c.nivel,c.grado,c.seccion,c.anio_escolar,
               COALESCE(d.nombre,'') AS docente,c.estado,
               (SELECT COUNT(*) FROM estudiante_curso ec WHERE ec.curso_id=c.id AND ec.fecha_fin IS NULL) AS estudiantes
        FROM cursos c LEFT JOIN docentes d ON d.id=c.docente_id
        ORDER BY c.anio_escolar DESC,c.grado,c.seccion,c.nombre
    """)

def cargar_curso(curso_id):
    if not curso_id:
        return "", "Secundaria", "", "", get_config("anio_escolar","2026-2027"), None, "Activo", "Selecciona una clase."
    df=db_df("SELECT * FROM cursos WHERE id=?", (int(curso_id),))
    if df.empty:
        return "", "Secundaria", "", "", "", None, "Activo", "No encontrada."
    r=df.iloc[0]
    return r["nombre"],r["nivel"],r["grado"],r["seccion"],r["anio_escolar"],r["docente_id"],r["estado"],"✅ Clase cargada."

def listar_estudiantes_clase(course_id):
    cid=course_id_from_choice(course_id)
    if not cid: return pd.DataFrame()
    return db_df("""
        SELECT e.id,e.matricula,e.nombre,e.fecha_nacimiento,
               CASE WHEN e.fecha_nacimiento IS NULL OR e.fecha_nacimiento='' THEN ''
                    ELSE CAST((julianday('now')-julianday(e.fecha_nacimiento))/365.2425 AS INTEGER) END AS edad_actual,
               e.edad_registro,e.correo,e.telefono,e.estado,ec.fecha_inicio
        FROM estudiantes e JOIN estudiante_curso ec ON ec.estudiante_id=e.id
        WHERE ec.curso_id=? AND ec.fecha_fin IS NULL
        ORDER BY e.nombre COLLATE NOCASE,e.id
    """, (cid,))

def crear_asignatura(nombre, codigo, grado, descripcion):
    if not nombre:
        return "Escribe el nombre de la asignatura."
    try:
        execute(
            """INSERT INTO asignaturas(nombre,codigo,grado,descripcion)
               VALUES (?,?,?,?)""",
            (nombre,codigo,grado,descripcion)
        )
        return "✅ Asignatura creada."
    except sqlite3.IntegrityError:
        return "❌ Ya existe una asignatura con ese nombre y grado."

def listar_asignaturas():
    return db_df("""
        SELECT id,nombre,codigo,grado,descripcion,
               CASE activa WHEN 1 THEN 'Activa' ELSE 'Inactiva' END AS estado
        FROM asignaturas ORDER BY grado,nombre
    """)

def asignar_asignatura_curso(curso_id, asignatura_id, docente_id=None):
    cid=course_id_from_choice(curso_id); aid=subject_id_from_choice(asignatura_id); did=docente_id_from_choice(docente_id)
    if not cid or not aid: return "Selecciona una clase y una asignatura."
    if did:
        row=db_df("SELECT id FROM docentes WHERE id=? AND estado='Activo'",(did,))
        if row.empty: return "❌ El docente seleccionado no está activo o no existe."
    execute("INSERT OR IGNORE INTO curso_asignatura(curso_id,asignatura_id,docente_id,activa) VALUES (?,?,?,1)",(cid,aid,did))
    execute("UPDATE curso_asignatura SET activa=1,docente_id=? WHERE curso_id=? AND asignatura_id=?",(did,cid,aid))
    return "✅ Asignatura asociada a la clase con el docente seleccionado."

def quitar_asignatura_curso(curso_id, asignatura_id):
    cid=course_id_from_choice(curso_id); aid=subject_id_from_choice(asignatura_id)
    if not cid or not aid: return "Selecciona una clase y una asignatura."
    execute("UPDATE curso_asignatura SET activa=0 WHERE curso_id=? AND asignatura_id=?",(cid,aid))
    return "✅ Asociación desactivada."

def listar_asignaturas_curso(curso_id):
    cid=course_id_from_choice(curso_id)
    if not cid: return pd.DataFrame()
    return db_df("""
        SELECT a.id,a.nombre,a.codigo,a.grado,COALESCE(d.nombre,'Sin docente') AS docente,
               COALESCE(d.especialidad,'') AS especialidad,
               CASE ca.activa WHEN 1 THEN 'Activa' ELSE 'Inactiva' END estado
        FROM curso_asignatura ca
        JOIN asignaturas a ON a.id=ca.asignatura_id
        LEFT JOIN docentes d ON d.id=ca.docente_id
        WHERE ca.curso_id=? ORDER BY a.grado,a.nombre
    """,(cid,))


# ============================================================
# 7. ESTUDIANTES — V5: estudiantes independientes + inscripción a clases + matrícula automática
# ============================================================

def normalizar_inicial(texto):
    texto = (texto or "").strip()
    if not texto: return ""
    return texto[0].upper()

def nombre_completo(nombres, apellidos):
    return " ".join([str(nombres or "").strip(), str(apellidos or "").strip()]).strip()

def calcular_edad(fecha_nacimiento, fecha_referencia=None):
    if not fecha_nacimiento: return None
    try:
        fn=datetime.strptime(str(fecha_nacimiento), "%Y-%m-%d").date()
        ref=fecha_referencia or date.today()
        edad=ref.year-fn.year-((ref.month,ref.day)<(fn.month,fn.day))
        return edad if edad>=0 else None
    except Exception:
        return None

def siguiente_orden_clase(course_id):
    cid=course_id_from_choice(course_id)
    if not cid: return 1
    row=db_df("SELECT COALESCE(MAX(orden_lista),0)+1 AS n FROM estudiantes e JOIN estudiante_curso ec ON ec.estudiante_id=e.id WHERE ec.curso_id=?", (cid,))
    return int(row.iloc[0]["n"]) if not row.empty else 1

def generar_matricula(nombres, apellidos, fecha_nacimiento, course_id):
    cid=course_id_from_choice(course_id)
    if not nombres or not apellidos or not fecha_nacimiento or not cid:
        return ""
    inicial_nombre=normalizar_inicial(nombres)
    inicial_apellido=normalizar_inicial(apellidos)
    try: anio=datetime.strptime(str(fecha_nacimiento), "%Y-%m-%d").year
    except Exception: return ""
    orden=siguiente_orden_clase(cid)
    return f"{inicial_nombre}{inicial_apellido}{anio}-{orden}"

def _matricula_unica(base):
    if not base: return ""
    n=1; candidato=base
    while not db_df("SELECT id FROM estudiantes WHERE matricula=?", (candidato,)).empty:
        n += 1
        candidato=f"{base}-{n}"
    return candidato

def guardar_estudiante(nombres, apellidos, fecha_nacimiento, curso_id, correo, telefono, estado="Activo"):
    cid=course_id_from_choice(curso_id)
    if not nombres or not apellidos or not fecha_nacimiento or not cid:
        return "❌ Completa nombres, apellidos, fecha de nacimiento y clase.", ""
    edad=calcular_edad(fecha_nacimiento)
    if edad is None:
        return "❌ La fecha de nacimiento no es válida.", ""
    curso=db_df("SELECT * FROM cursos WHERE id=? AND estado='Activo'", (cid,))
    if curso.empty:
        return "❌ La clase seleccionada no está activa.", ""
    nombre=nombre_completo(nombres,apellidos)
    base=generar_matricula(nombres,apellidos,fecha_nacimiento,cid)
    matricula=_matricula_unica(base)
    orden=siguiente_orden_clase(cid)
    try:
        sid=execute(
            """INSERT INTO estudiantes
               (matricula,nombre,curso,seccion,correo,telefono,estado,fecha_registro,fecha_nacimiento,edad_registro,curso_id,orden_lista)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
            (matricula,nombre,curso.iloc[0]["grado"],curso.iloc[0]["seccion"],correo or "",telefono or "",estado or "Activo",now_str(),str(fecha_nacimiento),edad,cid,orden)
        )
        execute("INSERT OR IGNORE INTO estudiante_curso(estudiante_id,curso_id,fecha_inicio,fecha_fin) VALUES (?,?,?,NULL)", (sid,cid,today_str()))
        return f"✅ Estudiante registrado. Matrícula automática: **{matricula}** | Edad al registrar: **{edad} años**.", matricula
    except sqlite3.IntegrityError as e:
        return f"❌ No se pudo registrar: {e}", ""

def buscar_estudiantes(texto="", course_id=None):
    texto=texto or ""; like=f"%{texto}%"; params=[like,like,like,like]
    join=""; extra=""
    cid=course_id_from_choice(course_id)
    if cid:
        join="JOIN estudiante_curso ec ON ec.estudiante_id=e.id"
        extra=" AND ec.curso_id=? AND ec.fecha_fin IS NULL"; params.append(cid)
    return db_df(f"""
        SELECT e.id,e.matricula,e.nombre,e.fecha_nacimiento,
               CASE WHEN e.fecha_nacimiento IS NULL OR e.fecha_nacimiento='' THEN ''
                    ELSE CAST((julianday('now')-julianday(e.fecha_nacimiento))/365.2425 AS INTEGER) END AS edad_actual,
               e.edad_registro,e.curso,e.seccion,e.correo,e.telefono,e.estado,e.fecha_registro
        FROM estudiantes e {join}
        WHERE (e.matricula LIKE ? OR e.nombre LIKE ? OR e.curso LIKE ? OR e.seccion LIKE ?){extra}
        ORDER BY e.nombre COLLATE NOCASE,e.id
    """, params)

def cargar_estudiante(matricula):
    if not matricula:
        return "", "", "", None, "", "", "", "Activo", "", "Selecciona una matrícula."
    df=db_df("SELECT * FROM estudiantes WHERE matricula=?", (str(matricula).strip(),))
    if df.empty:
        return "", "", "", None, "", "", "", "Activo", "", "No encontrado."
    r=df.iloc[0]
    partes=str(r["nombre"] or "").split()
    nombres=" ".join(partes[:-1]) if len(partes)>1 else str(r["nombre"] or "")
    apellidos=partes[-1] if len(partes)>1 else ""
    return r["matricula"],nombres,apellidos,r["curso_id"],r["fecha_nacimiento"],r["correo"],r["telefono"],r["estado"],calcular_edad(r["fecha_nacimiento"]),"✅ Cargado."

def actualizar_estudiante(matricula, nombres, apellidos, curso_id, fecha_nacimiento, correo, telefono, estado):
    if not matricula or not nombres or not apellidos or not fecha_nacimiento or not curso_id:
        return "Completa matrícula, nombres, apellidos, fecha de nacimiento y clase."
    sid_df=db_df("SELECT id,curso_id FROM estudiantes WHERE matricula=?", (str(matricula).strip(),))
    if sid_df.empty: return "❌ Estudiante no encontrado."
    sid=int(sid_df.iloc[0]["id"]); nuevo_cid=course_id_from_choice(curso_id); edad=calcular_edad(fecha_nacimiento); nombre=nombre_completo(nombres,apellidos)
    if edad is None: return "❌ Fecha de nacimiento inválida."
    viejo=sid_df.iloc[0]["curso_id"]
    execute("UPDATE estudiantes SET nombre=?,curso=?,seccion=?,correo=?,telefono=?,estado=?,fecha_nacimiento=?,edad_registro=?,curso_id=? WHERE matricula=?",
            (nombre, db_df("SELECT grado FROM cursos WHERE id=?",(nuevo_cid,)).iloc[0]["grado"], db_df("SELECT seccion FROM cursos WHERE id=?",(nuevo_cid,)).iloc[0]["seccion"], correo or "",telefono or "",estado or "Activo",str(fecha_nacimiento),edad,nuevo_cid,str(matricula).strip()))
    if viejo and int(viejo)!=int(nuevo_cid):
        execute("UPDATE estudiante_curso SET fecha_fin=? WHERE estudiante_id=? AND curso_id=? AND fecha_fin IS NULL", (today_str(),sid,int(viejo)))
        nuevo_orden=siguiente_orden_clase(nuevo_cid)
        execute("UPDATE estudiantes SET orden_lista=? WHERE id=?",(nuevo_orden,sid))
    execute("UPDATE estudiante_curso SET fecha_fin=NULL WHERE estudiante_id=? AND curso_id=?", (sid,int(nuevo_cid)))
    if db_df("SELECT id FROM estudiante_curso WHERE estudiante_id=? AND curso_id=?",(sid,int(nuevo_cid))).empty:
        execute("INSERT INTO estudiante_curso(estudiante_id,curso_id,fecha_inicio,fecha_fin) VALUES (?,?,?,NULL)",(sid,int(nuevo_cid),today_str()))
    return "✅ Estudiante actualizado."

def eliminar_estudiante(matricula):
    if not matricula: return "Indica la matrícula."
    execute("DELETE FROM estudiantes WHERE matricula=?", (str(matricula).strip(),))
    return "✅ Estudiante eliminado."

def matriculas():
    df=db_df("SELECT matricula FROM estudiantes ORDER BY nombre COLLATE NOCASE")
    return [x for x in df["matricula"].tolist()]

    df = db_df("SELECT matricula FROM estudiantes ORDER BY nombre")
    return [x for x in df["matricula"].tolist()]


# ============================================================
# 8. ASISTENCIA
# ============================================================

ESTADOS_ASISTENCIA = ["Presente","Ausente","Tardanza","Excusa/Permiso"]

def asistencia_del_dia(fecha, curso_id=None):
    cid=course_id_from_choice(curso_id)
    where = ""
    params = [fecha]
    if cid:
        where += " AND ec.curso_id=? AND ec.fecha_fin IS NULL"
        params.append(cid)

    return db_df(f"""
        SELECT e.id,e.matricula,e.nombre,e.curso,e.seccion,
               COALESCE(a.estado,'Sin registrar') AS estado,
               COALESCE(a.hora_entrada,'') AS hora_entrada,
               COALESCE(a.observacion,'') AS observacion
        FROM estudiantes e
        LEFT JOIN asistencia a
          ON a.estudiante_id=e.id AND a.fecha=?
        WHERE e.estado='Activo' {where}
        ORDER BY e.nombre
    """, params)

def guardar_asistencia(fecha, tabla):
    if tabla is None:
        return "No hay datos."
    if isinstance(tabla, pd.DataFrame):
        df = tabla.copy()
    else:
        df = pd.DataFrame(tabla)
    if df.empty:
        return "No hay estudiantes."
    conn = get_conn()
    count = 0
    for _, r in df.iterrows():
        estado = str(r.get("estado","")).strip()
        if estado not in ESTADOS_ASISTENCIA:
            continue
        conn.execute(
            """INSERT INTO asistencia(estudiante_id,fecha,estado,hora_entrada,observacion)
               VALUES (?,?,?,?,?)
               ON CONFLICT(estudiante_id,fecha)
               DO UPDATE SET estado=excluded.estado,
                             hora_entrada=excluded.hora_entrada,
                             observacion=excluded.observacion""",
            (
                int(r["id"]), fecha, estado,
                str(r.get("hora_entrada","")) if pd.notna(r.get("hora_entrada","")) else "",
                str(r.get("observacion","")) if pd.notna(r.get("observacion","")) else ""
            )
        )
        count += 1
    conn.commit()
    conn.close()
    return f"✅ {count} registros de asistencia guardados."

def marcar_todos(fecha, estado, curso_id=None):
    df = asistencia_del_dia(fecha, curso_id)
    if df.empty:
        return df, "No hay estudiantes."
    df["estado"] = estado
    if estado == "Tardanza":
        df["hora_entrada"] = datetime.now().strftime("%H:%M")
    return df, f"Marcados como {estado}. Pulsa Guardar."

def porcentaje_asistencia(course_id=None):
    cid=course_id_from_choice(course_id)
    params=[]; extra=""
    if cid:
        extra=" JOIN estudiante_curso ec ON ec.estudiante_id=e.id AND ec.curso_id=? AND ec.fecha_fin IS NULL"; params=[cid]
    return db_df(f"""
        SELECT e.matricula,e.nombre,e.curso,e.seccion,
               COUNT(a.id) AS dias_registrados,
               SUM(CASE WHEN a.estado='Presente' THEN 1 ELSE 0 END) AS presentes,
               SUM(CASE WHEN a.estado='Tardanza' THEN 1 ELSE 0 END) AS tardanzas,
               SUM(CASE WHEN a.estado='Ausente' THEN 1 ELSE 0 END) AS ausencias,
               ROUND(
                 100.0 * SUM(CASE WHEN a.estado IN ('Presente','Tardanza') THEN 1 ELSE 0 END)
                 / NULLIF(COUNT(a.id),0), 2
               ) AS porcentaje
        FROM estudiantes e{extra}
        LEFT JOIN asistencia a ON a.estudiante_id=e.id
        GROUP BY e.id
        ORDER BY porcentaje ASC, e.nombre COLLATE NOCASE
    """,params)


# ============================================================
# 9. INCIDENCIAS
# ============================================================

def guardar_incidencia(estudiante, fecha, hora, tipo, gravedad, descripcion, accion, observacion):
    sid = student_id_from_choice(estudiante)
    if not sid or not tipo:
        return "Selecciona estudiante y tipo."
    execute(
        """INSERT INTO incidencias
           (estudiante_id,fecha,hora,tipo,gravedad,descripcion,accion_tomada,observacion)
           VALUES (?,?,?,?,?,?,?,?)""",
        (sid,fecha,hora,tipo,gravedad,descripcion,accion,observacion)
    )
    return "✅ Incidencia registrada."

def listar_incidencias(course_id=None):
    return db_df("""
        SELECT i.id,e.matricula,e.nombre,e.curso,e.seccion,
               i.fecha,i.hora,i.tipo,i.gravedad,i.descripcion,
               i.accion_tomada,i.observacion
        FROM incidencias i JOIN estudiantes e ON e.id=i.estudiante_id
        ORDER BY i.fecha DESC,i.hora DESC
    """)


# ============================================================
# 10. PERMISOS
# ============================================================

def guardar_permiso(estudiante, fecha, tipo, desde, hasta, motivo, estado, observacion):
    sid = student_id_from_choice(estudiante)
    if not sid or not tipo:
        return "Selecciona estudiante y tipo."
    execute(
        """INSERT INTO permisos
           (estudiante_id,fecha,tipo,hora_desde,hora_hasta,motivo,estado,observacion)
           VALUES (?,?,?,?,?,?,?,?)""",
        (sid,fecha,tipo,desde,hasta,motivo,estado,observacion)
    )
    return "✅ Permiso registrado."

def listar_permisos(course_id=None):
    return db_df("""
        SELECT p.id,e.matricula,e.nombre,e.curso,e.seccion,
               p.fecha,p.tipo,p.hora_desde,p.hora_hasta,p.motivo,
               p.estado,p.observacion
        FROM permisos p JOIN estudiantes e ON e.id=p.estudiante_id
        ORDER BY p.fecha DESC
    """)

def cambiar_estado_permiso(permiso_id, estado):
    execute("UPDATE permisos SET estado=? WHERE id=?", (estado, permiso_id))
    return listar_permisos()


# ============================================================
# 11. COMPETENCIAS
# ============================================================

def listar_competencias(asignatura):
    aid = subject_id_from_choice(asignatura)
    if not aid:
        return pd.DataFrame(columns=["id","codigo","nombre","descripcion","orden","estado"])
    return db_df("""
        SELECT id,codigo,nombre,descripcion,orden,
               CASE activa WHEN 1 THEN 'Activa' ELSE 'Inactiva' END AS estado
        FROM competencias WHERE asignatura_id=?
        ORDER BY orden,id
    """, (aid,))

def crear_competencia(asignatura, codigo, nombre, descripcion, orden):
    aid = subject_id_from_choice(asignatura)
    if not aid or not codigo or not nombre or not descripcion:
        return "Completa asignatura, código, nombre y descripción."
    try:
        execute(
            """INSERT INTO competencias
               (asignatura_id,codigo,nombre,descripcion,orden)
               VALUES (?,?,?,?,?)""",
            (aid,codigo,nombre,descripcion,int(orden or 1))
        )
        return "✅ Competencia creada."
    except sqlite3.IntegrityError:
        return "❌ Ese código ya existe en la asignatura."

def actualizar_competencia(comp_id, codigo, nombre, descripcion, orden, activa):
    if not comp_id:
        return "Indica el ID de la competencia."
    execute(
        """UPDATE competencias
           SET codigo=?,nombre=?,descripcion=?,orden=?,activa=?
           WHERE id=?""",
        (codigo,nombre,descripcion,int(orden or 1),1 if activa else 0,comp_id)
    )
    return "✅ Competencia actualizada."

def cargar_competencia(comp_id):
    df = db_df("SELECT * FROM competencias WHERE id=?", (comp_id,))
    if df.empty:
        return "", "", "", "", True, "No encontrada."
    r=df.iloc[0]
    return r["codigo"],r["nombre"],r["descripcion"],int(r["orden"]),bool(r["activa"]),"✅ Cargada."


# ============================================================
# 12. CALIFICACIONES
# ============================================================

PERIODOS = [1,2,3,4]

def get_competencias_ids(asignatura):
    aid = subject_id_from_choice(asignatura)
    if not aid:
        return []
    return db_df("""
        SELECT id,codigo,nombre,descripcion,orden
        FROM competencias WHERE asignatura_id=? AND activa=1
        ORDER BY orden,id
    """, (aid,))

def guardar_calificacion(estudiante, asignatura, competencia, periodo, tipo, calificacion, observacion):
    sid = student_id_from_choice(estudiante)
    aid = subject_id_from_choice(asignatura)
    if not sid or not aid or not competencia:
        return "Selecciona estudiante, asignatura y competencia."
    try:
        nota = valid_grade(calificacion)
    except ValueError as e:
        return f"❌ {e}"
    if nota is None:
        # permite borrar una nota
        execute(
            """DELETE FROM calificaciones
               WHERE estudiante_id=? AND asignatura_id=? AND competencia_id=?
                 AND periodo=? AND tipo=?""",
            (sid,aid,int(competencia),int(periodo),tipo)
        )
        return "🗑️ Calificación eliminada."
    execute(
        """INSERT INTO calificaciones
           (estudiante_id,asignatura_id,competencia_id,periodo,tipo,
            calificacion,observacion,fecha_actualizacion)
           VALUES (?,?,?,?,?,?,?,?)
           ON CONFLICT(estudiante_id,asignatura_id,competencia_id,periodo,tipo)
           DO UPDATE SET calificacion=excluded.calificacion,
                         observacion=excluded.observacion,
                         fecha_actualizacion=excluded.fecha_actualizacion""",
        (sid,aid,int(competencia),int(periodo),tipo,nota,observacion,now_str())
    )
    return "✅ Calificación guardada."

def notas_estudiante(estudiante, asignatura):
    sid = student_id_from_choice(estudiante)
    aid = subject_id_from_choice(asignatura)
    comps = get_competencias_ids(asignatura)
    if not sid or not aid or comps.empty:
        return pd.DataFrame()
    rows=[]
    for _, c in comps.iterrows():
        row={
            "Competencia": f"{c['codigo']} - {c['nombre']}",
            "Descripción": c["descripcion"]
        }
        for p in PERIODOS:
            for t in ["P","RP"]:
                q=db_df("""
                    SELECT calificacion FROM calificaciones
                    WHERE estudiante_id=? AND asignatura_id=? AND competencia_id=?
                      AND periodo=? AND tipo=?
                """,(sid,aid,int(c["id"]),p,t))
                row[f"P{p}" if t=="P" else f"RP{p}"] = (
                    "" if q.empty or pd.isna(q.iloc[0]["calificacion"])
                    else q.iloc[0]["calificacion"]
                )
        row["Promedio"] = promedio_competencia(sid,aid,int(c["id"]))
        rows.append(row)
    out=pd.DataFrame(rows)
    out["Calificación final"] = promedio_final(sid,aid)
    return out

def promedio_competencia(sid, aid, cid):
    vals=[]
    use_rp = get_config("rp_reemplaza_p","1") == "1"
    for p in PERIODOS:
        p_df=db_df("""
            SELECT tipo,calificacion FROM calificaciones
            WHERE estudiante_id=? AND asignatura_id=? AND competencia_id=? AND periodo=?
        """,(sid,aid,cid,p))
        pval=None; rpval=None
        for _,r in p_df.iterrows():
            if r["tipo"]=="P": pval=r["calificacion"]
            elif r["tipo"]=="RP": rpval=r["calificacion"]
        if use_rp and rpval is not None:
            vals.append(float(rpval))
        else:
            if pval is not None:
                vals.append(float(pval))
            elif rpval is not None:
                vals.append(float(rpval))
    return round(sum(vals)/len(vals),2) if vals else None

def promedio_final(sid, aid):
    comps=get_competencias_ids(aid)
    vals=[]
    for _,c in comps.iterrows():
        x=promedio_competencia(sid,aid,int(c["id"]))
        if x is not None:
            vals.append(x)
    return round(sum(vals)/len(vals),2) if vals else None

def registro_grado(asignatura, course_id=None):
    """Matriz de registro: estudiantes x competencias x P/RP + promedios/final."""
    aid=subject_id_from_choice(asignatura)
    comps=get_competencias_ids(asignatura)
    if not aid or comps.empty:
        return pd.DataFrame()

    cid=course_id_from_choice(course_id)
    if cid:
        students=db_df("""SELECT e.id,e.matricula,e.nombre,e.curso,e.seccion FROM estudiantes e JOIN estudiante_curso ec ON ec.estudiante_id=e.id WHERE e.estado='Activo' AND ec.curso_id=? AND ec.fecha_fin IS NULL ORDER BY e.nombre COLLATE NOCASE,e.id""",(cid,))
    else:
        students=db_df("SELECT id,matricula,nombre,curso,seccion FROM estudiantes WHERE estado='Activo' ORDER BY nombre COLLATE NOCASE,id")
    rows=[]
    for _,s in students.iterrows():
        row={
            "ID": int(s["id"]),
            "Matrícula": s["matricula"],
            "Estudiante": s["nombre"],
            "Curso": s["curso"],
            "Sección": s["seccion"],
        }
        for _,c in comps.iterrows():
            cid=int(c["id"])
            prefix=c["codigo"]
            for p in PERIODOS:
                for t in ["P","RP"]:
                    q=db_df("""
                        SELECT calificacion FROM calificaciones
                        WHERE estudiante_id=? AND asignatura_id=? AND competencia_id=?
                          AND periodo=? AND tipo=?
                    """,(int(s["id"]),aid,cid,p,t))
                    value="" if q.empty or pd.isna(q.iloc[0]["calificacion"]) else q.iloc[0]["calificacion"]
                    row[f"{prefix}_P{p}" if t=="P" else f"{prefix}_RP{p}"] = value
            row[f"{prefix}_Promedio"] = promedio_competencia(int(s["id"]),aid,cid)
        row["Calificación final"] = promedio_final(int(s["id"]),aid)
        rows.append(row)
    return pd.DataFrame(rows)

def guardar_registro_grado(asignatura, tabla):
    aid=subject_id_from_choice(asignatura)
    comps=get_competencias_ids(asignatura)
    if not aid or comps.empty:
        return "Selecciona una asignatura con competencias."
    if tabla is None:
        return "No hay tabla."
    df=tabla if isinstance(tabla,pd.DataFrame) else pd.DataFrame(tabla)
    if df.empty:
        return "No hay estudiantes."

    conn=get_conn()
    count=0
    for _,r in df.iterrows():
        try:
            sid=int(r["ID"])
        except Exception:
            continue
        for _,c in comps.iterrows():
            cid=int(c["id"])
            prefix=c["codigo"]
            for p in PERIODOS:
                for t in ["RP"]:
                    col=f"{prefix}_RP{p}"
                    if col not in df.columns:
                        continue
                    raw=r[col]
                    if pd.isna(raw) or str(raw).strip()=="":
                        conn.execute(
                            """DELETE FROM calificaciones
                               WHERE estudiante_id=? AND asignatura_id=? AND competencia_id=?
                                 AND periodo=? AND tipo=?""",
                            (sid,aid,cid,p,t)
                        )
                        continue
                    try:
                        nota=valid_grade(raw)
                    except Exception:
                        continue
                    conn.execute(
                        """INSERT INTO calificaciones
                           (estudiante_id,asignatura_id,competencia_id,periodo,tipo,
                            calificacion,fecha_actualizacion)
                           VALUES (?,?,?,?,?,?,?)
                           ON CONFLICT(estudiante_id,asignatura_id,competencia_id,periodo,tipo)
                           DO UPDATE SET calificacion=excluded.calificacion,
                                         fecha_actualizacion=excluded.fecha_actualizacion""",
                        (sid,aid,cid,p,t,nota,now_str())
                    )
                    count+=1
    conn.commit()
    conn.close()
    return f"✅ Registro guardado. {count} celdas procesadas."

def opciones_competencias(asignatura):
    comps=get_competencias_ids(asignatura)
    if comps.empty:
        return gr.update(choices=[])
    return gr.update(
        choices=[(f"{r['codigo']} - {r['nombre']}", int(r["id"])) for _,r in comps.iterrows()]
    )

def cargar_nota_actual(estudiante, asignatura, competencia, periodo, tipo):
    sid=student_id_from_choice(estudiante)
    aid=subject_id_from_choice(asignatura)
    if not sid or not aid or not competencia:
        return "", ""
    df=db_df("""
        SELECT calificacion,observacion
        FROM calificaciones
        WHERE estudiante_id=? AND asignatura_id=? AND competencia_id=?
          AND periodo=? AND tipo=?
    """,(sid,aid,int(competencia),int(periodo),tipo))
    if df.empty:
        return "", ""
    return df.iloc[0]["calificacion"], df.iloc[0]["observacion"]


# ============================================================
# 13. COMPETENCIAS FUNDAMENTALES + RELACIONES
# ============================================================

CLIPBOARD = {"tipo": None, "data": None}

INSTRUMENTOS_EVALUACION = [
    "Listas de cotejo", "Rúbrica", "Escalas de estimación", "Registro anecdótico",
    "Diario de clase", "Diario del docente", "Guía de observación", "Prueba objetiva",
    "Examen oral", "Cuestionarios", "Prueba de ejecución práctica", "Portafolio de evidencias",
    "Proyectos de investigación", "Ensayo", "Mapa conceptual/mental", "Cuaderno de trabajo",
    "Reportes técnicos", "Estudios de caso", "Debate", "Exposición Oral",
    "Dramatización/juegos de rol", "Grabación de video/podcast escolar",
    "Coevaluación (evaluación entre pares)", "Plantillas de autoevaluación"
]
TIPOS_EVALUACION = ["Sumativa", "Diagnóstica", "Formativa"]


def fundamental_choices():
    df=db_df("SELECT id,codigo || ' - ' || nombre AS etiqueta FROM competencias_fundamentales WHERE activa=1 ORDER BY orden,id")
    return [(r["etiqueta"],int(r["id"])) for _,r in df.iterrows()]


def todas_fundamentales():
    return db_df("""
        SELECT id,codigo,nombre,descripcion,orden,
               CASE activa WHEN 1 THEN 'Activa' ELSE 'Inactiva' END estado
        FROM competencias_fundamentales ORDER BY orden,id
    """)


def crear_fundamental(codigo,nombre,descripcion,orden):
    if not codigo or not nombre or not descripcion:
        return "Completa código, nombre y descripción."
    try:
        execute("INSERT INTO competencias_fundamentales(codigo,nombre,descripcion,orden) VALUES (?,?,?,?)",(codigo.strip(),nombre.strip(),descripcion.strip(),int(orden or 1)))
        return "✅ Competencia fundamental creada."
    except sqlite3.IntegrityError:
        return "❌ El código de la competencia fundamental ya existe."


def actualizar_fundamental(fid,codigo,nombre,descripcion,orden,activa):
    if not fid:
        return "Indica el ID."
    try:
        execute("UPDATE competencias_fundamentales SET codigo=?,nombre=?,descripcion=?,orden=?,activa=? WHERE id=?",(codigo.strip(),nombre.strip(),descripcion.strip(),int(orden or 1),1 if activa else 0,int(fid)))
        return "✅ Competencia fundamental actualizada."
    except sqlite3.IntegrityError:
        return "❌ El código ya está utilizado."


def cargar_fundamental(fid):
    df=db_df("SELECT * FROM competencias_fundamentales WHERE id=?",(fid,))
    if df.empty:
        return "","","","",True,"No encontrada."
    r=df.iloc[0]
    return r["codigo"],r["nombre"],r["descripcion"],int(r["orden"]),bool(r["activa"]),"✅ Cargada."


def eliminar_fundamental(fid):
    if not fid:
        return "Indica el ID."
    execute("DELETE FROM competencias_fundamentales WHERE id=?",(int(fid),))
    return "🗑️ Competencia fundamental eliminada."


def competencias_con_fundamentales(asignatura):
    aid=subject_id_from_choice(asignatura)
    if not aid:
        return pd.DataFrame()
    return db_df("""
        SELECT c.id,c.codigo,c.nombre,c.descripcion,c.orden,
               COALESCE(GROUP_CONCAT(f.codigo || ' - ' || f.nombre, ' | '),'') AS competencias_fundamentales
        FROM competencias c
        LEFT JOIN competencia_fundamental_especifica m ON m.especifica_id=c.id
        LEFT JOIN competencias_fundamentales f ON f.id=m.fundamental_id
        WHERE c.asignatura_id=?
        GROUP BY c.id ORDER BY c.orden,c.id
    """,(aid,))


def asignar_competencia_fundamental(fid,cid):
    if not fid or not cid:
        return "Selecciona la competencia fundamental y la específica."
    try:
        execute("INSERT OR IGNORE INTO competencia_fundamental_especifica(fundamental_id,especifica_id) VALUES (?,?)",(int(fid),int(cid)))
        return "✅ Relación guardada. La CE conserva una sola descripción."
    except sqlite3.IntegrityError as e:
        return f"❌ No se pudo asignar: {e}"


def quitar_competencia_fundamental(fid,cid):
    if not fid or not cid:
        return "Selecciona ambos elementos."
    execute("DELETE FROM competencia_fundamental_especifica WHERE fundamental_id=? AND especifica_id=?",(int(fid),int(cid)))
    return "🗑️ Relación eliminada."


def listar_mapa_competencias(asignatura,fundamental=None):
    aid=subject_id_from_choice(asignatura)
    if not aid:
        return pd.DataFrame()
    params=[aid]
    extra=""
    if fundamental:
        extra=" AND f.id=?"
        params.append(int(fundamental))
    return db_df(f"""
        SELECT f.codigo AS CF_codigo,f.nombre AS CF_nombre,
               c.id AS CE_id,c.codigo AS CE_codigo,c.nombre AS CE_nombre,c.descripcion
        FROM competencias_fundamentales f
        JOIN competencia_fundamental_especifica m ON m.fundamental_id=f.id
        JOIN competencias c ON c.id=m.especifica_id
        WHERE c.asignatura_id=? {extra}
        ORDER BY f.orden,c.orden,c.id
    """,params)


def copiar_competencia_especifica(cid):
    df=db_df("SELECT codigo,nombre,descripcion,orden,activa FROM competencias WHERE id=?",(cid,))
    if df.empty:
        return "No encontrada.",""
    CLIPBOARD.update({"tipo":"especifica","data":df.iloc[0].to_dict()})
    return f"📋 Copiada: {df.iloc[0]['codigo']}", f"{df.iloc[0]['codigo']} | {df.iloc[0]['nombre']}"


def pegar_competencia_especifica(asignatura,codigo,nombre,descripcion,orden):
    if CLIPBOARD.get("tipo")!="especifica":
        return "No hay una competencia específica copiada."
    d=CLIPBOARD["data"]
    codigo=(codigo or f"{d['codigo']}-COPIA").strip()
    nombre=(nombre or d["nombre"]).strip()
    descripcion=descripcion or d["descripcion"]
    orden=int(orden or d["orden"] or 1)
    return crear_competencia(asignatura,codigo,nombre,descripcion,orden)


def opciones_competencias_fundamentales(asignatura):
    comps=get_competencias_ids(asignatura)
    if isinstance(comps,list) or comps is None or (hasattr(comps,"empty") and comps.empty):
        return gr.update(choices=[])
    return gr.update(choices=[(f"{r['codigo']} - {r['nombre']}",int(r['id'])) for _,r in comps.iterrows()])


# ============================================================
# 14. ACTIVIDADES EVALUATIVAS
# ============================================================

def guardar_actividad(estudiante,asignatura,competencia,periodo,fecha,titulo,descripcion,tipo_evaluacion,instrumento,calificacion,observacion):
    sid=student_id_from_choice(estudiante); aid=subject_id_from_choice(asignatura)
    if not sid or not aid or not competencia or not titulo:
        return "Selecciona estudiante, asignatura, competencia y título."
    try:
        nota=valid_grade(calificacion)
    except ValueError as e:
        return f"❌ {e}"
    if nota is None:
        return "❌ La actividad requiere una calificación."
    if int(periodo) not in PERIODOS:
        return "❌ Período inválido."
    execute("""INSERT INTO actividades
        (estudiante_id,asignatura_id,competencia_id,periodo,fecha,titulo,descripcion,tipo_evaluacion,instrumento,calificacion,observacion,fecha_registro)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",(sid,aid,int(competencia),int(periodo),fecha or today_str(),titulo,descripcion,tipo_evaluacion,instrumento,nota,observacion,now_str()))
    return "✅ Actividad registrada. P1-P4 se recalculan automáticamente."


def listar_actividades(asignatura=None,estudiante=None,periodo=None,course_id=None):
    params=[]; where=[]
    aid=subject_id_from_choice(asignatura); sid=student_id_from_choice(estudiante)
    if aid: where.append("a.asignatura_id=?"); params.append(aid)
    if sid: where.append("a.estudiante_id=?"); params.append(sid)
    if periodo: where.append("a.periodo=?"); params.append(int(periodo))
    cid=course_id_from_choice(course_id)
    join_course=""
    if cid:
        join_course=" JOIN estudiante_curso ec ON ec.estudiante_id=e.id AND ec.curso_id=? AND ec.fecha_fin IS NULL"
        params.append(cid)
    clause=("WHERE "+" AND ".join(where)) if where else ""
    return db_df(f"""
        SELECT a.id,e.matricula,e.nombre,asig.nombre AS asignatura,c.codigo AS competencia,
               a.periodo,a.fecha,a.titulo,a.tipo_evaluacion,a.instrumento,a.calificacion,a.observacion
        FROM actividades a JOIN estudiantes e ON e.id=a.estudiante_id{join_course}
        JOIN asignaturas asig ON asig.id=a.asignatura_id
        JOIN competencias c ON c.id=a.competencia_id
        {clause}
        ORDER BY a.fecha DESC,a.id DESC
    """,params)


def promedio_actividad_competencia(sid,aid,cid,periodo):
    df=db_df("SELECT calificacion FROM actividades WHERE estudiante_id=? AND asignatura_id=? AND competencia_id=? AND periodo=?",(sid,aid,cid,periodo))
    if df.empty: return None
    return round(float(df["calificacion"].mean()),2)


def promedio_actividades_general(asignatura=None,estudiante=None,periodo=None,course_id=None):
    df=listar_actividades(asignatura,estudiante,periodo,course_id)
    if df.empty:
        return pd.DataFrame(columns=["Actividades","Promedio general","Cantidad"])
    return pd.DataFrame([{
        "Actividades":len(df),
        "Promedio general":round(float(df["calificacion"].mean()),2),
        "Cantidad":len(df)
    }])


def resumen_actividades_por_competencia(asignatura,estudiante=None,periodo=None,course_id=None):
    aid=subject_id_from_choice(asignatura); sid=student_id_from_choice(estudiante); cid=course_id_from_choice(course_id)
    if not aid: return pd.DataFrame()
    params=[aid]; extra=""; join_course=""
    if cid:
        join_course=" JOIN estudiante_curso ec ON ec.estudiante_id=e.id AND ec.curso_id=? AND ec.fecha_fin IS NULL"
        params.append(cid)
    if sid: extra+=" AND a.estudiante_id=?"; params.append(sid)
    if periodo: extra+=" AND a.periodo=?"; params.append(int(periodo))
    return db_df(f"""
        SELECT c.codigo,c.nombre,a.periodo,COUNT(*) actividades,ROUND(AVG(a.calificacion),2) promedio
        FROM actividades a JOIN estudiantes e ON e.id=a.estudiante_id{join_course}
        JOIN competencias c ON c.id=a.competencia_id
        WHERE a.asignatura_id=? {extra}
        GROUP BY c.id,a.periodo ORDER BY c.orden,a.periodo
    """,params)


def eliminar_actividad(actividad_id):
    if not actividad_id: return "Indica el ID."
    execute("DELETE FROM actividades WHERE id=?",(int(actividad_id),))
    return "🗑️ Actividad eliminada."


# ============================================================
# 15. CALIFICACIONES POR COMPETENCIAS — P CALCULADA
# ============================================================

def periodo_competencia(sid,aid,cid,periodo):
    # Prioridad: actividades. Si no hay actividades, conserva P histórica de V3.
    avg=promedio_actividad_competencia(sid,aid,cid,periodo)
    if avg is not None: return avg
    df=db_df("SELECT calificacion FROM calificaciones WHERE estudiante_id=? AND asignatura_id=? AND competencia_id=? AND periodo=? AND tipo='P'",(sid,aid,cid,periodo))
    if not df.empty and pd.notna(df.iloc[0]["calificacion"]): return float(df.iloc[0]["calificacion"])
    return None


def recuperacion_competencia(sid,aid,cid,periodo):
    df=db_df("SELECT calificacion FROM calificaciones WHERE estudiante_id=? AND asignatura_id=? AND competencia_id=? AND periodo=? AND tipo='RP'",(sid,aid,cid,periodo))
    if df.empty or pd.isna(df.iloc[0]["calificacion"]): return None
    return float(df.iloc[0]["calificacion"])


def promedio_competencia(sid,aid,cid):
    vals=[]; use_rp=get_config("rp_reemplaza_p","1")=="1"
    for p in PERIODOS:
        pval=periodo_competencia(sid,aid,cid,p); rpval=recuperacion_competencia(sid,aid,cid,p)
        if use_rp and rpval is not None: vals.append(rpval)
        elif pval is not None: vals.append(pval)
        elif rpval is not None: vals.append(rpval)
    return round(sum(vals)/len(vals),2) if vals else None


def guardar_calificacion(estudiante,asignatura,competencia,periodo,tipo,calificacion,observacion):
    if str(tipo)=="P":
        return "ℹ️ La calificación P se calcula automáticamente con las actividades. Registra o modifica las actividades en el módulo Actividades."
    sid=student_id_from_choice(estudiante); aid=subject_id_from_choice(asignatura)
    if not sid or not aid or not competencia: return "Selecciona estudiante, asignatura y competencia."
    try: nota=valid_grade(calificacion)
    except ValueError as e: return f"❌ {e}"
    if nota is None:
        execute("DELETE FROM calificaciones WHERE estudiante_id=? AND asignatura_id=? AND competencia_id=? AND periodo=? AND tipo=?",(sid,aid,int(competencia),int(periodo),tipo))
        return "🗑️ Recuperación eliminada."
    execute("""INSERT INTO calificaciones(estudiante_id,asignatura_id,competencia_id,periodo,tipo,calificacion,observacion,fecha_actualizacion)
        VALUES (?,?,?,?,?,?,?,?)
        ON CONFLICT(estudiante_id,asignatura_id,competencia_id,periodo,tipo) DO UPDATE SET calificacion=excluded.calificacion,observacion=excluded.observacion,fecha_actualizacion=excluded.fecha_actualizacion""",
        (sid,aid,int(competencia),int(periodo),tipo,nota,observacion,now_str()))
    return "✅ RP guardada."


def cargar_nota_actual(estudiante,asignatura,competencia,periodo,tipo):
    sid=student_id_from_choice(estudiante); aid=subject_id_from_choice(asignatura)
    if not sid or not aid or not competencia: return "",""
    if str(tipo)=="P":
        v=periodo_competencia(sid,aid,int(competencia),int(periodo))
        return ("" if v is None else v), "Calculada automáticamente desde actividades."
    v=recuperacion_competencia(sid,aid,int(competencia),int(periodo))
    if v is None: return "",""
    df=db_df("SELECT observacion FROM calificaciones WHERE estudiante_id=? AND asignatura_id=? AND competencia_id=? AND periodo=? AND tipo=?",(sid,aid,int(competencia),int(periodo),tipo))
    return v, ("" if df.empty else df.iloc[0]["observacion"])


# ============================================================
# 16. CALIFICACIONES FINALES / BOLETÍN
# ============================================================

def datos_finales_estudiante(estudiante,asignatura):
    sid=student_id_from_choice(estudiante); aid=subject_id_from_choice(asignatura)
    if not sid or not aid: return None
    cf=promedio_final(sid,aid)
    row=db_df("SELECT * FROM calificaciones_finales WHERE estudiante_id=? AND asignatura_id=?",(sid,aid))
    data={"CF":cf,"completiva":None,"extraordinaria":None,"especial_cf":None,"especial_ce":None,"situacion":"","observacion":""}
    if not row.empty:
        r=row.iloc[0]
        for k in ["completiva","extraordinaria","especial_cf","especial_ce"]: data[k]=r[k]
        data["situacion"]=r["situacion"] or ""; data["observacion"]=r["observacion"] or ""
    return data


def calcular_completiva(cf,cec):
    if cf is None or cec is None: return None
    wcf=float(get_config("peso_completiva_cf","50"))/100
    wcec=float(get_config("peso_completiva_cec","50"))/100
    return round(cf*wcf+cec*wcec,2)


def calcular_extraordinaria(cf,ceex):
    if cf is None or ceex is None: return None
    wcf=float(get_config("peso_extra_cf","30"))/100
    wex=float(get_config("peso_extra_ceex","70"))/100
    return round(cf*wcf+ceex*wex,2)


def guardar_calificacion_final(estudiante,asignatura,cec,ceex,especial_cf,especial_ce,situacion,observacion):
    sid=student_id_from_choice(estudiante); aid=subject_id_from_choice(asignatura)
    if not sid or not aid: return "Selecciona estudiante y asignatura."
    cf=promedio_final(sid,aid)
    try:
        cecv=valid_grade(cec) if cec not in [None,""] else None
        ceexv=valid_grade(ceex) if ceex not in [None,""] else None
        ecf=valid_grade(especial_cf) if especial_cf not in [None,""] else None
        ece=valid_grade(especial_ce) if especial_ce not in [None,""] else None
    except ValueError as e: return f"❌ {e}"
    ccf=calcular_completiva(cf,cecv)
    cexf=calcular_extraordinaria(cf,ceexv)
    # Jerarquía para la nota final mostrada: especial C.E > especial C.F > extraordinaria > completiva > C.F.
    final=ece if ece is not None else (ecf if ecf is not None else (cexf if cexf is not None else (ccf if ccf is not None else cf)))
    umbral=float(get_config("umbral_aprobacion","70"))
    sit=(situacion or ("A" if final is not None and final>=umbral else ("R" if final is not None else ""))).upper()
    execute("""INSERT INTO calificaciones_finales(estudiante_id,asignatura_id,completiva,extraordinaria,especial_cf,especial_ce,situacion,observacion,fecha_actualizacion)
        VALUES (?,?,?,?,?,?,?,?,?)
        ON CONFLICT(estudiante_id,asignatura_id) DO UPDATE SET completiva=excluded.completiva,extraordinaria=excluded.extraordinaria,especial_cf=excluded.especial_cf,especial_ce=excluded.especial_ce,situacion=excluded.situacion,observacion=excluded.observacion,fecha_actualizacion=excluded.fecha_actualizacion""",
        (sid,aid,cecv,ceexv,ecf,ece,sit,observacion,now_str()))
    return f"✅ Guardado. C.F.={cf if cf is not None else '—'} | C.C.F.={ccf if ccf is not None else '—'} | C.EX.F.={cexf if cexf is not None else '—'} | Situación={sit or '—'}"


def calculo_final_completo(asignatura, course_id=None):
    aid=subject_id_from_choice(asignatura)
    if not aid: return pd.DataFrame()
    cid=course_id_from_choice(course_id)
    if cid:
        students=db_df("""SELECT e.id,e.matricula,e.nombre,e.curso,e.seccion FROM estudiantes e JOIN estudiante_curso ec ON ec.estudiante_id=e.id WHERE e.estado='Activo' AND ec.curso_id=? AND ec.fecha_fin IS NULL ORDER BY e.nombre COLLATE NOCASE,e.id""",(cid,))
    else:
        students=db_df("SELECT id,matricula,nombre,curso,seccion FROM estudiantes WHERE estado='Activo' ORDER BY nombre COLLATE NOCASE,id")
    rows=[]; umbral=float(get_config("umbral_aprobacion","70"))
    for _,s in students.iterrows():
        sid=int(s["id"]); cf=promedio_final(sid,aid)
        r=db_df("SELECT * FROM calificaciones_finales WHERE estudiante_id=? AND asignatura_id=?",(sid,aid))
        cec=ceex=ecf=ece=sit=obs=None
        if not r.empty:
            rr=r.iloc[0]; cec=rr["completiva"]; ceex=rr["extraordinaria"]; ecf=rr["especial_cf"]; ece=rr["especial_ce"]; sit=rr["situacion"]; obs=rr["observacion"]
        ccf=calcular_completiva(cf,cec); cexf=calcular_extraordinaria(cf,ceex)
        final=ece if ece is not None else (ecf if ecf is not None else (cexf if ceex is not None else (ccf if cec is not None else cf)))
        rows.append({"Matrícula":s["matricula"],"Estudiante":s["nombre"],"CF":cf,
                     "C.E.C":cec,"C.C.F":ccf,"C.E.EX":ceex,"C.EX.F":cexf,
                     "Especial C.F":ecf,"Especial C.E":ece,
                     "Calificación final":final,
                     "Situación final":sit or ("A" if final is not None and final>=umbral else ("R" if final is not None else "")),
                     "Observación":obs or ""})
    return pd.DataFrame(rows)


def boletin_final_estudiante(estudiante):
    sid=student_id_from_choice(estudiante)
    if not sid: return pd.DataFrame()
    subs=db_df("SELECT id,nombre,codigo,grado FROM asignaturas WHERE activa=1 ORDER BY grado,nombre")
    rows=[]; umbral=float(get_config("umbral_aprobacion","70"))
    for _,a in subs.iterrows():
        aid=int(a["id"]); cf=promedio_final(sid,aid)
        r=db_df("SELECT * FROM calificaciones_finales WHERE estudiante_id=? AND asignatura_id=?",(sid,aid))
        cec=ceex=ecf=ece=sit=None
        if not r.empty:
            rr=r.iloc[0]; cec=rr["completiva"]; ceex=rr["extraordinaria"]; ecf=rr["especial_cf"]; ece=rr["especial_ce"]; sit=rr["situacion"]
        ccf=calcular_completiva(cf,cec); cexf=calcular_extraordinaria(cf,ceex)
        final=ece if ece is not None else (ecf if ecf is not None else (cexf if ceex is not None else (ccf if cec is not None else cf)))
        rows.append({"Asignatura":a["nombre"],"Grado":a["grado"],"C.F.":cf,"C.E.C":cec,"C.C.F.":ccf,"C.E.EX":ceex,"C.EX.F.":cexf,"Especial C.F":ecf,"Especial C.E":ece,"Calificación final":final,"Situación":sit or ("A" if final is not None and final>=umbral else ("R" if final is not None else ""))})
    return pd.DataFrame(rows)


def exportar_boletin_pdf(estudiante):
    sid=student_id_from_choice(estudiante)
    if not sid: raise ValueError("Selecciona un estudiante.")
    stu=db_df("SELECT * FROM estudiantes WHERE id=?",(sid,))
    if stu.empty: raise ValueError("Estudiante no encontrado.")
    r=stu.iloc[0]
    data=boletin_final_estudiante(estudiante)
    ts=datetime.now().strftime("%Y%m%d_%H%M%S")
    path=REPORTES_DIR/f"Boletin_Final_{r['matricula']}_{ts}.pdf"
    doc=SimpleDocTemplate(str(path),pagesize=landscape(letter),rightMargin=24,leftMargin=24,topMargin=24,bottomMargin=24)
    styles=getSampleStyleSheet(); elements=[]
    centro=get_config("nombre_centro","Centro Educativo"); anio=get_config("anio_escolar","2026-2027")
    elements.append(Paragraph(f"<b>{centro}</b>",styles["Title"]))
    elements.append(Paragraph(f"<b>BOLETÍN FINAL — AÑO ESCOLAR {anio}</b>",styles["Heading2"]))
    elements.append(Paragraph(f"Estudiante: <b>{r['nombre']}</b> &nbsp;&nbsp; Matrícula: <b>{r['matricula']}</b> &nbsp;&nbsp; Curso: <b>{r['curso']}</b> &nbsp;&nbsp; Sección: <b>{r['seccion']}</b>",styles["BodyText"]))
    elements.append(Spacer(1,10))
    headers=["Asignatura","C.F.","C.E.C","C.C.F.","C.E.EX","C.EX.F.","Esp. C.F","Esp. C.E","Final","Situación"]
    rows=[headers]
    for _,x in data.iterrows():
        rows.append([str(x["Asignatura"]),*[("" if pd.isna(x[k]) else f"{float(x[k]):.2f}") for k in ["C.F.","C.E.C","C.C.F.","C.E.EX","C.EX.F.","Especial C.F","Especial C.E","Calificación final"]],str(x["Situación"] or "")])
    table=Table(rows,repeatRows=1,colWidths=[145,42,42,45,45,48,48,48,48,55])
    table.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.lightgrey),("GRID",(0,0),(-1,-1),0.5,colors.grey),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),8),("ALIGN",(1,1),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE")]))
    elements.append(table); elements.append(Spacer(1,10))
    elements.append(Paragraph("C.F. = Calificación Final del año. C.C.F. = resultado de la Completiva. C.EX.F. = resultado de la Extraordinaria. Las ponderaciones y el umbral de aprobación son configurables en Administración.",styles["BodyText"]))
    doc.build(elements)
    return str(path)


# ============================================================
# 13. PERFIL DEL ESTUDIANTE
# ============================================================

def perfil_estudiante(estudiante):
    sid=student_id_from_choice(estudiante)
    if not sid:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), "Selecciona un estudiante."

    datos=db_df("SELECT * FROM estudiantes WHERE id=?",(sid,))
    asistencia=db_df("""
        SELECT fecha,estado,hora_entrada,observacion
        FROM asistencia WHERE estudiante_id=? ORDER BY fecha DESC
    """,(sid,))
    incidencias=db_df("""
        SELECT fecha,hora,tipo,gravedad,descripcion,accion_tomada,observacion
        FROM incidencias WHERE estudiante_id=? ORDER BY fecha DESC
    """,(sid,))
    resumen = datos.to_dict("records")[0] if not datos.empty else {}
    texto = (
        f"### {resumen.get('nombre','')}\n"
        f"- Matrícula: {resumen.get('matricula','')}\n"
        f"- Curso: {resumen.get('curso','')}\n"
        f"- Sección: {resumen.get('seccion','')}\n"
        f"- Fecha de nacimiento: {resumen.get('fecha_nacimiento','')}\n"
        f"- Edad al registrar: {resumen.get('edad_registro','')} años\n"
        f"- Estado: {resumen.get('estado','')}"
    )
    return asistencia, incidencias, pd.DataFrame([resumen]), texto


# ============================================================
# 14. CALENDARIO
# ============================================================

def guardar_evento(fecha,hora_inicio,hora_fin,titulo,tipo,curso,seccion,descripcion):
    if not fecha or not titulo:
        return "Completa fecha y título."
    execute(
        """INSERT INTO calendario
           (fecha,hora_inicio,hora_fin,titulo,tipo,curso,seccion,descripcion)
           VALUES (?,?,?,?,?,?,?,?)""",
        (fecha,hora_inicio,hora_fin,titulo,tipo,curso,seccion,descripcion)
    )
    return "✅ Evento agregado."

def listar_calendario():
    return db_df("""
        SELECT id,fecha,hora_inicio,hora_fin,titulo,tipo,curso,seccion,descripcion
        FROM calendario ORDER BY fecha,hora_inicio
    """)


# ============================================================
# 15. ESTADÍSTICAS
# ============================================================

def grafico_asistencia(course_id=None):
    cid=course_id_from_choice(course_id); params=[]; extra=""
    if cid:
        extra=" JOIN estudiante_curso ec ON ec.estudiante_id=a.estudiante_id AND ec.curso_id=? AND ec.fecha_fin IS NULL"; params=[cid]
    df=db_df(f"""
        SELECT a.estado,COUNT(*) cantidad
        FROM asistencia a{extra} GROUP BY a.estado ORDER BY cantidad DESC
    """,params)
    fig=plt.figure(figsize=(8,4.5))
    if df.empty:
        plt.text(0.5,0.5,"Sin datos de asistencia",ha="center",va="center")
        plt.axis("off")
        return fig
    plt.bar(df["estado"],df["cantidad"])
    plt.title("Distribución de asistencia")
    plt.ylabel("Registros")
    plt.xticks(rotation=20)
    plt.tight_layout()
    return fig

def grafico_incidencias(course_id=None):
    cid=course_id_from_choice(course_id); params=[]; extra=""
    if cid:
        extra=" JOIN estudiante_curso ec ON ec.estudiante_id=i.estudiante_id AND ec.curso_id=? AND ec.fecha_fin IS NULL"; params=[cid]
    df=db_df(f"""
        SELECT i.gravedad,COUNT(*) cantidad
        FROM incidencias i{extra} GROUP BY i.gravedad ORDER BY cantidad DESC
    """,params)
    fig=plt.figure(figsize=(8,4.5))
    if df.empty:
        plt.text(0.5,0.5,"Sin incidencias registradas",ha="center",va="center")
        plt.axis("off")
        return fig
    plt.bar(df["gravedad"],df["cantidad"])
    plt.title("Incidencias por gravedad")
    plt.ylabel("Cantidad")
    plt.tight_layout()
    return fig

def resumen_calificaciones(asignatura, course_id=None):
    aid=subject_id_from_choice(asignatura)
    if not aid:
        return pd.DataFrame()
    cid=course_id_from_choice(course_id)
    if cid:
        students=db_df("""SELECT e.id,e.matricula,e.nombre,e.curso,e.seccion FROM estudiantes e JOIN estudiante_curso ec ON ec.estudiante_id=e.id WHERE ec.curso_id=? AND ec.fecha_fin IS NULL ORDER BY e.nombre COLLATE NOCASE,e.id""",(cid,))
    else:
        students=db_df("SELECT id,matricula,nombre,curso,seccion FROM estudiantes ORDER BY nombre COLLATE NOCASE,id")
    rows=[]
    for _,s in students.iterrows():
        final=promedio_final(int(s["id"]),aid)
        rows.append({
            "Matrícula":s["matricula"],
            "Estudiante":s["nombre"],
            "Curso":s["curso"],
            "Sección":s["seccion"],
            "Calificación final":final
        })
    return pd.DataFrame(rows)


# ============================================================
# 16. ALERTAS
# ============================================================

def alertas(course_id=None):
    asist=porcentaje_asistencia(course_id)
    alert_asist=asist[asist["porcentaje"].fillna(0)<80] if not asist.empty else pd.DataFrame()

    cid=course_id_from_choice(course_id); params=[]; extra=""
    if cid:
        extra=" JOIN estudiante_curso ec ON ec.estudiante_id=e.id AND ec.curso_id=? AND ec.fecha_fin IS NULL"; params=[cid]
    cal=db_df(f"""
        SELECT e.matricula,e.nombre,a.nombre AS asignatura,
               AVG(c.calificacion) AS promedio
        FROM calificaciones c
        JOIN estudiantes e ON e.id=c.estudiante_id{extra}
        JOIN asignaturas a ON a.id=c.asignatura_id
        GROUP BY e.id,a.id
        HAVING AVG(c.calificacion) < 70
        ORDER BY promedio
    """,params)
    return alert_asist, cal


# ============================================================
# 17. EXPORTACIÓN EXCEL
# ============================================================

def exportar_excel(asignatura, course_id=None):
    aid=subject_id_from_choice(asignatura)
    ts=datetime.now().strftime("%Y%m%d_%H%M%S")
    path=REPORTES_DIR/f"Reporte_Gestion_Clase_{ts}.xlsx"

    with pd.ExcelWriter(path,engine="openpyxl") as writer:
        db_df("SELECT * FROM estudiantes ORDER BY nombre COLLATE NOCASE").to_excel(writer,index=False,sheet_name="Estudiantes")
        db_df("""
            SELECT a.*,e.matricula,e.nombre
            FROM asistencia a JOIN estudiantes e ON e.id=a.estudiante_id
            ORDER BY a.fecha DESC,e.nombre
        """).to_excel(writer,index=False,sheet_name="Asistencia")
        porcentaje_asistencia(course_id).to_excel(writer,index=False,sheet_name="Resumen asistencia")
        listar_incidencias(course_id).to_excel(writer,index=False,sheet_name="Incidencias")
        listar_permisos(course_id).to_excel(writer,index=False,sheet_name="Permisos")
        listar_docentes().to_excel(writer,index=False,sheet_name="Docentes")
        listar_cursos().to_excel(writer,index=False,sheet_name="Cursos")
        listar_asignaturas().to_excel(writer,index=False,sheet_name="Asignaturas")
        db_df("""
            SELECT a.nombre AS asignatura,c.codigo,c.nombre AS competencia,
                   c.descripcion
            FROM competencias c JOIN asignaturas a ON a.id=c.asignatura_id
            ORDER BY a.nombre,c.orden
        """).to_excel(writer,index=False,sheet_name="Competencias")
        if aid:
            registro_grado(asignatura,course_id).to_excel(writer,index=False,sheet_name="Registro calificaciones")
            resumen_calificaciones(asignatura,course_id).to_excel(writer,index=False,sheet_name="Resumen notas")
            listar_actividades(asignatura,course_id=course_id).to_excel(writer,index=False,sheet_name="Actividades")
            resumen_actividades_por_competencia(asignatura,course_id=course_id).to_excel(writer,index=False,sheet_name="Promedio actividades")
            calculo_final_completo(asignatura,course_id).to_excel(writer,index=False,sheet_name="Calificaciones finales")
        todas_fundamentales().to_excel(writer,index=False,sheet_name="Competencias fundamentales")
        db_df("""
            SELECT f.codigo AS CF_codigo,f.nombre AS CF_nombre,c.codigo AS CE_codigo,c.nombre AS CE_nombre
            FROM competencia_fundamental_especifica m
            JOIN competencias_fundamentales f ON f.id=m.fundamental_id
            JOIN competencias c ON c.id=m.especifica_id
            ORDER BY f.orden,c.orden
        """).to_excel(writer,index=False,sheet_name="Mapa CF-CE")
        listar_calendario().to_excel(writer,index=False,sheet_name="Calendario")
    return str(path)

def backup_db():
    ts=datetime.now().strftime("%Y%m%d_%H%M%S")
    path=BACKUP_DIR/f"gestion_clase_v5_{ts}.db"
    shutil.copy2(DB_PATH,path)
    return str(path)


# ============================================================
# 18. PDF
# ============================================================

def exportar_pdf_registro(asignatura, course_id=None):
    aid=subject_id_from_choice(asignatura)
    if not aid:
        raise ValueError("Selecciona una asignatura.")

    ts=datetime.now().strftime("%Y%m%d_%H%M%S")
    path=REPORTES_DIR/f"Registro_Grado_{ts}.pdf"

    data=registro_grado(asignatura,course_id)
    comps=get_competencias_ids(asignatura)

    doc=SimpleDocTemplate(
        str(path), pagesize=landscape(letter),
        rightMargin=20,leftMargin=20,topMargin=20,bottomMargin=20
    )
    styles=getSampleStyleSheet()
    elements=[]

    subject=db_df("SELECT * FROM asignaturas WHERE id=?",(aid,))
    subject_name=subject.iloc[0]["nombre"] if not subject.empty else "Asignatura"

    title=Paragraph(
        f"<b>REGISTRO DE CALIFICACIONES — {subject_name}</b>",
        styles["Title"]
    )
    elements.append(title)
    elements.append(Spacer(1,8))

    # Tabla compacta: estudiante + promedio de cada competencia + final
    headers=["Matrícula","Estudiante"]
    for _,c in comps.iterrows():
        headers.append(f"{c['codigo']}<br/>Prom.")
    headers.append("Calificación<br/>final")

    rows=[headers]
    for _,r in data.iterrows():
        vals=[str(r["Matrícula"]),str(r["Estudiante"])]
        for _,c in comps.iterrows():
            v=r.get(f"{c['codigo']}_Promedio","")
            vals.append("" if pd.isna(v) else str(v))
        v=r.get("Calificación final","")
        vals.append("" if pd.isna(v) else str(v))
        rows.append(vals)

    table=Table(rows,repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.lightgrey),
        ("TEXTCOLOR",(0,0),(-1,0),colors.black),
        ("GRID",(0,0),(-1,-1),0.4,colors.grey),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),7),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("ALIGN",(2,1),(-1,-1),"CENTER"),
    ]))
    elements.append(table)
    elements.append(Spacer(1,10))

    elements.append(Paragraph(
        "<b>Nota:</b> El PDF resume promedios y calificación final. "
        "La matriz editable completa queda disponible en el módulo de calificaciones y en Excel.",
        styles["BodyText"]
    ))

    doc.build(elements)
    return str(path)


# ============================================================
# 19. CONFIGURACIÓN
# ============================================================

def cargar_config():
    return (
        get_config("anio_escolar","2026-2027"),
        get_config("nombre_centro","Centro Educativo"),
        float(get_config("max_calificacion","100")),
        get_config("rp_reemplaza_p","1")=="1",
        float(get_config("umbral_aprobacion","70")),
        float(get_config("peso_completiva_cf","50")),
        float(get_config("peso_completiva_cec","50")),
        float(get_config("peso_extra_cf","30")),
        float(get_config("peso_extra_ceex","70")),
    )

def guardar_config(anio, centro, max_nota, rp_reemplaza, umbral, pccf, pccec, pecf, pecex):
    set_config("anio_escolar",anio)
    set_config("nombre_centro",centro)
    set_config("max_calificacion",max_nota)
    set_config("rp_reemplaza_p","1" if rp_reemplaza else "0")
    set_config("umbral_aprobacion",umbral)
    set_config("peso_completiva_cf",pccf)
    set_config("peso_completiva_cec",pccec)
    set_config("peso_extra_cf",pecf)
    set_config("peso_extra_ceex",pecex)
    return "✅ Configuración guardada."


# ============================================================
# 20. INTERFAZ GRADIO
# ============================================================

CSS = """
.gradio-container {max-width: 1500px !important;}
h1,h2,h3 {margin-top: 8px;}
.small-note {font-size: 0.9em; opacity: .85;}
"""

with gr.Blocks(title="Sistema de Gestión de Clase V5.1", css=CSS, theme=gr.themes.Soft()) as app:

    # ---------------- LOGIN ----------------
    login_box=gr.Column(visible=True)
    with login_box:
        gr.Markdown("# 📚 Sistema de Gestión de Clase — V5")
        gr.Markdown(
            "### Control integral de estudiantes, asistencia, incidencias, permisos y calificaciones."
        )
        with gr.Row():
            login_user=gr.Textbox(label="Usuario", value="admin")
            login_pass=gr.Textbox(label="Contraseña", type="password", value="admin123")
        login_btn=gr.Button("🔐 Iniciar sesión", variant="primary")
        login_msg=gr.Markdown()

    # ---------------- APP ----------------
    main_box=gr.Column(visible=False)
    with main_box:
        with gr.Row():
            gr.Markdown("## 📚 Sistema de Gestión de Clase V5.1")
            logout_btn=gr.Button("Cerrar sesión", scale=0)

        with gr.Row():
            clase_activa=gr.Dropdown(choices=course_choices(),label="🏫 Clase activa",scale=4)
            clase_refrescar=gr.Button("🔄 Actualizar clases",scale=1)
        clase_info=gr.Markdown("Selecciona una clase para trabajar con su grupo de estudiantes.")
        
        with gr.Tabs():
        
            # GESTIÓN DE CLASES
            with gr.Tab("🏫 Clases"):
                gr.Markdown("""
                ## Gestión de clases
                Desde aquí puedes **crear, modificar, activar/desactivar y eliminar definitivamente** las clases.
                La clase se selecciona desde una lista; no es necesario escribir el ID manualmente.
                **Desactivar** conserva el historial. **Eliminar definitivamente** elimina la clase y sus relaciones de matrícula/asignaturas, pero conserva los estudiantes y sus registros.
                """)
                with gr.Row():
                    gc_selector=gr.Dropdown(choices=course_choices(), label="Clase seleccionada", scale=3)
                    gc_refresh=gr.Button("🔄 Actualizar lista", scale=1)
                with gr.Row():
                    gc_nombre=gr.Textbox(label="Nombre de la clase")
                    gc_nivel=gr.Textbox(label="Nivel", value="Secundaria")
                with gr.Row():
                    gc_grado=gr.Textbox(label="Grado")
                    gc_seccion=gr.Textbox(label="Sección")
                    gc_anio=gr.Textbox(label="Año escolar", value=get_config("anio_escolar","2026-2027"))
                with gr.Row():
                    gc_doc=gr.Number(label="ID docente (opcional)", precision=0)
                    gc_estado=gr.Dropdown(["Activo","Inactivo"], value="Activo", label="Estado")
                with gr.Row():
                    gc_new=gr.Button("➕ Crear nueva clase", variant="primary")
                    gc_load=gr.Button("📥 Cargar clase")
                    gc_update=gr.Button("✏️ Guardar cambios", variant="primary")
                    gc_deactivate=gr.Button("🚫 Desactivar")
                gc_confirm=gr.Checkbox(label="Confirmo que quiero eliminar definitivamente la clase seleccionada", value=False)
                with gr.Row():
                    gc_delete=gr.Button("🗑️ Eliminar definitivamente", variant="stop")
                gc_msg=gr.Markdown()
                gc_table=gr.Dataframe(label="Clases registradas", interactive=False)

                def refrescar_clases_ui():
                    choices=course_choices()
                    return gr.update(choices=choices, value=None), listar_cursos()

                gc_selector.change(
                    cargar_curso, [gc_selector],
                    [gc_nombre,gc_nivel,gc_grado,gc_seccion,gc_anio,gc_doc,gc_estado,gc_msg]
                )
                gc_load.click(
                    cargar_curso, [gc_selector],
                    [gc_nombre,gc_nivel,gc_grado,gc_seccion,gc_anio,gc_doc,gc_estado,gc_msg]
                )
                gc_new.click(
                    crear_curso, [gc_nombre,gc_nivel,gc_grado,gc_seccion,gc_anio,gc_doc], gc_msg
                ).then(refrescar_clases_ui, None, [gc_selector,gc_table])
                gc_update.click(
                    actualizar_curso, [gc_selector,gc_nombre,gc_nivel,gc_grado,gc_seccion,gc_anio,gc_doc,gc_estado], gc_msg
                ).then(refrescar_clases_ui, None, [gc_selector,gc_table])
                gc_deactivate.click(
                    eliminar_curso, [gc_selector], gc_msg
                ).then(refrescar_clases_ui, None, [gc_selector,gc_table])
                gc_delete.click(
                    eliminar_curso_definitivo, [gc_selector,gc_confirm], gc_msg
                ).then(refrescar_clases_ui, None, [gc_selector,gc_table])
                gc_refresh.click(refrescar_clases_ui, None, [gc_selector,gc_table])

            # DASHBOARD
            with gr.Tab("🏠 Dashboard"):
                gr.Markdown("## Resumen general")
                with gr.Row():
                    total_students=gr.Number(label="Estudiantes", interactive=False)
                    total_inc=gr.Number(label="Incidencias", interactive=False)
                    total_perm=gr.Number(label="Permisos", interactive=False)
                    total_att=gr.Number(label="Registros asistencia", interactive=False)
                dash_table=gr.Dataframe(label="Estudiantes con baja asistencia")
                dash_refresh=gr.Button("Actualizar dashboard")

            # ESTUDIANTES
            with gr.Tab("👨‍🎓 Estudiantes"):
                gr.Markdown("""
                ### Gestión de estudiantes por clase
                La matrícula se genera automáticamente como **primera letra de los nombres + primera letra de los apellidos + año de nacimiento + número de orden de inscripción en la clase**.
                El listado siempre se presenta alfabéticamente. La edad se calcula automáticamente a partir de la fecha de nacimiento.
                """)
                with gr.Row():
                    with gr.Column():
                        s_curso_id=gr.Dropdown(choices=course_choices(),label="Clase",value=None)
                        s_nombres=gr.Textbox(label="Nombres")
                        s_apellidos=gr.Textbox(label="Apellidos")
                        s_fecha_nac=gr.Textbox(label="Fecha de nacimiento (AAAA-MM-DD)")
                        s_edad=gr.Number(label="Edad al registrar",interactive=False)
                        s_matricula=gr.Textbox(label="Matrícula automática",interactive=False)
                        s_correo=gr.Textbox(label="Correo")
                        s_telefono=gr.Textbox(label="Teléfono")
                        s_estado=gr.Dropdown(["Activo","Inactivo"],value="Activo",label="Estado")
                        with gr.Row():
                            s_guardar=gr.Button("➕ Registrar",variant="primary")
                            s_actualizar=gr.Button("✏️ Actualizar")
                            s_eliminar=gr.Button("🗑️ Eliminar")
                        s_msg=gr.Markdown()
                    with gr.Column():
                        gr.Markdown("### Estudiantes de la clase seleccionada")
                        s_buscar=gr.Textbox(label="Buscar por matrícula, nombre, curso o sección")
                        s_tabla=gr.Dataframe(interactive=False)
                        with gr.Row():
                            s_refrescar=gr.Button("🔄 Buscar / actualizar")
                            s_cargar=gr.Button("Cargar por matrícula")
                s_nombres.change(lambda n,a,f,c: generar_matricula(n,a,f,c),[s_nombres,s_apellidos,s_fecha_nac,s_curso_id],s_matricula)
                s_fecha_nac.change(lambda n,a,f,c: (calcular_edad(f) or "",generar_matricula(n,a,f,c)),[s_nombres,s_apellidos,s_fecha_nac,s_curso_id],[s_edad,s_matricula])
                s_curso_id.change(lambda n,a,f,c: (generar_matricula(n,a,f,c),listar_estudiantes_clase(c)),[s_nombres,s_apellidos,s_fecha_nac,s_curso_id],[s_matricula,s_tabla])
                s_guardar.click(guardar_estudiante,[s_nombres,s_apellidos,s_fecha_nac,s_curso_id,s_correo,s_telefono,s_estado],[s_msg,s_matricula]).then(buscar_estudiantes,[s_buscar,s_curso_id],s_tabla)
                s_actualizar.click(actualizar_estudiante,[s_matricula,s_nombres,s_apellidos,s_curso_id,s_fecha_nac,s_correo,s_telefono,s_estado],s_msg).then(buscar_estudiantes,[s_buscar,s_curso_id],s_tabla)
                s_eliminar.click(eliminar_estudiante,[s_matricula],s_msg).then(buscar_estudiantes,[s_buscar,s_curso_id],s_tabla)
                s_refrescar.click(buscar_estudiantes,[s_buscar,s_curso_id],s_tabla)
                s_cargar.click(cargar_estudiante,[s_matricula],[s_matricula,s_nombres,s_apellidos,s_curso_id,s_fecha_nac,s_correo,s_telefono,s_estado,s_edad,s_msg])

            # ASISTENCIA
            with gr.Tab("🕘 Asistencia"):
                with gr.Row():
                    att_fecha=gr.Textbox(label="Fecha",value=today_str())
                    att_curso=gr.Dropdown(choices=course_choices(),label="Clase")
                    att_cargar=gr.Button("Cargar estudiantes",variant="primary")
                att_tabla=gr.Dataframe(
                    headers=["id","matricula","nombre","curso","seccion","estado","hora_entrada","observacion"],
                    datatype=["number","str","str","str","str","str","str","str"],
                    interactive=True,
                    label="Registro de asistencia"
                )
                with gr.Row():
                    att_pres=gr.Button("Todos PRESENTE")
                    att_aus=gr.Button("Todos AUSENTE")
                    att_tar=gr.Button("Todos TARDANZA")
                    att_exc=gr.Button("Todos EXCUSA/PERMISO")
                    att_guardar=gr.Button("💾 Guardar asistencia",variant="primary")
                att_msg=gr.Markdown()

                att_cargar.click(asistencia_del_dia,[att_fecha,att_curso],att_tabla)
                att_pres.click(lambda f,c: marcar_todos(f,"Presente",c),[att_fecha,att_curso],[att_tabla,att_msg])
                att_aus.click(lambda f,c: marcar_todos(f,"Ausente",c),[att_fecha,att_curso],[att_tabla,att_msg])
                att_tar.click(lambda f,c: marcar_todos(f,"Tardanza",c),[att_fecha,att_curso],[att_tabla,att_msg])
                att_exc.click(lambda f,c: marcar_todos(f,"Excusa/Permiso",c),[att_fecha,att_curso],[att_tabla,att_msg])
                att_guardar.click(guardar_asistencia,[att_fecha,att_tabla],att_msg)

            # INCIDENCIAS
            with gr.Tab("⚠️ Incidencias"):
                with gr.Row():
                    with gr.Column():
                        inc_est=gr.Dropdown(choices=student_choices(),label="Estudiante")
                        inc_fecha=gr.Textbox(label="Fecha",value=today_str())
                        inc_hora=gr.Textbox(label="Hora",value=datetime.now().strftime("%H:%M"))
                        inc_tipo=gr.Dropdown(
                            ["Académica","Conductual","Asistencia","Convivencia","Disciplina","Otra"],
                            label="Tipo"
                        )
                        inc_grav=gr.Dropdown(["Leve","Moderada","Grave"],value="Leve",label="Gravedad")
                        inc_desc=gr.Textbox(label="Descripción",lines=4)
                        inc_acc=gr.Textbox(label="Acción tomada",lines=3)
                        inc_obs=gr.Textbox(label="Observación",lines=3)
                        inc_btn=gr.Button("Registrar incidencia",variant="primary")
                        inc_msg=gr.Markdown()
                    with gr.Column():
                        inc_tabla=gr.Dataframe(label="Historial de incidencias",interactive=False)
                        inc_refresh=gr.Button("Actualizar")
                inc_btn.click(
                    guardar_incidencia,
                    [inc_est,inc_fecha,inc_hora,inc_tipo,inc_grav,inc_desc,inc_acc,inc_obs],
                    inc_msg
                ).then(listar_incidencias,[clase_activa],inc_tabla)
                inc_refresh.click(listar_incidencias,[clase_activa],inc_tabla)

            # PERMISOS
            with gr.Tab("📝 Permisos"):
                with gr.Row():
                    with gr.Column():
                        per_est=gr.Dropdown(choices=student_choices(),label="Estudiante")
                        per_fecha=gr.Textbox(label="Fecha",value=today_str())
                        per_tipo=gr.Dropdown(["Salida","Entrada tardía","Ausencia","Médico","Familiar","Otro"],label="Tipo")
                        per_desde=gr.Textbox(label="Hora desde")
                        per_hasta=gr.Textbox(label="Hora hasta")
                        per_motivo=gr.Textbox(label="Motivo",lines=3)
                        per_estado=gr.Dropdown(["Pendiente","Aprobado","Rechazado"],value="Pendiente",label="Estado")
                        per_obs=gr.Textbox(label="Observación",lines=2)
                        per_btn=gr.Button("Registrar permiso",variant="primary")
                        per_msg=gr.Markdown()
                    with gr.Column():
                        per_tabla=gr.Dataframe(interactive=False)
                        per_refresh=gr.Button("Actualizar")
                per_btn.click(
                    guardar_permiso,
                    [per_est,per_fecha,per_tipo,per_desde,per_hasta,per_motivo,per_estado,per_obs],
                    per_msg
                ).then(listar_permisos,[clase_activa],per_tabla)
                per_refresh.click(listar_permisos,[clase_activa],per_tabla)

            # CALIFICACIONES
            with gr.Tab("🎓 Calificaciones / Registro de Grado"):
                gr.Markdown("""
                ## Registro de calificaciones de Secundaria
                La estructura está preparada para trabajar por **competencias específicas**,
                **4 períodos (P1-P4)** y **recuperación (RP1-RP4)**, con promedio por
                competencia y calificación final.

                El ejemplo suministrado corresponde a **Matemática, 4to. grado** y presenta
                competencias CE-MAT1 a CE-MAT7, mientras que el registro agrupa las columnas
                en PC1, PC2, PC3 y PC4. El sistema permite editar tanto los códigos como las
                descripciones y la cantidad de competencias.
                """)
                with gr.Row():
                    cal_asig=gr.Dropdown(choices=subject_choices(),label="Asignatura",scale=2)
                    cal_est=gr.Dropdown(choices=student_choices(),label="Estudiante",scale=2)
                cal_comp=gr.Dropdown(label="Competencia específica")
                with gr.Row():
                    cal_periodo=gr.Dropdown([1,2,3,4],value=1,label="Período")
                    cal_tipo=gr.Dropdown(["P","RP"],value="P",label="Tipo")
                    cal_nota=gr.Number(label="Calificación",minimum=0,maximum=100)
                cal_obs=gr.Textbox(label="Observación")
                with gr.Row():
                    cal_cargar=gr.Button("Cargar nota")
                    cal_guardar=gr.Button("💾 Guardar nota",variant="primary")
                cal_msg=gr.Markdown()
                cal_detalle=gr.Dataframe(label="Registro del estudiante",interactive=False)

                gr.Markdown("### Registro general por asignatura")
                reg_asig=gr.Dropdown(choices=subject_choices(),label="Asignatura para registro")
                reg_tabla=gr.Dataframe(label="Matriz del registro de grado",interactive=True)
                with gr.Row():
                    reg_cargar=gr.Button("📋 Cargar registro")
                    reg_guardar=gr.Button("💾 Guardar cambios de la matriz",variant="primary")
                    reg_excel=gr.Button("📊 Exportar Excel")
                    reg_pdf=gr.Button("📄 Exportar PDF")
                reg_msg=gr.Markdown()
                reg_file=gr.File(label="Archivo generado")

                cal_asig.change(opciones_competencias,[cal_asig],cal_comp)
                cal_guardar.click(
                    guardar_calificacion,
                    [cal_est,cal_asig,cal_comp,cal_periodo,cal_tipo,cal_nota,cal_obs],
                    cal_msg
                ).then(notas_estudiante,[cal_est,cal_asig],cal_detalle)

                cal_cargar.click(
                    cargar_nota_actual,
                    [cal_est,cal_asig,cal_comp,cal_periodo,cal_tipo],
                    [cal_nota,cal_obs]
                )
                cal_est.change(notas_estudiante,[cal_est,cal_asig],cal_detalle)
                cal_asig.change(notas_estudiante,[cal_est,cal_asig],cal_detalle)

                reg_cargar.click(registro_grado,[reg_asig,clase_activa],reg_tabla)
                reg_guardar.click(
                    guardar_registro_grado,[reg_asig,reg_tabla],reg_msg
                ).then(registro_grado,[reg_asig],reg_tabla)

                reg_excel.click(exportar_excel,[reg_asig,clase_activa],reg_file)
                reg_pdf.click(exportar_pdf_registro,[reg_asig,clase_activa],reg_file)

            # COMPETENCIAS
            with gr.Tab("🧠 Competencias"):
                gr.Markdown("""
                ## Gestor de competencias fundamentales y específicas
                Una competencia específica (por ejemplo **CE-MAT1**) se almacena una sola vez.
                Puedes asociarla a cualquier competencia fundamental mediante el mapa CF ↔ CE,
                sin volver a copiar su descripción.
                """)
                with gr.Tab("Fundamentales"):
                    cf_tabla=gr.Dataframe(interactive=False,label="Competencias fundamentales")
                    with gr.Row():
                        with gr.Column():
                            cf_codigo_n=gr.Textbox(label="Código")
                            cf_nombre_n=gr.Textbox(label="Nombre")
                            cf_desc_n=gr.Textbox(label="Descripción",lines=5)
                            cf_orden_n=gr.Number(label="Orden",value=1,precision=0)
                            cf_crear=gr.Button("➕ Añadir fundamental",variant="primary")
                        with gr.Column():
                            cf_id=gr.Number(label="ID",precision=0)
                            cf_codigo=gr.Textbox(label="Código")
                            cf_nombre=gr.Textbox(label="Nombre")
                            cf_desc=gr.Textbox(label="Descripción",lines=5)
                            cf_orden=gr.Number(label="Orden",precision=0)
                            cf_activa=gr.Checkbox(label="Activa",value=True)
                            with gr.Row():
                                cf_cargar=gr.Button("Cargar")
                                cf_actualizar=gr.Button("💾 Editar",variant="primary")
                                cf_eliminar=gr.Button("🗑️ Eliminar")
                    cf_msg=gr.Markdown()
                    cf_ref=gr.Button("🔄 Actualizar")

                with gr.Tab("Específicas"):
                    comp_asig=gr.Dropdown(choices=subject_choices(),label="Asignatura")
                    comp_tabla=gr.Dataframe(interactive=False,label="Competencias específicas y sus fundamentales")
                    comp_refresh=gr.Button("🔄 Actualizar competencias")
                    with gr.Row():
                        with gr.Column():
                            gr.Markdown("#### Añadir / editar / codificar")
                            comp_id=gr.Number(label="ID",precision=0)
                            comp_codigo_n=gr.Textbox(label="Código (ej. CE-MAT1)")
                            comp_nombre_n=gr.Textbox(label="Nombre")
                            comp_desc_n=gr.Textbox(label="Descripción",lines=6)
                            comp_orden_n=gr.Number(label="Orden",value=1,precision=0)
                            comp_activa=gr.Checkbox(label="Activa",value=True)
                            with gr.Row():
                                comp_crear=gr.Button("➕ Añadir",variant="primary")
                                comp_cargar=gr.Button("Cargar ID")
                                comp_actualizar=gr.Button("💾 Editar / codificar",variant="primary")
                                comp_eliminar=gr.Button("🗑️ Eliminar")
                        with gr.Column():
                            gr.Markdown("#### Copiar / pegar")
                            comp_clip=gr.Textbox(label="Portapapeles",interactive=False)
                            comp_copiar=gr.Button("📋 Copiar competencia seleccionada")
                            comp_pegar=gr.Button("📌 Pegar como nueva")
                            gr.Markdown("Al pegar puedes cambiar el código para reutilizar la misma descripción sin duplicarla manualmente.")
                    comp_msg=gr.Markdown()

                with gr.Tab("Mapa CF ↔ CE"):
                    map_asig=gr.Dropdown(choices=subject_choices(),label="Asignatura")
                    with gr.Row():
                        map_cf=gr.Dropdown(choices=fundamental_choices(),label="Competencia fundamental")
                        map_ce=gr.Dropdown(label="Competencia específica")
                    with gr.Row():
                        map_add=gr.Button("🔗 Asignar CE a CF",variant="primary")
                        map_remove=gr.Button("✂️ Quitar relación")
                        map_refresh=gr.Button("🔄 Actualizar mapa")
                    map_msg=gr.Markdown()
                    map_tabla=gr.Dataframe(interactive=False,label="Relaciones CF ↔ CE")

                cf_ref.click(todas_fundamentales,None,cf_tabla)
                cf_crear.click(crear_fundamental,[cf_codigo_n,cf_nombre_n,cf_desc_n,cf_orden_n],cf_msg).then(todas_fundamentales,None,cf_tabla)
                cf_cargar.click(cargar_fundamental,[cf_id],[cf_codigo,cf_nombre,cf_desc,cf_orden,cf_activa,cf_msg])
                cf_actualizar.click(actualizar_fundamental,[cf_id,cf_codigo,cf_nombre,cf_desc,cf_orden,cf_activa],cf_msg).then(todas_fundamentales,None,cf_tabla)
                cf_eliminar.click(eliminar_fundamental,[cf_id],cf_msg).then(todas_fundamentales,None,cf_tabla)

                comp_asig.change(competencias_con_fundamentales,[comp_asig],comp_tabla)
                comp_refresh.click(competencias_con_fundamentales,[comp_asig],comp_tabla)
                comp_crear.click(crear_competencia,[comp_asig,comp_codigo_n,comp_nombre_n,comp_desc_n,comp_orden_n],comp_msg).then(competencias_con_fundamentales,[comp_asig],comp_tabla)
                comp_cargar.click(cargar_competencia,[comp_id],[comp_codigo_n,comp_nombre_n,comp_desc_n,comp_orden_n,comp_activa,comp_msg])
                comp_actualizar.click(actualizar_competencia,[comp_id,comp_codigo_n,comp_nombre_n,comp_desc_n,comp_orden_n,comp_activa],comp_msg).then(competencias_con_fundamentales,[comp_asig],comp_tabla)
                comp_eliminar.click(lambda cid: (eliminar_actividad(cid) if False else (execute("DELETE FROM competencias WHERE id=?",(int(cid),)) or "🗑️ Competencia específica eliminada.")),[comp_id],comp_msg).then(competencias_con_fundamentales,[comp_asig],comp_tabla)
                comp_copiar.click(copiar_competencia_especifica,[comp_id],[comp_msg,comp_clip])
                comp_pegar.click(pegar_competencia_especifica,[comp_asig,comp_codigo_n,comp_nombre_n,comp_desc_n,comp_orden_n],comp_msg).then(competencias_con_fundamentales,[comp_asig],comp_tabla)

                map_asig.change(opciones_competencias_fundamentales,[map_asig],map_ce).then(listar_mapa_competencias,[map_asig],map_tabla)
                map_add.click(asignar_competencia_fundamental,[map_cf,map_ce],map_msg).then(listar_mapa_competencias,[map_asig],map_tabla)
                map_remove.click(quitar_competencia_fundamental,[map_cf,map_ce],map_msg).then(listar_mapa_competencias,[map_asig],map_tabla)
                map_refresh.click(listar_mapa_competencias,[map_asig],map_tabla)

            # ACTIVIDADES
            with gr.Tab("📝 Actividades"):
                gr.Markdown("""
                ## Actividades evaluativas por período
                Cada actividad pertenece a **una competencia específica**. La calificación P1–P4
                del Registro de Competencias se obtiene automáticamente promediando las actividades
                de la misma competencia y período. La vista general es solo de verificación.
                """)
                with gr.Row():
                    act_est=gr.Dropdown(choices=student_choices(),label="Estudiante")
                    act_asig=gr.Dropdown(choices=subject_choices(),label="Asignatura")
                    act_comp=gr.Dropdown(label="Competencia específica")
                    act_periodo=gr.Dropdown([1,2,3,4],value=1,label="Período")
                with gr.Row():
                    act_fecha=gr.Textbox(label="Fecha",value=today_str())
                    act_titulo=gr.Textbox(label="Actividad / evidencia")
                act_desc=gr.Textbox(label="Descripción",lines=3)
                with gr.Row():
                    act_tipo=gr.Dropdown(TIPOS_EVALUACION,value="Formativa",label="1. Tipo de evaluación")
                    act_instr=gr.Dropdown(INSTRUMENTOS_EVALUACION,label="2. Instrumento de evaluación")
                    act_nota=gr.Number(label="Calificación",minimum=0,maximum=100)
                act_obs=gr.Textbox(label="Observación",lines=2)
                with gr.Row():
                    act_guardar=gr.Button("💾 Registrar actividad",variant="primary")
                    act_ref=gr.Button("🔄 Actualizar actividades")
                    act_id=gr.Number(label="ID para eliminar",precision=0)
                    act_eliminar=gr.Button("🗑️ Eliminar")
                act_msg=gr.Markdown()
                act_tabla=gr.Dataframe(label="Actividades registradas",interactive=False)
                gr.Markdown("### Verificación: promedio general de todas las actividades")
                act_general=gr.Dataframe(label="Promedio general",interactive=False)
                gr.Markdown("### Promedio por grupos de actividades (competencia × período)")
                act_resumen=gr.Dataframe(label="Promedios que alimentan P1–P4",interactive=False)
                act_asig.change(opciones_competencias_fundamentales,[act_asig],act_comp)
                act_guardar.click(guardar_actividad,[act_est,act_asig,act_comp,act_periodo,act_fecha,act_titulo,act_desc,act_tipo,act_instr,act_nota,act_obs],act_msg).then(listar_actividades,[act_asig,act_est,act_periodo,clase_activa],act_tabla).then(lambda a,e,c:(promedio_actividades_general(a,e,None,c),resumen_actividades_por_competencia(a,e,None,c)),[act_asig,act_est,clase_activa],[act_general,act_resumen])
                act_ref.click(listar_actividades,[act_asig,act_est,act_periodo,clase_activa],act_tabla).then(lambda a,e,c:(promedio_actividades_general(a,e,None,c),resumen_actividades_por_competencia(a,e,None,c)),[act_asig,act_est,clase_activa],[act_general,act_resumen])
                act_eliminar.click(eliminar_actividad,[act_id],act_msg).then(listar_actividades,[act_asig,act_est,act_periodo,clase_activa],act_tabla).then(lambda a,e,c:(promedio_actividades_general(a,e,None,c),resumen_actividades_por_competencia(a,e,None,c)),[act_asig,act_est,clase_activa],[act_general,act_resumen])

            # CALIFICACIONES
            with gr.Tab("🎓 Registro de Competencias"):
                gr.Markdown("""
                ## Registro de Grado de Secundaria
                **P1, P2, P3 y P4 son automáticas:** cada una es el promedio de las actividades
                del período pertenecientes a esa competencia. **RP1–RP4** se registran manualmente.
                La matriz conserva la estructura del documento de referencia: P/RP por competencia,
                promedio de competencias específicas y calificación final.
                """)
                with gr.Row():
                    cal_asig=gr.Dropdown(choices=subject_choices(),label="Asignatura",scale=2)
                    cal_est=gr.Dropdown(choices=student_choices(),label="Estudiante",scale=2)
                cal_comp=gr.Dropdown(label="Competencia específica")
                with gr.Row():
                    cal_periodo=gr.Dropdown([1,2,3,4],value=1,label="Período")
                    cal_tipo=gr.Dropdown(["P","RP"],value="P",label="Tipo")
                    cal_nota=gr.Number(label="Calificación",minimum=0,maximum=100)
                cal_obs=gr.Textbox(label="Observación")
                with gr.Row():
                    cal_cargar=gr.Button("Cargar nota")
                    cal_guardar=gr.Button("💾 Guardar RP",variant="primary")
                cal_msg=gr.Markdown()
                cal_detalle=gr.Dataframe(label="Registro del estudiante",interactive=False)
                gr.Markdown("### Matriz general")
                reg_asig=gr.Dropdown(choices=subject_choices(),label="Asignatura")
                reg_tabla=gr.Dataframe(label="Matriz del Registro de Grado",interactive=True)
                with gr.Row():
                    reg_cargar=gr.Button("📋 Cargar registro")
                    reg_guardar=gr.Button("💾 Guardar RP de la matriz",variant="primary")
                    reg_excel=gr.Button("📊 Excel")
                    reg_pdf=gr.Button("📄 PDF")
                reg_msg=gr.Markdown(); reg_file=gr.File(label="Archivo")
                cal_asig.change(opciones_competencias,[cal_asig],cal_comp)
                cal_guardar.click(guardar_calificacion,[cal_est,cal_asig,cal_comp,cal_periodo,cal_tipo,cal_nota,cal_obs],cal_msg).then(notas_estudiante,[cal_est,cal_asig],cal_detalle)
                cal_cargar.click(cargar_nota_actual,[cal_est,cal_asig,cal_comp,cal_periodo,cal_tipo],[cal_nota,cal_obs])
                cal_est.change(notas_estudiante,[cal_est,cal_asig],cal_detalle)
                cal_asig.change(notas_estudiante,[cal_est,cal_asig],cal_detalle)
                reg_cargar.click(registro_grado,[reg_asig,clase_activa],reg_tabla)
                reg_guardar.click(guardar_registro_grado,[reg_asig,reg_tabla],reg_msg).then(registro_grado,[reg_asig,clase_activa],reg_tabla)
                reg_excel.click(exportar_excel,[reg_asig,clase_activa],reg_file)
                reg_pdf.click(exportar_pdf_registro,[reg_asig,clase_activa],reg_file)

            # CALIFICACIÓN FINAL
            with gr.Tab("🏁 Calificación Final / Boletín"):
                gr.Markdown("""
                ## Cierre del año escolar
                La pantalla reproduce los campos del formato final suministrado: **C.F.,
                Calificación Completiva, Calificación Extraordinaria, Calificaciones Especiales
                y Situación Final A/R**. El documento muestra 50%/50% para Completiva y 30%/70%
                para Extraordinaria. El sistema permite modificar estas ponderaciones en Administración.
                """)
                fin_asig=gr.Dropdown(choices=subject_choices(),label="Asignatura")
                fin_tabla=gr.Dataframe(label="Cálculo final por estudiante",interactive=False)
                fin_ref=gr.Button("🔄 Calcular / actualizar")
                gr.Markdown("### Registrar componentes extraordinarios/especiales")
                with gr.Row():
                    fin_est=gr.Dropdown(choices=student_choices(),label="Estudiante")
                    fin_cec=gr.Number(label="C.E.C — Completiva",minimum=0,maximum=100)
                    fin_ceex=gr.Number(label="C.E.EX — Extraordinaria",minimum=0,maximum=100)
                with gr.Row():
                    fin_ecf=gr.Number(label="Especial C.F",minimum=0,maximum=100)
                    fin_ece=gr.Number(label="Especial C.E",minimum=0,maximum=100)
                    fin_sit=gr.Dropdown(["","A","R"],value="",label="Situación final (opcional)")
                fin_obs=gr.Textbox(label="Observación",lines=2)
                fin_guardar=gr.Button("💾 Guardar cierre",variant="primary")
                fin_msg=gr.Markdown()
                gr.Markdown("### Boletín final del estudiante")
                bol_est=gr.Dropdown(choices=student_choices(),label="Estudiante")
                bol_tabla=gr.Dataframe(interactive=False,label="Boletín")
                with gr.Row():
                    bol_ref=gr.Button("📋 Ver boletín")
                    bol_pdf=gr.Button("📄 Descargar boletín PDF",variant="primary")
                bol_file=gr.File(label="Boletín generado")
                fin_asig.change(calculo_final_completo,[fin_asig,clase_activa],fin_tabla)
                fin_ref.click(calculo_final_completo,[fin_asig,clase_activa],fin_tabla)
                fin_guardar.click(guardar_calificacion_final,[fin_est,fin_asig,fin_cec,fin_ceex,fin_ecf,fin_ece,fin_sit,fin_obs],fin_msg).then(calculo_final_completo,[fin_asig,clase_activa],fin_tabla).then(boletin_final_estudiante,[fin_est],bol_tabla)
                bol_ref.click(boletin_final_estudiante,[bol_est],bol_tabla)
                bol_pdf.click(exportar_boletin_pdf,[bol_est],bol_file)

            # PERFIL
            with gr.Tab("👤 Perfil del estudiante"):
                perf_est=gr.Dropdown(choices=student_choices(),label="Estudiante")
                perf_info=gr.Markdown()
                with gr.Row():
                    perf_datos=gr.Dataframe(label="Datos",interactive=False)
                    perf_asist=gr.Dataframe(label="Asistencia",interactive=False)
                perf_inc=gr.Dataframe(label="Incidencias",interactive=False)
                perf_btn=gr.Button("Cargar perfil")
                perf_btn.click(
                    perfil_estudiante,[perf_est],
                    [perf_asist,perf_inc,perf_datos,perf_info]
                )

            # ESTADÍSTICAS
            with gr.Tab("📊 Estadísticas"):
                stats_refresh=gr.Button("Actualizar estadísticas")
                stats_asist=gr.Plot(label="Asistencia")
                stats_inc=gr.Plot(label="Incidencias")
                stats_asist_table=gr.Dataframe(label="Porcentaje de asistencia",interactive=False)
                stats_cal_asig=gr.Dropdown(choices=subject_choices(),label="Asignatura")
                stats_cal=gr.Dataframe(label="Resumen de calificaciones",interactive=False)
                stats_refresh.click(
                    lambda c: (grafico_asistencia(c),grafico_incidencias(c),porcentaje_asistencia(c)),
                    [clase_activa],[stats_asist,stats_inc,stats_asist_table]
                )
                stats_cal_asig.change(resumen_calificaciones,[stats_cal_asig,clase_activa],stats_cal)

            # CALENDARIO
            with gr.Tab("📅 Calendario"):
                with gr.Row():
                    with gr.Column():
                        ev_fecha=gr.Textbox(label="Fecha",value=today_str())
                        ev_hi=gr.Textbox(label="Hora inicio")
                        ev_hf=gr.Textbox(label="Hora fin")
                        ev_titulo=gr.Textbox(label="Título")
                        ev_tipo=gr.Dropdown(["Clase","Examen","Actividad","Reunión","Entrega","Otro"],label="Tipo")
                        ev_curso=gr.Textbox(label="Curso")
                        ev_sec=gr.Textbox(label="Sección")
                        ev_desc=gr.Textbox(label="Descripción",lines=4)
                        ev_btn=gr.Button("➕ Agregar evento",variant="primary")
                        ev_msg=gr.Markdown()
                    with gr.Column():
                        ev_tabla=gr.Dataframe(interactive=False)
                        ev_refresh=gr.Button("Actualizar calendario")
                ev_btn.click(
                    guardar_evento,
                    [ev_fecha,ev_hi,ev_hf,ev_titulo,ev_tipo,ev_curso,ev_sec,ev_desc],
                    ev_msg
                ).then(listar_calendario,None,ev_tabla)
                ev_refresh.click(listar_calendario,None,ev_tabla)

            # REPORTES
            with gr.Tab("📁 Reportes y respaldo"):
                rep_asig=gr.Dropdown(choices=subject_choices(),label="Asignatura")
                with gr.Row():
                    rep_excel=gr.Button("📊 Generar Excel",variant="primary")
                    rep_pdf=gr.Button("📄 Generar PDF")
                    rep_backup=gr.Button("💾 Respaldar base de datos")
                rep_est=gr.Dropdown(choices=student_choices(),label="Estudiante para boletín final")
                rep_boletin=gr.Button("📄 Generar boletín final PDF")
                rep_file=gr.File(label="Archivo")
                rep_msg=gr.Markdown()
                rep_excel.click(exportar_excel,[rep_asig,clase_activa],rep_file)
                rep_pdf.click(exportar_pdf_registro,[rep_asig,clase_activa],rep_file)
                rep_backup.click(backup_db,None,rep_file)
                rep_boletin.click(exportar_boletin_pdf,[rep_est],rep_file)

            # USUARIOS / CENTRO
            with gr.Tab("⚙️ Administración"):
                gr.Markdown("### Usuarios")
                with gr.Row():
                    u_usuario=gr.Textbox(label="Usuario")
                    u_nombre=gr.Textbox(label="Nombre")
                    u_rol=gr.Dropdown(["admin","docente"],value="docente",label="Rol")
                    u_pass=gr.Textbox(label="Contraseña",type="password")
                u_crear=gr.Button("Crear usuario")
                u_msg=gr.Markdown()
                u_tabla=gr.Dataframe(interactive=False)
                u_ref=gr.Button("Actualizar usuarios")
                u_crear.click(crear_usuario,[u_usuario,u_nombre,u_rol,u_pass],u_msg).then(listar_usuarios,None,u_tabla)
                u_ref.click(listar_usuarios,None,u_tabla)

                gr.Markdown("### Docentes")
                gr.Markdown("Los datos de los docentes pueden ser consultados por todos los usuarios. Solo **Administrador** puede registrar, editar o desactivar docentes.")
                with gr.Row():
                    d_sel=gr.Dropdown(choices=docente_choices(),label="Docente para editar",allow_custom_value=False)
                    d_id=gr.Number(label="ID docente",precision=0,visible=False)
                with gr.Row():
                    d_nombre=gr.Textbox(label="Nombre")
                    d_cedula=gr.Textbox(label="Cédula")
                    d_correo=gr.Textbox(label="Correo")
                    d_tel=gr.Textbox(label="Teléfono")
                    d_esp=gr.Textbox(label="Especialidad / Área")
                    d_estado=gr.Dropdown(["Activo","Inactivo"],value="Activo",label="Estado")
                with gr.Row():
                    d_btn=gr.Button("➕ Registrar docente",variant="primary")
                    d_cargar=gr.Button("📥 Cargar")
                    d_actualizar=gr.Button("✏️ Actualizar")
                    d_desactivar=gr.Button("🚫 Desactivar")
                d_msg=gr.Markdown()
                d_tabla=gr.Dataframe(interactive=False)
                d_btn.click(crear_docente,[d_nombre,d_cedula,d_correo,d_tel,d_esp],d_msg).then(listar_docentes,None,d_tabla).then(lambda: gr.update(choices=docente_choices()),None,d_sel)
                d_cargar.click(cargar_docente,[d_sel],[d_nombre,d_cedula,d_correo,d_tel,d_esp,d_estado,d_msg])
                d_actualizar.click(actualizar_docente,[d_id,d_nombre,d_cedula,d_correo,d_tel,d_esp,d_estado],d_msg).then(listar_docentes,None,d_tabla).then(lambda: gr.update(choices=docente_choices()),None,d_sel)
                d_desactivar.click(eliminar_docente,[d_id],d_msg).then(listar_docentes,None,d_tabla).then(lambda: gr.update(choices=docente_choices()),None,d_sel)
                d_sel.change(lambda x: docente_id_from_choice(x),[d_sel],d_id)
                d_ref=gr.Button("🔄 Actualizar lista de docentes")
                d_ref.click(listar_docentes,None,d_tabla).then(lambda: gr.update(choices=docente_choices()),None,d_sel)

                gr.Markdown("### Cursos")
                with gr.Row():
                    c_nombre=gr.Textbox(label="Nombre del curso")
                    c_nivel=gr.Textbox(label="Nivel",value="Secundaria")
                    c_grado=gr.Textbox(label="Grado")
                    c_sec=gr.Textbox(label="Sección")
                    c_anio=gr.Textbox(label="Año escolar",value=get_config("anio_escolar","2026-2027"))
                    c_doc=gr.Number(label="ID docente",precision=0)
                    c_id=gr.Number(label="ID de clase para editar",precision=0)
                    c_estado=gr.Dropdown(["Activo","Inactivo"],value="Activo",label="Estado")
                with gr.Row():
                    c_btn=gr.Button("➕ Crear clase",variant="primary")
                    c_cargar=gr.Button("📥 Cargar clase")
                    c_actualizar=gr.Button("✏️ Actualizar clase")
                    c_eliminar=gr.Button("🚫 Desactivar clase")
                c_msg=gr.Markdown()
                c_tabla=gr.Dataframe(interactive=False)
                c_btn.click(crear_curso,[c_nombre,c_nivel,c_grado,c_sec,c_anio,c_doc],c_msg).then(listar_cursos,None,c_tabla).then(lambda: gr.update(choices=course_choices()),None,[clase_activa,s_curso_id,att_curso])
                c_cargar.click(cargar_curso,[c_id],[c_nombre,c_nivel,c_grado,c_sec,c_anio,c_doc,c_estado,c_msg])
                c_actualizar.click(actualizar_curso,[c_id,c_nombre,c_nivel,c_grado,c_sec,c_anio,c_doc,c_estado],c_msg).then(listar_cursos,None,c_tabla).then(lambda: gr.update(choices=course_choices()),None,[clase_activa,s_curso_id,att_curso])
                c_eliminar.click(eliminar_curso,[c_id],c_msg).then(listar_cursos,None,c_tabla).then(lambda: gr.update(choices=course_choices()),None,[clase_activa,s_curso_id,att_curso])

                gr.Markdown("### Asignaturas de la clase")
                with gr.Row():
                    ca_curso=gr.Dropdown(choices=course_choices(),label="Clase")
                    ca_asig=gr.Dropdown(choices=subject_choices(),label="Asignatura")
                    ca_docente=gr.Dropdown(choices=docente_choices(),label="Docente responsable de la asignatura")
                with gr.Row():
                    ca_add=gr.Button("🔗 Asociar asignatura + docente",variant="primary")
                    ca_remove=gr.Button("✂️ Quitar asociación")
                ca_msg=gr.Markdown(); ca_tabla=gr.Dataframe(interactive=False)
                ca_add.click(asignar_asignatura_curso,[ca_curso,ca_asig,ca_docente],ca_msg).then(listar_asignaturas_curso,[ca_curso],ca_tabla)
                ca_remove.click(quitar_asignatura_curso,[ca_curso,ca_asig],ca_msg).then(listar_asignaturas_curso,[ca_curso],ca_tabla)
                ca_curso.change(listar_asignaturas_curso,[ca_curso],ca_tabla)

                gr.Markdown("### Asignaturas")
                with gr.Row():
                    a_nombre=gr.Textbox(label="Nombre")
                    a_codigo=gr.Textbox(label="Código")
                    a_grado=gr.Textbox(label="Grado")
                a_desc=gr.Textbox(label="Descripción",lines=3)
                a_btn=gr.Button("Crear asignatura")
                a_msg=gr.Markdown()
                a_tabla=gr.Dataframe(interactive=False)
                a_btn.click(
                    crear_asignatura,[a_nombre,a_codigo,a_grado,a_desc],a_msg
                ).then(listar_asignaturas,None,a_tabla)

                gr.Markdown("### Configuración")
                with gr.Row():
                    cfg_anio=gr.Textbox(label="Año escolar",value=get_config("anio_escolar","2026-2027"))
                    cfg_centro=gr.Textbox(label="Nombre del centro",value=get_config("nombre_centro","Centro Educativo"))
                    cfg_max=gr.Number(label="Máxima calificación",value=float(get_config("max_calificacion","100")))
                    cfg_rp=gr.Checkbox(label="RP reemplaza P cuando existe",value=get_config("rp_reemplaza_p","1")=="1")
                with gr.Row():
                    cfg_umbral=gr.Number(label="Umbral de aprobación A/R",value=float(get_config("umbral_aprobacion","70")))
                    cfg_pccf=gr.Number(label="Completiva: % C.F.",value=float(get_config("peso_completiva_cf","50")))
                    cfg_pccec=gr.Number(label="Completiva: % C.E.C.",value=float(get_config("peso_completiva_cec","50")))
                    cfg_pecf=gr.Number(label="Extraordinaria: % C.F.",value=float(get_config("peso_extra_cf","30")))
                    cfg_pecex=gr.Number(label="Extraordinaria: % C.E.EX",value=float(get_config("peso_extra_ceex","70")))
                cfg_btn=gr.Button("Guardar configuración",variant="primary")
                cfg_msg=gr.Markdown()
                cfg_btn.click(guardar_config,[cfg_anio,cfg_centro,cfg_max,cfg_rp,cfg_umbral,cfg_pccf,cfg_pccec,cfg_pecf,cfg_pecex],cfg_msg)

            # ALERTAS
            with gr.Tab("🚨 Alertas"):
                alert_btn=gr.Button("Analizar alertas",variant="primary")
                alert_asist=gr.Dataframe(label="Baja asistencia (<80%)",interactive=False)
                alert_cal=gr.Dataframe(label="Promedios bajos (<70%)",interactive=False)
                alert_btn.click(alertas,[clase_activa],[alert_asist,alert_cal])

        # Actualización de combos
        refresh=gr.Button("🔄 Actualizar listas del sistema")
        refresh.click(
            refresh_all,[clase_activa],
            [inc_est,per_est,cal_est,cal_asig,reg_asig,comp_asig,perf_est,stats_cal_asig,rep_asig]
        )
    refresh.click(lambda: gr.update(choices=docente_choices()),None,ca_docente)

    login_btn.click(login,[login_user,login_pass],[login_box,main_box,login_msg])
    logout_btn.click(logout,None,[login_box,main_box,login_msg])

    # Cambiar de clase actualiza automáticamente las listas de estudiantes.
    clase_refrescar.click(lambda: gr.update(choices=course_choices()),None,clase_activa)
    def texto_clase_activa(c):
        cid=course_id_from_choice(c)
        if not cid:
            return "Selecciona una clase para trabajar con su grupo de estudiantes."
        df=db_df("SELECT nombre,grado,seccion,anio_escolar FROM cursos WHERE id=?",(cid,))
        if df.empty:
            return "Clase no encontrada."
        r=df.iloc[0]
        return f"Clase activa: **{r['nombre']}** — {r['grado']} {r['seccion']} — Año escolar {r['anio_escolar']}"

    def actualizar_listas_por_clase(c):
        upd=gr.update(choices=student_choices(c))
        sub=gr.update(choices=subject_choices(c))
        return (upd,upd,upd,upd,upd,upd,upd,upd,sub,sub,sub,sub,sub,sub,sub,gr.update(choices=course_choices(),value=c),gr.update(choices=course_choices(),value=c),gr.update(choices=course_choices(),value=c),texto_clase_activa(c))

    clase_activa.change(actualizar_listas_por_clase,[clase_activa],[inc_est,per_est,cal_est,act_est,fin_est,bol_est,perf_est,rep_est,cal_asig,reg_asig,comp_asig,act_asig,fin_asig,stats_cal_asig,rep_asig,s_curso_id,att_curso,ca_curso,clase_info])

    # Datos iniciales al abrir componentes de consulta
    dash_refresh.click(
        lambda c: (
            len(listar_estudiantes_clase(c)),
            len(listar_incidencias(c)),
            len(listar_permisos(c)),
            int(porcentaje_asistencia(c)["dias_registrados"].sum()) if not porcentaje_asistencia(c).empty else 0,
            porcentaje_asistencia(c)
        ),
        [clase_activa],[total_students,total_inc,total_perm,total_att,dash_table]
    )
    s_refrescar.click(lambda: buscar_estudiantes(""),None,s_tabla)
    inc_refresh.click(listar_incidencias,[clase_activa],inc_tabla)
    per_refresh.click(listar_permisos,[clase_activa],per_tabla)
    ev_refresh.click(listar_calendario,None,ev_tabla)
    u_ref.click(listar_usuarios,None,u_tabla)
    c_btn.click(listar_cursos,None,c_tabla)
    a_btn.click(listar_asignaturas,None,a_tabla)


# ============================================================
# 21. LANZAMIENTO
# ============================================================

print("="*70)
print("SISTEMA DE GESTIÓN DE CLASE V5")
print("="*70)
print(f"Base de datos: {DB_PATH}")
print("Usuario inicial: admin")
print("Contraseña inicial: admin123")
print("Cambia la contraseña creando un usuario nuevo y desactivando el inicial.")
print("="*70)

app.launch(share=True, debug=False)
