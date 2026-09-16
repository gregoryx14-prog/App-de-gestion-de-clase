# Gestión de Clase V5.1 Web — Etapa 20 + Fase de Pruebas y Estabilización

## Objetivo
Consolidar la versión acumulativa V5.1 Web y verificar el flujo completo antes de usar datos reales.

## Correcciones realizadas durante la estabilización
- Actualización parcial de permisos: ahora `PUT /api/permisos/{id}` permite cambiar solo los campos necesarios sin exigir todo el formulario.
- Corrección de importación CSV de estudiantes: se obtiene correctamente el `lastrowid`, se guardan grado/sección/edad y se conserva el formato de matrícula basado en iniciales.
- Corrección de exportación general XLSX: se genera correctamente un libro con múltiples hojas.
- Corrección del frontend: se eliminó la colisión de estado `comForm` entre Comunicados y Compromisos y se corrigieron sus manejadores.
- CORS configurable mediante `GESTION_CORS_ORIGINS`.
- Diagnóstico del sistema ampliado con comprobación de claves foráneas.

## Pruebas automatizadas
`tests/test_stage20.py` comprueba salud, login, autenticación y estado de SQLite.

`tests/test_e2e_stabilization.py` ejecuta un recorrido funcional acumulativo: autenticación, docentes, usuarios, año escolar, períodos, clases, asignaturas, matrícula, competencias, planificación, actividades, asistencia, incidencias, permisos, calificaciones, cierre, reportes PDF/XLSX, alertas, intervención, expediente, comunicaciones, agenda, horario, importación CSV, exportación, respaldo/restauración y controles de rol.

Resultado verificado:
```text
STAGE20_TEST_OK
E2E_STABILIZATION_OK
```

## Verificaciones técnicas
- `python -m py_compile` del backend y pruebas: OK.
- Integridad SQLite en pruebas: OK.
- Conflictos de horario: comprobados.
- Restricciones de acceso admin/docente: comprobadas en el flujo E2E.
- Reportes PDF/XLSX: generados durante la prueba.
- Backup y restore SQLite: comprobados durante la prueba.

## Frontend
Se intentó ejecutar `npm install && npm run build`, pero el entorno de ejecución agotó el tiempo disponible durante la instalación. Por eso esta fase **no declara un build Vite completo como validado**. El código fuente sí fue corregido para eliminar la colisión de estados detectada.

En una máquina local ejecutar:
```bash
cd frontend
npm install
npm run build
```

## Antes de producción
1. Cambiar `GESTION_SECRET_KEY` por una clave aleatoria de al menos 32 caracteres.
2. Configurar `GESTION_CORS_ORIGINS` con el dominio real del frontend.
3. Cambiar la contraseña inicial `admin / admin123`.
4. Crear una copia de seguridad antes de cargar datos reales.
5. Probar una restauración en una copia antes de restaurar sobre producción.
6. Para múltiples procesos/alta concurrencia, evaluar PostgreSQL.

## Docker
```bash
docker compose -f docker-compose.prod.yml up --build -d
```
