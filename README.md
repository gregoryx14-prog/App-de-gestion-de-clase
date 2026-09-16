# Gestión de Clase V5.1 Web — Etapa 20

Versión de consolidación final del proyecto. Mantiene los módulos acumulados de las etapas 1–19 y agrega herramientas de diagnóstico y preparación para despliegue.

## Módulos acumulados
1. Base V5.1 web y docentes.
2. Usuarios, autenticación, roles y vinculación docente.
3. Competencias, actividades y calificaciones.
4. Carga académica y control de acceso docente.
5. Asistencia, incidencias y permisos.
6. Cierre académico / Registro de Grado.
7. Reportes PDF/XLSX/CSV.
8. Dashboard e indicadores.
9. Expedientes estudiantiles.
10. Seguimiento y reuniones familiares.
11. Alertas de riesgo y planes de intervención.
12. Comunicaciones.
13. Agenda y calendario.
14. Horario escolar.
15. Años escolares y períodos académicos.
16. Planificación académica.
17. Evaluación docente avanzada y libro de calificaciones.
18. Seguridad, auditoría y configuración institucional.
19. Importación, exportación y respaldos.
20. Consolidación, diagnóstico y preparación de despliegue.

## Etapa 20
- `GET /api/health`: comprobación pública de disponibilidad e integridad SQLite.
- `GET /api/sistema/estado`: diagnóstico autenticado con versión, integridad y conteo de tablas.
- Panel administrativo **Estado del sistema**.
- Docker backend con `HEALTHCHECK`.
- Frontend con build multi-stage y Nginx.
- `docker-compose.prod.yml` para despliegue de frontend/backend.
- Prueba automatizada básica en `tests/test_stage20.py`.

## Ejecución local
### Backend
```bash
cd backend
pip install -r requirements.txt
uvicorn app.main:app --reload
```

### Frontend
```bash
cd frontend
npm install
npm run dev
```

## Producción con Docker
Antes de desplegar, cambie `GESTION_SECRET_KEY` por una clave aleatoria y privada. Revise también el dominio/origen permitido por CORS y los puertos publicados.

```bash
docker compose -f docker-compose.prod.yml up --build -d
```

## Pruebas
```bash
PYTHONPATH=backend python tests/test_stage20.py
```

## Nota de arquitectura
La aplicación mantiene SQLite para facilitar instalación local y despliegues pequeños. Para un centro con alta concurrencia o varios procesos de aplicación, PostgreSQL puede evaluarse posteriormente como una evolución de infraestructura; no es requisito para el funcionamiento actual.
