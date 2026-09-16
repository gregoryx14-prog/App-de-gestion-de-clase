# Fase de Pruebas y Estabilización — Gestión de Clase V5.1 Web

## Alcance
Se ejecutó una prueba de recorrido completo sobre una base SQLite temporal para no modificar datos del proyecto. La prueba cubre el flujo acumulativo de las etapas 1–20.

## Matriz de validación

| Área | Resultado |
|---|---|
| Salud y autenticación | OK |
| Usuarios y roles | OK |
| Docentes y especialidades | OK |
| Vinculación usuario-docente | OK |
| Años y períodos académicos | OK |
| Clases y asignaturas | OK |
| Matrículas y estudiantes | OK |
| Competencias fundamentales/específicas | OK |
| Planificación | OK |
| Actividades evaluativas | OK |
| Asistencia | OK |
| Incidencias | OK |
| Permisos | OK |
| Libro de calificaciones | OK |
| Cierre académico | OK |
| Reportes PDF/XLSX | OK |
| Alertas y planes de intervención | OK |
| Expedientes | OK |
| Comunicaciones | OK |
| Agenda | OK |
| Horario y detección de conflictos | OK |
| Importación CSV | OK |
| Exportación general | OK |
| Backup / restore | OK |
| Restricciones admin/docente | OK |
| Integridad SQLite | OK |
| Auditoría | OK |

## Incidencias encontradas y corregidas

1. **Permisos:** el `PUT` exigía el modelo completo aunque la operación normal puede necesitar cambiar únicamente estado/observación. Se añadió un modelo de actualización parcial.
2. **Importación de estudiantes:** se usaba `lastrowid` sobre la conexión en lugar del cursor. Corregido y además se completan grado, sección y edad.
3. **Exportación general:** se pasaba un `Workbook` a un helper que esperaba encabezados/filas. Se añadió un helper para serializar libros XLSX completos.
4. **Frontend:** había una colisión de estado `comForm` entre Comunicados y Compromisos. Se separó como `comForm` y `compromisoForm` y se corrigieron sus manejadores.
5. **Configuración:** CORS ahora admite `GESTION_CORS_ORIGINS` para adaptar el dominio de despliegue.
6. **Diagnóstico:** el estado del sistema ahora informa también errores de claves foráneas.

## Comandos de validación

```bash
python -m py_compile backend/app/main.py tests/test_stage20.py tests/test_e2e_stabilization.py tests/run_all.py
PYTHONPATH=backend python tests/run_all.py
```

Resultado obtenido:

```text
STAGE20_TEST_OK
E2E_STABILIZATION_OK
ALL_STABILIZATION_TESTS_OK
```

## Validación del frontend
Se intentó `npm install && npm run build`. El entorno de ejecución agotó el tiempo durante la instalación de dependencias, por lo que el build Vite no se marca como validado aquí. Debe ejecutarse localmente antes de un despliegue real.

## Estado recomendado
**Backend y flujo funcional: estabilizados mediante pruebas automatizadas.**

**Frontend: pendiente de una ejecución local de `npm install` + `npm run build` debido a la limitación del entorno de esta prueba.**

## Checklist antes de usar datos reales
- [ ] Cambiar `admin / admin123`.
- [ ] Definir `GESTION_SECRET_KEY` privada y aleatoria.
- [ ] Definir `GESTION_CORS_ORIGINS` con el dominio real.
- [ ] Ejecutar build del frontend.
- [ ] Probar backup y restore con una copia.
- [ ] Verificar que los reportes se abren correctamente.
- [ ] Crear usuarios/docentes reales y comprobar sus ámbitos de acceso.
