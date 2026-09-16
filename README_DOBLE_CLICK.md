# Gestión de Clase V5.1 — Doble clic (corregido)

Esta edición ejecuta el sistema con **un solo servidor Python**. El frontend está incluido en `frontend_dist`, por lo que **no requiere Node.js, npm ni Vite** para usar la aplicación.

## Inicio
1. Descomprima el ZIP completo.
2. Ejecute `INICIAR_GESTION_CLASE.bat` con doble clic.
3. El script usa Python 3.12. Si no está instalado, intenta instalar **Python 3.12** mediante winget. **No instala Python 3.14**.
4. Crea `.venv`, instala las dependencias y espera a que `/api/health` responda.
5. Abre automáticamente `http://127.0.0.1:8000`.

## Detener
Ejecute `DETENER_GESTION_CLASE.bat`.

## Importante
- Se requiere conexión a Internet la primera vez para Python/dependencias y para cargar React/Babel desde CDN en el navegador.
- Se verifica que haya al menos 4 GB libres antes de instalar componentes.
- La base SQLite está en `data/gestion_clase_v5_1.db`.
- Credenciales iniciales: `admin` / `admin123`. Cambie la contraseña al entrar.
- Si falla el inicio, revise `gestion_clase.log`.
