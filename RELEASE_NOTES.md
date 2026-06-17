# RELEASE NOTES — DCD 1.1.3.3

## Versión

**DCD 1.1.3.3 — Ajustes finales de visualización y cierre funcional**

## Autoría

- **Desarrollador / creador del programa:** Alberto Cabrera
- **Responsable funcional del proyecto:** Alberto Cabrera

## Cambios principales

- Cambio del mensaje de confirmación del envío automático de correo para mostrar: `Correo enviado correctamente a servicio de FSE`.
- Limpieza visual del portal de consulta: se ocultan filas analíticas con `Total plazas = 0`.
- Limpieza visual del PDF de Matriz_DCD: se ocultan filas de titulaciones con `Total = 0`, manteniendo la fila final `TOTAL`.
- El cambio es solo de presentación; no modifica registros, borradores, cálculos, Supabase ni la estructura oficial de la matriz.

## Base de datos

No requiere SQL nuevo respecto a DCD 1.1.3.2.

## Despliegue

Procedimiento habitual:

```bash
git add .
git commit -m "Actualizar DCD 1.1.3.3 ajustes finales visualizacion"
git push
```

Después, reiniciar la app en Streamlit Cloud.

## Nota sobre PDFs existentes

Los PDFs ya publicados antes de esta versión no cambian automáticamente. Para que el PDF vigente salga con la limpieza visual de filas `Total = 0`, genere una nueva publicación después de desplegar esta versión.
