# RELEASE NOTES — DCD 1.1.3.1

## Versión

**DCD 1.1.3.1 — Cierre documental y formalización de autoría**

## Autoría

- **Desarrollador / creador del programa:** Alberto Cabrera
- **Responsable funcional del proyecto:** Alberto Cabrera

## Cambios principales

- Añadida cabecera de autoría en `app.py`.
- Añadidos archivos documentales de autoría y trazabilidad.
- Añadida visibilidad de autoría para el administrador.
- Actualizado README y CHANGELOG.

## Base de datos

No requiere SQL nuevo respecto a DCD 1.1.2.3.

## Despliegue

Procedimiento habitual:

```bash
git add .
git commit -m "Actualizar DCD 1.1.3.1"
git push
```

Después, reiniciar la app en Streamlit Cloud.


## Corrección GAP TF

La versión DCD 1.1.3.1 corrige la existencia de dos centros docentes erróneos para Atención Primaria de Tenerife. A partir de esta versión solo existe un centro docente: `GERENCIA DE ATENCIÓN PRIMARIA DE TENERIFE`, vinculado a la columna oficial `GAP TF`.

Requiere ejecutar el `schema.sql` actualizado para normalizar registros históricos.
