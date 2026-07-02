# RELEASE NOTES — DCD 1.2.0 RC1

## Versión

**DCD 1.2.0 RC1 — Turnos y observaciones por titulación**

## Autoría

- **Desarrollador / creador del programa:** Alberto Cabrera
- **Responsable funcional del proyecto:** Alberto Cabrera

## Cambios principales

- Nueva captura de distribución de alumnos por turno en el Paso 4.
- Campos de alumnos en Mañana, Tarde, Rotatorio y Deslizante.
- Patrón semanal para turno deslizante de lunes a viernes con selección M/T/R.
- Nuevo campo de Observaciones por titulación/especialidad.
- Edición de registros ya introducidos para poder corregir turnos u observaciones sin eliminar el registro.
- Validación de que la suma de turnos coincide con el número total de alumnos.
- Bloqueo de envío si existen registros con distribución de turnos incoherente.
- Exportación de turnos y observaciones en Excel admin, Excel consulta y PDF consulta.
- Pulido del PDF de consulta para compactar patrones deslizantes y ajustar observaciones largas.
- Dashboard de consulta mantenido sin cambios para no comprometer la visualización territorial ya validada.

## Base de datos

Esta versión requiere la migración:

```text
supabase/migrations/2026_06_27_dcd_1_2_0_turnos_observaciones.sql
```

Columnas añadidas a `dcd_borradores` y `dcd_registros`:

```text
alumnos_manana
alumnos_tarde
alumnos_rotatorio
alumnos_deslizante
deslizante_lunes
deslizante_martes
deslizante_miercoles
deslizante_jueves
deslizante_viernes
observaciones_titulacion
```

## Validación realizada

- Rol usuario: captura, edición, guardado, recuperación de borrador, finalización y envío de correo correctos.
- Rol admin: publicación y Excel admin correctos.
- Rol consulta: dashboard, Excel consulta y PDF consulta correctos.

## Rama de desarrollo

```text
feature-turnos-dcd-1-2-0
```

## Despliegue

Procedimiento habitual en la rama beta/RC:

```bash
git add .
git commit -m "Preparar DCD 1.2.0 RC1 turnos y observaciones"
git push
```

Después, reiniciar la app beta en Streamlit Cloud.


## DCD 1.2.1 beta 2 — Multiusuario por centro docente

Versión beta para validar la configuración administrativa de centros multiusuario antes de activar la consolidación operativa.


## DCD 1.2.1 beta 3

- Corrección del registro de publicaciones en Supabase cuando existen valores NaN/NaT procedentes de tablas Pandas.
- No requiere cambios SQL.
- No modifica lógica de cálculo, Excel ni PDF.

## DCD 1.2.1 beta 4

- Añade consolidación parcial manual por administrador para centros multiusuario incompletos.
- Permite que entren al consolidado general las aportaciones finalizadas cuando falten usuarios, dejando trazabilidad de usuarios pendientes, motivo y administrador responsable.
- No requiere SQL nuevo.



## DCD 1.2.2 beta 2

- Ajuste de visualización del campo Detalle alumnos en la tabla de registros introducidos.
- Evita que Streamlit muestre listas internas como [object Object].
- No modifica guardado, recuperación, Excel, PDF, Supabase ni cálculos.

## DCD 1.2.2 beta 1

- Añadido detalle opcional por alumno en el Paso 4: Servicio y Curso/año.
- Persistencia en `detalle_alumnos` como JSONB en borradores y registros.
- Edición y recuperación de detalle por alumno en registros existentes.
- Nueva hoja `Detalle_Alumnos` en Excel admin y Excel de consulta.
- No se modifica el Dashboard ni se incorpora al PDF principal en esta beta.
