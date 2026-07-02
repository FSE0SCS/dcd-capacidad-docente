# CHANGELOG

## DCD 1.2.1 beta 2
- Rol admin entra directamente al Panel administrador.
- Trazabilidad de aportaciones por usuario en centros multiusuario.
- Estado de aportación por usuario en pestaña Multiusuario.
- El consolidado general solo incorpora centros multiusuario cuando todos los usuarios activos asignados han finalizado.

# CHANGELOG — DATOS CAPACIDAD DOCENTE (DCD)

## DCD 1.2.0 RC1

- Versión candidata estable de la nueva fase funcional de turnos y observaciones por titulación.
- Captura de distribución de alumnos por Mañana, Tarde, Rotatorio y Deslizante.
- Captura del patrón semanal del turno deslizante de lunes a viernes con M/T/R.
- Nuevo campo de Observaciones por titulación/especialidad, visible en Excel y PDF de consulta.
- Edición de registros ya introducidos en el Paso 4.
- Validación de coherencia entre número total de alumnos y suma de turnos.
- Guardado, recuperación de borrador, finalización y envío de correo validados.
- Excel admin, Excel consulta y PDF consulta validados con turnos y observaciones.
- Dashboard de consulta se mantiene sin añadir turnos para preservar la visualización territorial validada.

## DCD 1.1.3.4

- Incorporación de mapa visual de Canarias en el dashboard del rol consulta.
- El mapa muestra la capacidad docente total por isla sobre imagen institucional limpia.
- Cambio limitado a visualización: no modifica Supabase, publicaciones, Excel, PDF ni cálculos base.

## DCD 1.1.3.3

- Ajuste del mensaje de confirmación del envío automático de correo para no mostrar la dirección técnica del destinatario.
- Ocultación visual de filas con `Total = 0` en dashboard de consulta y PDF de Matriz_DCD.
- Cambio limitado a presentación: sin modificación de base de datos, cálculos de consolidación ni estructura oficial de la matriz.

## DCD 1.1.3.2

- Corrección del botón “Preparar Excel” para usuarios de rol consulta.
- Añadida compatibilidad `is_consulta_user()` con el visor externo.

## DCD 1.1.3.1

- Corrección GAP TF: centro único “GERENCIA DE ATENCIÓN PRIMARIA DE TENERIFE”, código GAPTF y columna Excel `GAP TF`.

## DCD 1.1.3

- Cierre documental del proyecto.
- Formalización de autoría en código y documentación.
- Inclusión de `AUTHORS.md`, `AUTHORSHIP.md`, `CHANGELOG.md` y `RELEASE_NOTES.md`.
- Identificación del desarrollador/creador del programa: Alberto Cabrera.
- Preparación para cierre estable y trazabilidad del paquete final.

## DCD 1.1.2.3

- Excel limitado para usuarios de consulta.
- Ajuste del logo en PDF en parte superior derecha.

## DCD 1.1.2.2

- Ajustes finales de PDF: logo, frase institucional, pie de firma y numeración de páginas.

## DCD 1.1.2

- Validaciones de calidad de datos.
- Comparativa entre publicaciones.

## DCD 1.1.1

- Auditoría ampliada.
- Backup administrativo.

## DCD 1.1.0

- Dashboard visual y presentación profesional.

## DCD 1.0.x

- MVP inicial.
- Supabase/PostgreSQL.
- Usuarios y roles.
- Guardado versionado.
- Publicaciones oficiales.
- Portal de consulta.


## DCD 1.1.3.5
- Corrección del PDF de consulta: se genera dinámicamente sin bloque interno de calidad de datos.
- Se filtran filas y resúmenes con total 0 en el PDF de consulta.
- Ajuste de coordenadas y proporciones del mapa de capacidad docente por isla.


## DCD 1.2.1 beta 2

- Añadida primera configuración administrativa de centros multiusuario.
- Nueva pestaña admin `Multiusuario` para marcar centros, definir usuarios previstos y asignar usuarios existentes.
- Preparación de trazabilidad para futura consolidación por centro.


## DCD 1.2.1 beta 3

- Corrección del registro de publicaciones en Supabase cuando existen valores NaN/NaT procedentes de tablas Pandas.
- No requiere cambios SQL.
- No modifica lógica de cálculo, Excel ni PDF.

## DCD 1.2.1 beta 4

- Añade consolidación parcial manual por administrador para centros multiusuario incompletos.
- Permite que entren al consolidado general las aportaciones finalizadas cuando falten usuarios, dejando trazabilidad de usuarios pendientes, motivo y administrador responsable.
- No requiere SQL nuevo.


## DCD 1.2.2 beta 1

- Añadido detalle opcional por alumno en el Paso 4: Servicio y Curso/año.
- Persistencia en `detalle_alumnos` como JSONB en borradores y registros.
- Edición y recuperación de detalle por alumno en registros existentes.
- Nueva hoja `Detalle_Alumnos` en Excel admin y Excel de consulta.
- No se modifica el Dashboard ni se incorpora al PDF principal en esta beta.
