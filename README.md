# DATOS CAPACIDAD DOCENTE (DCD 1.2.1 beta 2)

Aplicativo Streamlit para recoger datos de capacidad docente a partir del Excel base `listado_para_capacidad_docente.xlsx`.
## Autoría del proyecto

- **Desarrollador / creador del programa:** Alberto Cabrera
- **Responsable funcional del proyecto:** Alberto Cabrera
- **Desarrollado para:** F.S.E. – S.C.S.
- **Proyecto:** DATOS CAPACIDAD DOCENTE (DCD)
- **Versión candidata estable:** DCD 1.2.0 RC1
- **Año:** 2026

La autoría funcional y de desarrollo queda identificada en `app.py`, `README.md`, `AUTHORSHIP.md`, `CHANGELOG.md` y en la huella SHA256 del paquete de distribución.


Esta versión parte de la DCD 1.0 y añade:

- Pantalla específica de recordatorio antes de finalizar.
- Envío automático opcional por Mailgun.
- Control básico/avanzado de usuarios mediante `USERS_JSON`.
- Auditoría de acciones en Supabase.
- Revisión visible del mapeo entre centros docentes y columnas del Excel.
- Ajuste del campo Área: `HOSPITAL`, `ATENCION FAMILIAR Y COMUNITARIA` y retirada de `OTRAS UNIDADES DOCENTES`.
- Exportación Excel más completa, con estado, usuario, observaciones y datos generales en las hojas de salida.
- Control de estado `borrador` / `finalizado` con bloqueo suave de edición para expedientes finalizados.
- Totales automáticos en la hoja `Matriz_DCD`, incluida la fila total final.
- Panel administrador para generar un Excel consolidado desde los expedientes finalizados guardados en Supabase.
- Gestión de usuarios en Supabase con contraseñas hasheadas.
- Usuarios vinculados a un centro docente y borradores filtrados por centro.
- Reset de contraseña por administrador, sin visualizar contraseñas antiguas.
- Corrección del selector de centro docente en el panel de creación/actualización de usuarios.
- Guardado versionado: cada guardado crea una nueva versión y no sobrescribe la anterior.
- Panel administrador con estado de centros finalizados, pendientes y sin datos.
- Aviso por correo de centros pendientes, si Mailgun está configurado.
- Cambio visible de `Unidad Docente` a `Centro Docente`.
- Portal externo de consulta de publicación vigente con dashboard y descargas controladas.

- Mantenimiento avanzado de usuarios: activar, desactivar, resetear y eliminar con confirmación.
- Dashboard visual mejorado para portal de consulta y panel de publicaciones.
- Tarjetas compactas para evitar cortes en publicación vigente, versión y fecha.
- PDF preparado para incorporar `assets/logo.png` y frase institucional.
- Auditoría ampliada: login correcto/fallido, consultas externas, descargas PDF/Excel y backup admin.
- Panel administrador `Auditoría/Backup` con filtros y exportación segura de tablas principales.
- Panel administrador `Calidad de datos` con validaciones, alertas y comparativa entre publicaciones.


## DCD 1.2.0 RC1 — Turnos y observaciones por titulación

Esta versión candidata incorpora la nueva fase funcional de turnos y observaciones por titulación:

- Captura de alumnos por turno: Mañana, Tarde, Rotatorio y Deslizante.
- Patrón semanal para turno deslizante de lunes a viernes con valores M/T/R.
- Observaciones específicas por titulación/especialidad.
- Edición de registros ya introducidos en el Paso 4.
- Validación de que la suma de turnos coincide con el número total de alumnos.
- Persistencia en borradores y registros de Supabase.
- Recuperación correcta de borradores con turnos y observaciones.
- Exportación de turnos y observaciones a Excel admin, Excel consulta y PDF consulta.
- Dashboard de consulta mantenido sin cambios para preservar la visualización territorial validada.

## 1. Estructura del proyecto

```text
dcd_capacidad_docente/
├── app.py
├── requirements.txt
├── README.md
├── AUTHORS.md
├── AUTHORSHIP.md
├── CHANGELOG.md
├── RELEASE_NOTES.md
├── .gitignore
├── .streamlit/
│   └── secrets.toml.example
├── data/
│   └── listado_para_capacidad_docente.xlsx
├── assets/
│   └── logo.png  # opcional
└── supabase/
    └── schema.sql
```

## 2. Instalación local

Desde la carpeta del proyecto:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

En macOS/Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## 3. Secrets locales

Copia:

```text
.streamlit/secrets.toml.example
```

como:

```text
.streamlit/secrets.toml
```

Contenido mínimo:

```toml
APP_PASSWORD = "Capacidad2026"
SUPABASE_URL = "https://TU-PROYECTO.supabase.co"
SUPABASE_KEY = "TU_CLAVE_SERVICE_ROLE_O_SECRET"
```

Si no configuras usuarios avanzados, se entra con:

```text
Usuario: admin
Contraseña: Capacidad2026
```

## 4. Usuarios avanzados opcionales

Puedes añadir varios usuarios en los secretos:

```toml
USERS_JSON = '{"admin":{"password":"Capacidad2026","role":"admin","display_name":"Administrador"},"chuimi":{"password":"ClaveCHUIMI","role":"usuario","display_name":"CHUIMI"}}'
```

El control de usuarios permite vincular cada usuario a un centro docente concreto y filtrar sus borradores por ese centro.

## 5. Supabase

En Supabase, abre:

```text
SQL Editor > New query
```

Pega el contenido de:

```text
supabase/schema.sql
```

y ejecuta `Run`.

Si ya habías creado las tablas de versiones anteriores, este SQL añade las columnas necesarias para usuarios, versionado de borradores y control de centros pendientes sin borrar los datos existentes.

## 6. Mailgun opcional

Para activar el envío automático de correo, añade estos secrets:

```toml
MAILGUN_API_KEY = "key-xxxxxxxx"
MAILGUN_DOMAIN = "mg.tudominio.com"
MAILGUN_SENDER_EMAIL = "DCD <postmaster@mg.tudominio.com>"
MAILGUN_RECIPIENT_EMAIL = "destinatario@dominio.com"
```

Si Mailgun no está configurado, la app sigue funcionando: podrás descargar el Excel y enviarlo manualmente.

## 7. Subida a GitHub

Comprueba que `.streamlit/secrets.toml` no se sube. El `.gitignore` debe incluirlo.

Comandos habituales:

```bash
git add .
git commit -m "Actualizar DCD 1.0.9.1"
git push
```

## 8. Streamlit Community Cloud

En la app desplegada:

```text
Settings > Secrets
```

Pega los mismos secretos que usas en local, sin subirlos a GitHub.

Después de cambiar secretos, haz siempre reboot/restart de la app para que Streamlit los vuelva a cargar.

## 9. Flujo funcional

1. Login.
2. Aceptación de instrucciones.
3. Selección de área y centro docente.
4. Confirmación y revisión de mapeo Excel.
5. Entrada de datos por selectores dependientes.
6. Guardado de borrador.
7. Pantalla de recordatorio.
8. Descarga de Excel.
9. Guardado como finalizado.
10. Envío automático opcional.

## 10. Publicaciones oficiales DCD 1.0.9.1

Esta versión añade un módulo de publicaciones oficiales:

- Genera Excel consolidado desde Supabase.
- Genera PDF de la hoja `Matriz_DCD`.
- Guarda ambos archivos en Supabase Storage, bucket privado `dcd-publicaciones`.
- Registra cada publicación en `dcd_publicaciones`.
- Marca una única publicación como `publicacion_vigente = true`.
- Conserva las publicaciones anteriores como histórico para el administrador.
- Envía correo al administrador si Mailgun está configurado.
- Si todos los centros están finalizados, la app puede generar una publicación automática al finalizar el último centro.
- El administrador puede generar manualmente una publicación vigente, incluso con centros pendientes, dejando constancia en el registro.

Para esta versión es obligatorio ejecutar de nuevo:

```text
supabase/schema.sql
```

porque añade la tabla `dcd_publicaciones` y el bucket `dcd-publicaciones` en Supabase Storage.


## 11. Dashboard y análisis DCD 1.0.8

La versión 1.0.8 añade explotación analítica de la Matriz_DCD:

- Dashboard web en el panel de publicaciones.
- Resumen global de plazas.
- Resumen por provincia: Las Palmas y S/C Tenerife.
- Resumen por isla:
  - Gran Canaria = CHUIMI + HUGC DN + GAP GC.
  - Fuerteventura = GSS FV.
  - Lanzarote = GSS LZ.
  - Tenerife = GAP TF + CHUC + HUNSC.
  - La Palma = GSS LP.
  - La Gomera = GSS LG.
  - El Hierro = GSS EH.
- Resumen por centro docente/columna.
- Resumen por rama.
- Resumen por nivel de estudios.
- Resumen centro-rama y centro-nivel.
- Top titulaciones por número de plazas.
- Hojas analíticas añadidas al Excel consolidado.
- PDF de publicación con bloque inicial de dashboard/resumen antes de la matriz completa.

La versión 1.0.8 no requería cambios de base de datos respecto a la versión 1.0.7.


## 12. Cierre configurable DCD 1.0.9.1

La versión 1.0.8.2 incorpora y refuerza el bloque de cierre configurable:

- Fecha tope de cierre.
- Modo de cierre configurable desde el panel administrador:
  1. Publicar automáticamente solo cuando todos finalicen.
  2. Publicar automáticamente al llegar la fecha tope aunque falten centros.
  3. Solo publicación manual por admin.
- Avisos previos al administrador si hay centros pendientes y se acerca la fecha tope.
- Aviso manual de centros pendientes.
- Registro del motivo de publicación automática o por fecha tope.
- Evaluación manual del cierre automático desde el panel administrador.

Importante: Streamlit no ejecuta tareas en segundo plano de forma permanente. La fecha tope y los avisos se evalúan cuando un usuario accede a la app, cuando un centro finaliza expediente o cuando el admin pulsa evaluación manual.

Para esta versión es obligatorio ejecutar de nuevo:

```text
supabase/schema.sql
```

porque añade la tabla `dcd_configuracion`.


## 13. DCD 1.0.9.1 - Refuerzo de cierre automático

Esta versión deja explícitos y reforzados estos disparadores de cierre:

- Comprobación automática cuando un centro finaliza expediente.
- Comprobación automática cuando cualquier usuario entra en la app.
- Botón de administrador **Evaluar ahora cierre automático**.
- Aviso claro en panel de cierre: Streamlit no ejecuta procesos permanentes en segundo plano; la fecha tope se evalúa cuando la app es utilizada o mediante evaluación manual del admin.

No requiere SQL nuevo si ya se ejecutó el `schema.sql` de la 1.0.8.1.

## 14. Portal externo DCD 1.0.9.1

La versión 1.0.9 añade un portal de consulta para entidades externas.

Características:

- Nuevo rol de usuario: `consulta`.
- Los usuarios con rol `consulta` no acceden al flujo de carga ni al panel administrador.
- Solo ven la publicación marcada como vigente.
- Ven un dashboard de lectura construido desde el Excel de la publicación vigente.
- Pueden descargar el PDF de informe y el Excel consolidado vigentes.
- Las publicaciones históricas siguen estando reservadas al administrador.

No requiere SQL nuevo si ya existe la tabla `dcd_usuarios` creada en versiones anteriores.

Para crear un usuario externo:

1. Entrar como `admin`.
2. Ir a `Panel administrador > Usuarios`.
3. Crear usuario con rol `consulta`.
4. Asignar contraseña temporal.
5. El usuario externo accederá directamente al portal de publicación vigente.


## DCD 1.0.9.1 - Mantenimiento

- Ajuste visual de tarjetas de publicación vigente para evitar textos cortados.
- Mantenimiento avanzado de usuarios: activar, desactivar, resetear contraseña y eliminar con confirmación.
- Recomendación: usar desactivar para bajas ordinarias y eliminar solo cuando proceda limpiar la tabla de usuarios.
- Si no se ejecutó el SQL del rol consulta, ejecutar `supabase/schema.sql`.


## Logo institucional

Para incorporar el logo al PDF de publicación, coloque el archivo en:

```text
assets/logo.png
```

La aplicación lo detectará automáticamente al generar el PDF. Si no existe, el PDF se generará igualmente sin logo.

## 19. DCD 1.1.1 - Auditoría y backup administrativo

Esta versión no requiere SQL nuevo si ya existe la tabla `dcd_auditoria`. Añade:

- Registro de intentos de login fallidos.
- Registro controlado de consultas a la publicación vigente.
- Registro de descargas de PDF/Excel desde portal de consulta y panel administrador.
- Panel `Auditoría/Backup` para administradores.
- Filtros de auditoría por acción y usuario.
- Backup Excel de tablas principales:
  - `dcd_usuarios`
  - `dcd_borradores`
  - `dcd_registros`
  - `dcd_publicaciones`
  - `dcd_auditoria`
  - `dcd_configuracion`

Por seguridad, el backup excluye los hashes de contraseña por defecto. Solo deben incluirse si existe una necesidad técnica clara.

## 20. DCD 1.1.1.1 - Corrección de auditoría

Esta versión corrige la auditoría cuando Supabase tiene Row Level Security activado en la tabla `dcd_auditoria`.

Cambios:

- Añadidas políticas RLS para `dcd_auditoria`.
- Añadido botón de prueba manual de auditoría en `Panel administrador > Auditoría/Backup`.
- Si falla la auditoría, el panel muestra el último error detectado.

Es recomendable ejecutar de nuevo `supabase/schema.sql` en Supabase para dejar las políticas alineadas.

## 21. DCD 1.1.2 - Calidad de datos y comparativa

Esta versión añade una capa de revisión administrativa de calidad de datos sin modificar el flujo de carga de los centros docentes.

Incluye:

- Avisos básicos antes de finalizar un expediente: total 0, titulaciones duplicadas y valores altos.
- Nueva pestaña `Calidad de datos` en el panel administrador.
- Controles de centros pendientes, centros con borrador posterior, finalizados sin registros, duplicados y valores altos.
- Listado informativo de titulaciones del catálogo sin plazas.
- Comparativa entre las dos últimas publicaciones oficiales.
- Nuevas hojas de calidad en el Excel consolidado.
- Bloque de validaciones de calidad incorporado al PDF de publicación.

No requiere SQL nuevo respecto a la DCD 1.1.1.1.


## 22. DCD 1.1.2.1 - Pulido de interfaz

Esta versión incorpora ajustes de presentación y lenguaje de usuario:

- Se elimina el historial de versiones de la pantalla de acceso.
- El historial de versiones pasa a una pestaña propia del panel administrador.
- En la barra lateral se sustituyen referencias técnicas por textos transparentes para el usuario:
  - Base de Datos configurada.
  - Servidor correo configurado.
- Se oculta a usuarios no administradores la revisión del mapeo centro docente/columna Excel.
- El botón de finalización pasa a llamarse `Finalizar expediente en Base de Datos`.
- El mensaje de finalización al centro docente se simplifica para no mostrar información de centros pendientes.
- Se incorpora un intento de desplazamiento automático a la parte superior al entrar en el Paso 5.

No requiere SQL nuevo respecto a DCD 1.1.2.


## 23. DCD 1.1.2.3 - Ajuste final PDF

Versión de pulido del informe PDF publicado.

Incluye:

- Incorporación automática de `assets/logo.png` si existe.
- Frase institucional actualizada: “Informe desarrollado para la gestión y análisis de los Datos de Capacidad Docente del Servicio Canario de la Salud”.
- Pie de página con numeración.
- Pie de firma: “Jefatura del Servicio de Formacion Sanitaria Especializada”.

No requiere SQL nuevo respecto a DCD 1.1.2.1.


## 24. DCD 1.1.3.1 - Corrección GAP TF

Esta versión corrige la definición del centro docente de Atención Primaria de Tenerife.

- Se elimina la separación errónea entre GAP TF Norte y GAP TF Sur.
- Se mantiene un único centro docente: GERENCIA DE ATENCIÓN PRIMARIA DE TENERIFE.
- El centro queda vinculado a la columna oficial GAP TF de la Matriz_DCD.
- Se incorpora migración SQL para normalizar usuarios, borradores, registros y auditoría históricos que pudieran contener los nombres antiguos.

Esta versión requiere ejecutar el `supabase/schema.sql` actualizado.


## DCD 1.1.3.4 — Ajustes finales de visualización

- Ajuste estético del mensaje de confirmación de correo automático: “Correo enviado correctamente a servicio de FSE”.
- Limpieza visual del dashboard y del PDF publicado: las filas de Matriz_DCD con `Total = 0` se ocultan en la presentación para no ocupar espacio ni afear la lectura.
- El filtro es solo de visualización: no borra registros, no modifica Supabase, no altera borradores ni cambia la matriz oficial almacenada.


## DCD 1.1.3.4

- Mapa visual de Canarias en el dashboard del rol consulta con totales de capacidad docente por isla.
- Cambio exclusivamente visual, sin modificación de base de datos ni cálculos base.


### DCD 1.1.3.5
Corrección visual y de presentación para rol consulta: PDF de consulta sin calidad interna, sin filas con total 0 y mapa territorial ajustado.


### DCD 1.2.1 beta 2

Se inicia la fase multiusuario por centro docente. La primera beta añade una pestaña administrativa para configurar centros multiusuario y asignar usuarios existentes.


## DCD 1.2.1 beta 3

- Corrección del registro de publicaciones en Supabase cuando existen valores NaN/NaT procedentes de tablas Pandas.
- No requiere cambios SQL.
- No modifica lógica de cálculo, Excel ni PDF.
