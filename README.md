# DATOS CAPACIDAD DOCENTE (DCD 1.0.8.2)

Aplicativo Streamlit para recoger datos de capacidad docente a partir del Excel base `listado_para_capacidad_docente.xlsx`.

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

## 1. Estructura del proyecto

```text
dcd_capacidad_docente/
├── app.py
├── requirements.txt
├── README.md
├── .gitignore
├── .streamlit/
│   └── secrets.toml.example
├── data/
│   └── listado_para_capacidad_docente.xlsx
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
git commit -m "Actualizar DCD 1.0.8.2"
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

## 10. Publicaciones oficiales DCD 1.0.8.2

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


## 12. Cierre configurable DCD 1.0.8.2

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


## 13. DCD 1.0.8.2 - Refuerzo de cierre automático

Esta versión deja explícitos y reforzados estos disparadores de cierre:

- Comprobación automática cuando un centro finaliza expediente.
- Comprobación automática cuando cualquier usuario entra en la app.
- Botón de administrador **Evaluar ahora cierre automático**.
- Aviso claro en panel de cierre: Streamlit no ejecuta procesos permanentes en segundo plano; la fecha tope se evalúa cuando la app es utilizada o mediante evaluación manual del admin.

No requiere SQL nuevo si ya se ejecutó el `schema.sql` de la 1.0.8.1.
