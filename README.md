# DATOS CAPACIDAD DOCENTE (DCD 1.0.4)

Aplicativo Streamlit para recoger datos de capacidad docente a partir del Excel base `listado_para_capacidad_docente.xlsx`.

Esta versión parte de la DCD 1.0 y añade:

- Pantalla específica de recordatorio antes de finalizar.
- Envío automático opcional por Mailgun.
- Control básico/avanzado de usuarios mediante `USERS_JSON`.
- Auditoría de acciones en Supabase.
- Revisión visible del mapeo entre unidades docentes y columnas del Excel.
- Ajuste del campo Área: `HOSPITAL`, `ATENCION FAMILIAR Y COMUNITARIA` y retirada de `OTRAS UNIDADES DOCENTES`.
- Exportación Excel más completa, con estado, usuario, observaciones y datos generales en las hojas de salida.
- Control de estado `borrador` / `finalizado` con bloqueo suave de edición para expedientes finalizados.
- Totales automáticos en la hoja `Matriz_DCD`, incluida la fila total final.
- Panel administrador para generar un Excel consolidado desde los expedientes finalizados guardados en Supabase.

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

De momento el control de usuarios sirve para identificar quién entra y para auditoría. En una versión posterior se puede limitar cada usuario a una unidad docente concreta.

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

Si ya habías creado las tablas de DCD 1.0, este SQL añade la tabla nueva de auditoría sin borrar los datos existentes.

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
git commit -m "Actualizar DCD 1.0.4"
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
3. Selección de área y unidad docente.
4. Confirmación y revisión de mapeo Excel.
5. Entrada de datos por selectores dependientes.
6. Guardado de borrador.
7. Pantalla de recordatorio.
8. Descarga de Excel.
9. Guardado como finalizado.
10. Envío automático opcional.
