# DATOS CAPACIDAD DOCENTE (DCD 1.0)

Aplicativo MVP en Python/Streamlit para recoger datos de capacidad docente a partir del Excel base `listado_para_capacidad_docente.xlsx`.

## 1. Qué incluye esta primera versión

- Pantalla de acceso con contraseña.
- Pantalla de instrucciones y aceptación.
- Selección de área y unidad docente.
- Confirmación de datos seleccionados.
- Entrada de datos mediante selectores dependientes:
  - Nivel Estudio I
  - Nivel Estudio II
  - Rama
  - Titulación
  - Número de alumnos
- Tabla de registros introducidos.
- Eliminación de registros.
- Descarga de Excel final.
- Preparación para guardar/cargar borradores en Supabase.

## 2. Estructura del proyecto

```text
DCD_CAPACIDAD_DOCENTE/
├── app.py
├── requirements.txt
├── .gitignore
├── README.md
├── data/
│   └── listado_para_capacidad_docente.xlsx
├── supabase/
│   └── schema.sql
└── .streamlit/
    └── secrets.toml.example
```

## 3. Probar en local

Abre CMD o PowerShell en la carpeta del proyecto.

Crear entorno virtual:

```bash
python -m venv .venv
```

Activar entorno virtual en Windows:

```bash
.venv\Scripts\activate
```

Instalar dependencias:

```bash
pip install -r requirements.txt
```

Crear archivo de secretos local:

```bash
copy .streamlit\secrets.toml.example .streamlit\secrets.toml
```

Ejecutar la aplicación:

```bash
streamlit run app.py
```

Contraseña inicial:

```text
Capacidad2026
```

## 4. Configurar Supabase

1. Entrar en https://supabase.com
2. Crear cuenta o iniciar sesión.
3. Crear nuevo proyecto.
4. Nombre sugerido: `dcd-capacidad-docente`.
5. Guardar bien la contraseña de base de datos.
6. Elegir región europea si está disponible.
7. Esperar a que el proyecto se cree.

Después:

1. Ir a `SQL Editor`.
2. Pulsar `New query`.
3. Copiar el contenido de `supabase/schema.sql`.
4. Ejecutar la consulta.

## 5. Obtener claves de Supabase

Dentro del proyecto de Supabase:

1. Ir a `Project Settings`.
2. Entrar en `API`.
3. Copiar:
   - `Project URL`
   - una clave API para el servidor.

Para esta primera prueba se puede usar la clave `service_role` solo en los secretos de Streamlit, nunca en GitHub.

## 6. Configurar secretos locales

Editar `.streamlit/secrets.toml`:

```toml
APP_PASSWORD = "Capacidad2026"
SUPABASE_URL = "https://xxxx.supabase.co"
SUPABASE_KEY = "xxxxx"
```

Importante: `.streamlit/secrets.toml` no se sube a GitHub porque está en `.gitignore`.

## 7. Subir a GitHub

Crear un repositorio nuevo en GitHub, por ejemplo:

```text
dcd-capacidad-docente
```

Desde la carpeta del proyecto:

```bash
git init
git add .
git commit -m "Primera version DCD 1.0"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/dcd-capacidad-docente.git
git push -u origin main
```

Antes de subir, comprobar que no existe este archivo en el commit:

```text
.streamlit/secrets.toml
```

## 8. Desplegar en Streamlit Community Cloud

1. Entrar en https://streamlit.io/cloud
2. Crear nueva app.
3. Elegir el repositorio de GitHub.
4. Rama: `main`.
5. Archivo principal: `app.py`.
6. Desplegar.

Luego ir a los ajustes de la app en Streamlit Cloud y añadir los secretos:

```toml
APP_PASSWORD = "Capacidad2026"
SUPABASE_URL = "https://xxxx.supabase.co"
SUPABASE_KEY = "xxxxx"
```

## 9. Notas de diseño

La app no usa una tabla gigante editable porque Streamlit puede hacer reruns y provocar pérdida aparente de datos en formularios complejos. En su lugar, usa un patrón más estable:

```text
Seleccionar combinación -> introducir número de alumnos -> añadir/actualizar registro
```

Los datos introducidos se guardan en `st.session_state` durante la sesión. Si Supabase está configurado, además se pueden guardar como borrador y recuperar en otra sesión.

## 10. Pendiente para siguientes versiones

- Pantalla específica de recordatorio.
- Envío de correo automático.
- Control avanzado de usuarios.
- Auditoría de cambios.
- Revisión definitiva de mapeo entre unidades docentes y columnas del Excel.
