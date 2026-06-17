# RELEASE NOTES — DCD 1.1.3.4

## Versión

**DCD 1.1.3.4 — Mapa visual de capacidad docente por isla**

## Autoría

- **Desarrollador / creador del programa:** Alberto Cabrera
- **Responsable funcional del proyecto:** Alberto Cabrera

## Cambios principales

- Incorporación de una visualización territorial en el dashboard del rol consulta.
- El nuevo bloque muestra un mapa de Canarias con la capacidad docente total por isla.
- Se añade la imagen `assets/mapa_canarias.png` como recurso visual local de la aplicación.
- La visualización usa los datos ya calculados en `Resumen_Isla`; no introduce cálculos nuevos ni altera resultados.
- Se mantiene la limpieza visual incorporada en DCD 1.1.3.3 para no mostrar filas con valores 0 en dashboard/PDF.

## Alcance técnico

El cambio es exclusivamente visual y limitado al dashboard del rol consulta. No modifica:

- Supabase.
- Registros.
- Borradores.
- Usuarios.
- Publicaciones.
- Excel generado.
- PDF generado.
- Cálculos base de consolidación.

## Base de datos

No requiere SQL nuevo respecto a DCD 1.1.3.3.

## Despliegue

Procedimiento habitual:

```bash
git add .
git commit -m "Actualizar DCD 1.1.3.4 mapa dashboard consulta"
git push
```

Después, reiniciar la app en Streamlit Cloud.

## Prueba mínima

- Entrar con rol consulta.
- Abrir la publicación vigente.
- Confirmar que aparece el bloque “Visualización territorial”.
- Confirmar que el mapa muestra los totales por isla.
- Preparar Excel y PDF para comprobar que las descargas siguen funcionando.


## DCD 1.1.3.5
Versión correctiva de presentación para rol consulta: PDF limpio sin datos internos de calidad, ocultación de filas total 0 y mapa territorial ajustado.
