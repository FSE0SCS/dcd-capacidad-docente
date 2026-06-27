-- DCD 1.2.0 - Turnos y observaciones por titulación
alter table public.dcd_borradores
add column if not exists alumnos_manana integer default 0,
add column if not exists alumnos_tarde integer default 0,
add column if not exists alumnos_rotatorio integer default 0,
add column if not exists alumnos_deslizante integer default 0,
add column if not exists deslizante_lunes text,
add column if not exists deslizante_martes text,
add column if not exists deslizante_miercoles text,
add column if not exists deslizante_jueves text,
add column if not exists deslizante_viernes text,
add column if not exists observaciones_titulacion text;

alter table public.dcd_registros
add column if not exists alumnos_manana integer default 0,
add column if not exists alumnos_tarde integer default 0,
add column if not exists alumnos_rotatorio integer default 0,
add column if not exists alumnos_deslizante integer default 0,
add column if not exists deslizante_lunes text,
add column if not exists deslizante_martes text,
add column if not exists deslizante_miercoles text,
add column if not exists deslizante_jueves text,
add column if not exists deslizante_viernes text,
add column if not exists observaciones_titulacion text;
