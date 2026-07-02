-- ============================================================
-- DCD 1.2.2 beta 1
-- Servicios y curso/año por alumno
-- ============================================================

alter table public.dcd_borradores
add column if not exists detalle_alumnos jsonb not null default '[]'::jsonb;

alter table public.dcd_registros
add column if not exists detalle_alumnos jsonb not null default '[]'::jsonb;

create index if not exists idx_dcd_borradores_detalle_alumnos
on public.dcd_borradores using gin (detalle_alumnos);

create index if not exists idx_dcd_registros_detalle_alumnos
on public.dcd_registros using gin (detalle_alumnos);