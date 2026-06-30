-- ============================================================
-- DCD 1.2.1 beta 1
-- Soporte multiusuario por centro docente
-- ============================================================

create table if not exists public.dcd_centros_multiusuario_config (
    id bigserial primary key,
    centro_docente text not null unique,
    area text,
    multiusuario_activo boolean not null default false,
    usuarios_previstos integer not null default 1 check (usuarios_previstos >= 1),
    permite_consolidacion_parcial boolean not null default true,
    observaciones_config text,
    actualizado_por text,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now()
);

create table if not exists public.dcd_centros_multiusuario_usuarios (
    id bigserial primary key,
    centro_docente text not null,
    username text not null,
    email text,
    activo boolean not null default true,
    orden_participante integer,
    puede_finalizar_aportacion boolean not null default true,
    puede_consolidar_centro boolean not null default false,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now(),
    unique (centro_docente, username)
);

create table if not exists public.dcd_centros_multiusuario_consolidados (
    id bigserial primary key,
    centro_docente text not null,
    codigo_consolidado text not null unique,
    estado text not null default 'borrador',
    total_usuarios_previstos integer not null default 1,
    total_usuarios_finalizados integer not null default 0,
    consolidacion_parcial boolean not null default false,
    usuarios_pendientes text,
    consolidado_por text,
    motivo_consolidacion text,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now()
);

alter table public.dcd_borradores
add column if not exists usuario_aportacion text,
add column if not exists centro_multiusuario boolean default false,
add column if not exists codigo_consolidado_centro text,
add column if not exists aportacion_finalizada boolean default false,
add column if not exists fecha_finalizacion_aportacion timestamptz;

alter table public.dcd_registros
add column if not exists usuario_aportacion text,
add column if not exists centro_multiusuario boolean default false,
add column if not exists codigo_consolidado_centro text,
add column if not exists aportacion_finalizada boolean default false,
add column if not exists fecha_finalizacion_aportacion timestamptz;

create index if not exists idx_dcd_multi_config_centro
on public.dcd_centros_multiusuario_config (centro_docente);

create index if not exists idx_dcd_multi_usuarios_centro
on public.dcd_centros_multiusuario_usuarios (centro_docente);

create index if not exists idx_dcd_multi_usuarios_username
on public.dcd_centros_multiusuario_usuarios (username);

create index if not exists idx_dcd_borradores_usuario_aportacion
on public.dcd_borradores (usuario_aportacion);

create index if not exists idx_dcd_registros_usuario_aportacion
on public.dcd_registros (usuario_aportacion);

create index if not exists idx_dcd_registros_codigo_consolidado
on public.dcd_registros (codigo_consolidado_centro);
