-- ============================================================
-- Esquema Supabase para la webapp de Horarios IES Virgen de la Calle
-- Ejecutar en el SQL Editor de un proyecto Supabase NUEVO
-- ============================================================

-- 1) AUSENCIAS (registro diario de quién falta / guardias a cubrir)
create table ausencias (
  id         bigint generated always as identity primary key,
  fecha      date not null,
  codigo     text not null,        -- código de docente (hoja "Docentes")
  nombre     text not null,
  sesiones   text[] not null,      -- ej. ['1','2','3R']
  created_at timestamptz default now()
);

-- 2) RESERVAS (reserva puntual de un aula en una sesión/fecha)
create table reservas (
  id         bigint generated always as identity primary key,
  fecha      date not null,
  aula       text not null,
  sesion     text not null,
  docente    text not null,
  motivo     text,
  created_at timestamptz default now()
);

-- 3) SUSTITUCIONES (sustituciones de docente, con periodo de vigencia)
create table sustituciones (
  id                 bigint generated always as identity primary key,
  docente_original   text not null,
  codigo_original    text not null,
  docente_sustituto  text not null,
  fecha_inicio       date not null,
  fecha_fin          date,             -- null = sigue activa indefinidamente
  activa             boolean default true,
  created_at         timestamptz default now()
);

-- 4) ACEX (Actividades Complementarias y Extraescolares)
create table acex (
  id            bigint generated always as identity primary key,
  fecha         date not null,
  sesiones      text[] not null,
  docentes      text[],           -- nullable
  grupos        text[],           -- nullable
  observaciones text,             -- nullable
  created_at    timestamptz default now()
);

-- ============================================================
-- RLS (Row Level Security): la webapp usa la clave "anon" directamente
-- desde el navegador, así que hay que habilitar políticas explícitas
-- o el frontend no podrá leer/escribir nada.
-- ============================================================

alter table ausencias enable row level security;
alter table reservas enable row level security;
alter table sustituciones enable row level security;
alter table acex enable row level security;

-- Política simple para pruebas: acceso total con la clave anon.
-- ⚠️ Esto es deliberadamente permisivo, pensado para un entorno de
-- prueba aislado. No usar así en producción con datos reales sin
-- añadir autenticación.
create policy "anon_all_ausencias" on ausencias for all using (true) with check (true);
create policy "anon_all_reservas" on reservas for all using (true) with check (true);
create policy "anon_all_sustituciones" on sustituciones for all using (true) with check (true);
create policy "anon_all_acex" on acex for all using (true) with check (true);
