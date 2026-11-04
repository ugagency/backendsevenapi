-- Seven OS: perfil de cada usuario do painel, com o papel (admin, operador).
-- Base de permissao para tudo que vier depois: o backend consulta o papel aqui,
-- e politicas RLS futuras podem usar public.eh_admin().
-- Aditiva e idempotente.
--
-- Papeis:
--   admin    - altera credenciais do Coupa e, no futuro, gerencia usuarios e configuracoes
--   operador - usa o painel (roda o robo, ve o historico). Padrao de todo usuario novo.
-- Ninguem muda o proprio papel pelo front: nao existe politica de UPDATE para usuarios.
-- Promover alguem e feito pelo SQL Editor (ou por uma tela de admin, no futuro, via backend).

create table if not exists public.perfis (
  user_id       uuid primary key references auth.users (id) on delete cascade,
  email         text,
  nome          text,
  papel         text not null default 'operador' check (papel in ('admin', 'operador')),
  ativo         boolean not null default true,
  criado_em     timestamptz not null default now(),
  atualizado_em timestamptz not null default now()
);

comment on table public.perfis is 'Perfil e papel de cada usuario do painel Seven (admin, operador).';

-- Todo usuario novo do Auth ganha um perfil de operador.
create or replace function public.criar_perfil_novo_usuario()
returns trigger
language plpgsql
security definer
set search_path = ''
as $$
begin
  insert into public.perfis (user_id, email) values (new.id, new.email)
  on conflict (user_id) do nothing;
  return new;
end;
$$;

drop trigger if exists ao_criar_usuario_criar_perfil on auth.users;
create trigger ao_criar_usuario_criar_perfil
  after insert on auth.users
  for each row execute function public.criar_perfil_novo_usuario();

-- Usuarios que ja existem
insert into public.perfis (user_id, email)
select id, email from auth.users
on conflict (user_id) do nothing;

-- Usado por politicas RLS: o usuario logado e admin ativo?
create or replace function public.eh_admin()
returns boolean
language sql
stable
security definer
set search_path = ''
as $$
  select exists (
    select 1 from public.perfis
    where user_id = auth.uid() and papel = 'admin' and ativo
  );
$$;

revoke all on function public.eh_admin() from public;
grant execute on function public.eh_admin() to authenticated;
revoke all on function public.criar_perfil_novo_usuario() from public, anon, authenticated;

alter table public.perfis enable row level security;
revoke all on table public.perfis from anon;
revoke insert, update, delete, truncate on table public.perfis from authenticated;
grant select on table public.perfis to authenticated;

drop policy if exists "perfil: cada um ve o seu, admin ve todos" on public.perfis;
create policy "perfil: cada um ve o seu, admin ve todos" on public.perfis
  for select to authenticated
  using (user_id = (select auth.uid()) or public.eh_admin());

-- Administradores iniciais (definidos em 03/11/2026)
update public.perfis set papel = 'admin', atualizado_em = now()
where lower(email) in ('contactugagency@gmail.com', 'emanuele@sevensuprimentos.com.br');

-- Conferencia:
-- select email, papel, ativo from public.perfis order by papel, email;
