-- Seven OS, Spec 01. Aplicada em 20/11/2026 (migration os_spec01_acesso_perfis) em os e os_homolog.
-- Usuarios e perfis, contas/portais, contas liberadas por usuario, parametros. So objetos novos.
do $$
declare s text;
begin
  foreach s in array array['os','os_homolog'] loop
    execute format($f$
      create table if not exists %1$I.usuarios (
        user_id uuid primary key references auth.users(id) on delete cascade,
        nome text,
        email text,
        perfil text not null check (perfil in ('administrador','supervisor','cotacao','suporte')),
        ativo boolean not null default false,
        criado_em timestamptz not null default now(),
        convidado_por uuid
      )$f$, s);
    execute format($f$
      create table if not exists %1$I.contas_portal (
        id bigint generated always as identity primary key,
        portal text not null default 'coupa',
        conta_robo text,
        nome text not null,
        cnpj text,
        filial text,
        ativa boolean not null default true,
        criada_em timestamptz not null default now(),
        unique (portal, conta_robo)
      )$f$, s);
    execute format($f$
      create table if not exists %1$I.usuario_contas (
        user_id uuid not null references %1$I.usuarios(user_id) on delete cascade,
        conta_id bigint not null references %1$I.contas_portal(id) on delete cascade,
        primary key (user_id, conta_id)
      )$f$, s);
    execute format($f$
      create table if not exists %1$I.parametros (
        chave text primary key,
        valor jsonb not null,
        alterado_por uuid,
        alterado_em timestamptz not null default now()
      )$f$, s);
    execute format('revoke all on %I.usuarios, %I.contas_portal, %I.usuario_contas, %I.parametros from public, anon, authenticated', s, s, s, s);
    execute format('grant select, insert, update, delete on %I.usuarios, %I.contas_portal, %I.usuario_contas, %I.parametros to os_app', s, s, s, s);
    execute format('grant usage on sequence %I.contas_portal_id_seq to os_app', s);
    execute format($f$
      insert into %1$I.contas_portal (portal, conta_robo, nome)
      values ('coupa','filial1','filial1'), ('coupa','filial2','filial2')
      on conflict (portal, conta_robo) do nothing$f$, s);
  end loop;
end $$;
