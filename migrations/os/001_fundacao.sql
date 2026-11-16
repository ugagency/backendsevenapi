-- Seven OS, Fase 0. Aplicada em 16/11/2026 (migrations os_fase0_schemas_papel_auditoria e os_fase0_politicas_auditoria).
-- Cria so objetos novos: papel os_app (sem login nem senha ate definirem no painel), schemas os e os_homolog,
-- os.auditoria / os_homolog.auditoria (so INSERT e SELECT para os_app) e SELECT em public.eventos_coletados.
-- Testado: os_app le eventos_coletados; nao escreve nem apaga; nao le credenciais_coupa nem user_settings;
-- nao altera nem apaga auditoria; nao cria tabela em public; cria tabela em os_homolog.

do $$ begin
  if not exists (select 1 from pg_roles where rolname = 'os_app') then
    create role os_app nologin;
  end if;
end $$;

do $$
declare s text;
begin
  foreach s in array array['os','os_homolog'] loop
    execute format('create schema if not exists %I', s);
    execute format('revoke all on schema %I from public, anon, authenticated', s);
    execute format('grant usage, create on schema %I to os_app', s);
    execute format($f$
      create table if not exists %I.auditoria (
        id bigint generated always as identity primary key,
        em timestamptz not null default now(),
        user_id uuid,
        origem text not null check (origem in ('tela','exportacao','assistente','robo','carga','suporte')),
        entidade text not null,
        entidade_id text,
        acao text not null check (acao in ('criar','alterar','excluir','exportar','ver_restrito')),
        antes jsonb,
        depois jsonb,
        justificativa text
      )$f$, s);
    execute format('create index if not exists auditoria_em_idx on %I.auditoria (em)', s);
    execute format('create index if not exists auditoria_entidade_idx on %I.auditoria (entidade, entidade_id)', s);
    execute format('revoke all on %I.auditoria from public, anon, authenticated', s);
    execute format('grant select, insert on %I.auditoria to os_app', s);
    execute format('grant usage on sequence %I.auditoria_id_seq to os_app', s);
    execute format('alter table %I.auditoria enable row level security', s);
    execute format('create policy os_app_select on %I.auditoria for select to os_app using (true)', s);
    execute format('create policy os_app_insert on %I.auditoria for insert to os_app with check (true)', s);
  end loop;
end $$;

grant usage on schema public to os_app;
grant select on public.eventos_coletados to os_app;
grant os_app to postgres with set true;
