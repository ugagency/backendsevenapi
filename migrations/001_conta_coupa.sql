-- Seven OS: suporte a mais de uma conta Coupa (um CNPJ/filial por conta).
-- Aplicar ANTES de subir o main_api.py que grava a coluna "conta".
-- Idempotente: pode rodar mais de uma vez.

begin;

-- 1. Coluna da conta. Linhas existentes pertencem à conta original ("filial1",
--    que é o VALE_CONTA_ID padrão da conta configurada em VALE_USER / VALE_PASS).
--    Se a conta 1 tiver VALE_CONTA_ID diferente no .env, troque o default abaixo.
alter table public.eventos_coletados
  add column if not exists conta text not null default 'filial1';

-- 2. Remove a unicidade antiga (numero_evento, titulo). O mesmo evento da Vale
--    pode chegar às duas filiais e precisa ser gravado para cada uma.
do $$
declare r record;
begin
  for r in
    select conname
    from pg_constraint
    where conrelid = 'public.eventos_coletados'::regclass
      and contype = 'u'
      and conname <> 'eventos_coletados_conta_evento_titulo_key'
  loop
    execute format('alter table public.eventos_coletados drop constraint %I', r.conname);
  end loop;
end $$;

-- 3. Nova unicidade por conta (usada pelo upsert on_conflict="conta,numero_evento,titulo").
do $$
begin
  if not exists (
    select 1 from pg_constraint
    where conrelid = 'public.eventos_coletados'::regclass
      and conname = 'eventos_coletados_conta_evento_titulo_key'
  ) then
    alter table public.eventos_coletados
      add constraint eventos_coletados_conta_evento_titulo_key
      unique (conta, numero_evento, titulo);
  end if;
end $$;

commit;

-- Conferência depois de aplicar:
-- select conta, count(*) from public.eventos_coletados group by conta;
-- select conname, pg_get_constraintdef(oid) from pg_constraint
--   where conrelid = 'public.eventos_coletados'::regclass;
