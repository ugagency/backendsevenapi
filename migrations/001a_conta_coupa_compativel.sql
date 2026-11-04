-- Seven OS, etapa 0 - PARTE A: adiciona a conta Coupa SEM quebrar o robo que roda hoje.
-- Pode ser aplicada antes do deploy do backend novo. Idempotente.
--
-- O robo antigo continua funcionando: ele nao informa "conta" (entra o padrao 'filial1')
-- e seu upsert usa (numero_evento, titulo), que continua existindo.

begin;

alter table public.eventos_coletados
  add column if not exists conta text not null default 'filial1';

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

-- Conferencia:
-- select conta, count(*) from public.eventos_coletados group by conta;
