-- Seven OS, etapa 0 - PARTE B: remove a unicidade antiga (numero_evento, titulo).
-- APLICAR SOMENTE DEPOIS que o backend novo estiver no ar (usa on_conflict "conta,numero_evento,titulo").
--
-- Por que existe: o mesmo evento da Vale pode chegar as duas filiais. Com a regra antiga,
-- a segunda filial nao consegue gravar o evento. Se esta parte for aplicada ANTES do deploy,
-- o robo antigo passa a falhar ao gravar ("no unique constraint matching ON CONFLICT").
--
-- Rollback (so se nao houver linhas de duas contas com o mesmo evento e titulo):
--   alter table public.eventos_coletados add constraint eventos_coletados_numero_evento_titulo_key unique (numero_evento, titulo);

begin;

alter table public.eventos_coletados
  drop constraint if exists eventos_coletados_numero_evento_titulo_key;

commit;

-- Conferencia:
-- select conname, pg_get_constraintdef(oid) from pg_constraint where conrelid = 'public.eventos_coletados'::regclass;
