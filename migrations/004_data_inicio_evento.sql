-- Seven OS: guarda a data de INICIO do evento (o dia em que ele abriu no Coupa).
-- Aditiva e idempotente; coluna anulavel, entao o robo atual continua funcionando sem ela.
-- Aplicar ANTES do deploy do backend novo (junto com a 001a), pois o backend novo grava esta coluna.
--
-- Hoje, data_evento guarda o PRAZO do evento (coluna 'DATA' da planilha), nao o dia de inicio.
-- O nome foi mantido para nao quebrar o robo em producao. Linhas antigas ficam com data_inicio nula:
-- data_coleta e o dia da captura, que so coincide com o inicio quando a captura e feita no proprio dia.

alter table public.eventos_coletados
  add column if not exists data_inicio date;

comment on column public.eventos_coletados.data_inicio is 'Dia em que o evento abriu no Coupa (data de inicio da lista de eventos).';
comment on column public.eventos_coletados.data_evento is 'ATENCAO: guarda o PRAZO/data final do evento (texto dd/mm/aa), nao o inicio. Ver data_inicio.';

create index if not exists eventos_coletados_data_inicio_idx
  on public.eventos_coletados (data_inicio);

-- Conferencia depois de uma captura real:
-- select data_inicio, count(*) from public.eventos_coletados group by 1 order by 1 desc limit 10;
