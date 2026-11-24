-- Seven OS, assistente conversacional (agente no n8n com ferramentas de leitura).
-- Cada pergunta abre uma sessao curta: a API gera um codigo aleatorio, guarda so o hash, e o agente usa o codigo para
-- chamar a ferramenta de leitura em nome de quem perguntou. Expira em poucos minutos. So objetos novos, em os e os_homolog.
do $outer$
declare s text; ddl text := $ddl$
create table if not exists assistente_sessoes (
  token_hash text primary key, user_id uuid not null, conversa_id bigint references conversas(id) on delete cascade,
  criada_em timestamptz not null default now(), expira_em timestamptz not null,
  chamadas int not null default 0, ultima_consulta text, ultimos_parametros jsonb);
create index if not exists assistente_sessoes_expira_idx on assistente_sessoes (expira_em);
$ddl$;
begin
  foreach s in array array['os','os_homolog'] loop
    execute format('set local search_path = %I', s);
    execute ddl;
    execute format('revoke all on %I.assistente_sessoes from public, anon, authenticated', s);
    execute format('grant select, insert, update, delete on %I.assistente_sessoes to os_app', s);
  end loop;
end $outer$;
