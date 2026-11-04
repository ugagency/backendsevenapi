-- Seven OS, etapa intermediaria: credenciais do Coupa no banco, criptografadas.
-- Aditiva e idempotente: cria SO a tabela nova. Nao altera eventos_coletados nem o robo atual.
--
-- Seguranca:
--  * A senha e gravada cifrada pelo backend (Fernet, chave CREDENCIAIS_CHAVE so no servidor).
--    O banco nunca ve a senha em texto e o front nunca recebe a senha de volta.
--  * RLS ligado e SEM politicas: anon e authenticated (o front) nao leem nem escrevem.
--    So a chave secreta do backend (service_role / sb_secret_...) acessa, pois ela ignora RLS.

create table if not exists public.credenciais_coupa (
  conta_id       text primary key
                 check (conta_id ~ '^[a-z0-9_-]{1,30}$' and conta_id <> 'todas'),
  nome           text,
  usuario        text check (usuario is null or char_length(usuario) <= 200),
  senha_cifrada  text,
  atualizado_em  timestamptz not null default now(),
  atualizado_por text
);

comment on table public.credenciais_coupa is
  'Credenciais Coupa por filial. senha_cifrada = Fernet, chave so no backend. Acesso apenas via service_role.';

alter table public.credenciais_coupa enable row level security;
revoke all on table public.credenciais_coupa from anon, authenticated;

-- Conferencia (nao deve listar nenhuma politica, e a senha deve aparecer cifrada, comecando por gAAAA):
-- select policyname from pg_policies where tablename = 'credenciais_coupa';
-- select conta_id, nome, usuario, left(senha_cifrada, 5) as inicio, atualizado_em, atualizado_por from public.credenciais_coupa;
