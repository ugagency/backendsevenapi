# os_api

Servico novo do Seven OS. Separado do robo: container proprio, porta propria, deploy proprio.

## Variaveis de ambiente
| Variavel | Para que |
|---|---|
| `OS_DATABASE_URL` | Conexao do papel `os_app` (nunca o `postgres`) |
| `OS_SCHEMA` | `os_homolog` (padrao) ou `os` |
| `SUPABASE_URL`, `SUPABASE_KEY` | Validar login. So a chave publica |
| `OS_CORS_ORIGINS` | Origens do front, separadas por virgula. Vazio = nenhuma |
| `OS_CONVITE_REDIRECT` | Pagina do front onde o convidado cria a senha (`/definir-senha`). Tem que estar nas Redirect URLs do Supabase Auth |
| `OS_CONVITE_URL` | Opcional. Padrao: `SUPABASE_URL/functions/v1/os-convidar` |

O servico se recusa a subir se `SUPABASE_SECRET_KEY` estiver no ambiente.

## Rodar local
    pip install -r os_api/requirements.txt
    uvicorn os_api.main:app_factory --factory --port 8100

## Testes
    pip install pytest
    python -m pytest tests_os -q

## Compose (proposta; versionar o compose real quando os arquivos da VPS chegarem)
    os_api:
      build: { context: ., dockerfile: os_api/Dockerfile }
      env_file: .env.os
      ports: ["127.0.0.1:8100:8100"]
      restart: unless-stopped


## Convite de usuarios

Criar login exige a chave secreta do Supabase, e o os_api nao tem essa chave de proposito.
Quem cria o login e a Edge Function `supabase/functions/os-convidar`, que roda dentro do Supabase
(la a chave ja existe). O os_api repassa o token de quem esta logado; a funcao confere de novo que e
administrador ativo do Seven OS no schema informado antes de criar qualquer coisa.

- Modo `link` (padrao): a funcao devolve o link, o administrador envia por WhatsApp ou e-mail. Nao depende do e-mail do Supabase.
- Modo `email`: o Supabase envia. Sem SMTP proprio configurado, o Supabase so entrega para e-mails da equipe do projeto e poucos por hora.
- `POST /os/admin/usuarios/{id}/link-acesso`: link novo para convite expirado ou senha esquecida.
- O link nunca e gravado na auditoria nem no banco.

Publicar a funcao: `supabase functions deploy os-convidar --project-ref hwiyzclxtvggdhrxbsmc` (ou pelo MCP do Supabase).


## Hospedagem na Vercel (decisao de 24/11/2026)

O os_api roda na Vercel como funcao Python; o robo continua na VPS.
- Entrada: `api/index.py`; dependencias: `api/requirements.txt`; configuracao: `vercel.json` (regiao `cle1`, Ohio, a mesma do banco).
- `.vercelignore` deixa de fora o robo e todo o resto: so vao `api/` e `os_api/`.
- O pool do banco abre na primeira consulta (a Vercel pode nao rodar o lifespan). Usar `OS_POOL_MIN=0` e `OS_POOL_MAX=2`.
- Projeto Vercel separado (ex.: `seven-os-api`), importando este repositorio. Antes do merge, a Production Branch do
  projeto e `seven-os`; depois do merge, volta para `main`.
- Variaveis: `OS_DATABASE_URL`, `OS_SCHEMA`, `SUPABASE_URL`, `SUPABASE_KEY`, `OS_CORS_ORIGINS`, `OS_CORS_REGEX`,
  `OS_CONVITE_REDIRECT`, `OS_POOL_MIN`, `OS_POOL_MAX`. Nunca `SUPABASE_SECRET_KEY`.
- Limites conhecidos: primeiro pedido depois de um tempo parado e mais lento (a funcao "acorda"); caches em memoria
  valem so dentro de cada instancia.
