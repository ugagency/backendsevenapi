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
