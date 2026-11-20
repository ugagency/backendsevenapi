# os_api

Servico novo do Seven OS. Separado do robo: container proprio, porta propria, deploy proprio.

## Variaveis de ambiente
| Variavel | Para que |
|---|---|
| `OS_DATABASE_URL` | Conexao do papel `os_app` (nunca o `postgres`) |
| `OS_SCHEMA` | `os_homolog` (padrao) ou `os` |
| `SUPABASE_URL`, `SUPABASE_KEY` | Validar login. So a chave publica |
| `OS_CORS_ORIGINS` | Origens do front, separadas por virgula. Vazio = nenhuma |

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
