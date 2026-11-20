# Seven OS: como testar no seu computador

Nada aqui toca no robo nem na tela atual. O OS le `public.eventos_coletados` (so leitura) e grava no schema `os_homolog`.

## Uma vez so
1. Banco: aplicar `migrations/os/005_funcoes_captura_identificacao.sql` no Supabase (as 001 a 004 ja foram aplicadas) e depois rodar,
   no SQL Editor: `select os_homolog.sincronizar();` e `analyze;`. A primeira carga leva menos de um minuto.
2. Senha do papel do OS, no SQL Editor (escolha a senha; nao mande por chat): `alter role os_app with login password 'SUA_SENHA';`
3. Backend, na pasta do repositorio `backendsevenapi` (branch `seven-os`):
   - `pip install -r os_api/requirements.txt python-dotenv`
   - copiar `os_api/.env.os.example` para `.env.os` e preencher a senha e a chave publishable.
4. Front, na pasta `seven/frontend` (branch `seven-os`): `npm install` e, no `.env.local`, acrescentar
   `NEXT_PUBLIC_OS_API_URL=http://localhost:8100` (as variaveis do Supabase que ja existem continuam).

## Para rodar (dois terminais)
- Backend (cmd): `python -m dotenv -f .env.os run -- python -m uvicorn os_api.main:app_factory --factory --port 8100`
  Conferir: abrir http://localhost:8100/os/saude e ver `"banco": true`.
- Front: `npm run dev` e abrir http://localhost:3000/os com o login de sempre.

So entra no OS quem esta cadastrado e ativo em `os_homolog.usuarios` (hoje: contactugagency@gmail.com, perfil suporte).
Para ver precos e margem no teste, troque o perfil dele para administrador em Administracao, com outro usuario, ou peca a troca.

## Testes automaticos
- `python -m pytest tests_os -q` roda os testes sem banco.
- Com Postgres local: `tests_os/local/preparar.sh` e depois
  `OS_TEST_DATABASE_URL="postgresql://os_app@/seven_teste?host=/tmp&port=5544" python -m pytest tests_os -q` (13 testes de ponta a ponta).

## O que ainda nao esta ligado (dito na tela tambem)
- O robo continua sendo rodado pela tela atual. O OS traz as capturas pelo botao "Trazer capturas novas".
- Linhas repetidas do mesmo codigo no mesmo evento continuam vindo uma so, porque o OS le a tabela atual do robo.
- E-mail a fornecedor abre no programa de e-mail da pessoa; nao ha servico de envio contratado.
- O assistente interpreta a pergunta por regras, sem modelo de IA.
- Base historica da Seven (precos e fornecedores) nao foi importada: a tabela existe e esta vazia.
- Convite por e-mail e agenda automatica do robo: no corte.
