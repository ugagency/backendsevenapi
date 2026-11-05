# Spec 00: Fundação e segurança

- Etapa da proposta: nenhuma. Pré-requisito, fora do cronograma de 30 dias úteis.
- Repositórios: `backendsevenapi` e `seven`.
- Depende de: nada.
- Status: rascunho, 05/11/2026.

## Objetivo

Preparar o terreno para construir o OS no mesmo Supabase da produção sem risco para a Karen, e fechar as brechas de segurança que já existem hoje.

## Dentro

1. Backup do banco (dump) e cópia das planilhas do bucket `planilhas`. Conferir se o plano do Supabase tem backup automático.
2. Exportar a estrutura atual do banco para `migrations/000_base.sql`. A criação de `eventos_coletados` não está versionada hoje.
3. Branch `seven-os` nos dois repositórios.
4. Schemas `os` e `os_homolog`, com a mesma estrutura.
5. Papel Postgres `os_app`: acesso total a `os` e `os_homolog`, somente SELECT em `public.eventos_coletados`, nenhum acesso a `public.credenciais_coupa` nem à chave secreta.
6. Versionar `docker-compose.yml`, script de deploy e configuração do Nginx da VPS.
7. Esqueleto do serviço `os_api` (FastAPI) em container próprio, escutando só em 127.0.0.1, com rota própria no Nginx.
8. Tabela de auditoria (`os.auditoria`) e o mecanismo de gravação, já na fundação.
9. Fica para o corte (Spec 09): desligar o cadastro aberto, tirando "Criar Conta Nova" e criando perfil inativo por padrão.
10. Fica para o fechamento (Spec 09): repositório do backend privado, depois de confirmar como a VPS baixa o código.
11. Fica para o fechamento (Spec 09), depois de testar: corrigir os avisos do Supabase: `eh_admin()` executável por anônimo e proteção contra senha vazada desligada.
12. Manter `credenciais_coupa` (criada em 04/11, vazia) como a tabela de senhas do OS. Ela pode ser alterada para receber o portal e o que mais for preciso. A tabela antiga `user_settings` (2 linhas com e-mail e senha do Coupa em texto puro, que o código atual não usa) fica para o fechamento de segurança (Spec 09).
13. Remover do front a rota `/api/start-n8n` (aponta para um webhook externo e nada a chama) e o comentário sobre n8n em `page.tsx`. O n8n só volta no agente.

## Fora

- Restringir alteração e exclusão no bucket `planilhas`. Decisão de 05/11: fica para depois. Fica registrado como risco aberto.
- Trancar leitura e inserção públicas em `eventos_coletados` e fechar a inserção no bucket: ficam para o fim (Spec 09). O código já grava com a chave de servidor quando ela está configurada na VPS (não verificado lá).
- Remover o plano B do `.env` (Spec 01).

## Como funciona

- Migrations do OS só criam objetos em `os_homolog` e `os`. Aplicam primeiro em `os_homolog`, depois em `os`.
- Mudança em `public` é exceção: revisada, fora do horário da Karen e com volta escrita.
- Homologação: front em preview da branch `seven-os` com variáveis próprias; `os_api` de homologação aponta para `os_homolog`; robô em `ROBO_SIMULAR=1`, sem gravar em `eventos_coletados`.
- Carga pesada fora do horário da Karen, porque o banco é compartilhado com o robô.

## Dados

`os.auditoria`

| Coluna | Tipo | Nota |
|---|---|---|
| id | bigint identity | |
| em | timestamptz | default now() |
| user_id | uuid | null para ação do sistema |
| origem | text | tela, exportação, assistente, robô, carga, suporte |
| entidade | text | ex.: cotacao, fornecedor, produto, usuario |
| entidade_id | text | |
| acao | text | criar, alterar, excluir, exportar, ver_restrito |
| antes | jsonb | |
| depois | jsonb | |
| justificativa | text | obrigatória em alerta mantido |

Sem UPDATE nem DELETE para `os_app`. Só INSERT e SELECT.

## API

`os_api`, esqueleto:

- `GET /os/saude`: versão, conexão com o banco, schema em uso.
- Middleware de autenticação: valida o token do Supabase Auth em toda rota, exceto `/os/saude`.
- Middleware de auditoria: função única `auditar(...)` usada por todas as rotas que gravam.

## Telas

Nenhuma visível. Rota `/os` no front mostra "em construção" só para usuários com perfil do OS.

## Permissões

Nenhuma nova nesta spec. Perfis entram na Spec 01.

## Regras

- `os_api` nunca usa a chave secreta do Supabase.
- Nenhuma senha em código ou no repositório.

## Aceite

- Dump e cópia das planilhas guardados fora do bucket público, em local combinado.
- `000_base.sql` recria a estrutura atual em banco vazio.
- `os_app` falha ao tentar escrever em `eventos_coletados` e ao ler `credenciais_coupa`.
- `os_api` sobe e cai sem reiniciar o serviço atual.
- Planilha de uma data real igual à de produção depois do merge.
- Botão "Criar Conta Nova" fica para o corte.

## Testes

- Teste de permissão do papel `os_app` (escrita em `public` deve falhar).
- Deploy do `os_api` com o robô rodando: o robô termina normalmente.

## Pendências

- Não verificado: configuração atual da VPS (compose, deploy, Nginx) e do Vercel (se variáveis valem para Preview; se valerem, a branch cairia no backend de produção).
- Onde guardar o dump: o PC do usuário aguenta (menos de 40 MB), mas o dado é coberto pelo termo (item 1.2). Recomendado: armazenamento privado com acesso restrito.
- Bucket aberto para alteração e exclusão anônimas continua como risco até a decisão.
