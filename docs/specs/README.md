# Seven OS: specs de construção

Rascunho de 05/11/2026. Fonte: documento "Seven OS: escopo do produto", plano de evolução, proposta V3.5, termo de confidencialidade e decisões da SSYS de 05/11.

Cada spec é uma entrega. Termina em merge na main com o conteúdo novo invisível para a Karen (rotas `/os`, schema `os`). A ordem abaixo é a ordem de construção.

| Spec | Entrega | Etapa da proposta | Depende de |
|---|---|---|---|
| 00 | Fundação e segurança | Antes da Etapa 1 (fora do cronograma) | nada |
| 01 | Acesso, perfis e administração | Etapa 1 | 00 |
| 02 | Captura | Etapa 1 | 00, 01 |
| 03 | Identificação do produto e base histórica | Etapa 1 | 02 |
| 04 | Fila e ficha (primeiro somente leitura) | Leitura logo após a Etapa 1; completa na Etapa 4 | 03 |
| 05 | Fornecedores | Etapa 2 | 04 |
| 06 | Cotação | Etapa 3 | 05 |
| 07 | Painel de indicadores | Etapa 5 | 06 |
| 08 | Assistente | Etapa 5 | 06 |
| 09 | Exportação completa, corte e treinamento | Etapa 5 | todas |

A fila (04) vem antes de fornecedores e cotação, em versão somente leitura, para a Karen validar a identificação com casos reais cedo. A versão completa da fila (filtros, exportação, ordenação por valor) fecha na Etapa 4, como na proposta.

## Formato de cada spec

Objetivo, dentro e fora, como funciona, dados, API, telas, permissões, regras, aceite, testes e pendências.

## Regras que valem em todas as specs

1. Preço de compra nunca sai em mensagem a fornecedor nem vai a terceiro, inclusive ao modelo de IA.
2. Preço, margem e venda só aparecem para perfil autorizado, em tela, exportação, painel e assistente.
3. Toda alteração de preço, margem ou fornecedor tem autor e data, na auditoria.
4. O sistema não força identificação quando a descrição é insuficiente.
5. Sugestão de fornecedor exige confirmação humana e mostra a fonte.
6. O assistente responde só com dados registrados e não grava nada sem confirmação.
7. Dados da Seven não treinam modelos e não são usados fora dos Serviços.
8. O robô atual e a tela da Karen não param durante a construção.

## Contrato de produção (não quebrar)

- Rotas do serviço atual: `GET /contas`, `GET /status`, `POST /run-robot`, `GET /eu`, `GET/PUT /credenciais`.
- Storage: bucket `planilhas`, nome `eventos_[conta_]DDMMAA_timestamp.xlsx`.
- Planilha: 9 colunas, mais "Filiais" na unificada.
- `public.eventos_coletados` com upsert em (conta, numero_evento, titulo). O OS só lê.
- Login atual e a tela `/`.

## Portão de cada merge na main

- Planilha de uma data real idêntica à de produção.
- Contrato acima conferido.
- Migrations aplicadas antes em `os_homolog`, com dump feito.
- Publicação fora do horário da Karen e com `/status` sem execução em andamento.
- Volta escrita antes do merge.

## Decisões de 05/11 que as specs já seguem

- Preço de venda: a equipe digita; o sistema não calcula.
- Evento em dois CNPJs: uma linha na fila, com as filiais indicadas.
- Perfis: Administrador, Supervisor e Cotação. Suporte SSYS com acesso próprio, sem preço, auditado.
- Cotação vê preço de compra e não vê margem.
- Administradores da Seven: Emanuele e Fred, com e-mails definidos no onboarding presencial.
- Cada usuário vê só as contas e portais liberados para ele. Novas contas e portais são criados pela SSYS.
- Sistema exclusivo da Seven, sem camada multiempresa.
- Declinar é descarte interno, guardado com motivo e recorrência.
- Horários e rodadas do robô configuráveis pelo Administrador.
- WhatsApp, ligar e registrar retorno: depois da Etapa 5.
- Ficam no corte ou no fechamento (Spec 09): exigir login em `/run-robot` e `/status`, tirar "Criar Conta Nova", repositório privado e revogar `eh_admin()` para anônimo. Na Fase 0 sai a rota `/api/start-n8n` do front.
- Fechamento de segurança (trancar `eventos_coletados`, restringir o bucket, apagar `user_settings`) fica por último. `credenciais_coupa` fica e pode ser alterada.
- Login pelo Supabase Auth. Senhas dos portais cifradas no Supabase; `.env` como plano B temporário só para as duas contas atuais.
