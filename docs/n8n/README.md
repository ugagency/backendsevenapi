# Assistente do Seven OS no n8n (agente conversacional)

O assistente é um agente de IA no n8n (OpenRouter) que conversa com o usuário e lê os dados do Seven OS por uma ferramenta
da própria API. O agente não tem acesso ao banco nem escreve SQL.

## Como funciona
1. O usuário pergunta na tela. A API (`POST /os/assistente`) cria uma sessão curta (código aleatório, vale 4 minutos e no
   máximo 8 consultas) e chama o Webhook do n8n com: pergunta, últimas mensagens da conversa, data de hoje, primeiro nome do
   usuário, se ele vê valores, o código da sessão e o endereço da API.
2. O agente decide o que consultar e chama a ferramenta `consultar_seven`, que faz `POST {api}/os/assistente/ferramenta`
   com o cabeçalho `X-Assistente-Token`.
3. A API executa só a lista fechada de consultas (buscar_produto, historico_produto, fornecedores_produto, itens_por_prazo,
   marcas_sem_fornecedor, indicadores), com as contas e a visibilidade de preço e margem do perfil de quem perguntou.
4. O agente responde em texto. A API encerra a sessão, monta a tabela da última consulta para a tela e grava a conversa.
5. Se o n8n falhar ou demorar mais que `OS_N8N_TIMEOUT` (padrão 50 s), o assistente responde pelas regras internas.

Decisão do cliente (24/11/2026): o agente recebe tudo o que o perfil do usuário vê, inclusive preço e margem quando o perfil
permite. Esses dados passam pelo n8n e pelo OpenRouter; o termo (itens 3.2/3.3) precisa refletir isso.

## Montar no n8n
1. Criar um workflow novo e colar o conteúdo de `assistente-seven-os.json` no editor (Ctrl+V).
2. Conferir a credencial do nó OpenRouter Chat Model e o modelo.
3. Ativar o workflow e copiar a URL de produção do Webhook.
4. O salvamento de execuções está desligado no workflow, para perguntas e dados não ficarem no histórico do n8n.

O Webhook não usa credencial: o segredo é o caminho da URL. Trate a URL como senha.

## Ligar no os_api (projeto `seven-os-api` na Vercel)
- `OS_N8N_ASSISTENTE_URL` = URL de produção do Webhook
- `OS_N8N_TIMEOUT` (opcional, padrão 50)
- `OS_API_URL_PUBLICA` (opcional) = endereço público da API, se o detectado automaticamente não servir
Refazer o deploy. A resposta de `POST /os/assistente` traz `"motor": "ia"` ou `"motor": "regras"`.
