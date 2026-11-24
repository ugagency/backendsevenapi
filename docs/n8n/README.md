# Assistente do Seven OS no n8n

O n8n faz uma coisa só: recebe a pergunta, pede ao modelo (OpenRouter) para escolher UMA consulta da lista fechada e devolve
`{"consulta": "...", "parametros": {...}}`. Quem executa a consulta, aplica a permissão e mostra preço é sempre o `os_api`.

## O que sai do sistema para o n8n
Somente: texto da pergunta, data de hoje, dia da semana e a lista de nomes das consultas. Nunca histórico de conversa, preço,
margem, fornecedor ou qualquer dado da Seven. O `os_api` valida o retorno: consulta fora da lista ou parâmetro desconhecido é
descartado e o assistente usa as regras internas, sem erro para o usuário.

## Montar no n8n (uma vez)
1. Importar `assistente-seven-os.json` (Workflows > Import from file).
2. Credencial do nó Webhook (Header Auth): nome do cabeçalho `X-Seven-Token`, valor um segredo longo que você inventa.
3. Credencial do nó OpenRouter (Header Auth): cabeçalho `Authorization`, valor `Bearer <sua chave do OpenRouter>`.
4. O modelo está como `openai/gpt-4o-mini`; troque pelo que já usam no n8n, se preferir. O corpo já envia
   `provider.data_collection = deny`, para o OpenRouter recusar provedores que guardam ou treinam com os dados (termo, item 3.3).
5. O workflow já está com o salvamento de execuções desligado, para o texto das perguntas não ficar no histórico do n8n.
   Confira essa configuração depois de importar.
6. Ativar o workflow e copiar a URL de produção do Webhook (`.../webhook/seven-os-assistente`, não a de teste).

## Ligar no os_api (projeto `seven-os-api` na Vercel)
- `OS_N8N_ASSISTENTE_URL` = URL de produção do Webhook
- `OS_N8N_TOKEN` = o mesmo segredo do passo 2
- `OS_N8N_TIMEOUT` (opcional, padrão 12 segundos)
Depois, refazer o deploy. A resposta de `POST /os/assistente` passa a trazer `"motor": "ia"` ou `"motor": "regras"`.
