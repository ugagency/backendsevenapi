# Spec 08: Assistente

- Etapa da proposta: Etapa 5 (dias 26 a 30).
- Repositórios: `backendsevenapi` (`os_api`) e `seven` (`/os/assistente`).
- Depende de: Spec 06.
- Status: rascunho, 05/11/2026.

## Objetivo

Responder perguntas sobre os dados da Seven só com o que está registrado, como "já cotamos essa peça e por quanto?" e "o que encerra até sexta sem fornecedor?", sem enviar preço ao modelo de IA.

## Dentro

- Assistente no sistema, somente leitura nesta versão.
- Lista fechada de consultas.
- Conversas salvas por usuário.

## Fora

- Assistente no WhatsApp (fase seguinte, venda adicional).
- Gravação feita pelo assistente.

## Como funciona

1. A pergunta chega ao modelo, sem histórico nem preços anexados.
2. O modelo escolhe uma consulta de uma lista fechada: buscar produto, histórico, fornecedores, itens por prazo, marcas sem fornecedor, indicadores. Nunca escreve consulta livre ao banco.
3. O `os_api` executa a consulta com a permissão de quem perguntou.
4. O modelo recebe só dados não sensíveis e redige a frase.
5. Valores sensíveis aparecem direto na resposta, em tabela montada pelo sistema, sem passar pelo modelo.

Consequência: o assistente mostra preços, mas não raciocina sobre eles. Contas como "qual a variação?" vêm prontas do sistema.

## Dados

- `os.conversas`: id, user_id, criada_em.
- `os.mensagens`: conversa_id, papel, texto, consulta_usada, parametros, criada_em. Tabelas com valores sensíveis não são guardadas no texto do modelo.

## API

- `POST /os/assistente`: pergunta, resposta com texto, tabela do sistema e link para a ficha ou lista.
- `GET /os/assistente/conversas`.

## Regras

- Sem dado, responde que não tem. Nunca estima.
- Resposta traz link para a origem.
- OpenRouter configurado para recusar provedor que retém ou treina com os dados (termo, item 3.3).
- Teste automático: nenhuma chamada ao modelo contém campos de `cotacao_valores` ou `historico_importado` com valor.

## Aceite

- Responde as perguntas de exemplo da proposta.
- Nenhum preço enviado ao modelo, conferido no log das chamadas.
- Diz quando não sabe.
- Respeita o perfil: Cotação não recebe margem.

## Pendências

- Confirmação jurídica sobre o OpenRouter (termo, itens 3.2 e 3.3).
