# Spec 04: Fila e ficha

- Etapa da proposta: versão somente leitura logo após a Etapa 1; versão completa na Etapa 4 (dias 22 a 25).
- Repositórios: `backendsevenapi` (`os_api`) e `seven` (`/os`, `/os/item/[id]`).
- Depende de: Spec 03.
- Status: rascunho, 05/11/2026.

## Objetivo

Dar à equipe uma tela inicial com todos os itens em aberto, ordenados pelo que precisa de ação primeiro, e uma ficha que reúne tudo sobre cada item.

## Entrega em duas partes

1. Somente leitura (após a Etapa 1): abas, lista, ficha com identificação e histórico. A Karen valida a identificação com casos reais.
2. Completa (Etapa 4): filtros, ordenações, cartões do topo, exportação por perfil, indicação de quem está cotando.

## Como funciona

Item é um produto dentro de um evento. Várias linhas do mesmo evento e produto viram um item, com a contagem visível ("4 linhas no mesmo evento"). Evento em dois CNPJs é um item só, com as filiais indicadas.

Abas por prontidão:

| Aba | Critério |
|---|---|
| Pronto | Identidade em Certeza e ao menos um fornecedor vinculado |
| Falta fornecedor | Produto conhecido, sem fornecedor vinculado |
| Produto novo | Primeira aparição, sem histórico |
| Conferência manual | Identidade a conferir |

Andamento, separado da aba: A cotar, Em cotação, Enviada no Coupa, Declinada, Expirada, Ganha ou Perdida. A aba muda sozinha; o andamento muda por ação da equipe.

## Dados

- `os.itens`: id, produto_id, numero_evento, contas (array), prazo, quantidade_total, linhas, aba, andamento, responsavel_id, atualizado_em.
- `os.item_andamento`: histórico de mudanças de andamento, com autor, data e motivo.

## API

- `GET /os/fila`: aba, filtros (busca por descrição, código ou marca; UF; prefixo; data de captura; prazo; CNPJ), ordenação (prazo, valor estimado, recorrência), paginação.
- `GET /os/fila/cartoes`: em aberto, aguardando você, encerra em 48h, valor estimado (só com permissão).
- `GET /os/itens/{id}`: ficha.
- `PATCH /os/itens/{id}/andamento`: com motivo obrigatório em Declinada.
- `GET /os/fila/exportar`: exatamente a lista filtrada, nos limites do perfil, auditada.

## Telas

- `/os`: fila com abas, cartões, colunas (item com nome, marca e código; o que já sabemos; encerramento; estado), filtros e ordenação.
- `/os/item/[id]`, a ficha:

| Bloco | O que mostra |
|---|---|
| Cabeçalho | Nome, fabricante, referência, nível, UF, quantidade, prazo, ocorrências, CNPJs, link do evento no Coupa |
| Mesma peça, códigos diferentes | Códigos ligados, com data e motivo |
| Fornecedores | Entra na Spec 05 |
| Cotações anteriores | Data, código, fornecedor, compra, margem, resultado, conforme perfil |
| Nova cotação | Entra na Spec 06 |
| Declínios anteriores | Quantas vezes o produto foi declinado e por quê |

## Permissões

- Ver: todos os perfis, só nas contas liberadas.
- Valor estimado e margem: Administrador e Supervisor.

## Regras

- Item sai da fila aberta quando é enviado, declinado ou passa do prazo sem ação (Expirada). Continua consultável por filtro.
- Vincular fornecedor move o item de "Falta fornecedor" para "Pronto" sozinho.
- "Em cotação por Karen" aparece para os outros, sem bloquear.
- Item em conferência mostra primeiro as ações de identificação.
- Registro incerto mostra a origem ("importado da base", "informado pela equipe").

## Aceite

- Abas, filtros e ordenações funcionam com dados reais.
- Filtro por CNPJ.
- Exportação igual à tela e limitada pelo perfil.
- Usuário de Cotação não vê valor estimado nem margem.

## Testes

- Comparar contagem das abas com consulta direta ao banco.
- Exportação por perfil, coluna a coluna.

## Pendências

- Hora do prazo: se o Coupa não expuser, "encerra em 48h" vira "encerra até depois de amanhã".
- Valor estimado vem do último preço de venda. Sem histórico, não aparece.
