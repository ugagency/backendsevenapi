# Spec 06: Cotação

- Etapa da proposta: Etapa 3 (dias 16 a 21).
- Repositórios: `backendsevenapi` (`os_api`, robô para conferência com o Coupa) e `seven` (bloco na ficha).
- Depende de: Spec 05.
- Status: rascunho, 05/11/2026.

## Objetivo

Tirar a formação de preço da planilha e registrar cada cotação do início ao resultado, com alertas antes do envio. O envio ao Coupa continua manual.

## Dentro

- Preparar, conferir, enviar e registrar o resultado.
- Alertas: total no lugar de unitário, preço divergente, fornecedor sem preço.
- Versões de cada cotação, com autor e data.
- Conferência periódica com o preço enviado no Coupa, se o dado estiver acessível.

## Fora

- Cálculo automático do preço de venda (decisão de 05/11: a equipe digita).
- Regras de margem por prazo, imposto ou cliente; ICMS e NCM (fora da proposta).
- Envio automático da resposta no portal.

## Como funciona

1. Preparar: dentro da ficha, fornecedor, preço de compra unitário, IPI, frete, margem e preço de venda.
2. Conferir: o sistema compara com o histórico do mesmo produto (só produtos em Certeza).

| Alerta | Quando dispara | Ação exigida |
|---|---|---|
| Parece total, não unitário | Valor perto de último unitário x quantidade | Corrigir ou manter com justificativa. Bloqueia até resolver |
| Preço divergente | Diferença acima do limite configurado | Corrigir ou justificar |
| Fornecedor sem preço | Fornecedor escolhido sem preço | Preencher |

3. Enviar: a pessoa responde no Coupa e muda para "Enviada no Coupa". Item não cotado vai para "Declinada", com motivo.
4. Resultado: Ganha ou Perdida, ligada ao número do pedido de venda.

## Dados

- `os.cotacoes`: id, item_id, versao, fornecedor_id, conta_id (CNPJ que respondeu), criada_por, criada_em, situacao.
- `os.cotacao_valores`: cotacao_id, versao, preco_compra, ipi, frete, margem, preco_venda. Tabela separada por ser sensível; leitura só via `os_api` conforme perfil.
- `os.alertas_preco`: cotacao_id, tipo, valor_referencia, situacao, justificativa, resolvido_por.
- `os.resultados`: cotacao_id, resultado, pedido_venda, registrado_por, registrado_em.
- `os.conferencia_coupa`: numero_evento, preco_enviado_coupa, preco_registrado, diferenca, lido_em.

## API

- `POST /os/itens/{id}/cotacoes`, `POST /os/cotacoes/{id}/versao`.
- `POST /os/alertas/{id}/justificar`.
- `POST /os/cotacoes/{id}/resultado`.

## Telas

Bloco "Nova cotação" na ficha, com os alertas à vista, e histórico de versões.

## Permissões

- Ver e lançar: segue a matriz da Spec 01. Ver pendência 1.
- Margem e preço de venda: Administrador e Supervisor.

## Regras

- Valores sempre unitários, com a quantidade à vista.
- Alterar cria nova versão; nada é sobrescrito.
- Justificativa de alerta mantido vai para a auditoria.
- Cotação registra o CNPJ que respondeu.
- O protótipo usa 75% de unitário x quantidade como gatilho do primeiro alerta. Validar com casos reais.

## Aceite

- Cotação registrada do início ao resultado.
- Alerta dispara em caso real de total no lugar de unitário.
- Autor e data em toda alteração.
- Conferência com o Coupa aponta diferença (se o dado existir).

## Testes

- Casos reais da base com total lançado como unitário.
- Usuário de Cotação não recebe margem em nenhuma resposta da API.

## Pendências

1. Quem digita margem e preço de venda, já que Cotação não vê margem (ver Spec 01).
2. Limite de divergência: Wemerson.
3. Verificar no Coupa se o preço enviado aparece em eventos encerrados.
4. Resultado: na base, só 0,4% das linhas têm esse dado e o portal não informa quem venceu. A taxa de vitória depende da disciplina da equipe.
