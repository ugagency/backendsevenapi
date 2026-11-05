# Spec 03: Identificação do produto e base histórica

- Etapa da proposta: Etapa 1 (reconhecimento de produto, códigos, conferência manual, carga do histórico).
- Repositórios: `backendsevenapi` (`os_api` e scripts de carga) e `seven` (`/os/conferencia`).
- Depende de: Spec 02.
- Status: rascunho, 05/11/2026.

## Objetivo

Saber qual é a peça física por trás de cada linha, independente do código Vale, e carregar o histórico comercial da Seven ligado a essas peças.

## Dentro

- Cascata de identificação: descrição estruturada, fabricante e referência, IA para o resto, conferência manual.
- Níveis de confiança: Certeza, Provável, A conferir.
- Tela de conferência manual.
- Carga das linhas completas a partir das planilhas do Storage.
- Carga da base da Seven (BASE.xlsx) e da planilha de controle da Karen, com origem marcada.

## Fora

- Categoria dos itens (pendência da Seven).
- Preço enviado no Coupa (Spec 06).

## Como funciona: identificação

1. Descrição estruturada: os pares CAMPO:VALOR são normalizados e formam a assinatura do produto.
2. Fabricante e referência confirmam ou desempatam.
3. Modelo de IA (OpenRouter) só para o que sobrar, extraindo campos, sem decidir sozinho.
4. Descrição insuficiente, como a maioria do prefixo 35, vai para conferência manual.

| Nível | Quando | Efeito |
|---|---|---|
| Certeza | Assinatura igual, ou confirmada por fabricante e referência | Junta histórico automaticamente |
| Provável | Assinatura parecida, sem confirmação | Sugere juntar; pessoa confirma |
| A conferir | Informação insuficiente | Vai para conferência manual |

Retrato do banco em 05/11: prefixo 13 com 16.507 linhas (99,2% com descrição estruturada), 15 com 7.590 (98,2%), 35 com 3.382 (27,5%); 34.927 linhas no total (89,3%). É medição por padrão de texto, não taxa de identificação. A taxa real sai depois de rodar a cascata.

## Como funciona: base histórica

1. Recupera as linhas completas das planilhas do Storage (desde 01/04) e une com o banco, sem duplicar.
2. Roda a identificação em toda a base.
3. Importa a base da Seven, ligando cada registro à linha ou ao produto.
4. DECLINAR, PROCURAR, INTERNET e NECESSITA VALIDAÇÃO viram status, não cadastro de fornecedor.
5. Preço zerado vira "sem preço". Suspeita de total no lugar de unitário é sinalizada, nunca corrigida sozinha.
6. Wemerson confere uma amostra antes de valer.
7. Na virada, carga final com o que mudou desde a primeira.

## Dados

- `os.produtos`: id, nome, fabricante, referencia, assinatura, nivel, criado_por, criado_em.
- `os.produto_codigos`: produto_id, codigo, motivo, nivel, juntado_por, juntado_em, desfeito_em.
- `os.linha_produto`: linha_id, produto_id, nivel, metodo (estrutura, referencia, ia, pessoa).
- `os.conferencias`: linha_id, situacao (aberta, associada, novo_produto, nao_identificavel), decidido_por, decidido_em. A decisão vale para a próxima vez que o caso aparecer.
- `os.historico_importado`: id, origem (base_seven, planilha_karen, portal, equipe, ia), arquivo, linha_original jsonb, produto_id, fornecedor_texto, preco_compra, ipi, frete, margem, preco_venda, status, data_ref. Valores só saem pelo `os_api`, conforme perfil.

## API

- `GET /os/conferencia`: fila de conferência, com descrição, link do evento e produtos parecidos.
- `POST /os/conferencia/{linha}`: associar, criar novo ou marcar como não identificável.
- `POST /os/produtos/{id}/desfazer-juncao`.
- Scripts de carga versionados, rodando fora do horário da Karen.

## Telas

`/os/conferencia`: descrição, link no Coupa, sugestões de produto parecido, três ações.

## Permissões

- Resolver conferência e desfazer junção: Administrador, Supervisor, Cotação.
- Valores importados (compra, margem, venda) seguem a matriz da Spec 01.

## Regras

- Toda junção pode ser desfeita; a linha original nunca muda.
- Histórico e alerta de preço só comparam produtos em Certeza. Comparar por código gerou cerca de 65% de falso alerta na análise da base.
- O sistema nunca força identificação para esvaziar a conferência.
- Todo registro guarda a origem. Parte da planilha da Karen foi preenchida por IA sem marcação: importar como origem "planilha Karen" e sinalizar.
- A base original fica preservada e a carga pode ser refeita.
- Preço nunca vai para bucket público nem circula em canal aberto.

## Aceite

- Histórico carregado e conferido por amostra pelo Wemerson.
- Códigos diferentes da mesma peça juntos, como TRAMONTINA 56134/027.
- Prefixo 35 com descrição insuficiente em conferência manual.
- Preço oculto para perfil sem permissão.

## Testes

- Conjunto de casos reais anotados (mesma peça com códigos diferentes, mesmo código com peças diferentes) com taxa de acerto medida.
- Carga refeita do zero produz o mesmo resultado.

## Pendências

- BASE.xlsx: localizar o arquivo de 09/09 (11.577 linhas, jun e jul) e pedir agosto em diante.
- Planilha de controle da Karen: sem registro de recebimento. Confirmar.
- O prazo de 30 dias úteis da proposta conta a partir do envio da base completa.
