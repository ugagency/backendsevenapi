# Spec 05: Fornecedores

- Etapa da proposta: Etapa 2 (dias 10 a 15).
- Repositórios: `backendsevenapi` (`os_api`) e `seven` (`/os/fornecedores`, bloco na ficha).
- Depende de: Spec 04.
- Status: rascunho, 05/11/2026.

## Objetivo

Fornecedor deixa de ser texto digitado e vira cadastro único, ligado a produtos e marcas. Atacar a dor central da Karen: identificar o fornecedor.

## Dentro

- Cadastro único com variações de escrita agrupadas.
- Fabricante separado de fornecedor.
- Vínculo de fornecedor com produto e com marca.
- Sugestões com fonte, sempre confirmadas por pessoa.
- Solicitação de cotação por e-mail, sem preço de compra.

## Fora

- WhatsApp, ligar e registrar retorno: depois da Etapa 5 (decisão de 05/11).
- Leitura automática de e-mails de fornecedores e disparo em massa (fora da proposta).

## Como funciona

Como o vínculo nasce:

1. Da história: fornecedor com preço anterior no produto entra vinculado.
2. Da marca: quem atende a marca aparece como opção para produtos da mesma marca.
3. Da busca: para produto novo, o sistema sugere fornecedores e mostra a fonte de cada um.
4. Da equipe: toda sugestão precisa de confirmação. Confirmar vincula; descartar evita que volte.

E-mail:

1. Na ficha, a pessoa escolhe o fornecedor e clica em E-mail.
2. O texto vem pronto com peça, referência, quantidade, entrega e prazo de retorno.
3. Ela revisa, edita e envia. O envio fica registrado na ficha.
4. O retorno chega na caixa dela e o preço é lançado à mão.
5. Antes de enviar, o sistema verifica que a mensagem não tem preço de compra.

## Dados

- `os.fornecedores`: id, nome, cidade, uf, contato, telefone, email, ativo.
- `os.fornecedor_nomes`: variações de escrita (DANFOSS, DAFOSS), com quem confirmou a junção.
- `os.marcas`: id, nome, fabricante.
- `os.fornecedor_marcas`: fornecedor_id, marca_id.
- `os.produto_fornecedores`: produto_id, fornecedor_id, origem (historia, marca, busca, equipe), confirmado_por, confirmado_em.
- `os.sugestoes_fornecedor`: produto_id, fornecedor_id, fonte, situacao (aberta, confirmada, descartada).
- `os.solicitacoes_cotacao`: item_id, fornecedor_id, enviada_por, enviada_em, texto, canal (email).

## API

- `GET/POST/PATCH /os/fornecedores`, `POST /os/fornecedores/juntar`.
- `GET /os/produtos/{id}/fornecedores`, `POST /os/sugestoes/{id}/confirmar`, `POST /os/sugestoes/{id}/descartar`.
- `POST /os/solicitacoes`: monta, valida e envia o e-mail.

## Telas

- `/os/fornecedores`: busca, filtro por marca, marcas atendidas, produtos, taxa de vitória (quando houver resultado).
- Bloco na ficha: fornecedores do produto com contato, último preço (conforme perfil), data, cotações ganhas e sugestões.

## Permissões

- Vincular, confirmar, descartar e enviar e-mail: Administrador, Supervisor, Cotação.
- Último preço: segue a matriz da Spec 01.

## Regras

- Junção de nomes sugerida é confirmada por pessoa.
- Sugestão sempre mostra a fonte.
- E-mail nunca leva preço de compra; o envio é bloqueado se detectar.
- Toda vinculação e junção vai para a auditoria.

## Aceite

- Sem nomes duplicados conhecidos.
- Fabricante separado de fornecedor.
- Fornecedor reaparece na próxima ocorrência do produto.
- Sugestão mostra a fonte.
- E-mail sai sem preço de compra.

## Testes

- Base de jun e jul: 57,8% das linhas tinham fornecedor real. Conferir quantos viram cadastro e quantos sobram como status.
- Teste de bloqueio de e-mail com preço no texto.

## Pendências

- De qual caixa de e-mail sai a mensagem.
- Serviço de envio de e-mail é um terceiro novo e precisa ser comunicado à Seven (termo, item 3.2).
