# Spec 07: Painel de indicadores

- Etapa da proposta: Etapa 5 (dias 26 a 30).
- Repositórios: `backendsevenapi` (`os_api`) e `seven` (`/os/painel`).
- Depende de: Spec 06.
- Status: rascunho, 05/11/2026.

## Objetivo

Mostrar a saúde da operação e onde a Seven ganha ou perde dinheiro, sempre levando à lista de itens por trás de cada número.

## Indicadores

| Indicador | Como é calculado | Perfil |
|---|---|---|
| Itens capturados por dia | Linhas por dia, com falhas marcadas | Todos |
| Identificação por prefixo | % de linhas em Certeza por prefixo | Todos |
| Itens por estado | Contagem das abas da fila | Todos |
| Encerra nas próximas 48h | Itens abertos por prazo | Todos |
| Marcas sem fornecedor | Marcas recorrentes sem vínculo | Todos |
| Taxa de declínio por categoria | Declinados sobre capturados | Todos |
| Valor em aberto | Quantidade x último preço de venda | Administrador, Supervisor |
| Margem praticada | Margem registrada por período | Administrador, Supervisor |
| Divergências abertas | Reajuste provável, erro de lançamento, sem explicação | Administrador, Supervisor |
| Taxa de vitória | Ganhas sobre cotações com resultado | Administrador, Supervisor |

## API

- `GET /os/painel?periodo=&conta=`: só os indicadores permitidos ao perfil.
- `GET /os/painel/{indicador}/itens`: a lista por trás do número.

## Telas

`/os/painel`: período e filtro por CNPJ valem para o painel inteiro. Cada número abre a lista.

## Regras

- Indicador sem base suficiente diz "indisponível" e explica por quê.
- Taxa de vitória fica indisponível até haver resultado suficiente (limite em parâmetros).
- Classificar divergência é sugestão; quem decide é a pessoa.
- Alerta de falha do robô aparece no painel.

## Aceite

- Indicadores conferem com consulta direta à base.
- Usuário de Cotação não vê indicadores restritos, nem pela API.

## Pendências

- Categoria (Mobiliário, Rolamentos) não existe na captura. Pode vir do primeiro termo da descrição ou de uma lista curta mantida pela Seven. Sem isso, a taxa de declínio sai sem quebra por categoria.
- Limite de resultados para liberar a taxa de vitória.
