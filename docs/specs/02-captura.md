# Spec 02: Captura

- Etapa da proposta: Etapa 1 (CNPJs adicionais, execução diária, prazo de encerramento).
- Repositórios: `backendsevenapi` (serviço atual e `os_api`) e `seven` (`/os/captura`).
- Depende de: Specs 00 e 01.
- Status: rascunho, 05/11/2026.

## Objetivo

Transformar o robô em porta de entrada de dados: roda sozinho nos horários configurados, grava todas as linhas completas e mantém o botão manual da Karen.

## Dentro

- Gravação de todas as linhas, inclusive as repetidas do mesmo código no evento, em tabela nova.
- Agenda configurável (horários e rodadas por dia).
- Prazo em formato de data.
- Link do evento no Coupa guardado por linha.
- Histórico de execuções por conta, com sucesso, falha e contagem.
- Planilha gerada a partir do banco, por data, no mesmo layout de hoje.
- Alerta de falha.
- Marca: migrar de Gemini direto para OpenRouter.

## Fora

- Preço enviado em eventos encerrados: depende de verificar se a tela existe no Coupa (Spec 06).
- Outros portais além do Coupa.

## Como funciona

1. No horário da agenda, o `os_api` pede uma execução ao serviço atual, conta por conta, uma de cada vez.
2. O robô lê os eventos abertos e, além do que já faz hoje, grava cada linha inteira em `os.linhas_evento`.
3. A gravação nova é protegida: falha nela não interrompe a planilha nem a gravação em `eventos_coletados`.
4. A marca é extraída e a identificação (Spec 03) roda em seguida.
5. O botão "Rodar agora" continua, com data escolhida, como hoje.

A trava do robô hoje fica em memória (`ROBO_LOCK`). A execução passa a ser registrada em `os.execucoes`, para a agenda e a tela saberem o estado mesmo depois de um reinício.

## Ponto crítico: a planilha da Karen

Hoje a geração manual só traz o que aquela execução achou de novo. Se o robô automático capturar de manhã, a geração manual dela vem vazia. Por isso:

- A agenda automática só liga quando a tela Captura do OS (download por data a partir do banco) estiver validada com a Karen.
- Até lá, a captura continua só manual.

## Dados

- `os.linhas_evento`: id, conta_id, numero_evento, titulo_evento, link_evento, numero_linha, codigo, descricao, quantidade, unidade, uf, prazo (date), prazo_texto (como veio), marca, marca_origem (regra ou IA), capturado_em, execucao_id. Único em (conta_id, numero_evento, numero_linha).
- `os.linha_evento_contas`: liga a mesma linha lógica às contas em que apareceu (evento em dois CNPJs).
- `os.execucoes`: id, conta_id, origem (agenda ou manual), pedido_por, iniciada_em, terminada_em, situacao (rodando, sucesso, falha), linhas_lidas, linhas_novas, erro.
- `os.declinados` (descarte interno, decisão de 05/11): linha ou item, motivo, autor, em. Guardado para análise, com contagem de quantas vezes o produto voltou.

## API

- `GET /os/captura/execucoes`: histórico por conta.
- `POST /os/captura/rodar`: execução manual, com data e conta.
- `GET /os/captura/planilha?data=AAAA-MM-DD&conta=`: planilha do dia a partir do banco, no layout atual (9 colunas, mais "Filiais" na unificada).
- Interno: agendador no `os_api` lê `os.parametros` e chama o serviço atual.

## Telas

`/os/captura`: execuções por conta, situação, linhas, botão "Rodar agora", opção "incluir eventos já capturados" e download por data.

## Permissões

- Rodar e baixar: Administrador, Supervisor, Cotação, Suporte.
- Mudar agenda: só Administrador (Spec 01).

## Regras

- Eventos repetidos entre contas aparecem uma vez na fila, com os CNPJs indicados.
- Item declinado não volta à fila quando o mesmo evento é relido. Um evento novo com o mesmo produto é oportunidade nova e entra, mostrando quantas vezes o produto já foi declinado.
- Falha de login, conta sem resultado ou queda forte de volume geram alerta para a SSYS e aparecem no Painel.
- Teste de robô em homologação nunca grava em `eventos_coletados`.

## Aceite

- Robô roda sozinho nas duas contas nos horários configurados.
- Todas as linhas no banco, inclusive repetidas do mesmo código no mesmo evento.
- Planilha de uma data real gerada do banco igual à de produção.
- Prazo como data em 100% das linhas novas.
- Execução aparece na tela com contagem e situação.

## Testes

- Comparar `os.linhas_evento` com as planilhas do Storage de três datas reais.
- Simular falha na gravação nova: a planilha continua saindo.
- Deploy do `os_api` durante uma execução: o robô termina.

## Pendências

- Verificar no Coupa: hora do prazo (hoje só dd/mm/aa), se a lista mostra prorrogação e cancelamento (releitura de eventos abertos).
- Limite de queda de volume para alerta.
- Quem recebe o alerta de falha na SSYS.
