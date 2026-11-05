# Spec 01: Acesso, perfis e administração

- Etapa da proposta: Etapa 1 (acesso por usuário, credenciais de cada conta).
- Repositórios: `backendsevenapi` (`os_api`) e `seven` (`/os/admin`).
- Depende de: Spec 00.
- Status: rascunho, 05/11/2026.

## Objetivo

Garantir que cada pessoa veja e faça só o que o perfil dela permite, em todas as telas, exportações, painel e assistente, e que contas e portais sejam liberados por usuário.

## Dentro

- Login pelo Supabase Auth, cadastro só por convite.
- Três perfis da Seven: Administrador, Supervisor e Cotação.
- Acesso interno de suporte da SSYS, sem preço, auditado.
- Liberação de contas e portais por usuário.
- Telas de usuários, contas e portais, credenciais e parâmetros.
- Proteger as rotas do robô: `/run-robot` e `/status` exigem login e perfil.

## Fora

- Perfis editáveis pela Seven pela tela. Nesta fase a matriz é fixa no código e testada.
- Autenticação em dois fatores (avaliar depois).

## Como funciona

Dois eixos de acesso:

1. Perfil: o que a pessoa pode fazer.
2. Contas e portais: onde ela enxerga. Toda consulta do `os_api` filtra pelas contas liberadas.

O perfil do serviço atual (`public.perfis`, `admin` ou `operador`) continua valendo para a tela `/` até o corte. O OS usa a própria tabela de perfis.

## Matriz de permissões

| Permissão | Administrador | Supervisor | Cotação | Suporte SSYS |
|---|---|---|---|---|
| Ver fila, ficha e fornecedores | Sim | Sim | Sim | Sim |
| Rodar e baixar capturas | Sim | Sim | Sim | Sim |
| Vincular fornecedor e resolver conferência | Sim | Sim | Sim | Não |
| Ver preço de compra | Sim | Sim | Sim | Não |
| Ver margem e preço de venda | Sim | Sim | Não | Não |
| Lançar cotação | Sim | Sim | Sim (ver pendência 1) | Não |
| Ver valor em aberto e indicadores de margem | Sim | Sim | Não | Não |
| Exportar listas | Sim | Sim | Sem colunas restritas | Sem colunas restritas |
| Gerir usuários, contas, credenciais e parâmetros | Sim | Não | Não | Sim, sem ver senha |
| Ver auditoria | Sim | Sim | Não | Sim |

Permissões no código, como lista fixa: `fila.ver`, `captura.rodar`, `captura.baixar`, `identificacao.resolver`, `fornecedor.vincular`, `preco_compra.ver`, `margem.ver`, `venda.ver`, `cotacao.lancar`, `valor_aberto.ver`, `exportar`, `admin.usuarios`, `admin.contas`, `admin.credenciais`, `admin.parametros`, `auditoria.ver`.

## Dados

- `os.usuarios`: user_id (auth.users), nome, email, perfil (`administrador`, `supervisor`, `cotacao`, `suporte`), ativo (default false), criado_em, convidado_por.
- `os.contas_portal`: id, portal (`coupa` hoje), nome, cnpj, filial, ativa, criada_em. Prepara outros portais.
- `os.usuario_contas`: user_id, conta_id. Sem linha, sem acesso.
- `os.parametros`: chave, valor jsonb, alterado_por, alterado_em. Ex.: agenda do robô, limites de alerta, listas de motivo e categoria.

## Credenciais

- As senhas continuam em `public.credenciais_coupa`, cifradas com Fernet, lidas só pelo serviço do robô com a chave secreta. O `os_api` não tem acesso.
- A tela de credenciais do OS chama as rotas que já existem no serviço atual (`GET/PUT /credenciais`), que cifram e nunca devolvem senha.
- No `.env` da VPS ficam só a chave secreta do Supabase e a chave de cifra.
- Plano B temporário (decisão de 05/11): o `.env` vale só para as duas contas atuais e só quando a conta não tem registro no banco. Senha ilegível no banco bloqueia a conta, sem cair no `.env`. Conta ou portal novo usa só o banco.
- O robô registra no log a origem de cada senha (banco ou `.env`). O plano B sai quando o cadastro do OS estiver validado, e então as senhas saem do `.env`.
- Portal que não é Coupa: criar `credenciais_portal` com a mesma proteção quando o primeiro entrar.

## API

- `GET /os/eu`: perfil, permissões e contas liberadas.
- `GET/POST/PATCH /os/admin/usuarios`: convidar, ativar, desativar, mudar perfil e contas.
- `GET/POST/PATCH /os/admin/contas`: contas e portais.
- `GET/PUT /os/admin/parametros`.
- Toda rota checa permissão no servidor. O front só esconde botões; quem decide é o `os_api`.

## Telas

- `/os/admin/usuarios`: lista, convite por e-mail, perfil, contas liberadas, ativo.
- `/os/admin/contas`: contas e portais, situação da credencial (cadastrada, origem da última leitura), botão para cadastrar ou trocar senha.
- `/os/admin/parametros`: agenda do robô (horários e rodadas por dia), limites de alerta, listas.

## Regras

- Usuário novo nasce inativo e sem contas.
- Toda mudança de perfil, conta ou parâmetro vai para a auditoria.
- Acesso do suporte SSYS fica registrado na auditoria (termo, itens 3.1.b e 3.1.d).
- Colunas restritas somem também da exportação e das respostas do assistente.

## Aceite

- Usuário de Cotação não vê margem em nenhuma tela, exportação ou resposta da API.
- Usuário sem conta liberada não vê nenhum item daquela conta.
- `/run-robot` sem token retorna 401, e a Karen continua rodando pela tela `/` (o front passa a enviar o token antes de o backend exigir).
- Credencial nova cadastrada pela tela é usada pelo robô na próxima rodada, com origem "banco" no log.

## Testes

- Teste por perfil em cada rota com dado restrito.
- Teste de exportação por perfil.
- Teste de rota do robô em duas etapas: primeiro o front envia o token; depois o backend exige.

## Pendências

1. Quem digita preço de venda e margem. Cotação vê preço de compra e não vê margem. Se Cotação digitar o preço de venda vendo o de compra, ela calcula a margem de cabeça, e a restrição vira formalidade. Opções: (a) Cotação lança compra, IPI e frete, e Supervisor define venda; (b) Cotação digita venda sem ver histórico de margem, aceitando que ela deduz. Decidir com a Seven.
2. E-mails da Emanuele e do Fred e demais usuários: onboarding presencial.
3. Lista de motivos de declínio e de categorias: Seven.
