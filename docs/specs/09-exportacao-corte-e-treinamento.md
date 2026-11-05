# Spec 09: Exportação completa, corte e treinamento

- Etapa da proposta: Etapa 5 (dias 26 a 30) e homologação final.
- Repositórios: `backendsevenapi` e `seven`.
- Depende de: todas as anteriores.
- Status: rascunho, 05/11/2026.

## Objetivo

Fechar o que o termo exige (exportação completa e auditoria consultável), virar a operação da tela atual para o OS e treinar a equipe.

## Dentro

- Exportação da base inteira em formato aberto (termo, itens 4.2 e 4.3).
- Tela de auditoria: quem lançou ou alterou preço, margem e fornecedor, e quem exportou.
- Carga final de diferença da base da Seven.
- Corte: `/` redireciona para `/os`, com a tela atual dentro de Captura.
- Retirada do plano B do `.env` e apagamento das senhas do Coupa do `.env` da VPS.
- Treinamento da equipe.
- Fechamento de segurança, por último: trancar leitura e inserção públicas em `eventos_coletados`, restringir alteração e exclusão no bucket `planilhas`, apagar `user_settings` e trocar as senhas do Coupa. Antes, confirmar que o robô grava com a chave de servidor e que nenhum fluxo externo (n8n) lê `user_settings`.

## Como funciona: corte

1. Carga final com o que mudou na base da Seven desde a primeira carga.
2. Agenda automática ligada (se ainda não estiver), com a Karen usando o download por data do OS.
3. `/` passa a redirecionar para `/os`.
4. Uma semana de convivência: a tela antiga continua acessível por endereço direto.
5. Depois disso, retirada das estruturas antigas em `public`, só com volta escrita.

## API

- `POST /os/admin/exportar-base`: gera o pacote completo, auditado, só Administrador.
- `GET /os/auditoria`: filtros por entidade, autor e período.

## Telas

- `/os/admin/auditoria`.
- `/os/admin/exportar`.

## Aceite

- Exportação completa abre em planilha comum e bate com o banco.
- Equipe treinada, com roteiro de uso por perfil.
- Karen faz um dia inteiro de trabalho só no OS.
- Robô lê senha só do banco, conferido no log.

## Pendências

- Material de treinamento: formato (presencial, vídeo, roteiro escrito).
- Prazo da convivência entre a tela antiga e o OS.
