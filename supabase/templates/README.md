# Modelos de e-mail do Supabase Auth (Seven OS)

Onde colar: painel do Supabase > Authentication > Emails > Templates.

| Modelo no Supabase | Arquivo | Assunto |
|---|---|---|
| Invite user | `convite.html` | Seu acesso ao Seven OS |
| Reset password | `redefinir-senha.html` | Redefinir sua senha do Seven OS |

Os outros modelos (Confirm signup, Magic link, Change email) ficam como estao: sao usados pela tela atual
da Karen e nao devem mudar antes do merge.

O botao leva para `/definir-senha?convite=<codigo>` do Seven OS (a pagina troca o codigo pelo login).
O endereco vem do `redirectTo` que a funcao `os-convidar` manda (variavel `OS_CONVITE_REDIRECT`), que precisa
estar em Authentication > URL Configuration > Redirect URLs. Sem ele, o e-mail cai no link padrao do Supabase.

Requisito: SMTP proprio (Authentication > Emails > SMTP Settings). Sem ele o Supabase so entrega para
e-mails da equipe do projeto, com limite baixo por hora.

O logo vem de `https://seven-swart-delta.vercel.app/logo.png` (o front em producao).
