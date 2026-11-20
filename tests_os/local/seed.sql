-- Dados sinteticos para os testes de integracao (nada da Seven). Roda depois das migrations, como superusuario.
alter role os_app login;
insert into auth.users (email) values ('admin@teste'), ('supervisor@teste'), ('cotacao@teste'), ('suporte@teste'), ('cotacao1@teste')
on conflict (email) do nothing;
insert into os_homolog.usuarios (user_id, nome, email, perfil, ativo)
select id, initcap(split_part(email, '@', 1)), email,
       case split_part(email, '@', 1) when 'admin' then 'administrador' when 'supervisor' then 'supervisor' when 'suporte' then 'suporte' else 'cotacao' end, true
from auth.users where email like '%@teste' on conflict (user_id) do nothing;
insert into os_homolog.usuario_contas (user_id, conta_id)
select u.user_id, c.id from os_homolog.usuarios u cross join os_homolog.contas_portal c
where u.email <> 'cotacao1@teste' or c.conta_robo = 'filial1' on conflict do nothing;

delete from public.eventos_coletados where numero_evento like 'T%';
insert into public.eventos_coletados (conta, numero_evento, data_evento, uf, titulo, descricao, quantidade, unidade, marca, data_coleta)
select conta, ev, to_char(current_date + dias, 'DD/MM/YYYY'), 'MG', cod, descr, qtd, 'PC-peca', marca, now() - (case when dias < 0 then interval '40 days' else interval '0' end)
from (values
  ('filial1', 'T1001', 1,   '13900001', 'ROLAMENTO ESFERA;TIPO:RIGIDO;DIAMETRO INTERNO:25MM;DIAMETRO EXTERNO:52MM; FABRICANTE/REFERENCIA : SKF / 6205-2RS;', '10', 'SKF'),
  ('filial2', 'T1001', 1,   '13900001', 'ROLAMENTO ESFERA;TIPO:RIGIDO;DIAMETRO INTERNO:25MM;DIAMETRO EXTERNO:52MM; FABRICANTE/REFERENCIA : SKF / 6205-2RS;', '10', 'SKF'),
  ('filial1', 'T1002', 5,   '13900002', 'ROLAMENTO; APLICACAO: BOMBA; TIPO: ESFERA; FABRICANTE/REFERENCIA : SKF / 6205 2RS;', '2', 'SKF'),
  ('filial2', 'T1003', 3,   '35900003', 'KIT REPARO', '1', null),
  ('filial1', 'T1004', 10,  '13900001', '* ROLAMENTO ESFERA; TIPO: RIGIDO; DIAMETRO INTERNO: 25MM; DIAMETRO EXTERNO: 52MM | ROLAMENTO ESFERA', '4', 'SKF'),
  ('filial1', 'T0900', -30, '13900001', 'ROLAMENTO ESFERA;TIPO:RIGIDO;DIAMETRO INTERNO:25MM;DIAMETRO EXTERNO:52MM; FABRICANTE/REFERENCIA : SKF / 6205-2RS;', '6', 'SKF')
) v(conta, ev, dias, cod, descr, qtd, marca);
set role os_app;
select os_homolog.sincronizar();
reset role;
analyze;
