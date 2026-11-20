-- Seven OS, Specs 02 a 04: sincronizacao da captura, identificacao em cascata, itens da fila e juncao de produtos.
-- Le public.eventos_coletados (so SELECT) e grava apenas no proprio schema. Nao altera nada do robo.
do $outer$
declare s text; ddl text := $ddl$
create or replace function __S__.prazo_data(t text) returns date language plpgsql immutable as $f$
begin
  if t ~ '^\d{2}/\d{2}/\d{4}$' then return to_date(t, 'DD/MM/YYYY');
  elsif t ~ '^\d{2}/\d{2}/\d{2}$' then return to_date(t, 'DD/MM/YY');
  end if; return null;
exception when others then return null; end $f$;

create or replace function __S__.qtd_num(t text) returns numeric language plpgsql immutable as $f$
declare x text := replace(trim(coalesce(t, '')), ' ', '');
begin
  if x ~ '^\d+$' then return x::numeric;
  elsif x ~ '^\d{1,3}(\.\d{3})*(,\d+)?$' then return replace(replace(x, '.', ''), ',', '.')::numeric;
  elsif x ~ '^\d+[.,]\d+$' then return replace(x, ',', '.')::numeric;
  end if; return null;
exception when others then return null; end $f$;

create or replace function __S__.desc_norm(d text) returns text language sql immutable as $f$
  select upper(trim(regexp_replace(regexp_replace(regexp_replace(
    split_part(regexp_replace(coalesce(d, ''), '^\s*\*\s*', ''), ' | ', 1),
    '\s*([;:=])\s*', '\1', 'g'), '\s+', ' ', 'g'), ';+$', '')))
$f$;

create or replace function __S__.atualizar_abas(p_produtos bigint[] default null) returns int
language plpgsql set search_path = __S__ as $f$
declare v_n int;
begin
  update itens i set aba = x.aba, atualizado_em = now()
  from (
    select i2.id, case
      when p.nivel in ('conferir', 'nao_identificavel') then 'conferencia'
      when exists (select 1 from produto_fornecedores pf where pf.produto_id = i2.produto_id) then 'pronto'
      when exists (select 1 from itens a where a.produto_id = i2.produto_id and a.id <> i2.id
                   and coalesce(a.capturado_em, 'epoch') <= coalesce(i2.capturado_em, now())) then 'falta_fornecedor'
      else 'produto_novo' end as aba
    from itens i2 join produtos p on p.id = i2.produto_id
    where i2.andamento in ('a_cotar', 'em_cotacao') and (p_produtos is null or i2.produto_id = any(p_produtos))
  ) x where x.id = i.id and i.aba is distinct from x.aba;
  get diagnostics v_n = row_count; return v_n;
end $f$;

create or replace function __S__.montar_itens(p_eventos text[] default null) returns int
language plpgsql set search_path = __S__ as $f$
declare v_n int;
begin
  with base as (
    select l.numero_evento, lp.produto_id, l.conta_id, sum(l.quantidade) q, count(*) n, max(l.prazo) prazo,
           max(l.unidade) un, max(l.uf) uf, min(l.capturado_em) cap
    from linhas_evento l join linha_produto lp on lp.linha_id = l.id
    where p_eventos is null or l.numero_evento = any(p_eventos) group by 1, 2, 3),
  ag as (
    select numero_evento, produto_id, array_agg(conta_id order by conta_id) contas, max(q) q, max(n)::int n,
           max(prazo) prazo, max(un) un, max(uf) uf, min(cap) cap from base group by 1, 2)
  insert into itens (numero_evento, produto_id, contas, quantidade_total, linhas, prazo, unidade, uf, capturado_em)
  select numero_evento, produto_id, contas, q, n, prazo, un, uf, cap from ag
  on conflict (numero_evento, produto_id) do update
    set contas = excluded.contas, quantidade_total = excluded.quantidade_total, linhas = excluded.linhas,
        prazo = excluded.prazo, atualizado_em = now()
    where (itens.contas, itens.quantidade_total, itens.linhas, itens.prazo)
          is distinct from (excluded.contas, excluded.quantidade_total, excluded.linhas, excluded.prazo);
  get diagnostics v_n = row_count;
  if p_eventos is not null then
    delete from itens i where i.numero_evento = any(p_eventos) and i.andamento = 'a_cotar'
      and not exists (select 1 from cotacoes c where c.item_id = i.id)
      and not exists (select 1 from linhas_evento l join linha_produto lp on lp.linha_id = l.id
                      where l.numero_evento = i.numero_evento and lp.produto_id = i.produto_id);
  end if;
  return v_n;
end $f$;

create or replace function __S__.sincronizar(p_por uuid default null) returns jsonb
language plpgsql set search_path = __S__ as $f$
declare v_exec bigint; v_linhas int; v_prod int; v_melhora int; v_itens int; v_eventos text[]; v_produtos bigint[];
begin
  if not pg_try_advisory_xact_lock(hashtext('__S__.sincronizar')) then
    return jsonb_build_object('ocupado', true);
  end if;
  insert into execucoes (origem, pedido_por) values ('sincronizacao', p_por) returning id into v_exec;

  with novas as (
    select e.*, c.id as conta_id from public.eventos_coletados e
    join contas_portal c on c.portal = 'coupa' and c.conta_robo = e.conta
    where not exists (select 1 from linhas_evento l where l.origem_id = e.id)),
  ins as (
    insert into linhas_evento (conta_id, numero_evento, numero_linha, codigo, descricao, quantidade, quantidade_texto,
        unidade, uf, prazo, prazo_texto, marca, marca_origem, capturado_em, data_inicio, execucao_id, origem, origem_id)
    select n.conta_id, n.numero_evento,
      coalesce(m.mx, 0) + row_number() over (partition by n.conta_id, n.numero_evento order by n.id),
      substring(n.titulo from '^\s*(\d{6,10})'), n.descricao, qtd_num(n.quantidade), n.quantidade, n.unidade, n.uf,
      prazo_data(n.data_evento), n.data_evento,
      case when n.marca is null or trim(n.marca) = '' or upper(n.marca) like 'MARCA NAO ENCONTRADA%' then null else upper(trim(n.marca)) end,
      case when n.marca is null then null else 'captura' end, n.data_coleta, n.data_inicio, v_exec, 'eventos_coletados', n.id
    from novas n
    left join (select conta_id, numero_evento, max(numero_linha) mx from linhas_evento group by 1, 2) m
      on m.conta_id = n.conta_id and m.numero_evento = n.numero_evento
    returning numero_evento)
  select count(*), array_agg(distinct numero_evento) into v_linhas, v_eventos from ins;

  -- Cascata: o codigo Vale identifica a peca (descricoes do mesmo codigo variam em espaco, idioma e corte).
  -- Fabricante/referencia e descricao estruturada dao o nivel Certeza; sem isso o produto vai para conferencia.
  -- Mesma referencia em codigos diferentes NAO junta sozinha: vira sugestao (ref_chave), confirmada por pessoa.
  with calc0 as (
    select l.id, l.codigo, l.descricao, l.marca, desc_norm(l.descricao) dn,
           regexp_match(upper(coalesce(l.descricao, '')), 'FABRICANTE\s*/\s*REFERENCIA\s*:\s*([^/;|]+?)\s*/\s*([^;|]+)') fr
    from linhas_evento l where not exists (select 1 from linha_produto lp where lp.linha_id = l.id)),
  calc1 as (
    select *, (fr is not null and length(trim(fr[2])) >= 3 and trim(fr[2]) ~ '\d') tem_ref,
           (length(dn) >= 30 and length(dn) - length(replace(dn, ':', '')) >= 2) estruturada from calc0),
  calc as materialized (
    select id, codigo, descricao, dn,
      case when tem_ref then regexp_replace(trim(fr[1]), '\s+', ' ', 'g') else marca end fab,
      case when tem_ref then trim(fr[2]) end ref,
      case when tem_ref then regexp_replace(trim(fr[1]), '\s+', ' ', 'g') || '|' || regexp_replace(trim(fr[2]), '[\s\-\./]', '', 'g') end ref_chave,
      case when tem_ref then 2 when estruturada then 1 else 0 end qualidade,
      case when codigo is not null then 'V|' || codigo
           when tem_ref then 'R|' || regexp_replace(trim(fr[1]), '\s+', ' ', 'g') || '|' || regexp_replace(trim(fr[2]), '[\s\-\./]', '', 'g')
           when estruturada then 'E|' || md5(dn) else 'D|' || md5(dn) end as assinatura,
      left(coalesce(nullif(trim(case when dn like 'PARTES E PECAS%' and dn like '%NOME DO ITEM:%'
             then substring(dn from 'NOME DO ITEM:([^;]+)') else split_part(regexp_replace(dn, '^[A-Z]{2} \|\| ', ''), ';', 1) end), ''),
           case when codigo is not null then 'Codigo ' || codigo else 'Sem descricao' end), 120) as nome
    from calc1),
  melhor as (
    select distinct on (assinatura) * from calc order by assinatura, qualidade desc, (dn like 'EN ||%'), id),
  novos as (
    insert into produtos (nome, descricao, fabricante, referencia, ref_chave, assinatura, prefixo, nivel, metodo)
    select nome, left(descricao, 2000), fab, ref, ref_chave, assinatura, left(codigo, 2),
           case when qualidade > 0 then 'certeza' else 'conferir' end,
           case qualidade when 2 then 'referencia' when 1 then 'estrutura' else 'codigo' end
    from melhor on conflict (assinatura) do nothing returning id, assinatura),
  melhora as (
    update produtos p set nivel = 'certeza', nome = b.nome, descricao = left(b.descricao, 2000),
        fabricante = coalesce(b.fab, p.fabricante), referencia = coalesce(b.ref, p.referencia), ref_chave = coalesce(b.ref_chave, p.ref_chave),
        metodo = case b.qualidade when 2 then 'referencia' else 'estrutura' end
    from melhor b where p.assinatura = b.assinatura and b.qualidade > 0 and p.metodo <> 'pessoa'
      and (p.nivel = 'conferir' or (p.referencia is null and b.qualidade = 2)) returning p.id),
  todos as (
    select id, assinatura from novos
    union all
    select coalesce(p.juntado_em, p.id), p.assinatura from produtos p where p.assinatura in (select assinatura from calc)),
  liga as (
    insert into linha_produto (linha_id, produto_id, metodo)
    select c.id, t.id, case c.qualidade when 2 then 'referencia' when 1 then 'estrutura' else 'codigo' end
    from calc c join todos t on t.assinatura = c.assinatura
    on conflict (linha_id) do nothing returning produto_id)
  select (select count(*) from novos), (select count(*) from melhora), (select array_agg(distinct produto_id) from liga) into v_prod, v_melhora, v_produtos;

  insert into produto_codigos (produto_id, codigo, motivo)
  select distinct lp.produto_id, l.codigo, 'captura' from linha_produto lp join linhas_evento l on l.id = lp.linha_id
  where l.codigo is not null and lp.produto_id = any(coalesce(v_produtos, '{}'))
  on conflict do nothing;

  v_itens := montar_itens(case when coalesce(array_length(v_eventos, 1), 0) > 3000 then null else coalesce(v_eventos, '{}') end);
  perform atualizar_abas(case when coalesce(array_length(v_produtos, 1), 0) > 3000 then null else coalesce(v_produtos, '{}') end);

  update execucoes set terminada_em = now(), situacao = 'sucesso', linhas_lidas = v_linhas, linhas_novas = v_linhas where id = v_exec;
  return jsonb_build_object('execucao', v_exec, 'linhas_novas', v_linhas, 'produtos_novos', v_prod, 'itens_tocados', v_itens);
end $f$;

create or replace function __S__.juntar_produto(p_origem bigint, p_destino bigint, p_por uuid, p_motivo text default null) returns bigint
language plpgsql set search_path = __S__ as $f$
declare v_linhas bigint[]; v_eventos text[]; v_id bigint;
begin
  if p_origem = p_destino then raise exception 'origem e destino iguais'; end if;
  if (select juntado_em from produtos where id = p_destino) is not null then raise exception 'destino ja foi juntado a outro produto'; end if;
  if not exists (select 1 from produtos where id = p_origem and juntado_em is null) then raise exception 'origem inexistente ou ja juntada'; end if;
  select array_agg(lp.linha_id), array_agg(distinct l.numero_evento) into v_linhas, v_eventos
    from linha_produto lp join linhas_evento l on l.id = lp.linha_id where lp.produto_id = p_origem;
  insert into produto_juncoes (origem_id, destino_id, linhas, motivo, por)
    values (p_origem, p_destino, coalesce(v_linhas, '{}'), p_motivo, p_por) returning id into v_id;
  update linha_produto set produto_id = p_destino, metodo = 'pessoa' where produto_id = p_origem;
  insert into produto_codigos (produto_id, codigo, motivo, juntado_por)
    select p_destino, codigo, 'juncao:' || v_id, p_por from produto_codigos where produto_id = p_origem on conflict do nothing;
  insert into produto_fornecedores (produto_id, fornecedor_id, origem, confirmado_por)
    select p_destino, fornecedor_id, origem, confirmado_por from produto_fornecedores where produto_id = p_origem on conflict do nothing;
  update produtos set juntado_em = p_destino where id = p_origem or juntado_em = p_origem;
  update itens i set produto_id = p_destino where i.produto_id = p_origem
    and (i.andamento <> 'a_cotar' or exists (select 1 from cotacoes c where c.item_id = i.id))
    and not exists (select 1 from itens d where d.produto_id = p_destino and d.numero_evento = i.numero_evento);
  perform montar_itens(coalesce(v_eventos, '{}'));
  perform atualizar_abas(array[p_destino]);
  return v_id;
end $f$;

create or replace function __S__.desfazer_juncao(p_id bigint, p_por uuid) returns void
language plpgsql set search_path = __S__ as $f$
declare j produto_juncoes; v_eventos text[];
begin
  select * into j from produto_juncoes where id = p_id and desfeito_em is null for update;
  if not found then raise exception 'juncao inexistente ou ja desfeita'; end if;
  update linha_produto set produto_id = j.origem_id where linha_id = any(j.linhas) and produto_id = j.destino_id;
  update produtos set juntado_em = null where id = j.origem_id;
  delete from produto_codigos pc where pc.produto_id = j.destino_id and pc.motivo = 'juncao:' || p_id
    and not exists (select 1 from linha_produto lp join linhas_evento l on l.id = lp.linha_id
                    where lp.produto_id = j.destino_id and l.codigo = pc.codigo);
  select array_agg(distinct numero_evento) into v_eventos from linhas_evento where id = any(j.linhas);
  update itens i set produto_id = j.origem_id where i.produto_id = j.destino_id and i.numero_evento = any(coalesce(v_eventos, '{}'))
    and not exists (select 1 from linhas_evento l join linha_produto lp on lp.linha_id = l.id
                    where l.numero_evento = i.numero_evento and lp.produto_id = j.destino_id)
    and not exists (select 1 from itens d where d.produto_id = j.origem_id and d.numero_evento = i.numero_evento);
  perform montar_itens(coalesce(v_eventos, '{}'));
  perform atualizar_abas(array[j.origem_id, j.destino_id]);
  update produto_juncoes set desfeito_por = p_por, desfeito_em = now() where id = p_id;
end $f$;
$ddl$;
begin
  foreach s in array array['os','os_homolog'] loop
    execute replace(ddl, '__S__', s);
    execute format('revoke all on all functions in schema %I from public, anon, authenticated', s);
    execute format('grant execute on all functions in schema %I to os_app', s);
  end loop;
end $outer$;
