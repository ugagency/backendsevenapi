-- Seven OS, Specs 02 a 08: tabelas da operacao. So objetos novos, em os e os_homolog.
do $outer$
declare s text; ddl text := $ddl$
create table if not exists execucoes (
  id bigint generated always as identity primary key, conta_id bigint references contas_portal(id),
  origem text not null check (origem in ('agenda','manual','sincronizacao')), pedido_por uuid,
  iniciada_em timestamptz not null default now(), terminada_em timestamptz,
  situacao text not null default 'rodando' check (situacao in ('rodando','sucesso','falha')),
  linhas_lidas int, linhas_novas int, erro text);
create table if not exists linhas_evento (
  id bigint generated always as identity primary key, conta_id bigint not null references contas_portal(id),
  numero_evento text not null, numero_linha int not null, codigo text, descricao text,
  quantidade numeric, quantidade_texto text, unidade text, uf text, prazo date, prazo_texto text,
  marca text, marca_origem text, link_evento text, capturado_em timestamptz, data_inicio date,
  execucao_id bigint references execucoes(id), origem text not null default 'robo', origem_id bigint unique,
  unique (conta_id, numero_evento, numero_linha));
create index if not exists linhas_evento_evento_idx on linhas_evento (numero_evento);
create index if not exists linhas_evento_capt_idx on linhas_evento (capturado_em);
create table if not exists produtos (
  id bigint generated always as identity primary key, nome text not null, descricao text, fabricante text, referencia text,
  assinatura text not null unique, prefixo text, ref_chave text,
  nivel text not null check (nivel in ('certeza','provavel','conferir','nao_identificavel')),
  metodo text not null default 'estrutura', juntado_em bigint references produtos(id), criado_em timestamptz not null default now());
create index if not exists produtos_ref_chave_idx on produtos (ref_chave) where ref_chave is not null;
create table if not exists produto_codigos (
  produto_id bigint not null references produtos(id) on delete cascade, codigo text not null, motivo text,
  juntado_por uuid, juntado_em timestamptz not null default now(), primary key (produto_id, codigo));
create index if not exists produto_codigos_codigo_idx on produto_codigos (codigo);
create table if not exists linha_produto (
  linha_id bigint primary key references linhas_evento(id) on delete cascade,
  produto_id bigint not null references produtos(id), metodo text not null);
create index if not exists linha_produto_produto_idx on linha_produto (produto_id);
create table if not exists produto_juncoes (
  id bigint generated always as identity primary key, origem_id bigint not null references produtos(id),
  destino_id bigint not null references produtos(id), linhas bigint[] not null, motivo text,
  por uuid, em timestamptz not null default now(), desfeito_por uuid, desfeito_em timestamptz);
create table if not exists itens (
  id bigint generated always as identity primary key, produto_id bigint not null references produtos(id),
  numero_evento text not null, contas bigint[] not null default '{}', prazo date, quantidade_total numeric,
  linhas int not null default 1, unidade text, uf text,
  aba text not null default 'produto_novo' check (aba in ('pronto','falta_fornecedor','produto_novo','conferencia')),
  andamento text not null default 'a_cotar' check (andamento in ('a_cotar','em_cotacao','enviada','declinada','ganha','perdida')),
  responsavel_id uuid, capturado_em timestamptz, atualizado_em timestamptz not null default now(),
  unique (numero_evento, produto_id));
create index if not exists itens_prazo_idx on itens (prazo);
create index if not exists itens_produto_idx on itens (produto_id);
create table if not exists item_andamento (
  id bigint generated always as identity primary key, item_id bigint not null references itens(id) on delete cascade,
  de text, para text not null, motivo text, por uuid, em timestamptz not null default now());
create table if not exists declinados (
  id bigint generated always as identity primary key, item_id bigint not null references itens(id) on delete cascade,
  produto_id bigint not null references produtos(id), motivo text not null, por uuid, em timestamptz not null default now());
create index if not exists declinados_produto_idx on declinados (produto_id);
create table if not exists fornecedores (
  id bigint generated always as identity primary key, nome text not null, cidade text, uf text, contato text,
  telefone text, email text, observacao text, ativo boolean not null default true, criado_em timestamptz not null default now());
create unique index if not exists fornecedores_nome_uidx on fornecedores (upper(nome));
create table if not exists fornecedor_nomes (
  fornecedor_id bigint not null references fornecedores(id) on delete cascade, nome text not null,
  confirmado_por uuid, em timestamptz not null default now(), primary key (fornecedor_id, nome));
create table if not exists marcas (id bigint generated always as identity primary key, nome text not null unique, fabricante text);
create table if not exists fornecedor_marcas (
  fornecedor_id bigint not null references fornecedores(id) on delete cascade,
  marca_id bigint not null references marcas(id) on delete cascade, primary key (fornecedor_id, marca_id));
create table if not exists produto_fornecedores (
  produto_id bigint not null references produtos(id) on delete cascade,
  fornecedor_id bigint not null references fornecedores(id) on delete cascade,
  origem text not null check (origem in ('historia','marca','busca','equipe')),
  confirmado_por uuid, confirmado_em timestamptz not null default now(), primary key (produto_id, fornecedor_id));
create table if not exists sugestoes_fornecedor (
  id bigint generated always as identity primary key, produto_id bigint not null references produtos(id) on delete cascade,
  fornecedor_id bigint not null references fornecedores(id) on delete cascade, fonte text not null,
  situacao text not null default 'aberta' check (situacao in ('aberta','confirmada','descartada')),
  por uuid, em timestamptz not null default now(), unique (produto_id, fornecedor_id));
create table if not exists solicitacoes_cotacao (
  id bigint generated always as identity primary key, item_id bigint not null references itens(id) on delete cascade,
  fornecedor_id bigint not null references fornecedores(id), enviada_por uuid, enviada_em timestamptz not null default now(),
  assunto text, texto text not null, canal text not null default 'email');
create table if not exists cotacoes (
  id bigint generated always as identity primary key, item_id bigint not null references itens(id) on delete cascade,
  versao int not null, fornecedor_id bigint references fornecedores(id), conta_id bigint references contas_portal(id),
  criada_por uuid, criada_em timestamptz not null default now(),
  situacao text not null default 'rascunho' check (situacao in ('rascunho','pronta','enviada','substituida')),
  unique (item_id, versao));
create table if not exists cotacao_valores (
  cotacao_id bigint primary key references cotacoes(id) on delete cascade,
  preco_compra numeric, ipi numeric, frete numeric, preco_venda numeric);
create table if not exists alertas_preco (
  id bigint generated always as identity primary key, cotacao_id bigint not null references cotacoes(id) on delete cascade,
  tipo text not null check (tipo in ('parece_total','divergente','sem_preco')), valor_referencia numeric,
  situacao text not null default 'aberto' check (situacao in ('aberto','justificado')),
  justificativa text, resolvido_por uuid, resolvido_em timestamptz);
create table if not exists resultados (
  item_id bigint primary key references itens(id) on delete cascade, cotacao_id bigint references cotacoes(id),
  resultado text not null check (resultado in ('ganha','perdida')), pedido_venda text,
  registrado_por uuid, registrado_em timestamptz not null default now());
create table if not exists historico_importado (
  id bigint generated always as identity primary key, origem text not null, arquivo text, linha_original jsonb,
  produto_id bigint references produtos(id), codigo text, fornecedor_texto text, preco_compra numeric, ipi numeric,
  frete numeric, margem numeric, preco_venda numeric, status text, data_ref date);
create index if not exists historico_importado_produto_idx on historico_importado (produto_id);
create table if not exists conversas (id bigint generated always as identity primary key, user_id uuid not null, titulo text, criada_em timestamptz not null default now());
create table if not exists mensagens (
  id bigint generated always as identity primary key, conversa_id bigint not null references conversas(id) on delete cascade,
  papel text not null check (papel in ('usuario','assistente')), texto text not null, consulta_usada text, parametros jsonb,
  criada_em timestamptz not null default now());
$ddl$;
begin
  foreach s in array array['os','os_homolog'] loop
    execute format('set local search_path = %I', s);
    execute ddl;
    execute format('revoke all on all tables in schema %I from public, anon, authenticated', s);
    execute format('grant select, insert, update, delete on all tables in schema %I to os_app', s);
    execute format('grant usage on all sequences in schema %I to os_app', s);
    execute format('revoke update, delete, truncate on %I.auditoria from os_app', s);
  end loop;
end $outer$;
