"""Spec 05: cadastro unico de fornecedores, marcas, vinculo com produto, sugestoes e solicitacao de cotacao."""
import re
from typing import List, Optional

from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel

from . import db
from .acesso import UsuarioOS, exigir
from .auditoria import auditar
from .comum import exigir_conta, lista, num, origem, pode, um


class Fornecedor(BaseModel):
    nome: Optional[str] = None
    cidade: Optional[str] = None
    uf: Optional[str] = None
    contato: Optional[str] = None
    telefone: Optional[str] = None
    email: Optional[str] = None
    observacao: Optional[str] = None
    ativo: Optional[bool] = None
    marcas: Optional[List[str]] = None


class JuntarFornecedor(BaseModel):
    manter_id: int
    juntar_id: int


class Vinculo(BaseModel):
    fornecedor_id: int
    origem: str = "equipe"


class Solicitacao(BaseModel):
    item_id: int
    fornecedor_id: int
    assunto: Optional[str] = None
    texto: Optional[str] = None
    registrar: bool = False


CAMPOS = ("nome", "cidade", "uf", "contato", "telefone", "email", "observacao", "ativo")


def _marcas(conn, fid, nomes):
    conn.execute("delete from fornecedor_marcas where fornecedor_id = %s", (fid,))
    for n in sorted({m.strip().upper() for m in nomes if m and m.strip()}):
        mid = conn.execute("insert into marcas (nome) values (%s) on conflict (nome) do update set nome = excluded.nome returning id",
                           (n,)).fetchone()[0]
        conn.execute("insert into fornecedor_marcas (fornecedor_id, marca_id) values (%s, %s) on conflict do nothing", (fid, mid))


def _fornecedor(conn, fid):
    f = um(conn, """select f.*, coalesce((select array_agg(m.nome order by m.nome) from fornecedor_marcas fm join marcas m on m.id = fm.marca_id
                    where fm.fornecedor_id = f.id), '{}') marcas,
                    coalesce((select array_agg(n.nome order by n.nome) from fornecedor_nomes n where n.fornecedor_id = f.id), '{}') outros_nomes
                    from fornecedores f where f.id = %s""", (fid,))
    return f


def formatos_de_valor(v):
    """Formas comuns de escrever um valor em reais, para procurar no texto do e-mail."""
    f = float(v)
    br = f"{f:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return {f"{f:.2f}", f"{f:.2f}".replace(".", ","), br}


def texto_tem_preco(texto: str, precos_compra) -> Optional[str]:
    """Bloqueio da Spec 05: e-mail a fornecedor nunca leva preco. Devolve o motivo, ou None se estiver limpo."""
    if re.search(r"R\$\s*\d", texto, flags=re.I):
        return "O texto tem um valor em reais."
    for p in precos_compra:
        if p is None or float(p) == 0:
            continue
        for forma in formatos_de_valor(p):
            if re.search(r"(?<![\d.,])" + re.escape(forma) + r"(?![\d])", texto):
                return "O texto contem um preco de compra registrado para este produto."
    return None


def rascunho_email(item, produto, codigos, fornecedor, remetente):
    ref = f"{produto['fabricante'] or ''} {produto['referencia'] or ''}".strip()
    qtd = num(item["quantidade_total"])
    qtd_txt = (f"{qtd:g}" if qtd is not None else "a confirmar") + (f" {item['unidade']}" if item["unidade"] else "")
    linhas = [f"Prezados{', ' + fornecedor['contato'] if fornecedor.get('contato') else ''},", "",
              "Solicitamos cotacao para o item abaixo:", "", f"Item: {produto['nome']}"]
    if ref:
        linhas.append(f"Fabricante / referencia: {ref}")
    if produto.get("descricao"):
        linhas.append(f"Descricao: {produto['descricao'][:600]}")
    linhas += [f"Quantidade: {qtd_txt}", f"Local de entrega (UF): {item['uf'] or 'a confirmar'}", "",
               "Por favor informar preco unitario, IPI, frete, prazo de entrega e validade da proposta."]
    if item["prazo"]:
        linhas.append(f"Precisamos do retorno ate {item['prazo'].strftime('%d/%m/%Y')}.")
    linhas += ["", "Obrigado,", remetente]
    return f"Cotacao: {produto['nome'][:80]}", "\n".join(linhas)


def rotas_fornecedores() -> APIRouter:
    r = APIRouter(prefix="/os")

    @r.get("/fornecedores")
    def listar(busca: Optional[str] = None, marca: Optional[str] = None, inativos: bool = False,
               u: UsuarioOS = Depends(exigir("fila.ver"))):
        cond, par = ["true"], []
        if not inativos:
            cond.append("f.ativo")
        if busca and busca.strip():
            b = f"%{busca.strip()}%"
            cond.append("(f.nome ilike %s or f.cidade ilike %s or exists (select 1 from fornecedor_nomes n where n.fornecedor_id = f.id and n.nome ilike %s))")
            par += [b, b, b]
        if marca and marca.strip():
            cond.append("exists (select 1 from fornecedor_marcas fm join marcas m on m.id = fm.marca_id where fm.fornecedor_id = f.id and m.nome = %s)")
            par.append(marca.strip().upper())
        with db.conexao() as conn:
            rows = lista(conn, f"""select f.id, f.nome, f.cidade, f.uf, f.contato, f.telefone, f.email, f.ativo,
                coalesce((select array_agg(m.nome order by m.nome) from fornecedor_marcas fm join marcas m on m.id = fm.marca_id where fm.fornecedor_id = f.id), '{{}}') marcas,
                (select count(*) from produto_fornecedores pf where pf.fornecedor_id = f.id) produtos,
                (select count(*) from cotacoes c join resultados rs on rs.cotacao_id = c.id where c.fornecedor_id = f.id) com_resultado,
                (select count(*) from cotacoes c join resultados rs on rs.cotacao_id = c.id where c.fornecedor_id = f.id and rs.resultado = 'ganha') ganhas
                from fornecedores f where {' and '.join(cond)} order by f.nome limit 500""", par)
            marcas = [m["nome"] for m in lista(conn, "select nome from marcas order by nome")]
        return {"fornecedores": rows, "marcas": marcas}

    @r.post("/fornecedores", status_code=201)
    def criar(body: Fornecedor, u: UsuarioOS = Depends(exigir("fornecedor.vincular"))):
        nome = (body.nome or "").strip()
        if not nome:
            raise HTTPException(status_code=400, detail="Nome obrigatorio.")
        with db.conexao() as conn:
            igual = um(conn, """select f.id, f.nome from fornecedores f where upper(f.nome) = upper(%s)
                                or exists (select 1 from fornecedor_nomes n where n.fornecedor_id = f.id and upper(n.nome) = upper(%s)) limit 1""",
                       (nome, nome))
            if igual:
                raise HTTPException(status_code=409, detail=f"Ja existe: {igual['nome']} (id {igual['id']}).")
            fid = conn.execute("""insert into fornecedores (nome, cidade, uf, contato, telefone, email, observacao)
                                  values (%s, %s, %s, %s, %s, %s, %s) returning id""",
                               (nome, body.cidade, (body.uf or None) and body.uf.upper(), body.contato, body.telefone,
                                body.email, body.observacao)).fetchone()[0]
            if body.marcas is not None:
                _marcas(conn, fid, body.marcas)
            depois = _fornecedor(conn, fid)
            auditar(conn, origem=origem(u), entidade="fornecedor", entidade_id=fid, acao="criar", user_id=u.id,
                    depois={k: depois[k] for k in CAMPOS + ("marcas",)})
        return depois

    @r.get("/fornecedores/{fid}")
    def detalhe(fid: int, u: UsuarioOS = Depends(exigir("fila.ver"))):
        with db.conexao() as conn:
            f = _fornecedor(conn, fid)
            if not f:
                raise HTTPException(status_code=404, detail="Fornecedor nao encontrado.")
            f["produtos"] = lista(conn, """select p.id, p.nome, p.fabricante, p.referencia, pf.origem, pf.confirmado_em
                                           from produto_fornecedores pf join produtos p on p.id = pf.produto_id
                                           where pf.fornecedor_id = %s order by pf.confirmado_em desc limit 200""", (fid,))
        return f

    @r.patch("/fornecedores/{fid}")
    def alterar(fid: int, body: Fornecedor, u: UsuarioOS = Depends(exigir("fornecedor.vincular"))):
        dados = body.model_dump(exclude_unset=True)
        marcas = dados.pop("marcas", None)
        if "nome" in dados and not (dados["nome"] or "").strip():
            raise HTTPException(status_code=400, detail="Nome obrigatorio.")
        with db.conexao() as conn:
            antes = _fornecedor(conn, fid)
            if not antes:
                raise HTTPException(status_code=404, detail="Fornecedor nao encontrado.")
            for k, v in dados.items():
                if k in CAMPOS:
                    conn.execute(f"update fornecedores set {k} = %s where id = %s", (v.strip() if isinstance(v, str) else v, fid))
            if marcas is not None:
                _marcas(conn, fid, marcas)
            depois = _fornecedor(conn, fid)
            auditar(conn, origem=origem(u), entidade="fornecedor", entidade_id=fid, acao="alterar", user_id=u.id,
                    antes={k: antes[k] for k in CAMPOS + ("marcas",)}, depois={k: depois[k] for k in CAMPOS + ("marcas",)})
        return depois

    @r.post("/fornecedores/juntar")
    def juntar(body: JuntarFornecedor, u: UsuarioOS = Depends(exigir("fornecedor.vincular"))):
        """Duas escritas do mesmo fornecedor viram um cadastro. O nome juntado fica guardado como variacao."""
        if body.manter_id == body.juntar_id:
            raise HTTPException(status_code=400, detail="Escolha dois cadastros diferentes.")
        with db.conexao() as conn:
            a, b = _fornecedor(conn, body.manter_id), _fornecedor(conn, body.juntar_id)
            if not a or not b:
                raise HTTPException(status_code=404, detail="Fornecedor nao encontrado.")
            k, j = body.manter_id, body.juntar_id
            conn.execute("insert into fornecedor_nomes (fornecedor_id, nome, confirmado_por) values (%s, %s, %s) on conflict do nothing", (k, b["nome"], u.id))
            conn.execute("insert into fornecedor_nomes (fornecedor_id, nome, confirmado_por) select %s, nome, %s from fornecedor_nomes where fornecedor_id = %s on conflict do nothing", (k, u.id, j))
            conn.execute("insert into fornecedor_marcas (fornecedor_id, marca_id) select %s, marca_id from fornecedor_marcas where fornecedor_id = %s on conflict do nothing", (k, j))
            conn.execute("insert into produto_fornecedores (produto_id, fornecedor_id, origem, confirmado_por, confirmado_em) "
                         "select produto_id, %s, origem, confirmado_por, confirmado_em from produto_fornecedores where fornecedor_id = %s on conflict do nothing", (k, j))
            conn.execute("update cotacoes set fornecedor_id = %s where fornecedor_id = %s", (k, j))
            conn.execute("update solicitacoes_cotacao set fornecedor_id = %s where fornecedor_id = %s", (k, j))
            conn.execute("delete from sugestoes_fornecedor where fornecedor_id = %s", (j,))
            conn.execute("delete from fornecedores where id = %s", (j,))
            auditar(conn, origem=origem(u), entidade="fornecedor", entidade_id=k, acao="alterar", user_id=u.id,
                    antes={"juntado": {"id": j, "nome": b["nome"]}}, depois={"mantido": a["nome"]})
            return _fornecedor(conn, k)

    # ---------- vinculo com produto ----------
    @r.get("/produtos/{pid}/fornecedores")
    def do_produto(pid: int, u: UsuarioOS = Depends(exigir("fila.ver"))):
        with db.conexao() as conn:
            p = um(conn, "select id, fabricante from produtos where id = %s", (pid,))
            if not p:
                raise HTTPException(status_code=404, detail="Produto nao encontrado.")
            vinc = lista(conn, """
                select f.id, f.nome, f.cidade, f.uf, f.contato, f.telefone, f.email, pf.origem, pf.confirmado_em,
                  up.preco_compra ultimo_preco, up.criada_em ultimo_preco_em,
                  (select count(*) from cotacoes c join resultados rs on rs.cotacao_id = c.id join itens i on i.id = c.item_id
                   where c.fornecedor_id = f.id and i.produto_id = %s and rs.resultado = 'ganha') ganhas
                from produto_fornecedores pf join fornecedores f on f.id = pf.fornecedor_id
                left join lateral (select cv.preco_compra, c.criada_em from cotacoes c join cotacao_valores cv on cv.cotacao_id = c.id
                    join itens i on i.id = c.item_id where c.fornecedor_id = f.id and i.produto_id = %s and cv.preco_compra is not null
                    and c.situacao <> 'substituida' order by c.criada_em desc limit 1) up on true
                where pf.produto_id = %s order by f.nome""", (pid, pid, pid))
            sug = lista(conn, """
                select f.id, f.nome, f.cidade, f.uf, 'Atende a marca ' || m.nome as fonte
                from marcas m join fornecedor_marcas fm on fm.marca_id = m.id join fornecedores f on f.id = fm.fornecedor_id and f.ativo
                where m.nome = upper(%s)
                  and not exists (select 1 from produto_fornecedores pf where pf.produto_id = %s and pf.fornecedor_id = f.id)
                  and not exists (select 1 from sugestoes_fornecedor s where s.produto_id = %s and s.fornecedor_id = f.id and s.situacao = 'descartada')
                order by f.nome limit 20""", (p["fabricante"] or "", pid, pid)) if p["fabricante"] else []
        for v in vinc:
            preco = v.pop("ultimo_preco")
            if pode(u, "preco_compra.ver"):
                v["ultimo_preco"] = num(preco)
            else:
                v.pop("ultimo_preco_em", None)
        return {"vinculados": vinc, "sugestoes": sug}

    @r.post("/produtos/{pid}/fornecedores", status_code=201)
    def vincular(pid: int, body: Vinculo, u: UsuarioOS = Depends(exigir("fornecedor.vincular"))):
        if body.origem not in ("historia", "marca", "busca", "equipe"):
            raise HTTPException(status_code=400, detail="Origem invalida.")
        with db.conexao() as conn:
            if not um(conn, "select 1 x from produtos where id = %s", (pid,)) or not um(conn, "select 1 x from fornecedores where id = %s", (body.fornecedor_id,)):
                raise HTTPException(status_code=404, detail="Produto ou fornecedor nao encontrado.")
            conn.execute("insert into produto_fornecedores (produto_id, fornecedor_id, origem, confirmado_por) values (%s, %s, %s, %s) on conflict do nothing",
                         (pid, body.fornecedor_id, body.origem, u.id))
            conn.execute("""insert into sugestoes_fornecedor (produto_id, fornecedor_id, fonte, situacao, por) values (%s, %s, %s, 'confirmada', %s)
                            on conflict (produto_id, fornecedor_id) do update set situacao = 'confirmada', por = excluded.por, em = now()""",
                         (pid, body.fornecedor_id, body.origem, u.id))
            conn.execute("select atualizar_abas(%s)", ([pid],))
            auditar(conn, origem=origem(u), entidade="produto_fornecedor", entidade_id=f"{pid}:{body.fornecedor_id}", acao="criar",
                    user_id=u.id, depois={"produto_id": pid, "fornecedor_id": body.fornecedor_id, "origem": body.origem})
        return {"produto_id": pid, "fornecedor_id": body.fornecedor_id}

    @r.delete("/produtos/{pid}/fornecedores/{fid}")
    def desvincular(pid: int, fid: int, u: UsuarioOS = Depends(exigir("fornecedor.vincular"))):
        with db.conexao() as conn:
            n = conn.execute("delete from produto_fornecedores where produto_id = %s and fornecedor_id = %s", (pid, fid)).rowcount
            if not n:
                raise HTTPException(status_code=404, detail="Vinculo nao encontrado.")
            conn.execute("select atualizar_abas(%s)", ([pid],))
            auditar(conn, origem=origem(u), entidade="produto_fornecedor", entidade_id=f"{pid}:{fid}", acao="excluir", user_id=u.id,
                    antes={"produto_id": pid, "fornecedor_id": fid})
        return {"removido": True}

    @r.post("/produtos/{pid}/sugestoes/{fid}/descartar")
    def descartar(pid: int, fid: int, u: UsuarioOS = Depends(exigir("fornecedor.vincular"))):
        with db.conexao() as conn:
            conn.execute("""insert into sugestoes_fornecedor (produto_id, fornecedor_id, fonte, situacao, por) values (%s, %s, 'marca', 'descartada', %s)
                            on conflict (produto_id, fornecedor_id) do update set situacao = 'descartada', por = excluded.por, em = now()""",
                         (pid, fid, u.id))
            auditar(conn, origem=origem(u), entidade="sugestao_fornecedor", entidade_id=f"{pid}:{fid}", acao="alterar", user_id=u.id,
                    depois={"situacao": "descartada"})
        return {"descartada": True}

    # ---------- solicitacao de cotacao por e-mail ----------
    @r.post("/solicitacoes")
    def solicitar(body: Solicitacao, u: UsuarioOS = Depends(exigir("fornecedor.vincular"))):
        """Sem `registrar`, devolve o rascunho. Com `registrar`, valida que nao ha preco no texto e guarda na ficha.
        O envio em si sai da caixa de e-mail da pessoa (o front abre o programa de e-mail); nao ha servico de envio contratado."""
        with db.conexao() as conn:
            it = um(conn, "select * from itens where id = %s", (body.item_id,))
            if not it:
                raise HTTPException(status_code=404, detail="Item nao encontrado.")
            exigir_conta(u, it["contas"])
            f = um(conn, "select id, nome, contato, email from fornecedores where id = %s", (body.fornecedor_id,))
            if not f:
                raise HTTPException(status_code=404, detail="Fornecedor nao encontrado.")
            p = um(conn, "select id, nome, descricao, fabricante, referencia from produtos where id = %s", (it["produto_id"],))
            assunto, texto = rascunho_email(it, p, [], f, u.nome or u.email)
            if not body.registrar:
                return {"assunto": assunto, "texto": texto, "para": f["email"], "registrada": False}
            assunto, texto = (body.assunto or assunto).strip(), (body.texto or texto).strip()
            precos = [x["preco_compra"] for x in lista(conn, """
                select cv.preco_compra from cotacao_valores cv join cotacoes c on c.id = cv.cotacao_id join itens i on i.id = c.item_id where i.produto_id = %s
                union all select preco_compra from historico_importado where produto_id = %s""", (p["id"], p["id"]))]
            motivo = texto_tem_preco(assunto + "\n" + texto, precos)
            if motivo:
                raise HTTPException(status_code=400, detail=f"Envio bloqueado. {motivo} Retire o valor e tente de novo.")
            sid = conn.execute("insert into solicitacoes_cotacao (item_id, fornecedor_id, enviada_por, assunto, texto) values (%s, %s, %s, %s, %s) returning id",
                               (body.item_id, body.fornecedor_id, u.id, assunto, texto)).fetchone()[0]
            auditar(conn, origem=origem(u), entidade="solicitacao_cotacao", entidade_id=sid, acao="criar", user_id=u.id,
                    depois={"item_id": body.item_id, "fornecedor_id": body.fornecedor_id})
        return {"id": sid, "assunto": assunto, "texto": texto, "para": f["email"], "registrada": True}

    @r.get("/itens/{item_id}/solicitacoes")
    def solicitacoes(item_id: int, u: UsuarioOS = Depends(exigir("fila.ver"))):
        with db.conexao() as conn:
            it = um(conn, "select contas from itens where id = %s", (item_id,))
            if not it:
                raise HTTPException(status_code=404, detail="Item nao encontrado.")
            exigir_conta(u, it["contas"])
            return lista(conn, """select s.id, s.enviada_em, s.assunto, f.nome fornecedor, coalesce(nullif(us.nome, ''), us.email) por
                                  from solicitacoes_cotacao s join fornecedores f on f.id = s.fornecedor_id
                                  left join usuarios us on us.user_id = s.enviada_por where s.item_id = %s order by s.enviada_em desc""", (item_id,))

    return r
