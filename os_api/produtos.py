"""Spec 03: conferencia manual, juncao de codigos da mesma peca e desfazer."""
from typing import Optional

import psycopg
from fastapi import APIRouter, Depends, HTTPException, Query
from pydantic import BaseModel

from . import db
from .acesso import UsuarioOS, exigir
from .auditoria import auditar
from .comum import lista, origem, um

SQL_PRODUTO = """select p.id, p.nome, p.descricao, p.fabricante, p.referencia, p.prefixo, p.nivel, p.metodo,
  (select array_agg(codigo order by codigo) from produto_codigos pc where pc.produto_id = p.id) codigos,
  (select count(*) from itens i where i.produto_id = p.id) ocorrencias from produtos p """


class Decisao(BaseModel):
    acao: str  # associar | novo_produto | nao_identificavel
    destino_id: Optional[int] = None
    nome: Optional[str] = None
    fabricante: Optional[str] = None
    referencia: Optional[str] = None
    motivo: Optional[str] = None


class Juntar(BaseModel):
    origem_id: int
    motivo: Optional[str] = None


def rotas_produtos() -> APIRouter:
    r = APIRouter(prefix="/os")

    @r.get("/produtos")
    def buscar(busca: str = Query(..., min_length=2), limite: int = Query(20, ge=1, le=50),
               u: UsuarioOS = Depends(exigir("fila.ver"))):
        b = f"%{busca.strip()}%"
        with db.conexao() as conn:
            return lista(conn, SQL_PRODUTO + """where p.juntado_em is null and (p.nome ilike %s or p.fabricante ilike %s
                or p.referencia ilike %s or p.descricao ilike %s
                or exists (select 1 from produto_codigos pc where pc.produto_id = p.id and pc.codigo like %s))
                order by (p.nivel = 'certeza') desc, ocorrencias desc limit %s""", (b, b, b, b, busca.strip() + "%", limite))

    @r.get("/conferencia")
    def conferencia(pagina: int = Query(1, ge=1), busca: Optional[str] = None,
                    u: UsuarioOS = Depends(exigir("fila.ver"))):
        cond, par = "", [list(u.contas)]
        if busca and busca.strip():
            cond = "and (p.nome ilike %s or p.descricao ilike %s)"; par += [f"%{busca.strip()}%"] * 2
        with db.conexao() as conn:
            base = f"""from produtos p join lateral (
                  select count(*) abertos, min(i.prazo) proximo_prazo, min(i.id) item_id from itens i
                  where i.produto_id = p.id and i.contas && %s::bigint[] and i.andamento in ('a_cotar','em_cotacao')
                    and (i.prazo is null or i.prazo >= current_date)) a on a.abertos > 0
                where p.nivel = 'conferir' and p.juntado_em is null {cond}"""
            total = um(conn, "select count(*) n " + base, par)["n"]
            rows = lista(conn, """select p.id, p.nome, p.descricao, p.fabricante, p.prefixo, a.abertos, a.proximo_prazo, a.item_id,
                (select array_agg(codigo order by codigo) from produto_codigos pc where pc.produto_id = p.id) codigos,
                (select count(*) from itens i where i.produto_id = p.id) ocorrencias """ + base +
                         " order by a.proximo_prazo asc nulls last, p.id limit 30 offset %s", par + [(pagina - 1) * 30])
        return {"total": total, "pagina": pagina, "produtos": rows}

    @r.post("/conferencia/{produto_id}")
    def decidir(produto_id: int, body: Decisao, u: UsuarioOS = Depends(exigir("identificacao.resolver"))):
        with db.conexao() as conn:
            p = um(conn, "select id, nome, nivel, metodo, fabricante, referencia from produtos where id = %s and juntado_em is null for update", (produto_id,))
            if not p:
                raise HTTPException(status_code=404, detail="Produto nao encontrado.")
            if body.acao == "associar":
                if not body.destino_id:
                    raise HTTPException(status_code=400, detail="Informe o produto de destino.")
                try:
                    jid = conn.execute("select juntar_produto(%s, %s, %s, %s)",
                                       (produto_id, body.destino_id, u.id, body.motivo or "conferencia manual")).fetchone()[0]
                except psycopg.errors.RaiseException as e:
                    raise HTTPException(status_code=400, detail=str(e).splitlines()[0])
                depois = {"acao": "associar", "destino_id": body.destino_id, "juncao_id": jid}
            elif body.acao == "novo_produto":
                conn.execute("""update produtos set nivel = 'certeza', metodo = 'pessoa', nome = coalesce(nullif(%s, ''), nome),
                                fabricante = coalesce(nullif(%s, ''), fabricante), referencia = coalesce(nullif(%s, ''), referencia)
                                where id = %s""", ((body.nome or "").strip(), (body.fabricante or "").strip(),
                                                   (body.referencia or "").strip(), produto_id))
                conn.execute("select atualizar_abas(%s)", ([produto_id],))
                depois = {"acao": "novo_produto", "nivel": "certeza"}
            elif body.acao == "nao_identificavel":
                conn.execute("update produtos set nivel = 'nao_identificavel', metodo = 'pessoa' where id = %s", (produto_id,))
                conn.execute("select atualizar_abas(%s)", ([produto_id],))
                depois = {"acao": "nao_identificavel"}
            else:
                raise HTTPException(status_code=400, detail="Acao invalida.")
            auditar(conn, origem=origem(u), entidade="produto", entidade_id=produto_id, acao="alterar", user_id=u.id,
                    antes={"nivel": p["nivel"], "nome": p["nome"]}, depois=depois, justificativa=body.motivo)
        return depois

    @r.post("/produtos/{produto_id}/juntar")
    def juntar(produto_id: int, body: Juntar, u: UsuarioOS = Depends(exigir("identificacao.resolver"))):
        """Junta outro produto (origem) neste. Usado na ficha, em 'mesma peca, codigos diferentes'."""
        with db.conexao() as conn:
            try:
                jid = conn.execute("select juntar_produto(%s, %s, %s, %s)",
                                   (body.origem_id, produto_id, u.id, body.motivo or "mesma referencia")).fetchone()[0]
            except psycopg.errors.RaiseException as e:
                raise HTTPException(status_code=400, detail=str(e).splitlines()[0])
            auditar(conn, origem=origem(u), entidade="produto", entidade_id=produto_id, acao="alterar", user_id=u.id,
                    depois={"juntou": body.origem_id, "juncao_id": jid}, justificativa=body.motivo)
        return {"juncao_id": jid}

    @r.post("/juncoes/{juncao_id}/desfazer")
    def desfazer(juncao_id: int, u: UsuarioOS = Depends(exigir("identificacao.resolver"))):
        with db.conexao() as conn:
            try:
                conn.execute("select desfazer_juncao(%s, %s)", (juncao_id, u.id))
            except psycopg.errors.RaiseException as e:
                raise HTTPException(status_code=400, detail=str(e).splitlines()[0])
            auditar(conn, origem=origem(u), entidade="juncao", entidade_id=juncao_id, acao="excluir", user_id=u.id)
        return {"desfeita": juncao_id}

    return r
