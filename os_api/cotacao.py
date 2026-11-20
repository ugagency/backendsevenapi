"""Spec 06: cotacao em versoes, alertas antes do envio e resultado. O envio ao Coupa continua manual."""
from typing import Optional

from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel

from . import db
from .acesso import UsuarioOS, exigir
from .auditoria import auditar
from .comum import exigir_conta, lista, num, origem, pode, um, valores_visiveis

LIMITE_PADRAO_PCT = 30   # ate a Seven definir (pendencia da Spec 06)
FATOR_TOTAL = 0.75       # gatilho do alerta "parece total", como no prototipo; validar com casos reais


class NovaCotacao(BaseModel):
    fornecedor_id: Optional[int] = None
    conta_id: Optional[int] = None
    preco_compra: Optional[float] = None
    ipi: Optional[float] = None
    frete: Optional[float] = None
    preco_venda: Optional[float] = None


class Justificativa(BaseModel):
    justificativa: str


class Resultado(BaseModel):
    resultado: str
    pedido_venda: Optional[str] = None


def calcular_alertas(preco_compra, quantidade, referencia, limite_pct):
    """Regras puras dos alertas. `referencia` e o ultimo preco de compra unitario do mesmo produto (so Certeza)."""
    if preco_compra is None or preco_compra <= 0:
        return [("sem_preco", None)]
    if referencia is None or referencia <= 0:
        return []
    qtd = float(quantidade or 0)
    if qtd > 1 and preco_compra >= FATOR_TOTAL * referencia * qtd:
        return [("parece_total", referencia)]
    if abs(preco_compra - referencia) / referencia * 100 > limite_pct:
        return [("divergente", referencia)]
    return []


def _cotacoes_do_item(conn, u, item_id):
    rows = lista(conn, """
        select c.id, c.versao, c.situacao, c.criada_em, c.fornecedor_id, f.nome fornecedor, c.conta_id, ct.nome conta,
               coalesce(nullif(us.nome, ''), us.email) por, cv.preco_compra, cv.ipi, cv.frete, cv.preco_venda
        from cotacoes c left join cotacao_valores cv on cv.cotacao_id = c.id left join fornecedores f on f.id = c.fornecedor_id
        left join contas_portal ct on ct.id = c.conta_id left join usuarios us on us.user_id = c.criada_por
        where c.item_id = %s order by c.versao desc""", (item_id,))
    alertas = lista(conn, """select a.id, a.cotacao_id, a.tipo, a.valor_referencia, a.situacao, a.justificativa, a.resolvido_em
                             from alertas_preco a join cotacoes c on c.id = a.cotacao_id where c.item_id = %s order by a.id""", (item_id,))
    for a in alertas:
        ref = a.pop("valor_referencia")
        if pode(u, "preco_compra.ver"):
            a["valor_referencia"] = num(ref)
    for c in rows:
        c["valores"] = valores_visiveis(u, c)
        for k in ("preco_compra", "ipi", "frete", "preco_venda"):
            c.pop(k, None)
        c["alertas"] = [a for a in alertas if a["cotacao_id"] == c["id"]]
    return rows


def rotas_cotacao() -> APIRouter:
    r = APIRouter(prefix="/os")

    @r.get("/itens/{item_id}/cotacoes")
    def listar(item_id: int, u: UsuarioOS = Depends(exigir("fila.ver"))):
        with db.conexao() as conn:
            it = um(conn, "select contas from itens where id = %s", (item_id,))
            if not it:
                raise HTTPException(status_code=404, detail="Item nao encontrado.")
            exigir_conta(u, it["contas"])
            res = um(conn, "select resultado, pedido_venda, registrado_em from resultados where item_id = %s", (item_id,))
            return {"cotacoes": _cotacoes_do_item(conn, u, item_id), "resultado": res}

    @r.post("/itens/{item_id}/cotacoes", status_code=201)
    def lancar(item_id: int, body: NovaCotacao, u: UsuarioOS = Depends(exigir("cotacao.lancar"))):
        for campo in ("preco_compra", "ipi", "frete", "preco_venda"):
            v = getattr(body, campo)
            if v is not None and v < 0:
                raise HTTPException(status_code=400, detail=f"{campo} nao pode ser negativo.")
        with db.conexao() as conn:
            it = um(conn, "select i.*, p.nivel from itens i join produtos p on p.id = i.produto_id where i.id = %s for update of i", (item_id,))
            if not it:
                raise HTTPException(status_code=404, detail="Item nao encontrado.")
            exigir_conta(u, it["contas"])
            if it["andamento"] not in ("a_cotar", "em_cotacao", "enviada"):
                raise HTTPException(status_code=400, detail=f"Item em {it['andamento']} nao recebe nova cotacao.")
            conta = body.conta_id or (it["contas"][0] if len(it["contas"]) == 1 else None)
            if conta is not None and (conta not in it["contas"] or conta not in u.contas):
                raise HTTPException(status_code=400, detail="CNPJ invalido para este item.")
            if body.fornecedor_id and not um(conn, "select 1 x from fornecedores where id = %s", (body.fornecedor_id,)):
                raise HTTPException(status_code=404, detail="Fornecedor nao encontrado.")
            versao = um(conn, "select coalesce(max(versao), 0) + 1 v from cotacoes where item_id = %s", (item_id,))["v"]
            conn.execute("update cotacoes set situacao = 'substituida' where item_id = %s and situacao <> 'substituida'", (item_id,))
            cid = conn.execute("insert into cotacoes (item_id, versao, fornecedor_id, conta_id, criada_por) values (%s, %s, %s, %s, %s) returning id",
                               (item_id, versao, body.fornecedor_id, conta, u.id)).fetchone()[0]
            conn.execute("insert into cotacao_valores (cotacao_id, preco_compra, ipi, frete, preco_venda) values (%s, %s, %s, %s, %s)",
                         (cid, body.preco_compra, body.ipi, body.frete, body.preco_venda))
            # Historico e alerta so comparam produtos em Certeza (Spec 03).
            ref = None
            if it["nivel"] == "certeza":
                x = um(conn, """
                    select preco_compra from (
                      select cv.preco_compra, c.criada_em quando from cotacoes c join cotacao_valores cv on cv.cotacao_id = c.id
                        join itens i on i.id = c.item_id where i.produto_id = %s and i.id <> %s and cv.preco_compra > 0 and c.situacao <> 'substituida'
                      union all
                      select h.preco_compra, h.data_ref::timestamptz from historico_importado h where h.produto_id = %s and h.preco_compra > 0
                    ) q order by quando desc nulls last limit 1""", (it["produto_id"], item_id, it["produto_id"]))
                ref = float(x["preco_compra"]) if x else None
            lim = um(conn, "select valor from parametros where chave = 'alertas.divergencia_pct'")
            limite = float(lim["valor"]) if lim else LIMITE_PADRAO_PCT
            alertas = calcular_alertas(body.preco_compra, it["quantidade_total"], ref, limite)
            for tipo, vref in alertas:
                conn.execute("insert into alertas_preco (cotacao_id, tipo, valor_referencia) values (%s, %s, %s)", (cid, tipo, vref))
            conn.execute("update cotacoes set situacao = %s where id = %s", ("rascunho" if alertas else "pronta", cid))
            if it["andamento"] == "a_cotar":
                conn.execute("update itens set andamento = 'em_cotacao', responsavel_id = %s, atualizado_em = now() where id = %s", (u.id, item_id))
                conn.execute("insert into item_andamento (item_id, de, para, por) values (%s, 'a_cotar', 'em_cotacao', %s)", (item_id, u.id))
            if body.fornecedor_id:
                conn.execute("insert into produto_fornecedores (produto_id, fornecedor_id, origem, confirmado_por) values (%s, %s, 'historia', %s) on conflict do nothing",
                             (it["produto_id"], body.fornecedor_id, u.id))
                conn.execute("select atualizar_abas(%s)", ([it["produto_id"]],))
            auditar(conn, origem=origem(u), entidade="cotacao", entidade_id=cid, acao="criar", user_id=u.id,
                    depois={"item_id": item_id, "versao": versao, "fornecedor_id": body.fornecedor_id, "conta_id": conta,
                            "preco_compra": body.preco_compra, "ipi": body.ipi, "frete": body.frete, "preco_venda": body.preco_venda,
                            "alertas": [a[0] for a in alertas]})
            return {"cotacoes": _cotacoes_do_item(conn, u, item_id)}

    @r.post("/alertas/{alerta_id}/justificar")
    def justificar(alerta_id: int, body: Justificativa, u: UsuarioOS = Depends(exigir("cotacao.lancar"))):
        texto = body.justificativa.strip()
        if len(texto) < 5:
            raise HTTPException(status_code=400, detail="Escreva a justificativa.")
        with db.conexao() as conn:
            a = um(conn, """select a.id, a.tipo, a.situacao, a.cotacao_id, c.item_id, c.situacao cot_situacao, i.contas
                            from alertas_preco a join cotacoes c on c.id = a.cotacao_id join itens i on i.id = c.item_id
                            where a.id = %s for update of a""", (alerta_id,))
            if not a:
                raise HTTPException(status_code=404, detail="Alerta nao encontrado.")
            exigir_conta(u, a["contas"])
            if a["tipo"] == "sem_preco":
                raise HTTPException(status_code=400, detail="Preco em branco nao se justifica: lance uma nova versao com o preco.")
            if a["situacao"] != "aberto":
                raise HTTPException(status_code=400, detail="Alerta ja resolvido.")
            conn.execute("update alertas_preco set situacao = 'justificado', justificativa = %s, resolvido_por = %s, resolvido_em = now() where id = %s",
                         (texto, u.id, alerta_id))
            if a["cot_situacao"] == "rascunho" and not um(conn, "select 1 x from alertas_preco where cotacao_id = %s and situacao = 'aberto'", (a["cotacao_id"],)):
                conn.execute("update cotacoes set situacao = 'pronta' where id = %s", (a["cotacao_id"],))
            auditar(conn, origem=origem(u), entidade="alerta_preco", entidade_id=alerta_id, acao="alterar", user_id=u.id,
                    antes={"situacao": "aberto", "tipo": a["tipo"]}, depois={"situacao": "justificado"}, justificativa=texto)
            return {"cotacoes": _cotacoes_do_item(conn, u, a["item_id"])}

    @r.post("/itens/{item_id}/enviar")
    def enviar(item_id: int, u: UsuarioOS = Depends(exigir("cotacao.lancar"))):
        """A pessoa respondeu no Coupa e marca aqui. Bloqueia se houver alerta aberto."""
        with db.conexao() as conn:
            it = um(conn, "select id, andamento, contas from itens where id = %s for update", (item_id,))
            if not it:
                raise HTTPException(status_code=404, detail="Item nao encontrado.")
            exigir_conta(u, it["contas"])
            c = um(conn, "select id, situacao from cotacoes where item_id = %s and situacao <> 'substituida' order by versao desc limit 1", (item_id,))
            if not c:
                raise HTTPException(status_code=400, detail="Lance a cotacao antes de marcar como enviada.")
            if c["situacao"] == "rascunho":
                raise HTTPException(status_code=400, detail="Ha alerta aberto. Corrija ou justifique antes de enviar.")
            if it["andamento"] != "em_cotacao":
                raise HTTPException(status_code=400, detail=f"Item em {it['andamento']} nao pode ser marcado como enviado.")
            conn.execute("update cotacoes set situacao = 'enviada' where id = %s", (c["id"],))
            conn.execute("update itens set andamento = 'enviada', atualizado_em = now() where id = %s", (item_id,))
            conn.execute("insert into item_andamento (item_id, de, para, por) values (%s, 'em_cotacao', 'enviada', %s)", (item_id, u.id))
            auditar(conn, origem=origem(u), entidade="item", entidade_id=item_id, acao="alterar", user_id=u.id,
                    antes={"andamento": "em_cotacao"}, depois={"andamento": "enviada", "cotacao_id": c["id"]})
        return {"id": item_id, "andamento": "enviada"}

    @r.post("/itens/{item_id}/resultado")
    def resultado(item_id: int, body: Resultado, u: UsuarioOS = Depends(exigir("cotacao.lancar"))):
        if body.resultado not in ("ganha", "perdida"):
            raise HTTPException(status_code=400, detail="Resultado: ganha ou perdida.")
        with db.conexao() as conn:
            it = um(conn, "select id, andamento, contas from itens where id = %s for update", (item_id,))
            if not it:
                raise HTTPException(status_code=404, detail="Item nao encontrado.")
            exigir_conta(u, it["contas"])
            if it["andamento"] not in ("enviada", "ganha", "perdida"):
                raise HTTPException(status_code=400, detail="So item enviado no Coupa recebe resultado.")
            c = um(conn, "select id from cotacoes where item_id = %s and situacao = 'enviada' order by versao desc limit 1", (item_id,))
            conn.execute("""insert into resultados (item_id, cotacao_id, resultado, pedido_venda, registrado_por) values (%s, %s, %s, %s, %s)
                            on conflict (item_id) do update set resultado = excluded.resultado, pedido_venda = excluded.pedido_venda,
                              registrado_por = excluded.registrado_por, registrado_em = now()""",
                         (item_id, c and c["id"], body.resultado, (body.pedido_venda or "").strip() or None, u.id))
            conn.execute("update itens set andamento = %s, atualizado_em = now() where id = %s", (body.resultado, item_id))
            conn.execute("insert into item_andamento (item_id, de, para, por) values (%s, %s, %s, %s)", (item_id, it["andamento"], body.resultado, u.id))
            auditar(conn, origem=origem(u), entidade="item", entidade_id=item_id, acao="alterar", user_id=u.id,
                    antes={"andamento": it["andamento"]}, depois={"andamento": body.resultado, "pedido_venda": body.pedido_venda})
        return {"id": item_id, "andamento": body.resultado}

    return r
