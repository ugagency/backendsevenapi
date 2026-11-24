"""Spec 04: fila e ficha. Spec 02: declinio como descarte interno."""
import io
from typing import Optional

from fastapi import APIRouter, Depends, HTTPException, Query
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

from . import db
from .acesso import UsuarioOS, exigir
from .auditoria import auditar
from .comum import exigir_conta, lista, num, origem, pode, sincronizar_se_preciso, um, valores_visiveis

ABAS = ("pronto", "falta_fornecedor", "produto_novo", "conferencia")
ORDENS = {
    "prazo": "f.prazo asc nulls last, f.id",
    "recorrencia": "f.ocorrencias desc, f.prazo asc nulls last, f.id",
    "valor": "f.valor_estimado desc nulls last, f.prazo asc nulls last, f.id",
    "recentes": "f.capturado_em desc nulls last, f.id desc",
}
TRANSICOES = {
    "a_cotar": {"em_cotacao", "declinada"},
    "em_cotacao": {"a_cotar", "enviada", "declinada"},
    "enviada": {"em_cotacao", "ganha", "perdida", "declinada"},
    "declinada": {"a_cotar"},
    "ganha": {"enviada"},
    "perdida": {"enviada"},
}

ULTIMA_VENDA = """
  left join lateral (
    select cv.preco_venda from cotacoes c join cotacao_valores cv on cv.cotacao_id = c.id join itens ix on ix.id = c.item_id
    where ix.produto_id = i.produto_id and cv.preco_venda is not null and c.situacao <> 'substituida'
    order by c.criada_em desc limit 1) uv on true
  left join lateral (
    select h.preco_venda from historico_importado h where h.produto_id = i.produto_id and h.preco_venda is not null
    order by h.data_ref desc nulls last limit 1) hv on true
"""


class Andamento(BaseModel):
    para: str
    motivo: Optional[str] = None


def _filtros(u, busca, uf, prefixo, conta, andamento, situacao, prazo_ate, capturado_de, sem_fornecedor=None):
    cond, par = ["i.contas && %s::bigint[]"], [list(u.contas)]
    if conta:
        if conta not in u.contas:
            raise HTTPException(status_code=403, detail="Conta nao liberada para voce.")
        cond.append("%s = any(i.contas)"); par.append(conta)
    if situacao == "abertos":
        cond.append("i.andamento in ('a_cotar','em_cotacao') and (i.prazo is null or i.prazo >= current_date)")
    elif situacao == "expirados":
        cond.append("i.andamento in ('a_cotar','em_cotacao') and i.prazo < current_date")
    if andamento:
        cond.append("i.andamento = %s"); par.append(andamento)
    if uf:
        cond.append("i.uf = %s"); par.append(uf.upper())
    if prefixo:
        cond.append("p.prefixo = %s"); par.append(prefixo)
    if prazo_ate:
        cond.append("i.prazo <= %s"); par.append(prazo_ate)
    if capturado_de:
        cond.append("i.capturado_em >= %s"); par.append(capturado_de)
    if sem_fornecedor:
        cond.append("not exists (select 1 from produto_fornecedores pf where pf.produto_id = i.produto_id)")
    if busca and busca.strip():
        b = f"%{busca.strip()}%"
        cond.append("(p.nome ilike %s or p.fabricante ilike %s or p.referencia ilike %s or p.descricao ilike %s "
                    "or i.numero_evento = %s or exists (select 1 from produto_codigos pc where pc.produto_id = p.id and pc.codigo like %s))")
        par += [b, b, b, b, busca.strip(), busca.strip() + "%"]
    return " and ".join(cond), par


def _consulta(where):
    return f"""
    with f as (
      select i.id, i.numero_evento, i.prazo, i.quantidade_total, i.unidade, i.uf, i.aba, i.andamento, i.contas, i.linhas,
             i.capturado_em, i.responsavel_id, p.id produto_id, p.nome, p.fabricante, p.referencia, p.nivel, p.prefixo,
             (select count(*) from itens x where x.produto_id = p.id) ocorrencias,
             i.quantidade_total * coalesce(uv.preco_venda, hv.preco_venda) valor_estimado
      from itens i join produtos p on p.id = i.produto_id {ULTIMA_VENDA}
      where {where})
    """


def _linha_fila(u, r):
    d = {k: r[k] for k in ("id", "numero_evento", "prazo", "unidade", "uf", "aba", "andamento", "linhas", "ocorrencias")}
    d.update(quantidade=num(r["quantidade_total"]), contas=list(r["contas"]),
             produto={"id": r["produto_id"], "nome": r["nome"], "fabricante": r["fabricante"], "referencia": r["referencia"],
                      "nivel": r["nivel"], "codigos": r.get("codigos") or []},
             fornecedores=r.get("fornecedores", 0), declinios=r.get("declinios", 0), responsavel=r.get("responsavel"),
             expirado=bool(r.get("expirado")))
    if pode(u, "valor_aberto.ver"):
        d["valor_estimado"] = num(r["valor_estimado"])
    return d


ENRIQUECE = """
  select f.*, (f.prazo < current_date and f.andamento in ('a_cotar','em_cotacao')) expirado,
    (select array_agg(pc.codigo order by pc.codigo) from produto_codigos pc where pc.produto_id = f.produto_id) codigos,
    (select count(*) from produto_fornecedores pf where pf.produto_id = f.produto_id) fornecedores,
    (select count(*) from declinados d where d.produto_id = f.produto_id) declinios,
    (select coalesce(nullif(us.nome, ''), us.email) from usuarios us where us.user_id = f.responsavel_id) responsavel
  from f
"""


def rotas_fila() -> APIRouter:
    r = APIRouter(prefix="/os")

    @r.get("/fila")
    def fila(aba: Optional[str] = None, busca: Optional[str] = None, uf: Optional[str] = None, prefixo: Optional[str] = None,
             conta: Optional[int] = None, andamento: Optional[str] = None, situacao: str = "abertos",
             prazo_ate: Optional[str] = None, capturado_de: Optional[str] = None, sem_fornecedor: bool = False,
             ordem: str = "prazo", pagina: int = Query(1, ge=1), por_pagina: int = Query(50, ge=1, le=200),
             u: UsuarioOS = Depends(exigir("fila.ver"))):
        if aba and aba not in ABAS:
            raise HTTPException(status_code=400, detail="Aba invalida.")
        if ordem not in ORDENS:
            raise HTTPException(status_code=400, detail="Ordenacao invalida.")
        if ordem == "valor" and not pode(u, "valor_aberto.ver"):
            raise HTTPException(status_code=403, detail="Sem permissao para ordenar por valor.")
        where, par = _filtros(u, busca, uf, prefixo, conta, andamento, situacao, prazo_ate, capturado_de, sem_fornecedor)
        base = _consulta(where)
        with db.conexao() as conn:
            abas = {x["aba"]: x["n"] for x in lista(conn, base + "select aba, count(*) n from f group by aba", par)}
            filtro_aba, par2 = ("where f.aba = %s", par + [aba]) if aba else ("", list(par))
            total = sum(abas.values()) if not aba else abas.get(aba, 0)
            pagina_sql = (base + f", f2 as (select * from f {filtro_aba} order by {ORDENS[ordem]} limit %s offset %s) "
                          + ENRIQUECE.replace("from f\n", "from f2 f\n") + f" order by {ORDENS[ordem]}")
            rows = lista(conn, pagina_sql, par2 + [por_pagina, (pagina - 1) * por_pagina])
        return {"abas": {a: abas.get(a, 0) for a in ABAS}, "total": total, "pagina": pagina,
                "itens": [_linha_fila(u, x) for x in rows]}

    @r.get("/fila/cartoes")
    def cartoes(u: UsuarioOS = Depends(exigir("fila.ver"))):
        sincronizar_se_preciso()
        where, par = _filtros(u, None, None, None, None, None, "abertos", None, None)
        with db.conexao() as conn:
            c = um(conn, _consulta(where) + """
              select count(*) em_aberto,
                count(*) filter (where prazo <= current_date + 2) encerra_48h,
                count(*) filter (where responsavel_id = %s and andamento = 'em_cotacao') aguardando_voce,
                sum(valor_estimado) valor_estimado, count(*) filter (where valor_estimado is not null) com_valor
              from f""", par + [u.id])
        d = {"em_aberto": c["em_aberto"], "encerra_48h": c["encerra_48h"], "aguardando_voce": c["aguardando_voce"]}
        if pode(u, "valor_aberto.ver"):
            d.update(valor_estimado=num(c["valor_estimado"]), itens_com_valor=c["com_valor"])
        return d

    @r.get("/fila/exportar")
    def exportar(aba: Optional[str] = None, busca: Optional[str] = None, uf: Optional[str] = None,
                 prefixo: Optional[str] = None, conta: Optional[int] = None, andamento: Optional[str] = None,
                 situacao: str = "abertos", prazo_ate: Optional[str] = None, capturado_de: Optional[str] = None,
                 sem_fornecedor: bool = False, ordem: str = "prazo", u: UsuarioOS = Depends(exigir("exportar"))):
        from openpyxl import Workbook
        if ordem not in ORDENS or (ordem == "valor" and not pode(u, "valor_aberto.ver")):
            ordem = "prazo"
        where, par = _filtros(u, busca, uf, prefixo, conta, andamento, situacao, prazo_ate, capturado_de, sem_fornecedor)
        filtro_aba, par2 = ("where f.aba = %s", par + [aba]) if aba in ABAS else ("", list(par))
        sql = (_consulta(where) + f", f2 as (select * from f {filtro_aba} order by {ORDENS[ordem]} limit 20000) "
               + ENRIQUECE.replace("from f\n", "from f2 f\n") + f" order by {ORDENS[ordem]}")
        with db.conexao() as conn:
            rows = lista(conn, sql, par2)
            nomes = {c["id"]: c["nome"] for c in lista(conn, "select id, nome from contas_portal")}
            auditar(conn, origem=origem(u), entidade="fila", acao="exportar", user_id=u.id,
                    depois={"linhas": len(rows), "filtros": {"aba": aba, "busca": busca, "uf": uf, "prefixo": prefixo,
                                                             "conta": conta, "andamento": andamento, "situacao": situacao}})
        wb = Workbook(); ws = wb.active; ws.title = "Fila"
        cab = ["Evento", "Codigos", "Item", "Fabricante", "Referencia", "Identificacao", "UF", "Quantidade", "Unidade",
               "Encerramento", "Aba", "Andamento", "Filiais", "Ocorrencias", "Fornecedores", "Declinios"]
        com_valor = pode(u, "valor_aberto.ver")
        if com_valor:
            cab.append("Valor estimado")
        ws.append(cab)
        for x in rows:
            linha = [x["numero_evento"], ", ".join(x["codigos"] or []), x["nome"], x["fabricante"], x["referencia"], x["nivel"],
                     x["uf"], num(x["quantidade_total"]), x["unidade"], x["prazo"], x["aba"], x["andamento"],
                     ", ".join(nomes.get(c, str(c)) for c in x["contas"]), x["ocorrencias"], x["fornecedores"], x["declinios"]]
            if com_valor:
                linha.append(num(x["valor_estimado"]))
            ws.append(linha)
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        return StreamingResponse(buf, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                 headers={"Content-Disposition": 'attachment; filename="fila_seven_os.xlsx"'})

    @r.get("/itens/{item_id}")
    def ficha(item_id: int, u: UsuarioOS = Depends(exigir("fila.ver"))):
        with db.conexao() as conn:
            it = um(conn, """select i.*, (i.prazo < current_date and i.andamento in ('a_cotar','em_cotacao')) expirado,
                             (select coalesce(nullif(us.nome, ''), us.email) from usuarios us where us.user_id = i.responsavel_id) responsavel
                             from itens i where i.id = %s""", (item_id,))
            if not it:
                raise HTTPException(status_code=404, detail="Item nao encontrado.")
            exigir_conta(u, it["contas"])
            pid = it["produto_id"]
            p = um(conn, "select id, nome, descricao, fabricante, referencia, ref_chave, prefixo, nivel, metodo from produtos where id = %s", (pid,))
            codigos = lista(conn, "select codigo, motivo, juntado_em from produto_codigos where produto_id = %s order by codigo", (pid,))
            contas = lista(conn, "select id, nome, cnpj, filial from contas_portal where id = any(%s) order by id", (list(it["contas"]),))
            linhas = lista(conn, """select l.id, l.conta_id, l.numero_linha, l.codigo, l.descricao, l.quantidade, l.unidade, l.prazo_texto,
                                    l.marca, l.capturado_em, l.link_evento from linhas_evento l join linha_produto lp on lp.linha_id = l.id
                                    where l.numero_evento = %s and lp.produto_id = %s order by l.conta_id, l.numero_linha""",
                           (it["numero_evento"], pid))
            anteriores = lista(conn, """
                select i.id, i.numero_evento, i.prazo, i.quantidade_total, i.andamento, i.capturado_em,
                       c.id cotacao_id, f.nome fornecedor, cv.preco_compra, cv.ipi, cv.frete, cv.preco_venda, rs.resultado
                from itens i
                left join lateral (select * from cotacoes c where c.item_id = i.id and c.situacao <> 'substituida' order by versao desc limit 1) c on true
                left join cotacao_valores cv on cv.cotacao_id = c.id
                left join fornecedores f on f.id = c.fornecedor_id
                left join resultados rs on rs.item_id = i.id
                where i.produto_id = %s and i.id <> %s and i.contas && %s::bigint[] order by i.capturado_em desc nulls last limit 30""",
                               (pid, item_id, list(u.contas)))
            importado = lista(conn, """select origem, data_ref, fornecedor_texto, status, preco_compra, ipi, frete, preco_venda
                                       from historico_importado where produto_id = %s order by data_ref desc nulls last limit 30""", (pid,))
            declinios = lista(conn, """select d.em, d.motivo, i.numero_evento, coalesce(nullif(us.nome, ''), us.email) por
                                       from declinados d join itens i on i.id = d.item_id left join usuarios us on us.user_id = d.por
                                       where d.produto_id = %s order by d.em desc limit 30""", (pid,))
            andamentos = lista(conn, """select a.de, a.para, a.motivo, a.em, coalesce(nullif(us.nome, ''), us.email) por
                                        from item_andamento a left join usuarios us on us.user_id = a.por where a.item_id = %s order by a.em desc""", (item_id,))
            parecidos = lista(conn, """select q.id, q.nome, q.fabricante, q.referencia, q.nivel,
                                       (select array_agg(codigo order by codigo) from produto_codigos pc where pc.produto_id = q.id) codigos
                                       from produtos q where q.ref_chave = %s and q.id <> %s and q.juntado_em is null limit 10""",
                              (p["ref_chave"], pid)) if p["ref_chave"] else []
            juncoes = lista(conn, """select j.id, j.origem_id, o.nome origem_nome, j.motivo, j.em,
                                     (select array_agg(distinct l.codigo) from linhas_evento l where l.id = any(j.linhas) and l.codigo is not null) codigos
                                     from produto_juncoes j join produtos o on o.id = j.origem_id
                                     where j.destino_id = %s and j.desfeito_em is null order by j.em desc""", (pid,))
            ocorrencias = um(conn, "select count(*) n from itens where produto_id = %s", (pid,))["n"]
        for a in anteriores:
            v = valores_visiveis(u, a)
            for k in ("preco_compra", "ipi", "frete", "preco_venda"):
                a.pop(k, None)
            a["quantidade_total"] = num(a["quantidade_total"]); a["valores"] = v
        for h in importado:
            v = valores_visiveis(u, h)
            for k in ("preco_compra", "ipi", "frete", "preco_venda"):
                h.pop(k, None)
            h["valores"] = v
        for l in linhas:
            l["quantidade"] = num(l["quantidade"])
        item = {k: it[k] for k in ("id", "numero_evento", "prazo", "unidade", "uf", "aba", "andamento", "linhas", "capturado_em",
                                   "responsavel", "expirado")}
        item["quantidade"] = num(it["quantidade_total"])
        return {"item": item, "produto": p, "codigos": codigos, "contas": contas, "linhas": linhas, "ocorrencias": ocorrencias,
                "anteriores": anteriores, "historico_importado": importado, "declinios": declinios, "andamentos": andamentos,
                "mesma_referencia": parecidos, "juncoes": juncoes,
                "transicoes": sorted(TRANSICOES.get(it["andamento"], set()))}

    @r.patch("/itens/{item_id}/andamento")
    def mudar_andamento(item_id: int, body: Andamento, u: UsuarioOS = Depends(exigir("cotacao.lancar"))):
        with db.conexao() as conn:
            it = um(conn, "select id, produto_id, andamento, contas from itens where id = %s for update", (item_id,))
            if not it:
                raise HTTPException(status_code=404, detail="Item nao encontrado.")
            exigir_conta(u, it["contas"])
            if body.para not in TRANSICOES.get(it["andamento"], set()):
                raise HTTPException(status_code=400, detail=f"Nao da para ir de {it['andamento']} para {body.para}.")
            if body.para in ("enviada", "ganha", "perdida"):
                raise HTTPException(status_code=400, detail="Use o envio ou o registro de resultado da cotacao.")
            motivo = (body.motivo or "").strip()
            if body.para == "declinada" and not motivo:
                raise HTTPException(status_code=400, detail="Motivo obrigatorio para declinar.")
            conn.execute("update itens set andamento = %s, responsavel_id = case when %s = 'em_cotacao' then %s else responsavel_id end, "
                         "atualizado_em = now() where id = %s", (body.para, body.para, u.id, item_id))
            conn.execute("insert into item_andamento (item_id, de, para, motivo, por) values (%s, %s, %s, %s, %s)",
                         (item_id, it["andamento"], body.para, motivo or None, u.id))
            if body.para == "declinada":
                conn.execute("insert into declinados (item_id, produto_id, motivo, por) values (%s, %s, %s, %s)",
                             (item_id, it["produto_id"], motivo, u.id))
            if it["andamento"] == "declinada":
                conn.execute("select atualizar_abas(%s)", ([it["produto_id"]],))
            auditar(conn, origem=origem(u), entidade="item", entidade_id=item_id, acao="alterar", user_id=u.id,
                    antes={"andamento": it["andamento"]}, depois={"andamento": body.para}, justificativa=motivo or None)
        return {"id": item_id, "andamento": body.para}

    return r
