"""Spec 07: painel. Cada indicador leva a lista por tras do numero (link para a fila filtrada)."""
from typing import Optional

from fastapi import APIRouter, Depends, HTTPException, Query

from . import db
from .acesso import UsuarioOS, exigir
from .comum import lista, num, pode, um

MIN_RESULTADOS_PADRAO = 20  # ate a Seven definir (pendencia da Spec 07)


def calcular_painel(conn, u: UsuarioOS, dias: int, conta: Optional[int]):
    contas = [conta] if conta else list(u.contas)
    ab = "i.contas && %(c)s::bigint[] and i.andamento in ('a_cotar','em_cotacao') and (i.prazo is null or i.prazo >= current_date)"
    p = {"c": contas, "d": dias}
    ind = []

    serie = lista(conn, """select (capturado_em at time zone 'America/Sao_Paulo')::date dia, count(*) linhas from linhas_evento
                           where conta_id = any(%(c)s) and capturado_em >= now() - make_interval(days => %(d)s) group by 1 order by 1""", p)
    ultima = um(conn, "select max(capturado_em) em from linhas_evento where conta_id = any(%(c)s)", p)["em"]
    falhas = um(conn, "select count(*) n from execucoes where situacao = 'falha' and iniciada_em >= now() - make_interval(days => %(d)s)", p)["n"]
    ind.append({"chave": "capturados", "titulo": "Itens capturados por dia", "valor": sum(s["linhas"] for s in serie),
                "detalhe": f"em {len(serie)} dias com captura", "serie": serie, "ultima_captura": ultima, "falhas": falhas,
                "link": "/os/captura"})

    pref = lista(conn, """select coalesce(p.prefixo, 'sem codigo') prefixo, count(*) linhas,
                          round(100.0 * count(*) filter (where p.nivel = 'certeza') / count(*), 1) pct_certeza
                          from linhas_evento l join linha_produto lp on lp.linha_id = l.id join produtos p on p.id = lp.produto_id
                          where l.conta_id = any(%(c)s) and l.capturado_em >= now() - make_interval(days => %(d)s)
                          group by 1 order by 2 desc limit 8""", p)
    tot = sum(x["linhas"] for x in pref)
    ind.append({"chave": "identificacao", "titulo": "Identificacao por prefixo",
                "valor": round(sum(x["linhas"] * float(x["pct_certeza"]) for x in pref) / tot, 1) if tot else None,
                "unidade": "% em Certeza", "tabela": [{**x, "pct_certeza": num(x["pct_certeza"])} for x in pref],
                "link": "/os/conferencia", "disponivel": bool(tot), "motivo": None if tot else "Sem captura no periodo."})

    abas = {x["aba"]: x["n"] for x in lista(conn, f"select i.aba, count(*) n from itens i where {ab} group by 1", p)}
    ind.append({"chave": "estados", "titulo": "Itens por estado", "valor": sum(abas.values()), "abas": abas, "link": "/os"})

    n48 = um(conn, f"select count(*) n from itens i where {ab} and i.prazo <= current_date + 2", p)["n"]
    ind.append({"chave": "encerra_48h", "titulo": "Encerra nas proximas 48h", "valor": n48, "link": "/os?prazo=48h"})

    marcas = lista(conn, f"""select p.fabricante marca, count(*) itens from itens i join produtos p on p.id = i.produto_id
                             where {ab} and p.fabricante is not null
                               and not exists (select 1 from marcas m join fornecedor_marcas fm on fm.marca_id = m.id where m.nome = upper(p.fabricante))
                               and not exists (select 1 from produto_fornecedores pf where pf.produto_id = p.id)
                             group by 1 having count(*) >= 2 order by 2 desc limit 15""", p)
    ind.append({"chave": "marcas_sem_fornecedor", "titulo": "Marcas sem fornecedor", "valor": len(marcas), "tabela": marcas,
                "detalhe": "marcas com 2 ou mais itens abertos e nenhum fornecedor", "link": "/os?sem_fornecedor=1"})

    dec = um(conn, """select count(*) filter (where i.andamento = 'declinada') declinados, count(*) total from itens i
                      where i.contas && %(c)s::bigint[] and i.capturado_em >= now() - make_interval(days => %(d)s)""", p)
    ind.append({"chave": "declinio", "titulo": "Taxa de declinio",
                "valor": round(100.0 * dec["declinados"] / dec["total"], 1) if dec["total"] else None, "unidade": "%",
                "detalhe": f"{dec['declinados']} de {dec['total']} itens capturados no periodo. Sem quebra por categoria: a captura nao traz categoria.",
                "link": "/os?situacao=todos&andamento=declinada", "disponivel": bool(dec["total"])})

    if pode(u, "valor_aberto.ver"):
        v = um(conn, f"""select sum(i.quantidade_total * uv.preco_venda) valor, count(*) filter (where uv.preco_venda is not null) com, count(*) total
                         from itens i left join lateral (
                           select cv.preco_venda from cotacoes c join cotacao_valores cv on cv.cotacao_id = c.id join itens ix on ix.id = c.item_id
                           where ix.produto_id = i.produto_id and cv.preco_venda is not null and c.situacao <> 'substituida'
                           order by c.criada_em desc limit 1) uv on true where {ab}""", p)
        ind.append({"chave": "valor_aberto", "titulo": "Valor em aberto", "valor": num(v["valor"]), "unidade": "R$", "restrito": True,
                    "detalhe": f"{v['com']} de {v['total']} itens abertos tem preco de venda anterior", "link": "/os?ordem=valor",
                    "disponivel": v["com"] > 0, "motivo": None if v["com"] else "Nenhum item aberto tem preco de venda registrado ainda."})
    if pode(u, "margem.ver"):
        m = um(conn, """select count(*) n, avg((cv.preco_venda - (cv.preco_compra * (1 + coalesce(cv.ipi, 0) / 100) + coalesce(cv.frete, 0))) / cv.preco_venda * 100) media
                        from cotacoes c join cotacao_valores cv on cv.cotacao_id = c.id join itens i on i.id = c.item_id
                        where c.situacao in ('pronta','enviada') and cv.preco_venda > 0 and cv.preco_compra > 0 and i.contas && %(c)s::bigint[]
                          and c.criada_em >= now() - make_interval(days => %(d)s)""", p)
        ind.append({"chave": "margem", "titulo": "Margem praticada", "valor": round(float(m["media"]), 1) if m["media"] is not None else None,
                    "unidade": "%", "restrito": True, "detalhe": f"media de {m['n']} cotacoes no periodo",
                    "disponivel": m["n"] > 0, "motivo": None if m["n"] else "Sem cotacao com compra e venda no periodo."})
        d = lista(conn, """select a.tipo, count(*) n from alertas_preco a join cotacoes c on c.id = a.cotacao_id join itens i on i.id = c.item_id
                           where a.situacao = 'aberto' and c.situacao <> 'substituida' and i.contas && %(c)s::bigint[] group by 1""", p)
        ind.append({"chave": "divergencias", "titulo": "Divergencias abertas", "valor": sum(x["n"] for x in d), "restrito": True,
                    "tabela": d, "link": "/os?situacao=todos&andamento=em_cotacao"})
        minimo = um(conn, "select valor from parametros where chave = 'painel.min_resultados'")
        minimo = int(minimo["valor"]) if minimo else MIN_RESULTADOS_PADRAO
        rs = um(conn, """select count(*) n, count(*) filter (where r.resultado = 'ganha') ganhas from resultados r join itens i on i.id = r.item_id
                         where i.contas && %(c)s::bigint[]""", p)
        ok = rs["n"] >= minimo
        ind.append({"chave": "vitoria", "titulo": "Taxa de vitoria", "valor": round(100.0 * rs["ganhas"] / rs["n"], 1) if ok else None,
                    "unidade": "%", "restrito": True, "detalhe": f"{rs['ganhas']} ganhas em {rs['n']} com resultado",
                    "disponivel": ok, "motivo": None if ok else f"Indisponivel: ha {rs['n']} resultados registrados e o minimo e {minimo}."})
    for i in ind:
        i.setdefault("disponivel", True)
    return ind


def rotas_painel() -> APIRouter:
    r = APIRouter(prefix="/os")

    @r.get("/painel")
    def painel(dias: int = Query(30, ge=1, le=365), conta: Optional[int] = None, u: UsuarioOS = Depends(exigir("fila.ver"))):
        if conta and conta not in u.contas:
            raise HTTPException(status_code=403, detail="Conta nao liberada para voce.")
        with db.conexao() as conn:
            return {"dias": dias, "conta": conta, "indicadores": calcular_painel(conn, u, dias, conta)}

    return r
