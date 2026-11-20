"""Spec 02: capturas. Nesta fase o OS le o que o robo ja grava (sincronizacao); nada muda no robo."""
import io
import os
from typing import Optional

import httpx
from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.responses import StreamingResponse

from . import db
from .acesso import UsuarioOS, exigir
from .auditoria import auditar
from .comum import lista, num, origem, um

CABECALHO = ["Numero do evento", "Titulo", "UF(VALE)", "DATA", "DESCRIÇÃO", "QTDE", "UNID. MED", "pagina de descrição", "Marca"]


def rotas_captura() -> APIRouter:
    r = APIRouter(prefix="/os/captura")

    @r.get("/execucoes")
    def execucoes(u: UsuarioOS = Depends(exigir("fila.ver"))):
        with db.conexao() as conn:
            return lista(conn, """select e.id, e.origem, e.iniciada_em, e.terminada_em, e.situacao, e.linhas_novas, e.erro, c.nome conta,
                                  coalesce(nullif(us.nome, ''), us.email) por
                                  from execucoes e left join contas_portal c on c.id = e.conta_id left join usuarios us on us.user_id = e.pedido_por
                                  where e.linhas_novas is distinct from 0 or e.situacao <> 'sucesso' order by e.id desc limit 50""")

    @r.get("/dias")
    def dias(u: UsuarioOS = Depends(exigir("fila.ver"))):
        """Dias com captura, por conta, para o download por data."""
        with db.conexao() as conn:
            rows = lista(conn, """select (l.capturado_em at time zone 'America/Sao_Paulo')::date dia, l.conta_id, c.nome conta,
                                  count(*) linhas, count(distinct l.numero_evento) eventos
                                  from linhas_evento l join contas_portal c on c.id = l.conta_id
                                  where l.conta_id = any(%s) and l.capturado_em >= now() - interval '60 days'
                                  group by 1, 2, 3 order by 1 desc, 2""", (list(u.contas),))
            ultima = um(conn, "select max(capturado_em) em from linhas_evento where conta_id = any(%s)", (list(u.contas),))
        return {"dias": rows, "ultima_captura": ultima["em"]}

    @r.post("/sincronizar")
    def sincronizar(u: UsuarioOS = Depends(exigir("captura.rodar"))):
        """Traz para o OS o que o robo gravou desde a ultima vez. Leitura apenas em eventos_coletados."""
        with db.conexao() as conn:
            res = conn.execute("select sincronizar(%s)", (u.id,)).fetchone()[0]
        return res

    @r.post("/rodar")
    def rodar(request: Request, corpo: dict, u: UsuarioOS = Depends(exigir("captura.rodar"))):
        """Pede uma execucao ao servico do robo. So funciona onde OS_ROBO_URL estiver configurada (producao, apos o corte)."""
        url = os.getenv("OS_ROBO_URL", "").rstrip("/")
        if not url:
            raise HTTPException(status_code=503, detail="Neste ambiente o robo nao e acionado pelo OS. Use a tela atual de captura.")
        try:
            resp = httpx.post(f"{url}/run-robot", json=corpo, timeout=30,
                              headers={"Authorization": request.headers.get("authorization", "")})
        except httpx.HTTPError:
            raise HTTPException(status_code=502, detail="Servico do robo fora do ar.")
        with db.conexao() as conn:
            auditar(conn, origem=origem(u), entidade="captura", acao="criar", user_id=u.id, depois={"pedido": corpo, "status": resp.status_code})
        if resp.status_code >= 400:
            raise HTTPException(status_code=resp.status_code, detail=resp.text[:300])
        return resp.json()

    @r.get("/planilha")
    def planilha(data: str, conta: Optional[int] = None, u: UsuarioOS = Depends(exigir("captura.baixar"))):
        """Planilha do dia montada do banco, no layout da planilha atual (9 colunas; 'Filiais' na unificada)."""
        from openpyxl import Workbook
        contas = [conta] if conta else list(u.contas)
        if conta and conta not in u.contas:
            raise HTTPException(status_code=403, detail="Conta nao liberada para voce.")
        with db.conexao() as conn:
            rows = lista(conn, """
                select l.numero_evento, l.codigo, l.uf, l.prazo_texto, l.descricao, l.quantidade_texto, l.quantidade, l.unidade, l.link_evento, l.marca,
                       l.conta_id, c.nome conta
                from linhas_evento l join contas_portal c on c.id = l.conta_id
                where l.conta_id = any(%s) and (l.capturado_em at time zone 'America/Sao_Paulo')::date = %s
                order by l.numero_evento, l.codigo, l.numero_linha, l.conta_id""", (contas, data))
            if not rows:
                raise HTTPException(status_code=404, detail="Sem captura nessa data para as contas escolhidas.")
            auditar(conn, origem=origem(u), entidade="planilha", acao="exportar", user_id=u.id, depois={"data": data, "conta": conta, "linhas": len(rows)})
        wb = Workbook(); ws = wb.active; ws.title = "Eventos"
        unificada = conta is None and len(contas) > 1
        ws.append(CABECALHO + (["Filiais"] if unificada else []))
        if unificada:
            juntos = {}
            for x in rows:
                k = (x["numero_evento"], x["codigo"], x["descricao"], x["quantidade_texto"])
                juntos.setdefault(k, {"l": x, "contas": []})["contas"].append(x["conta"])
            saida = [(v["l"], ", ".join(dict.fromkeys(v["contas"]))) for v in juntos.values()]
        else:
            saida = [(x, None) for x in rows]
        for x, filiais in saida:
            linha = [x["numero_evento"], x["codigo"], x["uf"], x["prazo_texto"], x["descricao"],
                     x["quantidade_texto"] or num(x["quantidade"]), x["unidade"], x["link_evento"], x["marca"] or "MARCA NAO ENCONTRADA"]
            ws.append(linha + ([filiais] if unificada else []))
        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        nome = f"eventos_{'todas_' if unificada else ''}{data}.xlsx"
        return StreamingResponse(buf, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                 headers={"Content-Disposition": f'attachment; filename="{nome}"'})

    return r
