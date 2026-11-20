"""Spec 09: auditoria consultavel e exportacao completa da base em formato aberto (termo, itens 4.2 e 4.3)."""
import csv
import io
import zipfile
from datetime import datetime
from typing import Optional

from fastapi import APIRouter, Depends, HTTPException, Query
from fastapi.responses import StreamingResponse

from . import db
from .acesso import UsuarioOS, exigir, usuario_os
from .auditoria import auditar
from .comum import lista, origem

TABELAS_EXPORTACAO = (
    "usuarios", "contas_portal", "usuario_contas", "parametros", "execucoes", "linhas_evento", "produtos", "produto_codigos",
    "linha_produto", "produto_juncoes", "itens", "item_andamento", "declinados", "fornecedores", "fornecedor_nomes", "marcas",
    "fornecedor_marcas", "produto_fornecedores", "sugestoes_fornecedor", "solicitacoes_cotacao", "cotacoes", "cotacao_valores",
    "alertas_preco", "resultados", "historico_importado", "conversas", "mensagens", "auditoria")


def rotas_relatorios() -> APIRouter:
    r = APIRouter(prefix="/os")

    @r.get("/auditoria")
    def auditoria(entidade: Optional[str] = None, autor: Optional[str] = None, acao: Optional[str] = None,
                  de: Optional[str] = None, ate: Optional[str] = None, pagina: int = Query(1, ge=1),
                  u: UsuarioOS = Depends(exigir("auditoria.ver"))):
        cond, par = ["true"], []
        for campo, valor in (("a.entidade", entidade), ("a.acao", acao)):
            if valor:
                cond.append(f"{campo} = %s"); par.append(valor)
        if autor:
            cond.append("(us.email ilike %s or us.nome ilike %s)"); par += [f"%{autor}%"] * 2
        if de:
            cond.append("a.em >= %s"); par.append(de)
        if ate:
            cond.append("a.em < %s::date + 1"); par.append(ate)
        # Quem nao ve preco nao ve o antes/depois de registros com valores.
        ve_valores = "preco_compra.ver" in u.permissoes and "margem.ver" in u.permissoes
        with db.conexao() as conn:
            rows = lista(conn, f"""select a.id, a.em, a.origem, a.entidade, a.entidade_id, a.acao, a.antes, a.depois, a.justificativa,
                                   coalesce(nullif(us.nome, ''), us.email, 'sistema') autor
                                   from auditoria a left join usuarios us on us.user_id = a.user_id
                                   where {' and '.join(cond)} order by a.id desc limit 100 offset %s""", par + [(pagina - 1) * 100])
            entidades = [x["entidade"] for x in lista(conn, "select distinct entidade from auditoria order by 1")]
        if not ve_valores:
            for x in rows:
                if x["entidade"] in ("cotacao", "alerta_preco"):
                    x["antes"] = x["depois"] = None
                    x["oculto"] = True
        return {"pagina": pagina, "registros": rows, "entidades": entidades}

    @r.post("/admin/exportar-base")
    def exportar_base(u: UsuarioOS = Depends(usuario_os)):
        """Pacote completo em CSV (um arquivo por tabela), so para o Administrador. Fica registrado na auditoria."""
        if u.perfil != "administrador":
            raise HTTPException(status_code=403, detail="So o Administrador exporta a base completa.")
        buf = io.BytesIO()
        with db.conexao() as conn, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            contagem = {}
            for t in TABELAS_EXPORTACAO:
                with conn.cursor() as cur:
                    cur.execute(f"select * from {t}")
                    cols = [d.name for d in cur.description]
                    saida = io.StringIO(); w = csv.writer(saida); w.writerow(cols)
                    n = 0
                    for linha in cur:
                        w.writerow(linha); n += 1
                contagem[t] = n
                z.writestr(f"{t}.csv", saida.getvalue())
            z.writestr("LEIAME.txt", "Seven OS - exportacao completa da base\n"
                       f"Gerada em {datetime.now().isoformat(timespec='seconds')} por {u.email}\n"
                       "Um CSV por tabela, UTF-8, com cabecalho.\n" + "\n".join(f"{k}: {v} linhas" for k, v in contagem.items()) + "\n")
            auditar(conn, origem=origem(u), entidade="base", acao="exportar", user_id=u.id, depois=contagem)
        buf.seek(0)
        nome = f"seven_os_base_{datetime.now():%Y%m%d_%H%M}.zip"
        return StreamingResponse(buf, media_type="application/zip", headers={"Content-Disposition": f'attachment; filename="{nome}"'})

    return r
