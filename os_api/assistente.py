"""Spec 08: assistente somente leitura, com lista fechada de consultas.

O roteador escolhe UMA consulta da lista; quem executa e o os_api, com a permissao de quem perguntou.
Valores (preco, margem) so aparecem na tabela montada pelo sistema e nunca entram em texto enviado a um modelo.
A escolha da consulta e feita por um modelo de IA (OpenRouter, via workflow do n8n) quando OS_N8N_ASSISTENTE_URL esta
configurada; o que sai do servidor e SO o texto da pergunta e a data de hoje, sem historico, sem preco e sem dado da Seven.
O retorno do modelo e validado contra a lista fechada (consulta e parametros permitidos). Sem URL, com erro ou com resposta
invalida, cai nas regras de `escolher_consulta`, que nao dependem de nada externo.
"""
import os
import re
import unicodedata
from datetime import date, timedelta
from typing import Optional

import httpx
from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel
from psycopg.types.json import Jsonb

from . import db
from .acesso import UsuarioOS, exigir
from .comum import lista, num, pode, um, valores_visiveis
from .painel import calcular_painel

ABAS_NOME = {"pronto": "Pronto", "falta_fornecedor": "Falta fornecedor", "produto_novo": "Produto novo", "conferencia": "Conferência manual"}
CONSULTAS = ("historico_produto", "fornecedores_produto", "itens_por_prazo", "marcas_sem_fornecedor", "indicadores", "buscar_produto")
DIAS_SEMANA = {"segunda": 0, "terca": 1, "quarta": 2, "quinta": 3, "sexta": 4, "sabado": 5, "domingo": 6}
VAZIAS = {"ja", "cotamos", "cotou", "essa", "esse", "esta", "este", "peca", "item", "produto", "por", "quanto", "qual", "quais", "quem",
          "fornece", "fornecedor", "fornecedores", "vende", "de", "da", "do", "para", "o", "a", "os", "as", "e", "que", "tem", "temos",
          "historico", "preco", "precos", "foi", "com", "um", "uma", "no", "na", "ultima", "vez", "cotacao", "cotacoes", "buscar", "procurar",
          "mostre", "mostrar", "me", "sobre", "ja", "existe", "algum", "alguma"}


class Pergunta(BaseModel):
    pergunta: str
    conversa_id: Optional[int] = None


def _sem_acento(t):
    return "".join(c for c in unicodedata.normalize("NFD", t.lower()) if unicodedata.category(c) != "Mn")


def _termo(pergunta):
    aspas = re.search(r"[\"']([^\"']{2,})[\"']", pergunta)
    if aspas:
        return aspas.group(1).strip()
    palavras = [w for w in re.findall(r"[\w\-/\.]+", pergunta) if _sem_acento(w) not in VAZIAS]
    return " ".join(palavras).strip()


def escolher_consulta(pergunta: str, hoje: Optional[date] = None):
    """Devolve (consulta, parametros). Lista fechada; nunca gera SQL."""
    hoje = hoje or date.today()
    t = _sem_acento(pergunta)
    if re.search(r"\bmarcas?\b.*\bsem fornecedor", t):
        return "marcas_sem_fornecedor", {}
    if re.search(r"encerra|vence|prazo|fecha", t):
        ate = None
        m = re.search(r"(\d+)\s*dias?", t)
        if "hoje" in t:
            ate = hoje
        elif "amanha" in t:
            ate = hoje + timedelta(days=1)
        elif m:
            ate = hoje + timedelta(days=int(m.group(1)))
        else:
            for nome, n in DIAS_SEMANA.items():
                if nome in t:
                    ate = hoje + timedelta(days=(n - hoje.weekday()) % 7)
                    break
        if "semana" in t and ate is None:
            ate = hoje + timedelta(days=(6 - hoje.weekday()))
        return "itens_por_prazo", {"ate": (ate or hoje + timedelta(days=2)).isoformat(), "sem_fornecedor": "sem fornecedor" in t}
    if re.search(r"indicador|painel|resumo|como esta|quantos itens", t):
        return "indicadores", {}
    termo = _termo(pergunta)
    if re.search(r"fornece|fornecedor|quem vende", t):
        return "fornecedores_produto", {"termo": termo}
    if re.search(r"cotamos|cotou|cotacao|quanto|preco|historico", t):
        return "historico_produto", {"termo": termo}
    return "buscar_produto", {"termo": termo}


def validar_escolha(dados) -> Optional[tuple]:
    """Aceita so uma consulta da lista fechada e parametros conhecidos. Qualquer outra coisa e descartada (devolve None)."""
    if not isinstance(dados, dict) or dados.get("consulta") not in CONSULTAS:
        return None
    bruto = dados.get("parametros") or {}
    if not isinstance(bruto, dict):
        return None
    par = {}
    termo = bruto.get("termo")
    if isinstance(termo, str) and termo.strip():
        par["termo"] = termo.strip()[:120]
    ate = bruto.get("ate")
    if isinstance(ate, str):
        try:
            par["ate"] = date.fromisoformat(ate[:10]).isoformat()
        except ValueError:
            pass
    if isinstance(bruto.get("sem_fornecedor"), bool):
        par["sem_fornecedor"] = bruto["sem_fornecedor"]
    consulta = dados["consulta"]
    if consulta == "itens_por_prazo" and "ate" not in par:
        return None
    if consulta in ("historico_produto", "fornecedores_produto", "buscar_produto") and "termo" not in par:
        return None
    return consulta, par


def escolher_consulta_ia(pergunta: str, hoje: Optional[date] = None) -> Optional[tuple]:
    """Pede a um modelo (workflow do n8n -> OpenRouter) para escolher a consulta. Manda so a pergunta e a data.
    Devolve (consulta, parametros) ou None para usar as regras."""
    url = os.getenv("OS_N8N_ASSISTENTE_URL", "").strip()
    if not url:
        return None
    hoje = hoje or date.today()
    cab = {"X-Seven-Token": os.getenv("OS_N8N_TOKEN", "")}
    corpo = {"pergunta": pergunta, "hoje": hoje.isoformat(), "dia_da_semana": hoje.strftime("%A"), "consultas": list(CONSULTAS)}
    try:
        resp = httpx.post(url, json=corpo, headers=cab, timeout=float(os.getenv("OS_N8N_TIMEOUT", "12")))
        if resp.status_code != 200:
            return None
        return validar_escolha(resp.json())
    except (httpx.HTTPError, ValueError):
        return None


def _produtos(conn, termo, limite=5):
    if not termo or len(termo) < 2:
        return []
    b = f"%{termo}%"
    return lista(conn, """select p.id, p.nome, p.fabricante, p.referencia,
        (select array_agg(codigo order by codigo) from produto_codigos pc where pc.produto_id = p.id) codigos,
        (select count(*) from itens i where i.produto_id = p.id) ocorrencias
        from produtos p where p.juntado_em is null and (p.nome ilike %s or p.fabricante ilike %s or p.referencia ilike %s or p.descricao ilike %s
          or exists (select 1 from produto_codigos pc where pc.produto_id = p.id and pc.codigo = %s))
        order by ocorrencias desc limit %s""", (b, b, b, b, termo, limite))


def executar(conn, u: UsuarioOS, consulta: str, par: dict):
    """Executa a consulta com a permissao de quem perguntou. Devolve texto, tabela e link."""
    contas = list(u.contas)
    if consulta in ("historico_produto", "fornecedores_produto", "buscar_produto"):
        prods = _produtos(conn, par.get("termo"))
        if not prods:
            return {"texto": "Não encontrei esse produto nos registros. Tente o código Vale, a referência ou parte do nome entre aspas.", "tabela": None, "link": None}
        if consulta == "buscar_produto" or len(prods) > 1 and consulta != "buscar_produto" and par.get("termo", "") not in (prods[0]["codigos"] or []):
            if consulta == "buscar_produto" or len(prods) > 1:
                linhas = [[p["nome"], p["fabricante"] or "", p["referencia"] or "", ", ".join(p["codigos"] or []), p["ocorrencias"]] for p in prods]
                extra = "" if consulta == "buscar_produto" else " Diga qual deles (pelo código) para eu trazer o detalhe."
                return {"texto": f"Encontrei {len(prods)} produto(s) com esse termo.{extra}",
                        "tabela": {"colunas": ["Produto", "Fabricante", "Referência", "Códigos", "Ocorrências"], "linhas": linhas}, "link": None}
        p = prods[0]
        if consulta == "fornecedores_produto":
            fs = lista(conn, """select f.nome, f.cidade, f.uf, f.contato, f.telefone, f.email, pf.origem from produto_fornecedores pf
                                join fornecedores f on f.id = pf.fornecedor_id where pf.produto_id = %s order by f.nome""", (p["id"],))
            if not fs:
                return {"texto": f"{p['nome']}: nenhum fornecedor vinculado ainda.", "tabela": None, "link": None}
            return {"texto": f"{p['nome']}: {len(fs)} fornecedor(es) vinculado(s).",
                    "tabela": {"colunas": ["Fornecedor", "Cidade", "UF", "Contato", "Telefone", "E-mail", "Origem"],
                               "linhas": [[f["nome"], f["cidade"] or "", f["uf"] or "", f["contato"] or "", f["telefone"] or "", f["email"] or "", f["origem"]] for f in fs]},
                    "link": None}
        cots = lista(conn, """select i.id item_id, i.numero_evento, c.criada_em, c.situacao, f.nome fornecedor, i.quantidade_total,
                              cv.preco_compra, cv.ipi, cv.frete, cv.preco_venda, rs.resultado
                              from itens i join cotacoes c on c.item_id = i.id and c.situacao <> 'substituida'
                              left join cotacao_valores cv on cv.cotacao_id = c.id left join fornecedores f on f.id = c.fornecedor_id
                              left join resultados rs on rs.item_id = i.id
                              where i.produto_id = %s and i.contas && %s::bigint[] order by c.criada_em desc limit 10""", (p["id"], contas))
        apar = um(conn, "select count(*) n, max(capturado_em) ultima, min(id) item from itens where produto_id = %s and contas && %s::bigint[]", (p["id"], contas))
        link = f"/os/item/{cots[0]['item_id']}" if cots else (f"/os/item/{apar['item']}" if apar["item"] else None)
        if not cots:
            return {"texto": f"{p['nome']}: apareceu {apar['n']} vez(es) na captura, mas não há cotação registrada no sistema.", "tabela": None, "link": link}
        colunas = ["Data", "Evento", "Fornecedor", "Quantidade", "Situação", "Resultado"]
        chaves = [k for k in ("preco_compra", "ipi", "frete", "preco_venda", "margem_pct") if k in valores_visiveis(u, {})]
        nomes = {"preco_compra": "Compra (un.)", "ipi": "IPI %", "frete": "Frete", "preco_venda": "Venda (un.)", "margem_pct": "Margem %"}
        linhas = []
        for c in cots:
            v = valores_visiveis(u, c)
            linhas.append([c["criada_em"].strftime("%d/%m/%Y"), c["numero_evento"], c["fornecedor"] or "", num(c["quantidade_total"]),
                           c["situacao"], c["resultado"] or ""] + [v.get(k) for k in chaves])
        return {"texto": f"{p['nome']}: {len(cots)} cotação(ões) registrada(s). Os valores abaixo vêm direto do sistema.",
                "tabela": {"colunas": colunas + [nomes[k] for k in chaves], "linhas": linhas}, "link": link}
    if consulta == "itens_por_prazo":
        sem = "and not exists (select 1 from produto_fornecedores pf where pf.produto_id = i.produto_id)" if par.get("sem_fornecedor") else ""
        rows = lista(conn, f"""select i.id, i.numero_evento, i.prazo, i.quantidade_total, i.uf, i.aba, p.nome, p.fabricante from itens i
                               join produtos p on p.id = i.produto_id where i.contas && %s::bigint[] and i.andamento in ('a_cotar','em_cotacao')
                               and i.prazo >= current_date and i.prazo <= %s {sem} order by i.prazo, i.id limit 21""", (contas, par["ate"]))
        ate = date.fromisoformat(par["ate"]).strftime("%d/%m/%Y")
        qual = " sem fornecedor" if par.get("sem_fornecedor") else ""
        link = f"/os?prazo_ate={par['ate']}" + ("&sem_fornecedor=1" if par.get("sem_fornecedor") else "")
        if not rows:
            return {"texto": f"Nenhum item aberto{qual} encerra até {ate}.", "tabela": None, "link": link}
        mais = len(rows) > 20
        rows = rows[:20]
        return {"texto": f"{'Mais de 20' if mais else len(rows)} item(ns) aberto(s){qual} encerram até {ate}." + (" Mostro os 20 primeiros; a lista completa está na fila." if mais else ""),
                "tabela": {"colunas": ["Encerra", "Evento", "Item", "Fabricante", "Quantidade", "UF", "Situação"],
                           "linhas": [[x["prazo"].strftime("%d/%m/%Y"), x["numero_evento"], x["nome"], x["fabricante"] or "", num(x["quantidade_total"]), x["uf"] or "", ABAS_NOME.get(x["aba"], x["aba"])] for x in rows]},
                "link": link}
    ind = {i["chave"]: i for i in calcular_painel(conn, u, 30, None)}
    if consulta == "marcas_sem_fornecedor":
        t = ind["marcas_sem_fornecedor"]["tabela"]
        if not t:
            return {"texto": "Nenhuma marca recorrente está sem fornecedor entre os itens abertos.", "tabela": None, "link": "/os/painel"}
        return {"texto": f"{len(t)} marca(s) com 2 ou mais itens abertos e nenhum fornecedor.",
                "tabela": {"colunas": ["Marca", "Itens abertos"], "linhas": [[x["marca"], x["itens"]] for x in t]}, "link": "/os?sem_fornecedor=1"}
    linhas = []
    for i in ind.values():
        if not i.get("disponivel", True):
            linhas.append([i["titulo"], i.get("motivo") or "indisponível"])
        elif i.get("valor") is not None:
            linhas.append([i["titulo"], f"{i['valor']} {i.get('unidade', '')}".strip()])
    return {"texto": "Indicadores dos últimos 30 dias, conforme o seu perfil.", "tabela": {"colunas": ["Indicador", "Valor"], "linhas": linhas}, "link": "/os/painel"}


def rotas_assistente() -> APIRouter:
    r = APIRouter(prefix="/os/assistente")

    @r.post("")
    def perguntar(body: Pergunta, u: UsuarioOS = Depends(exigir("fila.ver"))):
        texto = body.pergunta.strip()
        if not 2 <= len(texto) <= 500:
            raise HTTPException(status_code=400, detail="Escreva uma pergunta de ate 500 caracteres.")
        escolhido = escolher_consulta_ia(texto)
        consulta, par = escolhido or escolher_consulta(texto)
        with db.conexao() as conn:
            conversa = body.conversa_id
            if conversa and not um(conn, "select 1 x from conversas where id = %s and user_id = %s", (conversa, u.id)):
                raise HTTPException(status_code=404, detail="Conversa nao encontrada.")
            if not conversa:
                conversa = conn.execute("insert into conversas (user_id, titulo) values (%s, %s) returning id", (u.id, texto[:80])).fetchone()[0]
            resp = executar(conn, u, consulta, par)
            conn.execute("insert into mensagens (conversa_id, papel, texto) values (%s, 'usuario', %s)", (conversa, texto))
            # So o texto e guardado; a tabela (que pode ter valores) e remontada pelo sistema a cada consulta.
            conn.execute("insert into mensagens (conversa_id, papel, texto, consulta_usada, parametros) values (%s, 'assistente', %s, %s, %s)",
                         (conversa, resp["texto"], consulta, Jsonb(par)))
        return {"conversa_id": conversa, "consulta": consulta, "motor": "ia" if escolhido else "regras", **resp}

    @r.get("/conversas")
    def conversas(u: UsuarioOS = Depends(exigir("fila.ver"))):
        with db.conexao() as conn:
            return lista(conn, "select id, titulo, criada_em from conversas where user_id = %s order by id desc limit 30", (u.id,))

    @r.get("/conversas/{cid}")
    def conversa(cid: int, u: UsuarioOS = Depends(exigir("fila.ver"))):
        with db.conexao() as conn:
            if not um(conn, "select 1 x from conversas where id = %s and user_id = %s", (cid, u.id)):
                raise HTTPException(status_code=404, detail="Conversa nao encontrada.")
            return lista(conn, "select papel, texto, consulta_usada, criada_em from mensagens where conversa_id = %s order by id", (cid,))

    return r
