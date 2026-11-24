"""Spec 08: assistente somente leitura.

Modo agente (quando OS_N8N_ASSISTENTE_URL esta configurada): um agente de IA no n8n (OpenRouter) conversa com o usuario e
consulta os dados chamando POST /os/assistente/ferramenta. A ferramenta so executa a lista fechada de consultas abaixo,
sempre com o perfil de quem perguntou (contas e visibilidade de preco/margem). O agente nao tem acesso ao banco nem SQL.
O acesso da ferramenta e por um codigo de sessao aleatorio, de uso curto (poucos minutos), criado a cada pergunta.
Decisao do cliente (24/11/2026): o agente recebe tudo o que o perfil do usuario ve, inclusive preco e margem quando o
perfil permite. Sem URL, com erro ou com resposta invalida, o assistente cai nas regras de `escolher_consulta`.
"""
import hashlib
import os
import re
import secrets
import unicodedata
from datetime import date, timedelta
from typing import Optional

import httpx
from fastapi import APIRouter, Depends, Header, HTTPException, Request
from pydantic import BaseModel
from psycopg.types.json import Jsonb

from . import db
from .acesso import UsuarioOS, carregar_usuario_os, exigir
from .permissoes import permissoes_do_perfil
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


DIAS_PT = ("segunda-feira", "terca-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sabado", "domingo")
SESSAO_MINUTOS = 4
SESSAO_MAX_CHAMADAS = 8


class Ferramenta(BaseModel):
    consulta: str
    termo: Optional[str] = None
    ate: Optional[str] = None
    sem_fornecedor: Optional[object] = None


def _hash(token: str) -> str:
    return hashlib.sha256(token.encode()).hexdigest()


def _url_publica(request: Optional[Request]) -> str:
    fixa = os.getenv("OS_API_URL_PUBLICA", "").strip().rstrip("/")
    if fixa or request is None:
        return fixa
    host = request.headers.get("x-forwarded-host") or request.headers.get("host") or ""
    local = host.startswith(("localhost", "127.0.0.1"))
    return f"{'http' if local else 'https'}://{host}"


def chamar_agente(pergunta: str, historico: list, token: str, api_url: str, u: UsuarioOS, hoje: Optional[date] = None) -> Optional[str]:
    """Pede a resposta ao agente do n8n. Devolve o texto ou None (para usar as regras)."""
    url = os.getenv("OS_N8N_ASSISTENTE_URL", "").strip()
    if not url or not api_url:
        return None
    hoje = hoje or date.today()
    corpo = {"pergunta": pergunta, "historico": historico, "hoje": hoje.isoformat(), "dia_da_semana": DIAS_PT[hoje.weekday()],
             "token": token, "api_url": api_url,
             "usuario": {"nome": (u.nome or "").split(" ")[0], "ve_valores": bool({"preco_compra.ver", "venda.ver", "margem.ver"} & set(u.permissoes))}}
    tok = os.getenv("OS_N8N_TOKEN", "").strip()
    cab = {"X-Seven-Token": tok} if tok else {}
    try:
        resp = httpx.post(url, json=corpo, headers=cab, timeout=float(os.getenv("OS_N8N_TIMEOUT", "50")))
        if resp.status_code != 200:
            return None
        dados = resp.json()
    except (httpx.HTTPError, ValueError):
        return None
    if isinstance(dados, list) and dados:
        dados = dados[0]
    texto = dados.get("resposta") or dados.get("output") if isinstance(dados, dict) else None
    if not isinstance(texto, str) or not texto.strip():
        return None
    return texto.strip()[:6000]


def _par_ferramenta(f: Ferramenta) -> Optional[tuple]:
    par = {}
    if f.termo:
        par["termo"] = f.termo
    if f.ate:
        par["ate"] = f.ate
    sf = f.sem_fornecedor
    if isinstance(sf, str):
        sf = sf.strip().lower() in ("true", "1", "sim", "yes")
    if isinstance(sf, bool):
        par["sem_fornecedor"] = sf
    return validar_escolha({"consulta": f.consulta, "parametros": par})


def _para_agente(resp: dict) -> dict:
    """Resultado em formato simples para o modelo ler."""
    t = resp.get("tabela") or {}
    linhas = [[x.isoformat() if hasattr(x, "isoformat") else x for x in l] for l in (t.get("linhas") or [])]
    return {"resumo": resp.get("texto"), "colunas": t.get("colunas") or [], "linhas": linhas}


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
    def perguntar(body: Pergunta, request: Request, u: UsuarioOS = Depends(exigir("fila.ver"))):
        texto = body.pergunta.strip()
        if not 2 <= len(texto) <= 500:
            raise HTTPException(status_code=400, detail="Escreva uma pergunta de ate 500 caracteres.")
        # 1) conversa, historico e sessao do agente. A conexao e devolvida antes de chamar o n8n, porque o agente
        #    chama a ferramenta desta mesma API enquanto esperamos.
        token = secrets.token_urlsafe(32)
        with db.conexao() as conn:
            conversa = body.conversa_id
            if conversa and not um(conn, "select 1 x from conversas where id = %s and user_id = %s", (conversa, u.id)):
                raise HTTPException(status_code=404, detail="Conversa nao encontrada.")
            if not conversa:
                conversa = conn.execute("insert into conversas (user_id, titulo) values (%s, %s) returning id", (u.id, texto[:80])).fetchone()[0]
            historico = list(reversed(lista(conn, "select papel, texto from mensagens where conversa_id = %s order by id desc limit 12", (conversa,))))
            conn.execute("delete from assistente_sessoes where expira_em < now() - interval '1 day'")
            conn.execute("insert into assistente_sessoes (token_hash, user_id, conversa_id, expira_em) values (%s, %s, %s, now() + make_interval(mins => %s))",
                         (_hash(token), u.id, conversa, SESSAO_MINUTOS))
        # 2) agente (n8n). Sem resposta valida, regras.
        resposta_ia = chamar_agente(texto, historico, token, _url_publica(request), u)
        # 3) tabela e link: refaz aqui a ultima consulta que o agente usou, para a tela mostrar os dados do sistema.
        with db.conexao() as conn:
            sess = um(conn, "update assistente_sessoes set expira_em = now() where token_hash = %s returning ultima_consulta, ultimos_parametros", (_hash(token),))
            if resposta_ia:
                consulta = sess and sess["ultima_consulta"]
                par = (sess and sess["ultimos_parametros"]) or {}
                base = executar(conn, u, consulta, par) if consulta else {"tabela": None, "link": None}
                resp = {"texto": resposta_ia, "tabela": base.get("tabela"), "link": base.get("link")}
            else:
                consulta, par = escolher_consulta(texto)
                resp = executar(conn, u, consulta, par)
            conn.execute("insert into mensagens (conversa_id, papel, texto) values (%s, 'usuario', %s)", (conversa, texto))
            conn.execute("insert into mensagens (conversa_id, papel, texto, consulta_usada, parametros) values (%s, 'assistente', %s, %s, %s)",
                         (conversa, resp["texto"], consulta, Jsonb(par)))
        return {"conversa_id": conversa, "consulta": consulta, "motor": "ia" if resposta_ia else "regras", **resp}

    @r.post("/ferramenta")
    def ferramenta(body: Ferramenta, x_assistente_token: str = Header(default="")):
        """Chamada pelo agente do n8n. So leitura, so a lista fechada, com o perfil de quem perguntou."""
        if not x_assistente_token:
            raise HTTPException(status_code=401, detail="Sessao do assistente ausente.")
        with db.conexao() as conn:
            sess = um(conn, """update assistente_sessoes set chamadas = chamadas + 1
                               where token_hash = %s and expira_em > now() and chamadas < %s returning user_id""",
                      (_hash(x_assistente_token), SESSAO_MAX_CHAMADAS))
            if not sess:
                raise HTTPException(status_code=401, detail="Sessao do assistente invalida ou expirada.")
            dados = carregar_usuario_os(str(sess["user_id"]))
            if dados is None:
                raise HTTPException(status_code=403, detail="Usuario sem acesso.")
            u = UsuarioOS(id=str(sess["user_id"]), email=dados["email"], nome=dados["nome"], perfil=dados["perfil"],
                          permissoes=permissoes_do_perfil(dados["perfil"]), contas=dados["contas"])
            if "fila.ver" not in u.permissoes:
                raise HTTPException(status_code=403, detail="Usuario sem acesso.")
            escolha = _par_ferramenta(body)
            if not escolha:
                return {"erro": "Consulta ou parametros invalidos. Use uma das consultas: " + ", ".join(CONSULTAS)
                        + ". historico_produto, fornecedores_produto e buscar_produto exigem termo; itens_por_prazo exige ate (AAAA-MM-DD)."}
            consulta, par = escolha
            conn.execute("update assistente_sessoes set ultima_consulta = %s, ultimos_parametros = %s where token_hash = %s",
                         (consulta, Jsonb(par), _hash(x_assistente_token)))
            return _para_agente(executar(conn, u, consulta, par))

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
