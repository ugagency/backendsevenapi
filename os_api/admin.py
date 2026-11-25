"""Rotas de administracao da Spec 01. Toda gravacao vai para a auditoria na mesma transacao."""
from typing import Any, List, Optional

import httpx
from fastapi import APIRouter, Depends, File, Header, HTTPException, Request, UploadFile
from pydantic import BaseModel

from . import carga_base, db, validacao
from .acesso import UsuarioOS, exigir
from .auditoria import auditar
from .permissoes import PERFIS


class NovoUsuario(BaseModel):
    email: str
    nome: str = ""
    perfil: str
    ativo: bool = False
    contas: List[int] = []
    convite: str = "link"  # "link": devolve o link para o administrador enviar; "email": o Supabase envia


class NovoLink(BaseModel):
    convite: str = "link"


class AlteraUsuario(BaseModel):
    nome: Optional[str] = None
    perfil: Optional[str] = None
    ativo: Optional[bool] = None
    contas: Optional[List[int]] = None


class NovaConta(BaseModel):
    portal: str = "coupa"
    nome: str
    cnpj: Optional[str] = None
    filial: Optional[str] = None
    conta_robo: Optional[str] = None


class AlteraConta(BaseModel):
    nome: Optional[str] = None
    cnpj: Optional[str] = None
    filial: Optional[str] = None
    ativa: Optional[bool] = None


class Valor(BaseModel):
    valor: Any


def _origem(u: UsuarioOS) -> str:
    return "suporte" if u.perfil == "suporte" else "tela"


SQL_USUARIO = (
    "select u.user_id::text, u.nome, u.email, u.perfil, u.ativo, "
    "coalesce(array_agg(uc.conta_id order by uc.conta_id) filter (where uc.conta_id is not null), '{}') "
    "from usuarios u left join usuario_contas uc on uc.user_id = u.user_id "
)


def _linha_usuario(r):
    return {"user_id": r[0], "nome": r[1] or "", "email": r[2] or "", "perfil": r[3], "ativo": r[4], "contas": list(r[5])}


def _usuario(conn, uid):
    r = conn.execute(SQL_USUARIO + "where u.user_id = %s group by u.user_id", (uid,)).fetchone()
    return _linha_usuario(r) if r else None


def _checar_contas(conn, ids):
    if not ids:
        return
    achadas = {r[0] for r in conn.execute("select id from contas_portal where id = any(%s)", (list(ids),)).fetchall()}
    faltam = set(ids) - achadas
    if faltam:
        raise HTTPException(status_code=400, detail=f"Conta inexistente: {sorted(faltam)}")


def _gravar_contas(conn, uid, ids):
    conn.execute("delete from usuario_contas where user_id = %s", (uid,))
    for c in sorted(set(ids)):
        conn.execute("insert into usuario_contas (user_id, conta_id) values (%s, %s)", (uid, c))


def _checar_perfil(perfil):
    if perfil not in PERFIS:
        raise HTTPException(status_code=400, detail=f"Perfil invalido. Use um de {list(PERFIS)}.")


def _convidar(cfg, token: str, email: str, modo: str, reenviar: bool = False, nome: str = "") -> dict:
    """Pede o login novo a Edge Function os-convidar. O os_api nao tem a chave secreta;
    a funcao confere de novo, com o token de quem esta logado, que e um administrador."""
    if not cfg.convite_url:
        raise HTTPException(status_code=503, detail="Convite indisponivel: OS_CONVITE_URL nao configurada.")
    try:
        r = httpx.post(
            cfg.convite_url,
            headers={"apikey": cfg.supabase_key, "Authorization": f"Bearer {token}"},
            json={"email": email, "schema": cfg.schema, "modo": modo, "redirect_to": cfg.convite_redirect or None,
                  **({"reenviar": True} if reenviar else {}), **({"nome": nome} if nome else {})},
            timeout=20,
        )
    except httpx.HTTPError:
        raise HTTPException(status_code=503, detail="Nao foi possivel falar com o servico de convite agora.")
    try:
        d = r.json()
    except ValueError:
        d = {}
    if r.status_code == 404:
        raise HTTPException(status_code=503, detail="Servico de convite ainda nao publicado no Supabase.")
    if r.status_code not in (200, 201) or not d.get("user_id"):
        raise HTTPException(status_code=502, detail=d.get("erro") or f"Convite recusado (status {r.status_code}).")
    return d


def rotas_admin() -> APIRouter:
    r = APIRouter(prefix="/os/admin")

    # ---------- usuarios ----------
    @r.get("/usuarios")
    def listar_usuarios(u: UsuarioOS = Depends(exigir("admin.usuarios"))):
        with db.conexao() as conn:
            rows = conn.execute(SQL_USUARIO + "group by u.user_id order by u.nome, u.email").fetchall()
        return [_linha_usuario(x) for x in rows]

    @r.post("/usuarios", status_code=201)
    def criar_usuario(body: NovoUsuario, request: Request, authorization: str = Header(default=""),
                      u: UsuarioOS = Depends(exigir("admin.usuarios"))):
        _checar_perfil(body.perfil)
        if body.convite not in ("link", "email"):
            raise HTTPException(status_code=400, detail="Convite deve ser 'link' ou 'email'.")
        email = body.email.strip().lower()
        with db.conexao() as conn:
            uid = conn.execute("select auth_id_por_email(%s)", (email,)).fetchone()[0]
            if uid is not None and _usuario(conn, str(uid)):
                raise HTTPException(status_code=409, detail="Essa pessoa ja esta cadastrada no Seven OS.")
            _checar_contas(conn, body.contas)
        convite = None
        cfg = request.app.state.cfg
        # fora da transacao: a chamada externa nao segura conexao do banco. Com login ja existente a funcao
        # so age se ele nunca foi usado (cadastro antigo sem confirmar): ai manda o convite tambem.
        if uid is None or cfg.convite_url:
            convite = _convidar(cfg, authorization[7:].strip(), email, body.convite, nome=body.nome.strip())
            uid = convite["user_id"]
        uid = str(uid)
        with db.conexao() as conn:
            if _usuario(conn, uid):
                raise HTTPException(status_code=409, detail="Essa pessoa ja esta cadastrada no Seven OS.")
            _checar_contas(conn, body.contas)
            conn.execute(
                "insert into usuarios (user_id, nome, email, perfil, ativo, convidado_por) values (%s, %s, %s, %s, %s, %s)",
                (uid, body.nome.strip(), email, body.perfil, body.ativo, u.id))
            _gravar_contas(conn, uid, body.contas)
            depois = _usuario(conn, uid)
            mandou = bool(convite and (convite.get("link") or convite.get("enviado")))
            if mandou:
                depois_aud = {**depois, "convite": body.convite}  # o link nunca vai para a auditoria
            else:
                depois_aud = depois
            auditar(conn, origem=_origem(u), entidade="usuario", entidade_id=uid, acao="criar", user_id=u.id, depois=depois_aud)
        resposta = dict(depois)
        if mandou:
            resposta["convite"] = body.convite
            if convite.get("link"):
                resposta["link_convite"] = convite["link"]
        return resposta

    @r.post("/usuarios/{uid}/link-acesso")
    def novo_link(uid: str, body: NovoLink, request: Request, authorization: str = Header(default=""),
                  u: UsuarioOS = Depends(exigir("admin.usuarios"))):
        """Convite expirado ou senha esquecida: gera (ou manda por e-mail) um link novo para definir a senha."""
        if body.convite not in ("link", "email"):
            raise HTTPException(status_code=400, detail="Convite deve ser 'link' ou 'email'.")
        with db.conexao() as conn:
            alvo = _usuario(conn, uid)
        if not alvo:
            raise HTTPException(status_code=404, detail="Usuario nao encontrado no Seven OS.")
        d = _convidar(request.app.state.cfg, authorization[7:].strip(), alvo["email"], body.convite, reenviar=True)
        with db.conexao() as conn:
            auditar(conn, origem=_origem(u), entidade="usuario", entidade_id=uid, acao="alterar", user_id=u.id,
                    depois={"email": alvo["email"], "link_acesso": body.convite})  # o link nunca vai para a auditoria
        return {"email": alvo["email"], "convite": body.convite, **({"link_convite": d["link"]} if d.get("link") else {})}

    @r.patch("/usuarios/{uid}")
    def alterar_usuario(uid: str, body: AlteraUsuario, u: UsuarioOS = Depends(exigir("admin.usuarios"))):
        if body.perfil is not None:
            _checar_perfil(body.perfil)
        if uid == u.id and (body.perfil is not None or body.ativo is not None):
            raise HTTPException(status_code=400, detail="Voce nao pode mudar o proprio perfil nem se desativar.")
        with db.conexao() as conn:
            antes = _usuario(conn, uid)
            if not antes:
                raise HTTPException(status_code=404, detail="Usuario nao encontrado no Seven OS.")
            perde_admin = antes["perfil"] == "administrador" and antes["ativo"] and (
                (body.perfil is not None and body.perfil != "administrador") or body.ativo is False)
            if perde_admin:
                outros = conn.execute(
                    "select count(*) from usuarios where perfil = 'administrador' and ativo and user_id <> %s", (uid,)
                ).fetchone()[0]
                if outros == 0:
                    raise HTTPException(status_code=409, detail="E o ultimo administrador ativo. Cadastre outro antes.")
            if body.contas is not None:
                _checar_contas(conn, body.contas)
            conn.execute(
                "update usuarios set nome = coalesce(%s, nome), perfil = coalesce(%s, perfil), ativo = coalesce(%s, ativo) where user_id = %s",
                (None if body.nome is None else body.nome.strip(), body.perfil, body.ativo, uid))
            if body.contas is not None:
                _gravar_contas(conn, uid, body.contas)
            depois = _usuario(conn, uid)
            auditar(conn, origem=_origem(u), entidade="usuario", entidade_id=uid, acao="alterar", user_id=u.id, antes=antes, depois=depois)
        return depois

    # ---------- contas e portais ----------
    SQL_CONTA = "select id, portal, conta_robo, nome, cnpj, filial, ativa from contas_portal "

    def _linha_conta(x):
        return {"id": x[0], "portal": x[1], "conta_robo": x[2], "nome": x[3], "cnpj": x[4], "filial": x[5], "ativa": x[6]}

    def _conta(conn, cid):
        x = conn.execute(SQL_CONTA + "where id = %s", (cid,)).fetchone()
        return _linha_conta(x) if x else None

    @r.get("/contas")
    def listar_contas(u: UsuarioOS = Depends(exigir("admin.contas"))):
        with db.conexao() as conn:
            return [_linha_conta(x) for x in conn.execute(SQL_CONTA + "order by id").fetchall()]

    @r.post("/contas", status_code=201)
    def criar_conta(body: NovaConta, u: UsuarioOS = Depends(exigir("admin.contas"))):
        if body.portal not in validacao.PORTAIS:
            raise HTTPException(status_code=400, detail=f"Portal ainda nao suportado. Hoje: {list(validacao.PORTAIS)}.")
        if not body.nome.strip():
            raise HTTPException(status_code=400, detail="Nome obrigatorio.")
        try:
            cnpj = validacao.cnpj(body.cnpj)
        except ValueError as e:
            raise HTTPException(status_code=400, detail=str(e))
        with db.conexao() as conn:
            if body.conta_robo and conn.execute(
                    "select 1 from contas_portal where portal = %s and conta_robo = %s", (body.portal, body.conta_robo)).fetchone():
                raise HTTPException(status_code=409, detail="Ja existe conta com esse identificador do robo.")
            cid = conn.execute(
                "insert into contas_portal (portal, conta_robo, nome, cnpj, filial) values (%s, %s, %s, %s, %s) returning id",
                (body.portal, body.conta_robo, body.nome.strip(), cnpj, body.filial)).fetchone()[0]
            depois = _conta(conn, cid)
            auditar(conn, origem=_origem(u), entidade="conta_portal", entidade_id=cid, acao="criar", user_id=u.id, depois=depois)
        return depois

    @r.patch("/contas/{cid}")
    def alterar_conta(cid: int, body: AlteraConta, u: UsuarioOS = Depends(exigir("admin.contas"))):
        try:
            cnpj = validacao.cnpj(body.cnpj) if body.cnpj is not None else None
        except ValueError as e:
            raise HTTPException(status_code=400, detail=str(e))
        with db.conexao() as conn:
            antes = _conta(conn, cid)
            if not antes:
                raise HTTPException(status_code=404, detail="Conta nao encontrada.")
            conn.execute(
                "update contas_portal set nome = coalesce(%s, nome), cnpj = coalesce(%s, cnpj), "
                "filial = coalesce(%s, filial), ativa = coalesce(%s, ativa) where id = %s",
                (None if body.nome is None else body.nome.strip() or None, cnpj, body.filial, body.ativa, cid))
            depois = _conta(conn, cid)
            auditar(conn, origem=_origem(u), entidade="conta_portal", entidade_id=cid, acao="alterar", user_id=u.id, antes=antes, depois=depois)
        return depois

    # ---------- parametros ----------
    @r.get("/parametros")
    def listar_parametros(u: UsuarioOS = Depends(exigir("admin.parametros"))):
        with db.conexao() as conn:
            rows = conn.execute("select chave, valor, alterado_por::text, alterado_em from parametros order by chave").fetchall()
        atuais = {x[0]: {"valor": x[1], "alterado_por": x[2], "alterado_em": x[3]} for x in rows}
        return {"conhecidos": sorted(validacao.PARAMETROS), "valores": atuais}

    @r.put("/parametros/{chave}")
    def gravar_parametro(chave: str, body: Valor, u: UsuarioOS = Depends(exigir("admin.parametros"))):
        try:
            valor = validacao.parametro(chave, body.valor)
        except KeyError:
            raise HTTPException(status_code=404, detail="Parametro desconhecido.")
        except ValueError as e:
            raise HTTPException(status_code=400, detail=str(e))
        from psycopg.types.json import Jsonb
        with db.conexao() as conn:
            antes = conn.execute("select valor from parametros where chave = %s", (chave,)).fetchone()
            conn.execute(
                "insert into parametros (chave, valor, alterado_por, alterado_em) values (%s, %s, %s, now()) "
                "on conflict (chave) do update set valor = excluded.valor, alterado_por = excluded.alterado_por, alterado_em = now()",
                (chave, Jsonb(valor), u.id))
            auditar(conn, origem=_origem(u), entidade="parametro", entidade_id=chave, acao="alterar", user_id=u.id,
                    antes=None if antes is None else {"valor": antes[0]}, depois={"valor": valor})
        return {"chave": chave, "valor": valor}

    @r.post("/importar-base")
    def importar_base(arquivo: UploadFile = File(...), u: UsuarioOS = Depends(exigir("admin.parametros"))):
        """Carga da base historica da Seven (BASE.xlsx). Administrador ou Suporte SSYS (a resposta so traz contagens,
        nenhum preco). Recarregar o mesmo arquivo nao duplica."""
        if u.perfil not in ("administrador", "suporte"):
            raise HTTPException(status_code=403, detail="So o Administrador ou o Suporte SSYS carrega a base da Seven.")
        nome = (arquivo.filename or "base.xlsx").split("/")[-1].split("\\")[-1]
        if not nome.lower().endswith((".xlsx", ".xlsm")):
            raise HTTPException(status_code=400, detail="Envie a planilha em Excel (.xlsx).")
        conteudo = arquivo.file.read()
        if len(conteudo) > 4_000_000:
            raise HTTPException(status_code=413, detail="Planilha maior que 4 MB. Divida por mes e envie em partes.")
        try:
            with db.conexao() as conn:
                res = carga_base.importar(conn, conteudo, nome, u.id)
                auditar(conn, origem="carga", entidade="historico_importado", acao="criar", user_id=u.id, depois=res)
        except ValueError as e:
            raise HTTPException(status_code=400, detail=str(e))
        return res

    return r
