"""Rotas de administracao da Spec 01. Toda gravacao vai para a auditoria na mesma transacao."""
from typing import Any, List, Optional

from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel

from . import db, validacao
from .acesso import UsuarioOS, exigir
from .auditoria import auditar
from .permissoes import PERFIS


class NovoUsuario(BaseModel):
    email: str
    nome: str = ""
    perfil: str
    ativo: bool = False
    contas: List[int] = []


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


def rotas_admin() -> APIRouter:
    r = APIRouter(prefix="/os/admin")

    # ---------- usuarios ----------
    @r.get("/usuarios")
    def listar_usuarios(u: UsuarioOS = Depends(exigir("admin.usuarios"))):
        with db.conexao() as conn:
            rows = conn.execute(SQL_USUARIO + "group by u.user_id order by u.nome, u.email").fetchall()
        return [_linha_usuario(x) for x in rows]

    @r.post("/usuarios", status_code=201)
    def criar_usuario(body: NovoUsuario, u: UsuarioOS = Depends(exigir("admin.usuarios"))):
        _checar_perfil(body.perfil)
        email = body.email.strip().lower()
        with db.conexao() as conn:
            uid = conn.execute("select auth_id_por_email(%s)", (email,)).fetchone()[0]
            if uid is None:
                raise HTTPException(status_code=404, detail="Essa pessoa ainda nao tem login no sistema. O convite por e-mail ainda nao esta disponivel.")
            uid = str(uid)
            if _usuario(conn, uid):
                raise HTTPException(status_code=409, detail="Essa pessoa ja esta cadastrada no Seven OS.")
            _checar_contas(conn, body.contas)
            conn.execute(
                "insert into usuarios (user_id, nome, email, perfil, ativo, convidado_por) values (%s, %s, %s, %s, %s, %s)",
                (uid, body.nome.strip(), email, body.perfil, body.ativo, u.id))
            _gravar_contas(conn, uid, body.contas)
            depois = _usuario(conn, uid)
            auditar(conn, origem=_origem(u), entidade="usuario", entidade_id=uid, acao="criar", user_id=u.id, depois=depois)
        return depois

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

    return r
