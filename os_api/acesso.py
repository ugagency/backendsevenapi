from dataclasses import dataclass
from typing import Tuple

from fastapi import Depends, HTTPException

from . import db
from .auth import Usuario, usuario_atual
from .permissoes import permissoes_do_perfil


@dataclass(frozen=True)
class UsuarioOS:
    id: str
    email: str
    nome: str
    perfil: str
    permissoes: frozenset
    contas: Tuple[int, ...]
    contas_detalhe: Tuple[dict, ...] = ()


def carregar_usuario_os(user_id: str):
    """Le o cadastro do OS. Devolve None se a pessoa nao existe ou esta inativa."""
    with db.conexao() as conn:
        linha = conn.execute(
            "select nome, email, perfil, ativo from usuarios where user_id = %s", (user_id,)
        ).fetchone()
        if not linha or not linha[3]:
            return None
        contas = conn.execute(
            "select uc.conta_id, c.nome, c.cnpj, c.filial from usuario_contas uc join contas_portal c on c.id = uc.conta_id "
            "where uc.user_id = %s and c.ativa order by uc.conta_id", (user_id,)
        ).fetchall()
    return {"nome": linha[0] or "", "email": linha[1] or "", "perfil": linha[2],
            "contas": tuple(c[0] for c in contas),
            "contas_detalhe": tuple({"id": c[0], "nome": c[1], "cnpj": c[2], "filial": c[3]} for c in contas)}


def usuario_os(u: Usuario = Depends(usuario_atual)) -> UsuarioOS:
    dados = carregar_usuario_os(u.id)
    if dados is None:
        # Login valido no Supabase mas sem cadastro ativo no OS: nao entra.
        raise HTTPException(status_code=403, detail="Sem acesso ao Seven OS. Peca a ativacao ao administrador.")
    return UsuarioOS(id=u.id, email=dados["email"] or u.email, nome=dados["nome"], perfil=dados["perfil"],
                     permissoes=permissoes_do_perfil(dados["perfil"]), contas=dados["contas"],
                     contas_detalhe=dados.get("contas_detalhe") or tuple({"id": c} for c in dados["contas"]))


def exigir(permissao: str):
    """Dependencia de rota: exige uma permissao. Uso: Depends(exigir('margem.ver'))."""
    def _checar(u: UsuarioOS = Depends(usuario_os)) -> UsuarioOS:
        if permissao not in u.permissoes:
            raise HTTPException(status_code=403, detail="Sem permissao para esta acao.")
        return u
    return _checar
