import time
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


# Cadastro do OS em memoria por poucos segundos: toda chamada precisa dele, e cada ida ao banco
# atravessa a rede ate o Supabase. Qualquer alteracao na administracao limpa o cache (main.py).
_CACHE_OS = {}
_TTL_OS = 20  # segundos


def limpar_cache_os() -> None:
    _CACHE_OS.clear()


def carregar_usuario_os(user_id: str):
    """Le o cadastro do OS numa unica consulta. Devolve None se a pessoa nao existe ou esta inativa."""
    agora = time.time()
    em_cache = _CACHE_OS.get(user_id)
    if em_cache and agora - em_cache[0] < _TTL_OS:
        return em_cache[1]
    with db.conexao() as conn:
        linhas = conn.execute(
            "select u.nome, u.email, u.perfil, u.ativo, c.id, c.nome, c.cnpj, c.filial "
            "from usuarios u left join usuario_contas uc on uc.user_id = u.user_id "
            "left join contas_portal c on c.id = uc.conta_id and c.ativa "
            "where u.user_id = %s order by c.id", (user_id,)
        ).fetchall()
    if not linhas or not linhas[0][3]:
        dados = None
    else:
        contas = [l for l in linhas if l[4] is not None]
        dados = {"nome": linhas[0][0] or "", "email": linhas[0][1] or "", "perfil": linhas[0][2],
                 "contas": tuple(c[4] for c in contas),
                 "contas_detalhe": tuple({"id": c[4], "nome": c[5], "cnpj": c[6], "filial": c[7]} for c in contas)}
    if len(_CACHE_OS) > 500:
        _CACHE_OS.clear()
    _CACHE_OS[user_id] = (agora, dados)
    return dados


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
