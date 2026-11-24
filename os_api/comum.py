"""Ajudas comuns: consultas com linhas em dicionario e regras de exibicao de valores por perfil."""
from decimal import Decimal

from fastapi import HTTPException
from psycopg.rows import dict_row

from .acesso import UsuarioOS


def lista(conn, sql, params=None):
    with conn.cursor(row_factory=dict_row) as cur:
        cur.execute(sql, params)
        return cur.fetchall()


def um(conn, sql, params=None):
    with conn.cursor(row_factory=dict_row) as cur:
        cur.execute(sql, params)
        return cur.fetchone()


def origem(u: UsuarioOS) -> str:
    return "suporte" if u.perfil == "suporte" else "tela"


def pode(u: UsuarioOS, permissao: str) -> bool:
    return permissao in u.permissoes


def exigir_conta(u: UsuarioOS, contas) -> None:
    """Item so e visivel para quem tem ao menos uma das contas dele liberada."""
    if not set(contas or []) & set(u.contas):
        raise HTTPException(status_code=404, detail="Item nao encontrado.")


def num(v):
    if v is None:
        return None
    return float(v) if isinstance(v, Decimal) else v


def margem_pct(compra, ipi, frete, venda):
    """Margem sobre a venda. O sistema nao calcula o preco de venda; so mostra a margem resultante."""
    if compra is None or venda is None or float(venda) == 0:
        return None
    custo = float(compra) * (1 + float(ipi or 0) / 100) + float(frete or 0)
    return round((float(venda) - custo) / float(venda) * 100, 2)


def valores_visiveis(u: UsuarioOS, v: dict) -> dict:
    """Unico lugar que decide que valor sai para cada perfil. Usado em tela, exportacao e assistente."""
    saida = {}
    if pode(u, "preco_compra.ver"):
        saida.update(preco_compra=num(v.get("preco_compra")), ipi=num(v.get("ipi")), frete=num(v.get("frete")))
    if pode(u, "venda.ver"):
        saida["preco_venda"] = num(v.get("preco_venda"))
    if pode(u, "margem.ver"):
        saida["margem_pct"] = margem_pct(v.get("preco_compra"), v.get("ipi"), v.get("frete"), v.get("preco_venda"))
    return saida


_ULTIMA_SYNC = 0.0


def sincronizar_se_preciso():
    """Traz para o OS o que o robo gravou, sem ninguem clicar. Incremental, protegido por trava no banco e
    limitado a uma vez por OS_SYNC_AUTO_SEG segundos (padrao 60; 0 desliga). Nunca derruba a tela que pediu."""
    import os
    import time
    global _ULTIMA_SYNC
    from . import db
    intervalo = float(os.getenv("OS_SYNC_AUTO_SEG", "60"))
    if intervalo <= 0 or time.monotonic() - _ULTIMA_SYNC < intervalo:
        return
    _ULTIMA_SYNC = time.monotonic()
    try:
        with db.conexao() as conn:
            conn.execute("select sincronizar(null)").fetchone()
    except Exception:  # a leitura da tela segue com o que ja esta no OS
        pass
