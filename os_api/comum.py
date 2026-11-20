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
