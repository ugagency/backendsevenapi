"""Matriz de permissoes da Spec 01. Lista fixa no codigo; o front so esconde botoes, quem decide e o servidor."""

TODAS = (
    "fila.ver", "captura.rodar", "captura.baixar", "identificacao.resolver", "fornecedor.vincular",
    "preco_compra.ver", "margem.ver", "venda.ver", "cotacao.lancar", "valor_aberto.ver", "exportar",
    "admin.usuarios", "admin.contas", "admin.credenciais", "admin.parametros", "auditoria.ver",
)

_BASE = {"fila.ver", "captura.rodar", "captura.baixar", "exportar"}
_OPERACAO = {"identificacao.resolver", "fornecedor.vincular", "preco_compra.ver", "venda.ver", "cotacao.lancar"}
_GESTAO = {"margem.ver", "valor_aberto.ver", "auditoria.ver"}
_ADMIN = {"admin.usuarios", "admin.contas", "admin.credenciais", "admin.parametros"}

PERMISSOES_POR_PERFIL = {
    "administrador": frozenset(_BASE | _OPERACAO | _GESTAO | _ADMIN),
    "supervisor": frozenset(_BASE | _OPERACAO | _GESTAO),
    # Cotacao: ve preco de compra e digita o de venda (decisao de 20/11); nunca ve margem calculada.
    "cotacao": frozenset(_BASE | _OPERACAO),
    # Suporte SSYS: sem preco (compra, venda, margem), com administracao (sem ver senha) e auditoria.
    "suporte": frozenset(_BASE | _ADMIN | {"auditoria.ver"}),
}

PERFIS = tuple(PERMISSOES_POR_PERFIL)


def permissoes_do_perfil(perfil: str) -> frozenset:
    return PERMISSOES_POR_PERFIL.get(perfil, frozenset())
