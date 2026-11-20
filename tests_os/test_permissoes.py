import pytest

from os_api.permissoes import PERFIS, PERMISSOES_POR_PERFIL, TODAS, permissoes_do_perfil


def tem(perfil, p):
    return p in permissoes_do_perfil(perfil)


def test_perfis_e_permissoes_conhecidas():
    assert set(PERFIS) == {"administrador", "supervisor", "cotacao", "suporte"}
    for perfil, ps in PERMISSOES_POR_PERFIL.items():
        assert ps <= set(TODAS), perfil


def test_perfil_desconhecido_nao_tem_nada():
    assert permissoes_do_perfil("qualquer") == frozenset()


@pytest.mark.parametrize("perfil", ["cotacao", "suporte"])
def test_sem_margem_nem_valor_em_aberto(perfil):
    assert not tem(perfil, "margem.ver")
    assert not tem(perfil, "valor_aberto.ver")


def test_cotacao_ve_compra_e_venda_mas_nao_administra():
    assert tem("cotacao", "preco_compra.ver") and tem("cotacao", "venda.ver") and tem("cotacao", "cotacao.lancar")
    assert not any(tem("cotacao", p) for p in ("admin.usuarios", "admin.contas", "admin.credenciais", "admin.parametros", "auditoria.ver"))


def test_suporte_sem_preco_e_sem_lancar():
    for p in ("preco_compra.ver", "venda.ver", "margem.ver", "cotacao.lancar", "fornecedor.vincular", "identificacao.resolver"):
        assert not tem("suporte", p)
    assert tem("suporte", "admin.usuarios") and tem("suporte", "auditoria.ver")


def test_supervisor_nao_administra():
    assert tem("supervisor", "margem.ver") and tem("supervisor", "auditoria.ver")
    assert not any(tem("supervisor", p) for p in ("admin.usuarios", "admin.contas", "admin.credenciais", "admin.parametros"))


def test_administrador_tem_tudo():
    assert set(TODAS) == set(permissoes_do_perfil("administrador"))
