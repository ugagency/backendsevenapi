import pytest
from fastapi.testclient import TestClient

from os_api import acesso, auth, validacao
from os_api.auth import Usuario
from os_api.config import Config
from os_api.main import criar_app


@pytest.fixture(autouse=True)
def _limpar():
    auth.limpar_cache()
    yield
    auth.limpar_cache()


def cliente(monkeypatch, perfil):
    monkeypatch.setattr(auth, "_validar_no_supabase", lambda c, t: Usuario("u1", "a@b.com"))
    monkeypatch.setattr(acesso, "carregar_usuario_os",
                        lambda uid: {"nome": "A", "email": "a@b.com", "perfil": perfil, "contas": ()})
    cfg = Config(database_url="", schema="os_homolog", supabase_url="https://x", supabase_key="k", cors_origens=())
    c = TestClient(criar_app(cfg, iniciar_banco=False))
    c.headers.update({"Authorization": "Bearer t"})
    return c


ROTAS = [
    ("get", "/os/admin/usuarios", None),
    ("post", "/os/admin/usuarios", {"email": "x@y.com", "perfil": "cotacao"}),
    ("patch", "/os/admin/usuarios/abc", {"nome": "X"}),
    ("get", "/os/admin/contas", None),
    ("post", "/os/admin/contas", {"nome": "X"}),
    ("patch", "/os/admin/contas/1", {"nome": "X"}),
    ("get", "/os/admin/parametros", None),
    ("put", "/os/admin/parametros/robo.agenda", {"valor": {"horarios": []}}),
]


@pytest.mark.parametrize("perfil", ["cotacao", "supervisor"])
@pytest.mark.parametrize("metodo,rota,corpo", ROTAS)
def test_sem_permissao_de_admin_da_403(monkeypatch, perfil, metodo, rota, corpo):
    c = cliente(monkeypatch, perfil)
    r = getattr(c, metodo)(rota, json=corpo) if corpo is not None else getattr(c, metodo)(rota)
    assert r.status_code == 403


def test_sem_login_da_401():
    cfg = Config(database_url="", schema="os_homolog", supabase_url="https://x", supabase_key="k", cors_origens=())
    assert TestClient(criar_app(cfg, iniciar_banco=False)).get("/os/admin/usuarios").status_code == 401


def test_nao_muda_o_proprio_perfil(monkeypatch):
    c = cliente(monkeypatch, "administrador")
    assert c.patch("/os/admin/usuarios/u1", json={"perfil": "cotacao"}).status_code == 400
    assert c.patch("/os/admin/usuarios/u1", json={"ativo": False}).status_code == 400


def test_perfil_invalido(monkeypatch):
    c = cliente(monkeypatch, "suporte")
    assert c.post("/os/admin/usuarios", json={"email": "x@y.com", "perfil": "chefe"}).status_code == 400


def test_portal_nao_suportado_e_cnpj_invalido(monkeypatch):
    c = cliente(monkeypatch, "administrador")
    assert c.post("/os/admin/contas", json={"nome": "X", "portal": "ariba"}).status_code == 400
    assert c.post("/os/admin/contas", json={"nome": "X", "cnpj": "11.222.333/0001-80"}).status_code == 400
    assert c.patch("/os/admin/contas/1", json={"cnpj": "123"}).status_code == 400


def test_parametro_desconhecido_e_invalido(monkeypatch):
    c = cliente(monkeypatch, "suporte")
    assert c.put("/os/admin/parametros/qualquer", json={"valor": 1}).status_code == 404
    assert c.put("/os/admin/parametros/robo.agenda", json={"valor": {"horarios": ["25:00"]}}).status_code == 400


def test_cnpj():
    assert validacao.cnpj("11.222.333/0001-81") == "11222333000181"
    assert validacao.cnpj("") is None
    for ruim in ("11.222.333/0001-80", "11111111111111", "123"):
        with pytest.raises(ValueError):
            validacao.cnpj(ruim)


def test_agenda():
    assert validacao.parametro("robo.agenda", {"horarios": ["13:00", "07:30"]}) == {"horarios": ["07:30", "13:00"]}
    assert validacao.parametro("robo.agenda", {"horarios": []}) == {"horarios": []}
    for ruim in ({"horarios": ["7:00"]}, {"horarios": ["07:00", "07:00"]}, {"horarios": ["07:00"] * 7}, ["07:00"], {"x": []}):
        with pytest.raises(ValueError):
            validacao.parametro("robo.agenda", ruim)


def test_listas_e_percentual():
    assert validacao.parametro("listas.motivos_declinio", [" Sem fornecedor ", "Prazo curto"]) == ["Sem fornecedor", "Prazo curto"]
    with pytest.raises(ValueError):
        validacao.parametro("listas.categorias", ["A", "a"])
    with pytest.raises(ValueError):
        validacao.parametro("alertas.divergencia_pct", True)
    assert validacao.parametro("alertas.divergencia_pct", 15) == 15
