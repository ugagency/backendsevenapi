import pytest
from fastapi.testclient import TestClient

from os_api import auth, db
from os_api.auditoria import auditar
from os_api.auth import Usuario
from os_api.config import Config
from os_api.main import criar_app


def cfg(**kw):
    base = dict(database_url="", schema="os_homolog", supabase_url="https://x.supabase.co",
                supabase_key="pub", cors_origens=())
    base.update(kw)
    return Config(**base)


@pytest.fixture(autouse=True)
def _limpar():
    auth.limpar_cache()
    yield
    auth.limpar_cache()


def cliente():
    return TestClient(criar_app(cfg(), iniciar_banco=False))


def test_saude_sem_login(monkeypatch):
    monkeypatch.setattr(db, "testar", lambda: True)
    r = cliente().get("/os/saude")
    assert r.status_code == 200
    assert r.json()["schema"] == "os_homolog" and r.json()["banco"] is True


def test_rota_protegida_sem_token():
    assert cliente().get("/os/eu").status_code == 401


def test_rota_protegida_token_invalido(monkeypatch):
    monkeypatch.setattr(auth, "_validar_no_supabase", lambda c, t: None)
    r = cliente().get("/os/eu", headers={"Authorization": "Bearer ruim"})
    assert r.status_code == 401


def test_rota_protegida_token_valido(monkeypatch):
    monkeypatch.setattr(auth, "_validar_no_supabase", lambda c, t: Usuario("u1", "a@b.com"))
    r = cliente().get("/os/eu", headers={"Authorization": "Bearer bom"})
    assert r.status_code == 200 and r.json() == {"id": "u1", "email": "a@b.com"}


def test_auth_fora_do_ar_da_503(monkeypatch):
    import httpx

    def falha(c, t):
        raise httpx.ConnectError("x")
    monkeypatch.setattr(auth, "_validar_no_supabase", falha)
    r = cliente().get("/os/eu", headers={"Authorization": "Bearer bom"})
    assert r.status_code == 503


def test_cache_nao_repete_chamada(monkeypatch):
    n = {"c": 0}

    def v(c, t):
        n["c"] += 1
        return Usuario("u1", "a@b.com")
    monkeypatch.setattr(auth, "_validar_no_supabase", v)
    c = cliente()
    for _ in range(3):
        c.get("/os/eu", headers={"Authorization": "Bearer bom"})
    assert n["c"] == 1


def test_rota_inexistente_nao_vaza():
    assert cliente().get("/os/qualquer").status_code in (401, 404)


def test_recusa_chave_secreta(monkeypatch):
    monkeypatch.setenv("SUPABASE_SECRET_KEY", "x")
    with pytest.raises(RuntimeError):
        Config.do_ambiente()


def test_schema_invalido(monkeypatch):
    monkeypatch.delenv("SUPABASE_SECRET_KEY", raising=False)
    monkeypatch.setenv("OS_SCHEMA", "public")
    with pytest.raises(RuntimeError):
        Config.do_ambiente()


class FakeConn:
    def __init__(self):
        self.chamadas = []

    def execute(self, sql, params=None):
        self.chamadas.append((sql, params))


def test_auditar_grava():
    c = FakeConn()
    auditar(c, origem="tela", entidade="fornecedor", acao="criar", user_id="u1", entidade_id=7,
            depois={"nome": "X"})
    sql, p = c.chamadas[0]
    assert sql.startswith("insert into auditoria")
    assert p[0] == "u1" and p[1] == "tela" and p[3] == "7" and p[4] == "criar"


@pytest.mark.parametrize("kw", [
    dict(origem="outra", entidade="x", acao="criar"),
    dict(origem="tela", entidade="x", acao="apagar"),
    dict(origem="tela", entidade="", acao="criar"),
])
def test_auditar_valida(kw):
    with pytest.raises(ValueError):
        auditar(FakeConn(), **kw)
