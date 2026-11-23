"""Conferencia local do login com a chave publica (ES256), sem ir ao Supabase."""
import json
import time

import jwt
import pytest
from cryptography.hazmat.primitives.asymmetric import ec

from os_api import auth
from os_api.config import Config

URL = "https://proj.supabase.co"


@pytest.fixture
def chave(monkeypatch):
    priv = ec.generate_private_key(ec.SECP256R1())
    pub = json.loads(jwt.algorithms.ECAlgorithm.to_jwk(priv.public_key()))
    pub.update(kid="k1", alg="ES256")
    chamadas = []

    class R:
        def raise_for_status(self): pass
        def json(self): return {"keys": [pub]}

    monkeypatch.setattr(auth.httpx, "get", lambda url, **k: chamadas.append(url) or R())
    monkeypatch.setattr(auth, "_validar_no_supabase", lambda c, t: pytest.fail("nao deveria ir ao Supabase"))
    auth.limpar_cache()
    yield priv, chamadas
    auth.limpar_cache()


def _token(priv, **extra):
    d = {"sub": "u1", "email": "a@b.com", "role": "authenticated", "aud": "authenticated",
         "iss": f"{URL}/auth/v1", "exp": int(time.time()) + 600, **extra}
    return jwt.encode(d, priv, algorithm="ES256", headers={"kid": "k1"})


def cfg():
    return Config(database_url="", schema="os_homolog", supabase_url=URL, supabase_key="pk", cors_origens=())


def test_token_valido_confere_local_e_baixa_a_chave_uma_vez(chave):
    priv, chamadas = chave
    assert auth._validar_local(cfg(), _token(priv)).id == "u1"
    assert auth._validar_local(cfg(), _token(priv, sub="u2")).id == "u2"
    assert chamadas == [f"{URL}/auth/v1/.well-known/jwks.json"]


@pytest.mark.parametrize("extra", [{"exp": int(time.time()) - 3600}, {"aud": "anon"}, {"iss": "https://outro/auth/v1"}, {"role": "anon"}])
def test_token_vencido_ou_de_outro_lugar_e_recusado(chave, extra):
    priv, _ = chave
    assert auth._validar_local(cfg(), _token(priv, **extra)) is None


def test_assinatura_de_outra_chave_e_recusada(chave):
    outra = ec.generate_private_key(ec.SECP256R1())
    assert auth._validar_local(cfg(), _token(outra)) is None


def test_formato_desconhecido_vai_para_o_supabase(chave):
    assert auth._validar_local(cfg(), "nao-e-jwt") is auth._INDEFINIDO
    hs = jwt.encode({"sub": "u1"}, "segredo", algorithm="HS256")
    assert auth._validar_local(cfg(), hs) is auth._INDEFINIDO
