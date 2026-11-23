import hashlib
import time
from dataclasses import dataclass

import httpx
import jwt
from fastapi import Depends, Header, HTTPException, Request

_CACHE = {}
_TTL = 30  # segundos


@dataclass(frozen=True)
class Usuario:
    id: str
    email: str


def _validar_no_supabase(cfg, token: str):
    """Pergunta ao Supabase Auth quem e o dono do token. Funciona para qualquer tipo de assinatura de JWT
    e nao precisa de segredo, so da chave publica."""
    r = httpx.get(
        f"{cfg.supabase_url}/auth/v1/user",
        headers={"apikey": cfg.supabase_key, "Authorization": f"Bearer {token}"},
        timeout=10,
    )
    if r.status_code != 200:
        return None
    d = r.json()
    if not d.get("id"):
        return None
    return Usuario(id=d["id"], email=d.get("email") or "")


# Conferencia local do login: o Supabase deste projeto assina os tokens com chave assimetrica (ES256),
# entao basta a chave PUBLICA (JWKS) para conferir a assinatura aqui mesmo, sem ir ate o Supabase a
# cada chamada (era a maior parte da demora: 150 a 600 ms por consulta). Se nao der para conferir
# localmente (token em outro formato, chave nova ainda nao baixada), cai na conferencia pelo Supabase.
_JWKS = {"em": 0.0, "chaves": {}}
_JWKS_TTL = 600  # segundos
_INDEFINIDO = object()


def _chaves_publicas(cfg, forcar: bool = False) -> dict:
    agora = time.time()
    if forcar or agora - _JWKS["em"] > _JWKS_TTL:
        _JWKS["em"] = agora  # tambem segura nova tentativa por 10 min se o Supabase nao responder
        try:
            r = httpx.get(f"{cfg.supabase_url}/auth/v1/.well-known/jwks.json", timeout=5)
            r.raise_for_status()
            _JWKS["chaves"] = {k["kid"]: jwt.PyJWK(k) for k in r.json().get("keys", []) if k.get("kid")}
        except Exception:
            pass
    return _JWKS["chaves"]


def _validar_local(cfg, token: str):
    """Usuario se o token confere; None se e invalido; _INDEFINIDO se nao da para decidir aqui."""
    if token.count(".") != 2:
        return _INDEFINIDO
    try:
        cab = jwt.get_unverified_header(token)
    except jwt.PyJWTError:
        return None
    if cab.get("alg") not in ("ES256", "RS256", "EdDSA"):
        return _INDEFINIDO  # segredo compartilhado (HS256): so o Supabase confere
    chaves = _chaves_publicas(cfg)
    if cab.get("kid") not in chaves:
        chaves = _chaves_publicas(cfg, forcar=True) if time.time() - _JWKS["em"] > 30 else chaves
        if cab.get("kid") not in chaves:
            return _INDEFINIDO
    chave = chaves[cab["kid"]]
    try:
        d = jwt.decode(token, chave.key, algorithms=[cab["alg"]], audience="authenticated",
                       issuer=f"{cfg.supabase_url}/auth/v1", leeway=30)
    except jwt.PyJWTError:
        return None
    if not d.get("sub") or d.get("role") != "authenticated":
        return None
    return Usuario(id=d["sub"], email=d.get("email") or "")


def usuario_atual(request: Request, authorization: str = Header(default="")) -> Usuario:
    if not authorization.lower().startswith("bearer "):
        raise HTTPException(status_code=401, detail="Login necessario.")
    token = authorization[7:].strip()
    chave = hashlib.sha256(token.encode()).hexdigest()
    agora = time.time()
    em_cache = _CACHE.get(chave)
    if em_cache and agora - em_cache[0] < _TTL:
        return em_cache[1]
    cfg = request.app.state.cfg
    try:
        u = _validar_local(cfg, token)
        if u is _INDEFINIDO:
            u = _validar_no_supabase(cfg, token)
    except httpx.HTTPError:
        raise HTTPException(status_code=503, detail="Nao foi possivel validar o login agora.")
    if u is None:
        raise HTTPException(status_code=401, detail="Login invalido ou expirado.")
    if len(_CACHE) > 500:
        _CACHE.clear()
    _CACHE[chave] = (agora, u)
    return u


def limpar_cache() -> None:
    _CACHE.clear()
    _JWKS.update(em=0.0, chaves={})
