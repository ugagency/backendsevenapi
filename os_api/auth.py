import hashlib
import time
from dataclasses import dataclass

import httpx
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


def usuario_atual(request: Request, authorization: str = Header(default="")) -> Usuario:
    if not authorization.lower().startswith("bearer "):
        raise HTTPException(status_code=401, detail="Login necessario.")
    token = authorization[7:].strip()
    chave = hashlib.sha256(token.encode()).hexdigest()
    agora = time.time()
    em_cache = _CACHE.get(chave)
    if em_cache and agora - em_cache[0] < _TTL:
        return em_cache[1]
    try:
        u = _validar_no_supabase(request.app.state.cfg, token)
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
