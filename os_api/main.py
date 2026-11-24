import time
from contextlib import asynccontextmanager

from fastapi import APIRouter, Depends, FastAPI
from fastapi.middleware.cors import CORSMiddleware

from . import VERSAO, db
from .acesso import UsuarioOS, limpar_cache_os, usuario_os
from .admin import rotas_admin
from .assistente import rotas_assistente
from .captura import rotas_captura
from .cotacao import rotas_cotacao
from .fila import rotas_fila
from .fornecedores import rotas_fornecedores
from .painel import rotas_painel
from .produtos import rotas_produtos
from .relatorios import rotas_relatorios
from .auth import usuario_atual
from .config import Config


LENTO_MS = 500


def criar_app(cfg: Config = None, iniciar_banco: bool = True) -> FastAPI:
    cfg = cfg or Config.do_ambiente()

    @asynccontextmanager
    async def ciclo(app):
        if iniciar_banco and cfg.database_url:
            db.iniciar(cfg)
        yield
        db.encerrar()

    app = FastAPI(title="Seven OS API", version=VERSAO, lifespan=ciclo)
    if iniciar_banco:
        db.configurar(cfg)
    app.state.cfg = cfg
    app.add_middleware(
        CORSMiddleware,
        allow_origins=list(cfg.cors_origens),  # lista vazia = nenhuma origem de navegador
        allow_origin_regex=cfg.cors_regex or None,
        allow_methods=["GET", "POST", "PATCH", "PUT", "DELETE"],
        expose_headers=["Content-Disposition"],
        allow_headers=["Authorization", "Content-Type"],
    )

    @app.middleware("http")
    async def tempo(request, call_next):
        """Mede cada chamada (cabecalho Server-Timing, visivel no F12 do navegador) e avisa no console
        as lentas. Alteracao na administracao limpa o cache de cadastro."""
        t0 = time.perf_counter()
        resposta = await call_next(request)
        ms = (time.perf_counter() - t0) * 1000
        resposta.headers["Server-Timing"] = f"os;dur={ms:.0f}"
        if request.method != "GET" and request.url.path.startswith("/os/admin"):
            limpar_cache_os()
        if ms > LENTO_MS:
            print(f"[os_api] lento {ms:.0f} ms {request.method} {request.url.path}", flush=True)
        return resposta

    # Unica rota sem login.
    publico = APIRouter(prefix="/os")

    @publico.get("/saude")
    def saude():
        return {"versao": VERSAO, "schema": cfg.schema, "banco": db.testar()}

    # Tudo o mais exige login valido (token do Supabase Auth). Perfis e permissoes entram na Spec 01.
    protegido = APIRouter(prefix="/os", dependencies=[Depends(usuario_atual)])

    @protegido.get("/eu")
    def eu(u: UsuarioOS = Depends(usuario_os)):
        return {"id": u.id, "email": u.email, "nome": u.nome, "perfil": u.perfil,
                "permissoes": sorted(u.permissoes), "contas": list(u.contas_detalhe)}

    app.include_router(publico)
    app.include_router(protegido)
    for rotas in (rotas_admin, rotas_fila, rotas_produtos, rotas_fornecedores, rotas_cotacao, rotas_captura,
                  rotas_painel, rotas_assistente, rotas_relatorios):
        app.include_router(rotas(), dependencies=[Depends(usuario_atual)])
    return app


def app_factory() -> FastAPI:
    return criar_app()
