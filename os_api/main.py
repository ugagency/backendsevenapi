from contextlib import asynccontextmanager

from fastapi import APIRouter, Depends, FastAPI
from fastapi.middleware.cors import CORSMiddleware

from . import VERSAO, db
from .acesso import UsuarioOS, usuario_os
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


def criar_app(cfg: Config = None, iniciar_banco: bool = True) -> FastAPI:
    cfg = cfg or Config.do_ambiente()

    @asynccontextmanager
    async def ciclo(app):
        if iniciar_banco and cfg.database_url:
            db.iniciar(cfg)
        yield
        db.encerrar()

    app = FastAPI(title="Seven OS API", version=VERSAO, lifespan=ciclo)
    app.state.cfg = cfg
    app.add_middleware(
        CORSMiddleware,
        allow_origins=list(cfg.cors_origens),  # lista vazia = nenhuma origem de navegador
        allow_methods=["GET", "POST", "PATCH", "PUT", "DELETE"],
        expose_headers=["Content-Disposition"],
        allow_headers=["Authorization", "Content-Type"],
    )

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
