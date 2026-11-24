import os
import threading
from contextlib import contextmanager

from psycopg_pool import ConnectionPool

from .config import Config

_pool = None
_cfg = None
_trava = threading.Lock()


def configurar(cfg: Config) -> None:
    """Guarda a configuracao; o pool abre na primeira consulta. Assim funciona tambem na Vercel,
    onde o servidor pode nao passar pelo inicio (lifespan) antes do primeiro pedido."""
    global _cfg
    _cfg = cfg


def iniciar(cfg: Config) -> None:
    """Abre o pool. Cada conexao fica com search_path restrito ao schema do ambiente.
    OS_POOL_MIN / OS_POOL_MAX ajustam o tamanho (na Vercel: 0 e 2, porque cada instancia tem o seu)."""
    global _pool, _cfg
    with _trava:
        if _pool is not None:
            return
        _cfg = cfg
        _pool = ConnectionPool(
            cfg.database_url,
            min_size=int(os.getenv("OS_POOL_MIN", "1")),
            max_size=int(os.getenv("OS_POOL_MAX", "5")),
            open=True,
            kwargs={"options": f"-c search_path={cfg.schema}"},
        )


def encerrar() -> None:
    global _pool
    if _pool is not None:
        _pool.close()
        _pool = None


@contextmanager
def conexao():
    if _pool is None:
        if _cfg is None or not _cfg.database_url:
            raise RuntimeError("Pool do banco nao iniciado.")
        iniciar(_cfg)
    with _pool.connection() as conn:
        yield conn


def testar() -> bool:
    try:
        with conexao() as conn:
            conn.execute("select 1")
        return True
    except Exception:
        return False
