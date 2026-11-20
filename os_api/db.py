from contextlib import contextmanager

from psycopg_pool import ConnectionPool

from .config import Config

_pool = None


def iniciar(cfg: Config) -> None:
    """Abre o pool. Cada conexao fica com search_path restrito ao schema do ambiente."""
    global _pool
    if _pool is not None:
        return
    _pool = ConnectionPool(
        cfg.database_url,
        min_size=1,
        max_size=5,
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
        raise RuntimeError("Pool do banco nao iniciado.")
    with _pool.connection() as conn:
        yield conn


def testar() -> bool:
    try:
        with conexao() as conn:
            conn.execute("select 1")
        return True
    except Exception:
        return False
