import os
from dataclasses import dataclass

SCHEMAS_PERMITIDOS = ("os", "os_homolog")


@dataclass(frozen=True)
class Config:
    database_url: str
    schema: str
    supabase_url: str
    supabase_key: str  # chave publica (publishable/anon). O os_api nunca usa a chave secreta.
    cors_origens: tuple

    @staticmethod
    def do_ambiente() -> "Config":
        schema = os.getenv("OS_SCHEMA", "os_homolog")
        if schema not in SCHEMAS_PERMITIDOS:
            raise RuntimeError(f"OS_SCHEMA invalido: {schema!r}. Use um de {SCHEMAS_PERMITIDOS}.")
        if os.getenv("SUPABASE_SECRET_KEY"):
            raise RuntimeError("O os_api nao deve receber SUPABASE_SECRET_KEY. Remova do ambiente deste servico.")
        origens = tuple(o.strip() for o in os.getenv("OS_CORS_ORIGINS", "").split(",") if o.strip())
        return Config(
            database_url=os.getenv("OS_DATABASE_URL", ""),
            schema=schema,
            supabase_url=os.getenv("SUPABASE_URL", "").rstrip("/"),
            supabase_key=os.getenv("SUPABASE_KEY", ""),
            cors_origens=origens,
        )
