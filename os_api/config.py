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
    convite_url: str = ""  # Edge Function os-convidar no Supabase (unica que cria login)
    convite_redirect: str = ""  # pagina do front onde o convidado define a senha
    cors_regex: str = ""  # opcional: previas da Vercel, cujo endereco muda a cada versao

    @staticmethod
    def do_ambiente() -> "Config":
        schema = os.getenv("OS_SCHEMA", "os_homolog")
        if schema not in SCHEMAS_PERMITIDOS:
            raise RuntimeError(f"OS_SCHEMA invalido: {schema!r}. Use um de {SCHEMAS_PERMITIDOS}.")
        if os.getenv("SUPABASE_SECRET_KEY"):
            raise RuntimeError("O os_api nao deve receber SUPABASE_SECRET_KEY. Remova do ambiente deste servico.")
        origens = tuple(o.strip() for o in os.getenv("OS_CORS_ORIGINS", "").split(",") if o.strip())
        supabase_url = os.getenv("SUPABASE_URL", "").rstrip("/")
        return Config(
            database_url=os.getenv("OS_DATABASE_URL", ""),
            schema=schema,
            supabase_url=supabase_url,
            supabase_key=os.getenv("SUPABASE_KEY", ""),
            cors_origens=origens,
            convite_url=os.getenv("OS_CONVITE_URL", f"{supabase_url}/functions/v1/os-convidar" if supabase_url else ""),
            convite_redirect=os.getenv("OS_CONVITE_REDIRECT", ""),
            cors_regex=os.getenv("OS_CORS_REGEX", ""),
        )
