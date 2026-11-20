from psycopg.types.json import Jsonb

ORIGENS = ("tela", "exportacao", "assistente", "robo", "carga", "suporte")
ACOES = ("criar", "alterar", "excluir", "exportar", "ver_restrito")


def auditar(conn, *, origem, entidade, acao, user_id=None, entidade_id=None,
            antes=None, depois=None, justificativa=None) -> None:
    """Unica porta de gravacao da auditoria. Toda rota que grava deve chamar esta funcao na mesma
    transacao da mudanca, para que mudanca e registro entrem juntos ou nenhum entre."""
    if origem not in ORIGENS:
        raise ValueError(f"origem invalida: {origem!r}")
    if acao not in ACOES:
        raise ValueError(f"acao invalida: {acao!r}")
    if not entidade:
        raise ValueError("entidade obrigatoria")
    conn.execute(
        "insert into auditoria (user_id, origem, entidade, entidade_id, acao, antes, depois, justificativa) "
        "values (%s, %s, %s, %s, %s, %s, %s, %s)",
        (
            user_id, origem, entidade,
            None if entidade_id is None else str(entidade_id), acao,
            None if antes is None else Jsonb(antes),
            None if depois is None else Jsonb(depois),
            justificativa,
        ),
    )
