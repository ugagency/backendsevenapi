"""Credenciais Coupa guardadas no banco (tabela credenciais_coupa), com a senha cifrada (Fernet).

A chave fica so no servidor (CREDENCIAIS_CHAVE). O banco guarda apenas o texto cifrado e o
front nunca recebe a senha de volta. Enquanto uma conta nao tiver credencial no banco, vale a
do .env (VALE_USER / VALE_PASS), para a troca ser gradual e sem parar o robo."""
import os
from cryptography.fernet import Fernet, InvalidToken

TABELA = "credenciais_coupa"
CAMPOS = "conta_id,nome,usuario,senha_cifrada,atualizado_em,atualizado_por"


def _fernet():
    chave = (os.getenv("CREDENCIAIS_CHAVE") or "").strip()
    if not chave:
        return None
    try:
        return Fernet(chave.encode())
    except Exception:
        raise RuntimeError("CREDENCIAIS_CHAVE invalida: gere uma com Fernet.generate_key()")


def chave_configurada() -> bool:
    try:
        return _fernet() is not None
    except RuntimeError:
        return False


def cifrar(senha: str) -> str:
    f = _fernet()
    if f is None:
        raise RuntimeError("CREDENCIAIS_CHAVE nao configurada no servidor")
    return f.encrypt(senha.encode("utf-8")).decode("ascii")


def decifrar(token: str) -> str:
    f = _fernet()
    if f is None:
        raise RuntimeError("CREDENCIAIS_CHAVE nao configurada no servidor")
    try:
        return f.decrypt(token.encode("ascii")).decode("utf-8")
    except InvalidToken:
        raise RuntimeError("senha do banco nao pode ser lida (a CREDENCIAIS_CHAVE mudou?)")


def ler_linhas(cliente) -> dict:
    dados = cliente.table(TABELA).select(CAMPOS).execute().data or []
    return {r["conta_id"]: r for r in dados}


def mesclar(contas_env: list, linhas: dict) -> list:
    """Junta as contas do .env com as linhas do banco. Para usuario e senha, o banco vence.
    Senha do banco ilegivel NAO cai para o .env: a conta fica sem senha e com o erro visivel,
    para ninguem rodar com uma senha antiga sem saber."""
    saida = []
    vistos = set()
    for c in contas_env:
        vistos.add(c["id"])
        saida.append(_uma(c["id"], c.get("nome"), c.get("user", ""), c.get("senha", ""), linhas.get(c["id"])))
    for conta_id in sorted(set(linhas) - vistos):
        saida.append(_uma(conta_id, None, "", "", linhas[conta_id]))
    return saida


def _uma(conta_id, nome_env, user_env, senha_env, linha):
    linha = linha or {}
    c = {
        "id": conta_id,
        "nome": (nome_env or linha.get("nome") or conta_id),
        "user": user_env, "fonte_usuario": "env" if user_env else None,
        "senha": senha_env, "fonte_senha": "env" if senha_env else None,
        "atualizado_em": linha.get("atualizado_em"),
        "atualizado_por": linha.get("atualizado_por"),
        "erro": None,
    }
    if linha.get("usuario"):
        c["user"], c["fonte_usuario"] = linha["usuario"], "banco"
    if linha.get("senha_cifrada"):
        try:
            c["senha"], c["fonte_senha"] = decifrar(linha["senha_cifrada"]), "banco"
        except RuntimeError as e:
            c["senha"], c["fonte_senha"], c["erro"] = "", None, str(e)
    return c


def pronta(c: dict) -> bool:
    return bool(c["user"] and c["senha"])
