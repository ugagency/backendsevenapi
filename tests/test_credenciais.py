"""Testes das credenciais no banco (cifra, mescla com .env, rotas). Sem rede: o Supabase e um stub.
Rodar na raiz do backend:  python tests/test_credenciais.py"""
import os, sys
RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)
from cryptography.fernet import Fernet

os.environ.update({
    "SUPABASE_URL": "https://x.supabase.co", "SUPABASE_KEY": "x" * 40, "SUPABASE_SECRET_KEY": "s" * 40,
    "CREDENCIAIS_CHAVE": Fernet.generate_key().decode(),
    "VALE_USER": "u1", "VALE_PASS": "p1", "VALE_CONTA_NOME": "Seven Matriz",
    "VALE_CONTA_NOME_2": "Seven Filial 2",          # filial2 sem usuario/senha no .env
})
for k in ("VALE_USER_2", "VALE_PASS_2", "ROBO_SIMULAR"):
    os.environ.pop(k, None)

class Tabela:
    def __init__(self, db): self.db, self.op, self.filtro = db, None, None
    def select(self, *_): self.op = "select"; return self
    def eq(self, col, val): self.filtro = (col, val); return self
    def limit(self, n): return self
    def upsert(self, linha, on_conflict=None):
        self.op = "upsert"; atual = self.db.setdefault(linha["conta_id"], {}); atual.update(linha); return self
    def execute(self):
        class R: pass
        linhas = [dict(v) for v in self.db.values()]
        if self.filtro:
            linhas = [x for x in linhas if x.get(self.filtro[0]) == self.filtro[1]]
        r = R(); r.data = linhas if self.op == "select" else []
        return r

class Auth:
    def get_user(self, token):
        class U: pass
        mapa = {"token-admin": ("u-admin", "ADMIN@seven.com"), "token-comum": ("u-comum", "comum@seven.com"),
                "token-inativo": ("u-inativo", "ex@seven.com"), "token-sem-perfil": ("u-novo", "novo@seven.com")}
        if token in mapa:
            u = U(); u.user = U(); u.user.id, u.user.email = mapa[token]; return u
        raise Exception("invalido")

class Cliente:
    def __init__(self):
        self.db = {}; self.auth = Auth()
        self.perfis = {"u-admin": {"user_id": "u-admin", "papel": "admin", "ativo": True},
                       "u-comum": {"user_id": "u-comum", "papel": "operador", "ativo": True},
                       "u-inativo": {"user_id": "u-inativo", "papel": "admin", "ativo": False}}
    def table(self, nome):
        assert nome in ("credenciais_coupa", "perfis"), nome
        return Tabela(self.db if nome == "credenciais_coupa" else self.perfis)

CLIENTE = Cliente()
import supabase as _s
_s.create_client = lambda url, key: CLIENTE
import main_api as m
import credenciais as cr
from fastapi.testclient import TestClient
api = TestClient(m.app)
ADM = {"Authorization": "Bearer token-admin"}

# 1) so o .env: filial1 pronta, filial2 conhecida mas sem senha (nao aparece para rodar)
assert list(m.CONTAS_COUPA) == ["filial1"], list(m.CONTAS_COUPA)
assert [c["id"] for c in api.get("/contas").json()["contas"]] == ["filial1"]
print("1) ok - conta sem senha nao aparece para rodar")

# 2) permissoes
assert api.get("/credenciais").status_code == 401
assert api.get("/credenciais", headers={"Authorization": "Bearer lixo"}).status_code == 401
assert api.get("/credenciais", headers={"Authorization": "Bearer token-comum"}).status_code == 403
assert api.get("/eu", headers={"Authorization": "Bearer token-comum"}).json()["admin"] is False
eu = api.get("/eu", headers=ADM).json()
assert eu["admin"] is True and eu["papel"] == "admin", eu
assert api.get("/credenciais", headers={"Authorization": "Bearer token-inativo"}).status_code == 403   # admin desativado
assert api.get("/credenciais", headers={"Authorization": "Bearer token-sem-perfil"}).status_code == 403  # sem perfil
assert api.get("/eu", headers={"Authorization": "Bearer token-sem-perfil"}).json()["admin"] is False
assert api.put("/credenciais/filial2", json={"senha": "x"}, headers={"Authorization": "Bearer token-comum"}).status_code == 403
print("2) ok - so papel admin ativo (tabela perfis) ve e altera")

# 3) salvar senha e usuario da filial2: fica cifrada no banco, nunca volta para o front
r = api.put("/credenciais/filial2", json={"usuario": " vendas@seven.com ", "senha": "Segredo#2026"}, headers=ADM)
assert r.status_code == 200, r.text
assert "Segredo" not in r.text and "senha_cifrada" not in r.text
linha = CLIENTE.db["filial2"]
assert linha["senha_cifrada"].startswith("gAAAA") and "Segredo" not in linha["senha_cifrada"]
assert linha["usuario"] == "vendas@seven.com" and linha["atualizado_por"] == "admin@seven.com"
lista = api.get("/credenciais", headers=ADM)
assert "Segredo" not in lista.text and "gAAAA" not in lista.text
f2 = [c for c in lista.json()["contas"] if c["id"] == "filial2"][0]
assert f2["pronta"] and f2["fonte_senha"] == "banco" and f2["tem_senha"]
assert [c["id"] for c in api.get("/contas").json()["contas"]] == ["filial1", "filial2"]
assert m.CONTAS_COUPA["filial2"]["senha"] == "Segredo#2026"
print("3) ok - senha cifrada no banco, nao volta ao front, conta passa a rodar")

# 4) banco vence o .env; trocar so a senha mantem o usuario
api.put("/credenciais/filial1", json={"senha": "nova1"}, headers=ADM)
f1 = m.CONTAS_COUPA["filial1"]
assert f1["senha"] == "nova1" and f1["fonte_senha"] == "banco" and f1["user"] == "u1" and f1["fonte_usuario"] == "env"
print("4) ok - banco vence o .env; campo vazio nao apaga")

# 5) validacoes
assert api.put("/credenciais/filial1", json={}, headers=ADM).status_code == 400
assert api.put("/credenciais/naoexiste", json={"senha": "x"}, headers=ADM).status_code == 404
assert api.put("/credenciais/filial1", json={"senha": "x" * 201}, headers=ADM).status_code == 400
print("5) ok - validacoes")

# 6) chave trocada: senha do banco ilegivel NAO cai para o .env
os.environ["CREDENCIAIS_CHAVE"] = Fernet.generate_key().decode()
m.atualizar_contas()
f1 = [c for c in m.CONTAS_TODAS if c["id"] == "filial1"][0]
assert not cr.pronta(f1) and "CREDENCIAIS_CHAVE" in f1["erro"] and "filial1" not in m.CONTAS_COUPA
r = api.post("/run-robot?data=031126&conta=filial1")
assert r.status_code == 500 and "Credenciais" in r.json()["detail"], r.text
print("6) ok - chave trocada bloqueia a conta com erro visivel, sem usar senha antiga")

# 7) banco fora do ar: segue com o .env
class Quebrado(Cliente):
    def table(self, nome):
        if nome == "perfis": return super().table(nome)
        raise ConnectionError("fora")
m.supabase_admin = Quebrado()
m.atualizar_contas()
assert m.CONTAS_COUPA["filial1"]["senha"] == "p1" and m.CONTAS_COUPA["filial1"]["fonte_senha"] == "env"
print("7) ok - banco fora do ar: usa o .env")
print("TODOS OS TESTES PASSARAM")
