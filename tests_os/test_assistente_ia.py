"""Escolha de consulta por IA (n8n): so a pergunta sai, a resposta e validada e ha volta para as regras."""
import httpx
from os_api import assistente as a


class _Resp:
    def __init__(self, codigo=200, dados=None):
        self.status_code = codigo
        self._d = dados

    def json(self):
        if self._d is None:
            raise ValueError("sem json")
        return self._d


def test_validar_aceita_so_lista_fechada():
    assert a.validar_escolha({"consulta": "indicadores"}) == ("indicadores", {})
    assert a.validar_escolha({"consulta": "drop table x"}) is None
    assert a.validar_escolha({"consulta": "historico_produto", "parametros": {"termo": "rolamento 6205"}}) == ("historico_produto", {"termo": "rolamento 6205"})
    assert a.validar_escolha({"consulta": "historico_produto", "parametros": {}}) is None
    assert a.validar_escolha({"consulta": "itens_por_prazo", "parametros": {"ate": "2026-11-28", "sem_fornecedor": True}}) == ("itens_por_prazo", {"ate": "2026-11-28", "sem_fornecedor": True})
    assert a.validar_escolha({"consulta": "itens_por_prazo", "parametros": {"ate": "amanha"}}) is None
    # parametro desconhecido (por exemplo SQL) e ignorado
    assert a.validar_escolha({"consulta": "buscar_produto", "parametros": {"termo": "x1", "sql": "select 1"}}) == ("buscar_produto", {"termo": "x1"})


def test_sem_url_usa_regras(monkeypatch):
    monkeypatch.delenv("OS_N8N_ASSISTENTE_URL", raising=False)
    assert a.escolher_consulta_ia("quais indicadores?") is None


def test_ia_manda_so_pergunta_e_data(monkeypatch):
    monkeypatch.setenv("OS_N8N_ASSISTENTE_URL", "https://n8n.exemplo/webhook/x")
    monkeypatch.setenv("OS_N8N_TOKEN", "segredo-teste")
    visto = {}

    def falso(url, json=None, headers=None, timeout=None):
        visto.update(url=url, json=json, headers=headers)
        return _Resp(200, {"consulta": "marcas_sem_fornecedor", "parametros": {}})
    monkeypatch.setattr(a.httpx, "post", falso)
    assert a.escolher_consulta_ia("marcas sem fornecedor") == ("marcas_sem_fornecedor", {})
    assert set(visto["json"]) == {"pergunta", "hoje", "dia_da_semana", "consultas"}
    assert visto["headers"]["X-Seven-Token"] == "segredo-teste"


def test_ia_com_falha_volta_para_regras(monkeypatch):
    monkeypatch.setenv("OS_N8N_ASSISTENTE_URL", "https://n8n.exemplo/webhook/x")
    for falha in (lambda *a, **k: (_ for _ in ()).throw(httpx.ConnectTimeout("t")), lambda *a, **k: _Resp(500, {}),
                  lambda *a, **k: _Resp(200, None), lambda *a, **k: _Resp(200, {"consulta": "inventada"})):
        monkeypatch.setattr(a.httpx, "post", falha)
        assert a.escolher_consulta_ia("qualquer coisa") is None
