"""Agente de IA (n8n): o que sai para o n8n, leitura da resposta, validacao da ferramenta e volta para as regras."""
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


class _U:
    nome = "Karen Silva"
    permissoes = frozenset({"fila.ver", "preco_compra.ver"})


def test_sem_url_usa_regras(monkeypatch):
    monkeypatch.delenv("OS_N8N_ASSISTENTE_URL", raising=False)
    assert a.chamar_agente("quais indicadores?", [], "tok", "https://api", _U()) is None


def test_agente_recebe_pergunta_historico_e_sessao(monkeypatch):
    monkeypatch.setenv("OS_N8N_ASSISTENTE_URL", "https://n8n.exemplo/webhook/x")
    monkeypatch.delenv("OS_N8N_TOKEN", raising=False)
    visto = {}

    def falso(url, json=None, headers=None, timeout=None):
        visto.update(json=json, headers=headers)
        return _Resp(200, [{"output": "  Cotamos 3 vezes.  "}])
    monkeypatch.setattr(a.httpx, "post", falso)
    hist = [{"papel": "usuario", "texto": "oi"}]
    assert a.chamar_agente("ja cotamos?", hist, "tok123", "https://api", _U()) == "Cotamos 3 vezes."
    assert set(visto["json"]) == {"pergunta", "historico", "hoje", "dia_da_semana", "token", "api_url", "usuario"}
    assert visto["json"]["token"] == "tok123" and visto["json"]["historico"] == hist
    assert visto["json"]["usuario"] == {"nome": "Karen", "ve_valores": True}
    assert visto["headers"] == {}


def test_agente_com_falha_volta_para_regras(monkeypatch):
    monkeypatch.setenv("OS_N8N_ASSISTENTE_URL", "https://n8n.exemplo/webhook/x")
    for falha in (lambda *a, **k: (_ for _ in ()).throw(httpx.ConnectTimeout("t")), lambda *a, **k: _Resp(500, {}),
                  lambda *a, **k: _Resp(200, None), lambda *a, **k: _Resp(200, {"output": "   "}), lambda *a, **k: _Resp(200, {"x": 1})):
        monkeypatch.setattr(a.httpx, "post", falha)
        assert a.chamar_agente("indicadores", [], "t", "https://api", _U()) is None


def test_parametros_da_ferramenta():
    F = a.Ferramenta
    assert a._par_ferramenta(F(consulta="itens_por_prazo", ate="2026-11-28", sem_fornecedor="true")) == ("itens_por_prazo", {"ate": "2026-11-28", "sem_fornecedor": True})
    assert a._par_ferramenta(F(consulta="historico_produto", termo="")) is None
    assert a._par_ferramenta(F(consulta="apagar_tudo", termo="x")) is None
    assert a._par_ferramenta(F(consulta="buscar_produto", termo="rolamento", ate="lixo")) == ("buscar_produto", {"termo": "rolamento"})
