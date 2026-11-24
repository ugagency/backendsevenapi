"""Testes de integracao: o os_api de verdade contra um Postgres local, conectado como os_app.

Preparar o banco: tests_os/local/preparar.sh
Rodar: OS_TEST_DATABASE_URL="postgresql://os_app@/seven_teste?host=/tmp&port=5544" python -m pytest tests_os -q
Sem a variavel, estes testes sao pulados. So o login (Supabase Auth) e simulado.
"""
import io
import os
import zipfile

import psycopg
import pytest
from fastapi.testclient import TestClient
from openpyxl import load_workbook

from os_api import auth, db
from os_api.auth import Usuario
from os_api.config import Config
from os_api.main import criar_app

URL = os.getenv("OS_TEST_DATABASE_URL")
pytestmark = pytest.mark.skipif(not URL, reason="OS_TEST_DATABASE_URL nao definida")


@pytest.fixture(scope="module")
def app():
    cfg = Config(database_url=URL, schema="os_homolog", supabase_url="https://x", supabase_key="k", cors_origens=())
    a = criar_app(cfg)
    with TestClient(a) as c:
        yield c
    db.encerrar()


@pytest.fixture(scope="module")
def ids():
    with psycopg.connect(URL) as conn:
        return {e.split("@")[0]: str(i) for i, e in conn.execute("select user_id, email from os_homolog.usuarios").fetchall()}


@pytest.fixture()
def como(app, ids, monkeypatch):
    def _entrar(quem):
        auth.limpar_cache()
        monkeypatch.setattr(auth, "_validar_no_supabase", lambda c, t: Usuario(ids[quem], f"{quem}@teste"))
        app.headers.update({"Authorization": f"Bearer {quem}"})
        return app
    return _entrar


def item(c, evento):
    r = c.get("/os/fila", params={"busca": evento, "situacao": "todos"}).json()
    return next(i for i in r["itens"] if i["numero_evento"] == evento)


ESTADO = {}


def test_01_fila_e_visibilidade_por_conta(como):
    c = como("admin")
    r = c.get("/os/fila").json()
    eventos = {i["numero_evento"] for i in r["itens"]}
    assert {"T1001", "T1002", "T1003", "T1004"} <= eventos and "T0900" not in eventos  # expirado fica fora dos abertos
    assert sum(r["abas"].values()) == r["total"]
    t1001 = item(c, "T1001")
    assert t1001["contas"] == [1, 2] and t1001["quantidade"] == 10 and t1001["ocorrencias"] == 3
    assert "valor_estimado" in t1001
    assert c.get("/os/fila", params={"busca": "6205"}).json()["total"] >= 2
    assert c.get("/os/fila", params={"conta": 2}).json()["total"] < r["total"]
    ESTADO.update(t1001=t1001["id"], t1002=item(c, "T1002")["id"], t1003=item(c, "T1003")["id"], t1004=item(c, "T1004")["id"],
                  t0900=item(c, "T0900")["id"], p1=t1001["produto"]["id"], p2=item(c, "T1002")["produto"]["id"],
                  p3=item(c, "T1003")["produto"]["id"])
    c = como("cotacao1")  # so filial1
    eventos = {i["numero_evento"] for i in c.get("/os/fila").json()["itens"]}
    assert "T1003" not in eventos and "T1001" in eventos
    assert c.get(f"/os/itens/{ESTADO['t1003']}").status_code == 404
    assert c.get("/os/fila", params={"conta": 2}).status_code == 403
    assert "valor_estimado" not in item(c, "T1001")
    assert c.get("/os/fila", params={"ordem": "valor"}).status_code == 403


def test_02_ficha(como):
    f = como("cotacao").get(f"/os/itens/{ESTADO['t1001']}").json()
    assert f["produto"]["fabricante"] == "SKF" and f["produto"]["nivel"] == "certeza"
    assert len(f["linhas"]) == 2 and len(f["contas"]) == 2 and f["ocorrencias"] == 3
    assert {a["numero_evento"] for a in f["anteriores"]} == {"T0900", "T1004"}
    assert [p["id"] for p in f["mesma_referencia"]] == [ESTADO["p2"]]


def test_03_conferencia(como):
    c = como("cotacao")
    conf = c.get("/os/conferencia").json()
    assert ESTADO["p3"] in [p["id"] for p in conf["produtos"]]
    assert como("suporte").post(f"/os/conferencia/{ESTADO['p3']}", json={"acao": "novo_produto"}).status_code == 403
    c = como("cotacao")
    assert c.post(f"/os/conferencia/{ESTADO['p3']}", json={"acao": "xx"}).status_code == 400
    assert c.post(f"/os/conferencia/{ESTADO['p3']}", json={"acao": "novo_produto", "nome": "KIT REPARO BOMBA"}).status_code == 200
    assert ESTADO["p3"] not in [p["id"] for p in c.get("/os/conferencia").json()["produtos"]]
    assert item(como("admin"), "T1003")["aba"] != "conferencia"


def test_04_fornecedores_vinculo_e_abas(como):
    c = como("cotacao")
    assert c.post("/os/fornecedores", json={"nome": ""}).status_code == 400
    f1 = c.post("/os/fornecedores", json={"nome": "Rolamentos Betim", "uf": "mg", "email": "vendas@rolbetim.test", "contato": "Joao"}).json()
    assert f1["uf"] == "MG"
    assert c.post("/os/fornecedores", json={"nome": "ROLAMENTOS BETIM"}).status_code == 409
    f2 = c.post("/os/fornecedores", json={"nome": "Rol. Betim Ltda", "marcas": ["skf", "NSK"]}).json()
    assert f2["marcas"] == ["NSK", "SKF"]
    ESTADO.update(f1=f1["id"], f2=f2["id"])
    pf = c.get(f"/os/produtos/{ESTADO['p1']}/fornecedores").json()
    assert pf["vinculados"] == [] and [s["id"] for s in pf["sugestoes"]] == [f2["id"]] and "SKF" in pf["sugestoes"][0]["fonte"]
    assert c.post(f"/os/produtos/{ESTADO['p1']}/sugestoes/{f2['id']}/descartar").status_code == 200
    assert c.get(f"/os/produtos/{ESTADO['p1']}/fornecedores").json()["sugestoes"] == []
    assert como("suporte").post(f"/os/produtos/{ESTADO['p1']}/fornecedores", json={"fornecedor_id": f1["id"]}).status_code == 403
    c = como("cotacao")
    assert c.post(f"/os/produtos/{ESTADO['p1']}/fornecedores", json={"fornecedor_id": f1["id"]}).status_code == 201
    assert item(c, "T1001")["aba"] == "pronto" and item(c, "T1004")["aba"] == "pronto"
    assert c.get("/os/fila", params={"aba": "pronto"}).json()["total"] >= 2
    # juntar duas escritas do mesmo fornecedor
    j = c.post("/os/fornecedores/juntar", json={"manter_id": f1["id"], "juntar_id": f2["id"]}).json()
    assert "Rol. Betim Ltda" in j["outros_nomes"] and j["marcas"] == ["NSK", "SKF"]
    assert c.get(f"/os/fornecedores/{f2['id']}").status_code == 404
    assert c.post("/os/fornecedores", json={"nome": "rol. betim ltda"}).status_code == 409  # variacao conhecida nao vira cadastro novo
    lst = c.get("/os/fornecedores", params={"busca": "Rol. Betim"}).json()["fornecedores"]
    assert [x["id"] for x in lst] == [f1["id"]] and lst[0]["produtos"] == 1


def test_05_solicitacao_sem_preco(como):
    c = como("cotacao")
    base = {"item_id": ESTADO["t1001"], "fornecedor_id": ESTADO["f1"]}
    rasc = c.post("/os/solicitacoes", json=base).json()
    assert rasc["registrada"] is False and "6205-2RS" in rasc["texto"] and "R$" not in rasc["texto"] and rasc["para"] == "vendas@rolbetim.test"
    r = c.post("/os/solicitacoes", json={**base, "registrar": True, "texto": rasc["texto"] + "\nTemos alvo de R$ 90,00."})
    assert r.status_code == 400 and "bloqueado" in r.json()["detail"].lower()
    assert c.post("/os/solicitacoes", json={**base, "registrar": True, "texto": rasc["texto"]}).json()["registrada"] is True
    assert len(c.get(f"/os/itens/{ESTADO['t1001']}/solicitacoes").json()) == 1


def test_06_cotacao_alertas_versoes_e_perfis(como):
    c = como("cotacao")
    i4, i1 = ESTADO["t1004"], ESTADO["t1001"]
    assert c.post(f"/os/itens/{i4}/enviar").status_code == 400  # sem cotacao
    r = c.post(f"/os/itens/{i4}/cotacoes", json={"fornecedor_id": ESTADO["f1"], "preco_compra": 100, "ipi": 5, "frete": 2, "preco_venda": 150})
    assert r.status_code == 201
    cot = r.json()["cotacoes"][0]
    assert cot["situacao"] == "pronta" and cot["alertas"] == [] and cot["conta"]
    assert cot["valores"] == {"preco_compra": 100.0, "ipi": 5.0, "frete": 2.0, "preco_venda": 150.0}  # Cotacao: sem margem
    assert item(c, "T1004")["andamento"] == "em_cotacao"
    assert c.post(f"/os/itens/{i4}/enviar").json()["andamento"] == "enviada"
    assert c.post(f"/os/itens/{i4}/resultado", json={"resultado": "talvez"}).status_code == 400
    assert c.post(f"/os/itens/{i4}/resultado", json={"resultado": "ganha", "pedido_venda": "PV-1"}).json()["andamento"] == "ganha"
    # item em dois CNPJs exige dizer qual respondeu
    r = c.post(f"/os/itens/{i1}/cotacoes", json={"fornecedor_id": ESTADO["f1"], "conta_id": 1, "preco_compra": 950, "preco_venda": 1200})
    cot = r.json()["cotacoes"][0]
    assert cot["situacao"] == "rascunho" and [a["tipo"] for a in cot["alertas"]] == ["parece_total"] and cot["alertas"][0]["valor_referencia"] == 100.0
    assert c.post(f"/os/itens/{i1}/enviar").status_code == 400  # alerta aberto bloqueia
    aid = cot["alertas"][0]["id"]
    assert c.post(f"/os/alertas/{aid}/justificar", json={"justificativa": "ok"}).status_code == 400
    j = c.post(f"/os/alertas/{aid}/justificar", json={"justificativa": "Lote minimo de 10, valor do lote confirmado"}).json()
    assert j["cotacoes"][0]["situacao"] == "pronta"
    v2 = c.post(f"/os/itens/{i1}/cotacoes", json={"fornecedor_id": ESTADO["f1"], "conta_id": 1, "preco_compra": 200, "preco_venda": 300}).json()["cotacoes"]
    assert [x["versao"] for x in v2] == [2, 1] and v2[1]["situacao"] == "substituida" and [a["tipo"] for a in v2[0]["alertas"]] == ["divergente"]
    v3 = c.post(f"/os/itens/{i1}/cotacoes", json={"conta_id": 1, "preco_venda": 300}).json()["cotacoes"]
    assert [a["tipo"] for a in v3[0]["alertas"]] == ["sem_preco"]
    assert c.post(f"/os/alertas/{v3[0]['alertas'][0]['id']}/justificar", json={"justificativa": "sem preco mesmo"}).status_code == 400
    v4 = c.post(f"/os/itens/{i1}/cotacoes", json={"fornecedor_id": ESTADO["f1"], "conta_id": 1, "preco_compra": 105, "preco_venda": 150}).json()["cotacoes"]
    assert v4[0]["situacao"] == "pronta" and v4[0]["versao"] == 4
    assert c.post(f"/os/itens/{i1}/cotacoes", json={"conta_id": 99, "preco_compra": 1}).status_code == 400
    assert c.post(f"/os/itens/{i1}/cotacoes", json={"preco_compra": -1}).status_code == 400
    # perfis
    adm = como("admin").get(f"/os/itens/{i1}/cotacoes").json()["cotacoes"][0]["valores"]
    assert adm["margem_pct"] == 30.0 and adm["preco_compra"] == 105.0
    sup = como("suporte").get(f"/os/itens/{i1}/cotacoes").json()["cotacoes"][0]
    assert sup["valores"] == {} and "valor_referencia" not in str(sup)
    assert como("suporte").post(f"/os/itens/{i1}/cotacoes", json={"preco_compra": 1}).status_code == 403
    ficha = como("cotacao").get(f"/os/itens/{i1}").json()
    ant = next(a for a in ficha["anteriores"] if a["numero_evento"] == "T1004")
    assert ant["resultado"] == "ganha" and ant["valores"]["preco_venda"] == 150.0 and "margem_pct" not in ant["valores"]
    assert "preco_compra" not in ant
    assert item(como("admin"), "T1002")["valor_estimado"] is None and item(como("admin"), "T0900")["valor_estimado"] == 6 * 150


def test_07_declinio_guarda_motivo_e_recorrencia(como):
    c = como("cotacao")
    i2 = ESTADO["t1002"]
    assert c.patch(f"/os/itens/{i2}/andamento", json={"para": "declinada"}).status_code == 400
    assert c.patch(f"/os/itens/{i2}/andamento", json={"para": "ganha"}).status_code == 400
    assert c.patch(f"/os/itens/{i2}/andamento", json={"para": "declinada", "motivo": "Sem fornecedor homologado"}).status_code == 200
    assert "T1002" not in {i["numero_evento"] for i in c.get("/os/fila").json()["itens"]}
    f = c.get(f"/os/itens/{i2}").json()
    assert f["declinios"][0]["motivo"] == "Sem fornecedor homologado" and f["andamentos"][0]["para"] == "declinada"
    assert item(c, "T1002")["declinios"] == 1
    assert como("suporte").patch(f"/os/itens/{i2}/andamento", json={"para": "a_cotar"}).status_code == 403
    assert como("cotacao").patch(f"/os/itens/{i2}/andamento", json={"para": "a_cotar"}).status_code == 200


def test_08_juntar_codigos_e_desfazer(como):
    c = como("cotacao")
    r = c.post(f"/os/produtos/{ESTADO['p1']}/juntar", json={"origem_id": ESTADO["p2"]})
    assert r.status_code == 200
    f = c.get(f"/os/itens/{ESTADO['t1001']}").json()
    assert {x["codigo"] for x in f["codigos"]} == {"13900001", "13900002"} and f["ocorrencias"] == 4 and len(f["juncoes"]) == 1
    assert item(c, "T1002")["produto"]["id"] == ESTADO["p1"]
    assert c.post(f"/os/produtos/{ESTADO['p1']}/juntar", json={"origem_id": ESTADO["p2"]}).status_code == 400
    assert c.post(f"/os/juncoes/{r.json()['juncao_id']}/desfazer").status_code == 200
    f = c.get(f"/os/itens/{ESTADO['t1001']}").json()
    assert {x["codigo"] for x in f["codigos"]} == {"13900001"} and f["ocorrencias"] == 3 and f["juncoes"] == []
    assert item(c, "T1002")["produto"]["id"] == ESTADO["p2"]
    assert c.get("/os/produtos", params={"busca": "139000"}).json()[0]["codigos"]


def test_09_painel_por_perfil(como):
    chaves = lambda quem: {i["chave"] for i in como(quem).get("/os/painel").json()["indicadores"]}
    restritos = {"valor_aberto", "margem", "divergencias", "vitoria"}
    assert not (chaves("cotacao") & restritos) and not (chaves("suporte") & restritos)
    adm = {i["chave"]: i for i in como("admin").get("/os/painel").json()["indicadores"]}
    assert restritos <= set(adm)
    assert adm["vitoria"]["disponivel"] is False and "mínimo" in adm["vitoria"]["motivo"]
    assert adm["margem"]["valor"] is not None and adm["capturados"]["valor"] >= 5
    cart = como("cotacao").get("/os/fila/cartoes").json()
    assert "valor_estimado" not in cart and cart["em_aberto"] >= 3 and cart["aguardando_voce"] >= 1
    assert "valor_estimado" in como("supervisor").get("/os/fila/cartoes").json()


def test_10_assistente_respeita_perfil_e_nao_inventa(como):
    r = como("cotacao").post("/os/assistente", json={"pergunta": "Ja cotamos 13900001 e por quanto?"}).json()
    assert r["consulta"] == "historico_produto" and "Compra (un.)" in r["tabela"]["colunas"] and "Margem %" not in r["tabela"]["colunas"]
    assert r["link"].startswith("/os/item/")
    r2 = como("admin").post("/os/assistente", json={"pergunta": "Ja cotamos 13900001 e por quanto?"}).json()
    assert "Margem %" in r2["tabela"]["colunas"]
    r3 = como("suporte").post("/os/assistente", json={"pergunta": "historico de preco do 13900001"}).json()
    assert not ({"Compra (un.)", "Venda (un.)", "Margem %"} & set(r3["tabela"]["colunas"]))
    r4 = como("cotacao").post("/os/assistente", json={"pergunta": "o que encerra em 5 dias sem fornecedor?", "conversa_id": r["conversa_id"]}).json()
    assert r4["consulta"] == "itens_por_prazo" and "sem_fornecedor=1" in r4["link"]
    r5 = como("cotacao").post("/os/assistente", json={"pergunta": "ja cotamos 'peca que nao existe xyz'?"}).json()
    assert r5["tabela"] is None and "Não encontrei" in r5["texto"]
    assert como("cotacao").post("/os/assistente", json={"pergunta": "painel", "conversa_id": 999999}).status_code == 404
    c = como("cotacao")
    assert len(c.get(f"/os/assistente/conversas/{r['conversa_id']}").json()) == 4
    assert como("admin").get(f"/os/assistente/conversas/{r['conversa_id']}").status_code == 404
    with psycopg.connect(URL) as conn:  # nenhuma mensagem guardada carrega valor
        textos = " ".join(t for (t,) in conn.execute("select texto from os_homolog.mensagens").fetchall())
    assert "105" not in textos and "150" not in textos


def test_11_captura_planilha_e_sincronizacao(como):
    c = como("cotacao")
    d = c.get("/os/captura/dias").json()
    assert d["ultima_captura"] and d["dias"]
    dia = d["dias"][0]["dia"]
    r = c.get("/os/captura/planilha", params={"data": dia})
    assert r.status_code == 200
    ws = load_workbook(io.BytesIO(r.content)).active
    cab = [x.value for x in ws[1]]
    assert cab == ["Numero do evento", "Titulo", "UF(VALE)", "DATA", "DESCRIÇÃO", "QTDE", "UNID. MED", "pagina de descrição", "Marca", "Filiais"]
    linhas = [[x.value for x in row] for row in ws.iter_rows(min_row=2)]
    t1001 = [l for l in linhas if l[0] == "T1001"]
    assert len(t1001) == 1 and t1001[0][9] == "filial1, filial2" and t1001[0][1] == "13900001"
    uma = load_workbook(io.BytesIO(c.get("/os/captura/planilha", params={"data": dia, "conta": 1}).content)).active
    assert len([x.value for x in uma[1]]) == 9
    assert c.get("/os/captura/planilha", params={"data": "2001-01-01"}).status_code == 404
    assert c.post("/os/captura/sincronizar").json()["linhas_novas"] == 0
    assert c.post("/os/captura/rodar", json={}).status_code == 503


def test_11b_captura_mostra_o_dia_escolhido_e_nao_o_dia_da_gravacao(como):
    """O robo grava em outro dia (capturado_em), mas a tela mostra o dia pedido para a captura (data_inicio)."""
    adm = URL.replace("os_app@", "postgres@")
    with psycopg.connect(adm) as conn:
        alvo = conn.execute("select id from os_homolog.linhas_evento where data_inicio is null and conta_id = 1 order by id limit 1").fetchone()[0]
        conn.execute("update os_homolog.linhas_evento set data_inicio = date '2026-11-20' where id = %s", (alvo,))
        conn.commit()
    try:
        c = como("cotacao")
        dias = [d["dia"] for d in c.get("/os/captura/dias").json()["dias"]]
        assert "2026-11-20" in dias
        assert c.get("/os/captura/planilha", params={"data": "2026-11-20", "conta": 1}).status_code == 200
    finally:
        with psycopg.connect(adm) as conn:
            conn.execute("update os_homolog.linhas_evento set data_inicio = null where id = %s", (alvo,))
            conn.commit()


def test_12_auditoria_e_exportacoes(como):
    assert como("cotacao").get("/os/auditoria").status_code == 403
    sup = como("suporte").get("/os/auditoria", params={"entidade": "cotacao"}).json()["registros"]
    assert sup and all(x["depois"] is None and x["oculto"] for x in sup)
    adm = como("admin").get("/os/auditoria", params={"entidade": "cotacao"}).json()
    assert adm["registros"][0]["depois"]["preco_compra"] == 105 and adm["registros"][0]["autor"] == "Cotacao"
    assert {"cotacao", "fornecedor", "item", "alerta_preco", "produto"} <= set(adm["entidades"])
    assert como("supervisor").post("/os/admin/exportar-base").status_code == 403
    z = zipfile.ZipFile(io.BytesIO(como("admin").post("/os/admin/exportar-base").content))
    assert {"cotacao_valores.csv", "linhas_evento.csv", "auditoria.csv", "LEIAME.txt"} <= set(z.namelist())
    assert "105" in z.read("cotacao_valores.csv").decode()
    cab = lambda quem: [x.value for x in load_workbook(io.BytesIO(como(quem).get("/os/fila/exportar").content)).active[1]]
    assert "Valor estimado" not in cab("cotacao") and "Valor estimado" in cab("admin")
    with psycopg.connect(URL) as conn:
        assert conn.execute("select count(*) from os_homolog.auditoria where acao = 'exportar'").fetchone()[0] >= 4
        with pytest.raises(psycopg.errors.InsufficientPrivilege):
            conn.execute("update os_homolog.auditoria set justificativa = 'x'")


def test_13_admin_real(como):
    c = como("suporte")
    us = c.get("/os/admin/usuarios").json()
    assert len(us) == 5
    assert c.put("/os/admin/parametros/robo.agenda", json={"valor": {"horarios": ["13:00", "07:00"]}}).json()["valor"] == {"horarios": ["07:00", "13:00"]}
    assert c.put("/os/admin/parametros/alertas.divergencia_pct", json={"valor": 15}).status_code == 200
    conta = c.post("/os/admin/contas", json={"nome": "Filial teste", "cnpj": "11.222.333/0001-81"}).json()
    assert conta["cnpj"] == "11222333000181"
    alvo = next(x for x in us if x["email"] == "cotacao1@teste")
    assert c.patch(f"/os/admin/usuarios/{alvo['user_id']}", json={"contas": [1, conta["id"]]}).json()["contas"] == [1, conta["id"]]
    adm = next(x for x in us if x["email"] == "admin@teste")
    assert c.patch(f"/os/admin/usuarios/{adm['user_id']}", json={"ativo": False}).status_code == 409  # ultimo administrador
    assert c.post("/os/admin/usuarios", json={"email": "ninguem@teste", "perfil": "cotacao"}).status_code == 503  # convite sem URL configurada
    eu = como("cotacao1").get("/os/eu").json()
    assert [x["id"] for x in eu["contas"]] == [1, conta["id"]] and eu["perfil"] == "cotacao"


def test_14_convite(como, app, monkeypatch):
    import dataclasses
    import httpx as _httpx
    from os_api import admin
    with psycopg.connect(URL) as conn:
        novo_id = str(conn.execute("select os_homolog.auth_id_por_email('convidado@novo')").fetchone()[0])
    chamadas = []

    def falso_post(url, headers=None, json=None, timeout=None):
        chamadas.append((url, headers, json))
        return _httpx.Response(201, json={"user_id": novo_id, "ja_existia": False, "modo": json["modo"], "link": "https://convite/abc"})

    monkeypatch.setattr(admin.httpx, "post", falso_post)
    cfg_antes = app.app.state.cfg
    app.app.state.cfg = dataclasses.replace(cfg_antes, convite_url="https://fn/os-convidar", convite_redirect="https://front/definir-senha")
    try:
        c = como("cotacao1")
        assert c.post("/os/admin/usuarios", json={"email": "Pessoa@Fora", "perfil": "cotacao"}).status_code == 403
        assert not chamadas
        c = como("admin")
        assert c.post("/os/admin/usuarios", json={"email": "x@fora", "perfil": "cotacao", "convite": "fax"}).status_code == 400
        r = c.post("/os/admin/usuarios", json={"email": "Pessoa@Fora", "nome": "Pessoa", "perfil": "cotacao", "contas": [1]})
        assert r.status_code == 201, r.text
        d = r.json()
        assert d["link_convite"] == "https://convite/abc" and d["convite"] == "link" and d["user_id"] == novo_id
        url, headers, corpo = chamadas[0]
        assert url == "https://fn/os-convidar" and headers["Authorization"] == "Bearer admin"
        assert corpo == {"email": "pessoa@fora", "schema": "os_homolog", "modo": "link", "redirect_to": "https://front/definir-senha", "nome": "Pessoa"}
        with psycopg.connect(URL) as conn:
            aud = conn.execute("select depois::text from os_homolog.auditoria where entidade = 'usuario' and entidade_id = %s", (novo_id,)).fetchone()[0]
        assert "convite" in aud and "https://convite" not in aud  # o link nunca vai para a auditoria
        # link novo para quem ja tem login (convite expirado ou senha esquecida)
        r = c.post(f"/os/admin/usuarios/{novo_id}/link-acesso", json={})
        assert r.status_code == 200 and r.json()["link_convite"] == "https://convite/abc"
        assert chamadas[-1][2]["reenviar"] is True and chamadas[-1][2]["email"] == "pessoa@fora"
        assert como("supervisor").post(f"/os/admin/usuarios/{novo_id}/link-acesso", json={}).status_code == 403
        c = como("admin")
        # servico de convite fora do ar: erro claro, nada gravado
        monkeypatch.setattr(admin.httpx, "post", lambda *a, **k: _httpx.Response(404, json={}))
        r = c.post("/os/admin/usuarios", json={"email": "outra@fora", "perfil": "cotacao"})
        assert r.status_code == 503 and "nao publicado" in r.json()["detail"]
    finally:
        app.app.state.cfg = cfg_antes
