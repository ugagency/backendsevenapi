"""Testes da planilha unificada (varias contas Coupa) e do modo simulacao.

Rodar na raiz do backend:  python tests/test_planilha_unificada.py
Nao usa Coupa, Selenium, Supabase nem rede: o robo e o envio sao substituidos por stubs.
"""
import os
import shutil
import sys
import tempfile

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)

os.environ.update({
    "SUPABASE_URL": "https://x.supabase.co", "SUPABASE_KEY": "x" * 40,
    "VALE_USER": "u1", "VALE_PASS": "p1", "VALE_CONTA_NOME": "Seven Matriz",
    "VALE_USER_2": "u2", "VALE_PASS_2": "p2", "VALE_CONTA_NOME_2": "Seven Filial 2",
    "VALE_USER_3": "u3", "VALE_PASS_3": "p3", "VALE_CONTA_ID_3": "todas",
})
os.environ.pop("ROBO_SIMULAR", None)

import supabase as _s
_s.create_client = lambda *a, **k: object()
import main_api as m
from openpyxl import load_workbook

SAIDA = tempfile.mkdtemp(prefix="seven_teste_")
m.SAIDA_TESTE = os.path.join(SAIDA, "saida_teste")

assert list(m.CONTAS_COUPA) == ["filial1", "filial2"], list(m.CONTAS_COUPA)  # id 'todas' e reservado

NE = "Titulo nao encontrado"
F1 = [("100", "10", "MG", "05/11/26", "PARAFUSO M8", "10", "UN", "", "ACME"),
      ("101", "10", "UF nao encontrada", "05/11/26", "VALVULA", "2", "UN", "", "Marca nao encontrada"),
      ("200", NE, "MG", "06/11/26", "ITEM A", "1", "UN", "", "X"),
      ("200", NE, "MG", "06/11/26", "ITEM B", "1", "UN", "", "X"),
      ("300", "10", "SP", "07/11/26", "BOMBA", "4", "UN", "", "Y")]
F2 = [("100", "10", "MG", "05/11/26", "PARAFUSO M8", "10", "UN", "", "ACME"),   # repetido
      ("100", "20", "MG", "05/11/26", "PORCA M8", "10", "UN", "", "ACME"),       # item novo no mesmo evento
      ("101", "10", "BA", "05/11/26", "VALVULA", "2", "UN", "", "GEMU"),         # completa UF e marca
      ("200", NE, "MG", "06/11/26", "ITEM A", "1", "UN", "", "X"),               # casa com o 1o ITEM A
      ("200", NE, "MG", "06/11/26", "ITEM B", "1", "UN", "", "X"),
      ("300", "10", "RJ", "07/11/26", "BOMBA", "9", "UN", "", "Y"),              # diverge UF e QTDE
      ("400", "10", "MG", "08/11/26", "CABO", "7", "M", "", "Z")]

modo = {"f2": "ok", "upload": "ok"}
chamadas = []


def robo_falso(data, filename, conta_id, pular=False, coletor=None, simular=False, dados_conta=None):
    assert dados_conta and dados_conta["id"] == conta_id  # o robo recebe o retrato da conta
    chamadas.append({"conta": conta_id, "simular": simular, "coletor": coletor is not None})
    if conta_id == "filial2" and modo["f2"] != "ok":
        return modo["f2"]
    if coletor is None:
        return "ok-individual"
    coletor.extend({"filial1": F1, "filial2": F2}[conta_id])
    return "ok"


enviados = []


def envio_falso(caminho, nome):
    if modo["upload"] != "ok":
        raise RuntimeError("storage fora")
    shutil.copy(caminho, os.path.join(SAIDA, nome))
    enviados.append(nome)


m.executar_robo_selenium = robo_falso
m.enviar_planilha = envio_falso


def ler(caminho):
    ws = load_workbook(caminho)["Eventos"]
    return [list(r) for r in ws.iter_rows(values_only=True)]


def rodar(contas, simular=False):
    m.ROBO_LOCK.acquire()  # a fila libera no final
    enviados.clear()
    chamadas.clear()
    m.executar_fila("051126", "051126", contas, False, simular)
    return m.ESTADO_ROBO["ultima_execucao"]["resultados"]


# 1) caso normal: uma planilha, evento repetido uma vez so
res = rodar(["filial1", "filial2"])
assert res == {"filial1": "ok", "filial2": "ok"}, res
assert len(enviados) == 1 and enviados[0].startswith("eventos_todas_051126_"), enviados
rows = ler(os.path.join(SAIDA, enviados[0]))
assert rows[0][:9] == m.CABECALHO_PLANILHA and rows[0][-1] == "Filiais"
corpo = rows[1:]
assert len(corpo) == 7, len(corpo)  # 12 recebidas - 5 repetidas
por = {(r[0], r[1], r[4]): r for r in corpo}
assert len(por) == len(corpo)
assert por[("100", "10", "PARAFUSO M8")][9] == "Seven Matriz, Seven Filial 2"
assert por[("100", "20", "PORCA M8")][9] == "Seven Filial 2"
assert por[("400", "10", "CABO")][9] == "Seven Filial 2"
assert por[("101", "10", "VALVULA")][2] == "BA" and por[("101", "10", "VALVULA")][8] == "GEMU"  # completou o que faltava
assert por[("300", "10", "BOMBA")][2] == "SP" and por[("300", "10", "BOMBA")][5] == "4"          # manteve a 1a conta
assert [int(r[0]) for r in corpo] == sorted(int(r[0]) for r in corpo)                            # ordenado por evento
assert sum(1 for r in corpo if r[0] == "200") == 2                                               # itens distintos nao fundidos
print("1) ok - planilha unica, repetidos fundidos, itens distintos preservados")

# 2) filial 2 falha: planilha so com a filial 1 e o erro aparece
modo["f2"] = "erro: Login no Coupa falhou"
res = rodar(["filial1", "filial2"])
corpo = ler(os.path.join(SAIDA, enviados[0]))[1:]
assert len(corpo) == 5 and all(r[9] == "Seven Matriz" for r in corpo)
assert res["filial2"].startswith("erro")
print("2) ok - falha de login numa filial nao gera planilha vazia nem esconde o erro")

# 3) falha no envio: erro em 'todas' e arquivo preservado
modo["f2"] = "ok"
modo["upload"] = "falha"
res = rodar(["filial1", "filial2"])
assert res["todas"].startswith("erro"), res
preservados = [f for f in os.listdir(tempfile.gettempdir()) if f.startswith("eventos_todas_051126_")]
assert preservados, "arquivo deveria ficar preservado para recuperacao"
for f in preservados:
    os.remove(os.path.join(tempfile.gettempdir(), f))
print("3) ok - falha no envio aparece e o arquivo e preservado")

# 4) uma conta so: comportamento de antes (sem coletor, sem planilha unificada)
modo["upload"] = "ok"
res = rodar(["filial2"])
assert enviados == [] and res == {"filial2": "ok-individual"}, (enviados, res)
assert chamadas == [{"conta": "filial2", "simular": False, "coletor": False}]
print("4) ok - conta unica segue o fluxo antigo")

# 5) simulacao: nada vai ao Storage; a planilha unificada fica na pasta de teste
res = rodar(["filial1", "filial2"], simular=True)
assert enviados == [], enviados
assert all(c["simular"] for c in chamadas)
arqs = os.listdir(m.SAIDA_TESTE)
assert len(arqs) == 1 and arqs[0].startswith("eventos_todas_"), arqs
assert len(ler(os.path.join(m.SAIDA_TESTE, arqs[0]))) == 8
print("5) ok - simulacao nao envia nada e salva a planilha em", m.SAIDA_TESTE)

shutil.rmtree(SAIDA, ignore_errors=True)
print("TODOS OS TESTES PASSARAM")
