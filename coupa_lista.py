"""
Leitura da lista de eventos do Coupa (fase 1 do robô), uma página por vez.

Separado do main_api.py para poder ser testado sem navegador e para registrar, por página,
por que cada linha foi aceita ou pulada. Antes, linhas puladas não deixavam rastro no log e
uma varredura que não encontrava o ponto de parada parecia "travada" sem explicação.

Regras de decisão por linha (mesma ordem e mesmo critério da versão anterior):
  poucas_colunas   menos de 7 colunas
  bandeira_amarela evento com bandeira amarela
  concluido        status contém "Concluído"
  data_invalida    data de início ilegível
  data_anterior    data de início anterior à pedida (encerra a varredura)
  outra_data       data de início posterior à pedida
  coletar          data de início igual à pedida
Erro "stale element" (o Coupa redesenhou a tabela no meio da leitura) agora é tentado de novo
em vez de descartar a linha.
"""

import time
from collections import Counter

from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By

XPATH_TBODY = '//*[@id="quote_request_table_tag"]'
TENTATIVAS_STALE = 3
PAUSA_STALE = 0.5


def ler_linha_evento(driver, i, hoje, parse_date_str, com_info=False):
    """Lê a linha i da tabela. Retorna um dict com 'estado' em: fim, pular, parar, coletar."""
    tbody = driver.find_element(By.XPATH, XPATH_TBODY)
    linhas = tbody.find_elements(By.TAG_NAME, "tr")
    if i >= len(linhas):
        return {"estado": "fim", "motivo": "fim_tabela"}
    linha = linhas[i]

    colunas = linha.find_elements(By.TAG_NAME, "td")
    if not colunas or len(colunas) < 7:
        return {"estado": "pular", "motivo": "poucas_colunas"}

    status_text = colunas[4].text.strip()
    data_inicio_str = colunas[2].text.strip()
    res = {"data_inicio_str": data_inicio_str, "status": status_text}
    res["info"] = (colunas[0].text.strip(), data_inicio_str, status_text)

    if linha.find_elements(By.CSS_SELECTOR, "img[src*='flag_yellow']"):
        res.update(estado="pular", motivo="bandeira_amarela")
        return res

    if "Concluído" in status_text:
        res.update(estado="pular", motivo="concluido")
        return res

    data_inicio = parse_date_str(data_inicio_str)
    if data_inicio is None:
        res.update(estado="pular", motivo="data_invalida")
        return res

    if data_inicio < hoje:
        res.update(estado="parar", motivo="data_anterior")
        return res

    if data_inicio != hoje:
        res.update(estado="pular", motivo="outra_data")
        return res

    res["numero"] = colunas[0].find_element(By.TAG_NAME, "a").text.strip()
    res["data_final"] = colunas[3].text.strip()
    res["estado"] = "coletar"
    return res


def varrer_pagina(driver, hoje, parse_date_str, num_linhas, ao_coletar, log,
                  tentativas=TENTATIVAS_STALE, pausa=PAUSA_STALE):
    """
    Percorre as linhas da página atual.
    ao_coletar(numero, data_final) é chamado para cada evento da data pedida e pode devolver
    um motivo (ex.: 'ja_no_banco') quando o evento não deve ser coletado.
    Retorna o resumo da página (motivos, datas e status vistos, primeira e última linha).
    """
    motivos = Counter()
    datas = Counter()
    status = Counter()
    coletados = 0
    encontrou_anterior = False
    anteriores = 0   # linhas com data de inicio anterior a pedida, qualquer que seja o status
    recentes = 0     # linhas com data de inicio igual ou posterior a pedida, qualquer status
    primeira = ultima = None

    for i in range(num_linhas):
        com_info = i == 0 or i == num_linhas - 1
        res = None
        for tentativa in range(tentativas):
            try:
                res = ler_linha_evento(driver, i, hoje, parse_date_str, com_info)
                break
            except StaleElementReferenceException:
                if tentativa == tentativas - 1:
                    motivos["erro_stale"] += 1
                    log(f"ERRO na linha {i}: tabela redesenhada, linha descartada apos {tentativas} tentativas")
                else:
                    time.sleep(pausa)
            except Exception as e:  # noqa: BLE001
                motivos["erro"] += 1
                msg = (str(e).strip().splitlines() or [type(e).__name__])[0]
                log(f"ERRO na linha {i}: {msg}")
                break
        if res is None:
            continue

        if res.get("info"):
            if primeira is None:
                primeira = res["info"]
            ultima = res["info"]
        if res.get("data_inicio_str") is not None:
            datas[res["data_inicio_str"] or "(vazia)"] += 1
            status[res["status"] or "(vazio)"] += 1
            d = parse_date_str(res["data_inicio_str"])
            if d is not None:
                if d < hoje:
                    anteriores += 1
                else:
                    recentes += 1

        estado = res["estado"]
        if estado == "fim":
            break
        if estado == "parar":
            motivos[res["motivo"]] += 1
            encontrou_anterior = True
            log(f"Data anterior encontrada ({res['data_inicio_str']}), encerrando coleta.")
            break
        if estado == "pular":
            motivos[res["motivo"]] += 1
            continue

        motivo = ao_coletar(res["numero"], res["data_final"])
        if motivo:
            motivos[motivo] += 1
        else:
            coletados += 1

    # Pagina so de eventos antigos (ex.: Concluidos de 2025): nada mais a coletar adiante.
    # O criterio por linha nao enxerga isso porque "Concluido" e descartado antes da data.
    pagina_antiga = (not encontrou_anterior and coletados == 0
                     and anteriores > 0 and recentes == 0)

    return {
        "coletados": coletados,
        "encontrou_anterior": encontrou_anterior,
        "pagina_antiga": pagina_antiga,
        "motivos": motivos,
        "datas": datas,
        "status": status,
        "primeira": primeira,
        "ultima": ultima,
    }


def resumo_texto(pagina, num_linhas, r):
    """Uma linha de log por página, para ver onde a varredura se perde."""
    def top(c, n=3):
        return ", ".join(f"{k}={v}" for k, v in c.most_common(n)) or "-"

    def pt(x):
        return "|".join(x) if x else "-"

    return (f"Pagina {pagina}: resumo | linhas={num_linhas} coletados={r['coletados']} "
            f"puladas=[{top(r['motivos'], 8)}] datas=[{top(r['datas'])}] "
            f"status=[{top(r['status'])}] primeira={pt(r['primeira'])} ultima={pt(r['ultima'])}")
