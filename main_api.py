import os
import re
from collections import Counter
import sys
import json
import time
import threading
import shutil
import tempfile
from datetime import datetime
from fastapi import FastAPI, BackgroundTasks, HTTPException, Header
from pydantic import BaseModel
from fastapi.middleware.cors import CORSMiddleware
from selenium import webdriver
from selenium.webdriver.common.by import By
from coupa_lista import varrer_pagina, resumo_texto
from coupa_descricao import coletar_descricoes
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook, load_workbook
from supabase import create_client, Client
from dotenv import load_dotenv
import credenciais

# Proteção contra caracteres não-ASCII (emojis, acentos, texto vindo do Gemini) quebrando
# o stdout/stderr no Windows, cujo encoding padrão do console (cp1252) não os suporta.
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# Carregar variáveis de ambiente (.env)
load_dotenv()

app = FastAPI()

# Configuração de CORS para permitir chamadas do Frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # No futuro, mude para o domínio real do seu Frontend no Vercel
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Configurações do Supabase (Estas devem estar no seu arquivo .env)
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
SUPABASE_BUCKET = os.getenv("SUPABASE_BUCKET", "relatorios")  # Nome do bucket no Supabase Storage

supabase = None  # sem credenciais so o modo simulacao funciona (ele nao usa o banco)
if not SUPABASE_URL or not SUPABASE_KEY:
    print("ERRO: SUPABASE_URL e SUPABASE_KEY não configurados no .env")
else:
    supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

# Cliente com a chave SECRETA (service_role / sb_secret_...): so ele acessa a tabela de
# credenciais (RLS sem politicas) e valida o login do painel. Nunca vai para o front.
SUPABASE_SECRET_KEY = os.getenv("SUPABASE_SECRET_KEY")
supabase_admin = None
if SUPABASE_URL and SUPABASE_SECRET_KEY:
    supabase_admin = create_client(SUPABASE_URL, SUPABASE_SECRET_KEY)
else:
    print("AVISO: SUPABASE_SECRET_KEY nao configurada: credenciais so do .env e gerenciador desligado.")

# Contas Coupa (uma por CNPJ/filial). A lista de contas vem do .env:
# conta 1 usa VALE_USER / VALE_PASS / VALE_CONTA_ID / VALE_CONTA_NOME; as seguintes, sufixo _2, _3...
# Usuario e senha podem vir do banco (tabela credenciais_coupa), que vence o .env; ver credenciais.py.
CONTA_ID_RE = re.compile(r"^[a-z0-9_-]{1,30}$")

def carregar_contas_env() -> list:
    contas, ids = [], set()
    for n in range(1, 10):
        suf = "" if n == 1 else f"_{n}"
        user = (os.getenv(f"VALE_USER{suf}") or "").strip()
        senha = os.getenv(f"VALE_PASS{suf}") or ""
        id_env = (os.getenv(f"VALE_CONTA_ID{suf}") or "").strip().lower()
        nome_env = (os.getenv(f"VALE_CONTA_NOME{suf}") or "").strip()
        if not (user or senha or id_env or nome_env):
            continue
        conta_id = id_env or f"filial{n}"
        if conta_id == "todas":
            print(f"ERRO: VALE_CONTA_ID{suf}='todas' e reservado (significa todas as contas). Conta ignorada.")
            continue
        if not CONTA_ID_RE.match(conta_id):
            print(f"ERRO: VALE_CONTA_ID{suf}='{conta_id}' invalido (use a-z, 0-9, _ ou -). Conta ignorada.")
            continue
        if conta_id in ids:
            print(f"ERRO: VALE_CONTA_ID{suf}='{conta_id}' repetido. Conta ignorada.")
            continue
        ids.add(conta_id)
        contas.append({"id": conta_id, "nome": nome_env or f"Filial {n}", "user": user, "senha": senha})
    return contas

CONTAS_TODAS = []   # todas as contas conhecidas, prontas ou nao (para o gerenciador)
CONTAS_COUPA = {}   # so as prontas (usuario e senha), na ordem: as que o robo pode rodar
CONTA_PADRAO = None

def atualizar_contas() -> None:
    """Relida a cada chamada (tabela pequena): trocar a senha no painel vale ja na proxima execucao,
    sem reiniciar o servidor. Se o banco falhar, segue com o .env e registra no log."""
    global CONTAS_TODAS, CONTA_PADRAO
    linhas = {}
    if supabase_admin is not None and credenciais.chave_configurada():
        try:
            linhas = credenciais.ler_linhas(supabase_admin)
        except Exception as e:
            log(f"AVISO: nao foi possivel ler credenciais do banco ({type(e).__name__}); usando so o .env.")
    CONTAS_TODAS = credenciais.mesclar(carregar_contas_env(), linhas)
    CONTAS_COUPA.clear()
    CONTAS_COUPA.update({c["id"]: c for c in CONTAS_TODAS if credenciais.pronta(c)})
    # A conta padrao (nome de arquivo sem id) e sempre a primeira conhecida, pronta ou nao.
    CONTA_PADRAO = CONTAS_TODAS[0]["id"] if CONTAS_TODAS else None

# Só um Chrome por vez: o VPS é compartilhado e duas sessões simultâneas
# no Coupa já foram observadas em teste. Uvicorn roda com 1 worker, então
# um lock de processo é suficiente.
ROBO_LOCK = threading.Lock()
ESTADO_ROBO = {"executando": False, "conta": None, "data": None, "inicio": None, "ultima_execucao": None}

def montar_evento_db(conta_id: str, r, data_inicio) -> dict:
    """Linha de eventos_coletados a partir de uma linha final da planilha (9 colunas).
    data_inicio: dia em que o evento abriu no Coupa (coluna 'Data de inicio' da lista; e o proprio
    dia pesquisado). Vai para a coluna data_inicio (tipo date). Atencao: data_evento guarda o PRAZO
    (coluna 'DATA' da planilha, a data final do evento), nome historico mantido para nao quebrar o robo atual."""
    return {
        "conta": conta_id,
        "numero_evento": str(r[0]).strip(),
        "titulo": str(r[1]) if len(r) > 1 and r[1] is not None else "",
        "uf": str(r[2]) if r[2] is not None else "",
        "data_evento": str(r[3]) if r[3] is not None else "",
        "data_inicio": data_inicio.isoformat() if data_inicio else None,
        "descricao": str(r[4]) if r[4] is not None else "",
        "quantidade": str(r[5]) if r[5] is not None else "",
        "unidade": str(r[6]) if r[6] is not None else "",
        "marca": str(r[8]) if len(r) > 8 and r[8] is not None else "",
    }

def parse_date_str(s: str):
    """Tenta vários formatos e retorna datetime.date ou None."""
    for fmt in ("%d/%m/%y", "%d/%m/%Y"):
        try:
            return datetime.strptime(s.strip(), fmt).date()
        except Exception:
            continue
    return None

def log(msg: str):
    print(f"[{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] {msg}", flush=True)

atualizar_contas()
if not CONTAS_COUPA:
    print("ERRO: nenhuma conta Coupa com usuario e senha (nem no banco, nem no .env).")

def obter_marcas_em_lote(descricoes: list, tamanho_lote: int = 60) -> list:
    """Extrai a marca do produto a partir da DESCRIÇÃO via Gemini, em lote.
    Retorna lista do mesmo tamanho/ordem de `descricoes`. Nunca propaga exceção."""
    resultado = ["Marca nao encontrada"] * len(descricoes)

    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        log("AVISO: GEMINI_API_KEY nao configurada. Pulando extracao de marca.")
        return resultado

    try:
        from google import genai
        from google.genai import types as genai_types
    except Exception as e:
        log(f"AVISO: SDK google-genai indisponivel ({e}). Pulando extracao de marca.")
        return resultado

    modelo = os.getenv("GEMINI_MODEL", "gemini-3.1-flash-lite")
    client = genai.Client(api_key=api_key)

    total = len(descricoes)
    for inicio in range(0, total, tamanho_lote):
        fim = min(inicio + tamanho_lote, total)
        lote = descricoes[inicio:fim]

        itens = "\n".join(f"{i}: {(desc or '').strip()}" for i, desc in enumerate(lote))
        prompt = (
            "Voce e um extrator de marcas de produtos industriais. Para cada item numerado "
            "abaixo, identifique a MARCA do produto (fabricante), seguindo estas regras:\n"
            "- So extraia a marca se ela estiver EXPLICITAMENTE escrita no texto. NUNCA "
            "invente ou deduza uma marca a partir do tipo de produto.\n"
            "- NUNCA retorne \"VALE\" como marca — e o nome do cliente comprador, nao do "
            "fabricante.\n"
            "- Se nao houver marca identificavel, retorne exatamente \"Marca nao encontrada\".\n"
            "- Retorne a marca em maiusculas, sem codigo de peca junto.\n"
            "- Responda APENAS com um JSON array puro, sem markdown, no formato: "
            '[{"indice": N, "marca": "..."}]\n\n'
            f"Itens:\n{itens}"
        )

        try:
            response = client.models.generate_content(
                model=modelo,
                contents=prompt,
                config=genai_types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0,
                ),
            )
            dados = json.loads(response.text)
            preenchidos = 0
            for item in dados:
                idx = item.get("indice")
                marca = item.get("marca")
                if isinstance(idx, int) and 0 <= idx < len(lote) and marca:
                    resultado[inicio + idx] = str(marca).strip().upper()
                    preenchidos += 1
            log(f"Marcas: lote {inicio}-{fim - 1}: {preenchidos}/{len(lote)} preenchidas.")
        except Exception as e:
            log(f"AVISO: falha ao obter marcas do lote {inicio}-{fim - 1}: {e}")
            continue

    return resultado

def executar_robo_selenium(data_usuario: str, filename: str, conta_id: str, pular_verificacao_duplicados: bool = False, coletor=None, simular: bool = False, dados_conta: dict = None) -> str:
    """Lógica principal do robô Selenium refatorada para rodar sem interface (headless).
    Retorna "ok" ou uma descrição curta do erro.

    coletor: se for uma lista, o robô NÃO envia a planilha da conta ao Storage; entrega as linhas
    finais (9 colunas) na lista, para a fila juntar várias contas numa planilha única.

    simular: loga no Coupa e coleta de verdade, mas NÃO consulta nem grava no banco e NÃO envia ao
    Storage. A planilha fica na pasta SAIDA_TESTE do servidor."""

    conta = dados_conta or CONTAS_COUPA[conta_id]
    inicio_total = time.time()
    log(f"=== ROBÔ INICIADO | conta={conta_id} | data={data_usuario} | arquivo={filename} ===")
    log(f"Credenciais da conta {conta_id}: usuario do {conta.get('fonte_usuario') or 'env'}, senha do {conta.get('fonte_senha') or 'env'}")
    if simular:
        log("MODO SIMULACAO: nada sera gravado no banco nem enviado ao Storage; a planilha ficara em " + SAIDA_TESTE)
    if pular_verificacao_duplicados:
        log("MODO TESTE: verificacao de eventos duplicados DESATIVADA - eventos ja existentes serao reprocessados")

    # 1. Preparar data (Aceita 6 ou 8 dígitos: DDMMAA ou DDMMAAAA)
    if len(data_usuario) == 6:
        HOJE_str = f"{data_usuario[:2]}/{data_usuario[2:4]}/{data_usuario[4:]}"
    else:
        HOJE_str = f"{data_usuario[:2]}/{data_usuario[2:4]}/{data_usuario[4:]}" # Pega os 4 dígitos do ano se houver

    HOJE = parse_date_str(HOJE_str)
    if not HOJE:
        log(f"ERRO: Não foi possível converter a data '{data_usuario}'")
        return "data invalida"

    # 2. Configurar Selenium Headless para o Servidor (Railway/Docker)
    chrome_options = Options()
    if os.getenv("ROBO_VISIVEL") != "1":
        chrome_options.add_argument("--headless=new") # IMPORTANTE: Modo invisível (ROBO_VISIVEL=1 mostra a janela, so para teste local)
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    
    # --- OTIMIZAÇÕES REDUÇÃO DE MEMÓRIA E CPU ---
    chrome_options.add_argument("--disable-extensions") # Não carregar extensões
    chrome_options.add_argument("--disable-software-rasterizer") # Exige menos da CPU onde já não tem GPU
    chrome_options.add_argument("--disable-logging") # Menos escrita de logs nativos do chrome
    chrome_options.add_argument("--log-level=3")     # Silencia saída desnecessária
    chrome_options.add_argument("--blink-settings=imagesEnabled=false") # 🚀 DESATIVA IMAGENS (salva MUITA RAM e Banda)
    chrome_options.page_load_strategy = 'eager' # 🚀 CARREGA MAIS RÁPIDO (não espera recursos inúteis da página, só o DOM)
    # Reduz sinais de automacao (autorizado pelo usuario): no PC de teste, sem isto o Coupa nao
    # respondia (ERR_TIMED_OUT); com isto a pagina de login carregou em ~2,5s (teste_chrome2).
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)
    
    # Localiza o binário do Chrome no Railway
    chrome_bin = os.getenv("CHROME_BIN")
    if chrome_bin:
        chrome_options.binary_location = chrome_bin
    
    # Inicializa o driver apenas UMA vez com as opções corretas
    driver = webdriver.Chrome(options=chrome_options)
    
    EXCEL_PATH = os.path.join(tempfile.gettempdir(), filename)
    ESTADOS = ['AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MT', 'MS', 'MG', 'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN', 'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO']

    try:
        # Prepara Planilha
        wb = Workbook()
        ws = wb.active
        ws.title = "Eventos"
        ws.append(["Numero do evento", "Titulo", "UF(VALE)", "DATA", "DESCRIÇÃO", "QTDE", "UNID. MED", "pagina de descrição", "Marca"])
        
        wait = WebDriverWait(driver, 20)
        log("Acessando página de login...")
        driver.get("https://vale.coupahost.com/sessions/supplier_login")

        # Login
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "user_login")))
        driver.find_element(By.ID, "user_login").send_keys(conta["user"])
        driver.find_element(By.ID, "user_password").send_keys(conta["senha"], Keys.RETURN)
        log(f"Login enviado (conta {conta_id}). Aguardando carregamento...")

        # Se o formulário de login continuar na tela, a credencial foi recusada.
        # Sem essa checagem o robô seguia e subia uma planilha vazia como se tivesse dado certo.
        try:
            WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.ID, "user_login")))
        except Exception:
            raise RuntimeError(f"Login no Coupa falhou para a conta {conta_id} (formulario de login ainda visivel)")

        # Filtro de data (mesma lógica)
        try:
            time_filter = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ch_start_time"]')))
            time_filter.click()
            time.sleep(5)
            time_filter = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ch_start_time"]')))
            time_filter.click()
            log("Filtro de data aplicado.")
        except:
            log("AVISO: Filtro de data não encontrado, continuando sem ele.")

        # Coleta de dados - Parte 1: Listagem de Eventos
        log(f"--- FASE 1: Coletando lista de eventos para {HOJE_str} ---")
        encontrou_ontem = False
        pagina = 1
        total_coletados = 0
        motivos_total = Counter()

        def coletar_evento(numero_evento, data_final):
            """Chamada para cada evento da data pedida. Devolve um motivo se NAO deve coletar."""
            nonlocal total_coletados
            # Verifica se o evento já existe no banco de dados
            if not pular_verificacao_duplicados and not simular:
                try:
                    res_db = (
                        supabase.table("eventos_coletados")
                        .select("id")
                        .eq("conta", conta_id)
                        .eq("numero_evento", numero_evento)
                        .execute()
                    )
                    if res_db.data:
                        log(f"AVISO: Evento {numero_evento} ja existe no banco. Pulando.")
                        return "ja_no_banco"
                except Exception as e_db:
                    log(f"ERRO ao verificar evento {numero_evento} no banco: {e_db}")
            total_coletados += 1
            log(f"  [{total_coletados}] Evento coletado: {numero_evento} | prazo: {data_final}")
            ws.append([numero_evento, '', '', data_final, '', '', '', '', ''])
            return None

        while True:
            log(f"Pagina {pagina}: aguardando tabela de eventos...")
            time.sleep(5)
            try:
                tbody = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="quote_request_table_tag"]')))
                # Usar index para evitar stale element reference
                num_linhas = len(tbody.find_elements(By.TAG_NAME, "tr"))
                log(f"Pagina {pagina}: {num_linhas} linhas encontradas na tabela.")

                resumo = varrer_pagina(driver, HOJE, parse_date_str, num_linhas, coletar_evento, log)
                motivos_total.update(resumo["motivos"])
                log(resumo_texto(pagina, num_linhas, resumo))
                if resumo["encontrou_anterior"]:
                    encontrou_ontem = True
            except Exception as e:
                log(f"ERRO ao acessar tabela na pagina {pagina}: {e}")
                break

            if encontrou_ontem: break
            try:
                proximo = driver.find_element(By.CLASS_NAME, "next_page")
                driver.execute_script("arguments[0].click();", proximo)
                pagina += 1
                time.sleep(3)
            except:
                log(f"Sem proxima pagina. Total de paginas percorridas: {pagina}.")
                break

        log(f"Fase 1: linhas puladas no total: {dict(motivos_total)}")
        wb.save(EXCEL_PATH)
        log(f"Fase 1 concluida: {total_coletados} evento(s) coletado(s). Iniciando detalhamento...")

        # --- DETALHA CADA EVENTO ---
        wb = load_workbook(EXCEL_PATH)
        ws = wb["Eventos"]

        todos_eventos = [row[0].value for row in ws.iter_rows(min_row=2) if row[0].value]
        total_eventos = len(todos_eventos)
        log(f"--- FASE 2: Detalhando {total_eventos} evento(s) ---")

        for idx_ev, row in enumerate(ws.iter_rows(min_row=2), start=1):
            evento = row[0].value
            if not evento:
                continue

            log(f"[{idx_ev}/{total_eventos}] Detalhando evento: {evento}")

            driver.get(f"https://vale.coupahost.com/quotes/external_responses/{evento}/edit")
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

            # --- VERIFICA EXISTÊNCIA DA PÁGINA DE DESCRIÇÃO ---
            try:
                botoes1 = driver.find_elements(By.XPATH, '//*[@id="pageContentWrapper"]/div[3]/div[2]/a[2]/span')
                if not botoes1:
                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    botoes2 = driver.find_elements(By.ID, 'quote_response_submit')
                    if botoes2:
                        log(f"  Evento {evento}: botao de submit encontrado, clicando...")
                        botoes2[0].click()
            except Exception:
                log(f"  ERRO ao verificar pagina de descricao do evento {evento}")
                row[7].value = "Erro ao verificar página de descrição"

            # Scroll e abre seção das informações
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            # Tenta encontrar o seletor principal ou o fallback conforme solicitado
            try:
                wait.until(EC.presence_of_element_located((By.CLASS_NAME, "s-expandLines")))
                seletor_atual = (By.CLASS_NAME, "s-expandLines")
                query_css = ".s-expandLines"
            except:
                log("⚠️ s-expandLines não encontrado, tentando fallback...")
                fallback_css = ".sidebar.-supplier.-borderLeft.flexPosition__element.-shrink.s-expandSidebar.-clickable"
                try:
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, fallback_css)))
                    seletor_atual = (By.CSS_SELECTOR, fallback_css)
                    query_css = fallback_css
                except:
                    log(f"⚠️ Nenhum seletor de expansão encontrado no evento {evento}")
                    continue

            elementos = driver.find_elements(*seletor_atual)

            if not elementos:
                log(f"⚠️ Nenhum elemento de expansão encontrado no evento {evento}")
                continue

            # Duplicar a linha do evento pelo número de elementos encontrados
            linhas_evento = [row]
            if len(elementos) > 1:
                for i in range(len(elementos) - 1):
                    nova_linha = [evento, row[1].value, row[2].value, row[3].value, '', '', '', '', '']
                    ws.append(nova_linha)
                wb.save(EXCEL_PATH)
                linhas_evento = [r for r in ws.iter_rows(min_row=2) if r[0].value == evento]

            # Percorre cada s-expandLines e coleta os dados (re-fetch a cada iteração, marca processed via JS)
            def click_element_retry(el, attempts=4, pause=0.4):
                from selenium.common.exceptions import (
                    StaleElementReferenceException,
                    ElementClickInterceptedException,
                    ElementNotInteractableException,
                    WebDriverException,
                )
                for _ in range(attempts):
                    try:
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                        time.sleep(0.15)
                        el.click()
                        return True
                    except (StaleElementReferenceException, ElementClickInterceptedException, ElementNotInteractableException, WebDriverException):
                        try:
                            driver.execute_script("arguments[0].click();", el)
                            return True
                        except Exception:
                            time.sleep(pause)
                return False

            # determina quantos existem no DOM no momento
            total = driver.execute_script(f"return document.querySelectorAll('{query_css}').length")
            if total == 0:
                log(f"⚠️ Nenhum elemento de expansão encontrado no evento {evento}")
                continue

            # duplicar linha já feito acima; garante linhas_evento atualizado
            linhas_evento = [r for r in ws.iter_rows(min_row=2) if r[0].value == evento]

            processed = 0
            idx = 0
            while processed < total and idx < total:
                # re-obtem a lista sempre
                try:
                    elementos = driver.find_elements(*seletor_atual)
                except Exception:
                    time.sleep(0.3)
                    elementos = driver.find_elements(*seletor_atual)

                if idx >= len(elementos):
                    # DOM encolheu — tenta refetch algumas vezes
                    retry_try = 0
                    while retry_try < 3 and idx >= len(elementos):
                        time.sleep(0.4)
                        elementos = driver.find_elements(By.CLASS_NAME, "s-expandLines")
                        retry_try += 1
                    if idx >= len(elementos):
                        log(f"⚠️ Índice {idx} fora do range atual ({len(elementos)}). Pulando.")
                        idx += 1
                        continue

                el = elementos[idx]

                # evita re-processar elemento já marcado
                try:
                    already = el.get_attribute('data-processed')
                except:
                    already = None

                if already:
                    idx += 1
                    processed += 1
                    continue

                # tenta clicar de forma robusta
                if not click_element_retry(el, attempts=4, pause=0.4):
                    log(f"⚠️ Falha ao clicar no expandLines index {idx} do evento {evento}")
                    # marca como processado para não travar loop
                    try:
                        driver.execute_script("arguments[0].setAttribute('data-processed','1')", el)
                    except Exception:
                        pass
                    idx += 1
                    processed += 1
                    continue

                # após clique, espera conteúdo de detalhe carregar (linha expandida)
                try:
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".s-itemsAndServicesLine.-expanded")))
                    time.sleep(0.25)
                except Exception:
                    time.sleep(0.4)

                # atualiza linhas_evento porque podem ter sido adicionadas
                linhas_evento = [r for r in ws.iter_rows(min_row=2) if r[0].value == evento]
                try:
                    linha_atual = linhas_evento[idx]
                except Exception:
                    # se não existir, tenta mapear para próxima disponível
                    if linhas_evento:
                        linha_atual = linhas_evento[-1]
                    else:
                        log(f"⚠️ Não há linha disponível para evento {evento} no idx {idx}")
                        # marca e segue
                        try:
                            driver.execute_script("arguments[0].setAttribute('data-processed','1')", el)
                        except Exception:
                            pass
                        idx += 1
                        processed += 1
                        continue

                # coleta campos — seletores por classe CSS, compatíveis com ambos layouts
                log(f"  Item {idx + 1}/{total}: coletando campos...")

                # Escopo: a linha atualmente expandida
                try:
                    wait.until(EC.presence_of_element_located(
                        (By.CSS_SELECTOR, ".s-itemsAndServicesLine.-expanded")
                    ))
                    linha_dom = driver.find_element(By.CSS_SELECTOR, ".s-itemsAndServicesLine.-expanded")
                except Exception:
                    time.sleep(1)
                    try:
                        linha_dom = driver.find_element(By.CSS_SELECTOR, ".s-itemsAndServicesLine.-expanded")
                    except Exception:
                        linha_dom = driver

                def buscar_texto(seletores, escopo=None):
                    """Tenta uma lista de seletores CSS e retorna o primeiro texto não vazio."""
                    alvo = escopo if escopo is not None else linha_dom
                    for sel in seletores:
                        try:
                            el = alvo.find_element(By.CSS_SELECTOR, sel)
                            txt = (el.text or "").strip()
                            if txt:
                                return txt
                        except Exception:
                            continue
                    return ""

                # coletar_descricoes vem de coupa_descricao.py (compartilhado com sevenmaisdeum.py)

                # --- QUANTIDADE ---
                qtde = buscar_texto([".s-quantity .s-value"])
                linha_atual[5].value = qtde if qtde else 'N/A'

                # --- UNIDADE ---
                unid = buscar_texto([".s-quantity .s-unit"])
                linha_atual[6].value = unid if unid else 'N/A'

                # --- DESCRIÇÃO (concatena todos os campos preenchidos) ---
                descri, campos_usados = coletar_descricoes(linha_dom)
                if descri:
                    linha_atual[4].value = descri
                    log(f"  Item {idx + 1}/{total}: descricao de {campos_usados}")
                else:
                    linha_atual[4].value = 'N/A'
                    log(f"  AVISO: nenhum campo de descricao encontrado no evento {evento} item {idx+1}")

                # --- TÍTULO (aceita separador "||" ou "|") ---
                titulo_raw = buscar_texto([".s-description p.s-textField"])
                if titulo_raw:
                    m_tit = re.match(r"^\s*(\d+)", titulo_raw)
                    linha_atual[1].value = m_tit.group(1) if m_tit else titulo_raw.split("|")[0].strip()
                else:
                    linha_atual[1].value = "Titulo nao encontrado"

                # --- UF (cascata: endereço de entrega -> varredura geral da linha) ---
                found = None
                endereco_txt = buscar_texto([
                    ".s-ship_to_address .s-itemsAndServicesAddressLine",
                    ".s-ship_to_address .addressLines",
                ])

                if endereco_txt:
                    texto = endereco_txt.upper()
                    # padrão "CEP CIDADE UF"
                    m_uf = re.search(r'\d{5}-?\d{3}\s+.+?\s+([A-Z]{2})\s*(?:\n|$)', texto)
                    if m_uf and m_uf.group(1) in ESTADOS:
                        found = m_uf.group(1)
                    if not found:
                        m_uf = re.search(r'-\s*([A-Z]{2})\s*-\s*BR', texto)
                        if m_uf and m_uf.group(1) in ESTADOS:
                            found = m_uf.group(1)
                    if not found:
                        for t in re.findall(r'\b[A-Z]{2}\b', texto):
                            if t in ESTADOS:
                                found = t
                                break

                # fallback: varre todo o texto da linha expandida
                if not found:
                    try:
                        combined = (linha_dom.text or "").upper()
                        m_uf = re.search(r'-\s*([A-Z]{2})\s*-\s*BR', combined)
                        if m_uf and m_uf.group(1) in ESTADOS:
                            found = m_uf.group(1)
                        if not found:
                            for sig in ESTADOS:
                                if re.search(r'\b' + re.escape(sig) + r'\b', combined):
                                    found = sig
                                    break
                    except Exception:
                        pass

                linha_atual[2].value = found if found else 'UF nao encontrada'
                if not found:
                    log(f"  AVISO: UF nao encontrada no evento {evento} item {idx+1}")

                log(f"  Item {idx + 1}/{total}: titulo={linha_atual[1].value!r} | uf={linha_atual[2].value!r} | qtde={linha_atual[5].value!r} | unid={linha_atual[6].value!r}")

                # fecha o detalhe (tenta vários métodos)
                try:
                    time.sleep(0.2)
                    fechar = None
                    try:
                        fechar = driver.find_element(By.CSS_SELECTOR, "button.button.s-cancel")
                    except Exception:
                        try:
                            fechar = driver.find_element(By.XPATH, "//button[contains(concat(' ', normalize-space(@class), ' '), ' s-cancel ') and contains(., 'Cancelar')]")
                        except Exception:
                            fechar = None
                    if fechar:
                        click_element_retry(fechar, attempts=3, pause=0.2)
                        time.sleep(0.25)
                except Exception:
                    pass

                # marca como processado (para não reprocessar se DOM reorganizar)
                try:
                    driver.execute_script("arguments[0].setAttribute('data-processed','1')", el)
                except Exception:
                    pass

                processed += 1
                idx += 1

            wb.save(EXCEL_PATH)
            log(f"[{idx_ev}/{total_eventos}] Evento {evento} finalizado.")

        log("--- FASE 3: Ordenando planilha e salvando no banco ---")
        linhas_finais = None
        # Ordena a planilha por "Numero do evento" (coluna A) para agrupar linhas com o mesmo número
        try:
            wb = load_workbook(EXCEL_PATH)
            ws = wb["Eventos"]
            rows = list(ws.iter_rows(min_row=2, values_only=True))

            descricoes = [r[4] if len(r) > 4 else '' for r in rows]
            log(f"--- Extraindo marcas via Gemini para {len(descricoes)} item(ns) ---")
            marcas = obter_marcas_em_lote(descricoes)
            rows = [r[:8] + (marcas[i],) for i, r in enumerate(rows)]

            def sort_key(row):
                v = row[0]
                if v is None:
                    return (1, "")
                s = str(v).strip()
                try:
                    return (0, int(s))      # números antes de strings, ordenados numericamente
                except Exception:
                    return (1, s.lower())   # strings ordenadas alfabeticamente

            rows_sorted = sorted(rows, key=sort_key)

            # remove linhas antigas (todas a partir da linha 2) e escreve ordenado
            if ws.max_row > 1:
                ws.delete_rows(2, ws.max_row - 1)
            
            eventos_para_inserir = []
            vistos_db = set()

            for r in rows_sorted:
                ws.append(list(r))
                ne = str(r[0]).strip()
                titulo = str(r[1]) if len(r) > 1 and r[1] is not None else ""
                chave = (ne, titulo)
                if ne and ne != "None" and chave not in vistos_db:
                    vistos_db.add(chave)
                    eventos_para_inserir.append(montar_evento_db(conta_id, r, HOJE))

            wb.save(EXCEL_PATH)
            linhas_finais = rows_sorted

            # Salva os novos eventos no banco de dados para evitar futuras duplicidades
            if simular:
                log(f"SIMULACAO: {len(eventos_para_inserir)} evento(s) NAO foram gravados no banco.")
            elif eventos_para_inserir:
                try:
                    # Upsert precisa de UPDATE quando o evento ja existe; a chave publica so tem INSERT
                    # (RLS), entao a gravacao usa a chave de servidor quando ela esta configurada.
                    cliente_gravacao = supabase_admin if supabase_admin is not None else supabase
                    cliente_gravacao.table("eventos_coletados").upsert(eventos_para_inserir, on_conflict="conta,numero_evento,titulo").execute()
                    log(f"Banco atualizado: {len(eventos_para_inserir)} evento(s) inserido(s)/atualizados.")
                except Exception as e_db:
                    log(f"ERRO ao inserir eventos no banco: {e_db}")
            else:
                log("Nenhum evento novo para inserir no banco.")

        except Exception as e:
            log(f"ERRO ao ordenar/salvar planilha ou banco: {e}")

        if coletor is not None:
            # Modo planilha unificada: quem envia o arquivo é a fila, depois de juntar as contas.
            if linhas_finais is None:
                raise RuntimeError("falha ao consolidar as linhas da conta (ver o log da fase 3)")
            coletor.extend(linhas_finais)
            log(f"Conta {conta_id}: {len(linhas_finais)} linha(s) entregues para a planilha unificada.")
            if os.path.exists(EXCEL_PATH):
                os.remove(EXCEL_PATH)
            return "ok"

        if simular:
            guardar_saida_teste(EXCEL_PATH, filename)
            if os.path.exists(EXCEL_PATH):
                os.remove(EXCEL_PATH)
            return "ok"

        log("--- FASE 4: Fazendo upload para o Supabase Storage ---")
        enviar_planilha(EXCEL_PATH, filename)

        if os.path.exists(EXCEL_PATH):
            os.remove(EXCEL_PATH)
        return "ok"

    except Exception as e:
        log(f"ERRO FATAL no robô (conta {conta_id}): {type(e).__name__}: {str(e).splitlines()[0] if str(e) else '(sem mensagem)'}")
        # Diagnóstico: em que página o robô estava quando falhou (não altera o fluxo)
        try:
            log(f"DIAGNOSTICO: url={driver.current_url} | titulo={driver.title!r}")
            if simular:
                os.makedirs(SAIDA_TESTE, exist_ok=True)
                png = os.path.abspath(os.path.join(SAIDA_TESTE, f"erro_{conta_id}_{int(time.time())}.png"))
                driver.save_screenshot(png)
                log(f"DIAGNOSTICO: imagem da pagina salva em {png}")
        except Exception as e_diag:
            log(f"DIAGNOSTICO indisponivel: {type(e_diag).__name__}")
        log(f"Arquivo Excel preservado para recuperacao manual em: {EXCEL_PATH}")
        return f"erro: {e}"
    finally:
        driver.quit()
        elapsed = time.time() - inicio_total
        log(f"=== ROBÔ FINALIZADO | conta={conta_id} | tempo total: {elapsed:.1f}s ===")

SAIDA_TESTE = os.getenv("SAIDA_TESTE", "saida_teste")

def guardar_saida_teste(caminho: str, filename: str):
    os.makedirs(SAIDA_TESTE, exist_ok=True)
    destino = os.path.abspath(os.path.join(SAIDA_TESTE, filename))
    shutil.copy(caminho, destino)
    log(f"SIMULACAO: planilha salva em {destino}")

def enviar_planilha(caminho: str, filename: str):
    with open(caminho, "rb") as f:
        supabase.storage.from_(SUPABASE_BUCKET).upload(
            path=filename,
            file=f,
            file_options={"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
        )
    res = supabase.storage.from_(SUPABASE_BUCKET).get_public_url(filename)
    log(f"Upload concluido. URL publica: {res}")

CABECALHO_PLANILHA = ["Numero do evento", "Titulo", "UF(VALE)", "DATA", "DESCRIÇÃO", "QTDE", "UNID. MED", "pagina de descrição", "Marca"]
_VALORES_VAZIOS = {"", "none", "n/a", "uf nao encontrada", "titulo nao encontrado", "marca nao encontrada"}
_CAMPOS_CONFERIDOS = {2: "UF", 3: "DATA", 4: "DESCRICAO", 5: "QTDE", 6: "UNID"}

def _vazio(v) -> bool:
    return v is None or str(v).strip().lower() in _VALORES_VAZIOS

def _chave_ordenacao(row):
    v = row[0]
    if v is None:
        return (1, "")
    s = str(v).strip()
    try:
        return (0, int(s))      # números antes de strings, ordenados numericamente
    except Exception:
        return (1, s.lower())   # strings ordenadas alfabeticamente

def montar_planilha_unificada(por_conta: dict, nomes: dict, caminho: str):
    """Junta as linhas de várias contas numa planilha só.

    Linha repetida entre contas (mesmo evento e mesmo item/título) aparece uma única vez, com a
    coluna "Filiais" dizendo em quais contas ela existe. Repetição dentro da MESMA conta não é
    unificada: são itens distintos que só compartilham número e título (ex.: "Titulo nao
    encontrado" em vários itens do mesmo evento). Por isso a k-ésima ocorrência de uma chave numa
    conta casa com a k-ésima ocorrência da mesma chave na outra.

    Retorna (linhas_recebidas, linhas_na_planilha)."""
    unicos = {}
    total = 0
    for conta_id, linhas in por_conta.items():
        vistas = {}
        for r in linhas:
            total += 1
            linha = (list(r) + [None] * 9)[:9]
            base_chave = (str(linha[0]).strip(), "" if linha[1] is None else str(linha[1]).strip())
            ordem = vistas.get(base_chave, 0)
            vistas[base_chave] = ordem + 1
            chave = (base_chave, ordem)

            item = unicos.get(chave)
            if item is None:
                unicos[chave] = {"linha": linha, "contas": [conta_id]}
                continue

            item["contas"].append(conta_id)
            base = item["linha"]
            divergentes = []
            for i in range(9):
                if _vazio(base[i]) and not _vazio(linha[i]):
                    base[i] = linha[i]  # completa o que a primeira conta não conseguiu ler
                elif i in _CAMPOS_CONFERIDOS and not _vazio(linha[i]) and str(base[i]).strip() != str(linha[i]).strip():
                    divergentes.append(_CAMPOS_CONFERIDOS[i])
            if divergentes:
                primeira = nomes.get(item["contas"][0], item["contas"][0])
                log(f"AVISO: evento {base_chave[0]} item {base_chave[1]!r} difere entre filiais em {', '.join(divergentes)}; mantido o valor de {primeira}.")

    ordenadas = sorted(unicos.values(), key=lambda it: _chave_ordenacao(it["linha"]))

    wb = Workbook()
    ws = wb.active
    ws.title = "Eventos"
    ws.append(CABECALHO_PLANILHA + ["Filiais"])
    for it in ordenadas:
        ws.append(it["linha"] + [", ".join(nomes.get(c, c) for c in it["contas"])])
    wb.save(caminho)
    return total, len(ordenadas)

def nome_arquivo(conta_id: str, data_formatada: str, conta_padrao: str = None) -> str:
    # A primeira conta mantém o nome de arquivo de sempre, para não quebrar
    # nada que dependa dele; as demais levam o identificador da conta.
    if conta_id == (conta_padrao or CONTA_PADRAO):
        return f"eventos_{data_formatada}_{int(time.time())}.xlsx"
    return f"eventos_{conta_id}_{data_formatada}_{int(time.time())}.xlsx"

def executar_fila(data: str, data_formatada: str, contas: list, pular_verificacao_duplicados: bool, simular: bool = False, retrato: dict = None, conta_padrao: str = None):
    """Roda as contas em sequência e libera o lock no final, aconteça o que acontecer.
    Com mais de uma conta, gera UMA planilha só, sem repetir o evento que chegou às duas."""
    resultados = {}
    por_conta = {}
    unificar = len(contas) > 1
    retrato = retrato or dict(CONTAS_COUPA)
    try:
        for conta_id in contas:
            ESTADO_ROBO.update({"conta": conta_id, "inicio": datetime.now().isoformat(timespec="seconds")})
            coletor = [] if unificar else None
            try:
                resultados[conta_id] = executar_robo_selenium(
                    data, nome_arquivo(conta_id, data_formatada, conta_padrao), conta_id, pular_verificacao_duplicados, coletor, simular,
                    retrato[conta_id]
                )
            except Exception as e:
                log(f"ERRO FATAL ao iniciar robô da conta {conta_id}: {e}")
                resultados[conta_id] = f"erro: {e}"
            if unificar and resultados[conta_id] == "ok":
                por_conta[conta_id] = coletor

        if unificar and por_conta:
            nome = f"eventos_todas_{data_formatada}_{int(time.time())}.xlsx"
            caminho = os.path.join(tempfile.gettempdir(), nome)
            try:
                total, unicas = montar_planilha_unificada(
                    por_conta, {c: retrato[c]["nome"] for c in por_conta}, caminho
                )
                log(f"Planilha unificada: {total} linha(s) recebidas, {unicas} na planilha ({total - unicas} repetida(s) entre filiais).")
                if simular:
                    guardar_saida_teste(caminho, nome)
                else:
                    enviar_planilha(caminho, nome)
                os.remove(caminho)
            except Exception as e:
                # O arquivo fica em /tmp para recuperação manual: o banco já recebeu os eventos,
                # então rodar de novo pularia tudo como "já existente".
                log(f"ERRO ao gerar/enviar a planilha unificada: {e}. Arquivo preservado em {caminho}")
                resultados["todas"] = f"erro: falha ao gerar a planilha unificada ({e})"
    finally:
        ESTADO_ROBO.update({
            "executando": False,
            "conta": None,
            "inicio": None,
            "ultima_execucao": {
                "data": data,
                "fim": datetime.now().isoformat(timespec="seconds"),
                "resultados": resultados,
            },
        })
        ROBO_LOCK.release()

@app.get("/")
def read_root():
    return {"status": "Backend Online", "projeto": "Seven Suprimentos - Automação Vale"}

@app.get("/contas")
def listar_contas():
    """Contas Coupa prontas para rodar. Nunca expõe usuário ou senha."""
    atualizar_contas()
    return {
        "padrao": CONTA_PADRAO,
        "contas": [{"id": c["id"], "nome": c["nome"]} for c in CONTAS_COUPA.values()],
    }

@app.get("/status")
def status_robo():
    return ESTADO_ROBO

@app.post("/run-robot")
def run_robot(data: str, background_tasks: BackgroundTasks, conta: str = "", pular_verificacao_duplicados: bool = False, simular: bool = False):
    """Inicia o robô como uma tarefa de fundo.
    conta: id da conta Coupa (ver /contas). Vazio = conta padrão.
    "todas" = todas em sequência, numa planilha única (evento repetido entre contas sai uma vez)."""
    if not data or len(data) not in [6, 8] or not data.isdigit():
        raise HTTPException(status_code=400, detail="Data deve estar no formato DDMMAA ou DDMMAAAA")
    # ROBO_SIMULAR=1 no .env força a simulação em qualquer chamada (útil para testar pelo front
    # numa máquina local sem arriscar o banco de produção).
    simular = simular or os.getenv("ROBO_SIMULAR") == "1"
    atualizar_contas()
    retrato = dict(CONTAS_COUPA)
    conta_padrao = CONTA_PADRAO
    if not CONTAS_COUPA:
        motivo = "nenhuma tem usuario e senha validos (ver Credenciais)" if CONTAS_TODAS else "nenhuma configurada no servidor"
        raise HTTPException(status_code=500, detail=f"Contas Coupa: {motivo}.")

    conta = (conta or "").strip().lower()
    if not conta:
        contas = [conta_padrao if conta_padrao in retrato else next(iter(retrato))]
    elif conta == "todas":
        contas = list(CONTAS_COUPA)
    elif conta in CONTAS_COUPA:
        contas = [conta]
    else:
        sem_senha = any(c["id"] == conta for c in CONTAS_TODAS)
        motivo = "esta sem usuario ou senha (ver Credenciais)" if sem_senha else "nao existe"
        raise HTTPException(status_code=400, detail=f"Conta '{conta}' {motivo}. Disponiveis: {', '.join(CONTAS_COUPA) or 'nenhuma'}")

    if not ROBO_LOCK.acquire(blocking=False):
        em_uso = ESTADO_ROBO.get("conta") or "outra conta"
        raise HTTPException(
            status_code=409,
            detail=f"O robô já está em execução ({em_uso}, data {ESTADO_ROBO.get('data')}). Aguarde terminar para iniciar outra.",
        )
    ESTADO_ROBO.update({"executando": True, "data": data, "conta": contas[0]})

    # Padroniza para DDMMAA para gerar o nome do arquivo, mas passa a data completa
    data_formatada = data if len(data) == 6 else f"{data[:4]}{data[6:]}"

    try:
        background_tasks.add_task(executar_fila, data, data_formatada, contas, pular_verificacao_duplicados, simular, retrato, conta_padrao)
    except Exception:
        ESTADO_ROBO["executando"] = False
        ROBO_LOCK.release()
        raise

    nomes = ", ".join(retrato[c]["nome"] for c in contas)
    return {
        "status": "Iniciado",
        "message": (
            f"SIMULACAO iniciada para a data {data} ({nomes}). Nada sera gravado no banco nem no Storage; a planilha ficara na pasta {SAIDA_TESTE} do servidor."
            if simular else
            f"Robô iniciado para a data {data} ({nomes}). O arquivo será enviado para o Supabase Storage em alguns minutos."
        ),
        "contas": contas,
        "modo_teste_duplicados": pular_verificacao_duplicados,
        "modo_simulacao": simular
    }

# ---------------------------------------------------------------------------
# Gerenciador de credenciais (painel). Exige login do painel (token do Supabase Auth) e
# papel 'admin' ativo na tabela public.perfis (Seven OS). A senha so entra: nenhuma rota
# devolve senha, nem cifrada.
# ---------------------------------------------------------------------------
def _usuario_do_token(authorization: str) -> dict:
    """Valida o token do painel e devolve {id, email, papel, ativo} (papel vem de public.perfis)."""
    if supabase_admin is None:
        raise HTTPException(status_code=503, detail="Gerenciador de credenciais indisponivel: SUPABASE_SECRET_KEY nao configurada no servidor.")
    if not authorization or not authorization.lower().startswith("bearer "):
        raise HTTPException(status_code=401, detail="Faca login no painel.")
    try:
        resp = supabase_admin.auth.get_user(authorization.split(" ", 1)[1].strip())
        user = resp.user if resp else None
    except Exception:
        user = None
    if not user or not user.id:
        raise HTTPException(status_code=401, detail="Sessao invalida ou expirada. Entre de novo.")
    try:
        dados = (supabase_admin.table("perfis").select("papel,ativo")
                 .eq("user_id", str(user.id)).limit(1).execute().data or [])
    except Exception as e:
        log(f"ERRO ao ler perfil do usuario: {type(e).__name__}")
        raise HTTPException(status_code=503, detail="Nao foi possivel ler o seu perfil. A tabela perfis existe?")
    perfil = dados[0] if dados else {}
    return {
        "id": str(user.id), "email": (user.email or "").lower(),
        "papel": perfil.get("papel"), "ativo": bool(perfil.get("ativo")),
    }

def _exigir_admin(authorization: str) -> str:
    u = _usuario_do_token(authorization)
    if u["papel"] != "admin" or not u["ativo"]:
        raise HTTPException(status_code=403, detail="Seu usuario nao pode alterar credenciais (precisa do papel admin).")
    email = u["email"]
    if not credenciais.chave_configurada():
        raise HTTPException(status_code=503, detail="CREDENCIAIS_CHAVE nao configurada (ou invalida) no servidor.")
    return email

def _publico(c: dict) -> dict:
    return {
        "id": c["id"], "nome": c["nome"], "usuario": c["user"] or None,
        "tem_senha": bool(c["senha"]), "fonte_usuario": c["fonte_usuario"], "fonte_senha": c["fonte_senha"],
        "pronta": credenciais.pronta(c), "erro": c["erro"],
        "atualizado_em": c["atualizado_em"], "atualizado_por": c["atualizado_por"],
    }

@app.get("/eu")
def quem_sou(authorization: str = Header(default="")):
    u = _usuario_do_token(authorization)
    return {
        "email": u["email"], "papel": u["papel"], "ativo": u["ativo"],
        "admin": u["papel"] == "admin" and u["ativo"],
        "credenciais_ativas": credenciais.chave_configurada(),
    }

@app.get("/credenciais")
def listar_credenciais(authorization: str = Header(default="")):
    _exigir_admin(authorization)
    atualizar_contas()
    return {"contas": [_publico(c) for c in CONTAS_TODAS]}

class NovaCredencial(BaseModel):
    usuario: str | None = None
    senha: str | None = None

@app.put("/credenciais/{conta_id}")
def salvar_credencial(conta_id: str, corpo: NovaCredencial, authorization: str = Header(default="")):
    email = _exigir_admin(authorization)
    atualizar_contas()
    atual = next((c for c in CONTAS_TODAS if c["id"] == conta_id), None)
    if atual is None:
        raise HTTPException(status_code=404, detail=f"Conta '{conta_id}' nao existe.")
    usuario = (corpo.usuario or "").strip()
    senha = corpo.senha or ""
    if not usuario and not senha:
        raise HTTPException(status_code=400, detail="Informe o usuario, a senha ou os dois.")
    if len(usuario) > 200 or len(senha) > 200:
        raise HTTPException(status_code=400, detail="Usuario ou senha longos demais.")
    linha = {"conta_id": conta_id, "nome": atual["nome"], "atualizado_por": email,
             "atualizado_em": datetime.now().astimezone().isoformat(timespec="seconds")}
    if usuario:
        linha["usuario"] = usuario
    if senha:
        linha["senha_cifrada"] = credenciais.cifrar(senha)
    try:
        supabase_admin.table(credenciais.TABELA).upsert(linha, on_conflict="conta_id").execute()
    except Exception as e:
        log(f"ERRO ao salvar credencial da conta {conta_id}: {type(e).__name__}")
        raise HTTPException(status_code=500, detail="Nao foi possivel salvar no banco. A tabela credenciais_coupa existe?")
    alterou = " e ".join(x for x, v in (("usuario", usuario), ("senha", senha)) if v)
    log(f"Credencial da conta {conta_id} atualizada ({alterou}) por {email}.")
    atualizar_contas()
    return {"ok": True, "conta": _publico(next(c for c in CONTAS_TODAS if c["id"] == conta_id))}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
