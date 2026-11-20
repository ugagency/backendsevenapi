"""
Captura da descrição do item na linha expandida de um evento do Coupa.

Compartilhado por main_api.py e sevenmaisdeum.py (não duplicar a lógica nos dois).

O Coupa renderiza os campos da linha em layouts diferentes conforme o evento:
  - layout antigo:  <dl class="customFieldsReview"><dt>Rótulo</dt><dd><span>valor</span></dd>
  - layout novo:    <div class="field s-showField"><div class="s-fieldLabel">Rótulo</div>
                    <p class="s-textField">valor</p></div>  (dentro de .s-request_line_custom_fields)
  - campo padrão:   .s-extended_description p.s-textField  ("Descrição")

Estratégia: lê os dois layouts, aceita qualquer rótulo dentro da área de campos
customizados da linha, procura em toda a linha campos cujo rótulo indique texto ou
descrição e, se nada for encontrado, usa o texto do título como último recurso.
"""

import re

from selenium.webdriver.common.by import By

# Ordem em que os campos conhecidos aparecem na planilha; os demais vão ao final.
CAMPOS_DESCRICAO = [
    "Texto do item",
    "Texto de compra do material",
    "Texto Breve",
    "Descrição",
]

# Valores que o Coupa usa para campo vazio.
_VALORES_VAZIOS = {"", "nenhum", "none", "n/a", "-", "--"}

# Rótulos fora da área de campos customizados que também contam como descrição.
_ROTULO_DESCRITIVO = re.compile(r"texto|descri|especifica|observa", re.IGNORECASE)

# Rótulos que nunca são descrição, mesmo que apareçam na área de campos customizados.
_ROTULOS_IGNORADOS = {
    "detalhes da solicitação", "detalhes", "anexos", "ncm", "necessário até",
    "endereço de entrega", "incoterms", "quantidade esperada", "nome do fabricante",
    "número da peça do fabricante",
}

_CLASSES_VAZIAS = {"none", "placeholderText", "s-noResponses"}


def _limpar_multilingue(txt):
    """Extrai apenas o bloco PT quando o Coupa concatena PT/EN/ES."""
    m = re.search(r"PT\s*\|\|\s*(.*?)\*{3,}", txt, re.DOTALL)
    return m.group(1).strip() if m else txt.strip()


def _texto(el):
    """Texto visível; se vier vazio (elemento oculto), usa textContent."""
    try:
        txt = (el.text or "").strip()
        if not txt:
            txt = (el.get_attribute("textContent") or "").strip()
        return txt
    except Exception:
        return ""


def _classes(el):
    try:
        return set((el.get_attribute("class") or "").split())
    except Exception:
        return set()


def _valor_util(el):
    """Retorna o texto do elemento de valor, ou "" se for marcador de campo vazio."""
    if el is None or _classes(el) & _CLASSES_VAZIAS:
        return ""
    txt = _texto(el)
    return "" if txt.lower() in _VALORES_VAZIOS else txt


def _valor_dd(dd_el):
    """Layout antigo: campo vazio = <span class="none">."""
    try:
        span = dd_el.find_element(By.TAG_NAME, "span")
        return _valor_util(span)
    except Exception:
        return _valor_util(dd_el)


def _registrar(encontrados, rotulo, valor):
    rotulo = (rotulo or "").strip().rstrip(":").strip()
    if not rotulo or not valor or rotulo in encontrados:
        return
    if rotulo.lower() in _ROTULOS_IGNORADOS:
        return
    encontrados[rotulo] = _limpar_multilingue(valor)


def _campos_layout_antigo(escopo, encontrados):
    try:
        for dt in escopo.find_elements(By.CSS_SELECTOR, "dl.customFieldsReview dt"):
            try:
                dd = dt.find_element(By.XPATH, "following-sibling::dd[1]")
            except Exception:
                continue
            _registrar(encontrados, _texto(dt), _valor_dd(dd))
    except Exception:
        pass


def _campos_layout_novo(escopo, encontrados, seletor_campos, so_descritivos):
    try:
        campos = escopo.find_elements(By.CSS_SELECTOR, seletor_campos)
    except Exception:
        return
    for campo in campos:
        try:
            rotulo = _texto(campo.find_element(By.CSS_SELECTOR, ".s-fieldLabel"))
        except Exception:
            continue
        if so_descritivos and not _ROTULO_DESCRITIVO.search(rotulo or ""):
            continue
        try:
            valor_el = campo.find_element(By.CSS_SELECTOR, "p.s-textField")
        except Exception:
            continue
        _registrar(encontrados, rotulo, _valor_util(valor_el))


def coletar_descricoes(escopo):
    """
    Recebe a linha expandida (.s-itemsAndServicesLine.-expanded) e retorna
    (texto_concatenado, lista_de_campos_usados). Texto vazio se nada for encontrado.
    """
    encontrados = {}

    # 1) Campos customizados, layout antigo (dl/dt/dd)
    _campos_layout_antigo(escopo, encontrados)

    # 2) Campos customizados, layout novo: qualquer rótulo dentro da área de campos da linha
    _campos_layout_novo(escopo, encontrados,
                        ".s-request_line_custom_fields .s-showField", so_descritivos=False)

    # 3) Campo padrão "Descrição"
    try:
        el = escopo.find_element(By.CSS_SELECTOR, ".s-extended_description p.s-textField")
        _registrar(encontrados, "Descrição", _valor_util(el))
    except Exception:
        pass

    # 4) Varredura geral da linha: qualquer campo cujo rótulo indique texto/descrição
    _campos_layout_novo(escopo, encontrados, ".s-showField", so_descritivos=True)

    # 5) Monta na ordem definida, descartando textos repetidos
    partes, ordem_log, vistos = [], [], set()
    ordem = [n for n in CAMPOS_DESCRICAO if n in encontrados]
    ordem += [n for n in encontrados if n not in CAMPOS_DESCRICAO]
    for nome in ordem:
        val = encontrados[nome]
        chave = " ".join(val.split()).lower()
        if not chave or chave in vistos:
            continue
        vistos.add(chave)
        partes.append(val)
        ordem_log.append(nome)

    # 6) Último recurso: texto do título ("35007833 || OUTRAS OBRAS FERRO OU ACO")
    if not partes:
        try:
            tit = _texto(escopo.find_element(By.CSS_SELECTOR, ".s-description p.s-textField"))
            resto = tit.split("||", 1)[1].strip() if "||" in tit else ""
            if resto:
                partes.append(resto)
                ordem_log.append("Titulo (fallback)")
        except Exception:
            pass

    return " | ".join(partes), ordem_log
