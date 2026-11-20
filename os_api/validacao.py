import re

PORTAIS = ("coupa",)  # novo portal entra aqui quando o robo souber captura-lo

_HORA = re.compile(r"([01]\d|2[0-3]):[0-5]\d")


def cnpj(valor):
    """Normaliza e valida CNPJ (14 digitos com digitos verificadores). Vazio vira None."""
    if valor is None or str(valor).strip() == "":
        return None
    d = re.sub(r"\D", "", str(valor))
    if len(d) != 14 or d == d[0] * 14:
        raise ValueError("CNPJ invalido.")
    def dv(base, pesos):
        s = sum(int(x) * p for x, p in zip(base, pesos))
        r = s % 11
        return "0" if r < 2 else str(11 - r)
    p1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    p2 = [6] + p1
    if d[12] != dv(d[:12], p1) or d[13] != dv(d[:13], p2):
        raise ValueError("CNPJ invalido.")
    return d


def _lista_texto(v, nome):
    if not isinstance(v, list) or not 1 <= len(v) <= 100:
        raise ValueError(f"{nome}: use uma lista com 1 a 100 itens.")
    itens = []
    for x in v:
        if not isinstance(x, str) or not x.strip() or len(x.strip()) > 80:
            raise ValueError(f"{nome}: cada item e um texto de 1 a 80 caracteres.")
        itens.append(x.strip())
    if len({i.lower() for i in itens}) != len(itens):
        raise ValueError(f"{nome}: itens repetidos.")
    return itens


def _agenda(v):
    if not isinstance(v, dict) or set(v) != {"horarios"} or not isinstance(v["horarios"], list):
        raise ValueError('robo.agenda: use {"horarios": ["07:00", "13:00"]}. Lista vazia desliga a agenda.')
    hs = v["horarios"]
    if len(hs) > 6:
        raise ValueError("robo.agenda: no maximo 6 rodadas por dia.")
    for h in hs:
        if not isinstance(h, str) or not _HORA.fullmatch(h):
            raise ValueError(f"robo.agenda: horario invalido {h!r}; use HH:MM.")
    if len(set(hs)) != len(hs):
        raise ValueError("robo.agenda: horarios repetidos.")
    return {"horarios": sorted(hs)}


def _percentual(v):
    if isinstance(v, bool) or not isinstance(v, (int, float)) or not 1 <= v <= 100:
        raise ValueError("alertas.divergencia_pct: numero entre 1 e 100.")
    return v


PARAMETROS = {
    "robo.agenda": _agenda,
    "alertas.divergencia_pct": _percentual,
    "listas.motivos_declinio": lambda v: _lista_texto(v, "listas.motivos_declinio"),
    "listas.categorias": lambda v: _lista_texto(v, "listas.categorias"),
}


def parametro(chave, valor):
    if chave not in PARAMETROS:
        raise KeyError(chave)
    return PARAMETROS[chave](valor)
