"""Carga da base historica da Seven (BASE.xlsx) no historico_importado (Spec 03).

Cada linha da aba BASE vira uma linha de historico com origem 'base_seven', ligada ao produto:
- com CODITEM: pelo codigo Vale (produto_codigos); se o codigo nao existir ainda, cria o produto com a mesma
  assinatura da captura ('V|codigo'), para a proxima captura cair no mesmo produto;
- sem CODITEM: pela mesma cascata da captura (referencia, descricao estruturada, descricao), sem criar produto
  quando a descricao e insuficiente (fica so no historico, sem produto).
Fornecedores com preco de compra viram cadastro e vinculo de origem 'historia'. Pseudo-fornecedores (DECLINAR,
PROCURAR, INTERNET...) ficam so no texto. A aba COD E PV marca os itens ganhos com o numero do pedido de venda.
Recarregar o mesmo arquivo nao duplica: a chave da linha e referencia + codigo + posicao da referencia na planilha.
"""
import io
import re
import unicodedata
from datetime import date, datetime
from typing import Optional

import openpyxl
from psycopg.types.json import Jsonb

ORIGEM = "base_seven"
PSEUDO = {"", "0", "DECLINAR", "DECLINADO", "NECESSITA VALIDACAO", "PROCURAR", "INTERNET", "VERIFICAR", "SEM FORNECEDOR",
          "NAO ENCONTRADO", "NAO LOCALIZADO", "AGUARDANDO", "PESQUISAR", "N/A", "-"}
COLUNAS = {"VENDEDOR": "vendedor", "NUMERO DE REFERENCIA": "referencia", "NUMERO DO EVENTO": "evento", "CODITEM": "codigo",
           "UF VALE": "uf", "DATA": "data", "DESCRICAO DO CLIENTE": "descricao", "QUANTIDADE": "quantidade",
           "UNID MED": "unidade", "FORNECEDOR": "fornecedor", "PRECO DE COMPRA UNITARIO AJUSTADO": "preco_compra",
           "IPI %": "ipi", "FRETE GERAL %": "frete", "MARGEM FINAL": "margem", "PRECO VENDA S/DIFAL": "preco_venda",
           "STATUS": "status"}


def _sem_acento(t) -> str:
    t = unicodedata.normalize("NFKD", str(t or ""))
    return "".join(c for c in t if not unicodedata.combining(c)).strip().upper()


def _num(v) -> Optional[float]:
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip().replace("R$", "").replace("%", "").strip()
        if re.fullmatch(r"-?\d{1,3}(\.\d{3})*(,\d+)?", s):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return None
    return None


def _codigo(v) -> Optional[str]:
    if v is None:
        return None
    s = str(int(v)) if isinstance(v, float) and v.is_integer() else str(v)
    m = re.search(r"\d{6,10}", s)
    return m.group(0) if m else None


def _data(v) -> Optional[date]:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return None


def nomes_fornecedor(texto) -> list:
    """'FATTO/BAWEX' -> ['FATTO', 'BAWEX']; pseudo-fornecedores ficam de fora."""
    if re.search(r"https?:|www\.|\.com|@|\?", str(texto or ""), re.I):
        return []  # link de loja ou e-mail: compra na internet, nao e cadastro de fornecedor
    saida = []
    for parte in re.split(r"\s*/\s*|\s+E\s+|\s*\+\s*", str(texto or ""), flags=re.I):
        nome = re.sub(r"\s+", " ", parte).strip(" .-").upper()
        if nome and _sem_acento(nome) not in PSEUDO and not nome.isdigit() and len(nome) >= 2 and nome.count("-") < 3:
            saida.append(nome)
    return saida


def ler_planilha(conteudo: bytes) -> tuple:
    """Devolve (linhas, ganhos). linhas: dicts normalizados; ganhos: {chave_ref_codigo: 'PV n'}."""
    wb = openpyxl.load_workbook(io.BytesIO(conteudo), read_only=True, data_only=True)
    aba = next((w for w in wb.worksheets if _sem_acento(w.title) == "BASE"), wb.worksheets[0])
    it = aba.iter_rows(values_only=True)
    cab = [COLUNAS.get(_sem_acento(c)) for c in next(it)]
    if not {"referencia", "descricao"} <= set(cab):
        raise ValueError("A aba BASE nao tem as colunas esperadas (NUMERO DE REFERENCIA, DESCRICAO DO CLIENTE...).")
    linhas, vistos = [], {}
    for n, r in enumerate(it, start=2):
        d = {k: v for k, v in zip(cab, r) if k}
        if not any(v not in (None, "") for v in d.values()):
            continue
        ref = str(d.get("referencia") or "").strip().upper()
        cod = _codigo(d.get("codigo"))
        base = f"{ref}|{cod or ''}"
        vistos[base] = vistos.get(base, 0) + 1
        compra = _num(d.get("preco_compra"))
        venda = _num(d.get("preco_venda"))
        forn = re.sub(r"\s+", " ", str(d.get("fornecedor") or "")).strip()
        linhas.append({
            "chave": f"{base}|{vistos[base]}", "linha": n, "referencia": ref, "codigo": cod,
            "descricao": str(d.get("descricao") or "").strip(), "fornecedor": forn or None,
            "preco_compra": compra if compra and compra > 0 else None, "ipi": _num(d.get("ipi")), "frete": _num(d.get("frete")),
            "margem": _num(d.get("margem")), "preco_venda": venda if venda and venda > 0 else None,
            "declinou": _sem_acento(forn) in ("DECLINAR", "DECLINADO"), "status_planilha": str(d.get("status") or "").strip(),
            "data": _data(d.get("data")),
            "original": {k: (v.isoformat() if isinstance(v, (date, datetime)) else v) for k, v in d.items()},
        })
    ganhos = {}
    pv = next((w for w in wb.worksheets if "PV" in _sem_acento(w.title)), None)
    if pv is not None:
        for r in list(pv.iter_rows(values_only=True))[1:]:
            if not r or not r[0] or not r[1]:
                continue
            m = re.search(r"PV\s*(\d+)", str(r[1]).upper())
            if m and "GANHEI" in _sem_acento(r[1]):
                ganhos[re.sub(r"\s", "", str(r[0]).upper())] = m.group(1)
    por_ref = {}  # a aba PV junta referencia (SEV + 6 digitos) e codigo Vale, as vezes com espaco ou ordem trocada
    for k, pv_n in ganhos.items():
        m = re.search(r"SEV\d{6}", k)
        if m:
            por_ref.setdefault(m.group(0), []).append((_codigo(k.replace(m.group(0), "")), pv_n))
    n_ref = {}
    for li in linhas:
        n_ref[li["referencia"]] = n_ref.get(li["referencia"], 0) + 1
    for li in linhas:  # o STATUS da propria BASE tambem traz "GANHEI - PV n"
        m = re.search(r"GANHEI.*PV\s*(\d+)", _sem_acento(li["status_planilha"]))
        pv_n = m.group(1) if m else None
        for cod, pv in por_ref.get(li["referencia"], []):
            if pv_n:
                break
            if cod and li["codigo"] == cod:
                pv_n = pv
            elif not li["codigo"] and (cod or n_ref[li["referencia"]] == 1):
                pv_n = pv
                li["codigo"] = cod  # a aba PV traz o codigo que faltou na BASE
            elif not cod and n_ref[li["referencia"]] == 1:
                pv_n = pv
        li["status"] = "ganha" if pv_n else ("declinada" if li["declinou"] else ("cotada" if li["preco_compra"] else "sem_preco"))
        li["pedido_venda"] = pv_n or None
    return linhas, ganhos


SQL_PRODUTOS = r"""
with b as (
  select chave, codigo, descricao, desc_norm(descricao) dn,
         regexp_match(upper(coalesce(descricao, '')), 'FABRICANTE\s*/\s*REFERENCIA\s*:\s*([^/;|]+?)\s*/\s*([^;|]+)') fr
  from _carga),
c1 as (select *, (fr is not null and length(trim(fr[2])) >= 3 and trim(fr[2]) ~ '\d') tem_ref,
              (length(dn) >= 30 and length(dn) - length(replace(dn, ':', '')) >= 2) estruturada from b),
c as (
  select chave, codigo, descricao, dn,
    case when tem_ref then regexp_replace(trim(fr[1]), '\s+', ' ', 'g') end fab, case when tem_ref then trim(fr[2]) end ref,
    case when tem_ref then regexp_replace(trim(fr[1]), '\s+', ' ', 'g') || '|' || regexp_replace(trim(fr[2]), '[\s\-\./]', '', 'g') end ref_chave,
    case when tem_ref then 2 when estruturada then 1 else 0 end qualidade,
    case when codigo is not null then 'V|' || codigo
         when tem_ref then 'R|' || regexp_replace(trim(fr[1]), '\s+', ' ', 'g') || '|' || regexp_replace(trim(fr[2]), '[\s\-\./]', '', 'g')
         when estruturada then 'E|' || md5(dn) end assinatura,
    left(coalesce(nullif(trim(case when dn like 'PARTES E PECAS%%' and dn like '%%NOME DO ITEM:%%'
           then substring(dn from 'NOME DO ITEM:([^;]+)') else split_part(dn, ';', 1) end), ''),
         case when codigo is not null then 'Codigo ' || codigo else 'Sem descricao' end), 120) nome
  from c1)
update _carga k set assinatura = c.assinatura, nome = c.nome, fab = c.fab, ref = c.ref, ref_chave = c.ref_chave, qualidade = c.qualidade
from c where c.chave = k.chave;

-- 1) pelo codigo Vale ja conhecido
update _carga k set produto_id = coalesce(p.juntado_em, p.id)
from produto_codigos pc join produtos p on p.id = pc.produto_id where pc.codigo = k.codigo;
-- 2) pela assinatura (referencia ou descricao estruturada, ou V|codigo de produto criado antes)
update _carga k set produto_id = coalesce(p.juntado_em, p.id)
from produtos p where k.produto_id is null and k.assinatura is not null and p.assinatura = k.assinatura;
-- 3) cria o que falta (so com codigo, referencia ou descricao estruturada)
with novos as (
  insert into produtos (nome, descricao, fabricante, referencia, ref_chave, assinatura, prefixo, nivel, metodo)
  select distinct on (assinatura) nome, left(descricao, 2000), fab, ref, ref_chave, assinatura, left(codigo, 2),
         case when qualidade > 0 then 'certeza' else 'conferir' end,
         case qualidade when 2 then 'referencia' when 1 then 'estrutura' else 'codigo' end
  from _carga where produto_id is null and assinatura is not null
  order by assinatura, qualidade desc
  on conflict (assinatura) do nothing returning id, assinatura)
update _carga k set produto_id = n.id from novos n where k.produto_id is null and k.assinatura = n.assinatura;
insert into produto_codigos (produto_id, codigo, motivo)
select distinct produto_id, codigo, 'carga:base_seven' from _carga where produto_id is not null and codigo is not null
on conflict do nothing;
"""


def importar(conn, conteudo: bytes, arquivo: str, user_id: Optional[str] = None) -> dict:
    linhas, ganhos = ler_planilha(conteudo)
    if not linhas:
        raise ValueError("A planilha nao tem linhas.")
    conn.execute("""create temp table _carga (chave text primary key, codigo text, descricao text, assinatura text, nome text,
                    fab text, ref text, ref_chave text, qualidade int, produto_id bigint) on commit drop""")
    with conn.cursor() as cur:
        with cur.copy("copy _carga (chave, codigo, descricao) from stdin") as cp:
            for li in linhas:
                cp.write_row((li["chave"], li["codigo"], li["descricao"]))
    antes_prod = conn.execute("select count(*) from produtos").fetchone()[0]
    conn.execute(SQL_PRODUTOS)
    produto_de = dict(conn.execute("select chave, produto_id from _carga").fetchall())
    produtos_novos = conn.execute("select count(*) from produtos").fetchone()[0] - antes_prod

    # historico: substitui as linhas desta planilha que ja tinham sido carregadas (mesma chave)
    chaves = [li["chave"] for li in linhas]
    substituidas = conn.execute("delete from historico_importado where origem = %s and linha_original->>'_chave' = any(%s)",
                                (ORIGEM, chaves)).rowcount
    with conn.cursor() as cur:
        with cur.copy("""copy historico_importado (origem, arquivo, linha_original, produto_id, codigo, fornecedor_texto,
                         preco_compra, ipi, frete, margem, preco_venda, status, data_ref) from stdin""") as cp:
            for li in linhas:
                orig = dict(li["original"], _chave=li["chave"], _linha=li["linha"], _pedido_venda=li["pedido_venda"])
                cp.write_row((ORIGEM, arquivo, Jsonb(orig), produto_de.get(li["chave"]), li["codigo"], li["fornecedor"],
                              li["preco_compra"], li["ipi"], li["frete"], li["margem"], li["preco_venda"], li["status"], li["data"]))

    # fornecedores: so os que deram preco (cotacao real), com vinculo ao produto
    pares = {}
    for li in linhas:
        if li["preco_compra"] and produto_de.get(li["chave"]):
            for nome in nomes_fornecedor(li["fornecedor"]):
                pares.setdefault(nome, set()).add(produto_de[li["chave"]])
    antes_f = conn.execute("select count(*) from fornecedores").fetchone()[0]
    ids = {}
    for nome in pares:
        row = conn.execute("""select f.id from fornecedores f where upper(f.nome) = %s
                              union select fn.fornecedor_id from fornecedor_nomes fn where upper(fn.nome) = %s limit 1""", (nome, nome)).fetchone()
        if row:
            ids[nome] = row[0]
        else:
            ids[nome] = conn.execute("insert into fornecedores (nome, observacao) values (%s, %s) returning id",
                                     (nome, "Criado pela carga da base da Seven (" + arquivo + ")")).fetchone()[0]
    vinculos = 0
    for nome, prods in pares.items():
        for pid in prods:
            vinculos += conn.execute("""insert into produto_fornecedores (produto_id, fornecedor_id, origem, confirmado_por)
                                        values (%s, %s, 'historia', %s) on conflict do nothing""", (pid, ids[nome], user_id)).rowcount
    fornecedores_novos = conn.execute("select count(*) from fornecedores").fetchone()[0] - antes_f
    try:
        conn.execute("select atualizar_abas(null)")
    except Exception:
        pass

    datas = [li["data"] for li in linhas if li["data"]]
    return {
        "arquivo": arquivo, "linhas": len(linhas), "substituidas": substituidas,
        "com_produto": sum(1 for li in linhas if produto_de.get(li["chave"])),
        "sem_produto": sum(1 for li in linhas if not produto_de.get(li["chave"])),
        "produtos_novos": produtos_novos, "com_preco_compra": sum(1 for li in linhas if li["preco_compra"]),
        "declinadas": sum(1 for li in linhas if li["status"] == "declinada"),
        "ganhas": sum(1 for li in linhas if li["status"] == "ganha"), "pedidos_na_aba_pv": len(ganhos),
        "fornecedores_novos": fornecedores_novos, "vinculos_novos": vinculos,
        "periodo": [min(datas).isoformat(), max(datas).isoformat()] if datas else None,
    }


if __name__ == "__main__":  # uso local: python -m os_api.carga_base BASE.xlsx  (usa OS_DATABASE_URL / OS_SCHEMA)
    import json
    import sys
    from . import db
    from .config import Config
    db.iniciar(Config.do_ambiente())
    with open(sys.argv[1], "rb") as f, db.conexao() as conn:
        print(json.dumps(importar(conn, f.read(), sys.argv[1].split("/")[-1]), ensure_ascii=False, indent=2))
