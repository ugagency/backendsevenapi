"""Teste da linha gravada em eventos_coletados (inclui data_inicio). Sem rede.
Rodar na raiz do backend:  python tests/test_evento_db.py"""
import os, sys
from datetime import date
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.update({"SUPABASE_URL": "https://x.supabase.co", "SUPABASE_KEY": "x" * 40, "VALE_USER": "u", "VALE_PASS": "p"})
os.environ.pop("SUPABASE_SECRET_KEY", None)
import supabase as _s
_s.create_client = lambda *a, **k: object()
import main_api as m

linha = ("303654", "13397589", "MA", "10/11/26", "INTERRUPTOR", "1", "PC-peca", None, "PLASSER")
e = m.montar_evento_db("filial2", linha, date(2026, 11, 3))
assert e["data_inicio"] == "2026-11-03", e                  # inicio = dia pesquisado, formato ISO (coluna date)
assert e["data_evento"] == "10/11/26", e                    # o prazo continua em data_evento
assert e["conta"] == "filial2" and e["numero_evento"] == "303654" and e["marca"] == "PLASSER"
assert set(e) == {"conta", "numero_evento", "titulo", "uf", "data_evento", "data_inicio", "descricao", "quantidade", "unidade", "marca"}
assert m.montar_evento_db("filial1", linha, None)["data_inicio"] is None
# data com ano de 4 digitos tambem vira date valida
assert m.parse_date_str("03/11/2026").isoformat() == m.parse_date_str("03/11/26").isoformat() == "2026-11-03"
print("TESTE OK: data_inicio gravada em ISO; prazo segue em data_evento")
