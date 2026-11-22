"""Sobe o os_api no seu PC lendo o arquivo .env.os. Uso: python rodar_os_api.py"""
import os, sys
aqui = os.path.dirname(os.path.abspath(__file__))
for linha in open(os.path.join(aqui, ".env.os"), encoding="utf-8"):
    linha = linha.strip()
    if linha and not linha.startswith("#") and "=" in linha:
        k, v = linha.split("=", 1); os.environ[k.strip()] = v.strip()
sys.path.insert(0, aqui)
import uvicorn
print("Seven OS API em http://localhost:8100  (confira: http://localhost:8100/os/saude)")
uvicorn.run("os_api.main:app_factory", factory=True, host="127.0.0.1", port=8100)
