"""Entrada do os_api na Vercel (funcao Python). So o os_api vai para la; o robo continua na VPS.
A Vercel chama o app ASGI abaixo para todos os caminhos (vercel.json)."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from os_api.main import criar_app  # noqa: E402

app = criar_app()
