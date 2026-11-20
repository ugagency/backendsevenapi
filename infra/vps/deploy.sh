#!/bin/bash
# Copia fiel do deploy.sh em uso na VPS (/opt/seven-backend/deploy.sh), versionada em 20/11/2026.
# Roda manualmente: o workflow do GitHub so imprime "deploy disparado" e nao atualiza a VPS.
# Atencao: recria o container do robo e derruba uma captura em andamento.
set -e
cd /opt/seven-backend
git pull origin main
docker compose up -d --build
