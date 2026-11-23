#!/usr/bin/env bash
# Sobe o os_api (Seven OS) na VPS para a previa, ANTES do merge, sem tocar no robo.
#  - Codigo em pasta separada (/opt/seven-os-api, branch seven-os). /opt/seven-backend e o container do robo nao mudam.
#  - Container proprio (seven-os-api), porta 8100 so em 127.0.0.1.
#  - Nginx: acrescenta "location /os/" ao site seven-backend (copia de seguranca antes; se o teste falhar, volta sozinho).
#    O "location /" do robo nao muda. Reload do nginx nao derruba conexoes.
# Uso (como root):  bash subir_os_previa.sh
# Rodar de novo atualiza para o ultimo commit da branch seven-os.
set -euo pipefail

PASTA=/opt/seven-os-api
SITE=/etc/nginx/sites-available/seven-backend
PREVIA="${PREVIA:-https://seven-git-seven-os-ugagencys-projects.vercel.app}"

echo "== 1/4 Codigo (branch seven-os) em $PASTA"
if [ -d "$PASTA/.git" ]; then
  git -C "$PASTA" fetch -q origin seven-os && git -C "$PASTA" reset -q --hard origin/seven-os
else
  git clone -q -b seven-os https://github.com/ugagency/backendsevenapi.git "$PASTA"
fi
git -C "$PASTA" log --oneline -1

echo "== 2/4 Configuracao ($PASTA/.env.os)"
if [ ! -f "$PASTA/.env.os" ]; then
  cat > "$PASTA/.env.os" <<EOF
# Seven OS - previa. NUNCA colocar SUPABASE_SECRET_KEY aqui (o os_api se recusa a subir com ela).
OS_DATABASE_URL=postgresql://os_app.hwiyzclxtvggdhrxbsmc:TROQUE_PELA_SENHA_DO_OS_APP@aws-1-us-east-2.pooler.supabase.com:5432/postgres
OS_SCHEMA=os_homolog
SUPABASE_URL=https://hwiyzclxtvggdhrxbsmc.supabase.co
SUPABASE_KEY=sb_publishable_3NdA6yPrFDnlwFo4Nsksow_eowBw99Y
OS_CORS_ORIGINS=$PREVIA,http://localhost:3000
# cada versao da previa ganha um endereco proprio na Vercel
OS_CORS_REGEX=https://seven-[a-z0-9-]+-ugagencys-projects\.vercel\.app
OS_CONVITE_REDIRECT=$PREVIA/definir-senha
EOF
  chmod 600 "$PASTA/.env.os"
  echo
  echo ">> Criei $PASTA/.env.os. Falta a senha do os_app:"
  echo ">>   nano $PASTA/.env.os   (troque TROQUE_PELA_SENHA_DO_OS_APP; se a senha tiver @, escreva %40)"
  echo ">> Depois rode este script de novo."
  exit 0
fi
if grep -q TROQUE_PELA_SENHA "$PASTA/.env.os"; then
  echo ">> Falta a senha do os_app em $PASTA/.env.os (nano $PASTA/.env.os) e rodar de novo."; exit 1
fi

echo "== 3/4 Container seven-os-api"
cd "$PASTA"
docker build -q -f os_api/Dockerfile -t seven-os-api . >/dev/null
docker rm -f seven-os-api >/dev/null 2>&1 || true
docker run -d --name seven-os-api --restart unless-stopped --env-file .env.os -p 127.0.0.1:8100:8100 seven-os-api >/dev/null
for i in $(seq 1 20); do
  if curl -fsS http://127.0.0.1:8100/os/saude >/dev/null 2>&1; then break; fi; sleep 1
done
curl -fsS http://127.0.0.1:8100/os/saude && echo || { echo ">> os_api nao respondeu. Ultimas linhas:"; docker logs --tail 30 seven-os-api; exit 1; }

echo "== 4/4 Nginx (rota /os/)"
if grep -q "location /os/" "$SITE"; then
  echo "rota /os/ ja existe, nada a fazer"
else
  COPIA="$SITE.antes-os-$(date +%Y%m%d%H%M%S)"
  cp "$SITE" "$COPIA"
  python3 - "$SITE" <<'PY'
import sys
p = sys.argv[1]; s = open(p).read()
bloco = """    location /os/ {
        proxy_pass http://127.0.0.1:8100;
        proxy_http_version 1.1;
        proxy_set_header Host              $host;
        proxy_set_header X-Real-IP         $remote_addr;
        proxy_set_header X-Forwarded-For   $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto https;
        proxy_read_timeout 120;
    }

"""
alvo = "    location / {"
if s.count(alvo) != 1:
    sys.exit("nao achei um unico 'location / {' no site; nada alterado")
open(p, "w").write(s.replace(alvo, bloco + alvo, 1))
PY
  if nginx -t 2>/dev/null; then
    systemctl reload nginx && echo "nginx recarregado (copia anterior: $COPIA)"
  else
    cp "$COPIA" "$SITE"; echo ">> nginx -t falhou; arquivo restaurado, nada mudou."; nginx -t; exit 1
  fi
fi

echo
echo "Teste por fora:"
curl -fsS https://srv1002700.hstgr.cloud/os/saude && echo
curl -fsS https://srv1002700.hstgr.cloud/ && echo "   (robo continua respondendo)"
