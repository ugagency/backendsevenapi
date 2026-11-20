# Infra da VPS (Seven)

Estado verificado em 20/11/2026, por leitura na VPS (srv1002700, pasta `/opt/seven-backend`).

## Como o Seven roda hoje
- Repositorio clonado em `/opt/seven-backend` (origin: github.com/ugagency/backendsevenapi, https, sem token na URL).
- Container `seven-backend`, porta 8000 so em 127.0.0.1. Nginx (site `seven-backend`, dominio srv1002700.hstgr.cloud, certificado Let's Encrypt) faz proxy de tudo (`location /`) para ela.
- `docker-compose.yml` e `deploy.sh` existem so na VPS (untracked). Copias fieis estao nesta pasta.
- Deploy e manual: `./deploy.sh` faz `git pull origin main` e `docker compose up -d --build`.
  O workflow `.github/workflows/deploy.yml` entra por SSH mas so executa `echo "deploy disparado"`.
- Nao ha cron nem webhook de deploy do Seven. (Verificado: crontab do root e /etc/cron.d.)
- A VPS e compartilhada com outros projetos (ModoSegu, n8n, Evolution e outros). Nao reiniciar a maquina sem planejar.

## Armadilhas
- `deploy.sh` recria o container do robo e derruba uma captura em andamento. Rodar fora do horario e com /status sem execucao.
- Se `docker-compose.yml` ou `deploy.sh` forem commitados na raiz da main, o `git pull` da VPS falha
  ("untracked working tree files would be overwritten"). Antes disso, na VPS: `mv docker-compose.yml docker-compose.yml.vps`
  e `mv deploy.sh deploy.sh.vps`. Por isso as copias ficam em `infra/vps/`.

## No corte
1. Fora do horario da Karen, /status sem execucao.
2. Dump novo do banco.
3. Merge do backend na main.
4. Na VPS: `cd /opt/seven-backend && git pull origin main`, criar `.env.os`, copiar `infra/vps/docker-compose.os.yml`
   para `docker-compose.yml`, `docker compose up -d --build`.
5. Nginx: aplicar `nginx-seven-backend.os.conf` (acrescenta `location /os/` para 127.0.0.1:8100; o `location /` do robo nao muda), depois `nginx -t && systemctl reload nginx`. Fazer copia do arquivo atual antes.
6. Teste rapido; se falhar, volta por `git revert` e `docker compose up -d --build`.

## Pontos de atencao
- O front fala com o backend pelo hostname da Hostinger (srv1002700.hstgr.cloud). Um dominio proprio da Seven ou da SSYS seria melhor para o produto; nao e urgente.
- CORS do robo esta aberto (`*`). Fecha no fechamento de seguranca.
