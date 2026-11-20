#!/bin/bash
# Monta um Postgres local de teste com o minimo do Supabase, aplica TODAS as migrations e carrega dados sinteticos.
# Uso: tests_os/local/preparar.sh [arquivo.csv de eventos_coletados, opcional]
# Depois: OS_TEST_DATABASE_URL="postgresql://os_app@/seven_teste?host=/tmp&port=5544" python -m pytest tests_os -q
set -e
cd "$(dirname "$0")/../.."
PORTA=${PGPORT_TESTE:-5544}; BANCO=${BANCO_TESTE:-seven_teste}
P="psql -h /tmp -p $PORTA -U postgres -v ON_ERROR_STOP=1 -q"
$P -c "drop database if exists $BANCO" -c "create database $BANCO" 2>/dev/null
P="$P -d $BANCO"
$P -f tests_os/local/stub_supabase.sql
$P -f migrations/000_base.sql
for f in migrations/os/0*.sql; do $P -f "$f" 2>&1 | grep -v NOTICE || true; done
if [ -n "$1" ]; then
  $P -c "\copy public.eventos_coletados from '$1' csv header" -c "select setval('public.eventos_coletados_id_seq', (select max(id) from public.eventos_coletados))" > /dev/null
fi
$P -f tests_os/local/seed.sql > /dev/null
echo "banco $BANCO pronto na porta $PORTA"
