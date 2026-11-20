-- Seven OS, Spec 01. Aplicada em 20/11/2026 (migration os_spec01_auth_id_por_email) em os e os_homolog.
-- Excecao controlada: o os_app nao le auth.users; esta funcao devolve so o id de um e-mail exato.
do $$
declare s text;
begin
  foreach s in array array['os','os_homolog'] loop
    execute format(
      'create or replace function %I.auth_id_por_email(p_email text) returns uuid language sql stable security definer set search_path = '''' as %L',
      s, 'select id from auth.users where lower(email) = lower(trim(p_email)) limit 1');
    execute format('revoke all on function %I.auth_id_por_email(text) from public, anon, authenticated', s);
    execute format('grant execute on function %I.auth_id_por_email(text) to os_app', s);
  end loop;
end $$;
