// Convite de usuario do Seven OS.
// Unico lugar que usa a chave secreta para criar login. Fica no Supabase (que ja tem a chave),
// para que nem o os_api nem o robo precisem dela.
// Regras:
//  - quem chama precisa estar logado e ser administrador ou suporte ativo do Seven OS no schema informado
//    (conferido aqui dentro com o proprio token; por isso verify_jwt fica desligado no gateway);
//  - modo "link": cria o login e devolve o link de convite (o administrador envia por WhatsApp ou e-mail);
//  - modo "email": cria o login e o Supabase envia o e-mail de convite;
//  - se a pessoa ja tem login, so devolve o id (nada e enviado);
//  - com "reenviar": true e a pessoa ja tendo login, gera um link novo para ela (re)definir a senha
//    (convite expirado ou senha esquecida).
import { createClient } from "npm:@supabase/supabase-js@2.45.4";
import postgres from "npm:postgres@3.4.4";

const SCHEMAS = new Set(["os", "os_homolog"]);
const URL_SUPA = Deno.env.get("SUPABASE_URL")!;
const CHAVE = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
const sql = postgres(Deno.env.get("SUPABASE_DB_URL")!, { max: 1, prepare: false });
const admin = createClient(URL_SUPA, CHAVE, { auth: { persistSession: false, autoRefreshToken: false } });

// Link limpo, no endereco do proprio Seven OS: a pagina /definir-senha troca o codigo pelo login
// (verifyOtp). Sem passar pela pagina do Supabase e sem depender da lista de Redirect URLs.
// Sem redirect_to, cai no link padrao do Supabase.
type Props = { action_link: string; hashed_token: string; verification_type: string };
const linkDe = (p: Props, redirectTo?: string) =>
  redirectTo
    ? `${redirectTo}?convite=${p.hashed_token}${p.verification_type === "recovery" ? "&tipo=senha" : ""}`
    : p.action_link;

const resp = (status: number, corpo: unknown) =>
  new Response(JSON.stringify(corpo), { status, headers: { "Content-Type": "application/json" } });

Deno.serve(async (req) => {
  if (req.method !== "POST") return resp(405, { erro: "Use POST." });
  const token = (req.headers.get("Authorization") || "").replace(/^Bearer\s+/i, "");
  if (!token) return resp(401, { erro: "Login necessario." });
  const { data: quem, error: eQuem } = await admin.auth.getUser(token);
  if (eQuem || !quem?.user) return resp(401, { erro: "Login invalido ou expirado." });

  let corpo: { email?: string; schema?: string; modo?: string; redirect_to?: string; reenviar?: boolean; nome?: string };
  try { corpo = await req.json(); } catch { return resp(400, { erro: "Corpo invalido." }); }
  const email = (corpo.email || "").trim().toLowerCase();
  const schema = corpo.schema || "";
  const modo = corpo.modo === "email" ? "email" : "link";
  if (!SCHEMAS.has(schema)) return resp(400, { erro: "Schema invalido." });
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return resp(400, { erro: "E-mail invalido." });

  const ok = await sql`select 1 from ${sql(schema)}.usuarios
                       where user_id = ${quem.user.id} and ativo and perfil in ('administrador', 'suporte')`;  // os perfis com admin.usuarios (os_api/permissoes.py)
  if (!ok.length) return resp(403, { erro: "Somente administrador ou suporte do Seven OS pode convidar." });

  const redirectTo = corpo.redirect_to || undefined;
  const meta = corpo.nome ? { nome: String(corpo.nome).slice(0, 120) } : undefined; // so para a saudacao na tela de senha
  const ja = await sql`select id from auth.users where lower(email) = ${email} limit 1`;
  if (ja.length && !corpo.reenviar) return resp(200, { user_id: ja[0].id, ja_existia: true });
  if (ja.length) {
    if (modo === "email") {
      const { error } = await admin.auth.resetPasswordForEmail(email, { redirectTo });
      if (error) return resp(502, { erro: `O Supabase nao enviou o e-mail: ${error.message}` });
      return resp(200, { user_id: ja[0].id, ja_existia: true, modo });
    }
    const { data, error } = await admin.auth.admin.generateLink({ type: "recovery", email, options: { redirectTo } });
    if (error) return resp(502, { erro: `O Supabase nao gerou o link: ${error.message}` });
    return resp(200, { user_id: ja[0].id, ja_existia: true, modo, link: linkDe(data.properties, redirectTo) });
  }

  if (modo === "email") {
    const { data, error } = await admin.auth.admin.inviteUserByEmail(email, { redirectTo, data: meta });
    if (error) return resp(502, { erro: `O Supabase nao enviou o convite: ${error.message}` });
    return resp(201, { user_id: data.user.id, ja_existia: false, modo });
  }
  const { data, error } = await admin.auth.admin.generateLink({ type: "invite", email, options: { redirectTo, data: meta } });
  if (error) return resp(502, { erro: `O Supabase nao gerou o convite: ${error.message}` });
  return resp(201, { user_id: data.user.id, ja_existia: false, modo, link: linkDe(data.properties, redirectTo) });
});
