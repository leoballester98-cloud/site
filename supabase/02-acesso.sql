-- ═══════════════════════════════════════════════════════════════════════════
--  Quem pode abrir o dashboard
--
--  Antes: "qualquer usuário logado lê tudo". O problema é que criar conta é
--  fácil — bastaria o cadastro estar aberto e qualquer pessoa viraria "logado".
--  Agora a permissão é por e-mail, checada no banco. Mesmo com cadastro aberto,
--  quem não estiver nesta lista não lê nada.
-- ═══════════════════════════════════════════════════════════════════════════

create table if not exists public.acesso (
  email text primary key,
  nome  text
);

insert into public.acesso (email, nome) values
  ('leoballester98@gmail.com', 'Leo'),
  ('iordneto95@gmail.com',     'Iordan')
on conflict (email) do nothing;

alter table public.acesso enable row level security;
-- ninguém lê esta tabela pela API: ela só é consultada pela função abaixo
drop policy if exists "ninguem le acesso" on public.acesso;

/* security definer: a função enxerga a tabela mesmo com o RLS ligado.
   Sem isso a policy consultaria acesso, que tem RLS, que consultaria a policy…
   e o Postgres entraria em recursão. */
create or replace function public.tem_acesso()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select exists (
    select 1 from public.acesso
    where lower(email) = lower(coalesce(auth.jwt() ->> 'email', ''))
  );
$$;

revoke all on function public.tem_acesso() from public;
grant execute on function public.tem_acesso() to authenticated;


-- ── Reescreve as policies de leitura usando a lista ────────────────────────
drop policy if exists "logado le eventos" on public.eventos;
create policy "autorizado le eventos" on public.eventos
  for select to authenticated using (public.tem_acesso());

drop policy if exists "logado le criativos" on public.criativos;
create policy "autorizado le criativos" on public.criativos
  for select to authenticated using (public.tem_acesso());

drop policy if exists "logado edita criativos" on public.criativos;
create policy "autorizado edita criativos" on public.criativos
  for update to authenticated using (public.tem_acesso()) with check (public.tem_acesso());

drop policy if exists "logado le diario" on public.criativos_diario;
create policy "autorizado le diario" on public.criativos_diario
  for select to authenticated using (public.tem_acesso());

drop policy if exists "logado le paginas" on public.paginas;
create policy "autorizado le paginas" on public.paginas
  for select to authenticated using (public.tem_acesso());

drop policy if exists "logado le consolidar" on public.consolidar;
create policy "autorizado le consolidar" on public.consolidar
  for select to authenticated using (public.tem_acesso());

-- O insert do quiz (anônimo) continua como estava: escreve e não lê.
