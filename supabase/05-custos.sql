-- ═══════════════════════════════════════════════════════════════════════════
--  Custos operacionais
--
--  Primeira tabela em que o dashboard ESCREVE. Até aqui ele só lia: a única
--  escrita permitida pela chave pública era INSERT em eventos, feito pelo quiz.
--  Aqui quem escreve é o usuário logado, e só quem está na lista de acesso.
-- ═══════════════════════════════════════════════════════════════════════════

create table if not exists public.custos (
  id        uuid primary key default gen_random_uuid(),
  nome      text not null check (length(btrim(nome)) between 1 and 120),
  valor     numeric(12,2) not null check (valor >= 0),
  -- data de COMPETÊNCIA: é por ela que o custo entra ou não no período filtrado
  data      date not null default (now() at time zone 'America/Sao_Paulo')::date,
  criado_em timestamptz not null default now(),
  criado_por text
);

create index if not exists custos_data_idx on public.custos (data);

alter table public.custos enable row level security;

/* tem_acesso() é SECURITY DEFINER (02-acesso.sql): consultar a lista de dentro
   de uma policy da própria tabela daria recursão. */
drop policy if exists custos_ler     on public.custos;
drop policy if exists custos_inserir on public.custos;
drop policy if exists custos_editar  on public.custos;
drop policy if exists custos_apagar  on public.custos;

create policy custos_ler     on public.custos for select using (public.tem_acesso());
create policy custos_inserir on public.custos for insert with check (public.tem_acesso());
create policy custos_editar  on public.custos for update using (public.tem_acesso())
                                                  with check (public.tem_acesso());
create policy custos_apagar  on public.custos for delete using (public.tem_acesso());

-- anon (a chave do quiz) não encosta nisto
revoke all on public.custos from anon;
grant select, insert, update, delete on public.custos to authenticated;
