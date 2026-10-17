-- ═══════════════════════════════════════════════════════════════════════════
--  Configurações do dashboard
--
--  Chave → valor JSON. Guarda regra que você muda de vez em quando e que os
--  dois precisam ver igual — a primeira é o critério das medalhas.
--  No navegador ficaria só na sua máquina, e o colaborador veria outra régua.
-- ═══════════════════════════════════════════════════════════════════════════

create table if not exists public.config (
  chave      text primary key,
  valor      jsonb not null,
  atualizado timestamptz not null default now()
);

alter table public.config enable row level security;

drop policy if exists config_ler    on public.config;
drop policy if exists config_gravar on public.config;

create policy config_ler    on public.config for select using (public.tem_acesso());
create policy config_gravar on public.config for all    using (public.tem_acesso())
                                                        with check (public.tem_acesso());

revoke all on public.config from anon;
grant select, insert, update, delete on public.config to authenticated;

-- critério atual, pra tabela não nascer vazia
insert into public.config (chave, valor) values ('medalhas', '[
  {"selo":"💎","nome":"Diamante","vendas":30,"roas":1.50},
  {"selo":"🥇","nome":"Ouro","vendas":50,"roas":1.25},
  {"selo":"🥈","nome":"Prata","vendas":30,"roas":1.25},
  {"selo":"🥉","nome":"Bronze","vendas":15,"roas":1.25}
]'::jsonb)
on conflict (chave) do nothing;
