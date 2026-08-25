-- ═══════════════════════════════════════════════════════════════════════════
--  Vendas vindas do webhook da Kiwify
--
--  Por que isto existe: as vendas que o dashboard já tem vêm do robô do Meta,
--  agregadas por ANÚNCIO. Elas não sabem qual preço a pessoa viu, nem qual
--  braço de experimento ela caiu — e sem isso um teste de preço não fecha,
--  porque a métrica que decide é receita por visitante.
--
--  O braço sai do próprio produto: os 9 links de R$67,90 são produtos
--  distintos na Kiwify, então o id que chega no webhook já diz o preço. Nada
--  precisa viajar do quiz até lá.
-- ═══════════════════════════════════════════════════════════════════════════

create table if not exists public.vendas (
  id         text primary key,          -- order_id da Kiwify: reentrega não duplica
  ts         timestamptz not null default now(),
  evento     text not null,             -- compra_aprovada | reembolso | chargeback | outro
  produto_id text,
  produto    text,
  valor      numeric(12,2) not null default 0,
  moeda      text,
  status     text,
  /* O payload inteiro fica guardado. Campo de webhook muda sem aviso, e quando
     mudar eu quero poder reprocessar do que chegou em vez de descobrir que o
     dado se perdeu na hora do parse. */
  bruto      jsonb
);

create index if not exists vendas_ts_idx      on public.vendas (ts);
create index if not exists vendas_produto_idx on public.vendas (produto_id);


/* De qual braço é cada produto. É tabela, e não código, pra você ligar um
   produto novo ao experimento sem deploy — e pra o mapa ficar visível. */
create table if not exists public.produto_braco (
  produto_id text primary key,
  chave      text not null,   -- o experimento (experimentos.chave)
  marca      text not null,   -- o braço (experimentos.braco_a / braco_b)
  rotulo     text
);


-- ── Quem pode o quê ────────────────────────────────────────────────────────
alter table public.vendas        enable row level security;
alter table public.produto_braco enable row level security;

drop policy if exists "logado le vendas" on public.vendas;
create policy "logado le vendas" on public.vendas
  for select to authenticated using (true);

drop policy if exists "logado le produto_braco" on public.produto_braco;
create policy "logado le produto_braco" on public.produto_braco
  for select to authenticated using (true);

drop policy if exists "logado edita produto_braco" on public.produto_braco;
create policy "logado edita produto_braco" on public.produto_braco
  for all to authenticated using (true) with check (true);

revoke all on public.vendas        from anon;
revoke all on public.produto_braco from anon;
-- quem escreve em vendas é a Edge Function, com a service role, que ignora RLS


/* Vendas e receita por braço do experimento, no mesmo recorte de data que o
   resto do dashboard. Reembolso e chargeback entram SUBTRAINDO: receita que
   voltou não é receita, e num teste de preço isso importa mais que o normal —
   preço mais alto costuma reembolsar mais, e é justo isso que precisa aparecer. */
create or replace function public.fn_experimento_vendas(
  p_exp text,
  p_de  date default null,
  p_ate date default null
)
returns table (marca text, vendas bigint, reembolsos bigint, receita numeric)
language sql
stable
security invoker
set search_path = public
as $$
  select pb.marca,
         count(*) filter (where v.evento = 'compra_aprovada')                as vendas,
         count(*) filter (where v.evento in ('reembolso','chargeback'))      as reembolsos,
         coalesce(sum(case when v.evento = 'compra_aprovada' then v.valor
                           when v.evento in ('reembolso','chargeback') then -v.valor
                           else 0 end), 0)                                   as receita
  from public.vendas v
  join public.produto_braco pb on pb.produto_id = v.produto_id
  where pb.chave = p_exp
    and (p_de  is null or (v.ts at time zone 'America/Sao_Paulo')::date >= p_de)
    and (p_ate is null or (v.ts at time zone 'America/Sao_Paulo')::date <= p_ate)
  group by pb.marca;
$$;

grant  execute on function public.fn_experimento_vendas(text, date, date) to authenticated;
revoke execute on function public.fn_experimento_vendas(text, date, date) from anon;
