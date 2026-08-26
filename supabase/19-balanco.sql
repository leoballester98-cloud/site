-- ═══════════════════════════════════════════════════════════════════════════
--  Balanço — como cada mês fechou, e onde o produto está no total
--
--  A pergunta que isto responde e que nenhuma outra aba responde: "estou no
--  lucro?". A Visão Geral mostra o período filtrado; aqui o mês é a unidade e o
--  acumulado atravessa todos eles.
--
--  Nada é digitado: sai do gasto do Meta, das vendas da Kiwify, dos custos
--  lançados e dos custos fixos. Fecha sozinho quando o mês vira.
-- ═══════════════════════════════════════════════════════════════════════════

/* A data em que o webhook da Kiwify entrou. Antes dela as vendas são as que o
   pixel do Meta atribuiu — a única coisa que existe pra aquele período.

   Fica em função, e não repetida em cada lugar que precisa dela, porque ela já
   vive no JavaScript do dashboard: duas cópias de uma data de corte é a receita
   pra uma ser mudada e a outra não, e aí os números param de bater sem que
   nada acuse. */
create or replace function public.kiwify_desde()
returns date language sql immutable as $$ select date '2026-08-25' $$;

grant execute on function public.kiwify_desde() to authenticated, anon;


create or replace function public.fn_balanco()
returns table (
  mes date, faturamento numeric, liquido numeric,
  anuncio numeric, operacional numeric, lucro numeric, vendas bigint
)
language sql
stable
security invoker
set search_path = public
as $$
  with
  /* Meta: o gasto sempre, e as vendas SÓ até o dia do corte. Depois dele quem
     conta venda é a Kiwify, e somar os dois contaria a mesma venda duas vezes. */
  meta as (
    select date_trunc('month', d.data)::date as mes,
           sum(d.gasto)                                                   as anuncio,
           sum(case when d.data < public.kiwify_desde()
                    then case when d.valor > 0 then d.valor else d.vendas * 37.90 end
                    else 0 end)                                           as faturamento,
           sum(case when d.data < public.kiwify_desde() then d.vendas else 0 end) as vendas,
           /* Antes do corte não existe my_commission. O líquido ali é estimado
              pela taxa do Celetus, que era R$33,50 por venda de R$37,90 — é o
              melhor que existe pra um período que não pode mais ser medido. */
           sum(case when d.data < public.kiwify_desde() then d.vendas * 33.50 else 0 end) as liquido
    from public.criativos_diario d
    group by 1
  ),
  kiwify as (
    select date_trunc('month', (v.ts at time zone 'America/Sao_Paulo')::date)::date as mes,
           sum(case when v.evento = 'compra_aprovada' then v.valor
                    when v.evento in ('reembolso','chargeback') then -v.valor else 0 end) as faturamento,
           sum(case when v.evento = 'compra_aprovada' then coalesce(v.liquido, 0)
                    when v.evento in ('reembolso','chargeback') then -coalesce(v.liquido, 0)
                    else 0 end)                                                           as liquido,
           count(*) filter (where v.evento = 'compra_aprovada')                           as vendas
    from public.vendas v
    where (v.ts at time zone 'America/Sao_Paulo')::date >= public.kiwify_desde()
    group by 1
  ),
  /* Custo avulso e ocorrência de custo fixo, pela data de competência. */
  op as (
    select date_trunc('month', c.data)::date as mes, sum(c.valor) as operacional
    from public.fn_custos(null, null) c
    group by 1
  ),
  meses as (
    select mes from meta union select mes from kiwify union select mes from op
  )
  select
    m.mes,
    coalesce(mt.faturamento, 0) + coalesce(kw.faturamento, 0)  as faturamento,
    coalesce(mt.liquido, 0)     + coalesce(kw.liquido, 0)      as liquido,
    coalesce(mt.anuncio, 0)                                     as anuncio,
    coalesce(op.operacional, 0)                                 as operacional,
    (coalesce(mt.liquido, 0) + coalesce(kw.liquido, 0))
      - coalesce(mt.anuncio, 0) - coalesce(op.operacional, 0)   as lucro,
    coalesce(mt.vendas, 0) + coalesce(kw.vendas, 0)             as vendas
  from meses m
  left join meta   mt on mt.mes = m.mes
  left join kiwify kw on kw.mes = m.mes
  left join op        on op.mes = m.mes
  order by m.mes desc;
$$;

grant  execute on function public.fn_balanco() to authenticated;
revoke execute on function public.fn_balanco() from anon;


-- ═══════════════════════════════════════════════════════════════════════════
--  Referências — números que você quer comparar depois
--
--  Separado do Histórico de propósito. Histórico é o que ACONTECEU: tem data e
--  passa. Referência é um número que continua servindo — "a conversão no
--  Celetus era 21,2%" vale daqui a seis meses, quando ninguém lembrar mais.
--
--  O rótulo é texto livre porque o contexto é seu: "sem VSL", "na h=2", "antes
--  do teste". Nenhuma lista de métricas guardaria isso.
--
--  `metrica` é opcional e liga a referência a um número que o dashboard sabe
--  calcular, pra ele mostrar o valor de hoje ao lado. Sem ela a referência
--  continua válida — só não compara sozinha. E é justamente o caso de "sem
--  VSL": uma situação que o dashboard não mede mais, e por isso vale guardar.
-- ═══════════════════════════════════════════════════════════════════════════

create table if not exists public.referencias (
  id        bigserial primary key,
  rotulo    text not null check (length(btrim(rotulo)) between 1 and 120),
  valor     numeric(14,4) not null,
  unidade   text not null default '%' check (unidade in ('%', 'R$', 'n')),
  metrica   text,          -- null = sem comparação automática
  data      date not null default (now() at time zone 'America/Sao_Paulo')::date,
  nota      text,
  criado_em timestamptz not null default now()
);

create index if not exists referencias_data_idx on public.referencias (data desc);

alter table public.referencias enable row level security;

drop policy if exists referencias_ler   on public.referencias;
drop policy if exists referencias_mexer on public.referencias;
create policy referencias_ler   on public.referencias for select using (public.tem_acesso());
create policy referencias_mexer on public.referencias for all
  using (public.tem_acesso()) with check (public.tem_acesso());

revoke all on public.referencias from anon;
grant select, insert, update, delete on public.referencias to authenticated;
