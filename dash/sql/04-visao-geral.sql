-- ═══════════════════════════════════════════════════════════════════════════
--  Série diária pra Visão Geral: gasto × faturamento por dia.
--  É a leitura que hoje não existe em nenhuma das duas abas — elas mostram
--  totais do período, não a evolução dentro dele.
-- ═══════════════════════════════════════════════════════════════════════════

create or replace function public.fn_diario(
  p_de  date default null,
  p_ate date default null
)
returns table (data date, gasto numeric, faturamento numeric, vendas bigint, checkouts bigint)
language sql
stable
security invoker
set search_path = public
as $$
  select
    d.data,
    sum(d.gasto)                                        as gasto,
    -- quando o pixel não devolve valor, cai no preço de tabela
    sum(case when d.valor > 0 then d.valor else d.vendas * 37.90 end) as faturamento,
    sum(d.vendas)                                       as vendas,
    sum(d.checkouts)                                    as checkouts
  from public.criativos_diario d
  where (p_de  is null or d.data >= p_de)
    and (p_ate is null or d.data <= p_ate)
  group by d.data
  order by d.data;
$$;

revoke execute on function public.fn_diario(date, date) from public;
grant  execute on function public.fn_diario(date, date) to authenticated;
