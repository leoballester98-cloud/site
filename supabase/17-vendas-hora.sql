-- ═══════════════════════════════════════════════════════════════════════════
--  Vendas por dia da semana × hora
--
--  Sai da `vendas` (Kiwify) e não da criativos_hora (Meta) por dois motivos:
--  a Kiwify tem a hora EXATA do pagamento, e conta a venda real em vez da
--  atribuída pelo pixel. A criativos_hora também só guarda uma janela curta,
--  então não serviria para achar padrão de semana.
--
--  Tudo convertido pra São Paulo antes de extrair dia e hora. Sem isso o
--  timestamp em UTC jogaria as vendas da noite pro dia seguinte — e o padrão
--  de fim de semana, que é o que se procura aqui, sairia embaralhado.
-- ═══════════════════════════════════════════════════════════════════════════

create or replace function public.fn_vendas_hora(
  p_de  date default null,
  p_ate date default null
)
returns table (dia smallint, hora smallint, vendas bigint, receita numeric)
language sql
stable
security invoker
set search_path = public
as $$
  select
    extract(dow  from (v.ts at time zone 'America/Sao_Paulo'))::smallint as dia,   -- 0 = domingo
    extract(hour from (v.ts at time zone 'America/Sao_Paulo'))::smallint as hora,
    count(*)::bigint                                                     as vendas,
    coalesce(sum(v.valor), 0)                                            as receita
  from public.vendas v
  where v.evento = 'compra_aprovada'
    and (p_de  is null or (v.ts at time zone 'America/Sao_Paulo')::date >= p_de)
    and (p_ate is null or (v.ts at time zone 'America/Sao_Paulo')::date <= p_ate)
  group by 1, 2;
$$;

grant  execute on function public.fn_vendas_hora(date, date) to authenticated;
revoke execute on function public.fn_vendas_hora(date, date) from anon;
