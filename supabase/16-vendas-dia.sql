-- ═══════════════════════════════════════════════════════════════════════════
--  Vendas da Kiwify, por dia — a Visão Geral passa a ler daqui
--
--  Por que trocar a fonte: as vendas que a Visão Geral mostrava vinham do
--  robô do Meta, ou seja, do que o PIXEL conseguiu atribuir. O pixel sempre
--  conta menos que a realidade. A Kiwify conta o que foi pago.
--
--  O gasto continua vindo do Meta, porque só ele sabe quanto cobrou. E a aba
--  Criativos continua no Meta inteira: lá a pergunta é "qual anúncio trouxe a
--  venda", e a Kiwify não sabe responder isso — ela não vê o anúncio.
--
--  `por_valor` é o que mata a adivinhação de order bump. O dashboard deduzia
--  quantos bumps existiam a partir do faturamento total, assumindo um preço
--  único; com o teste de preço no ar, essa conta virou lixo. Aqui vem a
--  contagem de pedidos POR VALOR COBRADO, e o líquido sai exato.
-- ═══════════════════════════════════════════════════════════════════════════

create or replace function public.fn_vendas_dia(
  p_de  date default null,
  p_ate date default null
)
returns table (
  data date, vendas bigint, reembolsos bigint, faturamento numeric,
  reembolsado numeric, por_valor jsonb
)
language sql
stable
security invoker
set search_path = public
as $$
  with base as (
    select (v.ts at time zone 'America/Sao_Paulo')::date as dia, v.evento, v.valor
    from public.vendas v
    where (p_de  is null or (v.ts at time zone 'America/Sao_Paulo')::date >= p_de)
      and (p_ate is null or (v.ts at time zone 'America/Sao_Paulo')::date <= p_ate)
  )
  select
    b.dia,
    count(*) filter (where b.evento = 'compra_aprovada')            as vendas,
    count(*) filter (where b.evento in ('reembolso','chargeback'))  as reembolsos,
    /* Reembolso e chargeback entram SUBTRAINDO: dinheiro que voltou não é
       faturamento, e num teste de preço isso pesa mais que o normal — preço
       mais alto costuma reembolsar mais. */
    coalesce(sum(case when b.evento = 'compra_aprovada' then b.valor
                      when b.evento in ('reembolso','chargeback') then -b.valor
                      else 0 end), 0)                               as faturamento,
    /* Bruto devolvido no dia. O prejuízo de um reembolso é o valor CHEIO: a
       Kiwify não devolve a taxa dela, então o líquido perde mais do que
       recebeu. Subtrair só o líquido faria reembolso parecer barato. */
    coalesce(sum(b.valor) filter (where b.evento in ('reembolso','chargeback')), 0)
                                                                    as reembolsado,
    (select jsonb_object_agg(x.valor::text, x.n)
       from (select b2.valor, count(*) as n
               from base b2
              where b2.dia = b.dia and b2.evento = 'compra_aprovada'
              group by b2.valor) x)                                 as por_valor
  from base b
  group by b.dia
  order by b.dia;
$$;

grant  execute on function public.fn_vendas_dia(date, date) to authenticated;
revoke execute on function public.fn_vendas_dia(date, date) from anon;


/* Venda lançada à mão. Vai pra MESMA tabela do webhook, não pra outra: duas
   tabelas de venda seriam duas contagens pra reconciliar, que é exatamente o
   problema que estamos saindo. O prefixo no id evita colidir com order_id da
   Kiwify, e a marca no `bruto` deixa dar para separar depois o que foi digitado
   do que chegou sozinho. */
create or replace function public.fn_venda_manual(
  p_valor numeric,
  p_data  date,
  p_nota  text default null
)
returns text
language plpgsql
security invoker
set search_path = public
as $$
declare novo_id text;
begin
  if p_valor is null or p_valor <= 0 then
    raise exception 'valor precisa ser maior que zero';
  end if;

  novo_id := 'manual:' || gen_random_uuid()::text;

  insert into public.vendas (id, ts, evento, produto_id, produto, valor, moeda, status, bruto)
  values (
    novo_id,
    (p_data + time '12:00') at time zone 'America/Sao_Paulo',
    'compra_aprovada',
    null,
    coalesce(nullif(btrim(p_nota), ''), 'Lançamento manual'),
    p_valor,
    'BRL',
    'manual',
    jsonb_build_object('manual', true, 'lancado_em', now(), 'nota', p_nota)
  );
  return novo_id;
end;
$$;

grant  execute on function public.fn_venda_manual(numeric, date, text) to authenticated;
revoke execute on function public.fn_venda_manual(numeric, date, text) from anon;
