-- ═══════════════════════════════════════════════════════════════════════════
--  O líquido para de ser conta minha e passa a ser dado da Kiwify
--
--  A Kiwify já manda `my_commission` no webhook: quanto cai na conta. Eu vinha
--  calculando (8,99% + R$2,49) e o payload real mostrou dois furos:
--
--    · Venda parcelada. O cliente paga R$44,09, mas o produtor recebe os mesmos
--      R$32,00 — o juro vai pra operadora. Minha fórmula, aplicada sobre o
--      valor cobrado, devolvia R$37,75 e inflava o lucro.
--    · Order bump. Ele chega como LINHA SEPARADA de R$17,00, e a taxa dele é
--      R$1,53 (só o percentual). Minha fórmula cobrava os R$2,49 fixos de novo
--      e devolvia R$12,98 em vez de R$15,47.
--
--  Ler o número dela também sobrevive a mudança de taxa sem eu tocar em nada.
--
--  E `product_base_price` diz o preço do PRODUTO, sem juros — é ele que separa
--  os braços do teste de preço. Pelo valor cobrado, uma venda parcelada de
--  R$67,90 poderia cair no braço de R$37,90.
-- ═══════════════════════════════════════════════════════════════════════════

alter table public.vendas add column if not exists liquido    numeric(12,2);
alter table public.vendas add column if not exists preco_base numeric(12,2);

/* Preenche o que já existe a partir do payload cru. É por isso que ele foi
   guardado inteiro desde o primeiro webhook: nenhuma venda precisa ser
   reprocessada na Kiwify, e nada se perdeu por eu ter lido os campos errados. */
update public.vendas
   set liquido    = ((bruto->'payload'->'Commissions'->>'my_commission')::numeric) / 100,
       preco_base = ((bruto->'payload'->'Commissions'->>'product_base_price')::numeric) / 100
 where bruto->'payload'->'Commissions'->>'my_commission' is not null
   and (liquido is null or preco_base is null);


/* Qual produto é order bump. Fica em tabela porque é decisão de negócio, não
   de código: bump não é cliente novo, é venda a mais pra quem você já tinha —
   e por isso ele conta em Vendas (como a Kiwify conta) mas NÃO entra no
   divisor do CPA. */
alter table public.produto_braco add column if not exists tipo text not null default 'principal';
alter table public.produto_braco drop constraint if exists produto_braco_tipo_ok;
alter table public.produto_braco
  add constraint produto_braco_tipo_ok check (tipo in ('principal', 'bump'));

/* Semeia com os produtos que já venderam. Sai dos DADOS, então não depende de
   eu adivinhar id nenhum. O Clubinho entra como bump; o resto, principal. */
insert into public.produto_braco (produto_id, chave, marca, rotulo, tipo)
select distinct v.produto_id, 'preco', '', v.produto,
       case when v.produto ilike '%clubinho%' then 'bump' else 'principal' end
  from public.vendas v
 where v.produto_id is not null
on conflict (produto_id) do update set rotulo = excluded.rotulo;


-- ── Leitura por dia, agora com o líquido de verdade ────────────────────────
drop function if exists public.fn_vendas_dia(date, date);

create function public.fn_vendas_dia(
  p_de  date default null,
  p_ate date default null
)
returns table (
  data date, vendas bigint, vendas_sem_bump bigint, reembolsos bigint,
  faturamento numeric, liquido numeric
)
language sql
stable
security invoker
set search_path = public
as $$
  select
    (v.ts at time zone 'America/Sao_Paulo')::date                        as data,
    /* Vendas conta tudo, inclusive bump — é assim que a Kiwify conta, e dois
       números diferentes pra mesma coisa geram desconfiança nos dois. */
    count(*) filter (where v.evento = 'compra_aprovada')                 as vendas,
    count(*) filter (where v.evento = 'compra_aprovada'
                      and coalesce(pb.tipo, 'principal') <> 'bump')      as vendas_sem_bump,
    count(*) filter (where v.evento in ('reembolso','chargeback'))       as reembolsos,
    coalesce(sum(case when v.evento = 'compra_aprovada' then v.valor
                      when v.evento in ('reembolso','chargeback') then -v.valor
                      else 0 end), 0)                                    as faturamento,
    /* Reembolso tira o LÍQUIDO que tinha entrado. A Kiwify estorna a comissão
       junto, então o prejuízo é o que você recebeu — não o valor cheio, como
       eu tinha assumido antes de ver o payload. */
    coalesce(sum(case when v.evento = 'compra_aprovada' then coalesce(v.liquido, 0)
                      when v.evento in ('reembolso','chargeback') then -coalesce(v.liquido, 0)
                      else 0 end), 0)                                    as liquido
  from public.vendas v
  left join public.produto_braco pb on pb.produto_id = v.produto_id
  where (p_de  is null or (v.ts at time zone 'America/Sao_Paulo')::date >= p_de)
    and (p_ate is null or (v.ts at time zone 'America/Sao_Paulo')::date <= p_ate)
  group by 1
  order by 1;
$$;

grant  execute on function public.fn_vendas_dia(date, date) to authenticated;
revoke execute on function public.fn_vendas_dia(date, date) from anon;


-- ── Teste de preço: o braço sai do PREÇO BASE, não de tabela ───────────────
/* Sem mapa pra manter: 37,90 é um braço, 67,90 é o outro, e o número vem do
   próprio pedido. Produto novo no mesmo preço entra sozinho. */
drop function if exists public.fn_experimento_vendas(text, date, date);

create function public.fn_experimento_vendas(
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
  select
    case when v.preco_base >= 50 then 'preco-68' else 'preco-38' end     as marca,
    count(*) filter (where v.evento = 'compra_aprovada')                 as vendas,
    count(*) filter (where v.evento in ('reembolso','chargeback'))       as reembolsos,
    coalesce(sum(case when v.evento = 'compra_aprovada' then coalesce(v.liquido, 0)
                      when v.evento in ('reembolso','chargeback') then -coalesce(v.liquido, 0)
                      else 0 end), 0)                                    as receita
  from public.vendas v
  left join public.produto_braco pb on pb.produto_id = v.produto_id
  where p_exp = 'preco'
    and v.preco_base is not null
    /* Bump fora: ele não tem preço próprio no teste — acompanha a venda, e
       contá-lo como venda de R$17,00 embaralharia os dois braços. */
    and coalesce(pb.tipo, 'principal') <> 'bump'
    and (p_de  is null or (v.ts at time zone 'America/Sao_Paulo')::date >= p_de)
    and (p_ate is null or (v.ts at time zone 'America/Sao_Paulo')::date <= p_ate)
  group by 1;
$$;

grant  execute on function public.fn_experimento_vendas(text, date, date) to authenticated;
revoke execute on function public.fn_experimento_vendas(text, date, date) from anon;
