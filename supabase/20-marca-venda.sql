-- ═══════════════════════════════════════════════════════════════════════════
--  A venda passa a carregar de qual BRAÇO ela veio
--
--  O problema que isto resolve: o teste de preço separava os braços pelo preço
--  do produto — 37,90 num lado, 67,90 no outro. Só que as campanhas atuais
--  também vendem a R$37,90, e elas continuam rodando. Todas as vendas delas
--  cairiam no braço A, afogando o teste num volume que nunca viu o experimento.
--
--  Agora o quiz cola a marca do braço no campo `sck` do link de checkout, a
--  Kiwify devolve em TrackingParameters.sck, e a venda chega sabendo de onde
--  veio. Venda sem marca é venda que não participou de teste nenhum.
--
--  Serve pra qualquer experimento futuro: é só mudar o valor da marca.
-- ═══════════════════════════════════════════════════════════════════════════

alter table public.vendas add column if not exists marca text;

create index if not exists vendas_marca_idx on public.vendas (marca);

/* As 12 vendas que já existem vieram de links sem marca — elas são de campanha
   normal, e ficam de fora de qualquer teste. Correto: nenhuma delas viu o
   experimento de preço, que ainda nem subiu. */
update public.vendas
   set marca = nullif(btrim(bruto->'payload'->'TrackingParameters'->>'sck'), '')
 where marca is null;


/* Separa por MARCA, não por preço. É a única forma de o braço de R$37,90 do
   teste não se confundir com as campanhas que vendem no mesmo preço. */
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
    v.marca,
    count(*) filter (where v.evento = 'compra_aprovada')                 as vendas,
    count(*) filter (where v.evento in ('reembolso','chargeback'))       as reembolsos,
    coalesce(sum(case when v.evento = 'compra_aprovada' then coalesce(v.liquido, 0)
                      when v.evento in ('reembolso','chargeback') then -coalesce(v.liquido, 0)
                      else 0 end), 0)                                    as receita
  from public.vendas v
  join public.experimentos e on e.chave = p_exp
                            and v.marca in (e.braco_a, e.braco_b)
  left join public.produto_braco pb on pb.produto_id = v.produto_id
  /* Bump fora: ele acompanha uma venda que já foi contada, e sozinho não tem
     preço próprio no teste. */
  where coalesce(pb.tipo, 'principal') <> 'bump'
    and (p_de  is null or (v.ts at time zone 'America/Sao_Paulo')::date >= p_de)
    and (p_ate is null or (v.ts at time zone 'America/Sao_Paulo')::date <= p_ate)
  group by v.marca;
$$;

grant  execute on function public.fn_experimento_vendas(text, date, date) to authenticated;
revoke execute on function public.fn_experimento_vendas(text, date, date) from anon;
