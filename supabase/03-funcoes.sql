-- ═══════════════════════════════════════════════════════════════════════════
--  Funções de leitura do dashboard
--
--  A conta acontece aqui, no Postgres, e o navegador só desenha. É o que faz
--  o dashboard abrir instantâneo: em vez de trazer todos os eventos e agregar
--  em JS (o que o Apps Script fazia), volta só o resultado.
--
--  security invoker de propósito: as funções respeitam o RLS de quem chamou.
--  Deslogado ou fora da lista de acesso, elas devolvem vazio.
-- ═══════════════════════════════════════════════════════════════════════════

-- ── Funil ──────────────────────────────────────────────────────────────────
-- Regra idêntica à de hoje: cada sessão conta pela MAIOR etapa que alcançou,
-- então a curva é sempre decrescente. Etapa 35 = clique em comprar, não é tela.
create or replace function public.fn_funil(
  p_de        date        default null,
  p_ate       date        default null,
  p_variantes text[]      default null,
  p_funil     text        default null,
  p_fonte     text        default null
)
returns jsonb
language sql
stable
security invoker
set search_path = public
as $$
  with filtrado as (
    select e.sessao, e.etapa
    from public.eventos e
    left join public.paginas pg on pg.codigo = e.pagina
    where e.tipo = 'etapa'
      and e.etapa is not null
      and (p_de  is null or (e.ts at time zone 'America/Sao_Paulo')::date >= p_de)
      and (p_ate is null or (e.ts at time zone 'America/Sao_Paulo')::date <= p_ate)
      and (p_variantes is null or e.variante = any (p_variantes))
      -- página vazia = ping anterior à separação: contava como Brasil/Meta
      and (p_funil is null or coalesce(pg.funil, 'brasil') = p_funil)
      and (p_fonte is null or coalesce(pg.fonte, 'meta')   = p_fonte)
  ),
  sess as (
    select sessao,
           max(case when etapa between 1 and 34 then etapa else 0 end) as maxi,
           bool_or(etapa = 35) as comprou
    from filtrado
    group by sessao
  ),
  base as (
    select count(*) filter (where maxi >= 1)  as visitantes,
           count(*) filter (where comprou)    as compraram,
           count(*) filter (where maxi >= 34) as fim
    from sess
  ),
  passos as (
    select n, (select count(*) from sess where maxi >= n) as sessoes
    from generate_series(1, 34) as n
  )
  select jsonb_build_object(
    'visitantes', (select visitantes from base),
    'compraram',  (select compraram  from base),
    'fim',        (select fim        from base),
    'etapas',     (select jsonb_agg(jsonb_build_object('n', n, 'sessoes', sessoes) order by n) from passos)
  );
$$;


-- ── Variantes de headline disponíveis no período ───────────────────────────
-- Alimenta o filtro. Sai dos dados, então variante nova aparece sozinha.
create or replace function public.fn_variantes(
  p_de    date default null,
  p_ate   date default null,
  p_funil text default null,
  p_fonte text default null
)
returns table (variante text, sessoes bigint, ini date, fim date)
language sql
stable
security invoker
set search_path = public
as $$
  select e.variante,
         count(distinct e.sessao) as sessoes,
         min((e.ts at time zone 'America/Sao_Paulo')::date) as ini,
         max((e.ts at time zone 'America/Sao_Paulo')::date) as fim
  from public.eventos e
  left join public.paginas pg on pg.codigo = e.pagina
  where e.tipo = 'etapa'
    and e.variante is not null and e.variante <> ''
    and (p_de  is null or (e.ts at time zone 'America/Sao_Paulo')::date >= p_de)
    and (p_ate is null or (e.ts at time zone 'America/Sao_Paulo')::date <= p_ate)
    and (p_funil is null or coalesce(pg.funil, 'brasil') = p_funil)
    and (p_fonte is null or coalesce(pg.fonte, 'meta')   = p_fonte)
  group by e.variante
  order by sessoes desc;
$$;


-- ── Criativos com as métricas do período ───────────────────────────────────
-- Junta as etiquetas (que você edita) com o somatório do diário no intervalo.
-- Os ajustes manuais de venda/valor entram aqui, como no dashboard atual.
create or replace function public.fn_criativos(
  p_de  date default null,
  p_ate date default null
)
returns table (
  anuncio text, campanha text, status text,
  formato text, estrutura text, angulo text, hook text, emocao text,
  amplificador text, arquetipo text, segmentacao text, prova text,
  gasto numeric, impressoes bigint, clicks bigint,
  v3s bigint, thruplay bigint, p25 bigint, lpv bigint,
  checkouts bigint, vendas bigint, valor numeric
)
language sql
stable
security invoker
set search_path = public
as $$
  select
    c.anuncio, c.campanha, c.status,
    c.formato, c.estrutura, c.angulo, c.hook, c.emocao,
    c.amplificador, c.arquetipo, c.segmentacao, c.prova,
    coalesce(sum(d.gasto), 0)                                  as gasto,
    coalesce(sum(d.impressoes), 0)                             as impressoes,
    coalesce(sum(d.clicks), 0)                                 as clicks,
    -- v2s é a reserva de quando a API não devolve o campo de 3s
    coalesce(sum(case when d.v3s > 0 then d.v3s else d.v2s end), 0) as v3s,
    coalesce(sum(d.thruplay), 0)                               as thruplay,
    coalesce(sum(d.p25), 0)                                    as p25,
    coalesce(sum(d.lpv), 0)                                    as lpv,
    coalesce(sum(d.checkouts), 0)                              as checkouts,
    coalesce(sum(d.vendas), 0) + c.vendas_ajuste               as vendas,
    coalesce(sum(d.valor), 0)  + c.valor_ajuste                as valor
  from public.criativos c
  left join public.criativos_diario d
    on d.anuncio = c.anuncio
   and (p_de  is null or d.data >= p_de)
   and (p_ate is null or d.data <= p_ate)
  -- linha absorvida por outro criativo some do dash inteiro
  where not exists (select 1 from public.consolidar k where k.anuncio = c.anuncio)
  group by c.anuncio, c.campanha, c.status, c.formato, c.estrutura, c.angulo,
           c.hook, c.emocao, c.amplificador, c.arquetipo, c.segmentacao, c.prova,
           c.vendas_ajuste, c.valor_ajuste;
$$;


grant execute on function public.fn_funil(date, date, text[], text, text) to authenticated;
grant execute on function public.fn_variantes(date, date, text, text)     to authenticated;
grant execute on function public.fn_criativos(date, date)                 to authenticated;

revoke execute on function public.fn_funil(date, date, text[], text, text) from anon;
revoke execute on function public.fn_variantes(date, date, text, text)     from anon;
revoke execute on function public.fn_criativos(date, date)                 from anon;
