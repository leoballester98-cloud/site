-- ═══════════════════════════════════════════════════════════════════════════
--  Separa as métricas por PLATAFORMA de anúncio (Meta, TikTok, …)
--
--  Só a criativos_diario ganha a coluna. A tabela `criativos` continua com o
--  nome como chave de propósito: o mesmo vídeo rodando nas duas plataformas é
--  o MESMO criativo, com as mesmas etiquetas (ângulo, hook, emoção). O que
--  muda entre plataformas é a métrica, e métrica mora no diário.
--
--  Tudo que já existe vira 'meta' pelo default, então nada precisa ser
--  reprocessado e o dashboard não muda de números.
-- ═══════════════════════════════════════════════════════════════════════════

alter table public.criativos_diario
  add column if not exists fonte text not null default 'meta';

/* A chave passa a incluir a fonte, senão um anúncio com o mesmo nome nas duas
   plataformas sobrescreveria o outro. O nome da constraint é descoberto em vez
   de escrito à mão: se ela tiver outro nome neste projeto, o script ainda
   funciona. */
do $$
declare nome text;
begin
  select conname into nome
  from pg_constraint
  where conrelid = 'public.criativos_diario'::regclass and contype = 'p';

  if nome is not null then
    execute format('alter table public.criativos_diario drop constraint %I', nome);
  end if;

  if not exists (
    select 1 from pg_constraint
    where conrelid = 'public.criativos_diario'::regclass and contype = 'p'
  ) then
    alter table public.criativos_diario add primary key (data, anuncio, fonte);
  end if;
end $$;

create index if not exists diario_fonte_idx on public.criativos_diario (fonte);


-- ── Leitura: as funções ganham filtro opcional de plataforma ────────────────

create or replace function public.fn_diario(
  p_de    date default null,
  p_ate   date default null,
  p_fonte text default null
)
returns table (data date, gasto numeric, faturamento numeric, vendas bigint, checkouts bigint)
language sql
stable
security invoker
set search_path = public
as $$
  select
    d.data,
    sum(d.gasto)                                                     as gasto,
    sum(case when d.valor > 0 then d.valor else d.vendas * 37.90 end) as faturamento,
    sum(d.vendas)                                                    as vendas,
    sum(d.checkouts)                                                 as checkouts
  from public.criativos_diario d
  where (p_de    is null or d.data >= p_de)
    and (p_ate   is null or d.data <= p_ate)
    and (p_fonte is null or d.fonte = p_fonte)
  group by d.data
  order by d.data;
$$;


/* Um total por plataforma. É o que alimenta os cartões da Visão Geral —
   antes o do TikTok era um espaço reservado com zeros escritos à mão. */
create or replace function public.fn_por_fonte(
  p_de  date default null,
  p_ate date default null
)
returns table (fonte text, gasto numeric, faturamento numeric,
               vendas bigint, checkouts bigint, ultima_venda date)
language sql
stable
security invoker
set search_path = public
as $$
  select
    d.fonte,
    sum(d.gasto)                                                     as gasto,
    sum(case when d.valor > 0 then d.valor else d.vendas * 37.90 end) as faturamento,
    sum(d.vendas)                                                    as vendas,
    sum(d.checkouts)                                                 as checkouts,
    max(d.data) filter (where d.vendas > 0)                          as ultima_venda
  from public.criativos_diario d
  where (p_de  is null or d.data >= p_de)
    and (p_ate is null or d.data <= p_ate)
  group by d.fonte
  order by sum(d.gasto) desc;
$$;


create or replace function public.fn_criativos(
  p_de    date default null,
  p_ate   date default null,
  p_fonte text default null
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
   and (p_de    is null or d.data >= p_de)
   and (p_ate   is null or d.data <= p_ate)
   and (p_fonte is null or d.fonte = p_fonte)
  where not exists (select 1 from public.consolidar k where k.anuncio = c.anuncio)
  group by c.anuncio, c.campanha, c.status, c.formato, c.estrutura, c.angulo,
           c.hook, c.emocao, c.amplificador, c.arquetipo, c.segmentacao, c.prova,
           c.vendas_ajuste, c.valor_ajuste;
$$;


/* As assinaturas antigas de 2 argumentos ficam órfãs depois do create or
   replace acima criar as de 3 — some com elas pra não haver duas funções com
   o mesmo nome e o Postgres não ter que adivinhar qual chamar. */
drop function if exists public.fn_diario(date, date);
drop function if exists public.fn_criativos(date, date);

grant execute on function public.fn_diario(date, date, text)    to authenticated;
grant execute on function public.fn_criativos(date, date, text) to authenticated;
grant execute on function public.fn_por_fonte(date, date)       to authenticated;

revoke execute on function public.fn_diario(date, date, text)    from anon;
revoke execute on function public.fn_criativos(date, date, text) from anon;
revoke execute on function public.fn_por_fonte(date, date)       from anon;
