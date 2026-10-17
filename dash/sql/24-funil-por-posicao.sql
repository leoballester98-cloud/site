-- ═══════════════════════════════════════════════════════════════════════════
--  Funil contado pela POSIÇÃO no caminho, não pelo número da etapa (14/10/2026)
--
--  Antes: tela contava por maxi >= etapa (o número gravado). Com telas que
--  gravam números fora de ordem (37, 4, 41, 3…), isso obrigava a marcar metade
--  do funil como 'vsl' (conta só quem viu), e a curva subia e descia.
--  Agora maxi é a maior ORDEM (posição) que a sessão alcançou entre as telas;
--  quem chegou na posição 20 passou pelas 19 anteriores.
--
--  Tipos: 'tela' = caminho principal (conta por posição); 'vsl' = vídeo, só
--  quem viu; 'ramo' = um dos dois lados de uma bifurcação (ainda não / já
--  tenta), só quem viu — por posição contaria quem passou pelo outro lado.
-- ═══════════════════════════════════════════════════════════════════════════
begin;

alter table public.etapas drop constraint if exists etapas_tipo_ok;
alter table public.etapas add constraint etapas_tipo_ok
  check (tipo in ('tela', 'vsl', 'ramo', 'compra'));

update public.etapas set tipo = 'tela'
 where etapa in (37, 4, 41, 43, 9, 42, 50, 51, 44, 45, 52, 32, 53, 49, 47, 48, 46);
update public.etapas set tipo = 'ramo' where etapa in (38, 39);

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
      and (p_funil is null or coalesce(pg.funil, 'brasil') = p_funil)
      and (p_fonte is null or coalesce(pg.fonte, 'meta')   = p_fonte)
  ),
  -- A última POSIÇÃO de tela do caminho: quem chegou nela chegou ao fim.
  ult as (select max(ordem) as n from public.etapas where tipo = 'tela'),
  sess as (
    select f.sessao,
           coalesce(max(t.ordem), 0)        as maxi,
           array_agg(distinct f.etapa)      as vistas
    from filtrado f
    left join public.etapas t on t.etapa = f.etapa and t.tipo = 'tela'
    group by f.sessao
  ),
  base as (
    select count(*) filter (where s.maxi >= 1)                  as visitantes,
           count(*) filter (where 35::smallint = any(s.vistas)) as compraram,
           count(*) filter (where s.maxi >= u.n)                as fim
    from sess s cross join ult u
  ),
  passos as (
    select et.etapa, et.ordem, et.nome, et.tipo, et.desde,
           case
             when et.tipo = 'tela'
               then (select count(*) from sess where maxi >= et.ordem)
             else (select count(*) from sess where et.etapa = any(vistas))
           end as sessoes
    from public.etapas et
    where et.tipo <> 'compra'
  )
  select jsonb_build_object(
    'visitantes', (select visitantes from base),
    'compraram',  (select compraram  from base),
    'fim',        (select fim        from base),
    'telas',      (select count(*) from public.etapas where tipo = 'tela'),
    'etapas',     (select jsonb_agg(jsonb_build_object(
                     'n',       ordem,
                     'etapa',   etapa,
                     'nome',    nome,
                     'tipo',    tipo,
                     'desde',   desde,
                     'sessoes', sessoes
                   ) order by ordem) from passos)
  );
$$;

commit;
notify pgrst, 'reload schema';

select etapa, ordem, tipo, nome from public.etapas order by ordem;
