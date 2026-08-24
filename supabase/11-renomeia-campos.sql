-- ═══════════════════════════════════════════════════════════════════════════
--  Renomeia duas etiquetas que mudaram de SIGNIFICADO, não só de nome
--
--  arquetipo -> fatia
--    Era o papel de quem fala no vídeo (especialista, quem já passou).
--    Passa a ser a fatia de público a quem o vídeo fala, com a motivação dela:
--    "a jovem que achou que seria rápido e tem tudo, menos o filho".
--
--  estrutura -> formato_conteudo
--    "Estrutura" se confundia com a estrutura invisível do roteiro, que é outra
--    coisa. Este campo guarda o tipo de conteúdo: diário, desabafo, tutorial.
--
--  Os VALORES não são tocados. Os que já estão gravados foram preenchidos com o
--  significado antigo, então continuam errados até serem revisados um a um —
--  renomear a coluna não corrige o conteúdo dela.
-- ═══════════════════════════════════════════════════════════════════════════

alter table public.criativos rename column arquetipo to fatia;
alter table public.criativos rename column estrutura to formato_conteudo;


/* fn_criativos devolve as colunas por nome, então precisa ser reescrita junto —
   senão ela procura uma coluna que não existe mais. */
create or replace function public.fn_criativos(
  p_de    date default null,
  p_ate   date default null,
  p_fonte text default null
)
returns table (
  anuncio text, campanha text, status text,
  formato text, formato_conteudo text, angulo text, hook text, headline text,
  emocao text, amplificador text, fatia text, segmentacao text, prova text,
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
    c.formato, c.formato_conteudo, c.angulo, c.hook, c.headline,
    c.emocao, c.amplificador, c.fatia, c.segmentacao, c.prova,
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
  group by c.anuncio, c.campanha, c.status, c.formato, c.formato_conteudo, c.angulo,
           c.hook, c.headline, c.emocao, c.amplificador, c.fatia, c.segmentacao, c.prova,
           c.vendas_ajuste, c.valor_ajuste;
$$;

grant  execute on function public.fn_criativos(date, date, text) to authenticated;
revoke execute on function public.fn_criativos(date, date, text) from anon;
