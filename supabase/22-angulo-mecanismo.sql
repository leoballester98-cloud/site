-- ═══════════════════════════════════════════════════════════════════════════
--  Ângulo do mecanismo — como o produto aparece no criativo
--
--  Vizinha de `angulo` e fácil de confundir com ela, então vale a distinção
--  por escrito: `angulo` é como quem fala se relaciona com quem assiste
--  (espelho, quem já saiu, assim como você só que pior). `angulo_mecanismo` é
--  COMO o produto entra na história: ela achou sozinha, alguém indicou, foi
--  feito por quem entende.
--
--  Ângulo do mecanismo forçado é o que faz um criativo soar como anúncio, e
--  por isso vale medir separado: dois vídeos com o mesmo ângulo podem ter
--  resultados opostos por causa disso.
-- ═══════════════════════════════════════════════════════════════════════════

alter table public.criativos add column if not exists angulo_mecanismo text;


-- ── fn_criativos precisa devolver o campo novo ─────────────────────────────
-- Criar a coluna na tabela não basta: o dashboard lê pela função, e ela lista
-- as colunas uma a uma. Sem isto a coluna aparece na tela mostrando '—' pra
-- todo mundo, e você preencheria a etiqueta sem nunca ver o valor de volta.
--
-- Esta é a função de 11-renomeia-campos.sql inteira, com o campo novo somado
-- em três lugares e nada mais tocado. drop antes do create porque mudar o
-- RETURNS TABLE muda a assinatura, e o Postgres recusa `create or replace`.

drop function if exists public.fn_criativos(date, date, text);

create function public.fn_criativos(
  p_de    date default null,
  p_ate   date default null,
  p_fonte text default null
)
returns table (
  anuncio text, campanha text, status text,
  formato text, formato_conteudo text, angulo text, angulo_mecanismo text,
  hook text, headline text,
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
    c.formato, c.formato_conteudo, c.angulo, c.angulo_mecanismo,
    c.hook, c.headline,
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
           c.angulo_mecanismo, c.hook, c.headline, c.emocao, c.amplificador, c.fatia, c.segmentacao, c.prova,
           c.vendas_ajuste, c.valor_ajuste;
$$;

grant  execute on function public.fn_criativos(date, date, text) to authenticated;
revoke execute on function public.fn_criativos(date, date, text) from anon;
