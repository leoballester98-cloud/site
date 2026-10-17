-- ═══════════════════════════════════════════════════════════════════════════
--  Ciclo Fértil — as etapas do quiz viram TABELA
--  Rodar no SQL Editor do Supabase. Pode rodar de novo sem estragar nada.
--
--  Antes: o dashboard tinha os 34 nomes de tela escritos dentro do JavaScript,
--  numa lista posicional. O banco guardava só o número. Pra encaixar uma VSL no
--  meio do funil eu teria que editar o número na função SQL E o nome no dash —
--  dois lugares, na mão, uma vez por VSL. É assim que o nome erra e ninguém vê.
--
--  Agora: uma linha por etapa. O quiz grava `etapa`, o dashboard lê `nome`.
--  VSL nova = um insert.
-- ═══════════════════════════════════════════════════════════════════════════

-- ── 1. Espaço pra crescer ──────────────────────────────────────────────────
-- Era 1..40: 34 telas + a 35 (clique em comprar) deixavam 5 números livres.
-- Com 3 VSLs sobrariam 2. Sobe pra 99 e o assunto morre.
alter table public.eventos drop constraint if exists eventos_etapa_ok;
alter table public.eventos add  constraint eventos_etapa_ok
  check (etapa is null or (etapa between 1 and 99));


-- ── 2. A tabela ────────────────────────────────────────────────────────────
create table if not exists public.etapas (
  etapa smallint primary key,       -- o número que o quiz grava
  ordem numeric  not null,          -- onde aparece no funil (a VSL 1 é 13.5)
  nome  text     not null,
  tipo  text     not null default 'tela',
  -- Dia em que a etapa passou a existir. Null = sempre existiu.
  -- Serve pro dashboard não mentir: num período que começa ANTES da VSL
  -- entrar no ar, a linha dela apareceria quase zerada — não porque as
  -- pessoas saíram, mas porque a tela não existia. Com a data, ele mostra "—".
  desde date,
  constraint etapas_tipo_ok  check (tipo in ('tela','vsl','compra')),
  constraint etapas_ordem_ok check (ordem > 0 and ordem < 100)
);

create unique index if not exists etapas_ordem_idx on public.etapas (ordem);

alter table public.etapas enable row level security;
drop policy if exists "logado le etapas" on public.etapas;
create policy "logado le etapas" on public.etapas
  for select to authenticated using (true);
-- O quiz não lê nada daqui: ele já sabe em que tela está.


-- ── 3. O conteúdo ──────────────────────────────────────────────────────────
-- As 34 telas, na ordem em que estão no quiz hoje. Vieram do array NOMES que
-- vivia no dashboard — este arquivo passa a ser a única cópia.
insert into public.etapas (etapa, ordem, nome, tipo) values
  ( 1,  1, '(abertura/idade)',                                  'tela'),
  ( 2,  2, '(prova social)',                                    'tela'),
  ( 3,  3, 'Seu ciclo menstrual é…',                            'tela'),
  ( 4,  4, 'Qual é o seu maior desejo nesse momento?',          'tela'),
  ( 5,  5, 'Em qual área você sente algo travado?',             'tela'),
  ( 6,  6, 'Há quanto tempo está tentando engravidar?',         'tela'),
  ( 7,  7, 'Você nota mudanças no muco ao longo do mês?',       'tela'),
  ( 8,  8, 'Você já notou muco tipo clara de ovo?',             'tela'),
  ( 9,  9, 'Usa sabonete íntimo/ducha/lenços com frequência?',  'tela'),
  (10, 10, 'Já mediu a temperatura ao acordar?',                'tela'),
  (11, 11, 'Você sabe quando está ovulando?',                   'tela'),
  (12, 12, 'Sente dor de um lado do ventre no meio do ciclo?',  'tela'),
  (13, 13, 'Como você identifica seu período fértil?',          'tela'),
  (14, 14, 'Como descreveria seu estresse hoje?',               'tela'),
  (15, 15, 'Tem um momento do dia só pra você?',                'tela'),
  (16, 16, 'Fica ansiosa quando a menstruação vai chegar?',     'tela'),
  (17, 17, 'Intensidade da pressão de engravidar?',             'tela'),
  (18, 18, 'Seu humor muda bastante no dia?',                   'tela'),
  (19, 19, 'Como descreveria seu sono?',                        'tela'),
  (20, 20, 'Frequência de atividade física?',                   'tela'),
  (21, 21, 'Como descreveria sua alimentação?',                 'tela'),
  (22, 22, 'Sente inchaço/cansaço/dores sem causa?',            'tela'),
  (23, 23, 'Ainda acredita que vai conseguir engravidar?',      'tela'),
  (24, 24, 'Tem prática de meditação/oração/respiração?',       'tela'),
  (25, 25, 'Consegue ter momentos de alegria e leveza?',        'tela'),
  (26, 26, 'Você se culpa quando o ciclo não vem?',             'tela'),
  (27, 27, 'Quantos dias dura seu ciclo?',                      'tela'),
  (28, 28, 'Quantos dias dura sua menstruação?',                'tela'),
  (29, 29, 'Em quantos dias seu desejo sexual aumenta?',        'tela'),
  (30, 30, '(RESULTADO diagnóstico — início do pitch)',         'tela'),
  (31, 31, 'Quando começou sua última menstruação?',            'tela'),
  (32, 32, '(loading)',                                         'tela'),
  (33, 33, 'gráfico fascinations',                              'tela'),
  (34, 34, 'Página de vendas',                                  'tela'),
  (35, 99, 'Clique em comprar',                                 'compra')
on conflict (etapa) do update
  set ordem = excluded.ordem, nome = excluded.nome, tipo = excluded.tipo;

-- A VSL 1. Entra entre a 13 e a 14: quem passa da 13 cai nela, quem sai dela
-- cai na 14. Grava 36 no banco pra não empurrar nenhum número antigo —
-- as etapas 1..34 continuam significando exatamente o que sempre significaram.
insert into public.etapas (etapa, ordem, nome, tipo, desde) values
  (36, 13.5, '▶ VSL 1 · o mecanismo', 'vsl', '2026-09-16')
on conflict (etapa) do update
  set ordem = excluded.ordem, nome = excluded.nome,
      tipo = excluded.tipo, desde = excluded.desde;

-- Quando a VSL 2 e a 3 entrarem, é só isto aqui — nada de código:
--   insert into public.etapas values (37, 31.5, '▶ VSL 2 · a prova',  'vsl', 'AAAA-MM-DD');
--   insert into public.etapas values (38, 33.5, '▶ VSL 3 · a venda',  'vsl', 'AAAA-MM-DD');


-- ── 4. O funil, lendo da tabela ────────────────────────────────────────────
-- Muda o que sai: cada etapa agora vem com nome, tipo e desde, e a ordem é
-- `ordem` (numérica), não o número gravado. O dashboard parou de saber nomes.
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
  -- A última tela do quiz sai da tabela. Se um dia entrar uma tela 35 de
  -- verdade, "chegou ao fim" acompanha sozinho.
  ult as (select max(etapa) as n from public.etapas where tipo = 'tela'),
  sess as (
    -- maxi = a tela mais funda que a sessão alcançou. Só conta tipo 'tela':
    -- VSL e clique em comprar não são profundidade de funil.
    -- vistas = tudo que a sessão disparou, pra saber quem viu a VSL e quem
    -- clicou em comprar sem precisar de uma coluna por evento.
    select f.sessao,
           coalesce(max(t.etapa), 0)        as maxi,
           array_agg(distinct f.etapa)      as vistas
    from filtrado f
    left join public.etapas t on t.etapa = f.etapa and t.tipo = 'tela'
    group by f.sessao
  ),
  base as (
    -- `ult` entra por cross join (uma linha só) em vez de subquery dentro do
    -- FILTER: o Postgres não aceita subquery ali.
    select count(*) filter (where s.maxi >= 1)                  as visitantes,
           count(*) filter (where 35::smallint = any(s.vistas)) as compraram,
           count(*) filter (where s.maxi >= u.n)                as fim
    from sess s cross join ult u
  ),
  passos as (
    select et.etapa, et.ordem, et.nome, et.tipo, et.desde,
           case
             -- Tela: retenção acumulada. Quem chegou mais fundo passou por aqui.
             when et.tipo = 'tela'
               then (select count(*) from sess where maxi >= et.etapa)
             -- VSL: só quem disparou o ping dela. Não dá pra deduzir da tela
             -- seguinte — o tráfego de antes do lançamento chegou na 14 sem
             -- nunca ter visto VSL nenhuma.
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
