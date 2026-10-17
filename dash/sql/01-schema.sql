-- ═══════════════════════════════════════════════════════════════════════════
--  Ciclo Fértil — schema do dashboard
--  Rodar UMA VEZ no SQL Editor do Supabase (projeto ciclo-fertil).
--  Pode rodar de novo sem estragar nada: tudo é "if not exists" / "or replace".
-- ═══════════════════════════════════════════════════════════════════════════

-- ── 1. De qual página veio o ping ──────────────────────────────────────────
-- Cada página do quiz é um par funil + fonte. É tabela, e não código, pra
-- página nova entrar sem eu mexer no dashboard.
create table if not exists public.paginas (
  codigo text primary key,
  funil   text not null,          -- 'brasil' | 'latam'
  fonte   text not null,          -- 'meta'   | 'tiktok'
  rotulo  text not null
);

insert into public.paginas (codigo, funil, fonte, rotulo) values
  ('v2',       'brasil', 'meta',   'Quiz BR — Meta'),
  ('tiktok',   'brasil', 'tiktok', 'Quiz BR — TikTok'),
  ('espanhol', 'latam',  'meta',   'Quiz Latam — Meta')
on conflict (codigo) do nothing;


-- ── 2. Pings do quiz ───────────────────────────────────────────────────────
create table if not exists public.eventos (
  id       bigserial primary key,
  ts       timestamptz not null default now(),
  sessao   text not null,
  tipo     text not null default 'etapa',
  etapa    smallint,                     -- 1..34 = tela, 35 = clique em comprar
  variante text,                         -- v0 / v1 / voriginal (teste de headline)
  pagina   text,                         -- vazio = registro anterior à separação
  dados    jsonb,                        -- respostas do quiz

  -- limites pra um insert malformado (ou abusivo) não sujar a base:
  -- o endpoint de insert é público por necessidade, o quiz é anônimo.
  constraint eventos_tipo_ok   check (tipo in ('etapa','respostas')),
  constraint eventos_etapa_ok  check (etapa is null or (etapa between 1 and 40)),
  constraint eventos_sessao_ok check (char_length(sessao) between 4 and 64),
  constraint eventos_var_ok    check (variante is null or char_length(variante) <= 24),
  constraint eventos_pag_ok    check (pagina   is null or char_length(pagina)   <= 24)
);

-- o funil sempre filtra por data e agrupa por sessão
create index if not exists eventos_ts_idx        on public.eventos (ts);
create index if not exists eventos_sessao_idx    on public.eventos (sessao);
create index if not exists eventos_ts_pagina_idx on public.eventos (ts, pagina);


-- ── 3. Criativos (o que hoje é a aba Criativos da planilha) ────────────────
-- As etiquetas são editadas por você; as métricas vêm do Meta.
create table if not exists public.criativos (
  anuncio       text primary key,        -- nome exato do anúncio no Meta
  ad_id         text,
  campanha      text,
  status        text,                    -- rodando / testado / o que você escrever
  formato       text,
  estrutura     text,
  angulo        text,
  hook          text,
  emocao        text,
  amplificador  text,
  arquetipo     text,
  segmentacao   text,
  prova         text,
  vendas_ajuste integer not null default 0,   -- venda que o Meta não marcou
  valor_ajuste  numeric not null default 0,
  atualizado    timestamptz
);

-- anúncios que são o MESMO criativo: as métricas somam na linha canônica
create table if not exists public.consolidar (
  anuncio   text primary key,            -- nome que aparece no Meta
  canonico  text not null                -- linha que recebe as métricas
);

insert into public.consolidar (anuncio, canonico) values
  ('tela original', 'ad1 - desabafo UCG'),
  ('tela 1',        'ad1 - desabafo UCG')
on conflict (anuncio) do nothing;


-- ── 4. Métricas do Meta, um dia por anúncio ────────────────────────────────
create table if not exists public.criativos_diario (
  data       date not null,
  anuncio    text not null,
  ad_id      text,
  campanha   text,
  gasto      numeric not null default 0,
  impressoes bigint  not null default 0,
  clicks     bigint  not null default 0,
  v3s        bigint  not null default 0,   -- plays de 3s (base do hook rate)
  v2s        bigint  not null default 0,   -- reserva quando v3s não vem da API
  thruplay   bigint  not null default 0,
  p25        bigint  not null default 0,
  lpv        bigint  not null default 0,
  checkouts  bigint  not null default 0,
  vendas     bigint  not null default 0,
  valor      numeric not null default 0,
  primary key (data, anuncio)
);

create index if not exists diario_data_idx    on public.criativos_diario (data);
create index if not exists diario_anuncio_idx on public.criativos_diario (anuncio);


-- ── 5. Quem pode o quê ─────────────────────────────────────────────────────
-- A chave publishable fica VISÍVEL no código da página — é assim por design.
-- Quem protege os dados são estas regras, não a chave.
alter table public.eventos          enable row level security;
alter table public.criativos        enable row level security;
alter table public.criativos_diario enable row level security;
alter table public.paginas          enable row level security;
alter table public.consolidar       enable row level security;

-- o quiz é anônimo: precisa INSERIR, e só isso. Não pode ler nada.
drop policy if exists "quiz insere evento" on public.eventos;
create policy "quiz insere evento" on public.eventos
  for insert to anon with check (true);

-- leitura do dashboard: só quem está logado
drop policy if exists "logado le eventos" on public.eventos;
create policy "logado le eventos" on public.eventos
  for select to authenticated using (true);

drop policy if exists "logado le criativos" on public.criativos;
create policy "logado le criativos" on public.criativos
  for select to authenticated using (true);

-- etiquetas e ajustes você edita pelo dashboard
drop policy if exists "logado edita criativos" on public.criativos;
create policy "logado edita criativos" on public.criativos
  for update to authenticated using (true) with check (true);

drop policy if exists "logado le diario" on public.criativos_diario;
create policy "logado le diario" on public.criativos_diario
  for select to authenticated using (true);

drop policy if exists "logado le paginas" on public.paginas;
create policy "logado le paginas" on public.paginas
  for select to authenticated using (true);

drop policy if exists "logado le consolidar" on public.consolidar;
create policy "logado le consolidar" on public.consolidar
  for select to authenticated using (true);

-- Nenhuma policy de INSERT/UPDATE para o robô do Meta: ele roda como service_role
-- na Edge Function, que ignora RLS por definição. A chave dele fica nos secrets
-- do projeto, nunca na página.
