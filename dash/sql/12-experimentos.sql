-- ═══════════════════════════════════════════════════════════════════════════
--  Testes A/B da tela 1 do quiz
--
--  Nenhuma tabela de evento muda. A `eventos` já grava o que o experimento
--  precisa: `variante` diz o braço e `etapa` diz até onde a sessão foi
--  (1..34 = tela, 35 = clique no checkout). O que falta é só um lugar pra
--  DESCREVER o experimento e uma função que leia os dois braços de uma vez.
--
--  A divisão do tráfego (50/50) e quais telas cada braço mostra vivem na
--  página, não aqui. É de propósito: uma página que precisa perguntar ao
--  banco o que desenhar é uma página que pisca — e piscar num teste A/B
--  envenena justamente a métrica da primeira tela.
-- ═══════════════════════════════════════════════════════════════════════════

create table if not exists public.experimentos (
  chave     text primary key,   -- o valor de ?ab= na URL
  rotulo    text not null,      -- como aparece no dashboard
  hipotese  text,               -- o que você espera provar, escrito ANTES
  braco_a   text not null,      -- marca gravada em eventos.variante, sem o 'v'
  braco_b   text not null,
  rotulo_a  text not null,
  rotulo_b  text not null,
  inicio    date,
  fim       date,               -- preenchido quando você encerra
  vencedor  text,               -- 'a' | 'b' | 'empate' | null enquanto roda
  nota      text,
  constraint experimentos_vencedor_ok
    check (vencedor is null or vencedor in ('a','b','empate'))
);

/* A hipótese fica escrita antes de existir dado. É o que separa um teste de
   uma pescaria: com ela no papel, um resultado contrário é uma resposta; sem
   ela, vira caça a qualquer número que tenha subido. */
insert into public.experimentos
  (chave, rotulo, hipotese, braco_a, braco_b, rotulo_a, rotulo_b, inicio)
values
  ('h3',
   'Tela 1 — h=3 vs original',
   'A tela com cartões de foto e headline longa passa mais gente da tela 1 pra tela 2 do que a lista original.',
   'ab-h3', 'ab-orig',
   'Nova (h=3)', 'Original',
   current_date)
on conflict (chave) do nothing;


/* Os dois braços numa consulta só, já no formato que o dashboard desenha.
   Cada sessão conta uma vez, pela MAIOR etapa que alcançou — mesma regra da
   fn_funil, senão os dois números do dashboard não bateriam entre si. */
create or replace function public.fn_experimento(
  p_exp text,
  p_de  date default null,
  p_ate date default null
)
returns table (
  lado text, marca text, rotulo text,
  sessoes bigint, tela1 bigint, tela2 bigint, fim bigint, checkout bigint
)
language sql
stable
security invoker
set search_path = public
as $$
  with e as (
    select * from public.experimentos where chave = p_exp
  ),
  bracos as (
    select 'a'::text as lado, e.braco_a as marca, e.rotulo_a as rotulo from e
    union all
    select 'b'::text,         e.braco_b,          e.rotulo_b          from e
  ),
  /* Uma linha por sessão. O group by inclui a variante porque é ela que liga
     a sessão ao braço; uma sessão só carrega uma. */
  ev as (
    select v.sessao, v.variante,
           max(case when v.etapa between 1 and 34 then v.etapa else 0 end) as maxi,
           bool_or(v.etapa = 35) as comprou
    from public.eventos v
    where v.tipo = 'etapa'
      and v.etapa is not null
      and (p_de  is null or (v.ts at time zone 'America/Sao_Paulo')::date >= p_de)
      and (p_ate is null or (v.ts at time zone 'America/Sao_Paulo')::date <= p_ate)
    group by v.sessao, v.variante
  )
  /* left join pra o braço sem nenhuma sessão ainda aparecer zerado em vez de
     sumir da tabela — braço que some parece braço que não existe. */
  select b.lado, b.marca, b.rotulo,
         count(x.sessao)                             as sessoes,
         count(x.sessao) filter (where x.maxi >= 1)  as tela1,
         count(x.sessao) filter (where x.maxi >= 2)  as tela2,
         count(x.sessao) filter (where x.maxi >= 34) as fim,
         count(x.sessao) filter (where x.comprou)    as checkout
  from bracos b
  left join ev x on x.variante = 'v' || b.marca
  group by b.lado, b.marca, b.rotulo
  order by b.lado;
$$;


-- ── Quem pode o quê ────────────────────────────────────────────────────────
alter table public.experimentos enable row level security;

drop policy if exists "logado le experimentos" on public.experimentos;
create policy "logado le experimentos" on public.experimentos
  for select to authenticated using (true);

/* Você edita rótulo, hipótese, nota, vencedor e data de fim pelo dashboard. */
drop policy if exists "logado edita experimentos" on public.experimentos;
create policy "logado edita experimentos" on public.experimentos
  for update to authenticated using (true) with check (true);

drop policy if exists "logado cria experimentos" on public.experimentos;
create policy "logado cria experimentos" on public.experimentos
  for insert to authenticated with check (true);

revoke all on public.experimentos from anon;

grant  execute on function public.fn_experimento(text, date, date) to authenticated;
revoke execute on function public.fn_experimento(text, date, date) from anon;
