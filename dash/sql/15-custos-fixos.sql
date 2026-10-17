-- ═══════════════════════════════════════════════════════════════════════════
--  Custos fixos — os que voltam todo mês
--
--  Guarda a REGRA, não as ocorrências. Cloud a R$100/mês vira uma linha só, e
--  cada mês do período é calculado na hora da leitura.
--
--  A alternativa seria um robô criando um lançamento por mês. Ela quebra de
--  dois jeitos previsíveis: se o robô falha um mês, o custo some daquele mês e
--  ninguém percebe; e mudar o valor obriga a corrigir todas as linhas já
--  criadas, ou a conviver com o histórico errado. Guardando a regra, o passado
--  se recalcula sozinho e nada precisa rodar de madrugada.
-- ═══════════════════════════════════════════════════════════════════════════

create table if not exists public.custos_fixos (
  id        uuid primary key default gen_random_uuid(),
  nome      text not null check (length(btrim(nome)) between 1 and 120),
  valor     numeric(12,2) not null check (valor >= 0),
  /* Até 28 de propósito: 29, 30 e 31 não existem em todo mês, e um custo que
     pula fevereiro seria um bug difícil de enxergar. */
  dia       smallint not null default 1 check (dia between 1 and 28),
  inicio    date not null default (now() at time zone 'America/Sao_Paulo')::date,
  fim       date,          -- null = ainda ativo. Encerrar preserva o histórico.
  criado_em timestamptz not null default now(),
  constraint custos_fixos_periodo_ok check (fim is null or fim >= inicio)
);

alter table public.custos_fixos enable row level security;

drop policy if exists custos_fixos_ler     on public.custos_fixos;
drop policy if exists custos_fixos_inserir on public.custos_fixos;
drop policy if exists custos_fixos_editar  on public.custos_fixos;
drop policy if exists custos_fixos_apagar  on public.custos_fixos;

create policy custos_fixos_ler     on public.custos_fixos for select using (public.tem_acesso());
create policy custos_fixos_inserir on public.custos_fixos for insert with check (public.tem_acesso());
create policy custos_fixos_editar  on public.custos_fixos for update using (public.tem_acesso())
                                                             with check (public.tem_acesso());
create policy custos_fixos_apagar  on public.custos_fixos for delete using (public.tem_acesso());

revoke all on public.custos_fixos from anon;
grant select, insert, update, delete on public.custos_fixos to authenticated;


/* Os avulsos e as ocorrências dos fixos numa lista só, já no formato que o
   dashboard desenha. Uma consulta só porque o total do período precisa somar
   os dois — separados, seria fácil um deles ficar de fora de alguma conta. */
create or replace function public.fn_custos(
  p_de  date default null,
  p_ate date default null
)
returns table (id text, nome text, valor numeric, data date, fixo boolean)
language sql
stable
security invoker
set search_path = public
as $$
  with lim as (
    /* Sem filtro de data, a janela vai do primeiro custo fixo até hoje: gerar
       ocorrências de um intervalo aberto não terminaria. */
    select coalesce(p_de,  (select min(inicio) from public.custos_fixos))            as de,
           coalesce(p_ate, (now() at time zone 'America/Sao_Paulo')::date)           as ate
  )
  select c.id::text, c.nome, c.valor, c.data, false
  from public.custos c
  where (p_de  is null or c.data >= p_de)
    and (p_ate is null or c.data <= p_ate)

  union all

  select f.id::text || ':' || to_char(x.oc, 'YYYY-MM'), f.nome, f.valor, x.oc, true
  from public.custos_fixos f, lim
  cross join lateral (
    select (date_trunc('month', g)::date + (f.dia - 1)) as oc
    from generate_series(date_trunc('month', lim.de), date_trunc('month', lim.ate),
                         interval '1 month') g
  ) x
  where x.oc between lim.de and lim.ate
    and x.oc >= f.inicio
    and (f.fim is null or x.oc <= f.fim)

  order by 4 desc;
$$;

grant  execute on function public.fn_custos(date, date) to authenticated;
revoke execute on function public.fn_custos(date, date) from anon;
