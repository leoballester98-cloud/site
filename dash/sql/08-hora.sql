-- ═══════════════════════════════════════════════════════════════════════════
--  Gasto e faturamento por HORA
--
--  Existe por um motivo só: com o filtro num dia único não há linha a desenhar,
--  porque linha liga dois pontos. Quebrado por hora, um dia vira até 24 pontos.
--
--  Guarda só uma janela curta (o robô reescreve os últimos dias a cada volta).
--  Hora × anúncio × dia multiplica rápido, e ninguém abre o detalhe de hora de
--  um dia de três meses atrás — pra isso o diário já serve.
-- ═══════════════════════════════════════════════════════════════════════════

create table if not exists public.criativos_hora (
  data      date    not null,
  hora      smallint not null check (hora between 0 and 23),
  anuncio   text    not null,
  gasto     numeric(12,2) not null default 0,
  checkouts integer not null default 0,
  vendas    integer not null default 0,
  valor     numeric(12,2) not null default 0,
  primary key (data, hora, anuncio)
);

create index if not exists criativos_hora_data_idx on public.criativos_hora (data);

alter table public.criativos_hora enable row level security;

drop policy if exists hora_ler on public.criativos_hora;
create policy hora_ler on public.criativos_hora for select using (public.tem_acesso());

revoke all on public.criativos_hora from anon;
grant select on public.criativos_hora to authenticated;
-- quem escreve é o robô, com a service role, que ignora RLS

/* Série por hora do intervalo. Devolve as 24 horas mesmo sem movimento: sem
   isso a linha pularia as horas vazias e mentiria sobre o formato do dia. */
create or replace function public.fn_horario(
  p_de  date default null,
  p_ate date default null
)
returns table (hora smallint, gasto numeric, faturamento numeric, vendas bigint, checkouts bigint)
language sql
stable
security invoker
set search_path = public
as $$
  select
    h::smallint                                        as hora,
    coalesce(sum(c.gasto), 0)                          as gasto,
    coalesce(sum(case when c.valor > 0 then c.valor else c.vendas * 37.90 end), 0) as faturamento,
    coalesce(sum(c.vendas), 0)::bigint                 as vendas,
    coalesce(sum(c.checkouts), 0)::bigint              as checkouts
  from generate_series(0, 23) as h
  left join public.criativos_hora c
    on c.hora = h
   and (p_de  is null or c.data >= p_de)
   and (p_ate is null or c.data <= p_ate)
  group by h
  order by h;
$$;

revoke execute on function public.fn_horario(date, date) from public;
grant  execute on function public.fn_horario(date, date) to authenticated;
