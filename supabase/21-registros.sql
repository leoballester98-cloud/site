-- ═══════════════════════════════════════════════════════════════════════════
--  Registros — o caderno da operação
--
--  Substitui `referencias`, que só sabia guardar NÚMERO: `valor` era not null,
--  o botão dizia "Guardar número" e a tabela tinha coluna Valor. Lembrete e
--  anotação de mudança não têm número, e por isso simplesmente não cabiam.
--
--  Aqui o tipo decide o que a linha carrega:
--    dado      um número que some do dashboard quando algo muda
--    lembrete  algo pra fazer ou vigiar; ganha `feito`
--    mudanca   o que foi mexido e quando
-- ═══════════════════════════════════════════════════════════════════════════

create table if not exists public.registros (
  id        bigserial primary key,
  tipo      text not null check (tipo in ('dado', 'lembrete', 'mudanca')),
  titulo    text not null check (length(btrim(titulo)) between 1 and 160),

  /* Opcional, ao contrário de referencias.valor. É a mudança que faz os outros
     dois tipos existirem. */
  valor     numeric(14,4),
  unidade   text check (unidade in ('%', 'R$', 'n')),

  /* `data` é quando aconteceu. `ate` fecha um PERÍODO e só interessa pro tipo
     dado: "a conversão era 21,2%" vale de um dia até outro, e é esse intervalo
     que diz onde o depois começa. Lembrete e mudança são pontuais. */
  data      date not null default (now() at time zone 'America/Sao_Paulo')::date,
  ate       date,

  nota      text,
  feito     boolean not null default false,
  criado_em timestamptz not null default now(),

  constraint registros_periodo_ok check (ate is null or ate >= data),
  /* Dado sem número não é dado — sem isto o tipo viraria só um rótulo, e a
     coluna Valor da tela mostraria vazio pra algo que se anuncia como número. */
  constraint registros_dado_tem_valor check (tipo <> 'dado' or valor is not null)
);

create index if not exists registros_data_idx on public.registros (feito, data desc);

alter table public.registros enable row level security;

drop policy if exists registros_ler   on public.registros;
drop policy if exists registros_mexer on public.registros;
create policy registros_ler   on public.registros for select using (public.tem_acesso());
create policy registros_mexer on public.registros for all
  using (public.tem_acesso()) with check (public.tem_acesso());

revoke all on public.registros from anon;
grant select, insert, update, delete on public.registros to authenticated;


/* Traz o que já existia como 'dado'. O `not exists` protege contra rodar duas
   vezes: sem ele, o segundo run duplicaria tudo.

   `referencias` fica de pé, com as linhas dentro. Só apague depois de conferir
   na tela que veio tudo — desfazer um drop é bem mais caro que ignorar uma
   tabela parada. */
insert into public.registros (tipo, titulo, valor, unidade, data, ate, nota, criado_em)
select 'dado', r.rotulo, r.valor, r.unidade, r.data, nullif(r.ate, r.data), r.nota, r.criado_em
from public.referencias r
where not exists (select 1 from public.registros);
