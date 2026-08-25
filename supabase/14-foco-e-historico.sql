-- ═══════════════════════════════════════════════════════════════════════════
--  1. Cada experimento mostra só as colunas da SUA pergunta
--  2. Histórico: o caderno da operação
-- ═══════════════════════════════════════════════════════════════════════════

/* Um teste de tela pergunta "quantas pessoas seguem?"; um de preço pergunta
   "quanto dinheiro entra?". Mostrar as duas coisas nos dois cartões não dá mais
   informação — dá mais coluna pra atravessar antes de achar a que importa, e
   uma tabela que exige garimpo é uma tabela que se lê errado.

   'tela'     -> visitantes, tela 1 → tela 2, chegou ao fim, clicou no checkout
   'dinheiro' -> visitantes, clicou no checkout, vendas, receita, R$/visitante */
alter table public.experimentos
  add column if not exists foco text not null default 'tela';

alter table public.experimentos drop constraint if exists experimentos_foco_ok;
alter table public.experimentos
  add constraint experimentos_foco_ok check (foco in ('tela', 'dinheiro'));

update public.experimentos set foco = 'dinheiro' where chave = 'preco';
update public.experimentos set foco = 'tela'     where chave = 'h3';


-- ═══════════════════════════════════════════════════════════════════════════
--  Histórico — o que estava valendo antes de você mexer
--
--  Existe porque o dashboard mede o AGORA. Quando você troca o checkout de
--  gateway, muda o preço ou reescreve a página de vendas, o número de antes
--  deixa de existir em qualquer lugar — e é exatamente contra ele que o novo
--  precisa ser comparado. Anotar é o que transforma uma mudança em experimento.
-- ═══════════════════════════════════════════════════════════════════════════

create table if not exists public.historico (
  id        bigserial primary key,
  data      date not null default (now() at time zone 'America/Sao_Paulo')::date,
  tipo      text not null default 'nota',
  titulo    text not null,
  texto     text,
  criado_em timestamptz not null default now(),
  constraint historico_tipo_ok check (tipo in ('nota','marco','decisao','numero'))
);

create index if not exists historico_data_idx on public.historico (data desc);

alter table public.historico enable row level security;

drop policy if exists "logado le historico" on public.historico;
create policy "logado le historico" on public.historico
  for select to authenticated using (true);

drop policy if exists "logado edita historico" on public.historico;
create policy "logado edita historico" on public.historico
  for all to authenticated using (true) with check (true);

revoke all on public.historico from anon;
