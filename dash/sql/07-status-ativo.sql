-- 'rodando' virou 'ativo'. O robô já grava o nome novo; isto acerta o que ficou
-- gravado antes, pra não conviverem dois nomes pro mesmo estado.
update public.criativos set status = 'ativo' where lower(btrim(status)) = 'rodando';
