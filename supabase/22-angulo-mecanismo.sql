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
