-- VSL 1 contada pela posição, como as outras telas (14/10/2026, já aplicado).
-- Antes era 'vsl' (azul no dash, só quem disparou o ping), o que fazia ela
-- aparecer com menos gente que a tela seguinte.
update public.etapas set tipo = 'tela' where etapa = 36;
select etapa, ordem, tipo, nome from public.etapas where etapa = 36;
