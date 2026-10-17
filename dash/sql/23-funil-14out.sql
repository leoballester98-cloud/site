-- ═══════════════════════════════════════════════════════════════════════════
--  Funil do dash na ordem real do quiz (14/10/2026)
--
--  - Sai o que não existe mais no quiz: 7, 8, 10, 12, 24, 25.
--  - Toda tela ganha a posição dela no caminho real (ordem 1..45).
--  - 'tela' só pra quem grava número em ordem crescente ao longo do caminho:
--    o funil conta 'tela' por maxi >= etapa. As que gravam fora da sequência
--    (4 objetivo, 9 sabonete) e as de número alto viram 'vsl', que conta só
--    quem viu.
--  - A ordem vai primeiro pra um valor provisório (etapa/1000) porque o
--    índice único em ordem não aceita duas linhas com o mesmo número nem
--    durante a troca.
-- ═══════════════════════════════════════════════════════════════════════════
begin;

delete from public.etapas where etapa in (7, 8, 10, 12, 24, 25);

update public.etapas set ordem = etapa / 1000.0 where tipo <> 'compra';

insert into public.etapas (etapa, ordem, nome, tipo) values
  ( 1,  1, '(abertura/idade)',                          'tela'),
  ( 2,  2, 'Mais de 1.800 mulheres (prova social)',     'tela'),
  (37,  3, 'Já começou a tentar engravidar?',           'vsl'),
  (38,  4, '↳ Ótimo momento pra começar (ainda não)',   'vsl'),
  (39,  5, '↳ Você está no caminho certo (já tenta)',   'vsl'),
  ( 4,  6, 'Qual é o seu principal objetivo?',          'vsl'),
  (41,  7, 'Nós sabemos como fazer isso acontecer',     'vsl'),
  ( 3,  8, 'Seu ciclo menstrual é…',                    'tela'),
  ( 5,  9, 'Em qual área sente algo travado?',          'tela'),
  ( 6, 10, 'Há quanto tempo está tentando?',            'tela'),
  (43, 11, 'Modo sobrevivência (resposta ao tempo)',    'vsl'),
  (11, 12, 'Você sabe quando está ovulando?',           'tela'),
  (13, 13, 'Como identifica seu período fértil?',       'tela'),
  ( 9, 14, 'Sabonete íntimo/ducha/lenços?',             'vsl'),
  (42, 15, 'Pote de plástico/parabenos?',               'vsl'),
  (36, 16, '▶ VSL 1 · o mecanismo',                     'vsl'),
  (14, 17, 'Como descreveria seu estresse?',            'tela'),
  (15, 18, 'Tem um momento do dia só pra você?',        'tela'),
  (16, 19, 'Ansiosa antes da menstruação?',             'tela'),
  (17, 20, 'Pressão de engravidar',                     'tela'),
  (50, 21, 'Estamos com você (fecho emocional)',        'vsl'),
  (18, 22, 'Humor muda ao longo do dia?',               'tela'),
  (19, 23, 'Sono',                                      'tela'),
  (20, 24, 'Atividade física',                          'tela'),
  (21, 25, 'Alimentação',                               'tela'),
  (22, 26, 'Inchaço, cansaço ou dores?',                'tela'),
  (51, 27, 'Mais leve e com energia (fecho organismo)', 'vsl'),
  (44, 28, 'Como se sente em relação ao corpo?',        'vsl'),
  (45, 29, 'Se sente acompanhada?',                     'vsl'),
  (23, 30, 'Ainda acredita que vai conseguir?',         'tela'),
  (26, 31, 'Se culpa quando o ciclo não vem?',          'tela'),
  (52, 32, 'Recupere a confiança (fecho propósito)',    'vsl'),
  (27, 33, 'Quantos dias dura seu ciclo?',              'tela'),
  (28, 34, 'Quantos dias dura a menstruação?',          'tela'),
  (29, 35, 'Em quantos dias o desejo sexual aumenta?',  'tela'),
  (32, 36, '⏳ Analisando suas respostas (barras)',      'vsl'),
  (30, 37, 'Seu perfil fértil (resultado)',             'tela'),
  (31, 38, 'Quando começou a última menstruação?',      'tela'),
  (53, 39, 'Baseado em pesquisas (centros)',            'vsl'),
  (49, 40, '⏳ Criando seu plano (anel)',                'vsl'),
  (47, 41, 'E-mail',                                    'vsl'),
  (48, 42, 'Nome',                                      'vsl'),
  (33, 43, 'Gráfico: o plano que vai ajudar',           'tela'),
  (46, 44, 'Raspadinha',                                'vsl'),
  (34, 45, 'Página de vendas',                          'tela')
on conflict (etapa) do update
  set ordem = excluded.ordem, nome = excluded.nome, tipo = excluded.tipo;

commit;

select etapa, ordem, tipo, nome, desde from public.etapas order by ordem;
