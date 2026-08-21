// ═══════════════════════════════════════════════════════════════════════════
//  atualizar-criativos — substitui o gatilho atualizarCriativos do Apps Script
//
//  Busca as métricas do Meta, grava um dia por anúncio em criativos_diario e
//  atualiza a linha de cada criativo preservando as etiquetas.
//
//  Roda com a service_role (variável de ambiente do próprio Supabase), então
//  ignora o RLS de propósito: é o único caminho que escreve nas tabelas de
//  criativos. O token do Meta fica nos secrets do projeto, nunca numa página.
// ═══════════════════════════════════════════════════════════════════════════

import { createClient } from 'jsr:@supabase/supabase-js@2';

const META_API_VER  = 'v21.0';
const META_CONTA    = 'act_700396725669598';
const CAMP_FILTRO   = 'Quiz';                 // só campanhas com isso no nome
const EXCLUIR_CAMP  = ['120251947450640652', '120251870932890652'];  // testes descartados
const EXCLUIR_ADS   = ['120251807360980652'];                        // cópia acidental do ad3
const DIAS          = 1095;                   // 3 anos, dentro do limite de 37 meses da API

const soma = (arr: any[] | undefined) =>
  (arr ?? []).reduce((s, x) => s + (Number(x.value) || 0), 0);

const hojeISO = () =>
  new Date(Date.now() - 3 * 3600 * 1000).toISOString().slice(0, 10);   // America/Sao_Paulo

const diasAtras = (iso: string, n: number) => {
  const d = new Date(iso + 'T12:00:00Z');
  d.setUTCDate(d.getUTCDate() - n);
  return d.toISOString().slice(0, 10);
};

/* Segue a paginação do Meta. O guard existe porque um erro de filtro poderia
   virar um laço infinito consumindo a cota da API. */
async function metaFetch(url: string) {
  const out: any[] = [];
  let next: string | null = url;
  for (let i = 0; next && i < 25; i++) {
    const resp = await fetch(next);
    const json = await resp.json();
    if (json.error) throw new Error('Meta API: ' + json.error.message);
    out.push(...(json.data ?? []));
    next = json.paging?.next ?? null;
  }
  return out;
}

Deno.serve(async (req) => {
  try {
    const token = Deno.env.get('META_TOKEN');
    if (!token) throw new Error('Falta o secret META_TOKEN no projeto.');

    const db = createClient(
      Deno.env.get('SUPABASE_URL')!,
      Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!,
      { auth: { persistSession: false } },
    );

    const ate   = hojeISO();
    const ontem = diasAtras(ate, 1);

    const fields = [
      'ad_id', 'ad_name', 'campaign_id', 'campaign_name', 'spend', 'impressions', 'clicks',
      'actions', 'action_values', 'video_thruplay_watched_actions',
      'video_continuous_2_sec_watched_actions', 'video_p25_watched_actions', 'video_play_actions',
    ].join(',');

    const params = new URLSearchParams({
      level: 'ad',
      // date_preset=maximum CORTA o dia de hoje (comprovado em 19/08/2026), e era
      // isso que deixava o filtro "Hoje" sempre zerado. Intervalo explícito resolve.
      time_range: JSON.stringify({ since: diasAtras(ate, DIAS), until: ate }),
      time_increment: '1',
      limit: '500',
      fields: fields + ',video_3_sec_watched_actions',
      filtering: JSON.stringify([{ field: 'campaign.name', operator: 'CONTAIN', value: CAMP_FILTRO }]),
      access_token: token,
    });

    const base = `https://graph.facebook.com/${META_API_VER}/${META_CONTA}/insights?`;
    let linhas: any[];
    try {
      linhas = await metaFetch(base + params.toString());
    } catch {
      // algumas contas rejeitam video_3_sec_watched_actions; refaz sem ele
      params.set('fields', fields);
      linhas = await metaFetch(base + params.toString());
    }

    // anúncios que são o mesmo criativo somam na linha canônica
    const { data: cons } = await db.from('consolidar').select('anuncio, canonico');
    const consolidar: Record<string, string> = {};
    (cons ?? []).forEach((c) => { consolidar[c.anuncio] = c.canonico; });

    /* Chaveado por (data, anúncio) porque a tabela tem PK nesse par e duas coisas
       geram colisão: a consolidação (tela original + tela 1 + ad1 viram a mesma
       linha) e anúncios com nome repetido em campanhas diferentes. O Apps Script
       não via isso — a planilha não tem chave e ele somava só na leitura. */
    const diarioMap: Record<string, any> = {};
    const ag: Record<string, any> = {};

    for (const r of linhas) {
      if (EXCLUIR_CAMP.includes(String(r.campaign_id))) continue;
      if (EXCLUIR_ADS.includes(String(r.ad_id))) continue;

      const nomeMeta = r.ad_name || r.ad_id;
      const nome = consolidar[nomeMeta] ?? nomeMeta;
      const consolidado = nome !== nomeMeta;

      const d = {
        gasto: Number(r.spend) || 0,
        impressoes: Number(r.impressions) || 0,
        clicks: Number(r.clicks) || 0,
        thruplay: soma(r.video_thruplay_watched_actions),
        v2s: soma(r.video_continuous_2_sec_watched_actions),
        v3s: soma(r.video_3_sec_watched_actions),
        p25: soma(r.video_p25_watched_actions),
        lpv: 0, checkouts: 0, vendas: 0, valor: 0,
      };

      for (const ac of (r.actions ?? [])) {
        const t = ac.action_type || '';
        const v = Number(ac.value) || 0;
        if (t === 'video_view') d.v3s += v;
        if (t === 'landing_page_view') d.lpv += v;
        if (['initiate_checkout', 'omni_initiated_checkout', 'offsite_conversion.fb_pixel_initiate_checkout'].includes(t)) {
          d.checkouts = Math.max(d.checkouts, v);   // max, não soma: a API repete o mesmo evento em vários action_type
        }
        if (['purchase', 'omni_purchase', 'offsite_conversion.fb_pixel_purchase'].includes(t)) {
          d.vendas = Math.max(d.vendas, v);
        }
      }
      for (const ac of (r.action_values ?? [])) {
        const t = ac.action_type || '';
        const v = Number(ac.value) || 0;
        if (['purchase', 'omni_purchase', 'offsite_conversion.fb_pixel_purchase'].includes(t)) {
          d.valor = Math.max(d.valor, v);           // já vem com o order bump somado
        }
      }

      const chave = r.date_start + '|' + nome;
      if (!diarioMap[chave]) {
        diarioMap[chave] = {
          data: r.date_start, anuncio: nome, ad_id: String(r.ad_id ?? ''),
          campanha: r.campaign_name ?? '',
          gasto: 0, impressoes: 0, clicks: 0, v3s: 0, v2s: 0,
          thruplay: 0, p25: 0, lpv: 0, checkouts: 0, vendas: 0, valor: 0,
        };
      }
      const dm = diarioMap[chave];
      for (const k of ['gasto','impressoes','clicks','v3s','v2s','thruplay','p25','lpv','checkouts','vendas','valor']) {
        dm[k] += (d as any)[k];
      }

      if (!ag[nome]) ag[nome] = { ad_id: r.ad_id, campanha: r.campaign_name, impressoes: 0, recente: false };
      const a = ag[nome];
      a.impressoes += d.impressoes;
      // entrega hoje ou ontem = ainda rodando. Só hoje seria instável: de manhã
      // cedo um anúncio ativo ainda não gastou e piscaria como testado.
      if (d.impressoes > 0 && String(r.date_start) >= ontem) a.recente = true;
      // a campanha do canônico não pode ser sobrescrita pela do anúncio absorvido
      if (!consolidado) a.campanha = r.campaign_name || a.campanha;
    }

    const diario = Object.values(diarioMap);

    /* Reescreve a tabela diária inteira. O erro do delete era engolido: se ele
       falhasse, o insert seguinte batia nas linhas antigas e o erro aparecia
       como "duplicate key", apontando pro lugar errado. */
    const del = await db.from('criativos_diario').delete().gte('data', '1900-01-01');
    if (del.error) throw new Error('Erro limpando diário: ' + del.error.message);

    /* upsert em vez de insert: idempotente. Se sobrar qualquer linha (delete que
       falhou, execução anterior interrompida no meio dos lotes), ela é
       sobrescrita em vez de derrubar a execução inteira. */
    let gravadas = 0;
    for (let i = 0; i < diario.length; i += 500) {
      const { error } = await db.from('criativos_diario')
        .upsert(diario.slice(i, i + 500), { onConflict: 'data,anuncio' });
      if (error) throw new Error('Erro gravando diário (lote ' + (i / 500 + 1) + '): ' + error.message);
      gravadas += Math.min(500, diario.length - i);
    }

    // linhas de criativo: cria as novas e atualiza as existentes SEM tocar nas etiquetas
    const { data: exist } = await db.from('criativos').select('anuncio, status');
    const statusAtual: Record<string, string> = {};
    (exist ?? []).forEach((c) => { statusAtual[c.anuncio] = String(c.status ?? '').trim().toLowerCase(); });

    const paraGravar = Object.keys(ag)
      .filter((nome) => ag[nome].impressoes > 0)     // ignora ad sem entrega
      .map((nome) => {
        const a = ag[nome];
        const st = statusAtual[nome];
        const auto = a.recente ? 'ativo' : 'testado';
        return {
          anuncio: nome,
          ad_id: String(a.ad_id ?? ''),
          campanha: a.campanha ?? '',
          // rótulo que você escreveu à mão (ex: descartado) é preservado
          // 'rodando' era o nome antigo de 'ativo' — segue sendo sobrescrito
          status: (!st || st === 'ativo' || st === 'rodando' || st === 'testado') ? auto : statusAtual[nome],
          atualizado: new Date().toISOString(),
        };
      });

    if (paraGravar.length) {
      const { error } = await db.from('criativos')
        .upsert(paraGravar, { onConflict: 'anuncio', ignoreDuplicates: false });
      if (error) throw new Error('Erro gravando criativos: ' + error.message);
    }

    return Response.json({
      ok: true,
      linhas_meta: linhas.length,
      dias_gravados: diario.length,
      criativos: paraGravar.length,
      ate,
    });
  } catch (e) {
    return Response.json({ ok: false, erro: String(e) }, { status: 500 });
  }
});
