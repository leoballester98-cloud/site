// ═══════════════════════════════════════════════════════════════════════════
//  kiwify-venda — recebe o webhook de venda da Kiwify
//
//  Existe por causa do teste de preço: as vendas que o dashboard já tem vêm do
//  robô do Meta, agregadas por anúncio, e não sabem qual preço a pessoa viu.
//  Aqui a venda chega com o id do produto, e o produto diz o braço.
//
//  Roda com a service_role, então ignora o RLS — é o único caminho que escreve
//  na tabela `vendas`. Precisa ser publicada SEM verificação de JWT: a Kiwify
//  não manda Authorization, e com a verificação ligada todo webhook toma 401.
//
//    supabase functions deploy kiwify-venda --no-verify-jwt
//
//  Assinatura: a Kiwify assina o corpo com HMAC e manda no query `?signature=`.
//  O segredo é o Token do webhook, que vive nos secrets como KIWIFY_TOKEN.
//  Sem o secret configurado a função ACEITA e marca a linha como não conferida
//  — assim o primeiro teste passa e você vê o payload, em vez de depurar às
//  cegas um 401. Com o secret, rejeita o que não bate.
// ═══════════════════════════════════════════════════════════════════════════

import { createClient } from 'jsr:@supabase/supabase-js@2';

/* A Kiwify usa HMAC-SHA1 do corpo cru. Precisa ser o corpo CRU, byte a byte:
   se der JSON.parse e re-serializar, a assinatura não bate mais — qualquer
   diferença de espaço ou de ordem de chave muda o hash. */
async function assinaturaOk(corpo: string, assinatura: string, segredo: string) {
  const enc = new TextEncoder();
  const chave = await crypto.subtle.importKey(
    'raw', enc.encode(segredo), { name: 'HMAC', hash: 'SHA-1' }, false, ['sign'],
  );
  const mac = await crypto.subtle.sign('HMAC', chave, enc.encode(corpo));
  const hex = Array.from(new Uint8Array(mac))
    .map((b) => b.toString(16).padStart(2, '0')).join('');
  return hex === assinatura.toLowerCase();
}

/* Os nomes de campo da Kiwify variam entre versões de webhook, e alguns vêm
   aninhados. Em vez de fixar um caminho, procura o primeiro que existir. */
function pega(obj: any, caminhos: string[][]): any {
  for (const c of caminhos) {
    let v = obj;
    for (const k of c) { v = v?.[k]; if (v === undefined || v === null) break; }
    if (v !== undefined && v !== null && v !== '') return v;
  }
  return null;
}

/* A Kiwify manda o valor em CENTAVOS, inteiro, em Commissions.charge_amount.
   Confirmado no payload real: 6134 para uma venda de R$61,34. Então divide por
   100 e pronto — sem adivinhar a unidade.

   A versão anterior chutava ("se for inteiro e >= 1000, é centavo") e passava
   nos testes com 3790 e 6790 justamente porque os dois preços do experimento
   caem acima do corte. Uma venda de R$9,90 chega como 990, ficaria abaixo dele
   e entraria como R$990,00. O palpite acertava por sorte de faixa de preço. */
function valorDe(dados: any): number {
  const cent = pega(dados, [['Commissions', 'charge_amount'], ['charge_amount']]);
  if (cent !== null && isFinite(Number(cent))) return Number(cent) / 100;

  /* Reserva, caso o campo suma numa versão futura do webhook. Aí a unidade
     volta a ser desconhecida e o palpite volta com ela — mas o payload cru fica
     guardado, então dá pra corrigir depois em vez de perder o dado. */
  const n = Number(String(pega(dados, [['order_total'], ['total'], ['price']]) ?? 0).replace(',', '.'));
  if (!isFinite(n)) return 0;
  return Number.isInteger(n) && Math.abs(n) >= 1000 ? n / 100 : n;
}

const EVENTOS: Record<string, string> = {
  'order_approved': 'compra_aprovada',
  'order.paid': 'compra_aprovada',
  'paid': 'compra_aprovada',
  'approved': 'compra_aprovada',
  'order_refunded': 'reembolso',
  'refunded': 'reembolso',
  'chargeback': 'chargeback',
};

Deno.serve(async (req) => {
  if (req.method !== 'POST') return new Response('ok', { status: 200 });

  const cru = await req.text();
  const segredo = Deno.env.get('KIWIFY_TOKEN');
  const assinatura = new URL(req.url).searchParams.get('signature') ?? '';

  let conferido = false;
  if (segredo) {
    conferido = assinatura ? await assinaturaOk(cru, assinatura, segredo) : false;
    if (!conferido) return new Response('assinatura invalida', { status: 401 });
  }

  let corpo: any;
  try { corpo = JSON.parse(cru); }
  catch { return new Response('json invalido', { status: 400 }); }

  const dados = corpo?.order ?? corpo?.data ?? corpo;

  const id = String(pega(dados, [['order_id'], ['id'], ['order_ref'], ['reference']]) ?? crypto.randomUUID());
  const statusCru = String(pega(dados, [['order_status'], ['status'], ['webhook_event_type']]) ?? '').toLowerCase();
  const eventoCru = String(corpo?.webhook_event_type ?? corpo?.event ?? statusCru).toLowerCase();

  const db = createClient(
    Deno.env.get('SUPABASE_URL')!,
    Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!,
    { auth: { persistSession: false } },
  );

  /* upsert e não insert: a Kiwify reentrega o mesmo evento quando a resposta
     demora, e cada reentrega viraria uma venda a mais na contagem. */
  const { error } = await db.from('vendas').upsert({
    id,
    evento: EVENTOS[eventoCru] ?? EVENTOS[statusCru] ?? 'outro',
    produto_id: pega(dados, [['product_id'], ['Product', 'product_id'], ['product', 'id'],
                             ['Commissions', 'product_id']])?.toString() ?? null,
    produto: pega(dados, [['product_name'], ['Product', 'product_name'], ['product', 'name']])?.toString() ?? null,
    valor: valorDe(dados),
    moeda: pega(dados, [['Commissions', 'currency'], ['currency']])?.toString() ?? 'BRL',
    status: statusCru || null,
    /* Guarda tudo, inclusive se a assinatura foi conferida: sem isso não dá pra
       distinguir depois um webhook legítimo de um que passou porque o secret
       ainda não estava configurado. */
    bruto: { conferido, recebido_em: new Date().toISOString(), payload: corpo },
  }, { onConflict: 'id' });

  if (error) {
    console.error('erro ao gravar venda', error);
    return new Response('erro', { status: 500 });
  }
  /* 200 sempre que gravou. A Kiwify reentrega em cima de qualquer outra coisa,
     e reentrega infinita de um payload que a gente já guardou não ajuda. */
  return new Response('ok', { status: 200 });
});
