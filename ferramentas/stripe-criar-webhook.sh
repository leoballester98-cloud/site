#!/usr/bin/env bash
# Cadastra na Stripe (MODO TESTE) o endereco do webhook que libera o acesso no
# app, e guarda o segredo de assinatura dele direto nos secrets do Supabase.
#
# A chave sk_test_ vem da area de transferencia (copie na Stripe antes de
# rodar) e e apagada dela logo em seguida. O segredo whsec_ que a Stripe
# devolve vai pro Supabase por um arquivo temporario que some no fim; nada
# disso aparece na tela nem fica no historico.
#
# Precisa do login no Supabase feito antes (npx supabase login).
set -euo pipefail

PROJETO=lfmjvgkbvkkexwytkfeg
URL="https://$PROJETO.supabase.co/functions/v1/purchase-webhook-stripe"

STRIPE_KEY=$(pbpaste | tr -d '[:space:]')
printf '' | pbcopy
case "$STRIPE_KEY" in
  sk_test_*|rk_test_*) ;;
  *) echo "O que esta copiado nao e uma chave de teste (sk_test_...). Nada foi criado."; exit 1 ;;
esac

resposta=$(curl -sS --fail-with-body https://api.stripe.com/v1/webhook_endpoints \
  -u "$STRIPE_KEY:" \
  -H "Idempotency-Key: cf-webhook-quiz-v1" \
  --data-urlencode "url=$URL" \
  -d "enabled_events[]=checkout.session.completed" \
  -d "enabled_events[]=checkout.session.async_payment_succeeded" \
  --data-urlencode "description=Libera o acesso ao app (planos do quiz)")
unset STRIPE_KEY

segredo=$(echo "$resposta" | python3 -c "import sys,json; print(json.load(sys.stdin).get('secret',''))")
if [ -z "$segredo" ]; then
  echo "A Stripe criou o webhook mas nao devolveu o segredo (isso acontece se ele ja existia)."
  echo "Pegue em Desenvolvedores → Webhooks → o endereco → Segredo de assinatura."
  exit 1
fi

tmp=$(mktemp); chmod 600 "$tmp"
trap 'rm -f "$tmp"' EXIT
printf 'STRIPE_WEBHOOK_SECRET=%s\n' "$segredo" > "$tmp"
unset segredo
npx --yes supabase secrets set --env-file "$tmp" --project-ref "$PROJETO"

echo
echo "✓ Webhook cadastrado na Stripe: $URL"
echo "✓ Segredo guardado no Supabase como STRIPE_WEBHOOK_SECRET"
