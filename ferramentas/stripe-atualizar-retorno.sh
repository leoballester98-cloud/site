#!/usr/bin/env bash
# Aponta o retorno dos 3 links de pagamento de TESTE ja criados pra pagina de
# obrigado, com ?plano= de cada um (os links foram criados antes dela existir).
# A chave sk_test_ vem da area de transferencia e e apagada dela em seguida.
set -euo pipefail

LINKS="$(cd "$(dirname "$0")" && pwd)/stripe-links-teste.json"
OBRIGADO='https://protocolociclofertil.com.br/quiz/v2/obrigado.html?session_id={CHECKOUT_SESSION_ID}'

STRIPE_KEY=$(pbpaste | tr -d '[:space:]')
printf '' | pbcopy
case "$STRIPE_KEY" in
  sk_test_*|rk_test_*) ;;
  *) echo "O que esta copiado nao e uma chave de teste (sk_test_...). Nada foi alterado."; exit 1 ;;
esac

for plano in essencial avancado imersao; do
  link_id=$(python3 -c "import json,sys; print(json.load(open(sys.argv[1]))[sys.argv[2]]['link_id'])" "$LINKS" "$plano")
  curl -sS --fail-with-body "https://api.stripe.com/v1/payment_links/$link_id" \
    -u "$STRIPE_KEY:" \
    -d "after_completion[type]=redirect" \
    --data-urlencode "after_completion[redirect][url]=$OBRIGADO&plano=$plano" > /dev/null
  echo "✓ $plano → obrigado.html?plano=$plano"
done
unset STRIPE_KEY
