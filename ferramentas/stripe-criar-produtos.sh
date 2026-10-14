#!/usr/bin/env bash
# Cria na Stripe os 3 planos do quiz (pagamento unico, em reais) e um link de
# pagamento pra cada um. Teste com sk_test_; real com sk_live_ E --live:
#   bash stripe-criar-produtos.sh --clipboard --live
#
# A chave e pedida na hora (ou lida da area de transferencia com --clipboard)
# e nao fica salva em lugar nenhum, nem no historico do terminal. O que sai no fim sao so os IDs e os links, que nao sao segredo,
# gravados em ferramentas/stripe-links-teste.json pro quiz e o webhook lerem.
#
# Rodar de novo no mesmo dia nao duplica nada: cada criacao leva uma chave de
# idempotencia fixa, e a Stripe devolve o mesmo objeto da primeira vez.
set -euo pipefail

# Com --clipboard a chave vem do que foi copiado (Cmd+C) e a area de
# transferencia e limpa logo depois. Sem o argumento, pede pra colar.
if [ "${1:-}" = "--clipboard" ]; then
  STRIPE_KEY=$(pbpaste | tr -d '[:space:]')
  printf '' | pbcopy
else
  read -rsp "Cole a chave secreta de TESTE da Stripe (sk_test_...) e aperte Enter: " STRIPE_KEY
  echo
fi
# Modo pela propria chave: sk_test_ cria no teste, sk_live_ no real (e so
# com --live junto, pra chave real nunca entrar por engano).
case "$STRIPE_KEY" in
  sk_test_*|rk_test_*) MODO=teste ;;
  sk_live_*|rk_live_*)
    if [ "${2:-}" != "--live" ]; then echo "Chave REAL sem --live. Nada foi criado."; exit 1; fi
    MODO=live ;;
  *) echo "Isso nao e uma chave da Stripe. Nada foi criado."; exit 1 ;;
esac
echo "Modo: $MODO"

API=https://api.stripe.com/v1
SAIDA="$(cd "$(dirname "$0")" && pwd)/stripe-links-$MODO.json"
# Pra onde ela vai depois de pagar. Cada link soma &plano=<tier>, que a pagina
# de obrigado usa pra contar a venda com o valor certo nos pixels.
OBRIGADO='https://plano.ciclofertil.com.br/quiz/v2/obrigado.html?session_id={CHECKOUT_SESSION_ID}'

stripe() { # stripe <caminho> <chave-de-idempotencia> [campos -d ...]
  local caminho=$1 idem=$2; shift 2
  curl -sS --fail-with-body "$API/$caminho" -u "$STRIPE_KEY:" \
    -H "Idempotency-Key: $idem" "$@"
}
campo() { python3 -c "import sys,json; print(json.load(sys.stdin)['$1'])"; }

echo '{' > "$SAIDA"
primeiro=1
while IFS='|' read -r plano semanas nome centavos; do
  produto=$(stripe products "cf-produto-$plano-$MODO-v1" \
    --data-urlencode "name=$nome" \
    --data-urlencode "description=Plano Ciclo Fértil de $semanas semanas. Acesso pelo app." \
    -d "metadata[plan]=$plano" \
    -d "metadata[semanas]=$semanas" | campo id)

  preco=$(stripe prices "cf-preco-$plano-$MODO-v1" \
    -d "product=$produto" \
    -d "currency=brl" \
    -d "unit_amount=$centavos" \
    -d "metadata[plan]=$plano" | campo id)

  link_json=$(stripe payment_links "cf-link-$plano-$MODO-v1" \
    -d "line_items[0][price]=$preco" \
    -d "line_items[0][quantity]=1" \
    -d "metadata[plan]=$plano" \
    -d "payment_intent_data[metadata][plan]=$plano" \
    -d "after_completion[type]=redirect" \
    --data-urlencode "after_completion[redirect][url]=$OBRIGADO&plano=$plano" \
    --data-urlencode "custom_text[submit][message]=Seu acesso ao app chega no seu e-mail logo depois da confirmação do pagamento.")
  link_id=$(echo "$link_json" | campo id)
  link_url=$(echo "$link_json" | campo url)

  [ $primeiro -eq 1 ] || echo ',' >> "$SAIDA"
  primeiro=0
  printf '  "%s": {"produto": "%s", "preco": "%s", "link_id": "%s", "url": "%s"}' \
    "$plano" "$produto" "$preco" "$link_id" "$link_url" >> "$SAIDA"
  echo "✓ $nome — $link_url"
done <<'PLANOS'
essencial|4|Teste 4 Semanas|4990
avancado|8|Plano de 8 Semanas|6990
imersao|12|Plano de 12 Semanas|9900
PLANOS
printf '\n}\n' >> "$SAIDA"

unset STRIPE_KEY
echo
echo "Pronto. Links salvos em: $SAIDA"
