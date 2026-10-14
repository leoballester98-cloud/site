#!/usr/bin/env bash
# Guarda no Supabase a chave que a funcao criar-checkout usa pra abrir o
# checkout embutido. Use uma chave RESTRITA (rk_...) so com "Checkout
# Sessions: Write". O modo sai da propria chave: rk_test_/sk_test_ vira
# STRIPE_SECRET_KEY_TESTE; rk_live_/sk_live_ vira STRIPE_SECRET_KEY_LIVE.
#
# Espera voce copiar a chave, le da area de transferencia e apaga de la;
# vai pro Supabase por um arquivo temporario que some no fim.
set -euo pipefail

PROJETO=lfmjvgkbvkkexwytkfeg
read -rp "Copie a chave restrita (rk_live_... ou rk_test_...) na Stripe e aperte Enter aqui: " _ || true
CHAVE=$(pbpaste | tr -d '[:space:]')
printf '' | pbcopy
case "$CHAVE" in
  rk_test_*|sk_test_*) VAR=STRIPE_SECRET_KEY_TESTE ;;
  rk_live_*|sk_live_*) VAR=STRIPE_SECRET_KEY_LIVE ;;
  "") echo "A area de transferencia estava vazia. Nada foi alterado."; exit 1 ;;
  pk_*) echo "Isso e a chave publicavel (pk_). Precisa da restrita (rk_). Nada foi alterado."; exit 1 ;;
  *) echo "O que esta copiado nao e uma chave rk_/sk_ (tem ${#CHAVE} caracteres). Nada foi alterado."; exit 1 ;;
esac

tmp=$(mktemp); chmod 600 "$tmp"
trap 'rm -f "$tmp"' EXIT
printf '%s=%s\n' "$VAR" "$CHAVE" > "$tmp"
unset CHAVE
cd /Users/leoballester/projetos/ciclofertil-app
npx --yes supabase secrets set --env-file "$tmp" --project-ref "$PROJETO"
echo
echo "✓ Chave guardada no Supabase como $VAR"
