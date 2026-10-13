#!/usr/bin/env bash
# Guarda a chave do Resend (re_...) nos secrets do Supabase, junto com o
# remetente. A chave vem da area de transferencia (copie no Resend antes de
# rodar) e e apagada dela em seguida; vai pro Supabase por um arquivo
# temporario que some no fim. Precisa do login no Supabase (npx supabase login).
set -euo pipefail

PROJETO=lfmjvgkbvkkexwytkfeg
REMETENTE='Ciclo Fértil <ola@ciclofertil.com.br>'

CHAVE=$(pbpaste | tr -d '[:space:]')
printf '' | pbcopy
case "$CHAVE" in
  re_*) ;;
  *) echo "O que esta copiado nao e uma chave do Resend (re_...). Nada foi alterado."; exit 1 ;;
esac

tmp=$(mktemp); chmod 600 "$tmp"
trap 'rm -f "$tmp"' EXIT
printf 'RESEND_API_KEY=%s\nRESEND_FROM=%s\n' "$CHAVE" "$REMETENTE" > "$tmp"
unset CHAVE
npx --yes supabase secrets set --env-file "$tmp" --project-ref "$PROJETO"

echo
echo "✓ RESEND_API_KEY guardada no Supabase"
echo "✓ Remetente: $REMETENTE"
