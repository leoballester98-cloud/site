# Plano de Testes A/B — Página de Vendas (Ciclo Fértil)

> Documento de planejamento dos testes de página. Objetivo: descobrir a combinação
> de **hero + mídia + headline** que mais engaja e converte o tráfego real, de forma
> estruturada, isolando uma variável por vez. **Tudo em cima da /v4/.**

---

## Regras do jogo (valem pra todos os testes)

- **Uma variável por teste.** Nunca rodar 2 testes no mesmo tráfego ao mesmo tempo.
- **Sequencial:** só começa a próxima fase quando a anterior tiver vencedor.
- **Isolamento:** nas variantes de página, muda **só** a variável testada; o resto é idêntico.
- **Mesmo criativo + mesmo público** entre as variantes de página — só a **URL** muda.
- **Mecânica no Meta:** ferramenta "Teste A/B" (ou ad sets duplicados apontando URLs
  diferentes), budget **50/50**.

### Métricas de decisão
- **Decisão final:** custo por compra (CPA) + taxa de conversão (CVR).
- **Sinal antecipado (mais rápido, em poucos dias):** **% de sessões ≥ 30s** no Clarity
  (o "engajado real" — filtro `Duração da sessão ≥ 0.5 minutos`). Serve pra ler tendência
  antes de acumular vendas suficientes.

### Realidade de volume (importante)
Com ~200 sessões reais / 3 dias e ~1-2 vendas/dia, atingir significância em **CVR**
por variante é **lento**. Por isso:
- Usar o **% engajado (≥30s)** como leitura rápida entre as fases.
- Deixar cada fase rodar **tempo/budget suficiente** — não decidir no dia 1.
- Testar **poucas variantes por vez** (2, no máximo 3) pra não diluir o tráfego.

---

## As fases

### Fase 0 — Concluir os criativos (em andamento)
- **Onde:** todos os criativos novos → `/v4/`.
- **O que fazer:** separar em **vencedores** (converteram) e **inconclusivos** (gastaram
  pouco). Os inconclusivos ganham **ad set isolado** com budget próprio pra concluir.
- **Saída:** lista de criativos vencedores → alimentam as próximas fases.

### Fase 1 — Teste de HERO (página)
- **Isola:** o hero (só ele muda; vídeo mantido nas duas).
- **Criativos:** só os **vencedores** da Fase 0.
- **Controle:** `/v4/` (hero atual)
- **Variante:** `/v4-testehero/` (hero do "sinal")
- **Decide:** qual hero engaja/converte mais.

### Fase 2 — Teste de MÍDIA (vídeo vs imagem)
- **Isola:** a mídia. Roda no **hero vencedor** da Fase 1.
- **Com vídeo:** `/[vencedor]/` (já existe)
- **Só imagem:** `/[vencedor]-img/` (nova página, sem vídeo)
  - ex.: `/v4-testehero/` vs `/v4-testehero-img/`  (ou `/v4/` vs `/v4-img/`)
- **Decide:** vídeo segura mais, ou imagem estática (mais leve/calma)?

### Fase 3 — Teste de 5 HEROS novos
- **Isola:** headline/hero. Roda na **mídia vencedora** da Fase 2.
- **Variantes:** `/v4-hero-a/` · `/v4-hero-b/` · `/v4-hero-c/` · `/v4-hero-d/` · `/v4-hero-e/`
- **Decide:** melhor headline. (Com pouco volume: roda 2-3 por vez, elimina os piores,
  sobe os próximos.)

---

## Todas as URLs

| Fase | Papel | URL | Status |
|------|-------|-----|--------|
| 0 | Criativos | `protocolociclofertil.com.br/v4/` | ✅ no ar |
| 1 | Controle (hero atual) | `protocolociclofertil.com.br/v4/` | ✅ no ar |
| 1 | Variante (hero do sinal) | `protocolociclofertil.com.br/v4-testehero/` | ✅ criada |
| 2 | Mídia — só imagem | `protocolociclofertil.com.br/v4-testehero-img/` | ✅ criada |
| 3 | Hero novo A | `protocolociclofertil.com.br/v4-hero-a/` | ✅ criada |
| 3 | Hero novo B | `protocolociclofertil.com.br/v4-hero-b/` | ✅ criada |
| 3 | Hero novo C | `protocolociclofertil.com.br/v4-hero-c/` | ✅ criada |
| 3 | Hero novo D | `protocolociclofertil.com.br/v4-hero-d/` | ✅ criada |
| 3 | Hero novo E | `protocolociclofertil.com.br/v4-hero-e/` | ✅ criada |

> Observação: a URL da Fase 2 (`-img`) depende do vencedor da Fase 1. Se a `/v4/` (atual)
> vencer, a variante de imagem vira `/v4-img/`.

---

## Fluxo resumido

```
Fase 0: criativos → /v4/  →  [vencedores]
   │
Fase 1: HERO      → /v4/  vs  /v4-testehero/          →  [hero vencedor]
   │
Fase 2: MÍDIA     → [vencedor]/  vs  [vencedor]-img/  →  [mídia vencedora]
   │
Fase 3: HEROS     → /v4-hero-a..e/                    →  [headline vencedora]
```

---

## Referências de método
- **Hopkins (Propaganda Científica):** medir tudo, específico > genérico, testar sempre.
- **Evaldo (Carta de 16 palavras):** cada hero deve sustentar o One Belief (ler os sinais
  + blindar hormônios na semana certa → engravidar).
- **Schwartz (Breakthrough):** entrar pela emoção onde a leitora está; mecanismo revelado
  descendo a página.
- **Scott Adams:** óbvio na 1ª olhada; hook entendido sozinho, mudo, em 1 segundo.

---

*Última atualização: 12/07/2026*
