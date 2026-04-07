# ORPEC Players Dashboard 2026

Dashboard de performance digital multicanal para os players ORPEC — Jan a Mar 2026.

> **Acesso:** [https://organizationivvvv.github.io/orpec-dashboard/](https://organizationivvvv.github.io/orpec-dashboard/)

## O que tem no dashboard

- **KPIs** — Top Instagram, Top Facebook, melhor posição geral, % performance
- **Posição ORPEC por plataforma** — Instagram, Facebook, LinkedIn, YouTube, SEO, PageSpeed, GMN
- **Gráficos por plataforma** — Instagram, Facebook, LinkedIn, YouTube, SEO, PageSpeed
- **Rankings** — Touchpoints digitais e performance ponderada por canal
- **Tabelas de dados brutos** — Todos os dados tabulados por plataforma, com ORPEC destacado

## Dados

- 6 players avaliados: ORPEC, Orguel, Mills, Pashal, SH, Atex
- Período: Janeiro, Fevereiro e Março 2026
- Fontes: planilha `Resumo ORPEC - PLAYERS 2026`

## Como atualizar

1. Edite o `index.html` com os dados novos
2. Commit + push:

```bash
git add index.html
git commit -m "update data"
git push
```

O deploy roda automaticamente via GitHub Pages (~1 min).

## Stack

- HTML/CSS/JS vanilla
- Chart.js (Chart.min.js local)
- GitHub Pages para hosting
- Fontes: Poppins + Inter (Google Fonts)

## Estrutura

```
orpec-dashboard/
├── index.html      # Dashboard completo (dados + visuais)
├── chart.min.js    # Chart.js (local, sem dependência de CDN)
├── .github/        # Workflow de deploy automatico
└── README.md       # Este arquivo
```
