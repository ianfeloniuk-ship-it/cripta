# Cripta

Carteira de investimentos gratuita, com estética de terror. Funciona no navegador e pode ser instalada como app.

**Abrir:** https://ianfeloniuk-ship-it.github.io/cripta/

- Ações, FIIs, ETFs e BDRs da B3: cotações da [brapi.dev](https://brapi.dev) (detalhes e dividendos pedem um token grátis, colocado no Altar).
- Ações e ETFs dos EUA, dólar, euro e índices: `tools/cotacoes.py`, rodado pelo GitHub Actions a cada 20 min no pregão, grava `dados/cotacoes.json`.
- Criptomoedas: [CoinGecko](https://www.coingecko.com).
- Importa os extratos em Excel da Área do Investidor da B3 (Altar → Pacto com a B3).
- Os dados da carteira ficam só no navegador de quem usa (localStorage). Nada pessoal vai para este repositório.

Para acompanhar uma ação americana que não está na lista, adicione o código em `tools/acoes_eua.txt` (ETFs em `tools/etfs_eua.txt`).
