"""Busca cotações de ações e ETFs dos EUA, dólar, euro e índices (Yahoo Finance)
e grava dados/cotacoes.json. Roda no GitHub Actions (ver .github/workflows/cotacoes.yml),
onde não há bloqueio de CORS; o site lê o arquivo do próprio domínio."""
import json, time, urllib.request, urllib.parse, pathlib

RAIZ = pathlib.Path(__file__).resolve().parent.parent
UA = {'User-Agent': 'Mozilla/5.0 (cripta-cotacoes)'}
EXTRA = {'USDBRL=X': 'fx', 'EURBRL=X': 'fx', '^BVSP': 'idx', '^GSPC': 'idx', '^IXIC': 'idx', '^DJI': 'idx'}

def lista(nome):
    return [l.strip() for l in (RAIZ / 'tools' / nome).read_text().splitlines() if l.strip()]

def spark(simbolos):
    url = 'https://query2.finance.yahoo.com/v7/finance/spark?' + urllib.parse.urlencode(
        {'symbols': ','.join(simbolos), 'range': '5d', 'interval': '1d'})
    for tentativa in range(3):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=30) as r:
                return json.load(r)['spark']['result']
        except Exception as e:
            print('falhou', simbolos[0], e); time.sleep(3 * (tentativa + 1))
    return []

def main():
    tipos = {s: 'us' for s in lista('acoes_eua.txt')}
    tipos.update({s: 'us-etf' for s in lista('etfs_eua.txt')})
    tipos.update(EXTRA)
    simb = list(tipos)
    out = {'ts': int(time.time()), 'fx': {}, 'idx': {}, 'us': {}}
    for i in range(0, len(simb), 20):
        for r in spark(simb[i:i + 20]):
            try:
                resp = r['response'][0]; m = resp['meta']
                fech = [c for c in resp['indicators']['quote'][0]['close'] if c is not None]
                preco = m.get('regularMarketPrice') or (fech[-1] if fech else None)
                ant = fech[-2] if len(fech) >= 2 else m.get('chartPreviousClose')
                if preco is None: continue
                var = round((preco / ant - 1) * 100, 2) if ant else None
                s = r['symbol']; t = tipos.get(s)
                if t == 'fx': out['fx'][s[:3]] = [round(preco, 4), var]
                elif t == 'idx': out['idx'][s] = [round(preco, 2), var, m.get('shortName') or s]
                else: out['us'][s] = [round(preco, 4), var, m.get('longName') or m.get('shortName') or s, 1 if t == 'us-etf' else 0]
            except Exception as e:
                print('ignorado', r.get('symbol'), e)
        time.sleep(0.5)
    if not out['us'] or 'USD' not in out['fx']:
        raise SystemExit('sem dados; mantendo o arquivo anterior')
    (RAIZ / 'dados').mkdir(exist_ok=True)
    (RAIZ / 'dados' / 'cotacoes.json').write_text(json.dumps(out, ensure_ascii=False, separators=(',', ':')))
    print(len(out['us']), 'ativos; dólar', out['fx']['USD'])

if __name__ == '__main__':
    main()
