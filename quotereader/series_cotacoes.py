import pandas as pd
from datetime import datetime


def txt_from_dataframe():
    with open('series_cotacoes.txt', 'w') as f:
        for k in tabs.keys():
            tabela = tabs[k]
            h = last_quote_asset[tabela[0][0]]
            f.write(f' {tabela[0][0]} : {h[0]:.2f} ({h[1]:.2f}%)\n\n')
            df = pd.DataFrame(tabela[1:], columns=tabela[0])
            txt = df.to_string(index=False).replace('(', ' ').replace(')', ' ')
            f.write(txt)
            f.write('\n\n\n')


def txt_from_tabs():
    with open('series_cotacoes.txt', 'w') as f:
        f.write(f'Arquivo gerado em {datetime.date(datetime.now())} às {datetime.time(datetime.now())}\n\n')
        for asset in tabs.keys():
            tabela = tabs[asset]
            # Código do ativo, última cotação e variação
            h = last_quote_asset[tabela[0][0]]
            lastq = h[0]
            lin = f'{asset:>6} : {lastq:.2f} ({h[1]:.2f}%)'
            f.write(lin + '\n\n')
            # Código do ativo e datas de vencimentos de Call e Put
            lin = f'{tabela[0][0]:>6}' + ' |' + '|'.join(f'{str(item)}'.center(14) for item in tabela[0][1:])
            f.write(lin + '\n')
            indicador = False
            for x in tabela[1:]:
                # Marcador da última cotação do ativo
                if not indicador and float(x[0]) >= round(lastq,2):
                    lin = f'{lastq:6.2f} ' + ('.' * (len(lin)-7))
                    # lin += '.' * (len(lin)-8)
                    f.write(lin + '\n')
                    indicador = True
                # Strike da linha atual
                lin = f'{x[0]:6.2f} '
                # Códigos e últimas cotações das opções do strike atual
                barra = True
                for item in x[1:]:
                    if barra:
                        sep = '|'
                    else:
                        sep = '.'
                    barra = not barra
                    if item:
                        op = item[0]
                        cot = item[1]
                        lin += sep + f'{op:6} {cot:5.2f}'.center(14)
                    else:
                        lin += sep + ''.center(14)
                f.write(lin + '\n')
            f.write('\n')


def generate_datasets():
    with pd.ExcelFile('series_autorizadas_cotacoes.xlsx') as xlsx:
        quotes = pd.read_excel(xlsx, 'Select')
        params = pd.read_excel(xlsx, 'Params')

    quotes = quotes[quotes.Last > 0]
    quotes['colHead'] = quotes.XprtnDt.astype(str) + ' ' + quotes.OptnTp.str[0]
    quotes.OptnTp = pd.Categorical(quotes.OptnTp)
    quotes.OptnStyle = pd.Categorical(quotes.OptnStyle)
    quotes.Asst = pd.Categorical(quotes.Asst)
    quotes.XprtnDt = pd.Categorical(quotes.XprtnDt, categories=sorted(quotes.XprtnDt.unique()), ordered=True)
    quotes.colHead = pd.Categorical(quotes.colHead, categories=sorted(quotes.colHead.unique()), ordered=True)

    # last_quote_asset = {}
    # tabs = {}
    for parm in params.itertuples():
        last_quote_asset[parm.Asset] = (parm.Last, parm.Var)
        df = quotes[(quotes.Asst == parm.Asset) & (quotes.ExrcPric >= parm.FromStrike) & (
                quotes.ExrcPric <= parm.ToStrike) & (quotes.XprtnDt.astype(str) >= parm.FromDate) & (
                            quotes.XprtnDt.astype(str) <= parm.ToDate)]
        vencimentos = sorted(df.colHead.unique())
        strikes = sorted(df.ExrcPric.unique())
        tabela = [[parm.Asset] + vencimentos]
        for st in strikes:
            linha = [st]
            for i in range(1, len(tabela[0])):
                op = df[['TckrSymb', 'Last']][(df.ExrcPric == st) & (df.colHead == tabela[0][i])].values
                if op.any():
                    tpl = tuple([op[0][0][4:], round(op[0][1], 2)])
                    linha.append(tpl)
                else:
                    linha.append('')
            tabela.append(linha)
        tabs[parm.Asset] = tabela


last_quote_asset = {}
tabs = {}
print('Carregando cotações...')
generate_datasets()
print('Gerando arquivo txt...')
txt_from_tabs()
print('Pronto.')
