import pandas as pd

with pd.ExcelFile('series_autorizadas_cotacoes.xlsx') as xlsx:
    quotes = pd.read_excel(xlsx, 'Select')
    params = pd.read_excel(xlsx, 'Params')

quotes = quotes[quotes.Last > 0]
quotes['colHead'] = quotes.XprtnDt.astype(str) + ' ' + quotes.OptnTp.str[0]

last_quote_asset = {}
tabs = {}
for parm in params.itertuples():
    last_quote_asset[parm.Asset] = (parm.Last, parm.Var)
    df = quotes[(quotes.Asst == parm.Asset) & (quotes.ExrcPric >= parm.FromStrike) & (
            quotes.ExrcPric <= parm.ToStrike) & (quotes.XprtnDt >= parm.FromDate) & (quotes.XprtnDt <= parm.ToDate)]
    vencimentos = sorted(df.colHead.unique())
    strikes = sorted(df.ExrcPric.unique())
    # vencimentos = sorted(quotes.colHead[(quotes.Asst == parm.Asset) & (quotes.ExrcPric >= parm.FromStrike) & (
    #         quotes.ExrcPric <= parm.ToStrike) & (quotes.XprtnDt >= parm.FromDate) & (
    #                                             quotes.XprtnDt <= parm.ToDate)].unique())
    # strikes = sorted(quotes.ExrcPric[(quotes.Asst == parm.Asset) & (quotes.ExrcPric >= parm.FromStrike) & (
    #         quotes.ExrcPric <= parm.ToStrike) & (quotes.XprtnDt >= parm.FromDate) & (
    #                                          quotes.XprtnDt <= parm.ToDate)].unique())
    tabela = [[parm.Asset] + vencimentos]
    for st in strikes:
        linha = [st]
        for i in range(1, len(tabela[0])):
            op = df[['TckrSymb', 'Last']][(df.ExrcPric == st) & (df.colHead == tabela[0][i])].values
            # op = quotes[['TckrSymb', 'Last']][
            #     (quotes.Asst == parm.Asset) & (quotes.ExrcPric == st) & (
            #             quotes.colHead == tabela[0][i])].values
            if op.any():
                tpl = tuple([op[0][0][4:], round(op[0][1], 2)])
                linha.append(tpl)
            else:
                linha.append('')
        tabela.append(linha)
    tabs[parm.Asset] = tabela

with open('texte.txt', 'w') as f:
    txt = ''
    for k in tabs.keys():
        tabela = tabs[k]
        h = last_quote_asset[tabela[0][0]]
        f.write(f' {tabela[0][0]} : {h[0]:.2f} ({h[1]:.2f}%)\n\n')
        df = pd.DataFrame(tabela[1:], columns=tabela[0])
        txt = df.to_string(index=False).replace('(', ' ').replace(')', ' ')
        f.write(txt)
        f.write('\n\n\n')

# with open('series_cotacoes.txt', 'w') as f:
#     for asset in tabs.keys():
#         header = [f'{asset:>6}'] + sorted(expirations.get(asset) + expirations.get(asset))
#         tabela = []
#         for stk in sorted(strikes.get(asset)):
#             lin = [stk]
#             for _ in range(len(header) - 1):
#                 lin.append('')
#             tabela.append(lin)
#
#         options = assets.get(asset)
#         options.sort(key=lambda o: getattr(o, 'sortkey'))
#         for op in options:
#             item = f'{op.symbol:6} {op.last:5.2f}'
#             for r in tabela:
#                 if r[0] == op.strike:
#                     if op.expiration in header:
#                         pos = header.index(op.expiration)
#                         if op.type == 'Put':
#                             pos += 1
#                         r[pos] = item
#                     else:
#                         raise Exception("Vencimento não encontrado!")
#                     break
#
#         # OUTPUT
#         param = params.get(asset)
#         lin = f'{asset:>6} : {param.last_quote:.2f} ({param.variation:.2f}%)'
#         f.write(lin + '\n\n')
#         lin = header[0] + ' |' + '|'.join(f'{str(item)}'.center(14) for item in header[1:])
#         f.write(lin + '\n')
#         indicador = False
#         for x in tabela:
#             if not indicador and float(x[0]) > param.last_quote:
#                 f.write('........\n')
#                 indicador = True
#             lin = f'{x[0]:6.2f} |'
#             lin = lin + '|'.join(f'{str(item)}'.center(14) for item in x[1:])
#             f.write(lin + '\n')
#         f.write('\n')
