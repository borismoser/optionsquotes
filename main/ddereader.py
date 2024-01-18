import ddeclient.dde_client as ddec

# topic
dde = ddec.DDEClient('profitchart', 'cot')
# item
tickers = ('PETR4', 'WEGE3')

for t in tickers:
    dde.advise(t)

for t in tickers:
    res = dde.request('PETR4')
    print(res)
