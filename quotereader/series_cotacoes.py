import pylightxl as xl
from dataclasses import dataclass


@dataclass
class Option:
    symbol: str
    type: str
    expiration: str
    strike: float
    last: float
    sortkey: str


@dataclass
class Params:
    last_quote: float
    variation: float
    from_strike: float
    to_strike: float
    from_date: str
    to_date: str


SYMBOL = 0
ASSET = 1
EXPDATE = 2
TYPE = 3
STRIKE = 4
STYLE = 5
LAST = 6

SHOW_LAST_QUOTE_ZERO = False
arquivo = "series_autorizadas_cotacoes.xlsx"
wb = xl.readxl(arquivo)

params = {}
first = True
for row in wb.ws('Params').rows:
    if first:
        first = False
        continue
    param = Params(
        last_quote=round(row[1], 2),
        variation=round(row[2], 2),
        from_strike=row[3],
        to_strike=row[4],
        from_date=row[5],
        to_date=row[6])
    params[row[0]] = param

ws = wb.ws('Select')
xrows, xcols = ws.size
assets = {}
expirations = {}
strikes = {}

for r in range(2, xrows+1):
    row = ws.row(r)
    asset = row[ASSET]
    lastq = round(row[LAST], 2)
    if (not SHOW_LAST_QUOTE_ZERO) and lastq == 0:
        continue

    st = row[STRIKE]
    ex = row[EXPDATE]
    param = params.get(asset)
    if (st < param.from_strike) or (st > param.to_strike) or (ex < param.from_date) or (ex > param.to_date):
        continue

    option = Option(
        symbol=row[SYMBOL][4:],
        type=row[TYPE],
        expiration=ex,
        strike=st,
        last=lastq,
        sortkey=f'{st:07.2f}{row[TYPE][0]}'
    )

    options = assets.get(asset, [])
    options.append(option)
    assets[asset] = options

    exp = expirations.get(asset, [])
    if option.expiration not in exp:
        exp.append(option.expiration)
    expirations[asset] = exp

    stk = strikes.get(asset, [])
    if option.strike not in stk:
        stk.append(option.strike)
    strikes[asset] = stk

with open('series_cotacoes.txt', 'w') as f:
    for asset in assets.keys():
        header = [f'{asset:>6}'] + sorted(expirations.get(asset) + expirations.get(asset))
        tabela = []
        for stk in sorted(strikes.get(asset)):
            lin = [stk]
            for _ in range(len(header) - 1):
                lin.append('')
            tabela.append(lin)

        options = assets.get(asset)
        options.sort(key=lambda o: getattr(o, 'sortkey'))
        for op in options:
            item = f'{op.symbol:6} {op.last:5.2f}'
            for r in tabela:
                if r[0] == op.strike:
                    if op.expiration in header:
                        pos = header.index(op.expiration)
                        if op.type == 'Put':
                            pos += 1
                        r[pos] = item
                    else:
                        raise Exception("Vencimento não encontrado!")
                    break

        # OUTPUT
        param = params.get(asset)
        lin = f'{asset:>6} : {param.last_quote:.2f} ({param.variation:.2f}%)'
        f.write(lin + '\n\n')
        lin = header[0] + ' |' + '|'.join(f'{str(item)}'.center(14) for item in header[1:])
        f.write(lin + '\n')
        indicador = False
        for x in tabela:
            if not indicador and float(x[0]) > param.last_quote:
                f.write('........\n')
                indicador = True
            lin = f'{x[0]:6.2f} |'
            lin = lin + '|'.join(f'{str(item)}'.center(14) for item in x[1:])
            f.write(lin + '\n')
        f.write('\n')

# test
