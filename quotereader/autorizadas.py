import pandas as pd
import pylightxl as xl

data = pd.read_csv("series_autorizadas.csv", low_memory=False, encoding='ANSI', delimiter=';', header=1)

data2 = data[['TckrSymb', 'Asst', 'XprtnDt', 'OptnTp', 'ExrcPric', 'OptnStyle']][
    ((data['Asst'] == 'BBDC4') | (data['Asst'] == 'BOVA11') | (data['Asst'] == 'PETR4')) & (
            (data['SgmtNm'] == 'EQUITY CALL') | (data['SgmtNm'] == 'EQUITY PUT'))]
data2.to_csv('series_autorizadas_filtradas.csv', sep=';', encoding='ANSI', index=False)

db = xl.readcsv('series_autorizadas_filtradas.csv', delimiter=';', ws='Select')
ws = db.ws('Select')
xrows, xcols = ws.size

ws.update_address(address='G1', val='Last')
for r in range(2, xrows + 1):
    addr = f'E{r}'
    #  Strike está chegando com decimal separada por vírgula
    prc = ws.address(address=addr)
    ws.update_address(address=addr, val=float(str(prc).replace(',', '.')))
    # Excel acusa erro ao tentar definir o endereço com a fórmula =RTD
    tck = ws.address(f'A{r}')
    ws.update_address(address=f'G{r}', val=f'RTD("rtdtrading.rtdserver";;"{tck}_B_0"; "ULT")')
xl.writexl(db=db, fn='series_autorizadas_cotacoes.xlsx')
