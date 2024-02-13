from datetime import datetime
from pathlib import Path
import pandas as pd
import pylightxl as xl
import shutil


data = pd.read_csv("series_autorizadas.csv", decimal=',',
                   low_memory=False, encoding='ANSI', delimiter=';', header=1)
data2 = data[['TckrSymb', 'Asst', 'XprtnDt', 'OptnTp', 'ExrcPric', 'OptnStyle']][
    data['Asst'].isin(['BBDC4', 'BOVA11', 'PETR4']) & data['SgmtNm'].isin(['EQUITY CALL', 'EQUITY PUT'])]
data2.to_csv('series_autorizadas_filtradas.csv', sep=';', encoding='ANSI', index=False)

db = xl.readcsv('series_autorizadas_filtradas.csv', delimiter=';', ws='Select')
ws = db.ws('Select')
xrows, xcols = ws.size

# Excel acusa erro ao tentar definir o endereço com a fórmula =RTD.
# Por isso, a fórmula está sendo colocada como texto, sendo necessário adicionar o = pelo Excel.
ws.update_address(address='G1', val='Last')
for r in range(2, xrows + 1):
    tck = ws.address(f'A{r}')
    ws.update_address(address=f'G{r}', val=f'RTD("rtdtrading.rtdserver";;"{tck}_B_0"; "ULT")')

if Path('series_autorizadas_cotacoes.xlsx').exists():
    shutil.copy('series_autorizadas_cotacoes.xlsx',
                f'series_autorizadas_cotacoes_{datetime.now():%Y-%m-%d_%H-%M-%S}.xlsx')

# xl.writexl(db=db, fn='series_autorizadas_cotacoes.xlsx')
