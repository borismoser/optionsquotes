import shutil
import sys
from datetime import datetime
from pathlib import Path
import pandas as pd
import pylightxl as xl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time
import os


def download_file():
    # Setup chrome options
    chrome_options = webdriver.ChromeOptions()

    # Set your download directory
    download_dir = os.getcwd()
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)
    prefs = {"download.default_directory": download_dir}
    chrome_options.add_experimental_option("prefs", prefs)

    # Setup WebDriver
    s = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=s, options=chrome_options)

    # Capture the start time
    start_time = time.time()

    try:
        # Open the webpage
        url_b3 = 'https://arquivos.b3.com.br/tabelas/InstrumentsConsolidated/' + datetime.now().strftime("%Y-%m-%d")
        driver.get(url_b3)
        #  'https://www.b3.com.br/en_us/market-data-and-indices/data-services/market-data/reports/daily-bulletin/data-on-exchange-listed-and-otc-assets-available-to-the-public/')

        # Find the download link and click it
        # download_link = driver.find_element(By.LINK_TEXT, 'Baixar arquivo completo')
        time.sleep(5)
        download_link = driver.find_element(By.PARTIAL_LINK_TEXT, 'complet')
        download_link.click()
        time.sleep(15)

        # Wait for the download to complete with a timeout
        timeout = 60  # seconds
        end_time = time.time() + timeout
        while time.time() < end_time:
            time.sleep(1)
            if not any(filename.endswith('.crdownload') for filename in os.listdir(download_dir)):
                break
        else:
            raise Exception("Download timed out.")

        # Get the name of the downloaded file
        downloaded_file = ""
        for file in os.listdir(download_dir):
            file_path = os.path.join(download_dir, file)
            if os.path.isfile(file_path) and os.path.getctime(file_path) > start_time:
                downloaded_file = file
                break

    finally:
        # Close the browser
        driver.quit()

    if not downloaded_file:
        raise Exception("Download failed or no new file found.")

    return downloaded_file


try:
    arq_series = download_file()
    print(f"Downloaded file: {arq_series}")
except Exception as e:
    print(f"Error: {e}")
    sys.exit()


data = pd.read_csv(arq_series, decimal=',',
                   low_memory=False, encoding='ANSI', delimiter=';', header=1)

# TODO Buscar parâmetros da planilha de cotações atual e usar os ativos no filtro das séries autorizadas (abaixo):

data = data[['TckrSymb', 'Asst', 'XprtnDt', 'OptnTp', 'ExrcPric', 'OptnStyle']][
    data['Asst'].isin(['BBDC4', 'BOVA11', 'PETR4', 'VALE3']) & data['SgmtNm'].isin(
        ['EQUITY CALL', 'EQUITY PUT'])]
data.to_csv('series_autorizadas_filtradas.csv', sep=';', encoding='ANSI', index=False)

db = xl.readcsv('series_autorizadas_filtradas.csv', delimiter=';', ws='Select')
ws = db.ws('Select')
xrows, xcols = ws.size

# Excel acusa erro ao tentar definir o endereço com a fórmula =RTD.
# Por isso, a fórmula está sendo colocada como texto, sendo necessário adicionar o = pelo Excel.
ws.update_address(address='G1', val='Last')
for r in range(2, xrows + 1):
    tck = ws.address(f'A{r}')
    ws.update_address(address=f'G{r}', val=f'RTD("rtdtrading.rtdserver";;"{tck}_B_0"; "ULT")')

bkp_file = ''
if Path('series_autorizadas_cotacoes.xlsx').exists():
    bkp_file = f'series_autorizadas_cotacoes_{datetime.now():%Y-%m-%d_%H-%M-%S}.xlsx'
    shutil.move('series_autorizadas_cotacoes.xlsx', bkp_file)

xl.writexl(db=db, fn='series_autorizadas_cotacoes.xlsx')
