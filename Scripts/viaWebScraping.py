import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

chrome_driver_path = 'chromedriver.exe'

chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920,1080")

service = Service(chrome_driver_path)

driver = webdriver.Chrome(service=service, options=chrome_options)
wait = WebDriverWait(driver, 20)

arquivo = pd.read_excel('V01.xlsx')
empresas = arquivo[['Nome', 'CNPJ']]

url = 'https://www.econodata.com.br/consulta-empresa'

xpath_campo_busca = '//*[@id="input-busca-nome-cnpj-detalhe-empresa"]'
xpath_click_resultado = '//*[@id="tabela-resultado-busca-empresa"]/div/table/tbody/tr[1]/td[1]/div/div/div/div'
xpath_sem_resultado = '//*[@id="__nuxt"]/div/div[4]/div/h2'
xpath_cnpj = '//*[@id="__nuxt"]/div[2]/div[2]/div[1]/div/div/div[3]/div[2]/h1/span'
xpath_razao_social = '//*[@id="__nuxt"]/div[2]/div[2]/div[1]/div/div/div[3]/div[2]/h2'
xpath_local = '//*[@id="__nuxt"]/div[2]/div[2]/div[1]/div/div/div[3]/div[2]/div[2]'


# Buscar dados de empresa x
def buscar_dados_empresa(termo_busca):
    try:
        driver.get(url)
        campo_busca = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_campo_busca)))
        campo_busca.clear()
        campo_busca.send_keys(termo_busca)
        campo_busca.send_keys(Keys.RETURN)
        time.sleep(2)

        try:
            click_resultado = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_click_resultado)))
            click_resultado.click()

            cnpj = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, xpath_cnpj))
            ).text.strip()

            razao_social = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, xpath_razao_social))
            ).text.strip()

            local = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, xpath_local))
            ).text.strip()

            return cnpj, razao_social, local

        except Exception as e:
            try:
                sem_resultado = wait.until(EC.visibility_of_element_located((By.XPATH, xpath_sem_resultado)))
                if sem_resultado.is_displayed() and sem_resultado.text == "Nenhum resultado encontrado.":
                    return "n/a", "n/a", "n/a"
            except:
                raise e

    except Exception as e:
        print(f"Erro ao buscar dados para {termo_busca}: {str(e)}")
        return "error", "error", "error"

resultados = []

print("Iniciando busca...")

for index, row in empresas.iterrows():
    nome_empresa = row['Nome']
    cnpj_empresa = row['CNPJ']
    print(f"Buscando dados para: {nome_empresa} / {cnpj_empresa} ({index + 1}/{len(empresas)})")

    resultado_busca = buscar_dados_empresa(nome_empresa)
    if resultado_busca == ("n/a", "n/a", "n/a"):
        resultado_busca = buscar_dados_empresa(cnpj_empresa)
        if resultado_busca == ("n/a", "n/a", "n/a"):
            resultado_busca = ("n/a", "n/a", "n/a")

    resultados.append(resultado_busca)

print("Processo concluído. Os resultados foram salvos em 'resultados_cnpjs.xlsx'.")

df_resultados = pd.DataFrame(resultados, columns=['CNPJ', 'Razao Social', 'Local'])
df_resultados.to_excel('resultados_cnpjs.xlsx', index=False)

driver.quit()
