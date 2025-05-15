from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
import os

# Configurações do navegador
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920,1080")


servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico, options=chrome_options)


pagina = 'https://masander.github.io/AlimenticiaLTDA-financeiro/'
navegador.get(pagina)
time.sleep(3)

# Dicionario para armazenar os dados
tabela_despesas = {
    "codigo": [],
    "data_compra": [],
    "categoria": [],
    "departamento": [],
    "preco": [],
    "empresa": [],
}

# Espera pela tabela aparecer
try:
    WebDriverWait(navegador, 10).until(
        EC.presence_of_element_located((By.TAG_NAME, "table"))
    )
    print("Tabela localizada com sucesso!")
except TimeoutException:
    print("Falha ao carregar a tabela.")
    navegador.quit()

# Coleta os dados da tabela
linhas = navegador.find_elements(By.TAG_NAME, "tr")

for linha in linhas:
    try:
        codigo = linha.find_element(By.CLASS_NAME, "td_id_despesa").text.strip()
        data_compra = linha.find_element(By.CLASS_NAME, "td_data").text.strip()
        categoria = linha.find_element(By.CLASS_NAME, "td_tipo").text.strip()
        departamento = linha.find_element(By.CLASS_NAME, "td_setor").text.strip()
        preco = linha.find_element(By.CLASS_NAME, "td_valor").text.strip()
        empresa = linha.find_element(By.CLASS_NAME, "td_fornecedor").text.strip()

        print(f"{codigo} | {categoria} | {preco}")

        tabela_despesas["codigo"].append(codigo)
        tabela_despesas["data_compra"].append(data_compra)
        tabela_despesas["categoria"].append(categoria)
        tabela_despesas["departamento"].append(departamento)
        tabela_despesas["preco"].append(preco)
        tabela_despesas["empresa"].append(empresa)

    except Exception as erro:
        print("Erro na leitura da linha:", erro)

# Exporta os dados
df_final = pd.DataFrame(tabela_despesas)
df_final.to_excel("dados_extraidos/despesas_resultado.xlsx", index=False)

print("Arquivo Excel salvo com sucesso!")


navegador.quit()

