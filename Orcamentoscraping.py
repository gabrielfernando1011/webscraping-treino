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


chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920,1080")

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico, options=chrome_options)

pagina = 'https://masander.github.io/AlimenticiaLTDA-financeiro/'
navegador.get(pagina)
time.sleep(3)


tabela_orcamentos = {
    "setor": [],
    "mes": [],
    "ano": [],
    "valor_previsto": [],
    "valor_realizado": [],
}


try:
    WebDriverWait(navegador, 10).until(
        EC.presence_of_element_located((By.TAG_NAME, "table"))
    )
    print("Tabela localizada com sucesso!")
except TimeoutException:
    print("Falha ao carregar a tabela.")
    navegador.quit()


linhas = navegador.find_elements(By.TAG_NAME, "tr")

for linha in linhas:
    try:
        setor = linha.find_element(By.CLASS_NAME, "td_setor").text.strip()
        mes = linha.find_element(By.CLASS_NAME, "td_mes").text.strip()
        ano = linha.find_element(By.CLASS_NAME, "td_ano").text.strip()
        valor_previsto = linha.find_element(By.CLASS_NAME, "td_valor_previsto").text.strip()
        valor_realizado = linha.find_element(By.CLASS_NAME, "td_valor_realizado").text.strip()

        print(f"{setor} - {mes} - {valor_previsto}")

        tabela_orcamentos["setor"].append(setor)
        tabela_orcamentos['mes'].append(mes)
        tabela_orcamentos['ano'].append(ano)
        tabela_orcamentos['valor_previsto'].append(valor_previsto)
        tabela_orcamentos['valor_realizado'].append(valor_realizado)
    except Exception as erro:
        print("Erro na leitura da linha:", erro)

orcamentos_adicionais = {
    "setor": [],
    "mes": [],
    "ano": [],
    "valor_previsto": [],
    "valor_realizado": [],
}


try:
    proxima_tabela = WebDriverWait(navegador, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//button[normalize-space(text())='Orçamentos']"))
    )
    navegador.execute_script("arguments[0].click();", proxima_tabela)
    proxima_tabela.click()

except Exception as e:
    print('Erro ao tentar avançar para a próxima página', e)


try:
    WebDriverWait(navegador, 10).until(
        EC.presence_of_all_elements_located((By.TAG_NAME, "tr"))
    )
    print('Elementos encontrados com sucesso')
except TimeoutException:
    print('Tempo de espera foi excedido, não foi encontrado nada')


orcamento_linhas = navegador.find_elements(By.TAG_NAME, "tr")

for linha in orcamento_linhas:
    try:
        setor = linha.find_element(By.CLASS_NAME, "td_setor").text.strip()
        mes = linha.find_element(By.CLASS_NAME, "td_mes").text.strip()
        ano = linha.find_element(By.CLASS_NAME, "td_ano").text.strip()
        valor_previsto = linha.find_element(By.CLASS_NAME, "td_valor_previsto").text.strip()
        valor_realizado = linha.find_element(By.CLASS_NAME, "td_valor_realizado").text.strip()

        print(f"{setor} - {mes} - {valor_previsto}")

        orcamentos_adicionais["setor"].append(setor)
        orcamentos_adicionais["mes"].append(mes)
        orcamentos_adicionais["ano"].append(ano)
        orcamentos_adicionais["valor_previsto"].append(valor_previsto)
        orcamentos_adicionais["valor_realizado"].append(valor_realizado)

    except Exception as erro:
        print('Erro ao coletar dados adicionais: ', erro)


navegador.quit()

df_orcamento = pd.DataFrame(tabela_orcamentos)
df_orcamento_adicional = pd.DataFrame(orcamentos_adicionais)


df_final = pd.concat([df_orcamento, df_orcamento_adicional], axis=0, ignore_index=True)


os.makedirs("dados_extraidos", exist_ok=True)
df_final.to_excel("dados_extraidos/orcamentos_completos.xlsx", index=False)

print("Arquivo Excel de orçamentos salvo com sucesso!")
