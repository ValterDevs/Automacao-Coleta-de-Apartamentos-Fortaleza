import sys, os, keyboard, json
import requests
from bs4 import BeautifulSoup
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from oauth2client.service_account import ServiceAccountCredentials
import undetected_chromedriver as uc
import gspread
from gspread.http_client import BackOffHTTPClient

# Função para atualizar a célula com tentativas de reconexão
def atualizar_celula(i, status_text):
    planilha = sheet
    tentativas = 3
    for tentativa in range(tentativas):
        try:
            # Atualize a célula
            planilha.update_cell(i, 2, status_text)
            print(f"Célula na linha {i} atualizada com sucesso.")
            break
        except Exception as e:
            print(f"Erro na tentativa {tentativa + 1}: {e}")
            if tentativa < tentativas - 1:
                print("Tentando novamente em 5 segundos...")
                time.sleep(5)
            else:
                print("Falha ao atualizar a célula após várias tentativas.")
                raise

# FUNÇÃO DE LOCALIZAR O XPATH E FAZER SCROLL
def localizar(driver, xpath, timeout=5):
    wait = WebDriverWait(driver, timeout)
    try:
        # Aguarda até que o elemento esteja presente no DOM
        elemento = wait.until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        
        # Faz scroll até o elemento ficar visível
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
        
        # Aguarda até estar visível e clicável
        elemento = wait.until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        
        return elemento
    except Exception as e:
        print(f"Erro ao localizar o elemento '{xpath}': {e}\n")
        return None

# Defina os escopos de acesso
escopos = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# Carregue as credenciais do arquivo JSON
credenciais = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scopes=escopos)

# Autentique-se e crie o cliente com BackoffClient
cliente = gspread.authorize(credenciais, http_client=BackOffHTTPClient)

# Abra a planilha pelo ID
sheet = cliente.open("Apartamentos-Fortaleza-CE").sheet1

# Obtenha todos os valores da planilha
data = sheet.get_all_values()

# Pega só a coluna B
col_A = sheet.col_values(1)  # 2 porque col_values começa em 1

primeira_linha_vazia = len(col_A) + 1  # próxima linha disponível

print("Linha em que será inserido novos dados:", primeira_linha_vazia)



driver = uc.Chrome()
driver.maximize_window()

# remover a flag "webdriver" do navigator
driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    "source": """
    Object.defineProperty(navigator, 'webdriver', {
      get: () => undefined
    })
    """
})

driver.get('https://www.zapimoveis.com.br/venda/apartamentos/ce+fortaleza/?transacao=venda&onde=%2CCear%C3%A1%2CFortaleza%2C%2C%2C%2C%2Ccity%2CBR%3ECeara%3ENULL%3EFortaleza%2C-3.73272%2C-38.527013%2C&tipos=apartamento_residencial&origem=busca-recente')
time.sleep(2.5)


for i in range(10):
    for a in range(1,31):


        localizar(driver, f'/html/body/section/div/div[3]/div[4]/div[1]/ul/li[{a}]/a/div/div[2]/div[1]/div/h2').click()

        try:
            localizar(driver, '/html/body/div[7]/section/section/section[2]/a[1]/div[2]/div[1]/img[1]').click
        except:
            pass


        variavel_reserva = ""

        abas = driver.window_handles

        driver.switch_to.window(abas[-1])

        url = driver.current_url

        time.sleep(2)

        try:
            valor = localizar(driver, '/html/body/div[3]/div[1]/div[1]/div[1]/div[3]/div/div[2]/div[1]/p').text
        except:
            valor = "Valor não informado"

        try:
            valor_cond = localizar(driver, '/html/body/div[3]/div[1]/div[1]/div[1]/div[3]/div/div[2]/div[2]/p').text
        except:
            valor_cond = "Valor do condomínio não informado"

        try:
            valor_iptu = localizar(driver, '/html/body/div[3]/div[1]/div[1]/div[1]/div[3]/div/div[2]/div[3]/p').text
        except:
            valor_iptu = "Valor do IPTU não informado"

        try:
            metros = localizar(driver, '/html/body/div[3]/div[1]/div[1]/div[1]/div[4]/div/div/div/div/ul/li[1]/span[2]').text
        except:
            metros = "Tamanho não informado"

        try:
            garagem = localizar(driver, '/html/body/div[3]/div[1]/div[1]/div[1]/div[4]/div/div/div/div/ul/li[4]/span[2]').text

        except:
            garagem = "Garagem não informada"

        try:
            quartos = localizar(driver, '/html/body/div[3]/div[1]/div[1]/div[1]/div[4]/div/div/div/div/ul/li[2]/span[2]').text
        except:
            quartos = "Quartos não informados"

        try:
            banheiros = localizar(driver, '/html/body/div[3]/div[1]/div[1]/div[1]/div[4]/div/div/div/div/ul/li[3]/span[2]').text
        except:
            banheiros = "Quantidade de banheiros não informado"

        try:
            andar = localizar(driver, '/html/body/div[3]/div[1]/div[1]/div[1]/div[4]/div/div/div/div/ul/li[6]/span[2]').text
        except:
            andar = "Andar nao informado"

        try:
            suite = localizar(driver, '/html/body/div[3]/div[1]/div[1]/div[1]/div[4]/div/div/div/div/ul/li[5]/span[2]').text
            if("andar" in suite):
                variavel_reserva = andar
                andar = suite
                suite = variavel_reserva
        except:
            suite = "Suítes não informadas"
        
        try:
            loc = localizar(driver, '/html/body/div[3]/div[1]/div[1]/section[1]/div[1]/p').text
        except:
            loc = "Localização não informada"

        if("andar" not in andar):
            andar = "Andar não informado"
        
        if("suítes" in garagem or "suíte" in garagem):
            variavel_reserva = garagem
            suite = variavel_reserva
            variavel_reserva = banheiros
            garagem = banheiros
            garagem = variavel_reserva

        try:
            anunciante = driver.find_element(By.XPATH, '/html/body/div[3]/div[2]/div[1]/div/div/div/div[1]/div[1]/div/div[2]/div/p').text
        except:
            anunciante = "Anunciante não encontrado"

        print(f'Valor: {valor} \n valor condominio: {valor_cond} \n valor iptu {valor_iptu} \n metros: {metros} \n garagem: {garagem} \n quartos: {quartos} \n banheiros: {banheiros} \n andar: {andar} \n suites: {suite}')
        print('fim do programa')


        dados = {
            "localizacao": loc,
            "andar": andar,
            "banheiros": banheiros,
            "quartos": quartos,
            "suites": suite,
            "valor": valor,
            "metro_quadrado": metros,
            "garagem": garagem,
            "condominio": valor_cond,
            "iptu": valor_iptu,
            "anunciante": anunciante,
            "link": url,
            "indice": a,
            "pagina": i 
        }

        linha_add = [loc, andar, banheiros, quartos, suite, valor,metros, garagem, valor_cond, valor_iptu, anunciante, url]

        sheet.append_row(linha_add)

        # Salvar em um arquivo JSON a ultima pagina e o ultima link
        with open("dados.json", "w", encoding="utf-8") as f:
            json.dump(dados, f, indent=4, ensure_ascii=False)

        driver.close()
        driver.switch_to.window(abas[0])

        time.sleep(3)


