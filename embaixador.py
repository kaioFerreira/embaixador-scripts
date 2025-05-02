import time
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def login_and_navigate_to_grid(driver):
    """
    Faz o login no site e navega até a página onde a grid de dados está presente.
    Retorna o container da grid.
    """
    driver.get("https://backoffice.locavibe.com.br/login")
    driver.maximize_window()

    # Clica em "Entrar"
    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/div/div/div[2]/div[1]/button'))
    )
    element.click()

    # Aguarda a presença do campo de usuário e preenche os dados
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "username")))
    driver.find_element(By.ID, "username").send_keys("kaiofhs@gmail.com")
    driver.find_element(By.ID, "password").send_keys("Kaka199#")
    driver.find_element(By.ID, "password").send_keys(Keys.RETURN)

    # Navega pela interface após o login
    element1 = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id=":r5:"]'))
    )
    element1.click()

    element2 = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id=":r4:"]/ul/li[1]/button'))
    )
    element2.click()

    element3 = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="sideBarNames"]/li[2]/a'))
    )
    element3.click()

    element4 = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/div/div/div/section[2]/div/div[2]/table/tbody/tr/td[1]/a'))
    )
    element4.click()

    element5 = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id=":ru:-tab-1"]'))
    )
    element5.click()

    # Aguarda a grid estar presente
    grid_container = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id=":ru:-tabpanel-1"]/div/section/div'))
    )

    return grid_container

def extract_grid_data():
    """
    Efetua o login, navega até a grid, rola a página para carregar todos os blocos e extrai os dados.
    Cada bloco deve conter 8 <label> com o padrão:
        - labels[1]: Cliente
        - labels[5]: Veículo
        - labels[7]: Data
    Os dados são logados no console e salvos em um arquivo Excel cujo nome inclui a data atual.
    """
    service = Service("./chromedriver")
    driver = webdriver.Chrome(service=service)

    grid_container = login_and_navigate_to_grid(driver)

    # Rola a grid para carregar todos os blocos
    previous_count = 0
    while True:
        blocks = grid_container.find_elements(By.XPATH, "./div")
        current_count = len(blocks)
        if current_count > previous_count:
            previous_count = current_count
            # Rola até o último bloco para carregar os próximos
            driver.execute_script("arguments[0].scrollIntoView(true);", blocks[-1])
            time.sleep(1)
        else:
            break

    # Re-obtem os blocos após a rolagem completa
    blocks = grid_container.find_elements(By.XPATH, "./div")
    dados = []

    for index, block in enumerate(blocks, start=1):
        labels = block.find_elements(By.TAG_NAME, "label")
        if len(labels) >= 8:
            cliente = labels[1].text.strip()
            veiculo = labels[5].text.strip()
            data = labels[7].text.strip()
            dados.append({"Cliente": cliente, "Veículo": veiculo, "Data": data})
            print(f"[Bloco {index}] Cliente: {cliente} | Veículo: {veiculo} | Data: {data}")
        else:
            print(f"[Bloco {index}] Estrutura inesperada: {block.text}")

    # Gera o nome do arquivo com a data atual
    today_str = datetime.now().strftime("%d-%m-%Y")
    filename = f"dados_{today_str}.xlsx"
    df = pd.DataFrame(dados)
    df.to_excel(filename, index=False)
    print(f"Dados salvos em {filename}")
    driver.quit()

def verify_clients_in_site(excel_file=None):
    """
    Lê a planilha com os dados extraídos e, após efetuar o login e navegar até a grid,
    verifica se cada cliente presente na planilha está no site.
    Imprime no console se o cliente foi encontrado ou não, e exibe um resumo final.
    Se excel_file não for informado, utiliza o arquivo do dia atual.
    """
    if excel_file is None:
        today_str = datetime.now().strftime("%d-%m-%Y")
        excel_file = f"dados_{today_str}.xlsx"
    
    df = pd.read_excel(excel_file)
    expected_clients = df["Cliente"].tolist()

    service = Service("./chromedriver")
    driver = webdriver.Chrome(service=service)

    grid_container = login_and_navigate_to_grid(driver)

    # Rola a grid para carregar todos os blocos
    previous_count = 0
    while True:
        blocks = grid_container.find_elements(By.XPATH, "./div")
        current_count = len(blocks)
        if current_count > previous_count:
            previous_count = current_count
            driver.execute_script("arguments[0].scrollIntoView(true);", blocks[-1])
            time.sleep(1)
        else:
            break

    blocks = grid_container.find_elements(By.XPATH, "./div")
    site_clients = []
    for block in blocks:
        labels = block.find_elements(By.TAG_NAME, "label")
        if len(labels) >= 8:
            cliente = labels[1].text.strip()
            site_clients.append(cliente)

    found_count = 0
    not_found_count = 0
    for cliente in expected_clients:
        if cliente in site_clients:
            print(f"Cliente encontrado: {cliente}")
            found_count += 1
        else:
            print(f"Cliente NÃO encontrado: {cliente}")
            not_found_count += 1

    print("\n=== Resumo ===")
    print(f"Total de clientes esperados: {len(expected_clients)}")
    print(f"Clientes encontrados: {found_count}")
    print(f"Clientes não encontrados: {not_found_count}")

    driver.quit()

# Para extrair e salvar os dados com o nome contendo a data de hoje:
extract_grid_data()

# Para verificar se os clientes da planilha (do dia atual) estão no site:
# verify_clients_in_site("dados_23-04-2025.xlsx")