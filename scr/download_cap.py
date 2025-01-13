from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
import os
import shutil
import datetime


def move_downloaded_files(download_start_time):
    # Caminho das pastas de origem e destino
    user_name = os.getlogin()
    source_folder = f"C:/Users/{user_name}/Downloads/"
    destination_folder = f"C:/Users/{user_name}/Votorantim/Financeiro - SUPRIMENTOS/02. SUP 02 GESTÃO DE LANÇAMENTO DE NOTAS FISCAIS/FOLHA DE SERVIÇO/Anexos SAP/TESTE/"
    
    # Criação da pasta de destino caso não exista
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)

    # Iteração sobre os arquivos na pasta de downloads
    for file_name in os.listdir(source_folder):
        file_path = os.path.join(source_folder, file_name)
        if os.path.isfile(file_path):
            # Verificar se o arquivo foi modificado após o início do download
            file_mod_time = datetime.datetime.fromtimestamp(os.path.getmtime(file_path))
            if file_mod_time >= download_start_time:
                # Mover o arquivo para a pasta de destino
                shutil.move(file_path, os.path.join(destination_folder, file_name))

    print("Arquivos movidos com sucesso!")


def download_cap():
    # Configurações do Selenium WebDriver
    options = Options()
    options.add_argument("--start-maximized")
    
    # Caminho do chromedriver (substitua com o caminho correto do seu chromedriver)
    service = Service("path/to/chromedriver")
    driver = webdriver.Chrome(service=service, options=options)
    
    # Aumenta o tempo de espera para garantir que a página seja carregada corretamente
    driver.implicitly_wait(10)
    
    # Link do ServiceNow
    driver.get("https://venergia.capworkflow.com/")
    
    # Inicia a contagem do tempo de download
    download_start_time = datetime.datetime.now()

    # Realiza o login na plataforma
    driver.find_element(By.XPATH, "/html/body/form/div[3]/div/div/div[2]/div/div/div[2]/div/button").click()
    time.sleep(3)

    # Insere o email de login
    driver.find_element(By.XPATH, "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[2]/div[2]/div/input[1]").send_keys("joao.souza.js1@votorantim.com")
    driver.find_element(By.XPATH, "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[4]/div/div/div/div/input").click()

    # Tratamento de erro de senha
    try:
        driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div/main/div/div/div/form/div[2]/div[2]/input").send_keys("123")
        driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div/main/div/div/div/form/div[2]/div[4]/span").click()
    except:
        pass

    print("Confirme após a autenticação")

    # Aguardar até que a página de tarefas seja carregada
    time.sleep(5)

    # Navega até a aba "TAREFAS"
    driver.find_element(By.XPATH, "/html/body/div/aside/div/nav/ul/li[2]/a").click()
    driver.execute_script("document.body.style.zoom='25%'")
    
    # Ajuste de colunas e seleção
    driver.find_element(By.XPATH, "/html/body/div/section/div/div[2]/div[1]/button").click()
    time.sleep(1)

    # Adiciona as colunas necessárias
    columns = [
        "Número", "Fluxo de Trabalho", "Etapa", "Solicitante", "Pendente desde", "Pendente Com",
        "Boletim de Medição", "Boletim de Medição (Pré-Fatura)", "Valor Bruto (R$)", "N° de Pedido de compra",
        "Empresa", "Fornecedor", "Comentários", "Comentários - Solicitante", "Data Pagamento", "Observação"
    ]

    input_box = driver.find_element(By.XPATH, "/html/body/div/section/div/div[2]/div[1]/div/div[1]/div/span/input[2]")
    
    for column in columns:
        input_box.send_keys(column)
        input_box.send_keys(Keys.TAB)
    time.sleep(2)

    driver.find_element(By.XPATH, "/html/body/div/section/div/div[2]/div[1]/div/div[1]/div/span/input[2]").send_keys("Observação")
    driver.find_element(By.XPATH, "/html/body/div/section/div/div[2]/div[1]/div/button").click()
    time.sleep(3)

    # Encontrar todos os elementos de linha <tr> e iterar sobre eles para buscar os links de download
    rows = driver.find_elements(By.XPATH, "//tbody/tr")
    row_count = len(rows)
    print(f"Número de linhas encontradas: {row_count}")

    # Itera sobre as linhas e busca os links de download
    for i in range(1, row_count + 1):
        try:
            # Tenta encontrar o link de download na 14ª coluna
            download_link = driver.find_element(By.XPATH, f"//tbody/tr[{i}]/td[14]/a")
        except:
            # Se não encontrar, tenta a 15ª coluna
            try:
                download_link = driver.find_element(By.XPATH, f"//tbody/tr[{i}]/td[15]/a")
            except:
                download_link = None

        if download_link:
            # Clica no link de download
            download_link.click()
            time.sleep(2)

    # Após o download, move os arquivos para a pasta de destino
    move_downloaded_files(download_start_time)

    # Fecha o navegador
    driver.quit()


if __name__ == "__main__":
    download_cap()