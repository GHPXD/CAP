import pandas as pd
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
import os

def get_excel_data(excel_path, sheet_name="Controle CAP"):
    """Carregar os dados da planilha Excel."""
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    return df

def approve_task(driver, data_hoje, data_ontem, excel_data):
    """Aprovar as tarefas com base nos dados da planilha."""
    for index, row in excel_data.iterrows():
        # Acessar as colunas conforme indicado
        solicitacao = row['Solicitação']
        folha_servico = row['FRS']
        status = row['Status Aprov.']
        data_aprovacao = row['Data Aprovação']
        
        # Validação conforme as condições
        if status == "Aprovado" and isinstance(solicitacao, (int, float)):
            # Validar se a data de aprovação é "hoje" ou "ontem"
            if data_aprovacao == data_hoje or data_aprovacao == data_ontem:
                print(f"Aprovando tarefa: {solicitacao} - {folha_servico}")
                
                try:
                    # Navegar até a tarefa na página web
                    driver.find_element(By.XPATH, f"//tr[td[1][contains(text(), '{solicitacao}')]]").click()
                    time.sleep(2)
                    
                    # Realizar a aprovação clicando no botão ou realizando a ação necessária
                    approve_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Aprovar')]")
                    approve_button.click()
                    time.sleep(2)
                    
                    print(f"Tarefa {solicitacao} aprovada com sucesso!")
                except Exception as e:
                    print(f"Erro ao tentar aprovar a tarefa {solicitacao}: {e}")
    print("Processo de aprovação concluído!")

def setup_driver():
    """Configurar o WebDriver do Selenium."""
    options = Options()
    options.add_argument("--start-maximized")
    
    chromedriver_path = os.getenv('CHROMEDRIVER_PATH', "path/to/chromedriver")  # Caminho do chromedriver
    service = Service(chromedriver_path)
    driver = webdriver.Chrome(service=service, options=options)
    
    driver.implicitly_wait(10)  # Aguardar o carregamento da página
    return driver

def login(driver, username, password):
    """Fazer login na plataforma."""
    driver.get("https://venergia.capworkflow.com/")
    time.sleep(3)
    
    try:
        driver.find_element(By.XPATH, "/html/body/form/div[3]/div/div/div[2]/div/div/div[2]/div/button").click()
        time.sleep(3)
        
        driver.find_element(By.XPATH, "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[2]/div[2]/div/input[1]").send_keys(username)
        driver.find_element(By.XPATH, "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[4]/div/div/div/div/input").click()
        time.sleep(2)
        
        driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div/main/div/div/div/form/div[2]/div[2]/input").send_keys(password)
        driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div/main/div/div/div/form/div[2]/div[4]/span").click()
        time.sleep(3)
    except Exception as e:
        print(f"Erro no login: {e}")
        driver.quit()

def main(excel_path, username, password):
    """Função principal para automatizar o processo de aprovação."""
    # Obter os dados da planilha
    excel_data = get_excel_data(excel_path)

    # Definir as datas de hoje e ontem
    data_hoje = datetime.date.today()
    data_ontem = data_hoje - datetime.timedelta(days=1)
    
    # Configurar o WebDriver
    driver = setup_driver()
    
    # Login na plataforma
    login(driver, username, password)

    # Aguardar a página carregar
    time.sleep(5)
    
    # Chamar a função de aprovação das tarefas
    approve_task(driver, data_hoje, data_ontem, excel_data)
    
    # Fechar o driver após a execução
    driver.quit()

if __name__ == "__main__":
    # Caminho para o arquivo Excel
    excel_path = os.path.join(os.getcwd(), "assets", "CAP.xlsx")  # Caminho para o arquivo na pasta assets
    
    # Definir o nome de usuário e senha (pode ser carregado de variáveis de ambiente ou outro meio)
    username = "email@email.com"
    password = "your_password"
    
    main(excel_path, username, password)