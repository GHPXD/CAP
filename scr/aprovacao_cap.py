import pandas as pd
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
import os


def get_excel_data(excel_path, sheet_name="Folha de Serviço"):
    # Carregar a planilha do Excel usando pandas
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    return df


def approve_task(driver, data_hoje, data_ontem, excel_data):
    # Verificar se as condições de aprovação estão atendidas
    for index, row in excel_data.iterrows():
        data_celula = row['Data']
        status = row['Status']
        numero = row['Número']
        fluxo_trabalho = row['Fluxo de Trabalho']
        
        # Validação conforme as condições no VBA
        if status == "Aprovado" and isinstance(numero, (int, float)) and isinstance(fluxo_trabalho, (int, float)):
            # Validação de data
            if data_celula == data_hoje or data_celula == data_ontem:
                print(f"Aprovando tarefa: {numero} - {fluxo_trabalho}")
                
                # Navegar até a tarefa na página web e realizar aprovação
                try:
                    # Aqui você vai navegar até a linha correspondente à tarefa no seu site
                    driver.find_element(By.XPATH, f"//tr[td[1][contains(text(), '{numero}')]]").click()
                    time.sleep(2)
                    
                    # Realizar a aprovação clicando no botão ou realizando a ação necessária
                    approve_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Aprovar')]")
                    approve_button.click()
                    time.sleep(2)
                    
                    print(f"Tarefa {numero} aprovada com sucesso!")
                except Exception as e:
                    print(f"Erro ao tentar aprovar a tarefa {numero}: {e}")

    print("Processo de aprovação concluído!")


def setup_driver():
    # Configurações do Selenium WebDriver
    options = Options()
    options.add_argument("--start-maximized")
    
    # Caminho do chromedriver (substitua com o caminho correto do seu chromedriver)
    service = Service("path/to/chromedriver")
    driver = webdriver.Chrome(service=service, options=options)
    
    # Aumenta o tempo de espera para garantir que a página seja carregada corretamente
    driver.implicitly_wait(10)
    return driver


def login(driver, username, password):
    # Acessa o portal de login e faz o login com as credenciais
    driver.get("https://venergia.capworkflow.com/")
    time.sleep(3)
    
    # Seleciona a opção Votorantim (Azure AD)
    driver.find_element(By.XPATH, "/html/body/form/div[3]/div/div/div[2]/div/div/div[2]/div/button").click()
    time.sleep(3)
    
    # Insere o e-mail
    driver.find_element(By.XPATH, "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[2]/div[2]/div/input[1]").send_keys(username)
    driver.find_element(By.XPATH, "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[4]/div/div/div/div/input").click()
    time.sleep(2)
    
    # Insere a senha
    driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div/main/div/div/div/form/div[2]/div[2]/input").send_keys(password)
    driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div/main/div/div/div/form/div[2]/div[4]/span").click()
    time.sleep(3)


def main(excel_path, username, password):
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
    # Defina o caminho da planilha, nome de usuário e senha
    excel_path = "path_to_your_excel_file.xlsx"
    username = "joao.souza.js1@votorantim.com"
    password = "your_password"
    
    main(excel_path, username, password)