import pandas as pd
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
import os
from tkinter import Tk, filedialog

def select_excel_file():
    """Abrir janela de diálogo para selecionar arquivo Excel."""
    root = Tk()
    root.withdraw()  # Esconde a janela principal do Tkinter
    
    file_path = filedialog.askopenfilename(
        title='Selecione o arquivo Excel',
        filetypes=[('Excel Files', '*.xlsx')],
        initialdir=os.getcwd()
    )
    
    return file_path if file_path else None

def setup_driver():
    """Configurar o WebDriver do Selenium."""
    options = Options()
    options.add_argument("--start-maximized")
    
    chromedriver_path = os.getenv('CHROMEDRIVER_PATH', "path/to/chromedriver")
    service = Service(chromedriver_path)
    driver = webdriver.Chrome(options=options)
    
    driver.implicitly_wait(10)
    return driver

def get_excel_data(excel_path, sheet_name="Controle CAP"):
    """Carregar os dados da planilha Excel."""
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
        
        # Validar colunas necessárias
        required_columns = ['Solicitação', 'FRS', 'Status Aprov.', 'Data Aprovação']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            raise ValueError(f"Colunas ausentes na planilha: {', '.join(missing_columns)}")
            
        # Converter coluna de data para datetime
        df['Data Aprovação'] = pd.to_datetime(df['Data Aprovação']).dt.date
        
        # Converter coluna de solicitação para numérico
        df['Solicitação'] = pd.to_numeric(df['Solicitação'], errors='coerce')
        
        return df
        
    except FileNotFoundError:
        print(f"Arquivo não encontrado: {excel_path}")
        return None
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {str(e)}")
        return None

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

def approve_task(driver, data_hoje, data_ontem, excel_data):
    """Aprovar as tarefas com base nos dados da planilha."""
    if excel_data is None or excel_data.empty:
        print("Sem dados para processar")
        return
        
    for index, row in excel_data.iterrows():
        try:
            solicitacao = row['Solicitação']
            folha_servico = row['FRS']
            status = row['Status Aprov.']
            data_aprovacao = row['Data Aprovação']
            
            if pd.isna(solicitacao) or pd.isna(status) or pd.isna(data_aprovacao):
                print(f"Linha {index + 2}: Dados incompletos, pulando...")
                continue
                
            if status == "Aprovado" and isinstance(solicitacao, (int, float)):
                if data_aprovacao == data_hoje or data_aprovacao == data_ontem:
                    print(f"Aprovando tarefa: {int(solicitacao)} - {folha_servico}")
                    
                    try:
                        driver.find_element(By.XPATH, f"//tr[td[1][contains(text(), '{int(solicitacao)}')]]").click()
                        time.sleep(2)
                        
                        approve_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Aprovar')]")
                        approve_button.click()
                        time.sleep(2)
                        
                        print(f"Tarefa {int(solicitacao)} aprovada com sucesso!")
                    except Exception as e:
                        print(f"Erro ao tentar aprovar a tarefa {int(solicitacao)}: {e}")
                        
        except Exception as e:
            print(f"Erro ao processar linha {index + 2}: {e}")
            continue
            
    print("Processo de aprovação concluído!")

def main(username, password):
    """Função principal para automatizar o processo de aprovação."""
    excel_path = select_excel_file()
    
    if not excel_path:
        print("Nenhum arquivo selecionado. Encerrando programa.")
        return
        
    excel_data = get_excel_data(excel_path)
    if excel_data is None:
        return

    data_hoje = datetime.date.today()
    data_ontem = data_hoje - datetime.timedelta(days=1)
    
    try:
        driver = setup_driver()
        login(driver, username, password)
        time.sleep(5)
        approve_task(driver, data_hoje, data_ontem, excel_data)
    except Exception as e:
        print(f"Erro durante a execução: {e}")
    finally:
        if 'driver' in locals():
            driver.quit()

if __name__ == "__main__":
    username = "email@email.com"
    password = "your_password"
    
    main(username, password)