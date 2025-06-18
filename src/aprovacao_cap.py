import pandas as pd
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
from tkinter import Tk, filedialog
import logging

# Configurar o logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def select_excel_file():
    """Abrir janela de diálogo para selecionar arquivo Excel."""
    root = Tk()
    root.withdraw()
    
    file_path = filedialog.askopenfilename(
        title='Selecione o arquivo Excel',
        filetypes=[('Excel Files', '*.xlsx'), ('All Files', '*.*')],
        initialdir=os.getcwd()
    )
    
    return file_path if file_path else None

def setup_driver(chromedriver_path):
    """Configurar o WebDriver do Selenium."""
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu") 
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    service = Service(chromedriver_path)
    driver = webdriver.Chrome(service=service, options=options) 

    logging.info("WebDriver configurado.")
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
        
        logging.info(f"Dados da planilha '{sheet_name}' carregados com sucesso.")
        return df
        
    except FileNotFoundError:
        logging.error(f"Arquivo não encontrado: {excel_path}")
        return None
    except Exception as e:
        logging.error(f"Erro ao ler o arquivo Excel: {str(e)}")
        return None

def login(driver, username, password):
    """Fazer login na plataforma."""
    driver.get("https://venergia.capworkflow.com/")
    
    wait = WebDriverWait(driver, 20)
    
    try:
        # Esperar pelo botão de login principal e clicar
        login_button_main = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Acessar')]")))
        login_button_main.click()
        logging.info("Clicou no botão 'Acessar'.")
        
        # Esperar pelo campo de email e preencher
        email_input = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@type='email' and @placeholder='Email, telefone ou Skype']")))
        email_input.send_keys(username)
        logging.info("Preencheu o email.")
        
        # Esperar pelo botão "Avançar" e clicar
        next_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='submit' and @value='Avançar']")))
        next_button.click()
        logging.info("Clicou em 'Avançar'.")
        
        # Esperar pelo campo de senha e preencher
        password_input = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@type='password' and @placeholder='Senha']")))
        password_input.send_keys(password)
        logging.info("Preencheu a senha.")
        
        # Esperar pelo botão "Entrar" e clicar
        signin_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(., 'Entrar')]")))
        signin_button.click()
        logging.info("Clicou em 'Entrar'.")

        # Esperar que o login seja bem-sucedido
        wait.until(EC.url_changes("https://venergia.capworkflow.com/Home/Index")) # Ou um elemento específico da homepage
        logging.info("Login realizado com sucesso.")

    except Exception as e:
        logging.error(f"Erro no login: {e}")
        driver.quit()
        raise 

def approve_task(driver, data_hoje, data_ontem, excel_data):
    """Aprovar as tarefas com base nos dados da planilha."""
    if excel_data is None or excel_data.empty:
        logging.info("Sem dados para processar.")
        return
    
    wait = WebDriverWait(driver, 15) 
        
    for index, row in excel_data.iterrows():
        try:
            solicitacao = row['Solicitação']
            folha_servico = row['FRS']
            status = row['Status Aprov.']
            data_aprovacao = row['Data Aprovação']
            
            if pd.isna(solicitacao) or pd.isna(status) or pd.isna(data_aprovacao):
                logging.warning(f"Linha {index + 2}: Dados incompletos ou inválidos, pulando. Solicitação: {solicitacao}, Status: {status}, Data: {data_aprovacao}")
                continue
            
            # Assegurar que 'Solicitação' é um inteiro para o XPATH
            solicitacao_int = int(solicitacao)
                
            if status == "Aprovado" and (data_aprovacao == data_hoje or data_aprovacao == data_ontem):
                logging.info(f"Tentando aprovar tarefa: Solicitação: {solicitacao_int} - FRS: {folha_servico}")
                
                try:
                    # Clicar na linha da tarefa
                    row_element = wait.until(EC.element_to_be_clickable((By.XPATH, f"//td[text()='{solicitacao_int}']/ancestor::tr")))
                    row_element.click()
                    logging.info(f"Clicou na linha da tarefa {solicitacao_int}.")
                    
                    # Esperar pelo botão "Aprovar" e clicar
                    approve_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Aprovar') or contains(text(), 'Approve')]")))
                    approve_button.click()
                    
                    # Confirmar a aprovação (se houver um popup ou similar)
                    try:
                        alert = wait.until(EC.alert_is_present())
                        alert.accept() # Aceita o alerta de confirmação, se houver
                        logging.info("Alerta de confirmação aceito.")
                    except:
                        pass # Não há alerta, continua
                    
                    logging.info(f"Tarefa {solicitacao_int} aprovada com sucesso!")

                except Exception as e:
                    logging.error(f"Erro ao tentar aprovar a tarefa {solicitacao_int}: {e}")
                    # Voltar à lista de tarefas se a aprovação falhar para tentar a próxima
                    driver.refresh()
                    wait.until(EC.url_contains("capworkflow.com/tasks")) # Espera a página de tarefas carregar novamente

            else:
                logging.info(f"Linha {index + 2}: Condições de aprovação não atendidas para Solicitação: {solicitacao_int}. Status: '{status}', Data: '{data_aprovacao}'.")
                
        except Exception as e:
            logging.error(f"Erro geral ao processar linha {index + 2}: {e}")
            continue
            
    logging.info("Processo de aprovação concluído!")

def main_aprovacao(username, password, chromedriver_path):
    """Função principal para automatizar o processo de aprovação."""
    logging.info("Iniciando processo de aprovação de CAP.")
    excel_path = select_excel_file()
    
    if not excel_path:
        logging.info("Nenhum arquivo Excel selecionado. Encerrando programa.")
        return
        
    excel_data = get_excel_data(excel_path)
    if excel_data is None:
        logging.error("Falha ao carregar dados do Excel. Encerrando.")
        return

    data_hoje = datetime.date.today()
    data_ontem = data_hoje - datetime.timedelta(days=1)
    
    driver = None
    try:
        driver = setup_driver(chromedriver_path)
        login(driver, username, password)
        # Navegar para a página de tarefas após o login bem-sucedido
        wait = WebDriverWait(driver, 10)
        tasks_tab = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, '/tasks') or contains(., 'TAREFAS')]")))
        tasks_tab.click()
        logging.info("Navegou para a aba 'TAREFAS'.")

        approve_task(driver, data_hoje, data_ontem, excel_data)
    except Exception as e:
        logging.critical(f"Erro crítico durante a execução do processo de aprovação: {e}")
    finally:
        if driver:
            driver.quit()
            logging.info("Navegador fechado.")

if __name__ == "__main__":
    username_placeholder = "seu_email@dominio.com"
    password_placeholder = "sua_senha_secreta"
    
    # É CRÍTICO que o ChromeDriver esteja no PATH ou que o caminho seja absoluto e correto.
    # Recomendado: coloque o chromedriver.exe na mesma pasta dos scripts ou na pasta "src".
    # ou, melhor ainda, configure uma variável de ambiente CHROMEDRIVER_PATH
    # Ex: CHROMEDRIVER_PATH = "C:/caminho/para/seu/chromedriver.exe"
    # ou use webdriver_manager para gerenciar o chromedriver automaticamente
    
    # Exemplo de uso de variável de ambiente:
    chromedriver_env_path = os.getenv('CHROMEDRIVER_PATH')
    if chromedriver_env_path:
        chromedriver_path = chromedriver_env_path
    else:
        possible_path_in_src = os.path.join(os.path.dirname(__file__), 'chromedriver.exe')
        chromedriver_path = possible_path_in_src
        
        if not os.path.exists(chromedriver_path):
            logging.warning(f"ChromeDriver não encontrado em: {chromedriver_path}. Por favor, baixe-o e defina a variável de ambiente 'CHROMEDRIVER_PATH' ou ajuste o caminho no código.")

    main_aprovacao(username_placeholder, password_placeholder, chromedriver_path)