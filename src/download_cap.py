from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import shutil
import datetime
import logging

# Configurar o logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def move_downloaded_files(download_start_time, source_folder, destination_folder):
    """Move arquivos baixados para a pasta de destino."""
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
        logging.info(f"Pasta de destino criada: {destination_folder}")

    moved_count = 0
    for file_name in os.listdir(source_folder):
        file_path = os.path.join(source_folder, file_name)
        if os.path.isfile(file_path):
            file_mod_time = datetime.datetime.fromtimestamp(os.path.getmtime(file_path))
            if file_mod_time >= download_start_time:
                try:
                    shutil.move(file_path, os.path.join(destination_folder, file_name))
                    moved_count += 1
                    logging.info(f"Arquivo '{file_name}' movido para '{destination_folder}'.")
                except Exception as e:
                    logging.error(f"Erro ao mover arquivo '{file_name}': {e}")

    logging.info(f"Total de {moved_count} arquivos movidos com sucesso.")

def download_cap_process(username, password, chromedriver_path, download_destination_folder=None):
    """Função principal para automatizar o processo de download de CAP."""
    logging.info("Iniciando processo de download de CAP.")

    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    
    # Definir a pasta de download para o Chrome, se especificado
    if download_destination_folder:
        prefs = {"download.default_directory" : download_destination_folder}
        options.add_experimental_option("prefs", prefs)
        logging.info(f"Download direto do Chrome configurado para: {download_destination_folder}")

    service = Service(chromedriver_path)
    driver = webdriver.Chrome(service=service, options=options)
    
    wait = WebDriverWait(driver, 20)
    
    driver.get("https://venergia.capworkflow.com/")
    
    # Inicia a contagem do tempo de download
    download_start_time = datetime.datetime.now()

    try:
        # Login na plataforma (reuso da lógica de aprovacao_cap.py para consistência)
        login_button_main = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Acessar')]")))
        login_button_main.click()
        logging.info("Clicou no botão 'Acessar'.")
        
        email_input = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@type='email' and @placeholder='Email, telefone ou Skype']")))
        email_input.send_keys(username)
        logging.info("Preencheu o email.")
        
        next_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='submit' and @value='Avançar']")))
        next_button.click()
        logging.info("Clicou em 'Avançar'.")
        
        password_input = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@type='password' and @placeholder='Senha']")))
        password_input.send_keys(password)
        logging.info("Preencheu a senha.")
        
        signin_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(., 'Entrar')]")))
        signin_button.click()
        logging.info("Clicou em 'Entrar'.")

        wait.until(EC.url_changes("https://venergia.capworkflow.com/Home/Index"))
        logging.info("Login realizado com sucesso.")

        # Navega até a aba "TAREFAS"
        tasks_tab = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, '/tasks') or contains(., 'TAREFAS')]")))
        tasks_tab.click()
        logging.info("Navegou para a aba 'TAREFAS'.")
        
        # Clicar no botão para ajustar colunas
        adjust_columns_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'btn-primary') and contains(., 'Ajustar Colunas')]")))
        adjust_columns_button.click()
        logging.info("Clicou em 'Ajustar Colunas'.")

        # Adiciona as colunas necessárias
        columns = [
            "Número", "Fluxo de Trabalho", "Etapa", "Solicitante", "Pendente desde", "Pendente Com",
            "Boletim de Medição", "Boletim de Medição (Pré-Fatura)", "Valor Bruto (R$)", "N° de Pedido de compra",
            "Empresa", "Fornecedor", "Comentários", "Comentários - Solicitante", "Data Pagamento", "Observação"
        ]

        input_box = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[contains(@class, 'form-control') and @placeholder='Buscar']")))
        
        for column in columns:
            input_box.send_keys(column)
            input_box.send_keys(Keys.TAB)
            logging.info(f"Adicionou coluna: {column}")
        
        # Clicar no botão "Aplicar" ou similar após adicionar colunas
        apply_columns_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Aplicar')]")))
        apply_columns_button.click()
        logging.info("Colunas ajustadas e aplicadas.")
        time.sleep(3)
        # Encontrar todos os elementos de linha <tr> e iterar sobre eles para buscar os links de download
        rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//tbody/tr")))
        row_count = len(rows)
        logging.info(f"Número de linhas encontradas na tabela: {row_count}")

        # Itera sobre as linhas e busca os links de download
        for i in range(1, row_count + 1):
            try:
                # Tenta encontrar o link de download na 14ª ou 15ª coluna (ajustar XPATHs se instável)
                download_link = None
                try:
                    download_link = driver.find_element(By.XPATH, f"//tbody/tr[{i}]/td[14]/a[contains(@href, 'download')]")
                except:
                    try:
                        download_link = driver.find_element(By.XPATH, f"//tbody/tr[{i}]/td[15]/a[contains(@href, 'download')]")
                    except:
                        pass
                
                if download_link:
                    download_link.click()
                    logging.info(f"Clicou no link de download na linha {i}.")
                    time.sleep(2)

            except Exception as e:
                logging.warning(f"Erro ao tentar baixar na linha {i}: {e}")
                continue

        # Após o download, move os arquivos para a pasta de destino padrão do usuário e depois para a pasta "teste"
        if not download_destination_folder:
            user_default_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
            # Definir a pasta de destino final (Documentos/teste)
            final_destination_folder = os.path.join(os.path.expanduser("~"), "Documents", "teste")
            move_downloaded_files(download_start_time, user_default_downloads, final_destination_folder)

    except Exception as e:
        logging.critical(f"Erro crítico durante a execução do processo de download: {e}")
    finally:
        driver.quit()
        logging.info("Navegador fechado.")

if __name__ == "__main__":
    username_placeholder = "seu_email@dominio.com"
    password_placeholder = "sua_senha_secreta"

    # Caminho do ChromeDriver
    chromedriver_env_path = os.getenv('CHROMEDRIVER_PATH')
    if chromedriver_env_path:
        chromedriver_path = chromedriver_env_path
    else:
        # Tenta um caminho relativo dentro do src/ ou um caminho específico
        possible_path_in_src = os.path.join(os.path.dirname(__file__), 'chromedriver.exe')
        # Ajuste para uma forma mais dinâmica ou peça ao usuário.
        chromedriver_path = "C:/Users/seuuser/Área de Trabalho/CAP-main/chromedriver.exe"

    if not os.path.exists(chromedriver_path):
        logging.warning(f"ChromeDriver não encontrado em: {chromedriver_path}. Por favor, baixe-o e defina a variável de ambiente 'CHROMEDRIVER_PATH' ou ajuste o caminho no código.")
        from tkinter import messagebox
        messagebox.showerror("Erro de Configuração", "ChromeDriver não encontrado. Por favor, verifique o arquivo chromedriver.exe e o caminho no código ou defina a variável de ambiente 'CHROMEDRIVER_PATH'.")
        exit() # Encerra o script se o driver não for encontrado

    download_cap_process(username_placeholder, password_placeholder, chromedriver_path)