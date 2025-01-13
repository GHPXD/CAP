# src/download_cap.py
import time
import os
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime

def download_cap():
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.set_page_load_timeout(300)  # Timeout para o carregamento da página

    download_start_time = datetime.now()

    driver.get("https://venergia.capworkflow.com/")
    driver.maximize_window()

    driver.find_element(By.XPATH, "/html/body/form/div[3]/div/div/div[2]/div/div/div[2]/div/button").click()
    time.sleep(3)
    driver.find_element(By.XPATH, "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[2]/div[2]/div/input[1]").send_keys("joao.souza.js1@votorantim.com")
    driver.find_element(By.XPATH, "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[4]/div/div/div/div/input").click()
    time.sleep(3)

    # Lógica de navegação no site...

    driver.quit()

    move_downloaded_files(download_start_time)

def move_downloaded_files(download_start_time):
    # Caminhos das pastas de origem e destino
    user_name = os.getlogin()
    source_folder = f"C:/Users/{user_name}/Downloads/"
    destination_folder = f"C:/Users/{user_name}/Votorantim/Financeiro - SUPRIMENTOS/02. SUP 02 GESTÃO DE LANÇAMENTO DE NOTAS FISCAIS/FOLHA DE SERVIÇO/Anexos SAP/TESTE/"

    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)

    for file_name in os.listdir(source_folder):
        file_path = os.path.join(source_folder, file_name)
        if os.path.getmtime(file_path) >= download_start_time.timestamp():
            shutil.move(file_path, os.path.join(destination_folder, file_name))

if __name__ == "__main__":
    download_cap()