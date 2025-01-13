# src/aprovacao_cap.py
import time
import os
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime

def aprovacao_cap():
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.set_page_load_timeout(300)

    driver.get("https://venergia.capworkflow.com/")
    driver.maximize_window()

    driver.find_element(By.XPATH, "/html/body/form/div[3]/div/div/div[2]/div/div/div[2]/div/button").click()
    time.sleep(3)
    driver.find_element(By.XPATH, "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[2]/div[2]/div/input[1]").send_keys("joao.souza.js1@votorantim.com")
    driver.find_element(By.XPATH, "/html/body/div/form[1]/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[3]/div/div/div/div[4]/div/div/div/div/input").click()
    time.sleep(3)

    # Lógica de navegação no site...

    driver.quit()

if __name__ == "__main__":
    aprovacao_cap()