import tkinter as tk
from tkinter import messagebox
import os
import sys
from src.aprovacao_cap import main_aprovacao
from src.download_cap import download_cap_process

USERNAME = "seu_email@dominio.com"
PASSWORD = "sua_senha_secreta"

# Determinar o caminho do chromedriver dinamicamente
def get_chromedriver_path():
    # Tenta obter do PATH do sistema
    try:
        from shutil import which
        chromedriver_path = which("chromedriver")
        if chromedriver_path:
            return chromedriver_path
    except ImportError:
        pass

    # Tenta a partir de uma variável de ambiente
    env_path = os.getenv('CHROMEDRIVER_PATH')
    if env_path and os.path.exists(env_path):
        return env_path
    
    script_dir = os.path.dirname(__file__) # src/
    project_root = os.path.abspath(os.path.join(script_dir, os.pardir))

    possible_paths = [
        os.path.join(project_root, 'chromedriver.exe'),
        os.path.join(script_dir, 'chromedriver.exe'),
        os.path.join(os.path.expanduser("~"), "Downloads", "chromedriver.exe")
    ]

    for path in possible_paths:
        if os.path.exists(path):
            return path
            
    try:
        from webdriver_manager.chrome import ChromeDriverManager
        driver_path = ChromeDriverManager().install()
        messagebox.showinfo("ChromeDriver", f"ChromeDriver baixado e configurado automaticamente em: {driver_path}")
        return driver_path
    except ImportError:
        messagebox.showwarning("Aviso", "A biblioteca 'webdriver_manager' não está instalada. Baixe 'chromedriver.exe' manualmente e coloque-o na raiz do projeto ou na pasta 'src/'.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao baixar ChromeDriver automaticamente: {e}. Por favor, baixe-o manualmente.")

    return None


def execute_download_cap():
    chromedriver_path = get_chromedriver_path()
    if not chromedriver_path:
        messagebox.showerror("Erro", "Não foi possível encontrar/configurar o ChromeDriver. O processo não pode continuar.")
        return
    try:
        # Passa as credenciais e o caminho do chromedriver diretamente
        download_cap_process(USERNAME, PASSWORD, chromedriver_path)
        messagebox.showinfo("Sucesso", "Processo de DownloadCAP concluído!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao rodar DownloadCAP: {e}")

def execute_aprovacao_cap():
    chromedriver_path = get_chromedriver_path()
    if not chromedriver_path:
        messagebox.showerror("Erro", "Não foi possível encontrar/configurar o ChromeDriver. O processo não pode continuar.")
        return
    try:
        # Passa as credenciais e o caminho do chromedriver diretamente
        main_aprovacao(USERNAME, PASSWORD, chromedriver_path)
        messagebox.showinfo("Sucesso", "Processo de AprovaçãoCAP concluído!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao rodar AprovaçãoCAP: {e}")

def create_interface():
    window = tk.Tk()
    window.title("Automação CAP - Selecione a Função")
    window.geometry("450x250") # Ajusta o tamanho da janela

    # Adicionar título
    title_label = tk.Label(window, text="Escolha uma função para rodar", font=("Arial", 16, "bold"))
    title_label.pack(pady=25)

    # Botão para rodar DownloadCAP
    btn_download = tk.Button(window, text="Rodar Download de CAP", width=35, height=2, command=execute_download_cap, font=("Arial", 12))
    btn_download.pack(pady=5)

    # Botão para rodar AprovaçãoCAP
    btn_aprovacao = tk.Button(window, text="Rodar Aprovação de CAP", width=35, height=2, command=execute_aprovacao_cap, font=("Arial", 12))
    btn_aprovacao.pack(pady=5)

    window.mainloop()

if __name__ == "__main__":
    create_interface()