import tkinter as tk
from tkinter import messagebox
import subprocess

def run_download_cap():
    try:
        subprocess.run(["python", "src/download_cap.py"])
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao rodar DownloadCAP: {e}")

def run_aprovacao_cap():
    try:
        subprocess.run(["python", "src/aprovacao_cap.py"])
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao rodar AprovaçãoCAP: {e}")

def create_interface():
    # Configuração da janela principal
    window = tk.Tk()
    window.title("Selecione a Função")
    window.geometry("400x200")

    # Adicionar título
    title_label = tk.Label(window, text="Escolha uma função para rodar", font=("Arial", 14))
    title_label.pack(pady=20)

    # Botão para rodar DownloadCAP
    btn_download = tk.Button(window, text="Rodar DownloadCAP", width=30, command=run_download_cap)
    btn_download.pack(pady=10)

    # Botão para rodar AprovaçãoCAP
    btn_aprovacao = tk.Button(window, text="Rodar AprovaçãoCAP", width=30, command=run_aprovacao_cap)
    btn_aprovacao.pack(pady=10)

    # Inicia a interface
    window.mainloop()

if __name__ == "__main__":
    create_interface()