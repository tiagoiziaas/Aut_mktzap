# config.py
import os
import tkinter as tk
from tkinter import filedialog

# Configurar caminhos


def selecionar_arquivo(tipo_arquivo):
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal do Tkinter
    caminho_arquivo = filedialog.askopenfilename(
        title=f"Selecione o arquivo {tipo_arquivo}",
        filetypes=[(f"Arquivos {tipo_arquivo}", "*.*")]
    )
    if not caminho_arquivo:
        print(f"Nenhum arquivo {tipo_arquivo} selecionado.")
    return caminho_arquivo

# Selecionar o arquivo de mensagem
mensagem_file_path = selecionar_arquivo("Mensagem")

# Selecionar o arquivo Excel
arquivo_excel_path = selecionar_arquivo("Excel")
# URL do site mktzap
mktzap_url = 'https://www.mktzap.com.br/'

# Credenciais de login
EMAIL = 'thaina.duarte@folhatech.com.br'
SENHA = 'Thaina@2024'
