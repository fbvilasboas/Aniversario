import sys
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk
import requests
from io import BytesIO
import win32com.client as win32
from datetime import datetime

# Função para baixar o arquivo Excel de um link SharePoint
def baixar_arquivo(link):
    try:
        # Faz o download do arquivo
        response = requests.get(link)
        response.raise_for_status()
        
        # Lê o conteúdo do arquivo em um DataFrame
        dados = pd.read_excel(BytesIO(response.content))
        
        # Verifica as colunas necessárias
        if 'Nome' not in dados.columns or 'Data Nascimento' not in dados.columns or 'E-mail' not in dados.columns:
            messagebox.showerror("Erro", "O arquivo deve conter as colunas 'Nome', 'Data Nascimento' e 'E-mail'.")
            return None
        
        # Converte a coluna de data de nascimento para datetime
        dados['Data Nascimento'] = pd.to_datetime(dados['Data Nascimento'], errors='coerce')
        return dados
    except requests.exceptions.RequestException as e:
        messagebox.showerror("Erro", f"Erro ao baixar o arquivo: {str(e)}")
        return None
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar o arquivo: {str(e)}")
        return None

# Função para enviar e-mail
def enviar_email(lista_aniversariantes, mes):
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mes = mes.upper()
        
        # Configurando e-mail
        assunto = f"Aniversariantes do mês {mes}"
        corpo = f"Segue a lista de aniversariantes do mês {mes}:\n\nANIVERSÁRIO   ANIVERSARIANTE\n\n" + "\n".join(lista_aniversariantes)

        mail.Subject = assunto
        mail.Body = corpo

        # Adicionar todos os e-mails na lista de destinatários
        destinatarios = ";".join(lista_emails)
        mail.To = destinatarios

        mail.Send()
        messagebox.showinfo("Sucesso", "E-mail enviado com sucesso!")
        print("O programa vai encerrar agora.")
        sys.exit()  # Encerra o programa
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao enviar e-mail: {str(e)}")

# Função para buscar aniversariantes do mês selecionado
def buscar_aniversariantes():
    # Obtém o mês selecionado no combobox
    mes = combobox_mes.get()
    if not mes:
        messagebox.showerror("Erro", "Por favor, selecione um mês.")
        return
    
    # Converte o nome do mês para o número correspondente
    mes_numero = meses.index(mes) + 1

    # Filtra aniversariantes do mês
    aniversariantes_mes = dados[
        dados['Data Nascimento'].dt.month == mes_numero
    ]
    
    if aniversariantes_mes.empty:
        messagebox.showinfo("Sem aniversariantes", f"Não há aniversariantes para o mês de {mes}.")
        return

    # Ordena os aniversariantes pelo dia do mês
    aniversariantes_mes = aniversariantes_mes.sort_values(by='Data Nascimento', key=lambda x: x.dt.day)

    # Formata a lista de aniversariantes
    lista_aniversariantes = [
        f"     {row['Data Nascimento'].strftime('%d/%m')}  ........ {row['Nome']}"
        for _, row in aniversariantes_mes.iterrows()
    ]

    # Exibe a lista e envia o e-mail
    enviar_email(lista_aniversariantes, mes)

# Interface gráfica com Tkinter
root = tk.Tk()
root.title("Aniversariantes do Mês")

# Variável global para armazenar os dados lidos
dados = None
lista_emails = []

# Lista de meses
meses = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]

# Campo para inserir o link do SharePoint
label_link = tk.Label(root, text="Insira o link do arquivo no SharePoint:")
label_link.pack(pady=10)

entry_link = tk.Entry(root, width=50)
entry_link.pack(pady=10)

# Botão para carregar o arquivo do link
btn_carregar = tk.Button(root, text="Carregar Arquivo", command=lambda: carregar_arquivo(entry_link.get()))
btn_carregar.pack(pady=10)

# Combobox e Label para seleção do mês
label = tk.Label(root, text="Selecione o mês:")
label.pack(pady=10)

combobox_mes = ttk.Combobox(root, values=meses, state="readonly")
combobox_mes.pack(pady=10)

# Botão para buscar aniversariantes
btn_buscar = tk.Button(root, text="Enviar Lista Aniversariantes do mês", command=lambda: buscar_aniversariantes())
btn_buscar.pack(pady=20)

# Função para carregar o arquivo e atualizar os dados globais
def carregar_arquivo(link):
    global dados, lista_emails
    dados = baixar_arquivo(link)
    if dados is not None:
        lista_emails = dados['E-mail'].dropna().tolist()

root.mainloop()
