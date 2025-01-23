import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from datetime import datetime

# Configurar variáveis de ambiente para credenciais do SharePoint
os.environ["SP_USER"] = "fernando.vilasboas@agu.gov.br"
os.environ["SP_PASSWORD"] = "AGU@#1234"

# Configuração do SharePoint
SHAREPOINT_URL = "https://agudf.sharepoint.com/sites/PRU6"
LISTA_NOME = "Nhttps://agudf.sharepoint.com/sites/PRU6/Lists/Quadro%20de%20pessoal%20PRU6%20%20apenas%20PSUs/AllItems.aspx?env=WebViewList&OR=Teams-HL&CT=1736945623048&clickparams=eyJBcHBOYW1lIjoiVGVhbXMtRGVza3RvcCIsIkFwcFZlcnNpb24iOiI0OS8yNDEyMDEwMDIxMSIsIkhhc0ZlZGVyYXRlZFVzZXIiOmZhbHNlfQ%3D%3D"  # Substitua pelo nome da sua lista no SharePoint

# Função para acessar lista do SharePoint
def acessar_lista_sharepoint():
    try:
        # Obter credenciais do ambiente
        usuario = os.getenv("SP_USER")
        senha = os.getenv("SP_PASSWORD")

        if not usuario or not senha:
            raise ValueError("Credenciais do SharePoint não configuradas.")

        # Conexão com o SharePoint
        ctx = ClientContext(SHAREPOINT_URL).with_credentials(UserCredential(usuario, senha))

        # Acessar os itens da lista
        lista = ctx.web.lists.get_by_title(LISTA_NOME)
        items = lista.items.get().execute_query()

        # Processar os itens da lista em um DataFrame
        dados_lista = [item.properties for item in items]
        df = pd.DataFrame(dados_lista)
        return df

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao acessar a lista do SharePoint: {str(e)}")
        return None

# Função para enviar e-mail
def enviar_email(lista_aniversariantes, mes):
    try:
        import win32com.client as win32
        
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mes = mes.upper()

        # Configurando e-mail
        assunto = f"Aniversariantes do mês {mes}"
        corpo = f"Segue a lista de aniversariantes do mês {mes}:\nANIVERSÁRIO   ANIVERSARIANTE\n\n" + "\n".join(lista_aniversariantes)

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

# Botão para carregar dados do SharePoint
btn_carregar = tk.Button(root, text="Carregar Dados do SharePoint", command=lambda: carregar_dados())
btn_carregar.pack(pady=10)

# Combobox e Label para seleção do mês
label = tk.Label(root, text="Selecione o mês:")
label.pack(pady=10)

combobox_mes = ttk.Combobox(root, values=meses, state="readonly")
combobox_mes.pack(pady=10)

# Botão para buscar aniversariantes
btn_buscar = tk.Button(root, text="Enviar Lista Aniversariantes do mês", command=lambda: buscar_aniversariantes())
btn_buscar.pack(pady=20)

# Função para carregar os dados do SharePoint
def carregar_dados():
    global dados, lista_emails
    dados = acessar_lista_sharepoint()
    if dados is not None:
        lista_emails = dados['E-mail'].dropna().tolist()

root.mainloop()
