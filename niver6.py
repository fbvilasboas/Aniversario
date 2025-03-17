import sys
import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk
import win32com.client as win32

from datetime import datetime

# Caminho fixo do arquivo Excel a ser lido
#caminho_excel = r"C:\Users\fernando.vilasboas\Desktop\NIVER\resultado.xlsx"

caminho_excel = r"resultado.xlsx"

def gerar_excel_de_iqy(caminho_arquivo_iqy, caminho_arquivo_excel):
    try:
        # Remove o arquivo Excel existente, se existir
        if os.path.exists(caminho_arquivo_excel):
            print(f"Removendo o arquivo existente: {caminho_arquivo_excel}")
            os.remove(caminho_arquivo_excel)
        
        # Verifica se o arquivo IQY existe
        if not os.path.exists(caminho_arquivo_iqy):
            print(f"Erro: O arquivo {caminho_arquivo_iqy} não foi encontrado.")
            return
        
        # Inicializa o Excel
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False  # Define como False para rodar em segundo plano
        
        # Normaliza o caminho para evitar erros
        caminho_arquivo_iqy = os.path.abspath(caminho_arquivo_iqy)
        caminho_arquivo_excel = os.path.abspath(caminho_arquivo_excel)

        # Abre o arquivo IQY no Excel
        print(f"Abrindo arquivo IQY: {caminho_arquivo_iqy}")
        workbook = excel.Workbooks.Open(caminho_arquivo_iqy)

        # Atualiza os dados da consulta
        print("Atualizando dados...")
        workbook.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()  # Aguarda consultas assíncronas serem finalizadas
        
        # Salva como um arquivo Excel
        print(f"Salvando o arquivo Excel em: {caminho_arquivo_excel}")
        workbook.SaveAs(caminho_arquivo_excel, FileFormat=51)  # FileFormat=51 é para arquivos .xlsx

        # Fecha o Excel
        workbook.Close(SaveChanges=False)
        excel.Quit()
        
        print(f"Arquivo Excel gerado com sucesso: {caminho_arquivo_excel}")
        ##
        # 
        lixo = input("O arquivo resultado.xlsx foi criado. Quando testar remover os e-mails ")
        lixo = input ("Enter para prosseguir")


    except Exception as e:
        print(f"Erro ao gerar o arquivo Excel: {e}")



# Função para carregar o arquivo Excel fixo
def carregar_arquivo():
    global dados, lista_emails

    caminho_iqy = r"consulta.iqy"
    caminho_excel = r"resultado.xlsx"


    gerar_excel_de_iqy(caminho_iqy, caminho_excel)
    try:
        # Verifica se o arquivo existe
        if not os.path.exists(caminho_excel):
            raise FileNotFoundError(f"O arquivo {caminho_excel} não foi encontrado.")

        # Lê o arquivo Excel
        dados = pd.read_excel(caminho_excel)

        # Verifica se as colunas necessárias estão presentes
        colunas_necessarias = {'Nome', 'Data Nascimento', 'E-mail'}
        colunas_arquivo = set(dados.columns)

        if not colunas_necessarias.issubset(colunas_arquivo):
            raise ValueError(
                f"O arquivo deve conter as colunas: {', '.join(colunas_necessarias)}.\n"
                f"Colunas encontradas: {', '.join(colunas_arquivo)}"
            )

        # Converte a coluna de data de nascimento para datetime
        dados['Data Nascimento'] = pd.to_datetime(dados['Data Nascimento'], errors='coerce')

        # Extrai os e-mails válidos
        lista_emails = dados['E-mail'].dropna().tolist()

    except Exception as e:
        # Mostra a mensagem de erro detalhada na interface gráfica
        messagebox.showerror("Erro", f"Erro ao processar o arquivo: {str(e)}")
        dados = None
        lista_emails = []

# Função para enviar e-mail
def enviar_email(lista_aniversariantes, mes):
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mes = mes.upper()

        # Obtém o nome do remetente
        nome_remetente = outlook.Session.CurrentUser.Name

        # Configurando e-mail
        assunto = f"Aniversariantes do mês {mes}"
        corpo = f"Segue a lista de aniversariantes do mês {mes}:\n\nANIVERSÁRIO   ANIVERSARIANTE\n\n" + "\n".join(lista_aniversariantes) + f"\n\nAtenciosamente,\n\n{nome_remetente}"

        mail.Subject = assunto
        mail.Body = corpo

        # Adicionar todos os e-mails na lista de destinatários
        #destinatarios = ";".join(lista_emails)
        destinatarios = "fernando.vilasboas@agu.gov.br"
        
        #mail.To = destinatarios
        #usar linha abaixo no lugar do mail.To , quando quiser que NÃO APAREÇAM, no cabeçalho do e-mail, TODOS OS DESTINATÁRIOS
        mail.BCC = destinatarios 

        #o comando mail.Send() envia os e-mails . Vou só imprimir a lista de emails
        mail.Send()
        #
        print ("lista de destinatarios\n")
        print(destinatarios)
        #
        #
        #
        

        messagebox.showinfo("Sucesso", "E-mail enviado com sucesso!")

        if os.path.exists(caminho_excel):
            print(f"Removendo o arquivo existente: {caminho_excel}")
            os.remove(caminho_excel)

        print("O programa vai encerrar agora.")
        sys.exit()  # Encerra o programa
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao enviar e-mail: {str(e)}")

# Função para buscar aniversariantes do mês selecionado
def buscar_aniversariantes():
    if dados is None:
        messagebox.showerror("Erro", "Os dados do arquivo ainda não foram carregados.")
        return

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
    
    
    #imprimindo lista de aniversariantes
   
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

# Carrega automaticamente o arquivo no início do programa
carregar_arquivo()

# Combobox e Label para seleção do mês
label = tk.Label(root, text="Selecione o mês:")
label.pack(pady=10)

combobox_mes = ttk.Combobox(root, values=meses, state="readonly")
combobox_mes.pack(pady=10)

# Botão para buscar aniversariantes
btn_buscar = tk.Button(root, text="Enviar Lista Aniversariantes do mês", command=buscar_aniversariantes)
btn_buscar.pack(pady=20)

root.mainloop()
