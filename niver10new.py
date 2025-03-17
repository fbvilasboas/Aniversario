import sys
import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk
import win32com.client as win32
from datetime import datetime


# Função para obter o diretório do executável ou do script
def get_executable_dir():
    if getattr(sys, 'frozen', False):
        # Se o programa está sendo executado como um executável
        return sys._MEIPASS  # Caminho temporário onde os arquivos embutidos estão
    else:
        # Se o programa está sendo executado como um script
        return os.path.dirname(os.path.abspath(__file__))


# Função para gerar o arquivo Excel a partir do arquivo IQY
def gerar_excel_de_iqy(caminho_arquivo_iqy, caminho_arquivo_excel):
    try:
        # Remove o arquivo Excel existente, se existir
        if os.path.exists(caminho_arquivo_excel):
            os.remove(caminho_arquivo_excel)

        # Verifica se o arquivo IQY existe
        if not os.path.exists(caminho_arquivo_iqy):
            raise FileNotFoundError(f"O arquivo {caminho_arquivo_iqy} não foi encontrado.")

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

    except Exception as e:
        print(f"Erro ao gerar o arquivo Excel: {e}")
        raise


# Função para carregar o arquivo Excel fixo
def carregar_arquivo():
    global dados, lista_emails

    # Caminhos dos arquivos
    caminho_iqy = os.path.join(get_executable_dir(), "consulta.iqy")
    caminho_excel = os.path.join(os.path.dirname(sys.executable), "resultado.xlsx")

    # Gera o arquivo Excel a partir do IQY
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
def enviar_email(lista_aniversariantes, mes, caminho_excel):
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mes = mes.upper()

        # Obtém o nome do remetente
        nome_remetente = outlook.Session.CurrentUser.Name

        # Caminho da imagem
        caminho_imagem = os.path.join(get_executable_dir(), "FELIZ.PNG")

        # Configurando o conteúdo do e-mail em HTML
        corpo_html = f"""
        <html>
            <head>
                <style>
                    body {{
                        font-family: Times New Roman, sans-serif;
                        color: #333;
                        background-color: #f0f0f0;
                        padding: 20px;
                    }}
                    .header {{
                        font-size: 24px;
                        font-weight: bold;
                        color: #13348d;
                        margin-bottom: 20px;
                    }}
                    .lista {{
                        font-size: 16px;
                        margin-top: 20px;
                    }}
                    .footer {{
                        margin-top: 40px;
                        font-size: 24px;
                        color: #555;
                        margin-bottom: 20px;
                    }}
                </style>
            </head>
            <body>
                <div class="header">
                    Aniversariantes do mês de {mes}
                </div>
                <img src="cid:FELIZ_PNG" alt="Imagem de felicitações" width="672" height="380">
                <div class="lista">
                    <ul>
                        {''.join(
            f"<li><b>{aniversariante.split('........')[0]}</b>: {aniversariante.split('........')[1]}</li>"
            for aniversariante in lista_aniversariantes
        )}
                    </ul>
                </div>
                <div class="header">
                    A PRU6 deseja a todos os aniversariantes muita saúde, paz e realizações. Parabéns!
                </div>
            </body>
        </html>
        """

        # Configurando o e-mail
        mail.Subject = f"\n\nAniversariantes do mês {mes}"
        mail.HTMLBody = corpo_html

        # Anexar a imagem FELIZ.PNG
        if os.path.exists(caminho_imagem):
            anexo = mail.Attachments.Add(caminho_imagem)
            anexo.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "FELIZ_PNG")
        else:
            messagebox.showwarning("Aviso", f"Imagem {caminho_imagem} não encontrada!")

        # Lista de destinatários de teste
        destinatarios = "fernando.vilasboas@agu.gov.br"
        mail.to = destinatarios

        # Enviar e-mail
        mail.Send()
        print("\n\nLista de destinatarios\n")
        print(destinatarios)

        messagebox.showinfo("Sucesso", "E-mail enviado com sucesso!")

        # Remove o arquivo Excel após o envio
        if os.path.exists(caminho_excel):
            os.remove(caminho_excel)

        # Encerra o programa
        sys.exit()

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

    # Exibe a lista e envia o e-mail
    caminho_excel = os.path.join(os.path.dirname(sys.executable), "resultado.xlsx")
    enviar_email(lista_aniversariantes, mes, caminho_excel)


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