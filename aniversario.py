import sys
import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk
import win32com.client as win32
from datetime import datetime

# para gerar executável
#pyinstaller --onefile -w --add-data "consulta.iqy;." --add-data "FELIZ.PNG;." aniversario.py

# Função para obter o diretório do executável ou do script
def get_executable_dir():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS  # Caminho temporário onde os arquivos embutidos estão
    else:
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
        #excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False  # Executa o Excel em segundo plano (sem exibir a interface)
        excel.DisplayAlerts = False  # Desativa alertas do Excel



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
                        font-family: Georgia, sans-serif;
                        color: #333;
                        background-color: #f0f0f0;
                        padding: 20px;
                    }}
                    .header {{
                        font-family: Georgia, sans-serif;
                        font-size: 24px;
                        font-weight: bold;
                        color: #13348d;
                        margin-bottom: 20px;
                    }}
                    .lista {{
                        font-family: Sans-serif, sans-serif;
                        font-size: 18px;
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
                    Cumprimente os colegas aniversariantes do mês de {mes}:
                </div>
                <div class="lista">
                    <ul>
                        {''.join(
            f"<li><b>{aniversariante.split('........')[0]}</b>: {aniversariante.split('........')[1]}</li>"
            for aniversariante in lista_aniversariantes
        )}
                    </ul>
                </div>
                <img src="cid:FELIZ_PNG" alt="Imagem de felicitações" width="750" height="350">
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

        # Adicionar todos os e-mails na lista de destinatários
        #destinatarios = ";".join(lista_emails)

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

    # Formata a lista de aniversariantes (copia igual da lista)
    #lista_aniversariantes = [
    #    f"     {row['Data Nascimento'].strftime('%d/%m')}  ........ {row['Nome']}"
    #    for _, row in aniversariantes_mes.iterrows()
    #]


    # Formata a lista de aniversariantes  (nomes apareçam com a primeira letra em maiúscula , sendo que "da" "das" "de", "do" e "dos" estejam em minúsculo

    lista_aniversariantes = [
        f"     {row['Data Nascimento'].strftime('%d/%m')}  ........ " +
        " ".join(word.lower() if word.lower() in {"da","das","de","do","dos"} else word.title() for word in row['Nome'].split())
        for _, row in aniversariantes_mes.iterrows()
    ]

    # Exibe a lista e envia o e-mail
    caminho_excel = os.path.join(os.path.dirname(sys.executable), "resultado.xlsx")
    enviar_email(lista_aniversariantes, mes, caminho_excel)


# Função para exibir a tela de boas-vindas
def exibir_tela_boas_vindas():
    # Cria a janela de boas-vindas
    tela_boas_vindas = tk.Tk()
    tela_boas_vindas.title("Bem-vindo")

    # Define o tamanho da janela (400x200 pixels)
    largura = 600
    altura = 300

    # Centraliza a janela
    largura_tela = tela_boas_vindas.winfo_screenwidth()
    altura_tela = tela_boas_vindas.winfo_screenheight()
    pos_x = (largura_tela // 2) - (largura // 2)
    pos_y = (altura_tela // 2) - (altura // 2)
    tela_boas_vindas.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

    # Adiciona um rótulo de boas-vindas
    label_boas_vindas = tk.Label(
        tela_boas_vindas,
        text="Bem-vindo ao programa de envio de lista de aniversariantes!\n\n\nCarregando dados...",
        font=("Arial", 14),
        justify="center"
    )
    label_boas_vindas.pack(expand=True)

    # Atualiza a interface para exibir a tela de boas-vindas
    tela_boas_vindas.update()

    # Carrega os dados em segundo plano
    carregar_arquivo()

    # Fecha a tela de boas-vindas após o carregamento
    tela_boas_vindas.destroy()

    # Exibe a tela principal
    exibir_tela_principal()


# Função para exibir a tela principal
def exibir_tela_principal():
    global combobox_mes

    # Cria a janela principal
    root = tk.Tk()
    root.title("Aniversariantes do Mês")

    # Define o tamanho da janela (600x300 pixels)
    largura = 600
    altura = 300

    # Centraliza a janela
    largura_tela = root.winfo_screenwidth()
    altura_tela = root.winfo_screenheight()
    pos_x = (largura_tela // 2) - (largura // 2)
    pos_y = (altura_tela // 2) - (altura // 2)
    root.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

    root.resizable(False, False)  # Bloqueia redimensionamento

    # Texto explicativo
    texto_explicativo = tk.Label(
        root,
        text="Este programa cria a lista de aniversariantes do mês e envia para todos os servidores da PRU6 via e-mail.",
        wraplength=500,
        justify="center",
        font=("Arial", 14)
    )
    texto_explicativo.pack(pady=20)

    # Combobox e Label para seleção do mês
    label = tk.Label(root, text="Selecione o mês:", font=("Verdana", 14))
    label.pack(pady=20)

    # Cria o Combobox
    combobox_mes = ttk.Combobox(root, values=meses, state="readonly", width=20)
    combobox_mes.pack(pady=10)

    # Botão para buscar aniversariantes
    btn_buscar = ttk.Button(root, text="ENVIAR LISTA", command=buscar_aniversariantes)
    btn_buscar.configure(cursor="hand2")
    btn_buscar.pack(pady=20)

    # Inicia o loop da interface gráfica
    root.mainloop()


# Lista de meses
meses = [
    "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
    "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"
]

# Variável global para armazenar os dados lidos
dados = None
lista_emails = []

# Exibe a tela de boas-vindas ao iniciar o programa
exibir_tela_boas_vindas()