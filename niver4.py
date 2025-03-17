import os
import win32com.client as win32

def gerar_excel_de_iqy(caminho_arquivo_iqy, caminho_arquivo_excel):
    try:
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
    except Exception as e:
        print(f"Erro ao gerar o arquivo Excel: {e}")

# Caminho do arquivo IQY e do Excel gerado
caminho_iqy = r"C:\Users\fernando.vilasboas\Desktop\NIVER\consulta.iqy"
caminho_excel = r"C:\Users\fernando.vilasboas\Desktop\NIVER\resultado.xlsx"

# Gera o arquivo Excel a partir do arquivo IQY
gerar_excel_de_iqy(caminho_iqy, caminho_excel)

