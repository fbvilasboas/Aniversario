import win32com.client as win32

def gerar_excel_de_iqy(caminho_iqy, caminho_excel):
    try:
        # Inicia o Excel
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False  # Oculta o Excel
        excel.DisplayAlerts = False  # Desativa alertas

        # Abre o arquivo .iqy
        workbook = excel.Workbooks.Open(caminho_iqy)

        # Salva como .xlsx
        workbook.SaveAs(caminho_excel, FileFormat=51)  # 51 = xlsx
        workbook.Close(SaveChanges=False)

        # Fecha o Excel
        excel.Quit()

        print(f"Arquivo Excel gerado com sucesso: {caminho_excel}")
    except Exception as e:
        print(f"Erro ao gerar o arquivo Excel: {e}")
    finally:
        # Garante que o Excel seja fechado
        if 'excel' in locals():
            excel.Quit()

# Caminhos dos arquivos
caminho_iqy = r"C:\Users\fernando.vilasboas\Desktop\NIVER\consulta.iqy"
caminho_excel = r"C:\Users\fernando.vilasboas\Desktop\NIVER\consulta.xlsx"

# Gera o arquivo Excel
gerar_excel_de_iqy(caminho_iqy, caminho_excel)