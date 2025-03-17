from cx_Freeze import setup, Executable
import os

# Configuração do nome do arquivo principal
script_principal = "niver7.py"  # Substitua pelo nome do arquivo do seu programa Python

# Adicione aqui os arquivos necessários, como imagens, templates ou arquivos externos
arquivos_incluidos = [
    #"resultado.xlsx",  # Substitua pelo caminho dos arquivos adicionais
    "consulta.iqy",
]

# Configuração adicional do ambiente
os.environ['TCL_LIBRARY'] = r"C:\Users\fernando.vilasboas\AppData\Local\Programs\Python\Python312\tcl\tcl8.6"  # Substitua pelo caminho correto
os.environ['TK_LIBRARY'] = r"C:\Users\fernando.vilasboas\AppData\Local\Programs\Python\Python312\tcl\tk8.6"    # Substitua pelo caminho correto

# Executável a ser gerado
executables = [
    Executable(
        script=script_principal, 
        #base="Win32GUI",  # Usa a interface gráfica do Windows (não exibe o console)
        target_name="aniversariantes.exe"  # Nome do executável gerado
    )
]

# Configuração do setup
setup(
    name="Aniversariantes do Mês",
    version="1.0",
    description="Programa para enviar lista de aniversariantes via e-mail",
    options={
        "build_exe": {
            "packages": ["os", "sys", "tkinter", "pandas", "win32com.client"],  # Dependências usadas no programa
            "include_files": arquivos_incluidos,  # Arquivos adicionais incluídos
        }
    },
    executables=executables
)
