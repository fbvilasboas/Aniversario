# ANIVERSARIANTES DO MÊS

Um programa em Python com objetivo de facilitar o trabalho do gabinete ao informar a lista dos aniversariantes do mês, via e-mail, para todos os integrantes da PRU6 

## Arquivo de entrada

O programa usa os dados da lista de integrantes da PRU6 para buscar os aniversariantes do mês: 
Arquivo sharepoint no endereço: https://agudf.sharepoint.com/sites/PRU6/Lists/Quadro%20de%20pessoal%20PRU6%20%20apenas%20PSUs/AllItems.aspx?env=WebViewList&OR=Teams-HL&CT=1736945623048&clickparams=eyJBcHBOYW1lIjoiVGVhbXMtRGVza3RvcCIsIkFwcFZlcnNpb24iOiI0OS8yNDEyMDEwMDIxMSIsIkhhc0ZlZGVyYXRlZFVzZXIiOmZhbHNlfQ%3D%3D

Para ler a lista acima, o programa usa um arquivo consulta.iqyy

Esse arquivo consulta.iqy extrai os dados para uma planilha excel temporária "resultado.xlsx": 

## Instalação do Pyinstaler

Para gerar o executável desta aplicação é necessário instalar o **pyinstaller**:
    python.exe -m pip install --upgrade pip
    python -m pip install pyinstaller

Após instalação verifique se o pyinstaller está no *%PATH%*.

O pyinstaller precisará das bibliotecas e dependências do robô instaladas para gerar o binário corretamente.
  
É recomendado o uso de um ambiente virtual python para evitar a interferência na configuração da máquina.

Para gerar um ambiente python virtual utilize o comando a seguir:

    python -m venv .venv 
   .\.venv\Scripts\activate

As dependências usadas pelo pyinstaller estão no arquivo **'requirements.txt'**. Elas podem ser instaladas usando o comando abaixo.

    python -m pip install -r requirements-dev.txt

Se novas bibliotecas forem inclusas no projeto, pode ser necessário atualizar o arquivo requirements.txt. Se for esse o caso atualize com o comando abaixo:

    python -m pip freeze > requirements-dev.txt

## Geração do executável

Para gerar o executável, rode o comando a seguir:

pyinstaller --onefile -w --add-data "consulta.iqy;." --add-data "FELIZ.PNG;." aniversario_ok.py


O Comando pyinstaller gera 2 pastas, './build' e './dist/'.
A pasta build contém arquivos intermediários criados na construção do binário.
A pasta './dist/' contém o binário final gerado.
