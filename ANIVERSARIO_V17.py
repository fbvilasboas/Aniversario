import sys
import os
import time
import re
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import win32com.client as win32
from datetime import datetime

from PIL import Image, ImageTk  # miniatura só na UI


# ============================================================
# Utilidades de caminho (compatível PyCharm / PyInstaller)
# ============================================================

def fixar_cwd():
    if getattr(sys, "frozen", False):
        os.chdir(os.path.dirname(sys.executable))
    else:
        os.chdir(os.path.dirname(os.path.abspath(__file__)))

fixar_cwd()


def get_executable_dir():
    if getattr(sys, "frozen", False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def get_output_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


# ============================================================
# IQY → Excel (silencioso)
# ============================================================

def _extrair_url_do_iqy(caminho_arquivo_iqy: str) -> str:
    with open(caminho_arquivo_iqy, "r", encoding="utf-8", errors="ignore") as f:
        for linha in f:
            linha = linha.strip()
            if linha.lower().startswith("http"):
                return linha
    raise ValueError("URL não encontrada dentro do arquivo .iqy")


def gerar_excel_de_iqy(caminho_arquivo_iqy, caminho_arquivo_excel):
    excel = None
    wb = None

    if os.path.exists(caminho_arquivo_excel):
        os.remove(caminho_arquivo_excel)

    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False

        try:
            excel.AutomationSecurity = 1
        except Exception:
            pass

        try:
            url = _extrair_url_do_iqy(caminho_arquivo_iqy)
            wb = excel.Workbooks.Add()
            ws = wb.Worksheets(1)
            qt = ws.QueryTables.Add(Connection=f"URL;{url}", Destination=ws.Range("A1"))
            qt.BackgroundQuery = False
            qt.Refresh(False)
            wb.SaveAs(caminho_arquivo_excel, FileFormat=51)
            return
        except Exception:
            if wb:
                wb.Close(False)
            wb = None

        wb = excel.Workbooks.Open(caminho_arquivo_iqy)
        wb.RefreshAll()

        try:
            excel.CalculateUntilAsyncQueriesDone()
        except Exception:
            time.sleep(8)

        wb.SaveAs(caminho_arquivo_excel, FileFormat=51)

    finally:
        if wb:
            wb.Close(False)
        if excel:
            excel.Quit()


# ============================================================
# Dados
# ============================================================

def carregar_arquivo():
    global dados, lista_emails

    caminho_iqy = os.path.join(get_executable_dir(), "consulta.iqy")
    caminho_excel = os.path.join(get_output_dir(), "resultado.xlsx")

    gerar_excel_de_iqy(caminho_iqy, caminho_excel)

    dados = pd.read_excel(caminho_excel)
    dados["Data Nascimento"] = pd.to_datetime(dados["Data Nascimento"], errors="coerce")
    lista_emails = dados["E-mail"].dropna().astype(str).tolist()


def _normalizar_lista_emails(texto):
    if not texto:
        return []
    partes = re.split(r"[;,\s]+", texto.strip())
    vistos, saida = set(), []
    for e in partes:
        if e and e.lower() not in vistos:
            vistos.add(e.lower())
            saida.append(e)
    return saida


# ============================================================
# Miniatura (UI) — PATCH CRÍTICO
# ============================================================

def atualizar_miniatura(caminho):
    global thumb_photo, thumb_placeholder

    if label_thumb is None:
        return

    # garante placeholder (pixels)
    if thumb_placeholder is None:
        thumb_placeholder = tk.PhotoImage(width=96, height=96)

    label_thumb.configure(image=thumb_placeholder, text="")
    thumb_photo = None

    if not caminho or not os.path.exists(caminho):
        return

    try:
        img = Image.open(caminho)
        img.thumbnail((96, 96))
        thumb_photo = ImageTk.PhotoImage(img)
        label_thumb.configure(image=thumb_photo)
    except Exception:
        pass


def escolher_imagem():
    global caminho_imagem_selecionada

    root = tk.Tk()
    root.attributes("-topmost", True)
    root.withdraw()

    caminho = filedialog.askopenfilename(
        title="Selecione a imagem",
        filetypes=[("Imagens", "*.png;*.jpg;*.jpeg;*.gif;*.bmp")],
        parent=root,
    )
    root.destroy()

    if caminho:
        caminho_imagem_selecionada = caminho
        label_imagem_valor.configure(text=os.path.basename(caminho))
        atualizar_miniatura(caminho)
    else:
        caminho_imagem_selecionada = ""
        label_imagem_valor.configure(text="Nenhuma imagem selecionada")
        atualizar_miniatura("")


# ============================================================
# Envio de e-mail (NÃO ALTERADO)
# ============================================================

def enviar_email(lista_aniversariantes, mes, destinatarios):
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)

    itens = "".join(f"<li>{x}</li>" for x in lista_aniversariantes)

    corpo = f"""
    <html>
    <body style="background:#f0f0f0">
      <table width="100%">
        <tr><td align="center">
          {"<img src='cid:IMG' width='800'>" if caminho_imagem_selecionada else ""}
          <table width="800" style="background:#fff">
            <tr><td style="padding:24px;font-family:Arial;font-size:18px">
              <h2>ANIVERSARIANTES DO MÊS DE {mes}</h2>
              <ul>{itens}</ul>
            </td></tr>
          </table>
        </td></tr>
      </table>
    </body>
    </html>
    """

    mail.Subject = f"Aniversariantes do mês - {mes}"
    mail.HTMLBody = corpo

    if caminho_imagem_selecionada:
        anexo = mail.Attachments.Add(caminho_imagem_selecionada)
        anexo.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "IMG"
        )

    mail.BCC = ";".join(destinatarios)
    mail.Send()

    messagebox.showinfo(
        "Sucesso",
        f"E-mail enviado com sucesso para {len(destinatarios)} pessoa(s)."
    )
    sys.exit()


# ============================================================
# UI
# ============================================================

def buscar_aniversariantes():
    mes = combobox_mes.get()
    mes_num = meses.index(mes) + 1

    aniversariantes = dados[dados["Data Nascimento"].dt.month == mes_num]

    lista = [
        f"{row['Data Nascimento'].strftime('%d/%m')} : "
        + " ".join(
            w.lower() if w.lower() in {"da", "das", "de", "do", "dos"} else w.title()
            for w in row["Nome"].split()
        )
        for _, row in aniversariantes.iterrows()
    ]

    dest = []
    if var_enviar_todos.get():
        dest.extend(lista_emails)
    dest.extend(_normalizar_lista_emails(entry_destinatarios.get()))

    enviar_email(lista, mes, dest)

def exibir_tela_boas_vindas():
    tela_boas_vindas = tk.Tk()
    tela_boas_vindas.title("Bem-vindo")

    largura = 650
    altura = 320

    largura_tela = tela_boas_vindas.winfo_screenwidth()
    altura_tela = tela_boas_vindas.winfo_screenheight()
    pos_x = (largura_tela // 2) - (largura // 2)
    pos_y = (altura_tela // 2) - (altura // 2)
    tela_boas_vindas.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

    label_boas_vindas = tk.Label(
        tela_boas_vindas,
        text="Bem-vindo ao programa de envio de lista de aniversariantes!\n\nCarregando dados...",
        font=("Arial", 16),
        justify="center",
    )
    label_boas_vindas.pack(expand=True)

    tela_boas_vindas.update()
    carregar_arquivo()
    tela_boas_vindas.destroy()

    exibir_tela_principal()


def exibir_tela_principal():
    global combobox_mes, entry_destinatarios, var_enviar_todos
    global label_imagem_valor, label_thumb, thumb_placeholder

    root = tk.Tk()
    root.title("Aniversariantes do Mês")
    root.geometry("860x560")
    root.resizable(False, False)

    tk.Label(
        root,
        text="Este programa cria a lista de aniversariantes do mês e envia via e-mail.",
        font=("Arial", 20),
        wraplength=820,
        justify="center"
    ).pack(pady=16)

    frame_mes = ttk.Frame(root)
    frame_mes.pack(pady=8)
    ttk.Label(frame_mes, text="Selecione o mês:", font=("Arial", 14, "bold")).pack(side=tk.LEFT, padx=10)
    combobox_mes = ttk.Combobox(frame_mes, values=meses, width=22, state="readonly", font=("Arial", 14))
    combobox_mes.pack(side=tk.LEFT)

    frame_dest = ttk.LabelFrame(root, text="Destinatários", padding=12)
    frame_dest.pack(fill=tk.X, padx=20, pady=12)

    ttk.Label(frame_dest, text="E-mail(s) manual(is):", font=("Arial", 13)).grid(row=0, column=0, padx=8, pady=8)
    entry_destinatarios = ttk.Entry(frame_dest, width=52, font=("Arial", 13))
    entry_destinatarios.grid(row=0, column=1, padx=8, pady=8)
    ttk.Label(frame_dest, text="(separe por ; ou ,)", font=("Arial", 13)).grid(row=1, column=1, sticky="w")

    var_enviar_todos = tk.BooleanVar(value=False)
    ttk.Checkbutton(frame_dest, text="Enviar para todos (lista do SharePoint)", variable=var_enviar_todos)\
        .grid(row=2, column=1, sticky="w", padx=8, pady=8)

    frame_img = ttk.LabelFrame(root, text="Imagem do e-mail", padding=12)
    frame_img.pack(fill=tk.X, padx=20, pady=12)

    ttk.Button(frame_img, text="Escolher imagem...", command=escolher_imagem)\
        .grid(row=0, column=0, padx=10, pady=10)

    label_thumb = tk.Label(frame_img, bd=1, relief="solid")
    label_thumb.grid(row=0, column=1, padx=18, pady=8)

    thumb_placeholder = tk.PhotoImage(width=96, height=96)
    label_thumb.configure(image=thumb_placeholder)

    label_imagem_valor = ttk.Label(frame_img, text="Nenhuma imagem selecionada", font=("Arial", 16))
    label_imagem_valor.grid(row=0, column=2, padx=14)

    btn = ttk.Button(root, text="ENVIAR LISTA", command=buscar_aniversariantes)
    btn.pack(pady=18)

    caminho_padrao = os.path.join(get_executable_dir(), "FELIZ_OLD.PNG")
    if os.path.exists(caminho_padrao):
        atualizar_miniatura(caminho_padrao)
        label_imagem_valor.configure(text="FELIZ_OLD.PNG")

    root.mainloop()


# ============================================================
# Init
# ============================================================

meses = [
    "JANEIRO","FEVEREIRO","MARÇO","ABRIL","MAIO","JUNHO",
    "JULHO","AGOSTO","SETEMBRO","OUTUBRO","NOVEMBRO","DEZEMBRO"
]

dados = None
lista_emails = []
caminho_imagem_selecionada = ""
label_thumb = None
label_imagem_valor = None
thumb_photo = None
thumb_placeholder = None

carregar_arquivo()
exibir_tela_boas_vindas()
