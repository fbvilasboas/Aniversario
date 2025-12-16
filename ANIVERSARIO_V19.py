# ANIVERSARIO_V19.py
# Patch: ao lado da miniatura, escolher inserir a imagem ANTES ou DEPOIS da lista no corpo do e-mail.

import sys
import os
import time
import re
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import win32com.client as win32

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


def localizar_imagem_padrao() -> str:
    candidatos = [
        os.path.join(get_executable_dir(), "FELIZ_OLD.PNG"),
        os.path.join(get_executable_dir(), "FELIZ_OLD.png"),
        os.path.join(get_output_dir(), "FELIZ_OLD.PNG"),
        os.path.join(get_output_dir(), "FELIZ_OLD.png"),
        os.path.join(os.getcwd(), "FELIZ_OLD.PNG"),
        os.path.join(os.getcwd(), "FELIZ_OLD.png"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "FELIZ_OLD.PNG"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "FELIZ_OLD.png"),
    ]
    for p in candidatos:
        if p and os.path.exists(p):
            return p
    return ""


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

        # 1) Tenta QueryTable por URL
        try:
            url = _extrair_url_do_iqy(caminho_arquivo_iqy)
            wb = excel.Workbooks.Add()
            ws = wb.Worksheets(1)
            qt = ws.QueryTables.Add(Connection=f"URL;{url}", Destination=ws.Range("A1"))
            qt.BackgroundQuery = False
            qt.Refresh(False)
            wb.SaveAs(caminho_arquivo_excel, FileFormat=51)  # xlsx
            return
        except Exception:
            if wb:
                wb.Close(False)
            wb = None

        # 2) Fallback: abrir IQY como Excel espera
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
# Miniatura (UI) — placeholder em pixels
# ============================================================

def atualizar_miniatura(caminho):
    global thumb_photo, thumb_placeholder

    if label_thumb is None:
        return

    if thumb_placeholder is None:
        thumb_placeholder = tk.PhotoImage(width=96, height=96)

    # sempre mantém modo imagem (pixels)
    label_thumb.configure(image=thumb_placeholder, text="")
    thumb_photo = None

    if not caminho or not os.path.exists(caminho):
        return

    try:
        img = Image.open(caminho)
        img.thumbnail((96, 96))
        thumb_photo = ImageTk.PhotoImage(img)
        label_thumb.configure(image=thumb_photo)
    except Exception as e:
        try:
            print("[ERRO miniatura]", e, "arquivo:", caminho)
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
# Envio de e-mail (agora com opção ANTES/DEPOIS da lista)
# ============================================================

def enviar_email(lista_aniversariantes, mes, destinatarios):
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)

    itens = "".join(f"<li>{x}</li>" for x in lista_aniversariantes)

    tem_imagem = bool(caminho_imagem_selecionada) and os.path.exists(caminho_imagem_selecionada)
    imagem_html = "<img src='cid:IMG' width='800' style='display:block;border:0;'>" if tem_imagem else ""

    if tem_imagem and (var_posicao_imagem.get() == "DEPOIS"):
        bloco_imagem_antes = ""
        bloco_imagem_depois = imagem_html
    else:
        bloco_imagem_antes = imagem_html
        bloco_imagem_depois = ""

    corpo = f"""
    <html>
    <body style="background:#f0f0f0">
      <table width="100%">
        <tr><td align="center">
          <table width="800" style="background:#fff">
            <tr>
              <td style="padding:24px;font-family:Arial;font-size:18px">
                {bloco_imagem_antes}
                <h2>ANIVERSARIANTES DO MÊS DE {mes}</h2>
                <ul>{itens}</ul>
                {bloco_imagem_depois}
              </td>
            </tr>
          </table>
        </td></tr>
      </table>
    </body>
    </html>
    """

    mail.Subject = f"Aniversariantes do mês - {mes}"
    mail.HTMLBody = corpo

    if tem_imagem:
        anexo = mail.Attachments.Add(caminho_imagem_selecionada)
        anexo.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "IMG"
        )

    # BCC para não expor lista
    mail.To = ""
    mail.BCC = ";".join(destinatarios)

    mail.Send()

    messagebox.showinfo("Sucesso", f"E-mail enviado com sucesso para {len(destinatarios)} pessoa(s).")
    sys.exit()


# ============================================================
# UI
# ============================================================

def buscar_aniversariantes():
    if dados is None or dados.empty:
        messagebox.showerror("Erro", "Dados não carregados.")
        return

    mes = combobox_mes.get()
    if not mes:
        messagebox.showerror("Erro", "Selecione o mês.")
        return

    mes_num = meses.index(mes) + 1
    aniversariantes = dados[dados["Data Nascimento"].dt.month == mes_num].copy()

    if aniversariantes.empty:
        messagebox.showinfo("Sem aniversariantes", f"Não há aniversariantes para o mês de {mes}.")
        return

    aniversariantes["dia"] = aniversariantes["Data Nascimento"].dt.day
    aniversariantes = aniversariantes.sort_values("dia")

    lista = [
        f"{row['Data Nascimento'].strftime('%d/%m')} : "
        + " ".join(
            w.lower() if w.lower() in {"da", "das", "de", "do", "dos"} else w.title()
            for w in str(row["Nome"]).split()
        )
        for _, row in aniversariantes.iterrows()
        if pd.notna(row["Data Nascimento"]) and pd.notna(row["Nome"])
    ]

    destinatarios = []
    if var_enviar_todos.get():
        destinatarios.extend(lista_emails)

    destinatarios.extend(_normalizar_lista_emails(entry_destinatarios.get()))

    vistos = set()
    final = []
    for e in destinatarios:
        el = e.lower()
        if el not in vistos:
            vistos.add(el)
            final.append(e)

    if not final:
        messagebox.showerror("Erro", "Informe pelo menos um e-mail ou marque 'Enviar para todos'.")
        return

    enviar_email(lista, mes, final)


def exibir_tela_boas_vindas():
    tela = tk.Tk()
    tela.title("Bem-vindo")

    largura = 650
    altura = 320
    largura_tela = tela.winfo_screenwidth()
    altura_tela = tela.winfo_screenheight()
    pos_x = (largura_tela // 2) - (largura // 2)
    pos_y = (altura_tela // 2) - (altura // 2)
    tela.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

    label = tk.Label(
        tela,
        text="Bem-vindo ao programa de envio de lista de aniversariantes!\n\nCarregando dados...",
        font=("Arial", 16),
        justify="center",
    )
    label.pack(expand=True)

    tela.update()
    try:
        carregar_arquivo()
    except Exception as e:
        tela.destroy()
        messagebox.showerror("Erro", f"Falha ao carregar dados do SharePoint:\n{e}")
        return

    tela.destroy()
    exibir_tela_principal()


# ============================================================
# EXIBIR TELA PRINCIPAL
# ============================================================

def exibir_tela_principal():
    global combobox_mes, entry_destinatarios, var_enviar_todos
    global label_imagem_valor, label_thumb, thumb_placeholder
    global caminho_imagem_selecionada, var_posicao_imagem

    root = tk.Tk()
    root.title("Aniversariantes do Mês")

    largura = 860
    altura = 620
    largura_tela = root.winfo_screenwidth()
    altura_tela = root.winfo_screenheight()
    pos_x = (largura_tela // 2) - (largura // 2)
    pos_y = (altura_tela // 2) - (altura // 2)
    root.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")
    root.resizable(False, False)

    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass

    style.configure("Big.TLabel", font=("Arial", 15))
    style.configure("Big.TLabelframe.Label", font=("Arial", 15, "bold"))
    style.configure("Big.TEntry", font=("Arial", 14))
    style.configure("Big.TCombobox", font=("Arial", 14))
    style.configure("Big.TCheckbutton", font=("Arial", 14))
    style.configure("Big.TRadiobutton", font=("Arial", 14))
    style.configure("Big.TButton", font=("Arial", 15, "bold"))

    texto_explicativo = tk.Label(
        root,
        text="Este programa cria a lista de aniversariantes do mês e envia via e-mail.",
        wraplength=820,
        justify="center",
        font=("Arial", 18),
    )
    texto_explicativo.pack(pady=14)

    # Mês
    frame_mes = ttk.Frame(root)
    frame_mes.pack(pady=10)

    ttk.Label(frame_mes, text="Selecione o mês:", style="Big.TLabel").pack(side=tk.LEFT, padx=10)
    combobox_mes = ttk.Combobox(
        frame_mes,
        values=meses,
        state="readonly",
        width=18,
        style="Big.TCombobox",
    )
    combobox_mes.pack(side=tk.LEFT, padx=10)

    # Destinatários
    frame_dest = ttk.LabelFrame(root, text="Destinatários", padding=12, style="Big.TLabelframe")
    frame_dest.pack(fill=tk.X, padx=20, pady=12)

    ttk.Label(frame_dest, text="E-mail(s) manual(is):", style="Big.TLabel").grid(
        row=0, column=0, sticky="w", padx=6, pady=8
    )
    entry_destinatarios = ttk.Entry(frame_dest, width=70, style="Big.TEntry")
    entry_destinatarios.grid(row=0, column=1, sticky="w", padx=6, pady=8)

    ttk.Label(frame_dest, text="(separe por ; ou ,)", style="Big.TLabel").grid(
        row=1, column=1, sticky="w", padx=6
    )

    var_enviar_todos = tk.BooleanVar(value=False)
    ttk.Checkbutton(
        frame_dest,
        text="Enviar para todos (lista do SharePoint)",
        variable=var_enviar_todos,
        style="Big.TCheckbutton",
    ).grid(row=2, column=1, sticky="w", padx=6, pady=10)

    frame_dest.grid_columnconfigure(1, weight=1)

    # Imagem + posição
    frame_img = ttk.LabelFrame(root, text="Imagem do e-mail", padding=12, style="Big.TLabelframe")
    frame_img.pack(fill=tk.X, padx=20, pady=12)

    btn_img = ttk.Button(frame_img, text="Escolher imagem...", command=escolher_imagem, style="Big.TButton")
    btn_img.grid(row=0, column=0, sticky="w", padx=6, pady=8)

    label_thumb = tk.Label(frame_img, bd=1, relief="solid")
    label_thumb.grid(row=0, column=1, sticky="w", padx=10, pady=8)

    thumb_placeholder = tk.PhotoImage(width=96, height=96)
    label_thumb.configure(image=thumb_placeholder)

    label_imagem_valor = ttk.Label(frame_img, text="Nenhuma imagem selecionada", style="Big.TLabel")
    label_imagem_valor.grid(row=0, column=2, sticky="w", padx=10, pady=8)

    # >>> NOVO: escolha ANTES/DEPOIS ao lado da miniatura <<<
    var_posicao_imagem = tk.StringVar(value="ANTES")

    frame_pos = ttk.Frame(frame_img)
    frame_pos.grid(row=0, column=3, sticky="w", padx=10, pady=8)

    ttk.Radiobutton(
        frame_pos,
        text="ANTES DA LISTA",
        value="ANTES",
        variable=var_posicao_imagem,
        style="Big.TRadiobutton",
    ).pack(anchor="w")

    ttk.Radiobutton(
        frame_pos,
        text="DEPOIS DA LISTA",
        value="DEPOIS",
        variable=var_posicao_imagem,
        style="Big.TRadiobutton",
    ).pack(anchor="w")

    frame_img.grid_columnconfigure(2, weight=1)

    # Default: carregar FELIZ_OLD no startup (miniatura + nome)
    caminho_padrao = localizar_imagem_padrao()
    if caminho_padrao:
        caminho_imagem_selecionada = caminho_padrao
        label_imagem_valor.configure(text=os.path.basename(caminho_padrao))
        atualizar_miniatura(caminho_imagem_selecionada)
    else:
        caminho_imagem_selecionada = ""
        atualizar_miniatura("")

    # Botão enviar
    btn_buscar = ttk.Button(root, text="ENVIAR LISTA", command=buscar_aniversariantes, style="Big.TButton")
    btn_buscar.configure(cursor="hand2")
    btn_buscar.pack(pady=22)

    root.mainloop()


# ============================================================
# Init
# ============================================================

meses = [
    "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
    "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"
]

dados = None
lista_emails = []

combobox_mes = None
entry_destinatarios = None
var_enviar_todos = None

caminho_imagem_selecionada = ""
var_posicao_imagem = None

label_thumb = None
label_imagem_valor = None

thumb_photo = None
thumb_placeholder = None

exibir_tela_boas_vindas()
