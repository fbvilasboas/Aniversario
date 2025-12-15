import sys
import os
import time
import re
import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import win32com.client as win32
from datetime import datetime

# para gerar executável
# pyinstaller --onefile -w --add-data "consulta.iqy;." --add-data "FELIZ_OLD.PNG;." aniversario_ok.py


def get_executable_dir():
    if getattr(sys, "frozen", False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def get_output_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _extrair_url_do_iqy(caminho_arquivo_iqy: str) -> str:
    with open(caminho_arquivo_iqy, "r", encoding="utf-8", errors="ignore") as f:
        for linha in f:
            linha = linha.strip()
            if linha.lower().startswith("http"):
                return linha
    raise ValueError("Não encontrei a URL dentro do arquivo .iqy (formato inesperado).")


def gerar_excel_de_iqy(caminho_arquivo_iqy, caminho_arquivo_excel, mostrar_excel=False):
    caminho_arquivo_iqy = os.path.abspath(caminho_arquivo_iqy)
    caminho_arquivo_excel = os.path.abspath(caminho_arquivo_excel)

    if not os.path.exists(caminho_arquivo_iqy):
        raise FileNotFoundError(f"O arquivo {caminho_arquivo_iqy} não foi encontrado.")

    if os.path.exists(caminho_arquivo_excel):
        try:
            os.remove(caminho_arquivo_excel)
        except PermissionError:
            raise PermissionError(f"Feche o arquivo '{caminho_arquivo_excel}' (ele está em uso).")

    excel = None
    wb = None

    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = bool(mostrar_excel)
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False

        try:
            excel.AutomationSecurity = 1  # msoAutomationSecurityLow
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
            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
            except Exception:
                pass
            wb = None

        # 2) Fallback: abrir o IQY como o Excel espera
        wb = excel.Workbooks.Open(caminho_arquivo_iqy)

        try:
            wb.RefreshAll()
        except Exception:
            pass

        try:
            excel.CalculateUntilAsyncQueriesDone()
        except Exception:
            time.sleep(8)

        wb.SaveAs(caminho_arquivo_excel, FileFormat=51)

    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass


def carregar_arquivo():
    global dados, lista_emails

    caminho_iqy = os.path.join(get_executable_dir(), "consulta.iqy")
    caminho_excel = os.path.join(get_output_dir(), "resultado.xlsx")

    # Silencioso (sem janela do Excel)
    gerar_excel_de_iqy(caminho_iqy, caminho_excel, mostrar_excel=False)

    try:
        if not os.path.exists(caminho_excel):
            raise FileNotFoundError(f"O arquivo {caminho_excel} não foi encontrado.")

        dados = pd.read_excel(caminho_excel)

        colunas_necessarias = {"Nome", "Data Nascimento", "E-mail"}
        colunas_arquivo = set(dados.columns)

        if not colunas_necessarias.issubset(colunas_arquivo):
            raise ValueError(
                f"O arquivo deve conter as colunas: {', '.join(colunas_necessarias)}.\n"
                f"Colunas encontradas: {', '.join(colunas_arquivo)}"
            )

        dados["Data Nascimento"] = pd.to_datetime(dados["Data Nascimento"], errors="coerce")
        lista_emails = dados["E-mail"].dropna().tolist()

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar o arquivo: {str(e)}")
        dados = None
        lista_emails = []


def _normalizar_lista_emails(texto: str) -> list[str]:
    if not texto:
        return []

    partes = re.split(r"[;,\s]+", texto.strip())
    emails = [p.strip() for p in partes if p.strip()]

    vistos = set()
    saida = []
    for e in emails:
        el = e.lower()
        if el not in vistos:
            vistos.add(el)
            saida.append(e)
    return saida


def escolher_imagem():
    global caminho_imagem_selecionada

    root = tk.Tk()
    root.attributes("-topmost", True)
    root.withdraw()

    caminho = filedialog.askopenfilename(
        title="Selecione a imagem para anexar",
        filetypes=[
            ("Imagens", "*.png;*.jpg;*.jpeg;*.gif;*.bmp"),
            ("Todos os arquivos", "*.*"),
        ],
        parent=root,
    )
    root.destroy()

    if caminho:
        caminho_imagem_selecionada = caminho
        label_imagem_valor.configure(text=os.path.basename(caminho_imagem_selecionada))
    else:
        caminho_imagem_selecionada = ""
        label_imagem_valor.configure(text="Nenhuma imagem selecionada")


# --- SUBSTITUA a função enviar_email inteira por esta versão ---
# (o resto do seu programa pode ficar igual)

def enviar_email(lista_aniversariantes, mes, caminho_excel, destinatarios: list[str], caminho_imagem: str):
    try:
        outlook = win32.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)
        mes = mes.upper()

        cid_img = "IMG_TOPO"

        itens_html = "".join(
            f"<li><b>{a.split('.')[0]}</b>: {a.split('.')[1]}</li>"
            for a in lista_aniversariantes
        )

        corpo_html = f"""
        <html>
        <head>
          <meta charset="utf-8">
        </head>
        <body style="margin:0;padding:0;background-color:#f0f0f0;">
          <table width="100%" cellpadding="0" cellspacing="0" border="0">
            <tr>
              <td align="center">

                <!-- IMAGEM TOPO -->
                {"<img src='cid:IMG_TOPO' width='800' style='max-width:900%;display:block;border:0;'>" if caminho_imagem else ""}

                <!-- CONTEÚDO -->
                <table width="800" cellpadding="0" cellspacing="0" border="0"
                       style="max-width:900px;background:#ffffff;">
                  <tr>
                    <td style="padding:24px;font-family:Georgia,sans-serif;color:#111;">
                      <div style="font-size:22px;font-weight:bold;color:#13348d;
                                  margin-bottom:16px;text-align:center;">
                        ANIVERSARIANTES DO MÊS DE {mes}
                      </div>

                      <ul style="font-family:Arial,sans-serif;
                                 font-size:18px;line-height:1.6;
                                 padding-left:20px;margin:0;">
                        {itens_html}
                      </ul>
                    </td>
                  </tr>
                </table>

              </td>
            </tr>
          </table>
        </body>
        </html>
        """

        mail.Subject = f"Aniversariantes do mês - {mes}"
        mail.HTMLBody = corpo_html

        # Anexa imagem como inline (topo)
        if caminho_imagem and os.path.exists(caminho_imagem):
            anexo = mail.Attachments.Add(caminho_imagem)
            anexo.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid_img
            )

        # BCC para não expor lista
        mail.To = ""
        mail.BCC = ";".join(destinatarios)

        mail.Send()

        if os.path.exists(caminho_excel):
            os.remove(caminho_excel)

        qtd = len(destinatarios)
        messagebox.showinfo("Sucesso",f"E-mail enviado com sucesso para {qtd} pessoa(s)."
        )

        sys.exit()

    except Exception as e:
        messagebox.showerror("Erro", str(e))


def buscar_aniversariantes():
    if dados is None:
        messagebox.showerror("Erro", "Os dados do arquivo ainda não foram carregados.")
        return

    mes = combobox_mes.get()
    if not mes:
        messagebox.showerror("Erro", "Por favor, selecione um mês.")
        return

    mes_numero = meses.index(mes) + 1
    aniversariantes_mes = dados[dados["Data Nascimento"].dt.month == mes_numero]

    if aniversariantes_mes.empty:
        messagebox.showinfo("Sem aniversariantes", f"Não há aniversariantes para o mês de {mes}.")
        return

    aniversariantes_mes = aniversariantes_mes.sort_values(by="Data Nascimento", key=lambda x: x.dt.day)

    lista_aniversariantes = [
        f"     {row['Data Nascimento'].strftime('%d/%m')}  . "
        + " ".join(
            word.lower() if word.lower() in {"da", "das", "de", "do", "dos"} else word.title()
            for word in row["Nome"].split()
        )
        for _, row in aniversariantes_mes.iterrows()
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

    caminho_excel = os.path.join(get_output_dir(), "resultado.xlsx")
    enviar_email(lista_aniversariantes, mes, caminho_excel, final, caminho_imagem_selecionada.strip())


def exibir_tela_boas_vindas():
    tela_boas_vindas = tk.Tk()
    tela_boas_vindas.title("Bem-vindo")

    largura = 600
    altura = 300

    largura_tela = tela_boas_vindas.winfo_screenwidth()
    altura_tela = tela_boas_vindas.winfo_screenheight()
    pos_x = (largura_tela // 2) - (largura // 2)
    pos_y = (altura_tela // 2) - (altura // 2)
    tela_boas_vindas.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

    label_boas_vindas = tk.Label(
        tela_boas_vindas,
        text="Bem-vindo ao programa de envio de lista de aniversariantes!\n\n\nCarregando dados.",
        font=("Arial", 14),
        justify="center",
    )
    label_boas_vindas.pack(expand=True)

    tela_boas_vindas.update()
    carregar_arquivo()
    tela_boas_vindas.destroy()

    exibir_tela_principal()


def exibir_tela_principal():
    global combobox_mes, entry_destinatarios, var_enviar_todos
    global caminho_imagem_selecionada, label_imagem_valor

    root = tk.Tk()
    root.title("Aniversariantes do Mês")

    largura = 760
    altura = 500

    largura_tela = root.winfo_screenwidth()
    altura_tela = root.winfo_screenheight()
    pos_x = (largura_tela // 2) - (largura // 2)
    pos_y = (altura_tela // 2) - (altura // 2)
    root.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

    root.resizable(False, False)

    texto_explicativo = tk.Label(
        root,
        text="Este programa cria a lista de aniversariantes do mês e envia via e-mail.",
        wraplength=700,
        justify="center",
        font=("Arial", 14),
    )
    texto_explicativo.pack(pady=12)

    # Mês
    frame_mes = ttk.Frame(root)
    frame_mes.pack(pady=8)

    ttk.Label(frame_mes, text="Selecione o mês:", font=("Verdana", 12)).pack(side=tk.LEFT, padx=8)
    combobox_mes = ttk.Combobox(frame_mes, values=meses, state="readonly", width=18)
    combobox_mes.pack(side=tk.LEFT, padx=8)

    # Destinatários
    frame_dest = ttk.LabelFrame(root, text="Destinatários", padding=10)
    frame_dest.pack(fill=tk.X, padx=20, pady=10)

    ttk.Label(frame_dest, text="E-mail(s) manual(is):").grid(row=0, column=0, sticky="w", padx=5, pady=6)
    entry_destinatarios = ttk.Entry(frame_dest, width=70)
    entry_destinatarios.grid(row=0, column=1, sticky="w", padx=5, pady=6)
    ttk.Label(frame_dest, text="(separe por ; ou ,)").grid(row=1, column=1, sticky="w", padx=5)

    var_enviar_todos = tk.BooleanVar(value=False)
    ttk.Checkbutton(
        frame_dest,
        text="Enviar para todos (lista do SharePoint)",
        variable=var_enviar_todos,
    ).grid(row=2, column=1, sticky="w", padx=5, pady=8)

    frame_dest.grid_columnconfigure(1, weight=1)

    # Imagem
    frame_img = ttk.LabelFrame(root, text="Imagem do e-mail", padding=10)
    frame_img.pack(fill=tk.X, padx=20, pady=10)

    ttk.Button(frame_img, text="Escolher imagem...", command=escolher_imagem).grid(
        row=0, column=0, sticky="w", padx=5, pady=6
    )

    label_imagem_valor = ttk.Label(frame_img, text="Nenhuma imagem selecionada")
    label_imagem_valor.grid(row=0, column=1, sticky="w", padx=10, pady=6)

    # default: se existir a antiga, já preenche
    caminho_padrao = os.path.join(get_executable_dir(), "FELIZ_OLD.PNG")
    if os.path.exists(caminho_padrao):
        caminho_imagem_selecionada = caminho_padrao
        label_imagem_valor.configure(text=os.path.basename(caminho_imagem_selecionada))
    else:
        caminho_imagem_selecionada = ""

    frame_img.grid_columnconfigure(1, weight=1)

    # Botão enviar
    btn_buscar = ttk.Button(root, text="ENVIAR LISTA", command=buscar_aniversariantes)
    btn_buscar.configure(cursor="hand2")
    btn_buscar.pack(pady=18)

    root.mainloop()


meses = [
    "JANEIRO",
    "FEVEREIRO",
    "MARÇO",
    "ABRIL",
    "MAIO",
    "JUNHO",
    "JULHO",
    "AGOSTO",
    "SETEMBRO",
    "OUTUBRO",
    "NOVEMBRO",
    "DEZEMBRO",
]

dados = None
lista_emails = []

combobox_mes = None
entry_destinatarios = None
var_enviar_todos = None

caminho_imagem_selecionada = ""
label_imagem_valor = None

exibir_tela_boas_vindas()
