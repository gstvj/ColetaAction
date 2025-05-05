import os
import platform
import socket
import subprocess
import psutil
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk

# Paleta de cores Action Infinte
COR_FUNDO = "#1c1f2b"
COR_PRIMARIA = "#f97300"
COR_TEXTO = "#f4f4f4"
COR_BOTAO = "#f97300"
COR_BOTAO_TEXTO = "#ffffff"
COR_CAIXA = "#2b2e3b"

def coletar_info(nome_completo, setor):
    info = f"Nome Completo: {nome_completo}\n"
    info += f"Setor: {setor}\n"
    info += f"Nome do Computador: {socket.gethostname()}\n"
    info += f"Sistema Operacional: {platform.system()} {platform.release()} ({platform.version()})\n"

    try:
        serial = subprocess.check_output("wmic bios get serialnumber", shell=True).decode().split("\n")[1].strip()
    except:
        serial = "Não disponível"
    info += f"Número de Série: {serial}\n"

    try:
        modelo = subprocess.check_output("wmic computersystem get model", shell=True).decode().split("\n")[1].strip()
    except:
        modelo = "Não disponível"
    info += f"Modelo do Dispositivo: {modelo}\n"

    try:
        cpu = subprocess.check_output("wmic cpu get name", shell=True).decode().split("\n")[1].strip()
    except:
        cpu = platform.processor()
    info += f"Processador: {cpu}\n"

    ram_total = round(psutil.virtual_memory().total / (1024**3), 2)
    info += f"Memória RAM: {ram_total} GB\n"

    info += "Armazenamento:\n"
    for part in psutil.disk_partitions():
        try:
            usage = psutil.disk_usage(part.mountpoint)
            total = round(usage.total / (1024**3), 2)
            used = round(usage.used / (1024**3), 2)
            free = round(usage.free / (1024**3), 2)
            info += f"  {part.device} - Total: {total} GB | Usado: {used} GB | Livre: {free} GB\n"
        except:
            continue

    info += "Conectividade de Rede:\n"
    try:
        adapters = psutil.net_if_stats()
        for iface, stats in adapters.items():
            speed = "Gigabit" if stats.speed >= 1000 else f"{stats.speed} Mbps"
            info += f"  {iface} - {'Ativo' if stats.isup else 'Inativo'} - {speed}\n"
    except:
        info += "  Não foi possível obter informações de rede.\n"

    return info

def salvar_e_enviar(nome_completo, setor):
    if not nome_completo.strip():
        messagebox.showerror("Erro", "Por favor, digite seu nome e sobrenome.")
        return
    if setor.startswith("➤"):
        messagebox.showerror("Erro", "Por favor, selecione seu setor.")
        return

    data = coletar_info(nome_completo, setor)
    nome_arquivo = f"{nome_completo}-{setor}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    pasta_destino = r"\\loja\Arquivos\z.DDGP\REGISTROS\infos"

    try:
        os.makedirs(pasta_destino, exist_ok=True)
        caminho_final = os.path.join(pasta_destino, nome_arquivo)
        with open(caminho_final, "w", encoding="utf-8") as f:
            f.write(data)
        messagebox.showinfo("Sucesso", f"Informações registradas com sucesso!")
        janela.destroy()
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao salvar arquivo:\n{e}")

# Interface
caminhoicone = r"\\loja\Arquivos\z.DDGP\REGISTROS\.das\favicon.ico"

janela = tk.Tk()
janela.iconbitmap(caminhoicone)
janela.title("Registro de Equipamento - ActiON Infinte")
janela.geometry("450x320")
janela.configure(bg=COR_FUNDO)
janela.resizable(False, False)

style = ttk.Style()
style.theme_use("clam")

# Estilos personalizados
style.configure("Custom.TButton", foreground=COR_BOTAO_TEXTO, background=COR_BOTAO,
                font=("Segoe UI", 10, "bold"), borderwidth=0)
style.map("Custom.TButton",
          background=[("active", "#ff8124"), ("disabled", "#666")])

style.configure("Custom.TCombobox",
                fieldbackground=COR_CAIXA, background=COR_CAIXA,
                foreground=COR_TEXTO, selectforeground=COR_TEXTO,
                selectbackground=COR_CAIXA, arrowcolor=COR_BOTAO)

titulo = tk.Label(janela, text="Registro de PCs/Notebook", bg=COR_FUNDO,
                  fg=COR_PRIMARIA, font=("Segoe UI", 16, "bold"))
titulo.pack(pady=(20, 5))

subtitulo = tk.Label(janela, text="Preencha seu nome sobrenome e setor abaixo",
                     bg=COR_FUNDO, fg=COR_TEXTO, font=("Segoe UI", 10))
subtitulo.pack()

entrada_nome = ttk.Entry(janela, font=("Segoe UI", 11), width=40)
entrada_nome.pack(pady=(10, 10))

# Setores disponíveis
setores = ["➤ Selecione o setor", "Estoque", "Financeiro", "RH", "Suporte", "Compras", "Marketing", "Agendamento", "Comercial", "Recepção", "Pos-venda", "Reversão de cancelamento", "Caixa", "Manutenção carros"]
#var_setor = tk.StringVar(value=setores[0])
#seletor = ttk.Combobox(janela, textvariable=var_setor, values=setores, style="Custom.TCombobox", state="readonly", width=37)
#seletor.pack(pady=(5, 15))

var_setor = tk.StringVar(value=setores[0])
option_menu = tk.OptionMenu(janela, var_setor, *setores)
option_menu.config(width=30, font=("Segoe UI", 10), fg="gray")  # muda a cor do texto
option_menu.pack(pady=(5, 15))

botao = ttk.Button(janela, text="Enviar", style="Custom.TButton",
                   command=lambda: salvar_e_enviar(entrada_nome.get(), var_setor.get()))
botao.pack(pady=(10, 20))

footer = tk.Label(janela, text="v1.0.0 GustavoCarvalho", bg=COR_FUNDO,
                  fg="#666", font=("Segoe UI", 8))
footer.pack(side="bottom", pady=5)

janela.mainloop()
