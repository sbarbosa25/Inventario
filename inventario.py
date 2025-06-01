import csv
import platform
import socket
import getpass
import psutil
import datetime
import unicodedata

try:
    import wmi
    _wmi_available = True
except ImportError:
    _wmi_available = False

import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox

def remover_acentos(texto: str) -> str:
    """
    Remove acentuação de um texto, retornando apenas caracteres ASCII equivalentes.
    """
    nfkd = unicodedata.normalize('NFKD', texto)
    return "".join([c for c in nfkd if not unicodedata.combining(c)])

def coletar_informacoes(progresso_callback=None):
    """
    Coleta diversas informações do sistema e retorna uma lista de tuplas (chave, valor),
    já sem acentuação.
    A cada etapa concluída, chama progresso_callback(etapa_atual) para atualizar a barra.
    """
    info = []
    etapas = [
        "data_hora", "sistema_host", "cpu", "memoria", "disco",
        "temperatura", "programas", "usb", "ip", "interfaces", "serial"
    ]
    total_etapas = len(etapas)
    etapa_atual = 0

    # 1) Data e hora
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    info.append(("Data e Hora da Coleta", now))
    etapa_atual += 1
    if progresso_callback: progresso_callback(etapa_atual, total_etapas)

    # 2) Sistema e Host
    info.append(("Nome do Host", socket.gethostname()))
    info.append(("Usuário Logado", getpass.getuser()))
    info.append(("Sistema Operacional", f"{platform.system()} {platform.release()} ({platform.version()})"))
    info.append(("Arquitetura", platform.machine()))
    info.append(("Processador (model string)", platform.processor()))
    etapa_atual += 1
    if progresso_callback: progresso_callback(etapa_atual, total_etapas)

    # 3) CPU
    try:
        cpu_fisico = psutil.cpu_count(logical=False)
        cpu_total = psutil.cpu_count(logical=True)
        cpu_freq = psutil.cpu_freq()
        info.append(("Cores Físicos", str(cpu_fisico)))
        info.append(("Cores Lógicos", str(cpu_total)))
        if cpu_freq:
            info.append(("Frequência Máxima do CPU (MHz)", f"{cpu_freq.max:.2f}"))
            info.append(("Frequência Atual do CPU (MHz)", f"{cpu_freq.current:.2f}"))
    except Exception as e:
        info.append(("Erro ao coletar CPU", str(e)))
    etapa_atual += 1
    if progresso_callback: progresso_callback(etapa_atual, total_etapas)

    # 4) Memória RAM
    try:
        vm = psutil.virtual_memory()
        total_ram_gb = vm.total / (1024 ** 3)
        info.append(("Memória RAM Total (GB)", f"{total_ram_gb:.2f}"))
        info.append(("Memória RAM Disponível (GB)", f"{vm.available / (1024 ** 3):.2f}"))
        info.append(("Uso de RAM (%)", f"{vm.percent}"))
    except Exception as e:
        info.append(("Erro ao coletar Memória RAM", str(e)))
    etapa_atual += 1
    if progresso_callback: progresso_callback(etapa_atual, total_etapas)

    # 5) Espaço em Disco
    try:
        discos = psutil.disk_partitions(all=False)
        for part in discos:
            try:
                uso = psutil.disk_usage(part.mountpoint)
                total_gb = uso.total / (1024 ** 3)
                livre_gb = uso.free / (1024 ** 3)
                info.append((f"Disco: {part.device} - Montagem: {part.mountpoint}", ""))
                info.append(("  Total (GB)", f"{total_gb:.2f}"))
                info.append(("  Livre (GB)", f"{livre_gb:.2f}"))
                info.append(("  Uso (%)", f"{uso.percent}"))
            except PermissionError:
                continue
    except Exception as e:
        info.append(("Erro ao coletar Disco", str(e)))
    etapa_atual += 1
    if progresso_callback: progresso_callback(etapa_atual, total_etapas)

    # 6) Temperaturas
    try:
        temps = psutil.sensors_temperatures()
        if temps:
            for nome_sensor, entradas in temps.items():
                for entrada in entradas:
                    label = entrada.label if entrada.label else nome_sensor
                    info.append((f"Temperatura [{label}]", f"{entrada.current:.1f} °C"))
        else:
            info.append(("Temperaturas", "Não disponível neste sistema"))
    except Exception as e:
        info.append(("Erro ao coletar Temperatura", str(e)))
    etapa_atual += 1
    if progresso_callback: progresso_callback(etapa_atual, total_etapas)

    # 7) Programas Instalados (Windows via WMI)
    programas = []
    if _wmi_available and platform.system() == "Windows":
        try:
            c = wmi.WMI()
            for produto in c.Win32_Product():
                nome = produto.Name or ""
                versao = produto.Version or ""
                if versao:
                    programas.append(f"{nome} ({versao})")
                else:
                    programas.append(nome)
            if programas:
                info.append(("Programas Instalados (via WMI)", "; ".join(programas)))
            else:
                info.append(("Programas Instalados (via WMI)", "Nenhum encontrado"))
        except Exception as e:
            info.append(("Erro ao listar Programas Instalados", str(e)))
    else:
        info.append(("Programas Instalados", "Não suportado neste SO"))
    etapa_atual += 1
    if progresso_callback: progresso_callback(etapa_atual, total_etapas)

    # 8) Quantidade de Dispositivos/Portas USB (Windows via WMI)
    if _wmi_available and platform.system() == "Windows":
        try:
            c = wmi.WMI()
            contagem_usb = sum(1 for _ in c.Win32_USBHub())
            info.append(("Quantidade de Dispositivos/Portas USB", str(contagem_usb)))
        except Exception as e:
            info.append(("Erro ao contar USB via WMI", str(e)))
    else:
        info.append(("Contagem USB", "Não suportado neste SO"))
    etapa_atual += 1
    if progresso_callback: progresso_callback(etapa_atual, total_etapas)

    # 9) Endereço IP Local
    try:
        hostname = socket.gethostname()
        ip_local = socket.gethostbyname(hostname)
        info.append(("Endereço IP Local", ip_local))
    except Exception as e:
        info.append(("Erro ao obter IP Local", str(e)))
    etapa_atual += 1
    if progresso_callback: progresso_callback(etapa_atual, total_etapas)

    # 10) Interfaces de Rede (IPv4 e MAC)
    try:
        if_addrs = psutil.net_if_addrs()
        for interface, adicionrs in if_addrs.items():
            for addr in adicionrs:
                if addr.family == socket.AF_INET:
                    info.append((f"Interface {interface} - IPv4", addr.address))
                elif hasattr(socket, "AF_PACKET") and addr.family == socket.AF_PACKET:
                    info.append((f"Interface {interface} - MAC", addr.address))
    except Exception:
        pass
    etapa_atual += 1
    if progresso_callback: progresso_callback(etapa_atual, total_etapas)

    # 11) Número de Série do Equipamento (Windows via WMI)
    serial = "Não disponível neste SO"
    if _wmi_available and platform.system() == "Windows":
        try:
            c = wmi.WMI()
            bios_info = c.Win32_BIOS()
            if bios_info and bios_info[0].SerialNumber:
                serial = bios_info[0].SerialNumber.strip()
        except Exception as e:
            serial = f"Erro ao obter serial: {e}"
    info.append(("Número de Série do Equipamento", serial))
    etapa_atual += 1
    if progresso_callback: progresso_callback(etapa_atual, total_etapas)

    # Remover acentos de todas as chaves e valores:
    info_ascii = []
    for chave, valor in info:
        chave_limpa = remover_acentos(str(chave))
        valor_limpo = remover_acentos(str(valor))
        info_ascii.append((chave_limpa, valor_limpo))

    return info_ascii

def salvar_csv(caminho, dados):
    """
    Recebe o caminho (com extensão .csv) e uma lista de tuplas (chave, valor) e
    salva tudo num arquivo CSV com delimitador ponto e vírgula.
    """
    with open(caminho, mode='w', newline='', encoding='utf-8') as arquivo:
        writer = csv.writer(arquivo, delimiter=';')
        writer.writerow(["Item", "Valor"])
        for chave, valor in dados:
            writer.writerow([chave, valor])

class InventarioApp:
    def __init__(self, root):
        self.root = root
        root.title("Inventário de Equipamentos - Solução TI RJ")
        root.geometry("650x450")
        root.resizable(False, False)

        # Título
        titulo = tk.Label(root, text="Solução TI RJ - Inventário de Equipamentos",
                          font=("Helvetica", 14, "bold"))
        titulo.pack(pady=(10, 0))

        subtitulo = tk.Label(root, text="Uso Reservado - Dúvidas procurar Samuel ou Jeniffer", font=("Helvetica", 10, "italic"), fg="red")
        subtitulo.pack()

        # Barra de progresso
        self.progresso = ttk.Progressbar(root, orient="horizontal", length=600, mode="determinate")
        self.progresso.pack(pady=(10, 5))

        # Área de texto para log
        self.texto = scrolledtext.ScrolledText(root, width=80, height=15, state="disabled")
        self.texto.pack(pady=(5, 10))

        # Botões
        botoes_frame = tk.Frame(root)
        botoes_frame.pack()
        self.btn_iniciar = tk.Button(botoes_frame, text="Iniciar Busca", width=15, command=self.iniciar_coleta)
        self.btn_iniciar.grid(row=0, column=0, padx=5)
        btn_sair = tk.Button(botoes_frame, text="Sair", width=10, command=root.quit)
        btn_sair.grid(row=0, column=1, padx=5)

    def log(self, mensagem: str):
        """
        Adiciona uma linha no log (ScollText), habilitando temporariamente a edição.
        """
        self.texto.configure(state="normal")
        self.texto.insert(tk.END, mensagem + "\n")
        self.texto.see(tk.END)
        self.texto.configure(state="disabled")
        self.root.update_idletasks()

    def progresso_callback(self, etapa_atual, total):
        """
        Atualiza a barra de progresso conforme a etapa.
        """
        valor = int((etapa_atual / total) * 100)
        self.progresso["value"] = valor
        self.log(f"[{etapa_atual}/{total}] Concluída fase: {etapas_map[etapa_atual-1]}")
        self.root.update_idletasks()

    def iniciar_coleta(self):
        """
        Desabilita o botão, limpa o log e inicia a coleta de informações.
        """
        self.btn_iniciar.config(state="disabled")
        self.progresso["value"] = 0
        self.texto.configure(state="normal")
        self.texto.delete("1.0", tk.END)
        self.texto.configure(state="disabled")

        self.log("Iniciando coleta de informações...")

        # Executar a coleta passo a passo
        dados = coletar_informacoes(progresso_callback=self.progresso_callback)

        self.log("Coleta concluída. Preparando para salvar CSV...")
        self.progresso["value"] = 100
        self.root.update_idletasks()

        # Diálogo “Salvar como” para CSV
        caminho = filedialog.asksaveasfilename(
            title="Salvar Inventário como...",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")]
        )
        if caminho:
            try:
                if not caminho.lower().endswith(".csv"):
                    caminho += ".csv"
                salvar_csv(caminho, dados)
                self.log(f"Arquivo CSV salvo em: {caminho}")
                messagebox.showinfo("Concluído", "Inventário gerado com sucesso!")
            except Exception as e:
                self.log(f"Erro ao salvar CSV: {e}")
                messagebox.showerror("Erro ao salvar", f"Não foi possível salvar o arquivo:\n{e}")
        else:
            self.log("Operação de salvamento cancelada pelo usuário.")

        self.btn_iniciar.config(state="normal")

# Mapeamento de nomes das fases (para exibir no log)
etapas_map = [
    "Data e Hora",
    "Sistema e Host",
    "CPU",
    "Memória RAM",
    "Espaço em Disco",
    "Temperaturas",
    "Programas Instalados",
    "Contagem USB",
    "IP Local",
    "Interfaces de Rede",
    "Número de Série"
]

if __name__ == "__main__":
    root = tk.Tk()
    app = InventarioApp(root)
    root.mainloop()


