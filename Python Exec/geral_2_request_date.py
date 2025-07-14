import tkinter as tk
from tkinter import messagebox
from datetime import timedelta, datetime as dt
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import os
import time
import logging
import socket
import getpass

# Nome da máquina que usa o caminho local
HOST_LOCAL = "JAGUAR-ANTIGO"

# Coleta informações do sistema
hostname = socket.gethostname()
ip_address = socket.gethostbyname(hostname)
usuario_logado = getpass.getuser()


# Configuração de logging
os.makedirs("logs", exist_ok=True)  # Cria diretório de logs se não existir
log_path = f'logs/automacao_rel_geral_{dt.now().strftime("%Y-%m-%d_%H-%M-%S")}.log'
logging.basicConfig(
    filename=log_path,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logging.getLogger().addHandler(logging.StreamHandler())  # Mostra log também no terminal

# Primeira linha com informações do sistema
logging.info(f"Início da automação - IP: {ip_address}, Computador: {hostname}, Usuário: {usuario_logado}")

# Variáveis globais para armazenar as datas
data_inicial_dt = None
data_final_dt = None
data_meio_dt = None

#para chamar as funções
from function_limpar_diretorios import limpar_pasta
from function_con_sieg_completo import realizar_pesquisas_e_download
#from function_salvar_dominio_completo import gerar_relatorios_conferencia
#from function_excel_macro import abrir_excel_e_maximizar
#from function_enviar_email import enviar_email_relatorio
#from function_mover_arquivos_completo import organizar_arquivos

# Caminhos fixos usados na automação
if hostname == HOST_LOCAL:
    caminho_historico = r"F:\Arquivos Digitais\18 - T.I\Relatório Geral de Notas"
else:
    caminho_historico = r"\\192.168.1.200\arquivosdigitais$\18 - T.I\Relatório Geral de Notas"
caminho_exe_dominio = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Domínio Contábil\Domínio Escrita Fiscal"
senha_dominio = "16129252"
caminho_r1 = r'F:\Arquivos Digitais\18 - T.I\Relatório Geral de Notas\Resumido - Relatório de Cupons Fiscais.xls'
caminho_r2 = r'F:\Arquivos Digitais\18 - T.I\Relatório Geral de Notas\Resumido - Relatório de Entradas.xls'
caminho_r3 = r'F:\Arquivos Digitais\18 - T.I\Relatório Geral de Notas\Resumido - Relatório de Saídas.xls'
caminho_r4 = r'F:\Arquivos Digitais\18 - T.I\Relatório Geral de Notas\Empresas.xls'
caminho_macro = r"C:\Users\aloyr\Documents\Macros\Conferência de Saídas.xlsm"

def calcular_datas():
    global data_inicial_dt, data_final_dt, data_meio_dt

    data_inicial_str = entrada_inicial.get()
    data_final_str = entrada_final.get()
    logging.info(f"Entrada do usuário - Data Inicial: {data_inicial_str}, Data Final: {data_final_str}")

    try:
        data_inicial_dt = dt.strptime(data_inicial_str, '%d/%m/%Y')
        data_final_dt = dt.strptime(data_final_str, '%d/%m/%Y')

        if data_inicial_dt > data_final_dt:
            messagebox.showinfo("Aviso", "Data inicial é posterior à data final. Invertendo as datas.")
            logging.warning("Data inicial maior que a final. Realizando inversão.")
            data_inicial_dt, data_final_dt = data_final_dt, data_inicial_dt

        dif_dias = (data_final_dt - data_inicial_dt).days
        data_meio_dt = data_inicial_dt + timedelta(days=dif_dias // 2)

        logging.info(f"Datas calculadas: Inicial = {data_inicial_dt}, Final = {data_final_dt}, Meio = {data_meio_dt}")
        janela.destroy()  # Fecha a interface após capturar as datas
    except ValueError:
        logging.error("Formato de data inválido.")
        messagebox.showerror("Erro", "Por favor, digite as datas no formato dd/mm/aaaa.")

# Criando a janela principal
janela = tk.Tk()
janela.title("Período do Relatório")
janela.geometry("300x200")

# Datas padrão (início do mês até hoje)
data_hoje = dt.today()
data_inicio_mes = data_hoje.replace(day=1)

# Labels e entradas
tk.Label(janela, text="Data Inicial (dd/mm/aaaa):").pack()
entrada_inicial = tk.Entry(janela)
entrada_inicial.insert(0, data_inicio_mes.strftime('%d/%m/%Y'))
entrada_inicial.pack()

tk.Label(janela, text="Data Final (dd/mm/aaaa):").pack()
entrada_final = tk.Entry(janela)
entrada_final.insert(0, data_hoje.strftime('%d/%m/%Y'))
entrada_final.pack()

# Botão para calcular
tk.Button(janela, text="Gerar Relatório", command=calcular_datas).pack(pady=10)

# Iniciar a interface
logging.info("Interface de entrada iniciada.")
tk.mainloop()

# Continuação da automação após a janela ser fechada
if data_inicial_dt and data_final_dt and data_meio_dt:
    data_inicial = data_inicial_dt.strftime('%d/%m/%Y')
    data_final = data_final_dt.strftime('%d/%m/%Y')
    data_meio = data_meio_dt.strftime('%d/%m/%Y')

    logging.info("Iniciando automação com as datas capturadas.")
    logging.info(f"Data Inicial: {data_inicial}, Data Final: {data_final}, Data Meio: {data_meio}")

    # Executa as funções da automação
    try: 
        logging.info("Iniciando realizar_pesquisas_e_download")
        realizar_pesquisas_e_download(data_inicial, data_final, caminho_historico)

        # logging.info("Iniciando gerar_relatorios_conferencia")
        # gerar_relatorios_conferencia(data_inicial, data_final, senha_dominio, caminho_exe_dominio, caminho_historico, caminho_r1, caminho_r2, caminho_r3, caminho_r4)

        # logging.info("Iniciando organizar_arquivos")
        # caminho_pasta_organizada = organizar_arquivos()

        logging.info("Automação concluída com sucesso.")
    except Exception as e:
        logging.exception("Erro durante a execução da automação:")
