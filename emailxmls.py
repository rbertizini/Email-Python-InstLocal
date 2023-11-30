import os
import win32com.client
from datetime import datetime, timedelta

# Logs
print("Início")

# Diretório XMLs
diretorio = r"\\192.168.197.22\d$\WebApp\LKM ARDFe OCR Async\PRD\Arquivos\02 disponibilizado\GE"

# Diretório local
arquivo_ultima_data = "ultima_data.txt"

# Data controle
data_ontem = datetime.now() - timedelta(days=1)

# Lê a última data processada (se existir)
try:
    with open(arquivo_ultima_data, "r") as file:
        ultima_data_str = file.read()
    data_ontem = datetime.strptime(ultima_data_str, "%Y-%m-%d")    
except FileNotFoundError:
    data_ontem = datetime.now() - timedelta(days=1)  

# Formatação de data de controle
data_formatada = data_ontem.strftime("%d/%m/%Y")

# Logs
print(f"Data inicial: {data_formatada}")

# Lista de arquivos do dia de ontem
arquivos_ontem = [
    os.path.join(diretorio, arquivo) 
    for arquivo in os.listdir(diretorio) 
    if os.path.isfile(os.path.join(diretorio, arquivo)) and os.path.getctime(os.path.join(diretorio, arquivo)) > data_ontem.timestamp()
]
arquivos_qtd = len(arquivos_ontem)

# Logs
print(f"Arquivos: {arquivos_qtd}")

# Configuração do Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
mensagem = outlook.CreateItem(0)  # 0 significa mensagem de e-mail

# Configurações do e-mail
mensagem.Subject = f"XMLs GE - {data_formatada}"
mensagem.Body = f"Bom dia\n\nSegue anexo arquivos XMLs obtidos a partir do dia {data_formatada}.\nForam anexados {arquivos_qtd} arquivos\n\nObrigado\nRenato Martins"
mensagem.To = "aieza_martinez@lkm.com.br"
#mensagem.To = "renato_martins@lkm.com.br"

# Anexa os arquivos ao e-mail
for arquivo in arquivos_ontem:
    anexo = os.path.abspath(arquivo)
    mensagem.Attachments.Add(anexo)

# Envia o e-mail
mensagem.Send()

# Logs
print("E-mail enviado")

# Atualizando data de processamento
data_ontem = datetime.now() - timedelta(days=1)  

# Formatação de data de controle
data_formatada = data_ontem.strftime("%d/%m/%Y")

# Armazenando data de processamento d-1
try:
    with open(arquivo_ultima_data, "w") as file:
        file.write(data_ontem.strftime("%Y-%m-%d"))    
except FileNotFoundError:
    print("Erro ao gerar o arquivo")

# Logs
print(f"Arquivo atualizado: {data_formatada}")

# Logs
print("Fim")