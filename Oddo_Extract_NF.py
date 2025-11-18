import sys
import pandas as pd
import win32com.client
import subprocess
#import time
import pyautogui
import os
import PyPDF2
import time
import psutil
from time import sleep
from Sap_NF_Extract_Oddo import SapGui
import send2trash


def Extract_NF(user_id, tabela):
    directory2 = "" # pasta onde estao as NFs antigas
    try:
        # Percorre os arquivos na pasta
        for filename in os.listdir(directory2):
            # Verifica se o arquivo é uma planilha Excel
            send2trash.send2trash(f'{directory2}\\{filename}')
            # Exibe a lista de arquivos
        print("NFs antigas excluídas e preparação para extração de novas NFs...")
    except FileNotFoundError:
        print(f"A pasta '{directory2}' não existe.")
 
    
    
    print("Iniciando extração de NFs")
    
    tabela.to_excel(f"Pasta onde salvaremos os dados", index=False)
    
    print("Extração concluída e planilha gerada - Information_Oddo.xlsx")
    print("Agora iremos gerar a PDF da NF. Se não for necessario pode encerrar o script")
    print("Nota fiscal sera gerada em 10 segundos")
    
    sleep(10)
    
    print("Aguardando disponibilização das NFs no SAP: 15 minutos")
    sleep(240)
    print("Aguardando disponibilização das NFs no SAP: 10 minutos")
    sleep(300)
    print("Aguardando disponibilização das NFs no SAP: 5 minutos")
    sleep(240)
    print("Aguardando disponibilização das NFs no SAP: 1 minuto")
    sleep(60)


    tabela['Nota Fiscal'] = SapGui().Execute_SAP(tabela, user_id)

    sleep(2)

    sys.exit(0)
