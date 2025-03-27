import os
import json
import logging as log
import win32com.client as client
from tqdm import tqdm

import tempfile as tf
from pdf2image import convert_from_path as p2i

with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assistant', 'configPDFC.json'), 'r') as config_file:
    config = json.load(config_file)
email_field = config.get('email_field')

outlook = client.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace('MAPI')
try:
    inbox = namespace.Folders[email_field].Folders['TESTE PDFC']      # ALTERAR  - inbox = namespace.GetDefaultFolder(6)
except Exception as e:                                                # REMOVER #
    log.error(f'Erro ao acessar caixa de entrada: {e}')               # ------- #

try:
    for item in tqdm(inbox.Items, desc='Processando e-mails', unit=' e-mail'):
        if item.Class == 43 and item.FlagStatus == 0:
            try:
                sender = item.SenderEmailAddress
                sender = 'pedro@bosch.com'
                domain = sender.split('@')[1].split('.')[0] 
                receipt_date = item.ReceivedTime
                receipt_year = receipt_date.year

                '''
                pasta_destino = None
                nome_pasta_destino = 'precessed'
                for pasta in inbox.Folders:
                     if pasta.Name.lower() == nome_pasta_destino.lower():
                        pasta_destino = pasta
                        break
                '''     
                pasta_destino = namespace.Folders[email_field].Folders['Processed'] 

                if pasta_destino:
                    item.Move(pasta_destino)

            except Exception as e:
                            log.error(f'Erro ao processar: {e}')  
except Exception as e:
    log.error(f'Erro ao processar: {e}')  