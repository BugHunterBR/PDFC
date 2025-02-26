import os
import json
import logging as log
import win32com.client as win
from tqdm import tqdm
import pdfplumber as ppl
import easyocr as ocr
import tempfile as tf
from pdf2image import convert_from_path as p2i
import cv2
import numpy as np                                              # USADA EM correct_skew
import zipfile
import py7zr
import rarfile
import tarfile
import shutil

# pip install rarfile --proxy="http://rb-proxy-de.bosch.com:8080"
# http://rb-proxy-ca1.bosch.com:8080

log.basicConfig(filename=os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assistant', 'Processamento_PDFC.log'), level=log.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assistant', 'configPDFC.json'), 'r') as config_file:
    config = json.load(config_file)
    
email_field = config.get('email_field')
base_folder_path = config.get('base_folder_path')
MarckCheck = int(config.get('MarckCheck', 0))
MarckRed = int(config.get('MarckRed', 0))

reader = ocr.Reader(['pt','en'], model_storage_directory=os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ocr'))

def extract_domain(email):  
    try:
        part_domain = email.split('@')[1]
        main_domain = part_domain.split('.')[0]        
        return main_domain.upper() 
    except (IndexError, AttributeError):
        return None

def create_folder_year(base_folder_path, year):
    base_folder_year = os.path.join(base_folder_path, f'Arquivo {year}')
    if not os.path.exists(base_folder_year):
        os.makedirs(base_folder_year)
        log.info(f'Pasta criada: {base_folder_year}')
        return base_folder_year
    return base_folder_year

def save_temp(attachment):
    file_extension = os.path.splitext(attachment.FileName)[1] 
    try:
        with tf.NamedTemporaryFile(delete=False, suffix=file_extension) as file_temp:
            temp_path = file_temp.name
            attachment.SaveAsFile(temp_path)
        return temp_path
    except Exception as e:
        log.error(f'Erro ao criar arquivo temporário {attachment.FileName}: {e}')
        return None
    
def save_attachment(attachment, destination_folder_year, domain):
    try:
        destination_folder_domain = os.path.join(destination_folder_year, domain)
        destination_attachment = os.path.join(destination_folder_domain, attachment.FileName)
        if not os.path.exists(destination_folder_domain):
            os.makedirs(destination_folder_domain)
            log.info(f'Pasta criada: {destination_folder_domain}')
        if not os.path.exists(destination_attachment):
            attachment.SaveAsFile(destination_attachment)
            log.info(f'Anexo salvo em: {destination_attachment}') 
        return destination_attachment
    except Exception as e:
        log.error(f'Erro ao salvar o anexo {attachment.FileName}: {e}')
        return None
    
def status_checkmark(item, status):
    try:
        item.MarkAsTask(status)
        item.FlagStatus = status
        item.Save()    
    except Exception as e:
        log.error(f'Erro ao marcar email {attachment.FileName}: {e}')
        return None

def correct_skew(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    coords = np.column_stack(np.where(thresh > 0))
    angle = cv2.minAreaRect(coords)[-1]
    if angle < -45:
        angle = -(90 + angle)
    else:
        angle = -angle
    (h, w) = image.shape[:2]
    center = (w // 2, h // 2)
    M = cv2.getRotationMatrix2D(center, angle, 1.0)
    rotated = cv2.warpAffine(image, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
    return rotated

def clear_path(extract_path):
    # Verifica se a pasta existe
    if os.path.exists(extract_path):
        # Itera sobre os arquivos e subdiretórios dentro da pasta
        for item in os.listdir(extract_path):
            item_path = os.path.join(extract_path, item)
            try:
                # Se for um diretório, remove recursivamente
                if os.path.isdir(item_path):
                    shutil.rmtree(item_path)
                else:
                    # Se for um arquivo, remove o arquivo
                    os.remove(item_path)
            except Exception as e:
                print(f"Erro ao tentar remover {item_path}: {e}")

def move_files_to_root(extract_path):
    for root, _, files in os.walk(extract_path):
        for file in files:
            src = os.path.join(root, file)
            dest = os.path.join(extract_path, file)
            if src != dest:
                shutil.move(src, dest)

   
try:
    outlook = win.Dispatch('Outlook.Application').GetNamespace('MAPI')
    try:
        inbox = outlook.Folders[email_field].Folders['TESTE PDFC']      # ATENÇÃO - ALTERAR
    except Exception as e:
        try:
            inbox = outlook.Folders[email_field].Folders['Inbox']
        except Exception as e:
            log.error(f'Erro ao acessar caixa de entrada: {e}')

    try:
        for item in tqdm(inbox.Items, desc="Processando e-mails", unit=" e-mail"):
            if item.Class == 43 and item.FlagStatus == 0: 
                try:
                    sender = item.SenderEmailAddress
                    receipt_date = item.ReceivedTime
                    receipt_year = receipt_date.year
                    sender = 'pedro@bosch.com'
                    domain = extract_domain(sender)

                    if domain is not None:
                        destination_folder_year = create_folder_year(base_folder_path, receipt_year)
                        for attachment in item.Attachments:
################################################################################################################################################

                            try:
                                if attachment.FileName.lower().endswith(('.pdf')):
                                    file_temp = save_temp(attachment)
                                    if file_temp:                                                               # SE O ARQUIVO EXISTIR
                                        with ppl.open(file_temp) as pdf:                                        # ABRE O ARQUIVO COM PPL E O IDENTIFICA PO PDF
                                            for page in pdf.pages:                                              # ITERA CADA PAGINA DO PDF
                                                text = page.extract_text()                                      # EXTRAI O TEXTO DE CADA PAGINA
                                                if text and text.strip():                                       # VERIFICA SE TEXT NÃO É NONE E USA STRIP PARA REMOVER ESPAÇOS (CASO O PDF ESTEJA VAZIO E POSSUA APENAS ESPAÇOS)
                                                    log.info(f"Texto extraído do PDF {text}: {text_content[:100]}...")
                                                    ...
                                                    
                                                    '''
                                                    LOGICA PARA IDENTIFICAÇÃO SE O CONTEUDO É CERTIFICADO
                                                    '''

                                                    #save_attachment(file_temp)

                                                # REALIZAR PROCESSO DE VERIFICAÇÃO PARA VER REALMENTE É UMA IMAGEM
                                                else:                                                                           # SE O PDF ESTIVER VAZIO (PODE SER UMA POSSIVEL IMAGEM)
                                                    pdf_img = p2i(file_temp, dpi=300, poppler_path='poppler-24.08.0\\Library\\bin')               # Uma lista de imagens, onde cada imagem representa uma página do PDF.
                                                    for i, img_convert in enumerate(pdf_img):                                   # Itera sobre cada imagem e á aloca uma por vez em img_convert
                                                        image = cv2.cvtColor(np.array(img_convert), cv2.COLOR_RGB2BGR)          # Converter imagem PIL para OpenCV               
                                                        improved_img = correct_skew(image)                                   # Corrigir inclinação da imagem
                                                        result_ocr = reader.readtext(improved_img)                           # Processar a imagem com OCR
                                                        #print(f"Resultado OCR página {i+1}: {result_ocr}")

                                                        '''
                                                        LOGICA PARA IDENTIFICAÇÃO SE O CONTEUDO É CERTIFICADO
                                                        '''

                                            # VERIFICAR ONDE DEVE SER POSTO (SALVAMENTO DE IDNEIFCAÇÃO DE PDF -> IMG -> OCR)           
                                            
                                            if result_ocr:                                                                      # Se a ocr não estiver vazia
                                                save_attachment(file_temp)                                                      # Chama a função que salva o arquivo
                                                                                                                            
                                    os.remove(file_temp)
                                    status_checkmark(item, MarckCheck)
                           
                            except Exception as a:
                                log.error(f'Erro ao processar arquivo PDF {attachment.FileName}: {a}')  
                            
################################################################################################################################################

                            try:    
                                if attachment.FileName.lower().endswith(('.jpg', '.png')):
                                    file_temp = save_temp(attachment)
                                    if file_temp:                                               # Se o arquivo temporario existir
                                        result_ocr = reader.readtext(file_temp)                 # Faz a leitura usando OCR
                                        if result_ocr:                                          # Se a leitura da OCR não retornar vazio
                                            ...
                                    
                                        '''
                                        LOGICA PARA IDENTIFICAÇÃO SE O CONTEUDO É CERTIFICADO                      
                                        '''
                                        # save_attachment()
                                        
                                        os.remove(file_temp)
                                    status_checkmark(item, MarckCheck)        
                               
                            except Exception as a:
                                log.error(f'Erro ao processar arquivo de imagem {attachment.FileName}: {a}')
                            
################################################################################################################################################

                            try:       
                                if attachment.FileName.lower().endswith(('.zip', '.7z', '.rar', '.tar', '.gz')):
                                    file_temp = save_temp(attachment)
                                    if file_temp:
                                        ext = os.path.splitext(file_temp)[1].lower()
                                        file_paths = []
                                        extract_path = tf.gettempdir() # Diretório temporário do sistema
                                        #---------------------------------------------------------------------------------#
                                        if ext == '.zip':
                                            extract_path = os.path.join(extract_path, "extracted_zip")           # cria o caminho onde os arquivos serão extraidos e adiciona a pasta extracted_7z para ser criada
                                            os.makedirs(extract_path, exist_ok=True)
                                            try:
                                                with zipfile.ZipFile(file_temp, 'r') as zip_ref:
                                                    zip_ref.extractall(path=extract_path)
                                                    file_paths = zip_ref.namelist()

                                                    move_files_to_root(extract_path)

                                            except Exception as e:
                                                log.info(f"Erro ao extrair .zip: {e}")
                                            
                                            # VERIFICANDO SE OS ARQUIVOS EXTRAIDOS SÃO PDF
                                            # POSSIBILIDADE DE CRIAR FUNÇÃO - USADA EM ZIP E 7Z
                                            for file in os.listdir(extract_path):
                                                if file.lower().endswith(".pdf"):
                                                    file_path = os.path.join(extract_path, file)

                                                    if file_path:
                                                        with ppl.open(file_path) as pdf:
                                                            text_content = []
                                                            for page in pdf.pages:
                                                                text = page.extract_text()
                                                                if text and text.strip():
                                                                    text_content.append(text)
                                                            
                                                            if text_content:
                                                                log.info(f"Texto extraído do PDF {file}: {text_content[:100]}...")
                                                                '''

                                                                LOGICA PARA IDENTIFICAÇÃO SE O CONTEUDO É CERTIFICADO
                                                                
                                                                '''
                                                            else:
                                                                pdf_img = p2i(file_path, dpi=300, poppler_path='poppler-24.08.0\\Library\\bin')
                                                                for i, img_convert in enumerate(pdf_img):
                                                                    image = cv2.cvtColor(np.array(img_convert), cv2.COLOR_RGB2BGR)
                                                                    improved_img = correct_skew(image)
                                                                    result_ocr = reader.readtext(improved_img)
                                                                    print(f"Resultado OCR página {i+1}: {result_ocr}")
                                                                    '''
                                                                    LOGICA PARA IDENTIFICAÇÃO SE O CONTEUDO É CERTIFICADO
                                                                    '''

                                                                    '''
                                                                    LOGICA PARA APLICAR O SALVAMENTO NA PASTA CORRETA
                                                                    NESTA PARTE DO CAMINHO O ARQUIVO ESTA SENDO SALVO EM extracted_7z
                                                                    '''

                                                                    '''
                                                                    APAGAR O CONTEUDO DA PASTA extracted_7z PARA QUE POSSA RECEBER NOVOS CONTEUDOS POSTERIORMENTE
                                                                    '''
                                            clear_path(extract_path)

                                            if len(file_paths) > 0 or len(text_content) > 0: 
                                                file_paths.clear()
                                                text_content.clear()

                                            
                                        #---------------------------------------------------------------------------------#
                                        elif ext == '.7z':                                                      # verifica se a extenção é .7z
                                            extract_path = os.path.join(extract_path, "extracted_7z")           # cria o caminho onde os arquivos serão extraidos e adiciona a pasta extracted_7z para ser criada
                                            os.makedirs(extract_path, exist_ok=True)                            # cria a pasta de acordo com o caminho de extract_path
                                            try:
                                                with py7zr.SevenZipFile(file_temp, "r") as sevenz_ref:
                                                    '''
                                                    Abre o arquivo do caminho file_temp em somente leitura e o referencia como sevenz_ref
                                                    '''
                                                # with                                                          - garante que o arquivo seja automaticamente fechado após o uso    
                                                # sevenz_ref                                                    - passa a ser o objeto de referência do arquivo .7z   
                                                # py7zr.SevenZipFile(...)                                       - é a classe usada para abrir um arquivo .7z.
                                                # file_temp                                                     - é o caminho do arquivo .7z salvo temporariamente.
                                                # "r"                                                           - significa modo de leitura (read mode).
                                                    sevenz_ref.extractall(path=extract_path)
                                                    '''
                                                    Extrai o arquivos na pasta extracted_7z
                                                    '''
                                                    # sevenz_ref                                                - o objeto de referência do arquivo .7z (o proprio arquivo compacto referenciado)  
                                                    # extractall()                                              - método da biblioteca py7zr que extrai todo o conteúdo do arquivo .7z
                                                    # path=extract_path                                         - define o caminho onde os arquivos serão armazenados (na pasta extracted_7z)
                                                    file_paths = sevenz_ref.getnames()
                                                    '''
                                                    Retorna uma lista dos arquivos que estão dentro do .7z (sevenz_ref)
                                                    '''
                                                    # .getnames()                                               - retorna uma lista dos arquivos que estão dentro de sevenz_ref (arquivo .7z), o resultado da lista é armazenado em file_paths
                                                
                                                # Movendo arquivos para a raiz
                                                move_files_to_root(extract_path)

                                            except Exception as e:
                                                print(f"Erro ao extrair .7z: {e}")

                                            # Processamento de PDFs extraídos
                                            for file in os.listdir(extract_path):
                                                if file.lower().endswith(".pdf"):                       # identifica se é pdf
                                                    file_path = os.path.join(extract_path, file)

                                                    if file_path:
                                                        with ppl.open(file_path) as pdf:
                                                            text_content = []
                                                            for page in pdf.pages:
                                                                text = page.extract_text()
                                                                if text and text.strip():
                                                                    text_content.append(text)

                                                            # Se houver texto, realiza identificação do certificado
                                                            if text_content:
                                                                log.info(f"Texto extraído do PDF {file}: {text_content[:100]}...")
                                                                '''

                                                                LOGICA PARA IDENTIFICAÇÃO SE O CONTEUDO É CERTIFICADO
                                                                
                                                                '''
                                                            else:
                                                                pdf_img = p2i(file_path, dpi=300, poppler_path='poppler-24.08.0\\Library\\bin')
                                                                for i, img_convert in enumerate(pdf_img):
                                                                    image = cv2.cvtColor(np.array(img_convert), cv2.COLOR_RGB2BGR)
                                                                    improved_img = correct_skew(image)
                                                                    result_ocr = reader.readtext(improved_img)
                                                                    print(f"Resultado OCR página {i+1}: {result_ocr}")
                                                                    '''
                                                                    LOGICA PARA IDENTIFICAÇÃO SE O CONTEUDO É CERTIFICADO
                                                                    '''

                                                                    '''
                                                                    LOGICA PARA APLICAR O SALVAMENTO NA PASTA CORRETA
                                                                    NESTA PARTE DO CAMINHO O ARQUIVO ESTA SENDO SALVO EM extracted_7z
                                                                    '''

                                                                    '''
                                                                    APAGAR O CONTEUDO DA PASTA extracted_7z PARA QUE POSSA RECEBER NOVOS CONTEUDOS POSTERIORMENTE
                                                                    '''

                                                                    
                                            clear_path(extract_path)

                                            if len(file_paths) > 0 or len(text_content) > 0: # limpa a lista para receber novos arquivos
                                                file_paths.clear()
                                                text_content.clear()      
                                        #---------------------------------------------------------------------------------#            
                                        
                                        elif ext == '.rar':
                                            with rarfile.RarFile(file_temp, 'r') as rar_ref:
                                                file_paths = rar_ref.namelist()
                                        
                                        #---------------------------------------------------------------------------------#
                                        
                                        elif ext in ['.tar', '.gz']:
                                            with tarfile.open(file_temp, 'r:*') as tar_ref:
                                                file_paths = [f.name for f in tar_ref.getmembers() if f.isfile()]
                                    
                                        os.remove(file_temp)
                                    
                                    status_checkmark(item, MarckCheck)       
                                                 
                            #VERIFICAR EXCEPT
                            except Exception as a:
                                log.error(f'Erro ao processar o arquivo compactado {attachment.FileName}: {a}')            
                
################################################################################################################################################

                except Exception as b:
                    log.error(f'Erro ao extrair informações do email {sender}: {b}') 

                    item.MarkAsTask(MarckRed)         
                    item.FlagStatus = MarckRed
                    item.Save()
                  
    except Exception as c:
        log.error(f'Erro ao iterar e-mails: {c}')
    
except Exception as d:
    log.error(f'Erro de processamento: {d}')
    print(f'Erro de processamento: {d}')

                                        