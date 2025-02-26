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

# pip install rarfile --proxy='http://rb-proxy-de.bosch.com:8080'
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
    
def extract_text_from_compact_file(file_path, poppler_path='poppler-24.08.0\\Library\\bin', dpi=300):
    pdf_img = p2i(file_path, dpi=dpi, poppler_path=poppler_path)

    text_content = []
    
    for img_convert in pdf_img:
        image = cv2.cvtColor(np.array(img_convert), cv2.COLOR_RGB2BGR)
        improved_img = correct_skew(image)
        result_ocr = reader.readtext(improved_img)
        
        page_text = ' '.join([text[1] for text in result_ocr])
        
        if page_text:
            text_content.append(page_text)
    return text_content

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
    if os.path.exists(extract_path):
        for item in os.listdir(extract_path):
            item_path = os.path.join(extract_path, item)
            try:
                if os.path.isdir(item_path):
                    shutil.rmtree(item_path)
                else:
                    os.remove(item_path)
            except Exception as e:
                print(f'Erro ao tentar remover {item_path}: {e}')

def move_files_to_root(extract_path):
    for root, _, files in os.walk(extract_path):
        for file in files:
            src = os.path.join(root, file)
            dest = os.path.join(extract_path, file)
            if src != dest:
                shutil.move(src, dest)

def extract_files(file_temp, ext):
    """ Extrai arquivos compactados para uma pasta temporária e retorna a lista de arquivos extraídos. """
    extract_path = os.path.join(tf.gettempdir(), f"extracted_{ext.strip('.')}")
    os.makedirs(extract_path, exist_ok=True)
    
    try:
        if ext == '.zip':
            with zipfile.ZipFile(file_temp, 'r') as zip_ref:
                zip_ref.extractall(path=extract_path)
                file_paths = zip_ref.namelist()

        elif ext == '.7z':
            with py7zr.SevenZipFile(file_temp, 'r') as sevenz_ref:
                sevenz_ref.extractall(path=extract_path)
                file_paths = sevenz_ref.getnames()

        elif ext == '.rar':
            with rarfile.RarFile(file_temp, 'r') as rar_ref:
                rar_ref.extractall(path=extract_path)
                file_paths = rar_ref.namelist()

        # TESTAR - VERIFICAR SE ESTA FUNCIONAL
        elif ext in ('.tar', '.gz'):
            with tarfile.open(file_temp, 'r:*') as tar_ref:
                tar_ref.extractall(path=extract_path)
                file_paths = tar_ref.getnames()
                #  gzip.GzipFile

        move_files_to_root(extract_path)
        return extract_path, file_paths

    except Exception as e:
        log.error(f"Erro ao extrair {ext}: {e}")
        return None, []

def process_pdfs_compressed(extract_path):
    """ Processa arquivos PDF extraídos e extrai o texto. """
    text_content = []
    
    for file in os.listdir(extract_path):
        if file.lower().endswith('.pdf'):
            file_path = os.path.join(extract_path, file)

            try:
                with ppl.open(file_path) as pdf:
                    for page in pdf.pages:
                        text = page.extract_text()
                        if text and text.strip():
                            text_content.append(text)

                if not text_content:
                    text_content = extract_text_from_compact_file(file_path)

                log.info(f'Texto extraído do PDF {file}: {text_content[:100]}...')

            except Exception as e:
                log.error(f"Erro ao processar PDF {file}: {e}")

    return text_content







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
        for item in tqdm(inbox.Items, desc='Processando e-mails', unit=' e-mail'):
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

                            try:
                                if attachment.FileName.lower().endswith(('.pdf')):
                                    file_temp = save_temp(attachment)
                                    if file_temp:
                                        with ppl.open(file_temp) as pdf:
                                            for page in pdf.pages:
                                                text = page.extract_text()
                                                if text and text.strip():
                                                    
                                                    
                                                    '''

                                                    LOGICA PARA IDENTIFICAÇÃO SE O CONTEUDO É CERTIFICADO
                                                    
                                                    '''
                                                    #save_attachment(file_temp)

                                                    log.info(f'Texto extraído do PDF {attachment.FileName}: {text[:100]}...')

                                                else:
                                                    pdf_img = p2i(file_temp, dpi=300, poppler_path='poppler-24.08.0\\Library\\bin')               # Uma lista de imagens, onde cada imagem representa uma página do PDF.
                                                    for i, img_convert in enumerate(pdf_img):                                   # Itera sobre cada imagem e á aloca uma por vez em img_convert
                                                        image = cv2.cvtColor(np.array(img_convert), cv2.COLOR_RGB2BGR)          # Converter imagem PIL para OpenCV               
                                                        improved_img = correct_skew(image)                                   # Corrigir inclinação da imagem
                                                        result_ocr = reader.readtext(improved_img)                           # Processar a imagem com OCR
                                                    
                                                        '''
                                                        LOGICA PARA IDENTIFICAÇÃO SE O CONTEUDO É CERTIFICADO
                                                        '''

                                                    log.info(f'Texto extraído do PDF/IMG {attachment.FileName}: {result_ocr[:100]}...')
                                                                                                                            
                                    os.remove(file_temp)
                                    status_checkmark(item, MarckCheck)
                           
                            except Exception as e:
                                log.error(f'Erro ao processar arquivo PDF {attachment.FileName}: {e}')  
                            
                            try:    
                                if attachment.FileName.lower().endswith(('.jpg', '.png')):
                                    file_temp = save_temp(attachment)
                                    if file_temp:
                                        result_ocr = reader.readtext(file_temp)
                                        if result_ocr:
                                            ...
                                    
                                        '''
                                        LOGICA PARA IDENTIFICAÇÃO SE O CONTEUDO É CERTIFICADO                      
                                        '''
                                        # save_attachment()

                                        log.info(f'Texto extraído do JPG/PNG {attachment.FileName}: {result_ocr[:100]}...')
                                        
                                        
                                    os.remove(file_temp)
                                    status_checkmark(item, MarckCheck)        
                               
                            except Exception as e:
                                log.error(f'Erro ao processar arquivo de imagem {attachment.FileName}: {e}')

                            try:       
                                if attachment.FileName.lower().endswith(('.zip', '.7z', '.rar', '.tar', '.gz')):
                                    file_temp = save_temp(attachment)
                                    if file_temp:
                                        ext = os.path.splitext(file_temp)[1].lower()
                                        extract_path, file_paths = extract_files(file_temp, ext)

                                        if extract_path:
                                            text_content = process_pdfs_compressed(extract_path)
                                            
                                            '''
                                            LOGICA PARA IDENTIFICAÇÃO SE O CONTEUDO É CERTIFICADO   

                                            USAR text_content (ONDE O TEXTO É RETORNADO)                   
                                            '''
                                            # save_attachment()

                                            clear_path(extract_path)
                                    
                                    os.remove(file_temp)        
                                    status_checkmark(item, MarckCheck)       
                                                 
                            except Exception as e:
                                log.error(f'Erro ao processar o arquivo compactado {attachment.FileName}: {e}')            
                
                except Exception as e:
                    log.error(f'Erro ao extrair informações do email {sender}: {e}') 

                    item.MarkAsTask(MarckRed)         
                    item.FlagStatus = MarckRed
                    item.Save()
                  
    except Exception as e:
        log.error(f'Erro ao iterar e-mails: {e}')
    
except Exception as e:
    log.error(f'Erro de processamento: {e}')
