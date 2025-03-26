import os
import shutil

extract_path = fr'C:\Users\KRP1PO\AppData\Local\Temp\extracted_zip'

folder_files = os.listdir(extract_path)

if folder_files:
    for i in folder_files:
        file_remove = os.path.join(extract_path, i)
        if os.path.isfile(file_remove):
           os.remove(file_remove)
        elif os.path.isdir(file_remove):
            shutil.rmtree(file_remove)           

    