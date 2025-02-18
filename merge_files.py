import os
import pandas as pd
import re

def remove_special_characters(text):

    text = text.strip()
    text = re.sub(r'[^\u4e00-\u9fa5a-zA-Z]+', '', text)
    return text

def merge_excel_files(input_folder, output_file):
    writer = pd.ExcelWriter(output_file, engine='openpyxl')

    for filename in os.listdir(input_folder):
        if filename.endswith('.xlsx') or filename.endswith('.xls'):
            file_path = os.path.join(input_folder, filename)
            
            df = pd.read_excel(file_path)
            
            sheet_name = remove_special_characters(os.path.splitext(filename)[0])

            df.to_excel(writer, sheet_name=sheet_name, index=False)
    writer.close()


input_folder = r""
output_file = ""
merge_excel_files(input_folder, output_file)
