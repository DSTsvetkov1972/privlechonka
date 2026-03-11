import os
import sys 
sys.path.append(os.getcwd())

import pandas as pd
import zipfile
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from datetime import datetime
from processors.config import files_dir, sent_dir, decisions_dir, prepared_file_path, summary_sent_path
from processors.fns import get_last_sent_number, create_doc_file
from colorama import Fore


def sent_maker():
    if not os.path.exists(prepared_file_path):
        print(Fore.MAGENTA, 'Нет файлов подготовленных для отправки на актирование', Fore.RESET)
        return

    prepared_df = pd.read_excel(prepared_file_path)
    if prepared_df.empty:
        print(Fore.MAGENTA, 'Файл подготовленный для актирования пуст', Fore.RESET)
        return

    while os.path.exists(os.path.join(files_dir, '~$сводка по отправленным на актирование.xlsx')):
        os.startfile(summary_sent_path)
        print(Fore.RED,'Закройте файл "сводка по отправленным на актирование.xlsx" и нажмите ввод', Fore.RESET, end='')
        input()

    while os.path.exists(os.path.join(files_dir, '~$подготовлено к передаче на актирование.xlsx')):
        os.startfile(prepared_file_path)
        print(Fore.RED, 'Закройте файл "подготовлено к передаче на актирование.xlsx" и нажмите ввод', Fore.RESET, end='')        
        input()    

    # получаем номер предыдущей отправки на актирование
    last_sent_number = get_last_sent_number()

    sent_file_short_name = f'{ os.getcwd().split('\\')[-1] }_{ int(last_sent_number)+1 }_{ str(datetime.now())[:19].replace(':', '-') }'
    
    sent_file_xlsx = sent_file_short_name + '.xlsx'
    sent_file_xlsx_path = os.path.join(sent_dir, sent_file_xlsx)
    
    sent_file_doc = sent_file_short_name + '.doc'
    sent_file_doc_path = os.path.join(sent_dir, sent_file_doc)    
    
    sent_file_zip = sent_file_short_name + '.zip'
    sent_file_zip_path = os.path.join(sent_dir, sent_file_zip)

    
    prepared_df['файл'] = sent_file_xlsx


    if os.path.exists(summary_sent_path):
        summary_prepared_df = pd.read_excel(summary_sent_path)
        summary_prepared_df = pd.concat([summary_prepared_df, prepared_df])
    else:
        summary_prepared_df = prepared_df
    

    os.rename(prepared_file_path, sent_file_xlsx_path)

    create_doc_file(prepared_df,sent_file_doc_path)

    with zipfile.ZipFile(sent_file_zip_path, 'w') as zipf:
        desicion_file_names = prepared_df['Файл с решением ЭС'].drop_duplicates()

        for file in desicion_file_names:
            zipf.write(os.path.join(decisions_dir, file), arcname=file)

    # ceatre_zip_file(summary_prepared_df,sent_file_doc_path)
    summary_prepared_df.to_excel(summary_sent_path, index=False)


    print(Fore.GREEN, f'Передан на актирование { sent_file_short_name } (.xlsx, .doc, .zip)', Fore.RESET)
        



    

if __name__ == '__main__':
    sent_maker()    
