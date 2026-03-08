import os
import sys 
sys.path.append(os.getcwd())

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from datetime import datetime
from processors.config import source_file_path, prepared_file_path, summary_prepared_path, summary_sent_path
from processors.fns import get_last_sent_number
from colorama import Fore


def sent_maker():


    while os.path.exists(os.path.join('files', '~$сводка по отправленным на актирование.xlsx')):
        os.startfile(summary_sent_path)
        print(Fore.RED,'Закройте файл "сводка по отправленным на актирование.xlsx" и нажмите ввод', Fore.RESET, end='')
        input()

    while os.path.exists(os.path.join('files', '~$подготовлено к передаче на актирование.xlsx')):
        os.startfile(prepared_file_path)
        print(Fore.RED, 'Закройте файл "подготовлено к передаче на актирование.xlsx" и нажмите ввод', Fore.RESET, end='')        
        input()    

    last_sent_number = get_last_sent_number()

    sent_file_name = f'sent_{ int(last_sent_number)+1 }_{ str(datetime.now())[:19].replace(':', '-') }.xlsx'
    sent_file_path = os.path.join(
        os.getcwd(), 'files', 'sent',
        sent_file_name
        )

    if os.path.exists(prepared_file_path):
        prepared_df = pd.read_excel(prepared_file_path)

        if not prepared_df.empty:
            prepared_df['файл'] = sent_file_name

            if os.path.exists(summary_sent_path):
                summary_prepared_df = pd.read_excel(summary_sent_path)
                summary_prepared_df = pd.concat([summary_prepared_df, prepared_df])
            else:
                summary_prepared_df = prepared_df

              

            os.rename(prepared_file_path, sent_file_path)
            summary_prepared_df.to_excel(summary_sent_path, index=False)

            print(Fore.GREEN, f'Передан на актирование ./files/sent/{ sent_file_name }', Fore.RESET)
            
        else:
            print(Fore.MAGENTA, 'Файл подготовленный для актирования пуст', Fore.RESET)
    else:
        print(Fore.MAGENTA, 'Нет файлов подготовленных для отправки на актирование', Fore.RESET)
    

if __name__ == '__main__':
    sent_maker()    
