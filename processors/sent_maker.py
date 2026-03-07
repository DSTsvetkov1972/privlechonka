import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from datetime import datetime

from config import source_file_path, prepared_file_path, summary_prepared_path, summary_sent_path
from fns import get_last_sent_number
from colorama import Fore


while os.path.exists(os.path.join('files', '~$сводка по отправленным на актирование.xlsx')):
    os.startfile(summary_sent_path)
    input('Закройте файл "сводка по отправленным на актирование.xlsx" и нажмите ввод')

while os.path.exists(os.path.join('files', '~$подготовлено к передаче на актирование.xlsx')):
    os.startfile(prepared_file_path)
    input('Закройте файл "подготовлено к передаче на актирование.xlsx" и нажмите ввод')    

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

        sent_file_name    

        os.rename(prepared_file_path, sent_file_path)
        summary_prepared_df.to_excel(summary_sent_path, index=False)
        
    else:
        print('Файл подготовленный для актирования пуст')
else:
    print('Нет файлов подготовленных для отправки на актирование')
    
