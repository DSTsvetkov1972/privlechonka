import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
import sys

sys.path.append(os.getcwd())
from processors.fns import source_file_checker
from colorama import Fore

from processors.config import source_file_path, prepared_file_path, summary_source_path, summary_sent_path

def source_checker():
    while os.path.exists(os.path.join(os.getcwd(), 'files', '~$сводка по исходному файлу.xlsx')):
        os.startfile(summary_source_path)
        print(Fore.RED, 'Закройте файл "сводка по исходному файлу.xlsx" и нажмите ввод', Fore.RESET, end='')
        input()
        
    sorce_checker_res = source_file_checker(source_file_path, prepared_file_path, summary_sent_path)
    summary_df = sorce_checker_res['check_df']   

    summary_df.to_excel(summary_source_path, index=False, sheet_name='Sheet1')


    wb = load_workbook(summary_source_path)

    ws = wb['Sheet1']

    ws.freeze_panes = 'A2'

    # Устанавливаем ширину для конкретной колонки
    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 32
    ws.column_dimensions['C'].width = 32
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 16
    ws.column_dimensions['F'].width = 22
    ws.column_dimensions['G'].width = 22
    ws.column_dimensions['H'].width = 26
    ws.column_dimensions['I'].width = 42
    ws.column_dimensions['J'].width = 26

    for col in 'ABCDEFGHIJ':
        cell = ws[f'{ col }1']
        cell.font = Font(bold=True)  # Жирный шрифт
        cell.alignment = Alignment(horizontal='center', vertical='center')  # Выравнивание по центру

    wb.save(summary_source_path)

    os.startfile(summary_source_path)

    print(Fore.GREEN, 'Сводка по файлу "свод по каждому заказу.xlsx" подготовлена и будет открыта поверх всех окон на рабочем столе!' ,Fore.RESET)

if __name__ == '__main__':
    source_checker()




