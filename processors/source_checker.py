import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.filters import FilterColumn, Filters
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
    ws.row_dimensions[1].height = 45

    ws.auto_filter.ref = f"A1:{ get_column_letter(ws.max_column) }{ len(summary_df) }"
    col_filter = FilterColumn(colId=1)  # колонка C (индекс 2)
    # col_filter.filters = Filters(filter=["скрытый"])
    ws.auto_filter.filterColumn.append(col_filter)

    ws.freeze_panes = 'A2'

    # Устанавливаем ширину для конкретной колонки


    # ws.column_dimensions['A'].width = 12
    # ws.column_dimensions['B'].width = 16
    # ws.column_dimensions['C'].width = 22
    # ws.column_dimensions['D'].width = 16
    # ws.column_dimensions['E'].width = 20
    # ws.column_dimensions['F'].width = 22
    # ws.column_dimensions['G'].width = 18
    # ws.column_dimensions['H'].width = 18
    # ws.column_dimensions['I'].width = 18
    # ws.column_dimensions['J'].width = 18
    # ws.column_dimensions['K'].width = 18    
    # ws.column_dimensions['L'].width = 18    


    for col in range(1, ws.max_column+1):
        cell = ws.cell(column=col, row=1)
        cell.font = Font(bold=True)  # Жирный шрифт
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.column_dimensions[get_column_letter(col)].width = 18


    for col in range(1, ws.max_column+1):
        for row in range(2, ws.max_row+1):
            cell = ws.cell(column=col, row=row)    
            cell.alignment = Alignment(horizontal='center', vertical='center')  # Выравнивание по центру


    wb.save(summary_source_path)

    os.startfile(summary_source_path)

    print(Fore.GREEN, 'Сводка по файлу "свод по каждому заказу.xlsx" подготовлена и будет открыта поверх всех окон на рабочем столе!' ,Fore.RESET)

if __name__ == '__main__':
    source_checker()




