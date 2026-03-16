import os
import warnings

warnings.filterwarnings("ignore", message="Conditional Formatting extension is not supported and will be removed")

template_column_names = [
    'Номер контейнера',
    'Дата сдачи в депо КНР',
    'Депо сдачи в КНР',
    'Ставка доплаты, руб',
    'Дата актирования',
    'Курс из iSales',
    'Ставка из iSales',
    'Номер решения ЭС',
    'Файл с решением ЭС',
    'Комментарий'
    ]

files_dir = os.path.join(os.getcwd(), 'Файлы')
tmp_files_dir = os.path.join(files_dir, 'tmp_files')

decisions_dir = os.path.join(files_dir, 'Решения ЭС')
sent_dir = os.path.join(files_dir, 'Отправленные на актирование')

source_file_path = os.path.join(os.getcwd(), 'свод по каждому заказу.xlsx')
summary_source_path = os.path.join(tmp_files_dir, 'сводка по исходному файлу.xlsx')
prepared_file_path = os.path.join(tmp_files_dir, 'подготовлено к передаче на актирование.xlsx')
summary_prepared_path = os.path.join(tmp_files_dir, 'сводка подготовки к передаче на актирование.xlsx')
summary_sent_path = os.path.join(files_dir, 'Отправленные на актирование', '.Реестр отправленных на актирование.xlsx')
