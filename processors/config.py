import os

template_column_names = ['Номер контейнера', 'Дата сдачи в депо КНР', 'Депо сдачи в КНР', 'Ставка доплаты, руб', 'Дата актирования', 'Курс из iSales', 'Ставка из iSales']

source_file_path = os.path.join(os.getcwd(), 'свод по каждому заказу.xlsx')
summary_source_path = os.path.join(os.getcwd(), 'files', 'сводка по исходному файлу.xlsx')
prepared_file_path = os.path.join(os.getcwd(), 'files', 'подготовлено к передаче на актирование.xlsx')
summary_prepared_path = os.path.join(os.getcwd(), 'files', 'сводка подготовки к передаче на актирование.xlsx')
summary_sent_path = os.path.join(os.getcwd(), 'files', 'сводка по отправленным на актирование.xlsx')

sent_dir = os.path.join(os.getcwd(), 'files', 'sent')

if not os.path.exists(sent_dir):
    os.makedirs(sent_dir)
