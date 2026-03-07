import pandas as pd
import os
from colorama import Fore
from config import template_column_names, summary_sent_path


def depo_cost_parser(s):

    try:
        # s = str(s)
        s = s.replace(' ', '').replace('\n', ';')
        #print(s)
        l = s.split(';')
        l = [depo_cost.split(':') for depo_cost in l if depo_cost != '']

        #print(l)
        l = [(depo_cost[0].split(','), depo_cost[1]) for depo_cost in l]

        res = {}
        for depos, cost in l:
            for depo in depos:
                res[depo] = float(cost)
        return(True, res)

    except Exception as e:
        return (False, str(e))
    

def sorce_checker(source_file_path, prepared_file_path, summary_sent_path):
    print('считываем листы в память')

    all_sheets_dict = pd.read_excel(source_file_path, sheet_name=None)
    correct_sheets_dict = {} # В этот словарь будут помещенны датафреймы листов, в которых нет ошибок

    #print(all_sheets_dict)

    summary_list = []

    for sheet, df in all_sheets_dict.items():
        # print(Fore.MAGENTA, sheet, Fore.RESET)
        df_columns = list(df.columns)
        
        columns_not_in_template = [df_column for df_column in df_columns if df_column not in template_column_names]
        if columns_not_in_template:
            columns_not_in_template_err = str(columns_not_in_template)[1:-1]
        else:
            columns_not_in_template_err = '-'

        columns_not_in_df = [template_column for template_column in template_column_names if template_column not in df_columns]
        if columns_not_in_df:
            columns_not_in_df_err = str(columns_not_in_df)[1:-1]
        else:
            columns_not_in_df_err = '-'

        if 'Курс из iSales' in df_columns and len(df)!=0:
            try:
                currency_rate_cell_value = df['Курс из iSales'].iloc[0]
                currency_rate = float(str(currency_rate_cell_value).replace(',', '.'))
                currency_rate_err = '-'
            except ValueError:
                currency_rate_err = f'не число: "{currency_rate_cell_value}"' 
        else:
            currency_rate_err = "не заполнено"


        if 'Ставка из iSales' in df_columns and len(df)!=0:
            depo_cost_cell_value = df['Ставка из iSales'].iloc[0]
            depo_cost_parser_res = depo_cost_parser(depo_cost_cell_value)
            
            if depo_cost_parser_res[0]:
                depo_cost_err = '-' 
            else:
                depo_cost_err = depo_cost_parser_res[1]
        else:
            depo_cost_err = "не заполнено"


        if 'Номер контейнера' in df_columns and len(df)>1:
            conts_qty = len((df['Номер контейнера']))
            last_cont = df['Номер контейнера'].iloc[conts_qty-1]

        if 'Депо сдачи в КНР' in df_columns and len(df)>1:
            depos = df['Депо сдачи в КНР']
            depos = depos.dropna()
            depos = depos[depos != '']
            depos_qty = len(depos)
        

        summary_list.append(
            [sheet,
            columns_not_in_template_err,
            columns_not_in_df_err,
            currency_rate_err,
            depo_cost_err,
            conts_qty,
            last_cont,
            depos_qty
            ])
        
        if (columns_not_in_template_err == '-' and
            columns_not_in_df_err == '-' and
            currency_rate_err == '-' and
            depo_cost_err == '-'):
            correct_sheets_dict[sheet] = all_sheets_dict[sheet]




    summary_df = pd.DataFrame(
        summary_list,
        columns = ['Заказ',
                   'Колонки которых нет в шаблоне',
                   'Нет колонок из шаблона',
                   'Курс из iSales',
                   'Ставка из iSales',
                   'Контейнеров на листе',
                   'Последний контейнер',
                   'Заполненных депо сдачи']
                   )
    
    if os.path.exists(prepared_file_path):
        prepared_df = pd.read_excel(
            prepared_file_path,
            dtype=str)
        prepared_df = prepared_df.rename(columns={'Депо сдачи в КНР': 'Подготовленно к передаче на актирование'})
        prepared_grouped = prepared_df.groupby('Заказ')['Подготовленно к передаче на актирование'].count()
    else:
        prepared_grouped = pd.DataFrame(columns=['Заказ','Подготовленно к передаче на актирование'])

    
    if os.path.exists(summary_sent_path):
        sent_df = pd.read_excel(
            summary_sent_path,
            dtype=str)
        sent_df = sent_df.rename(columns={'Депо сдачи в КНР': 'Передано на актирование'})
        sent_grouped = sent_df.groupby('Заказ')['Передано на актирование'].count()
    else:
        sent_grouped = pd.DataFrame(columns=['Заказ','Передано на актирование'])


    check_df = pd.merge(summary_df, prepared_grouped, on='Заказ', how='left')
    check_df = pd.merge(check_df, sent_grouped, on='Заказ', how='left')

    return {'check_df': check_df,
            'correct_sheets_dict': correct_sheets_dict}    


def get_last_sent_number():
    sent_dir = os.path.join(os.getcwd(), 'files', 'sent')
    sent_files = list(os.walk(sent_dir))[0][2]

    if not sent_files:
        return 0
    else:
        last_sent_file = max(sent_files)
        last_number = last_sent_file.split('_')[1]
        return last_number


if __name__ == '__main__':

    s = """Chengdu, Chongqing, Shanghai: 355;
Huangpu, Dalang, Shilong, Ningbo: 455;
Tianjin, Qingdao: 355;
Dalian: 155;
Changsha, Zhengzhou: 255;
Yiwu, Wuhan: 355;
Hefei: 305;
Taicang, Xi`an: 355;
Xiamen, Shenzhen(Yantian): 455"""
    print(depo_cost_parser(s))