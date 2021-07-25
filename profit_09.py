#!/usr/bin/env python
# -*- coding: utf-8 -*-
#

import os
import time
import datetime
import subprocess

import pandas as pd
import xlsxwriter


def main(args):
    infolder = './in/'
    outfolder = './out/'
    outfile = outfolder + 'out_' + time.strftime("%Y%m%d_%H%M%S") + '.xlsx'

    out_sheet_name = 'ProfitSum'  # имя листа выходного файла
    # Столбцы выходного файла
    cols = ['Trade #', 'Symbol Name', 'Order #', 'Type', 'Signal', 'Date', 'Time',
            'Price', 'Contracts', 'Profit ()', 'Месяц P', 'Месяц E', 'SUM Profit P', 'SUM Profit E', 'Filename']

    df_result = pd.DataFrame(columns=cols)  # результирующая таблица
    df_address = pd.DataFrame(
        columns=['b', 'e'])  # таблица диапазанов по которым считаем суммы формулами excel в xlsxwriter

    in_sheet_name = 'List of Trades'  # имя листа входного файла
    # Заголовки столбцов входного файла (считаем что заголовки находятся на 3 строке (header = 2))
    ucols = ['Trade #', 'Symbol Name', 'Order #', 'Type', 'Signal', 'Date', 'Time', 'Price', 'Contracts', 'Profit ()']

    address_shift = 0  # смещение в выходной таблице для записи сумм

    infiles = os.listdir(infolder)

    for filename in infiles:  # по всем файлам

        print('\n' + filename)

        filepath = infolder + filename
        # tmp_df - исходный лист отдельного файла
        tmp_df = pd.read_excel(filepath, in_sheet_name, header=2, usecols=ucols)  # header = 2 - заголовки на 3 строке
        tmp_df['Filename'] = filename
        tmp_df['Месяц P'] = tmp_df['Date'].dt.month
        tmp_df['Date'] = pd.to_datetime(tmp_df['Date'], errors='coerce')  # в дату
        tmp_df['Date'] = tmp_df['Date'].dt.date

        for m in range(1, 13):  # выборка по месяцам

            df_month = tmp_df.loc[tmp_df['Месяц P'] == m].copy()  # df_month - отдельный месяц

            if not df_month.empty:
                print(datetime.date(1900, m, 1).strftime('%B'))

                sum_month = df_month['Profit ()'].sum()  # искомая сумма
                # добавить отдельный месяц в результат:
                df_result = pd.concat([df_result, df_month], sort=False, ignore_index=True)
                df_result.loc[address_shift, "SUM Profit P"] = sum_month  # прописать сумму

                address_shift_b = address_shift + 2  # начало диапазона суммы в excel
                address_shift += df_month.shape[0]
                address_shift_e = address_shift + 1  # конец диапазона суммы в excel
                # диапазон в таблицу диапазонов:
                df_address = df_address.append({'b': address_shift_b, 'e': address_shift_e}, ignore_index=True)

    writer = pd.ExcelWriter(outfile, engine='xlsxwriter')  # используем XlsxWriter как движек экспорта

    df_result.to_excel(writer, sheet_name=out_sheet_name,
                       index=False)  # Convert the dataframe to an XlsxWriter Excel object.
    workbook = writer.book  # Get the xlsxwriter objects from the dataframe writer object.
    worksheet = writer.sheets[out_sheet_name]  # получили наш лист

    # Немного форматирования:

    worksheet.set_column('A:A', 8)
    worksheet.set_column('B:B', 14)
    # worksheet.set_column('C:C', 10)
    # worksheet.set_column('D:D', 14)
    # worksheet.set_column('E:E', 14)
    worksheet.set_column('F:F', 12)
    worksheet.set_column('G:G', 8)
    worksheet.set_column('H:H', 10)
    worksheet.set_column('I:I', 10)
    worksheet.set_column('J:J', 10)
    worksheet.set_column('M:M', 12)
    worksheet.set_column('N:N', 12)

    worksheet.set_row(0, 33)  # Высота заголовка
    worksheet.freeze_panes(1, 2)  # закрепить области

    # автофильтр
    worksheet.autofilter('A1:O' + str(df_result.shape[0] + 1).strip())
    # print('A1:O' + str(df_result.shape[0]+1).strip())

    # прописываем формулы
    # месяца
    for row_num in range(1, df_result.shape[0] + 1):
        worksheet.write_formula(row_num, 11, "=MONTH(F%d)" % (row_num + 1), None,
                                '')  # parameters None, '' force recalculate result libreoffice

    address_list = df_address.values.tolist()  # таблицу смещений в список
    # суммы
    for a in address_list:
        worksheet.write_formula(a[0] - 1, 13, "=SUM(J%d:J%d)" % (a[0], a[1]), None,
                                '')  # parameters None, '' force recalculate result libreoffice

    workbook.close()  # close() записывает все данные в файл xlsx и закрывает его:
    writer.save()  # Close the Pandas Excel writer and output the Excel file.

    # subp = subprocess.Popen(["xdg-open", outfile]) # Lin
    subp = subprocess.Popen(os.path.abspath(outfile), shell=True)  # Win #, creationflags=DETACHED_PROCESS) #
    return 0


if __name__ == '__main__':
    import sys

    sys.exit(main(sys.argv))
