# Author Loik Andrey 7034@balancedv.ru
# TODO
#  Часть 1
#  1. Научится читать комментарии и выделять из них кратность
#  2. В формулу расчета ДМОср добавить учёт кратности по позиции

import pandas as pd
import os

FOLDER = 'Исходные данные'
SALES_NAME = 'ОТ'
MIN_STOCK_NAME = 'МО'
NEW_FILE_NAME = 'Анализ установленного МО.xlsx'


def Run():
    salesFilelist = search_file(SALES_NAME)  # запускаем функцию по поиску файлов и получаем список файлов
    minStockFilelist = search_file(MIN_STOCK_NAME)  # запускаем функцию по поиску файлов и получаем список файлов
    df_sales = create_df(salesFilelist, SALES_NAME)
    df_minStock = create_df(minStockFilelist, MIN_STOCK_NAME)
    df_general = concat_df(df_sales, df_minStock)
    df_general = transfer_MO(df_general)
    df_write_xlsx(df_general)


def search_file(name):
    """
    :param name: Поиск всех файлов в папке FOLDER, в наименовании которых, содержится name
    :return: filelist список с наименованиями фалов
    """
    filelist = []
    for item in os.listdir(FOLDER):
        if name in item and item.endswith('.xlsx'):  # если файл содержит name и с расширением .xlsx, то выполняем
            # Добавляем в список папку и имя файла для последующего обращения из списка
            filelist.append(FOLDER + "/" + item)
        else:
            pass
    return filelist


def create_df(file_list, add_name):
    """
    :param file_list: Загружаем в DataFrame файлы из file_list
    :param add_name: Добавляем add_name в наименование колонок DataFrame
    :return: df_result Дата фрэйм с данными из файлов
    """
    df = pd.DataFrame()
    for filename in file_list:  # проходим по каждому элементу списка файлов
        # print(filename)  # для тестов выводим в консоль наименование файла с которым проходит работа
        df = read_my_excel(filename)

        if add_name == MIN_STOCK_NAME:
            df_search_header = df.iloc[:15, :2]  # для ускорения работы выбираем из DataFrame первую колонку и 15 строк
            # print (df_search_header)
            # создаём маску и отмечаем True строку где есть слово "Номенклатура", остальные False
            mask = (df_search_header == 'Номенклатура')
            # Преобразуем Dataframe согласно маски. После обработки все значения будут NaN кроме нужного нам.
            # В этой же строке кода удаляем все строки со значением NaN и далее получаем индекс оставшейся строки
            f = df_search_header[mask].dropna(axis=0,
                                              how='all').index.values  # Удаление пустых колонок, если axis=0, то строк
            # print (df.iloc[:15, :2])
            df = df.iloc[int(f):, :]  # Убираем все строки с верха DF до заголовков
            df = df.dropna(axis=1, how='all')  # Убираем пустые колонки
            # df.iloc[0, :] = df.iloc[0, :] + ' ' + add_name # Добавляем в наименование тип данных
            df.iloc[0, 0] = 'Код'
            df.iloc[0, 1] = 'Номенклатура'
            df.columns = df.iloc[
                0]  # Значения из найденной строки переносим в заголовки DataFrame для простоты дальнейшего обращения
            df.columns.name = None
            df = df.iloc[2:, :]  # Убираем две строки с верха DF
            df['Номенклатура'] = df['Номенклатура'].str.strip()  # Удалить пробелы с обоих концов строки в ячейке
            df.set_index(['Номенклатура'], inplace=True)  # переносим колонки в индекс, для упрощения дальнейшей работы
            df.iloc[:, 1:2] = df.iloc[:, 1:2].fillna(0)

            # print (df.iloc[:10, 1:2])
            # return df.to_excel('test.xlsx')

        # Добавляем преобразованный DF в результирующий DF

        # Добавляем в результирующий DF по продажам расчётные данные
        elif add_name == SALES_NAME:
            df_search_header = df.iloc[:15, :6]  # для ускорения работы выбираем из DataFrame первую колонку и 15 строк
            # создаём маску и отмечаем True строку где есть слово "Номенклатура", остальные False
            mask = (df_search_header == 'Номенклатура')
            # Преобразуем Dataframe согласно маски. После обработки все значения будут NaN кроме нужного нам.
            # В этой же строке кода удаляем все строки со значением NaN и далее получаем индекс оставшейся строки
            f = df[mask].dropna(axis=0, how='all').index.values  # Удаление пустых колонок, если axis=0, то строк
            col = df[mask].dropna(axis=1, how='all').columns.values
            df = df.iloc[int(f) + 1:, :]  # Убираем все строки с верха DF до заголовков
            df.iloc[0, col] = 'Номенклатура'
            df = df.dropna(axis=1, how='all')  # Убираем пустые колонки
            # Значения из найденной строки переносим в заголовки DataFrame для простоты дальнейшего обращения
            df.columns = df.iloc[0]
            df.columns.name = None
            df = df.iloc[2:, :]  # Убираем две строки с верха DF
            df['Номенклатура'] = df['Номенклатура'].str.strip()  # Удалить пробелы с обоих концов строки в ячейке
            df.set_index(['Номенклатура'], inplace=True)  # переносим колонки в индекс, для упрощения дальнейшей работы
            list = ['ДМО', 'ДМОk']
            df = df[list]
            # Получаем из комментария в эксель значения кратности
            from openpyxl import load_workbook
            d = {'Номенклатура': [], 'Кратность': []}
            df_multiplicity = pd.DataFrame(data=d)
            wb = load_workbook(filename)
            ws = wb["TDSheet"]  # or whatever sheet name
            crat = 'Кратность'
            for row in ws.rows:
                i = 1
                value_crat = 1
                for cell in row:
                    if crat in str(cell.comment) and i == 1:
                        index = str(cell.comment).find(crat)
                        value_crat = int(str(cell.comment)[index + 14: index + 16])
                        i += 1

                value = row[5].value
                if value is not None and value != 'Номенклатура':
                    df_multiplicity.loc[len(df_multiplicity)] = [value, value_crat]

            try:
                # Удалить пробелы с обоих концов строки в ячейке
                df_multiplicity['Номенклатура'] = df_multiplicity['Номенклатура'].str.strip()
                print('Удаляем пробелы из Номенклатуры для сопоставления')
            except:
                print('Нет пробелов в Номенклатуре')
            df_multiplicity.set_index(['Номенклатура'], inplace=True)
            # df_multiplicity.to_excel('test1.xlsx')
            # df = pd.concat([df, df_multiplicity], axis=1, ignore_index=False)
            df = concat_df(df, df_multiplicity)
            df['Кратность'] = df['Кратность'].fillna(1)

            try:
                df['ДМО'] = df['ДМО'].str.replace(',', '.').astype(
                    float)  # заменяем запятые на точки и преобразуем в числовой формат
                df['ДМОk'] = df['ДМОk'].str.replace(',', '.').astype(
                    float)  # заменяем запятые на точки и преобразуем в числовой формат
                print('Заменили тип данных с текст на число в колнках ДМО и ДМОk')
            except:
                print('Колнки ДМО и ДМОk не требуют преобразований')

            df[['ДМО', 'ДМОk']] = df[['ДМО', 'ДМОk']].fillna(0)
            df['ДМОср'] = round((df['ДМО'] + df['ДМОk'] + 0.00000001) / 2)  # округляем по мат. правилам
            df['ДМОср'] = round((df['ДМОср'] + 0.00000001) / df['Кратность']) * df[
                'Кратность']  # округляем до кратности
            mask_05 = df['ДМОk'] == 0.5
            df['ДМОср'] = df['ДМОср'].mask(mask_05, 0.5)

            # df.to_excel('test.xlsx')
            # TODO Объединить с фалом продаж
        # df_result = payment(df_result)
    return df


def read_my_excel(file_name):
    """
    Пытаемся прочитать файл xlxs, если не получается, то исправляем ошибку и опять читаем файл
    :param file_name: Имя файла для чтения
    :return: DataFrame
    """
    print('Попытка загрузки файла:' + file_name)
    try:
        if SALES_NAME in file_name:
            df = pd.read_excel(file_name, sheet_name='TDSheet', header=None, skipfooter=1, engine='openpyxl')
        else:
            df = pd.read_excel(file_name, sheet_name='TDSheet', header=None, skipfooter=0, engine='openpyxl')
        return (df)
    except KeyError as Error:
        print(Error)
        df = None
        if str(Error) == "\"There is no item named 'xl/sharedStrings.xml' in the archive\"":
            bug_fix(file_name)
            print('Исправлена ошибка: ', Error, f'в файле: \"{file_name}\"\n')
            if SALES_NAME in file_name:
                df = pd.read_excel(file_name, sheet_name='TDSheet', header=None, skipfooter=1, engine='openpyxl')
            else:
                df = pd.read_excel(file_name, sheet_name='TDSheet', header=None, skipfooter=0, engine='openpyxl')
            return df
        else:
            print('Ошибка: >>' + str(Error) + '<<')


def bug_fix(file_name):
    """
    Переименовываем не корректное имя файла в архиве excel
    :param file_name: Имя excel файла
    """
    import shutil
    from zipfile import ZipFile
    from rarfile import RarFile

    # Создаем временную папку
    tmp_folder = '/temp/'
    os.makedirs(tmp_folder, exist_ok=True)

    # Распаковываем excel как zip в нашу временную папку и удаляем excel
    try:
        with ZipFile(file_name) as excel_container:
            excel_container.extractall(tmp_folder)
    except:
        with RarFile(file_name) as excel_container:
            excel_container.extractall(tmp_folder)
    os.remove(file_name)

    # Переименовываем файл с неверным названием
    wrong_file_path = os.path.join(tmp_folder, 'xl', 'SharedStrings.xml')
    correct_file_path = os.path.join(tmp_folder, 'xl', 'sharedStrings.xml')
    os.rename(wrong_file_path, correct_file_path)

    # Запаковываем excel обратно в zip и переименовываем в исходный файл
    try:
        shutil.make_archive(f'{FOLDER}/correct_file', 'zip', tmp_folder)
    except:
        shutil.make_archive(f'{FOLDER}/correct_file', 'rar', tmp_folder)
    os.rename(f'{FOLDER}/correct_file.zip', file_name)


def concat_df(df1, df2):
    df = pd.concat([df1, df2], axis=1, ignore_index=False)
    return df


def transfer_MO(df):
    # print(df.head())
    # print(df.info())
    # print(df.iloc[:5, :])

    # =ЕСЛИ(И(C3>0;C3<1);C3;ЕСЛИ(И(C3>=1;H3>=1);H3+ОСТАТ(C3;1);ЕСЛИ(H3=0;0,33;H3)))

    # Добавляем к ДМОср остаток от МО подразделения
    df['ДМОср тех'] = (df.iloc[:, 5] % 1)
    df['ДМОср тех'] = df['ДМОср тех'] + df['ДМОср']

    # Подставляем в ДМОср значения МО если 0<МО<1
    df['МО перенос'] = df['ДМОср']
    df['МО перенос'] = df['МО перенос'].mask((df['ДМОср'] == 0), 0.33)
    mask_MO1 = (df.iloc[:, 5] >= 1).values & (df['ДМОср'] >= 1).values
    df['МО перенос'] = df['МО перенос'].mask(mask_MO1, df['ДМОср тех'].values)
    mask_MO2 = (df.iloc[:, 5] > 0) & (df.iloc[:, 5] < 1)
    df['МО перенос'] = df['МО перенос'].mask(mask_MO2.values, df.iloc[:, 5].values)
    df = df.drop(['ДМОср тех'], axis=1)
    df['Расхождения'] = df['МО перенос'] - df.iloc[:, 5]
    df['Номенклатура'] = df.index
    df.set_index(['Код', 'Номенклатура'], inplace=True)
    # print(df.iloc[255:260, 6])
    # print(df.iloc[:10, [6, 7]])

    df.to_excel('test2.xlsx')
    return df


def df_write_xlsx(df):
    # Сохраняем в переменные значения конечных строк и столбцов
    row_end, col_end = len(df), len(df.columns)
    row_end_str, col_end_str = str(row_end), str(col_end)

    # Сбрасываем встроенный формат заголовков pandas
    pd.io.formats.excel.ExcelFormatter.header_style = None

    # Создаём эксель и сохраняем данные
    name_file = NEW_FILE_NAME
    sheet_name = 'Данные'  # Наименование вкладки для сводной таблицы
    writer = pd.ExcelWriter(name_file, engine='xlsxwriter')
    workbook = writer.book
    df.to_excel(writer, sheet_name=sheet_name)
    wks1 = writer.sheets[sheet_name]  # Сохраняем в переменную вкладку для форматирования

    # Получаем словари форматов для эксель
    header_format, con_format, border_storage_format_left, border_storage_format_right, \
    name_format, MO_format, data_format = format_custom(workbook)

    # Форматируем таблицу
    wks1.set_default_row(12)
    wks1.set_row(0, 20, header_format)
    wks1.set_column('A:A', 12, name_format)
    wks1.set_column('B:B', 32, name_format)
    wks1.set_column('C:H', 10, data_format)
    wks1.set_column('I:I', 12, data_format)

    # Делаем жирным рамку между складами и форматируем колонку с МО по всем складам
    wks1.set_column(2, 2, None, border_storage_format_left)
    wks1.set_column(5, 5, None, border_storage_format_right)
    wks1.set_column(6, 6, None, border_storage_format_left)
    wks1.set_column(7, 7, None, border_storage_format_right)
    wks1.set_column(7, 7, None, MO_format)

    # Добавляем фильтр в первую колонку
    wks1.autofilter(0, 0, row_end + 1, col_end + 1)

    # Сохраняем файл
    writer.save()
    return


def format_custom(workbook):
    header_format = workbook.add_format({
        'font_name': 'Arial',
        'font_size': '7',
        'align': 'center',
        'valign': 'top',
        'text_wrap': True,
        'bold': True,
        'bg_color': '#F4ECC5',
        'border': True,
        'border_color': '#CCC085'
    })

    border_storage_format_left = workbook.add_format({
        'num_format': '# ### ##0.00',
        'font_name': 'Arial',
        'font_size': '8',
        'left': 2,
        'left_color': '#000000',
        'bottom': True,
        'bottom_color': '#CCC085',
        'top': True,
        'top_color': '#CCC085',
        'right': True,
        'right_color': '#CCC085',
    })
    border_storage_format_right = workbook.add_format({
        'num_format': '# ### ##0.00',
        'font_name': 'Arial',
        'font_size': '8',
        'right': 2,
        'right_color': '#000000',
        'bottom': True,
        'bottom_color': '#CCC085',
        'top': True,
        'top_color': '#CCC085',
        'left': True,
        'left_color': '#CCC085',
    })

    name_format = workbook.add_format({
        'font_name': 'Arial',
        'font_size': '8',
        'align': 'left',
        'valign': 'top',
        'text_wrap': True,
        'bold': False,
        'border': True,
        'border_color': '#CCC085'
    })

    MO_format = workbook.add_format({
        'num_format': '# ### ##0.00;;',
        'bold': True,
        'font_name': 'Arial',
        'font_size': '8',
        'font_color': '#FF0000',
        'right': 2,
        'right_color': '#000000',
        'bottom': True,
        'bottom_color': '#CCC085',
        'top': True,
        'top_color': '#CCC085',
        'left': True,
        'left_color': '#CCC085',
    })
    data_format = workbook.add_format({
        'num_format': '# ### ##0.00',
        'font_name': 'Arial',
        'font_size': '8',
        'text_wrap': True,
        'border': True,
        'border_color': '#CCC085'
    })
    con_format = workbook.add_format({
        'bg_color': '#FED69C',
    })

    return header_format, con_format, border_storage_format_left, border_storage_format_right, \
           name_format, MO_format, data_format


if __name__ == '__main__':
    Run()
