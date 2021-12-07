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


def Run():

    salesFilelist = search_file(SALES_NAME)  # запускаем функцию по поиску файлов и получаем список файлов
    minStockFilelist = search_file(MIN_STOCK_NAME)  # запускаем функцию по поиску файлов и получаем список файлов
    df_sales = create_df (salesFilelist, SALES_NAME)
    df_minStock = create_df (minStockFilelist, MIN_STOCK_NAME)
    df_general = concat_df (df_sales, df_minStock)
    df_general = transfer_MO (df_general)
    #df_general.to_excel('test.xlsx')

def search_file(name):
    """
    :param name: Поиск всех файлов в папке FOLDER, в наименовании которых, содержится name
    :return: filelist список с наименованиями фалов
    """
    filelist = []
    for item in os.listdir(FOLDER):
        if name in item and item.endswith('.xlsx'): # если файл содержит name и с расширенитем .xlsx, то выполняем
            filelist.append(FOLDER + "/" + item) # добаляем в список папку и имя файла для последующего обращения из списка
        else:
            pass
    return filelist

def create_df (file_list, add_name):
    """
    :param file_list: Загружаем в DataFrame файлы из file_list
    :param add_name: Добавляем add_name в наименование колонок DataFrame
    :return: df_result Дата фрэйм с данными из файлов
    """
    #df_result = []

    for filename in file_list: # проходим по каждому элементу списка файлов
        print (filename) # для тестов выводим в консоль наименование файла с которым проходит работа
        df = read_my_excel(filename)
        if add_name == MIN_STOCK_NAME:
            df_search_header = df.iloc[:15, :2] # для ускорения работы выбираем из DataFrame первую колонку и 15 строк
            # print (df_search_header)
            # создаём маску и отмечаем True строку где есть слово "Номенклатура", остальные False
            mask = (df_search_header == 'Номенклатура')
            # Преобразуем Dataframe согласно маски. После обработки все значения будут NaN кроме нужного нам.
            # В этой же строке кода удаляем все строки со значением NaN и далее получаем индекс оставшейся строки
            f = df_search_header[mask].dropna(axis=0, how='all').index.values # Удаление пустых колонок, если axis=0, то строк
            # print (df.iloc[:15, :2])
            df = df.iloc[int(f):, :] # Убираем все строки с верха DF до заголовков
            df = df.dropna(axis=1, how='all')  # Убираем пустые колонки
            #df.iloc[0, :] = df.iloc[0, :] + ' ' + add_name # Добавляем в наименование тип данных
            df.iloc[0, 0] = 'Код'
            df.iloc[0, 1] = 'Номенклатура'
            df.columns = df.iloc[0] # Значения из найденной строки переносим в заголовки DataFrame для простоты дальнейшего обращения
            df.columns.name = None
            df = df.iloc[2:, :] # Убираем две строки с верха DF
            df['Номенклатура'] = df['Номенклатура'].str.strip() # Удалить пробелы с обоих концов строки в ячейке
            df.set_index(['Номенклатура'], inplace=True) # переносим колонки в индекс, для упрощения дальнейшей работы
            df.iloc[:, 1:2] = df.iloc[:, 1:2].fillna(0)

            #print (df.iloc[:10, 1:2])
            #return df.to_excel('test.xlsx')

        # Добавляем преобразованный DF в результирующий DF
        #df_result = concat_df(df_result, df)

        # Добавляем в результирующий DF по продажам расчётные данные
        elif add_name == SALES_NAME:
            df_search_header = df.iloc[:15, :5]  # для ускорения работы выбираем из DataFrame первую колонку и 15 строк
            # создаём маску и отмечаем True строку где есть слово "Номенклатура", остальные False
            mask = (df_search_header == 'Номенклатура')
            # Преобразуем Dataframe согласно маски. После обработки все значения будут NaN кроме нужного нам.
            # В этой же строке кода удаляем все строки со значением NaN и далее получаем индекс оставшейся строки
            f = df[mask].dropna(axis=0, how='all').index.values  # Удаление пустых колонок, если axis=0, то строк
            df = df.iloc[int(f) + 1:, :]  # Убираем все строки с верха DF до заголовков
            df = df.dropna(axis=1, how='all')  # Убираем пустые колонки
            df.iloc[0, 3] = 'Номенклатура'
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
                value_crat = 1
                i = 1
                for cell in row:
                    if crat in str(cell.comment) and i == 1:
                        index = str(cell.comment).find(crat)
                        value_crat = int(str(cell.comment)[index + 14])
                        df_multiplicity.loc[len(df_multiplicity)] = [row[4].value, value_crat]
                        i += 1
            df_multiplicity.set_index(['Номенклатура'], inplace=True)
            #df_multiplicity.to_excel('test1.xlsx')
            #df = pd.concat([df, df_multiplicity], axis=1, ignore_index=False)
            df = concat_df(df, df_multiplicity)
            df['Кратность'] = df['Кратность'].fillna(1)
            df[['ДМО','ДМОk']] = df[['ДМО','ДМОk']].fillna(0)
            #df['ДМОср_тех'] = (df['ДМО'] + df['ДМОk']) / 2 # создаём базовое среднее для проверки
            df['ДМОср'] = round((df['ДМО'] + df['ДМОk'] + 0.00000001) / 2) # округляем по мат. правилам
            df['ДМОср'] = round((df['ДМОср'] + 0.00000001) / df['Кратность']) * df['Кратность'] # округляем до кратности
            mask_05 = df['ДМОk'] == 0.5
            df['ДМОср'] = df['ДМОср'].mask(mask_05, 0.5)

            #df.to_excel('test.xlsx')
            # TODO Объединить с фалом продаж
        #df_result = payment(df_result)
    return df

def read_my_excel (file_name):
    """
    Пытаемся прочитать файл xlxs, если не получается, то исправляем ошибку и опять читаем файл
    :param file_name: Имя файла для чтения
    :return: DataFrame
    """
    print ('Попытка загрузки файла:'+file_name)
    try:
        if SALES_NAME in file_name:
            df = pd.read_excel(file_name, sheet_name='TDSheet', header=None, skipfooter=1, engine='openpyxl')
        else:
            df = pd.read_excel(file_name, sheet_name='TDSheet', header=None, skipfooter=0, engine='openpyxl')
        return (df)
    except KeyError as Error:
        print (Error)
        df = None
        if str(Error) == "\"There is no item named 'xl/sharedStrings.xml' in the archive\"":
            bug_fix (file_name)
            print('Исправлена ошибка: ', Error, f'в файле: \"{file_name}\"\n')
            if SALES_NAME in file_name:
                df = pd.read_excel(file_name, sheet_name='TDSheet', header=None, skipfooter=1, engine='openpyxl')
            else:
                df = pd.read_excel(file_name, sheet_name='TDSheet', header=None, skipfooter=0, engine='openpyxl')
            return df
        else:
            print('Ошибка: >>' + str(Error) + '<<')

def bug_fix (file_name):
    """
    Переименовываем не корректное имя файла в архиве excel
    :param file_name: Имя excel файла
    """
    import shutil
    from zipfile import ZipFile

    # Создаем временную папку
    tmp_folder = '/temp/'
    os.makedirs(tmp_folder, exist_ok=True)

    # Распаковываем excel как zip в нашу временную папку и удаляем excel
    with ZipFile(file_name) as excel_container:
        excel_container.extractall(tmp_folder)
    os.remove(file_name)

    # Переименовываем файл с неверным названием
    wrong_file_path = os.path.join(tmp_folder, 'xl', 'SharedStrings.xml')
    correct_file_path = os.path.join(tmp_folder, 'xl', 'sharedStrings.xml')
    os.rename(wrong_file_path, correct_file_path)

    # Запаковываем excel обратно в zip и переименовываем в исходный файл
    shutil.make_archive(f'{FOLDER}/correct_file', 'zip', tmp_folder)
    os.rename(f'{FOLDER}/correct_file.zip', file_name)

def concat_df (df1, df2):
    df = pd.concat([df1, df2], axis=1, ignore_index=False, levels=['Номенклатура'])
    return df

def transfer_MO (df):
    print (df.head())
    print (df.iloc[:5, 3:4])
    # =ЕСЛИ(И(C3>0;C3<1);C3;ЕСЛИ(И(C3>=1;H3>=1);H3+ОСТАТ(C3;1);ЕСЛИ(H3=0;0,33;H3)))
    mask_MO1 = (df.iloc[:, 5:6] > 0) == (df.iloc[:, 5:6] < 1)
    #ask_MO2 = (df.iloc[:, 5:6] >= 1) == (df.iloc[:, 3:4] >= 1)
    mask_MO2 = (df.iloc[:, 5:6]) == (df.iloc[:, 3:4])
    # TODO Исправить ошибку: Можно сравнивать только объекты DataFrame с одинаковой меткой
    print (mask_MO2)
    #mask_MO2.to_excel('test2.xlsx')

if __name__ == '__main__':
    Run()

