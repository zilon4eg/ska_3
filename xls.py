import xlwings as xw
import datetime
from win32com.client import Dispatch
from winsys import fs
import os


def directory_access(directory_path):
    directory = fs.dir(directory_path)
    all_access = []
    for ace in directory.security().dacl:
        user = str(ace.trustee)[str(ace.trustee).find('\\') + 1:]
        access_flags = fs.FILE_ACCESS.names_from_value(ace.access)
        if user not in ['Администратор', 'Прошедшие проверку', 'СИСТЕМА', 'Администраторы', 'Пользователи', 'Администраторы домена']:

            access = []
            # ===(чтение)===
            read = ['ALL_ACCESS', 'GENERIC_EXECUTE', 'GENERIC_READ', 'GENERIC_WRITE', 'LIST_DIRECTORY', 'READ_ATTRIBUTES', 'READ_DATA', 'READ_EA', 'STANDARD_RIGHTS_READ', 'STANDARD_RIGHTS_WRITE', 'SYNCHRONIZE', 'READ_CONTROL']
            # ===(запись)===
            write = ['ADD_FILE', 'ADD_SUBDIRECTORY', 'ALL_ACCESS', 'APPEND_DATA', 'CREATE_PIPE_INSTANCE', 'GENERIC_EXECUTE', 'GENERIC_READ', 'GENERIC_WRITE', 'WRITE_ATTRIBUTES', 'WRITE_DATA', 'WRITE_EA', 'SYNCHRONIZE']
            # ===(чтение и выполнение + список содержимого папки + чтение)===
            execution = ['ALL_ACCESS', 'GENERIC_EXECUTE', 'GENERIC_READ', 'GENERIC_WRITE', 'LIST_DIRECTORY', 'READ_ATTRIBUTES', 'READ_DATA', 'READ_EA', 'TRAVERSE', 'STANDARD_RIGHTS_READ', 'STANDARD_RIGHTS_WRITE', 'SYNCHRONIZE', 'READ_CONTROL']
            # ===(изменение + чтение и выполнение + список содержимого папки + чтение + запись)===
            modification = ['ADD_FILE', 'ADD_SUBDIRECTORY', 'ALL_ACCESS', 'APPEND_DATA', 'CREATE_PIPE_INSTANCE', 'GENERIC_EXECUTE', 'GENERIC_READ', 'GENERIC_WRITE', 'LIST_DIRECTORY', 'READ_ATTRIBUTES', 'READ_DATA', 'READ_EA', 'TRAVERSE', 'WRITE_ATTRIBUTES', 'WRITE_DATA', 'WRITE_EA', 'STANDARD_RIGHTS_READ', 'STANDARD_RIGHTS_WRITE', 'SYNCHRONIZE', 'DELETE', 'READ_CONTROL']
            # ===(полный доступ)===
            full_access = ['ADD_FILE', 'ADD_SUBDIRECTORY', 'ALL_ACCESS', 'APPEND_DATA', 'CREATE_PIPE_INSTANCE', 'DELETE_CHILD', 'GENERIC_EXECUTE', 'GENERIC_READ', 'GENERIC_WRITE', 'LIST_DIRECTORY', 'READ_ATTRIBUTES', 'READ_DATA', 'READ_EA', 'TRAVERSE', 'WRITE_ATTRIBUTES', 'WRITE_DATA', 'WRITE_EA', 'STANDARD_RIGHTS_READ', 'STANDARD_RIGHTS_WRITE', 'SYNCHRONIZE', 'DELETE', 'READ_CONTROL', 'WRITE_DAC', 'WRITE_OWNER']

            # set.issubset(other) или set <= other - все элементы set принадлежат other
            if set(full_access).issubset(set(access_flags)):
                access.append('Полный доступ')
            if set(modification).issubset(set(access_flags)):
                access.append('Изменение')
            if set(write).issubset(set(access_flags)):
                access.append('Запись')
            if set(execution).issubset(set(access_flags)):
                access.append('Чтение и выполнение')
            if set(read).issubset(set(access_flags)):
                access.append('Чтение')
            access.append(user)
            access.append(directory_path[directory_path.rfind('\\') + 1:])
            access.reverse()
            for i in range(5):
                if len(access) < 7:
                    access.append(None)
            all_access.append(access)
    return all_access


def dir_list(root_path):
    return list(item for item in os.listdir(root_path) if os.path.isdir(os.path.join(root_path, item)))


def directories_access(root_path):
    access = []
    for directory in dir_list(root_path):
        path = f'{root_path}\\{directory}'
        for dir in directory_access(path):
            access.append(dir)
    return access


def auto_size_column(book_path, sheet_name):
    """выбирает оптимальную ширину столбцов листа"""
    # from win32com.client import Dispatch
    excel = Dispatch('Excel.Application')
    wb = excel.Workbooks.Open(book_path)
    excel.Worksheets(sheet_name).Activate()
    excel.ActiveSheet.Columns.AutoFit()
    wb.Save()
    wb.Close()
    excel.Quit()


def difference_lists(list0, list1):
    """вычитает один список списков из другого"""
    set0 = set(map(lambda x: tuple(x), list0))
    set1 = set(map(lambda x: tuple(x), list1))
    set_result = set0 - set1
    list_result = list(map(lambda x: list(x), set_result))
    return list_result


def ws_now(work_book):
    """создает название листа на основании текущей даты (если такой лист уже есть, добавляет время)"""
    ws_name1 = datetime.datetime.today().strftime('%d-%m-%Y')
    ws_name2 = datetime.datetime.today().strftime('%d-%m-%Y %H-%M-%S')
    return ws_name1 if ws_name1 not in list(sheet.name for sheet in work_book.sheets) else ws_name2


def number_to_letter(number):
    """преобразует число в букву"""
    return 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[number - 1]


def coloring_on_difference(ws, difference, data, selected_color):
    """
    Окрашивает изменившиеся ячейки относительно дельты
    :param ws: рабочий лист
    :param difference: дельта изменений
    :param data: данные листа
    :param selected_color: выбранный цвет заливки
    """
    for id, str_data in enumerate(data):
        for str_diff in difference:
            print(difference)
            print(str_diff)
            if str_data[0] == str_diff[0] and str_data[1] == str_diff[1]:
                start_range = number_to_letter(1)
                end_range = number_to_letter(len(str_diff))
                ws.range(f'{start_range}{id + 1}:{end_range}{id + 1}').color = selected_color
                print(f'{start_range}{id + 1}:{end_range}{id + 1}')


def add_missing_line(ws, difference, data, selected_color):
    """
    Добавляет сервисные строки с удаленными разрешениями и окрашивает их в красный цвет
    :param ws: рабочий лист
    :param difference: дельта изменений в данных
    :param data: данные рабочего листа
    :param selected_color: цвет в который красим (красный)
    """
    line_for_write = len(data) + 4
    for str_data in data:
        for str_diff in difference:
            if [str_diff[0], str_diff[1]] not in list([element[0], element[1]] for element in data):
                start_range = number_to_letter(1)
                end_range = number_to_letter(len(str_diff))
                ws.range(f'{start_range}{line_for_write}:{end_range}{line_for_write}').value = str_diff
                ws.range(f'{start_range}{line_for_write}:{end_range}{line_for_write}').font.color = r'#780000'
                ws.range(f'{start_range}{line_for_write}:{end_range}{line_for_write}').color = selected_color
                difference.remove(str_diff)
                line_for_write += 1
            elif str_data[0] == str_diff[0] and str_data[1] == str_diff[1] and len(str_diff) > len(str_data):
                start_range = number_to_letter(1)
                end_range = number_to_letter(len(str_diff))
                ws.range(f'{start_range}{line_for_write}:{end_range}{line_for_write}').value = str_diff
                ws.range(f'{start_range}{line_for_write}:{end_range}{line_for_write}').font.color = r'#780000'
                start_range = number_to_letter(1 + len(str_data))
                end_range = number_to_letter(len(str_data) + len(str_diff) - len(str_data))
                ws.range(f'{start_range}{line_for_write}:{end_range}{line_for_write}').color = selected_color
                line_for_write += 1


def filter_missing_line(data):
    """удаляет все сервисные строки с удаленными разрешениями"""
    return list(item for item in data if item[0] != None)


def del_none(data):
    """удаляет все значения None из списка"""
    if data:
        if isinstance(data[0], list):
            return list(list(point for point in item if point != None) for item in data)
        else:
            return list(item for item in data if item != None)
    else:
        return data



def get_all_path():
    """читает пути для работы из файла"""
    if os.path.exists(f'{os.getcwd()}\\config.ini'):
        all_path = {}
        with open(f'{os.getcwd()}\\config.ini', 'r', encoding='cp1251') as file:
            for id, string in enumerate(file):
                if id == 1:
                    all_path['dir'] = string.strip().lower()
                elif id == 3:
                    all_path['file'] = string.strip().lower()
        return all_path
    else:
        return {'file': f'{os.getcwd()}\\access_list.xlsx', 'dir': f'{os.getcwd()}'}


if __name__ == '__main__':
    path = get_all_path()
    # заносим в переменную путь к файлу excel
    file_path = path['file']
    # заносим в переменную путь к каталогу права на папки которого нужно получить
    dir_path = path['dir']
    # заносим в переменную перечень используемых цветов заливки
    color = {'green': r'#CCFFCC', 'red': r'#FF9999', 'white': None}

    # открываем рабочую книгу (wb)
    try:
        wb = xw.Book(file_path)
        # заносим в переменную имя рабочего листа (ws)
        ws_now_name = ws_now(wb)
        # используем функцию добавления листа
        wb.sheets.add(ws_now_name)
    except:
        # если не удалось открыть, создаем книгу
        wb = xw.Book()
        # заносим в переменную имя рабочего листа (ws)
        ws_now_name = datetime.datetime.today().strftime('%d-%m-%Y')
        # переименовываем активный лист
        xw.books.active.sheets.active.name = ws_now_name

    # инициализируем рабочий лист (ws)
    ws_now = wb.sheets[ws_now_name]
    # получаем список разрешений и записываем их на лист файла
    data_ws_now = directories_access(dir_path) #актуальный список разрешений
    ws_now.range('A1').value = data_ws_now

    if len(list(sheet.name for sheet in wb.sheets)) > 1:
        # инициализируем 'предыдущий' лист (ws)
        ws_previous = wb.sheets[1]
        # заносим в переменную 'data_ws_previous' содержимое листа ws_previous
        data_ws_previous = ws_previous.range('A1').expand().value

        if data_ws_now and data_ws_previous:
            """дальше по циклу идем только если оба листа с данными"""
            # если на листе только одна строка, правим состав переменной и упаковываем список строки с список листа
            if not isinstance(data_ws_previous[0], list):
                data_ws_previous = [data_ws_previous]
            if not isinstance(data_ws_now[0], list):
                data_ws_now = [data_ws_now]

            # удаляем из переменной сервисную информацию об отсутствубщих разрешениях
            data_ws_previous = filter_missing_line(data_ws_previous)
            # вычитаем содержимое текущего листа из предыдущего и удаляем все значения None из списков разрешений
            difference_np = difference_lists(data_ws_now, data_ws_previous)
            difference_np = del_none(difference_np)
            # с помощью функции окрашиваем в зеленый цвет "появившиеся" разрешения
            coloring_on_difference(ws_now, difference_np, data_ws_now, color['green'])
            # вычитаем содержимое предыдущего листа из текущего и удаляем все значения None из списков разрешений
            difference_pn = difference_lists(data_ws_previous, data_ws_now)
            difference_pn = del_none(difference_pn)
            # удаляем все значения None из списка разрешений
            data_ws_now = del_none(data_ws_now)
            # с помощью функции удаляем избыточную окрашенность списка разрешений
            coloring_on_difference(ws_now, difference_pn, data_ws_now, color['white'])
            # с помощью функции добавляем в конец листа список удаленных разрешений
            add_missing_line(ws_now, difference_pn, data_ws_now, color['red'])

    # сохраняем книгу
    wb.save(file_path)
    # закрываем процесс
    wb.app.quit()
    # с помощью функции выставляем оптимальное значение ширины стролбцов
    auto_size_column(file_path, ws_now_name)
