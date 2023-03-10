import openpyxl
import time


def load_file(file_name, read_only=True, data_only=False):
    """
    :param file_name: имя файла
    :param read_only: файл открывается на чтение?
    :param data_only: определяет, будут ли содержать ячейки с формулами - формулу (по умолчанию) или только значение,
    сохраненное/посчитанное при последнем чтении листа Excel.
    :return: загруженный документ
    """
    try:
        return openpyxl.load_workbook(filename=file_name, read_only=read_only, data_only=data_only)
    except FileNotFoundError:
        print(f'Отсутствует файл - {file_name}. Исправьте это и запустите программу снова :)')
        time.sleep(10)
        raise FileNotFoundError('Добавьте файл.')


def sheet_to_iter(sheet, condition_selection, exc=(0,)) -> tuple:
    """
    :param sheet: имя вкладки в exel документе
    :param condition_selection: условие отбора
    :param exc: строки, которые исключаются из выборки
    :return: кортеж с данными
    """
    data = tuple(
        tuple(map(lambda x: x.value, row)) for i, row in enumerate(sheet) if (
            (i not in exc) and condition_selection(row)
        )
    )
    return data


def clear_sheet(book, sheet, name):
    try:
        sheet.delete_cols(0, 100)
        book.save(name)
    except PermissionError:
        print('Закройте файл Report.xlsx. После этого запустите программу снова :)')
        time.sleep(10)
        raise PermissionError('Перезапустите файл после устранения замечаний.')


def main():
    print("Welcome. Please don't distract me. I'm working...\n")
    time.sleep(1)

    # info - документ, где содержится информация о интересущем месяце, годе и названием файлов
    info_file_name = '!info.xlsx'
    info_book = load_file(file_name=info_file_name)
    info_sheet = info_book.worksheets[0]
    month = info_sheet['A2'].value
    zap_sheet_name = str(info_sheet['B2'].value)  # название вкладки = равно году в котором ведётся учёт
    zap_file_name = info_sheet['C2'].value
    data_file_name = info_sheet['D2'].value
    report_file_name = info_sheet['E2'].value

    # zap - документ, который ведёт менеджер
    zap_book = load_file(file_name=zap_file_name)
    # zap_sheet = zap_book.worksheets[-1]  # последняя страница файла из списка - 2022
    # zap_sheet = zap_book.worksheets[name_sheet] # если обращаться по индексу
    # zap_sheet = zap_book.active # выберет активный лист
    zap_sheet = zap_book[zap_sheet_name]
    zap_func = (lambda row: row[7].value == month)
    zap_data = sheet_to_iter(zap_sheet, zap_func, exc=(0,))

    # data - документ, который формируется 1С
    data_book = load_file(file_name=data_file_name)
    data_sheet = data_book.worksheets[0]
    data_func = (lambda row: all(c.value is None for c in row[1:7]))
    data_data = sheet_to_iter(data_sheet, data_func, exc=range(6))
    # print(*data_data, sep='\n')

    # запись данных в файл report
    report_book = load_file(file_name=report_file_name, read_only=False)
    report_sheet = report_book.worksheets[0]  # вкладка в документе куда будет записывать информация
    clear_sheet(report_book, report_sheet, report_file_name)
    report_sheet['A1'] = 'Контрагент'  # 0
    report_sheet['B1'] = '№ счета\n№ договора\n№ спецификации'  # 1
    report_sheet['C1'] = 'Дата выставления\nсчета / договора / спец-и'  # 2
    report_sheet['D1'] = 'Дата отгрузки'  # 3
    report_sheet['E1'] = 'Сумма с НДС, р.'  # 4
    report_sheet['F1'] = 'Дата оплаты факт'  # 5
    report_sheet['G1'] = 'Маржа, р.'  # 6
    report_row = 2  # первую строку заполнил вручную, поэтому начинаю со 2 строки

    flag = False
    for zap_row in zap_data:
        for data_row in data_data:
            if zap_row[1] == data_row[0] and zap_row[4] == data_row[14]:
                data = list(zap_row[:6])
                data.append(data_row[15])
                data[3] = data_row[7]

                for report_col, value in enumerate(data, start=1):  # запись в файл
                    cell = report_sheet.cell(row=report_row, column=report_col)
                    if report_col in (3, 4, 6):
                        cell.number_format = 'DD.MM.YYYY'
                    elif report_col in (5, 7):
                        cell.number_format = '# ##0.00 [$₽-419]'
                    cell.value = value

                report_row += 1
                flag = True

    report_book.save(report_file_name)

    if not flag:
        print('WARNING!!! None of the lines are written. Check the input data.')
    else:
        print('AWESOME!!! The report file was successfully recorded.\nHave a good day!')
    time.sleep(10)


if __name__ == '__main__':
    main()
