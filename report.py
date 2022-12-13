import openpyxl


def main():
    # ex - документ, куда будет записываться информация
    name_ex = 'example.xlsx'
    book_ex = openpyxl.load_workbook(filename=name_ex)
    sheet_ex = book_ex.worksheets[0]

    sheet_ex_info = book_ex['info']
    month = sheet_ex_info['A2'].value
    file_zp = sheet_ex_info['B2'].value
    file_on = sheet_ex_info['C2'].value
    # запись файла происходит дальше


    # zp - документ, который ведёт менеджер
    book_zp = openpyxl.load_workbook(filename=file_zp, read_only=True, data_only=True)
    sheet_zp = book_zp.worksheets[-1]  # последняя страница файла из списка - 2022
    func_zp = (lambda row: row[7] == month)
    data_zp = sheet_to_iter(sheet_zp, func_zp, exc=(0,))

    # on - документ, который формируется 1С
    book_on = openpyxl.load_workbook(filename=file_on, read_only=True, data_only=True)
    sheet_on = book_on.worksheets[0]
    func_on = (lambda row: all(c is None for c in row[1:7]))
    data_on = sheet_to_iter(sheet_on, func_on, exc=range(6))



    # запись данных в файл
    clear_sheet(book_ex, sheet_ex, name_ex)
    sheet_ex['A1'] = 'Контрагент'
    sheet_ex['B1'] = '''№ счета
№ договора
№ спецификации'''
    sheet_ex['C1'] = '''Дата выставления
счета / договора / спец-и'''
    sheet_ex['D1'] = 'Дата отгрузки'
    sheet_ex['E1'] = 'Сумма с НДС, р.'
    sheet_ex['F1'] = 'Дата оплаты факт'
    sheet_ex['G1'] = 'Маржа, р.'

    row_ex = 2 # потому что первую строку заполнил вручную
    for row_zp in data_zp:
        for row_on in data_on:
            if row_zp[1] == row_on[0] and row_zp[4] == row_on[13]:
                for col_ex, value in enumerate(list(row_zp[:6]) + [row_on[14]], 1):
                    cell = sheet_ex.cell(row=row_ex, column=col_ex)
                    cell.value = value
                row_ex += 1
    book_ex.save(name_ex)
    print('Success.')


def sheet_to_iter(sheet, condition_selection, exc=(0,)):
    data = tuple(filter(condition_selection,
                        (
                            tuple(map(lambda x: x.value, row)) for i, row in enumerate(sheet) if not (i in exc)
                        )
                        ))
    return data


def clear_sheet(book, sheet, name):
    sheet.delete_cols(0, 100)
    book.save(name)


main()
