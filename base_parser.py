import openpyxl
from loguru import logger as log
import random
import re
import os


def number_cell_processing(file: str, get_str_number, get_value_cell: str) -> int:
    short = False
    long = False
    if re.findall(r'([3789]){1}\d{1,}', get_value_cell):
        for s in re.finditer(r'([3789]){1}\d{1,}', get_value_cell):
            get_match = s.group()
            if get_match[0] == '9' and len(get_match) == 10:
                return '+7' + get_match
            elif get_match[0] != '9' and len(get_match) == 11:
                return '+7' + get_match[1:]
            elif len(get_match) < 11:
                short = True
            elif len(get_match) > 11:
                long = True

        get_match = next(re.finditer(r'([3789]){1}\d{1,}', get_value_cell)).group()
        if short:
            log.bind(short=True).info(f"({file}) Short number value in {get_str_number} row. Value ({len(get_match)}) '{get_value_cell}'")
            return 0
        elif long:
            log.bind(long=True).info(f"({file}) Long number value in {get_str_number} row. Value ({len(get_match)}) '{get_value_cell}'")
            return 0
    else:
        log.bind(wrong=True).info(f"({file}) Wrong value in {get_str_number} row. Value '{get_value_cell}'")


def data_exctraction(file: str, get_row: list) -> dict():
    data_dict = dict()

    if get_row[0] and get_row[3]:
        get_number = number_cell_processing(file, get_row[0], get_row[3])
        if get_number:
            data_dict['name'] = get_row[2]

            balance = 0.0
            if get_row[4]:
                try: 
                    balance = float(get_row[4].replace(',', '.'))
                except ValueError:
                    log.bind(wrong=True).error(f"({file}) Error in balance in {get_row[0]}. Value: {get_row[4]}. Skipped.")
            data_dict['balance'] = balance

            total_costs = 0.0
            if get_row[6]:
                try:
                    total_costs = float(get_row[6].split()[0].replace(',', '.'))
                except ValueError:
                    log.bind(wrong=True).error(f"({file}) Error in total costs in {get_row[0]}. Value: {get_row[6]}. Skipped.")
            data_dict['total_costs'] = total_costs

            return {get_number: data_dict}

    return dict()


def data_save(numbers_dict: dict) -> None:
    boder_style = openpyxl.styles.borders.Border(
        left=openpyxl.styles.borders.Side(style='thin'), 
        right=openpyxl.styles.borders.Side(style='thin'), 
        top=openpyxl.styles.borders.Side(style='thin'), 
        bottom=openpyxl.styles.borders.Side(style='thin')
    )
    writebook = openpyxl.Workbook()
    active_sheet = writebook.active
    active_sheet.title = 'User base'
    active_sheet.cell(row=1, column=1).value = 'N п/п'
    active_sheet.cell(row=1, column=2).value = 'Phone number'
    active_sheet.cell(row=1, column=3).value = 'User name'
    active_sheet.cell(row=1, column=4).value = 'Balance'
    active_sheet.cell(row=1, column=5).value = 'Total costs'
    active_sheet.column_dimensions["A"].width = 10
    active_sheet.column_dimensions["B"].width = 15
    active_sheet.column_dimensions["C"].width = 60
    active_sheet.column_dimensions["D"].width = 15
    active_sheet.column_dimensions["E"].width = 15

    for index, get_number in enumerate(numbers_dict.keys()):
        active_sheet.cell(row=index + 2, column=1).value = index + 1
        active_sheet.cell(row=index + 2, column=2).value = get_number
        active_sheet.cell(row=index + 2, column=3).value = numbers_dict[get_number]['name']
        active_sheet.cell(row=index + 2, column=4).value = numbers_dict[get_number]['balance']
        active_sheet.cell(row=index + 2, column=4).number_format = '0.00'
        active_sheet.cell(row=index + 2, column=5).value = numbers_dict[get_number]['total_costs']
        active_sheet.cell(row=index + 2, column=5).number_format = '0.00'

    for row in active_sheet.iter_rows(min_row=1, max_col=active_sheet.max_column, max_row=active_sheet.max_row):
        for cell in row:
            cell.border = boder_style

    print('Saving file...')
    try:
        writebook.save('base.xlsx')
    except PermissionError:
        writebook.save('base'+str(random.randint(10000,99999))+'.xlsx')


if __name__ == '__main__':
    path = os.path.join(os.getcwd(), 'bases')
    files = [
        os.path.join(path, get_file) 
        for get_file in os.listdir(path) 
        if (os.path.isfile(os.path.join(path, get_file))
                and get_file[0].isalpha() and 'xls' in get_file.split('.')[1])
    ]
    log.add("long.log", filter=lambda record: "long" in record["extra"], mode='w')
    log.add("short.log", filter=lambda record: "short" in record["extra"], mode='w')
    log.add("error.log", filter=lambda record: "wrong" in record["extra"], mode='w')
    numbers_base = dict()

    for file in files:
        wbook = openpyxl.load_workbook(file)
        sheet = wbook.active
        print(f'Total records: {sheet.max_row}')

        for index in range(1, sheet.max_row + 1):
            if sheet.cell(row=index, column=4).value:
                cells = [
                    str(sheet.cell(row=index, column=get_column_num).value)
                    if sheet.cell(row=index, column=get_column_num).value else ''
                    for get_column_num in range(1, sheet.max_column + 1)
                ]

                returned_dict = data_exctraction(file, cells)

                if returned_dict.keys():
                    get_number = list(returned_dict.keys())[0]
                    if not numbers_base.get(get_number):
                        numbers_base.update(returned_dict)
                    else:
                        numbers_base[get_number]['balance'] = max(
                            returned_dict[get_number]['balance'],
                            numbers_base[get_number]['balance']
                        )
                        numbers_base[get_number]['total_costs'] = max(
                            returned_dict[get_number]['total_costs'],
                            numbers_base[get_number]['total_costs']
                        )

    print(f'Total found: {len(numbers_base.keys())}')
    data_save(numbers_base)
