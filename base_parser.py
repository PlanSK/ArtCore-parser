import openpyxl
from loguru import logger as log
import random
import re
import os


def number_cell_processing(get_str_number, get_value_cell: str) -> int:
    number = 0

    if re.match(r'.*[7-8]{1}\d{10}', get_value_cell):
        if len(re.findall(r'[7-8]{1}\d{10,}', get_value_cell)[0]) == 11:
            number = '+7' + re.findall(r'[7-8]{1}\d{10}', get_value_cell)[0][1:]
        elif len(re.findall(r'[7-8]{1}\d{10,}', get_value_cell)[0]) > 11:
            leng = len(re.findall(r'[7-8]{1}\d{10,}', get_value_cell)[0])
            log.info(f"Long number value in {get_str_number} row. Value ({leng}) '{get_value_cell}'")
    if re.match(r'.*[3]{1}\d{10}', get_value_cell):
        if len(re.findall(r'[3]{1}\d{10,}', get_value_cell)[0]) == 11:
            number = '+7' + re.findall(r'[3]{1}\d{10}', get_value_cell)[0][1:]
        elif len(re.findall(r'[3]{1}\d{10,}', get_value_cell)[0]) > 11:
            leng = len(re.findall(r'[3]{1}\d{10,}', get_value_cell)[0])
            log.info(f"Long number value in {get_str_number} row. Value ({leng}) '{get_value_cell}'")
    elif re.match(r'.*[9]{1}\d{9}', get_value_cell) and len(re.findall(r'[9]{1}\d{9,}', get_value_cell)[0]) == 10:
        number = '+7' + re.findall(r'[9]{1}\d{9,}', get_value_cell)[0]
    elif re.match(r'.*[7-8]{1}\d{,9}', get_value_cell):
        leng = len(re.findall(r'[7-8]{1}\d{,9}', get_value_cell)[0])
        log.info(f"Short number value in {get_str_number} row. Value ({leng}) '{get_value_cell}'")
    else:
        log.info(f"Wrong value in {get_str_number} row. Value '{get_value_cell}'")

    return number


def data_exctraction(get_row: list) -> dict():
    data_dict = dict()

    if get_row[0] and get_row[3]:
        get_number = number_cell_processing(get_row[0], get_row[3])
        if get_number:
            data_dict['name'] = get_row[2]

            balance = 0.0
            if get_row[4]:
                try: 
                    balance = float(get_row[4].replace(',', '.'))
                except ValueError:
                    log.error(f"Error in balance in {get_row[0]}. Value: {get_row[4]}. Skipped.")
            data_dict['balance'] = balance

            total_costs = 0.0
            if get_row[6]:
                try:
                    total_costs = float(get_row[6].split()[0].replace(',', '.'))
                except ValueError:
                    log.error(f"Error in total costs in {get_row[0]}. Value: {get_row[6]}. Skipped.")
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
    files = [os.path.join(path, get_file) for get_file in os.listdir(path) if os.path.isfile(os.path.join(path, get_file))]
    log.add("error.log", format="{level} | {message}", mode='w')
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

                returned_dict = data_exctraction(cells)

                if returned_dict.keys():
                    for key in returned_dict.keys():
                        get_number = key
                    if not numbers_base.get(get_number):
                        numbers_base.update(returned_dict)
                    else:
                        numbers_base[get_number]['balance'] = max(returned_dict[get_number]['balance'], numbers_base[get_number]['balance'])
                        numbers_base[get_number]['total_costs'] = max(returned_dict[get_number]['total_costs'], numbers_base[get_number]['total_costs'])
                        log.info(f"{get_number} already exists in base. New data - B: {numbers_base[get_number]['balance']} TC: {numbers_base[get_number]['total_costs']}")

    print(f'Total found: {len(numbers_base.keys())}')
    data_save(numbers_base)
