import openpyxl
from loguru import logger as log
import random


def number_cell_processing(get_str_number, get_value_cell: str) -> int:
    num_string = ''
    number_list = list()

    for fragment in get_value_cell.split():
        for char in fragment:
            if char.isdigit():
                num_string += char
        if num_string:
            number_list.append(num_string)
            num_string = ''

    if not len(number_list):
        log.info(f"Wrong value in {get_str_number} row. Value '{get_value_cell}'")
    else:
        for get_num in number_list:
            if len(get_num) == 11:
                if 7 > int(get_num[0]) > 8:
                    continue
                else:
                    if int(get_num[0]) == 8:
                        get_num = '7' + get_num[1:]
                    return get_num
            elif len(get_num) == 10 and get_num[0] == 9:
                return '7' + get_num
        log.info(f"Short or long values in {get_str_number} row. Value ({len(get_value_cell)})'{get_value_cell}'")

    return 0


def data_exctraction(get_row: list) -> dict():
    data_dict = dict()

    if get_row[0] and get_row[3]:
        get_number = number_cell_processing(get_row[0], get_row[3])
        if get_number:
            data_dict['name'] = get_row[2]

            balance = 0.0
            if get_row[4]:
                balance = float(get_row[4].replace(',', '.'))
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
    writebook = openpyxl.Workbook()
    active_sheet = writebook.active
    active_sheet.cell(row=1, column=1).value = 'N п/п'
    active_sheet.cell(row=1, column=2).value = 'Phone number'
    active_sheet.cell(row=1, column=3).value = 'User name'
    active_sheet.cell(row=1, column=4).value = 'Balance'
    active_sheet.cell(row=1, column=5).value = 'Total costs'

    for index, get_number in enumerate(numbers_dict.keys()):
        active_sheet.cell(row=index + 2, column=1).value = index + 1
        active_sheet.cell(row=index + 2, column=2).value = get_number
        active_sheet.cell(row=index + 2, column=3).value = numbers_dict[get_number]['name']
        active_sheet.cell(row=index + 2, column=4).value = numbers_dict[get_number]['balance']
        active_sheet.cell(row=index + 2, column=5).value = numbers_dict[get_number]['total_costs']
    try:
        writebook.save('base.xlsx')
    except PermissionError:
        writebook.save('base'+str(random.randint(10000,99999))+'.xlsx')


if __name__ == '__main__':
    file = 'yamash.xlsx'
    log.add("error.log", format="{level} | {message}", mode='w')

    wbook = openpyxl.load_workbook(file)
    sheet = wbook.active
    print(f'Total records: {sheet.max_row}')

    numbers_base = dict()

    for index in range(1, sheet.max_row + 1):
        if sheet.cell(row=index, column=4).value:
            cells = [
                str(sheet.cell(row=index, column=get_column_num).value)
                if sheet.cell(row=index, column=get_column_num).value else ''
                for get_column_num in range(1, sheet.max_column + 1)
            ]

            returned_dict = data_exctraction(cells)

            if returned_dict.values():
                if not numbers_base.get(len(returned_dict.keys())):
                    numbers_base.update(returned_dict)
                else:
                    log.info(f"{returned_dict.keys()} already exists in base.")

    data_save(numbers_base)

    print(f'Total found: {len(numbers_base.keys())}')
