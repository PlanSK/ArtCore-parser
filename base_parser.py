import openpyxl
from loguru import logger as log


def num_extractor(get_row: str) -> int:
    if get_row[3]:
        get_value_cell = get_row[3]
    else:
#        log.info(f"None phone number in {get_row[0]} row.")
        return 0

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
        log.info(f"Wrong value in {get_row[0]} row. Value '{get_value_cell}'")
    else:
        for get_num in number_list:
            if len(get_num) == 11:
                if 7 > int(get_num[0]) > 8:
                    continue
                else:
                    return int(get_num)
            elif len(get_num) == 10 and get_num[0] == 9:
                return int('8' + get_num)
        log.info(f"Short or long values in {get_row[0]} row. Value ({len(get_value_cell)})'{get_value_cell}'")
    return 0


file = 'yamash.xlsx'
log.add("error.log", format="{level} | {message}", mode='w')

wbook = openpyxl.load_workbook(file)
sheet = wbook.active
print(f'Total records: {sheet.max_row}')

numbers_list = list()

for index in range(1, sheet.max_row + 1):
    cells = [
        str(sheet.cell(row=index, column=get_column_num).value)
        if sheet.cell(row=index, column=get_column_num).value else ''
        for get_column_num in range(1, sheet.max_column + 1)
    ]
    if num_extractor(cells):
        numbers_list.append(cells[3])

print(f'Total found: {len(numbers_list)}')
