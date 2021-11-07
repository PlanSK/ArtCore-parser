from datetime import date, datetime
from loguru import logger as log
import os


import openpyxl

import gsheets
import excel
import parser


if __name__ == '__main__':
    balance_bonus = 300.0
    base_table_key = '1EsL801iyUvi7TtcHOubsnKyYt37VgeCoVPDlcVNi2HA'
    upload_table_key = '1n5A-fkR0LYCdTbdiBVLcRZ4-AKjEoetCE2bgGJEIym0'
    start_time = datetime.now()
    path = os.path.join(os.getcwd(), 'bases')
    files = [
        os.path.join(path, get_file) 
        for get_file in os.listdir(path) 
        if (os.path.isfile(os.path.join(path, get_file))
                and get_file[0].isalpha() and 'xls' in get_file.split('.')[1])
    ]
    
    log.add(
        "long.log",
        filter=lambda record: "long" in record["extra"], mode='w'
    )
    log.add(
        "short.log",
        filter=lambda record: "short" in record["extra"], mode='w'
    )
    log.add(
        "error.log",
        filter=lambda record: "wrong" in record["extra"], mode='w'
    )
    
    numbers_base = gsheets.gsheets_load(base_table_key)

    for file in files:
        wbook = openpyxl.load_workbook(file)
        sheet = wbook.active
        print(f'Total records: {sheet.max_row}')

        for row in sheet.iter_rows(
            min_row=1, 
            max_row=sheet.max_row, 
            max_col=sheet.max_column
        ):

            get_row = [cell.value for cell in row]
            if get_row[3]:
                cells = [
                    str(get_cell)
                    if get_cell else ''
                    for get_cell in get_row
                ]

                get_number, returned_dict = parser.data_exctraction(file, *cells)

                if get_number:
                    if not numbers_base.get(get_number):
                        numbers_base[get_number] = returned_dict
                    else:
                        numbers_base[get_number]['balance'] = max(
                            returned_dict['balance'],
                            numbers_base[get_number]['balance']
                        )
                        numbers_base[get_number]['total_costs'] = max(
                            returned_dict['total_costs'],
                            numbers_base[get_number]['total_costs']
                        )

    if balance_bonus:
        print(f'Add balance bonus {balance_bonus} RUB.')
        for number in numbers_base.keys():
            numbers_base[number]['balance'] += balance_bonus

    print(f'Total found: {len(numbers_base.keys())} records.')
    excel.data_save(numbers_base)
    print('Saving to Google sheets...')
    gsheets.gsheets_save(upload_table_key, numbers_base)
    print(f"Time work: {datetime.now() - start_time}")