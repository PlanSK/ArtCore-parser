from datetime import datetime
import os
import json

import openpyxl
from loguru import logger as log

import gsheets
import excel
import parser
import hash


if __name__ == '__main__':
    balance_bonus = 0.0
    base_table_key = '1EsL801iyUvi7TtcHOubsnKyYt37VgeCoVPDlcVNi2HA'
    upload_table_key = '1n5A-fkR0LYCdTbdiBVLcRZ4-AKjEoetCE2bgGJEIym0'
    checksum_file = 'checksum.lst'
    base_file_name = 'base.json'
    excel_file_name = 'base.xlsx'
    start_time = datetime.now()
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
    checksum_list = hash.checksum_list(checksum_file)

    path = os.path.join(os.getcwd(), 'bases')
    files = [
        os.path.join(path, get_file) 
        for get_file in os.listdir(path) 
        if (os.path.isfile(os.path.join(path, get_file))
                and get_file[0].isalpha() and 'xls' in get_file.split('.')[1])
    ]

    default_checksum = ''

    if os.path.exists(base_file_name):
        with open(base_file_name, 'r', encoding='utf-8') as base_data:
            numbers_base = json.load(base_data)

        default_checksum = hash.checksum_dict(numbers_base)
        gsheets_base = gsheets.gsheets_load(base_table_key)

        if checksum_list.get(base_table_key) and checksum_list[base_table_key] != hash.checksum_dict(gsheets_base):
            checksum_list.update({base_table_key: hash.checksum_dict(gsheets_base)})
            for get_number, returned_dict in gsheets_base.items():
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
        else:
            print('Hash gsheets not changed. Skipped.')
    else:
        numbers_base = gsheets.gsheets_load(base_table_key)

    for file in files:
        if not hash.checksum_check(file, checksum_list) or not default_checksum:
            wbook = openpyxl.load_workbook(file)
            sheet = wbook.active

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

            path, file_name = os.path.split(file)
            checksum_list.update({file_name: hash.checksum_gen(file)})
        else:
            print(f'{file} is skiped. Checksum not changed.')

    with open(checksum_file, 'w', encoding='utf-8') as new_checksum_file:
        json.dump(checksum_list, new_checksum_file, indent=4)

    if default_checksum != hash.checksum_dict(numbers_base):
        if balance_bonus:
            print(f'Add balance bonus {balance_bonus} RUB.')
            for number in numbers_base.keys():
                numbers_base[number]['balance'] += balance_bonus

        print(f'Total found: {len(numbers_base.keys())} records.')
        if excel_file_name:
            excel.data_save(numbers_base, excel_file_name)

        print('Saving to Google sheets...')
        gsheets.gsheets_save(upload_table_key, numbers_base)

        print('Saving json file...')
        with open(base_file_name, 'w', encoding='utf-8') as json_file:
            json.dump(numbers_base, json_file, indent=4,)
    
    print(f'{len(numbers_base.keys())} records in base.')
    print(f"Time work: {datetime.now() - start_time}")