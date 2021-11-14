import re
from loguru import logger as log 


def number_cell_processing(
    file: str, 
    get_str_number: str, 
    get_value_cell: str
) -> int:

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
            log.bind(short=True).info(
                f"({file}) Short number value in {get_str_number} row. "
                f"Value ({len(get_match)}) '{get_value_cell}'"
            )
            return 0
        elif long:
            log.bind(long=True).info(
                f"({file}) Long number value in {get_str_number} row. "
                f"Value ({len(get_match)}) '{get_value_cell}'"
            )
            return 0
    else:
        log.bind(wrong=True).info(
            f"({file}) Wrong value in {get_str_number} row. "
            f"Value '{get_value_cell}'"
        )


def data_exctraction(
    file: str,
    order_number: str,
    _,
    personal_name: str,
    phone_number: str,
    get_balance: str,
    __,
    costs: str,
    *remaining
) -> dict():

    data_dict = dict()

    if order_number and phone_number:
        get_number = number_cell_processing(file, order_number, phone_number)
        if get_number:
            data_dict['name'] = re.sub(r'[^А-Яа-яёЁ\s]', '', personal_name).strip()

            balance = 0.0
            if get_balance:
                try: 
                    balance = float(get_balance.replace(',', '.'))
                except ValueError:
                    log.bind(wrong=True).error(
                        f"({file}) Error in balance in {order_number}. "
                        f"Value: {get_balance}. Skipped."
                    )
            data_dict['balance'] = balance

            total_costs = 0.0
            if costs:
                try:
                    total_costs = float(costs.split()[0].replace(',', '.'))
                except ValueError:
                    log.bind(wrong=True).error(
                        f"({file}) Error in total costs in {order_number}. "
                        f"Value: {costs}. Skipped."
                    )
            data_dict['total_costs'] = total_costs
            data_dict['file'] = [f'{file}:{order_number}']

            return get_number, data_dict

    return None, None
