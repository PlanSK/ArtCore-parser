import gspread
from parser import number_cell_processing


def analyze_row(number, card, fname, name, oname, phone_number, rname, roname, *args) -> dict():
    if fname:
        if name and oname:
            fio = ' '.join([fname, name, oname])
        elif not name and not oname and rname and roname:
            fio = ' '.join([fname, rname, roname])
        elif name and not oname:
            fio = ' '.join([fname, name])

        if name and not phone_number and card:
            phone_number = card

        if phone_number and number_cell_processing('GSheets', number, phone_number):
            return {
                number_cell_processing('GSheets', number, phone_number): {
                    'name': fio,
                    'balance': float(0),
                    'total_costs': float(0)
                }
            }

    return dict()


gc = gspread.service_account(filename='et_creds.json')

sh = gc.open_by_key("1EsL801iyUvi7TtcHOubsnKyYt37VgeCoVPDlcVNi2HA")

list_of_rows = sh.sheet1.get_all_values()

for num, row in enumerate(list_of_rows):
    params = [cell.strip() for cell in [num] + row[4:9] + row[15:17]]
    number = analyze_row(*params)

    print(f"{num}: {number}")