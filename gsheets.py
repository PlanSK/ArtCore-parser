import gspread
from . import parser


def analyze_row(
    number: str,
    card: str,
    fname: str,
    name: str,
    oname: str,
    phone_number: str,
    rname: str,
    roname: str,
    *args
) -> dict():
    if fname:
        fio = fname
        
        if name:
            fio += ' ' + name
        elif not name and rname:
            fio += ' ' + rname
        
        if oname:
            fio += ' ' + oname
        elif not oname and roname:
            fio += ' ' + roname

        if name and not phone_number and card:
            phone_number = card

        if phone_number and parser.number_cell_processing('GSheets', number, phone_number):
            return {
                parser.number_cell_processing('GSheets', number, phone_number): {
                    'name': fio,
                    'balance': float(0),
                    'total_costs': float(0)
                }
            }

    return dict()

def gsheets_load() -> dict:
    gc = gspread.service_account(filename='et_creds.json')

    sh = gc.open_by_key("1EsL801iyUvi7TtcHOubsnKyYt37VgeCoVPDlcVNi2HA")

    list_of_rows = sh.sheet1.get_all_values()
    base = dict()


    for num, row in enumerate(list_of_rows):
        params = [cell.strip() for cell in [str(num)] + row[4:9] + row[15:17]]
        if analyze_row(*params):
            base.update(analyze_row(*params))
    
    return base