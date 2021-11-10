import gspread
import gspread_formatting
import parser

from gspread_formatting.models import TextFormat


GSHEETS_API_KEY = 'esports_cred.json'


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

        if phone_number and parser.number_cell_processing(
            'GSheets',
            number,
            phone_number
        ):
            return {
                parser.number_cell_processing(
                    'GSheets',
                    number,
                    phone_number): {
                    'name': fio,
                    'balance': float(0),
                    'total_costs': float(0)
                }
            }

    return dict()


def gsheets_load(table_key: str) -> dict:
    google_connect = gspread.service_account(filename=GSHEETS_API_KEY)
    gsheet = google_connect.open_by_key(table_key)

    list_of_rows = gsheet.sheet1.get_all_values()
    base = dict()


    for num, row in enumerate(list_of_rows):
        params = [cell.strip() for cell in [str(num)] + row[4:9] + row[15:17]]
        if analyze_row(*params):
            base.update(analyze_row(*params))
    
    return base


def gsheets_save(table_key: str, numbers: dict):
    google_connect = gspread.service_account(filename=GSHEETS_API_KEY)
    gsheet = google_connect.open_by_key(table_key)
    worksheet = gsheet.sheet1
    worksheet.clear()
    write_list = []

    float_style = {
        'numberFormat': {
            'type': 'NUMBER',
            'pattern': '#,##0.00'
        }
    }
    
    borders_style = gspread_formatting.Border(
        style='SOLID',
        color=gspread_formatting.Color(0, 0, 0),
        width=1
    )
    table_style = gspread_formatting.CellFormat(
        borders=gspread_formatting.Borders(
            top=borders_style,
            bottom=borders_style,
            left=borders_style,
            right=borders_style
        )
    )
    title_format = gspread_formatting.CellFormat(
        backgroundColor=gspread_formatting.color(1, 0, 0),
        horizontalAlignment='CENTER',
        textFormat=TextFormat(
            foregroundColor=gspread_formatting.color(1, 1, 1),
            bold=True,
            fontSize=10
        )
    )
    
    worksheet.update('A1', [['N', 'Phone number', 'Name', 'Balance', 'Total costs']])
    
    for number, (phone, data) in enumerate(numbers.items(), start=1):
        write_list.append([
            number,
            phone,
            data['name'],
            data['balance'],
            data['total_costs']
        ])
    gspread_formatting.set_column_widths(worksheet, [ 
        ('A:', 50),
        ('B:', 100),
        ('C:', 300)
    ])
    
    range_table = 'E'+str(len(write_list) + 1)
    
    worksheet.update('A2:' + range_table, write_list)
    gspread_formatting.format_cell_ranges(
        worksheet,
        [('A1:'+ range_table, table_style), ('A1:E1', title_format)]
    )
    worksheet.format('D2:' + range_table, float_style)
