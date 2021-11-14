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
                    'total_costs': float(0),
                    'file': ['gsheets']
                }
            }

    return dict()


def gsheets_load_data(table_key: str) -> list:
    google_connect = gspread.service_account(filename=GSHEETS_API_KEY)
    gsheet = google_connect.open_by_key(table_key)

    return gsheet.sheet1.get_all_values()

def gsheets_form_load(table_key: str) -> dict:
    list_of_rows = gsheets_load_data(table_key)
    base = dict()

    for num, row in enumerate(list_of_rows):
        params = [cell.strip() for cell in [str(num)] + row[4:9] + row[15:17]]
        if analyze_row(*params):
            base.update(analyze_row(*params))
    
    return base


def gsheets_guest_load(table_key: str) -> dict:
    list_of_rows = gsheets_load_data(table_key)
    base = dict()
    for num, row in enumerate(list_of_rows):
        oname = ''
        if len(row[1].split()) == 2:
            fname, name = row[1].split()
        if len(row[1].split()) == 3:
            fname, name, oname = row[1].split()
        if analyze_row(
            number=str(num),
            phone_number=row[2],
            fname=fname,
            name=name,
            oname=oname,
            card='',
            rname='',
            roname=''
        ):
            base.update(
                analyze_row(
                    number=str(num),
                    phone_number=row[2],
                    fname=fname,
                    name=name,
                    oname=oname,
                    card='',
                    rname='',
                    roname=''
                )
            )

    return base    


def gsheets_save(table_key: str, numbers: dict, sheet: int = 0, debug: bool = False):
    google_connect = gspread.service_account(filename=GSHEETS_API_KEY)
    gsheet = google_connect.open_by_key(table_key)
    worksheet = gsheet.get_worksheet(sheet)
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

    cells_list = ['N', 'Phone number', 'Name', 'Balance', 'Total costs', 'Loyality']
    range_table = 'F'
    width_cells_list = [ 
        ('A:', 50),
        ('B:', 100),
        ('C:', 300),
        ('F:', 200),
    ]

    for number, (phone, data) in enumerate(numbers.items(), start=1):
        data_list = [
            number,
            phone,
            data['name'],
            data['balance'],
            data['total_costs'],
            data['loyality']
        ]

        if debug:
            data_list.append(','.join(data['file']))

        write_list.append(data_list)

    if debug:
        cells_list.append('File name')
        range_table = 'G'
        width_cells_list.append(('G:', 300))

    worksheet.update('A1', [cells_list])

    gspread_formatting.set_column_widths(worksheet, width_cells_list)

    range_table += str(len(write_list) + 1)

    worksheet.update('A2:' + range_table, write_list)
    gspread_formatting.format_cell_ranges(
        worksheet,
        [('A1:'+ range_table, table_style), (f'A1:{range_table[0]}1', title_format)]
    )
    worksheet.format('D2:' + range_table, float_style)
