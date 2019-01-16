#coding:utf-8

from openpyxl import load_workbook
from openpyxl.styles.borders import Border, Side


# FILE_PATH = r'/home/user/work/odoo12/repositories/all-euro-data-2018-2019.xlsx'

#FILE_PATH = r'D:\py_projects\serega\test.xlsx'
SHEET_DATA = {}
COLUMN_NAMES = ['H_S', 'H_M', 'A_S', 'A_M']


def check_headers(sheet):
    headers = [sheet.cell(row=1, column=i).value for i in range(1, sheet.max_column+1)]
    if 'H_S1' in headers:
        return True
    return False


def prepare_sheet(sheet, max_column):
    column = max_column + 11

    for name in COLUMN_NAMES:
        for i in range(1, 8, 1):
            sheet.cell(row=1, column=column).value = "%s%s" % (name, i)
            column += 1


def take_data(sheet, row_num, team):
    name = sheet.cell(row=row_num, column=team).value
    data = {'score': [], 'missed': []}
    for j in range(row_num, 1, -1):
        if name == sheet.cell(row=j, column=3).value:
            data['score'].append(sheet.cell(row=j, column=5).value)
            data['missed'].append(sheet.cell(row=j, column=6).value)

        elif name == sheet.cell(row=j, column=4).value:
            data['score'].append(sheet.cell(row=j, column=6).value)
            data['missed'].append(sheet.cell(row=j, column=5).value)
        if len(data['score']) == 7:
            SHEET_DATA.update({
                tuple((row_num, team)): data
            })
            break
    if data['score']:
        SHEET_DATA.update({
            tuple((row_num, team)): data
        })


def set_border(sheet, row_num, column):
    sheet.cell(row=row_num, column=column).border = Border(
        left=Side(style='medium'))


def write_data(max_column, sheet):
    for val in SHEET_DATA:
        data = SHEET_DATA.get(val)
        team = val[1]
        row_num = val[0]

        if team == 3:
            column_score = max_column + 11
        else:
            column_score = max_column + 25

        column_miss = column_score + 7

        set_border(sheet, row_num, column_score)
        for d in data.get('score'):
            sheet.cell(row=row_num, column=column_score).value = d
            column_score += 1

        set_border(sheet, row_num, column_miss)
        for d in data.get('missed'):
            sheet.cell(row=row_num, column=column_miss).value = d
            column_miss += 1


def run(file_path):
    wb = load_workbook(file_path)
    sheets = wb.worksheets
    team_from_columns = [3, 4]

    all_sheets_count = len(sheets)
    counter = 1

    for sheet in sheets:
        print("Processed sheet - %s, from %s" % (counter, all_sheets_count))

        check_sheet = check_headers(sheet)
        max_column = sheet.max_column

        if not check_sheet:
            prepare_sheet(sheet, max_column)
        elif check_sheet:
            max_column = max_column - 38

        for row_num in range(sheet.max_row, 1, -1):
            for team in team_from_columns:
                take_data(sheet, row_num, team)

        write_data(max_column, sheet)
        SHEET_DATA.clear()
        counter += 1

    wb.save(file_path)
