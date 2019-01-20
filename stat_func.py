#coding:utf-8

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from openpyxl.styles.borders import Border, Side


# FILE_PATH = r'/home/user/work/odoo12/repositories/all-euro-data-2018-2019.xlsx'

#FILE_PATH = r'D:\py_projects\serega\test.xlsx'
SHEET_DATA = {}
COLUMN_NAMES = ['H_S', 'H_M', 'A_S', 'A_M']
DA_INDEX = column_index_from_string('DA')


def check_headers(sheet):
    headers = [sheet.cell(row=1, column=i).value for i in range(1, sheet.max_column+1)]
    if 'H_S1' in headers:
        return True
    return False


def prepare_sheet(sheet):
    column = DA_INDEX

    for name in COLUMN_NAMES:
        for i in range(1, 8, 1):
            sheet.cell(row=1, column=column).value = "%s%s" % (name, i)
            column += 1


def take_data(sheet, row_num, team):
    name = sheet.cell(row=row_num, column=team).value
    data = {'score': [], 'missed': []}
    first = True

    for j in range(row_num, 1, -1):
        # if name == sheet.cell(row=j, column=3).value and first:
        #     first = False
        #     continue

        if name == sheet.cell(row=j, column=3).value:
            data['score'].append(sheet.cell(row=j, column=5).value)
            data['missed'].append(sheet.cell(row=j, column=6).value)

        elif name == sheet.cell(row=j, column=4).value:
            data['score'].append(sheet.cell(row=j, column=6).value)
            data['missed'].append(sheet.cell(row=j, column=5).value)

        elif len(data['score']) == 7:
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


def write_data(sheet):
    for val in SHEET_DATA:
        data = SHEET_DATA.get(val)
        team = val[1]
        row_num = val[0]

        if team == 3:
            column_score = DA_INDEX
        else:
            column_score = DA_INDEX + 14

        column_miss = column_score + 7

        set_border(sheet, row_num, column_score)
        for d in data.get('score'):
            sheet.cell(row=row_num, column=column_score).value = d
            column_score += 1

        set_border(sheet, row_num, column_miss)
        for d in data.get('missed'):
            sheet.cell(row=row_num, column=column_miss).value = d
            column_miss += 1
