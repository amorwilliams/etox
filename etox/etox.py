 # encoding: utf-8

import xlrd
import sys

from sheet import shm

args = {}

def export_single_book():
    file_path = args.input
    output_path = args.output

    shm.add_work_book(file_path)
    sheet_name_list = shm.get_sheet_name_list()

    for sheet_name in sheet_name_list:
        if shm.is_ref_sheet(sheet_name):
            continue

        print('Exporting: %s' % sheet_name)

        py = shm.get_sheet(sheet_name).to_python()
        f = file(output_path + sheet_name + '.py', 'w')
        f.write(str(py).encode('UTF-8'))
        f.close()

        json = shm.export_json(sheet_name)
        f = file(output_path + sheet_name + '.json', 'w')
        f.write(json.encode('UTF-8'))
        f.close()

        lua = shm.export_lua(sheet_name)
        f = file(output_path + sheet_name + '.lua', 'w')
        f.write(lua.encode('UTF-8'))
        f.close()


def export_main_book():
    file_path = args.input
    output_path = args.output

    wb = xlrd.open_workbook(file_path)
    sh = wb.sheet_by_index(0)

    workbook_path_list = []
    sheet_list = []

    for row in range(sh.nrows):
        type = sh.cell(row, 0).value

        if type == '__workbook__':
            pass
        else:
            sheet_list.append([])
            sheet = sheet_list[-1]
            sheet.append(type)

        for col in range(1, sh.ncols):
            value = sh.cell(row, col).value

            if type == '__workbook__' and value != '':
                workbook_path_list.append(value)
            elif value != '':
                sheet.append(value)

    for workbook_path in workbook_path_list:
        shm.add_work_book(workbook_path + '.xlsx')

    for sheet in sheet_list:
        if '->' in sheet[0]:
            sheet_name = sheet[0].split('->')[0]
            sheet_output_name = sheet[0].split('->')[1]
        else:
            sheet_output_name = sheet_name = sheet[0]

        sheet_output_field = sheet[1:]

        # py = shm.get_sheet(sheet_name).to_python(sheet_output_field)
        # f = file(output_path + sheet_name + '.py', 'w')
        # f.write(str(py).encode('UTF-8'))
        # f.close()

        print('Exporting: %s -> %s' % (sheet_name, sheet_output_name))

        json = shm.export_json(sheet_name, sheet_output_field)
        f = file(output_path + sheet_name + '.json', 'w')
        f.write(json.encode('UTF-8'))
        f.close()

        lua = shm.export_lua(sheet_name, sheet_output_field)
        f = file(output_path + sheet_name + '.lua', 'w')
        f.write(lua.encode('UTF-8'))
        f.close()


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='Export excle to assigned file, now support [json, lua]', prog='etox')
    parser.add_argument('-v', '--version', action='version', version='%(prog)s v1.0')
    parser.add_argument('-m', help='Use mainbook mode, otherwise will use singlebook mode.', action='store_true')
    parser.add_argument("-i", '--input', help='Input filename', type=str)
    parser.add_argument("-o", '--output', help='Output filepath', type=str)
    args = parser.parse_args()

    if not args.m:
        export_single_book()
    else:
        export_main_book()
