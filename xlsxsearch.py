import openpyxl as xl
from openpyxl.utils import get_column_letter

import sys
from pathlib import Path
import os
from copy import copy
import PySimpleGUI as sg

import squery

col_width = []

def gen_file_list(folder):
    files = []
    p = Path(folder)
    for i in p.glob('**/*.xlsx'):
        if os.fspath(i).find("xlsxsearch_") < 0:
            files.append(os.fspath(i))
    return files

def run_search(files, searchstring, query_func, dest_filename, fix_search, w, status):
    status[0]('')
    status[1]('')
    status[2]('')
    w.Refresh()
    wb = xl.Workbook()
    ws1 = wb.active
    ws1.title = searchstring
    first = True
    dest_row = 2
    for filename in files:
        workbook = xl.load_workbook(filename)
        # Get the active worksheet
        ws = workbook.active
        max_row = ws.max_row
        max_column = ws.max_column
        if len(col_width) < max_column:
            col_width.extend([0] * (max_column - len(col_width)))
        if first:
            status[0]("Took formatting from: " + filename)
            w.Refresh()
            #ws1.column_dimensions = copy(ws.column_dimensions)
            ws1.sheet_view.rightToLeft = copy(ws.sheet_view)
            ws1.sheet_format = copy(ws.sheet_format)
            #ws1.dimensions = copy(ws.dimensions)
            ws1.sheet_properties = copy(ws.sheet_properties)
            #ws1.merged_cells = copy(ws.merged_cells)
            ws1.page_margins = copy(ws.page_margins)
            ws1.page_setup = copy(ws.page_setup)
            ws1.print_options = copy(ws.print_options)
            copy_row(ws, ws1, 1, 1)
            first = False
        # ws.rows[1:] means we'll skip the first row (the header row).
        for row in range(1, max_row+1):
            found = False
            for column in range(1, max_column+1):
                c = str(ws.cell(row=row, column=column).value)
                if c and query_func(c):
                    found = True
                    break
            if found:
                status[1]("Found " + str(dest_row - 1) + (" row" if dest_row == 2 else " rows"))
                w.Refresh()
                copy_row(ws, ws1, row, dest_row, fix_search, searchstring)
                dest_row += 1
        dims = {}
        for row in ws1.rows:
            for cell in row:
                if cell.value:
                    dims[cell.column] = max((dims.get(cell.column, 0), len(str(cell.value))))
        for column in range(len(col_width)):
            ws1.column_dimensions[get_column_letter(column+1)].width = col_width[column]
    wb.save(dest_filename)
    status[2]("Done!")

def copy_row(ws_s, ws_d, src_row, dest_row, fix_search=False, searchstring=''):
    max_column = ws_s.max_column
    ws_d.row_dimensions[dest_row] = copy(ws_s.row_dimensions[src_row])
    for column in range(1, max_column + 1):
        cs = ws_s.cell(row=src_row, column=column)
        if cs.value and fix_search:
            new_value = str(cs.value).replace(searchstring, "__"+searchstring+"__")
        else:
            new_value = cs.value
        cd = ws_d.cell(row=dest_row, column=column, value=new_value)
        if cs.value:
            #print(ws_s.column_dimensions[column].width)
            col_width[column-1] = max(col_width[column-1], ws_s.column_dimensions[get_column_letter(column)].width)
        if cs.has_style:
            cd.font = copy(cs.font)
            cd.border = copy(cs.border)
            cd.fill = copy(cs.fill)
            cd.number_format = copy(cs.number_format)
            cd.protection = copy(cs.protection)
            cd.alignment = copy(cs.alignment)
        if cs.hyperlink:
            cd.hyperlink = copy(cs.hyperlink)
        if cs.comment:
            cd.comment = copy(cs.comment)

#
# start
#
sfb = sg.FolderBrowse()
dfb = sg.FolderBrowse()
status = [None] * 3
for i in range(3):
    status[i] = sg.Text('', size=(100, 1))
layout = [
    [sg.Text('xlsxsearch (21-oct-2021) - Copyright 2020-2021 by Udi Finkelstein', size=(100, 1))],
    [sg.Text('https://github.com/udif/xlsxsearch/', size=(100, 1))],
    [sg.Text('', size=(100, 1))],
    [sg.Text('Source Folders with XLSX files', size=(30, 1)),
     sg.InputText(key='-SFOLDER-', size=(100, 1),
                  default_text = sfb.InitialFolder), sfb],
    [sg.Text('Destination for search result XLSX files', size=(30, 1)),
     sg.InputText(key='-DFOLDER-', size=(100, 1),
                  default_text = dfb.InitialFolder), dfb],
    [sg.Text('Search query', size=(30, 1)), sg.InputText('', size=(100, 1), key='-QUERY-')],
    [sg.Checkbox('Mark results with __XX__', key='-CB-', size=(20,1), default=True)],
    [sg.Text('', size=(100, 1))],
    [sg.Submit()],
    [sg.Text('', size=(100, 1))],
    [status[0]],
    [status[1]],
    [status[2]]
    ]

w = sg.Window("xlsxsearch", layout, margins=(80, 50))

folder = ''
while True:
    folder = "."
    file_list = []
    event, values = w.read()
    # End program if user closes window or
    # presses the OK button
    if event == "Submit":
        fix_search = values["-CB-"]
        sfolder = values["-SFOLDER-"]
        dfolder = values["-DFOLDER-"]
        file_list = gen_file_list(sfolder)
        searchstring = values["-QUERY-"]
        query_func = squery.squery_compile(searchstring)
        dest_file = os.fspath(Path(dfolder).joinpath("xlsxsearch_" + searchstring + ".xlsx"))
        try:
            run_search(file_list, searchstring, query_func, dest_file, fix_search, w, status)
        except OSError as err:
            status[2]('OS Error: {}'.format(err))
            w.Refresh()

    if event == sg.WIN_CLOSED:
        break
    if event == sg.WIN_CLOSED:
        w.close()
        sys.exit(0)
    if event == "-SFOLDER-":
        pass
w.close()


#print(files)
