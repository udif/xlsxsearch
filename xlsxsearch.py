import openpyxl
import xlsxwriter

import sys
from pathlib import Path
import os
from copy import copy
import PySimpleGUI as sg
import pprint

import squery

col_width = []
pp = pprint.PrettyPrinter(indent=4)

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
    #
    # openpyxl
    #
    wb_w = openpyxl.Workbook()
    ws_w = wb_w.active
    ws_w.title = searchstring
    #
    # xlsxwriter
    #
    (h, t) = os.path.split(dest_filename)
    wb_w2 = xlsxwriter.Workbook(os.path.join(h, "_" + t))
    ws_w2 = wb_w2.add_worksheet()
    ws_w2.activate()
    #
    first = True
    dest_row = 2
    for filename in files:
        wb_r = openpyxl.load_workbook(filename)
        # Get the active worksheet
        ws_r = wb_r.active
        max_row = ws_r.max_row
        max_column = ws_r.max_column
        if len(col_width) < max_column:
            col_width.extend([0] * (max_column - len(col_width)))
        if first:
            status[0]("Took formatting from: " + filename)
            w.Refresh()
            #ws_w.column_dimensions = copy(ws_r.column_dimensions)
            ws_w.sheet_view.rightToLeft = copy(ws_r.sheet_view.rightToLeft)
            if ws_r.sheet_view.rightToLeft:
                ws_w2.right_to_left()
            ws_w.sheet_format = copy(ws_r.sheet_format) # https://openpyxl.readthedocs.io/en/stable/api/openpyxl.worksheet.dimensions.html?highlight=SheetFormatProperties#openpyxl.worksheet.dimensions.SheetFormatProperties
            #ws_w.dimensions = copy(ws_r.dimensions)
            ws_w.sheet_properties = copy(ws_r.sheet_properties) #https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/3.0/openpyxl/worksheet/properties.py #WorksheetProperties
            #ws_w.merged_cells = copy(ws_r.merged_cells)
            ws_w.page_margins = copy(ws_r.page_margins)
            # page_margins is broken into 3 lines in xlsxwriter
            ws_w2.set_margins(ws_r.page_margins.left, ws_r.page_margins.right, ws_r.page_margins.top, ws_r.page_margins.bottom)
            ws_w2.set_header(options={'margin': ws_r.page_margins.header})
            ws_w2.set_footer(options={'margin': ws_r.page_margins.footer})
            ws_w.page_setup = copy(ws_r.page_setup) # https://xlsxwriter.readthedocs.io/page_setup.html and check https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/3.0/openpyxl/worksheet/page.py
            ws_w.print_options = copy(ws_r.print_options)
            openpyxl_copy_row(ws_r, ws_w, 1, 1)
            xlsxwriter_copy_row(ws_r, ws_w2, 1, 0)
            first = False
        # ws_r.rows[1:] means we'll skip the first row (the header row).
        for row in range(1, max_row+1):
            found = False
            for column in range(1, max_column+1):
                c = str(ws_r.cell(row=row, column=column).value)
                if c and query_func(c):
                    found = True
                    break
            if found:
                status[1]("Found " + str(dest_row - 1) + (" row" if dest_row == 2 else " rows"))
                w.Refresh()
                openpyxl_copy_row(ws_r, ws_w, row, dest_row, fix_search, searchstring)
                xlsxwriter_copy_row(ws_r, ws_w2, row, dest_row - 1, fix_search, searchstring)
                dest_row += 1
        dims = {}
        for row in ws_w.rows:
            for cell in row:
                if cell.value:
                    dims[cell.column] = max((dims.get(cell.column, 0), len(str(cell.value))))
        for column in range(len(col_width)):
            if col_width[column] > 0:
                ws_w.column_dimensions[openpyxl.utils.get_column_letter(column+1)].width = col_width[column]
                ws_w2.set_column(dest_row - 1, column, col_width[column])
    wb_w.save(dest_filename)
    wb_w2.close()
    status[2]("Done!")

def openpyxl_copy_row(ws_s, ws_d, src_row, dest_row, fix_search=False, searchstring=''):
    max_column = ws_s.max_column
    ws_d.row_dimensions[dest_row] = copy(ws_s.row_dimensions[src_row])
    for column in range(1, max_column + 1):
        cs = ws_s.cell(row=src_row, column=column)
        if cs.value:
            if fix_search:
                new_value = str(cs.value).replace(searchstring, "__"+searchstring+"__")
            else:
                new_value = cs.value
            cd = ws_d.cell(row=dest_row, column=column, value=new_value)
            #print(ws_s.column_dimensions[column].width)
            col_width[column-1] = max(col_width[column-1], ws_s.column_dimensions[openpyxl.utils.get_column_letter(column)].width)
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

def xlsxwriter_copy_row(ws_s, ws_d, src_row, dest_row, fix_search=False, searchstring=''):
    max_column = ws_s.max_column
    rd = ws_s.row_dimensions[src_row]
    ws_d.set_row(dest_row, rd.ht, None, {'hidden' : rd.hidden, 'level' : rd.outlineLevel , 'collapsed' : rd.collapsed})
    for column in range(max_column):
        cs = ws_s.cell(row=src_row, column=column + 1)
        if cs.value:
            if fix_search:
                new_value = str(cs.value).replace(searchstring, "__"+searchstring+"__")
            else:
                new_value = cs.value
            ws_d.write_string(dest_row, column, new_value)
            #print(ws_s.column_dimensions[column].width)
            col_width[column] = max(col_width[column], ws_s.column_dimensions[openpyxl.utils.get_column_letter(column + 1)].width)
            if cs.has_style:

                print("Cell:")
                pp.pprint (cs.alignment)
            #    cd.border = copy(cs.border)
            #    cd.fill = copy(cs.fill)
            #    cd.number_format = copy(cs.number_format)
            #    cd.protection = copy(cs.protection)
            #    cd.alignment = copy(cs.alignment)
            #if cs.hyperlink:
            #    cd.hyperlink = copy(cs.hyperlink)
            if cs.comment:
                ws_d.write_comment(dest_row, column, cs.comment)

#
# start
#
sfb = sg.FolderBrowse()
dfb = sg.FolderBrowse()
status = [None] * 3
for i in range(3):
    status[i] = sg.Text('', size=(100, 1))
layout = [
    [sg.Text('xlsxsearch (21-oct-2021, $Id$) - Copyright 2020-2021 by Udi Finkelstein', size=(100, 1))],
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
    [sg.Checkbox('optional "and"', key='-IA-', size=(20,1), default=True, tooltip='If selected, then multiple words without AND or OR are assumed to have implicit AND operator between them')],
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
        (o, a, n) = squery.squery_get_keywords()
        a = list(filter(lambda n: len(n) > 0, a))
        if values["-IA-"]:
            a.append('')
        squery.squery_set_keywords(o, a, n)
        sfolder = values["-SFOLDER-"]
        dfolder = values["-DFOLDER-"]
        file_list = gen_file_list(sfolder)
        searchstring = values["-QUERY-"]
        query_func = squery.squery_compile(searchstring)
        if query_func is None:
            status[2]("Illegal query!")
            continue
        filename = searchstring.replace('"', "'")
        dest_file = os.fspath(Path(dfolder).joinpath("xlsxsearch_" + filename + ".xlsx"))
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
