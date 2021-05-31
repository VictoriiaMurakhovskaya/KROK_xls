# -*- coding: utf-8 -*-

import xlrd
import xlsxwriter
from tkinter import font
import tkinter
import re
import traceback

# параметры формируемой таблицы
# можно менять здесь или интерактивно
# высота строк
row_width = 10

# строка заголовков
headers_row = 1

# начало перечня студентов
start_row = 5

# ширина колонок
columns_width = {'A': 10,
                 'B': 40,
                 'C': 10,
                 'D': 8,
                 'E': 8,
                 'F': 8}

# заголовки столбцов таблицы
head = ['Код освітнього компоненту/ Component code',
        'Назва дисципліни / Course Title',
        'Кількість кредитів Європейської кредитної трансферно - накопичувальної системи / ECTS Credits',
        'Оцінка за шкалою заклада вищої освіти / Institutional Grade']


def insert_subheader(wb, ws, name, row_number, p1, p2, size):
    sub_header_format = wb.add_format({'bold': True,
                                       'font_size': str(size),
                                       'font_name': 'Calibri',
                                       'align': 'left',
                                       'text_wrap': True,
                                       'valign': 'vcenter',
                                       })
    sub_header_numbers = wb.add_format({'bold': True,
                                        'font_size': str(size),
                                        'font_name': 'Calibri',
                                        'align': 'center',
                                        'text_wrap': True,
                                        'valign': 'vcenter',
                                        })
    sub_header_format.set_border()
    sub_header_format.set_indent(1)
    sub_header_numbers.set_border()
    ws.write(row_number, 0, '', sub_header_format)
    ws.write(row_number, 1, name, sub_header_format)
    for k in range(2, 4):
        ws.write(row_number, k, '', sub_header_format)
    if p1 != 0:
        ws.write(row_number, 2, p1, sub_header_numbers)
        #ws.write(row_number, 4, p2, sub_header_numbers)


def is_number(value):
    try:
        return int(value)
    except:
        return 0


def make_file(file_name, save_dir, w1, w2, w3, w4, w5, w6, font, size):
    global headers_row
    columns_width = {'A': w1,
                     'B': w2,
                     'C': w3,
                     'D': w4,
                     'E': w5,
                     'F': w6}
    # установка форматов
    header_dict = {'bold': True,
                   'font_size': str(size),
                   'font_name': font,
                   'align': 'center',
                   'text_wrap': True,
                   'valign': 'vcenter'}

    row_numbers_dict = {'bold': False,
                        'font_size': str(size),
                        'font_name': font,
                        'text_wrap': True,
                        'valign': 'vcenter',
                        'align': 'center'}

    row_strings_dict = {'bold': False,
                        'font_size': str(size),
                        'font_name': font,
                        'text_wrap': True,
                        'valign': 'vcenter',
                        'align': 'left'}

    rb = xlrd.open_workbook(file_name, formatting_info=False)
    sheet = rb.sheet_by_index(0)

    # предметы
    subjs = []
    subjs_params = {}
    subj_min = 0
    subj_max = 0
    subj_list = 1
    for i in range(1, sheet.ncols):
        cell = str(sheet.cell(headers_row - 1, i).value.strip())
        if cell:
            if subj_min == 0: subj_min = i
            subj_max = i
            subjs.append(subj_list)
            subjs_params.update({subj_list: [sheet.cell(headers_row - 1, i).value,
                                             sheet.cell(headers_row, i).value.strip(),
                                             sheet.cell(headers_row + 1, i).value,
                                             sheet.cell(headers_row + 2, i).value,
                                             sheet.cell(headers_row + 3, i).value,
                                             sheet.cell(headers_row - 2, i).value]})
            subj_list += 1
    # имена студентов
    names = []
    marks = {}
    for i in range(headers_row + 3, sheet.nrows):
        cell = str(sheet.cell(i, 1).value)
        if cell.find(' ') > 0:
            names.append(cell)
            marks_of_student = []
            for j in range(subj_min, subj_max + 1):
                marks_of_student.append(sheet.cell(i, j).value)
            marks.update({cell: marks_of_student})

    for student in names:
        match = re.search(r'[А-Яа-яЁёЇїІіЄє]*', student)
        s_name = match.group()
        wb = xlsxwriter.Workbook(save_dir + '/' + s_name + '.xlsx', {'in_memory': True})
        header_format = wb.add_format(header_dict)
        header_format.set_border()
        row_numbers = wb.add_format(row_numbers_dict)
        row_numbers.set_border()
        row_strings = wb.add_format(row_strings_dict)
        row_strings.set_border()
        row_strings.set_indent(1)
        ws = wb.add_worksheet('Total')
        for col in columns_width.keys():
            ws.set_column(col + ':' + col, columns_width[col])
        # заголовок таблицы
        for i in range(0, len(head)):
            ws.write(0, i, head[i], header_format)

        # строки таблицы
        count = 0
        count_rows = 0
        count_n = 1
        total_credits = 0
        total_hours = 0
        for subj in subjs:
            if subjs_params[subj][2] == 'P':
                count_rows += 1
                count += 1
                insert_subheader(wb, ws, subjs_params[subj][1], count_rows, 0, 0, size)
            else:
                score = is_number(marks[student][count])
                count += 1
                if score > 0:
                    if subjs_params[subj][3]:
                        total_credits += int(subjs_params[subj][3])
                    if subjs_params[subj][4]:
                        total_hours += int(subjs_params[subj][4])
                    count_rows += 1
                    ws.write(count_rows, 0, subjs_params[subj][0], row_numbers)
                    for k, n in zip([0, 1, 3], range(0, 3)):
                        ws.write(count_rows, n, ('-' if not subjs_params[subj][k] else subjs_params[subj][k]),
                                 row_strings if k == 1 else row_numbers)
                    ws.write(count_rows, 3, score, row_numbers)
                    count_n += 1
        insert_subheader(wb, ws, 'Загальна кількість кредитів Європейської кредитної трансферно - накопичувальної системи'
                                 ' / Total ECTS Credits',
                         count_rows + 1, total_credits, total_hours, size)
        try:
            wb.close()
        except :
            print(
            '>>> traceback <<<')
            traceback.print_exc()
            print(
            '>>> end of traceback <<<')
    return True


def installed_fonts():
    root = tkinter.Tk()
    font_list = list(font.families())
    root = None
    # return [{'label': font_list[i], 'value': i} for i in range(0, len(font_list))]
    return font_list


if __name__ == "__main__":
    make_file('sample_file.xls', 'Out', 10, 40, 10, 8, 8, 8, {'label': 'Calibri'}, '10')
