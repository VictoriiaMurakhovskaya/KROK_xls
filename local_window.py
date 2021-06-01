# -*- coding: utf-8 -*-

from tkinter import Tk, Label, filedialog, font, LEFT, StringVar
from tkinter import messagebox as mb
from tkinter import filedialog as fd
from tkinter.ttk import Combobox, Spinbox, Button
from main import make_file
import sys


def generate_table():
    if (file_name := fd.askopenfilename(defaultextension='.xlsx',
                                        filetypes=[('Excel таблиці', '*.xlsx'),
                                                   ('Excel таблиці (2007)', '*.xls')])) == '':
        mb.showwarning(title='Помилка', message='Некоректно обрано файл')
        return

    if (savedir := filedialog.askdirectory()) == '':
        mb.showwarning(title='Помилка', message='Некоректно обрані файли')
        return

    if make_file(file_name, savedir, int(clmn_values_str[0].get()),
                 int(clmn_values_str[1].get()), int(clmn_values_str[2].get()),
                 int(clmn_values_str[3].get()),
                 0, 0,
                 list(font.families())[cmb_font.current()],
                 int(f_size.get())):
        mb.showinfo(title='Інформація', message='Файли сформовано')
    else:
        mb.showwarning(title='Помилка', message='Помилка обробки файла')


def finish_it():
    sys.exit()


if __name__ == '__main__':
    window = Tk()
    f_size = StringVar(window)
    window.title("Формування Excel")
    window.geometry('450x135')
    font_list = list(font.families())
    lbl1 = Label(window, text='Тип шрифта', justify=LEFT)
    lbl1.grid(column=0, row=0, padx=5, pady=5, columnspan=2)
    cmb_font = Combobox(window, width=20, height=5, values=list(font.families()), state='readonly')
    cmb_font.current(1)
    cmb_font.grid(column=2, row=0, padx=5, pady=5, columnspan=5)
    lbl2 = Label(window, text='Размер шрифта', justify=LEFT)
    lbl2.grid(column=7, row=0, padx=5, pady=5, columnspan=3)
    font_size = Spinbox(window, from_=0, to=300, width=3, textvariable=f_size)
    f_size.set("9")
    font_size.grid(column=10, row=0, padx=5, pady=5, columnspan=2)
    lbl3 = Label(window, text='Размеры колонок вывода', justify=LEFT)
    lbl3.grid(column=0, row=1, padx=5, pady=5, columnspan=4)
    clmn_labels = []
    clmn_sizes = []
    clmn_values = [10, 40, 10, 8]
    clmn_values_str = []
    for i in range(0, 4):
        clmn_labels.append(Label(window, text=str(i + 1) + ': ', justify=LEFT))
        clmn_labels[i].grid(column=2 * i, row=2, padx=5, pady=5)
        clmn_values_str.append(StringVar(window))
        clmn_values_str[i].set(str(clmn_values[i]))
        clmn_sizes.append(Spinbox(window, from_=5, to=40, width=2, textvariable=clmn_values_str[i]))
        clmn_sizes[i].grid(column=2 * i + 1, row=2, padx=5, pady=5)
    btn = Button(window, text="Сформувати", width=15, command=generate_table)
    btn.grid(column=0, row=3, pady=5, padx=5, columnspan=3)
    btn2 = Button(window, text="Вийти", width=15, command=finish_it)
    btn2.grid(column=3, row=3, pady=5, padx=5, columnspan=3)
    window.mainloop()
