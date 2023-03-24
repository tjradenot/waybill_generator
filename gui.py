import tkinter as tk
from tkinter import ttk
from tkcalendar import Calendar
from waybill import main, check_files
import time
from settings import gui_data

class Application:

    def __init__(self):
        # Настройки основного окна
        self.window = tk.Tk()
        # https://www.iconarchive.com/show/material-icons-by-pictogrammers/excavator-icon.html
        self.window.iconbitmap('icon.ico')
        self.window.title(gui_data['title'])
        self.window.geometry(f"{gui_data['width']}x{gui_data['height']}")
        self.window.resizable(False, False)

        # Закрыть программу по нажатиюю Escape
        self.window.bind('<Escape>', lambda event: self.window.quit())

        # Хранит 0 или 1. Нужна для функции show_date.
        self.flag = 0

        # VIN
        self.vin_frame = ttk.Frame(self.window)

        self.vin_label = ttk.Label(
            self.vin_frame, text='VIN:', font='Calibri 14')
        self.vin_label.pack(side='left')

        # Хранит ВИН, который ввел пользователь
        self.vin_user_input = tk.StringVar(value='')

        self.vin_entry = ttk.Entry(
            self.vin_frame, font='Calibri 14 bold', textvariable=self.vin_user_input)
        self.vin_entry.insert(0, '')
        self.vin_entry.focus_set()  # Активный курсор в текстовом поле ввода
        self.vin_entry.pack(side='left', padx=6)

        self.vin_frame.pack()

        # Календарь
        self.calendar = Calendar(self.window, date_pattern='dd.mm.y')
        self.calendar.pack(pady=10)

        # Выбор дат
        # Для хранения и изменения дат в переменных 'начало периода' и 'конец периода'
        self.start_date_user_input = tk.StringVar(
            value='00.00.0000')
        self.end_date_user_input = tk.StringVar(value='00.00.0000')

        self.dates_frame = ttk.LabelFrame(self.window, text='Период')

        self.start_date_label = ttk.Label(
            self.dates_frame, text='Начало периода:', font='Calibri 14')

        # sticky=tk.W прикрепляется к западу
        self.start_date_label.grid(row=0, column=0, sticky=tk.W, padx=3)

        self.start_date_show_text = ttk.Label(
            self.dates_frame, font='Calibri 14 bold', textvariable=self.start_date_user_input)
        self.start_date_show_text.grid(row=0, column=1, sticky=tk.E)

        self.end_date_label = ttk.Label(
            self.dates_frame, text='Конец периода:', font='Calibri 14')
        self.end_date_label.grid(row=1, column=0, sticky=tk.W, pady=5, padx=3)

        self.end_date_show_text = ttk.Label(
            self.dates_frame, font='Calibri 14 bold', textvariable=self.end_date_user_input)
        self.end_date_show_text.grid(row=1, column=1, sticky=tk.E)

        self.dates_frame.pack(pady=6)

        # Смена фрейм
        self.shift_frame = ttk.Frame(self.window)
        self.shift_user_input = tk.StringVar(value='все')

        # Смена текст
        self.shift_title = ttk.Label(
            self.shift_frame, text='Смена:', font='Calibri 14')
        self.shift_title.pack(side='left')

        # Выпадающий список для выбора смен
        self.shift_dropdown_list = ttk.Combobox(
            master=self.shift_frame,
            width=5,
            state="readonly",
            values=['все', 'день', 'ночь'],
            font='Calibri 14 bold',
            textvariable=self.shift_user_input
        )
        self.shift_dropdown_list.pack(side='left', padx=6)

        self.shift_frame.pack()

        # Передает event "<<CalendarSelected>>" в функцию self.show_date. Пользователь выбрал дату и выбранная дата передается в show_date.
        self.calendar.bind('<<CalendarSelected>>', self.show_date)

        # Кнопка запуска
        self.run_button_style = ttk.Style()
        self.run_button_style.configure('run.TButton', font=('Calibri 18'))
        self.run_button = ttk.Button(
            text='Запустить', style='run.TButton', width=20, command=lambda: main(self.create_data()))
        self.run_button.pack(pady=3)

        self.window.mainloop()

    def show_date(self, event):
        """Записывает выбранные даты в переменные 'начало периода' и 'конец периода'"""
        if self.flag == 0:
            self.start_date_user_input.set(self.calendar.get_date())
            self.flag = 1
        elif self.flag == 1:
            self.end_date_user_input.set(self.calendar.get_date())
            self.flag = 0

    def create_data(self):
        """Возвращает словарь. ВИН - ключ. Список - начало периода, конец периода, смена."""
        return {self.vin_user_input.get(): [self.start_date_user_input.get(), self.end_date_user_input.get(), self.shift_user_input.get()]}


if __name__ == '__main__':
    if check_files():
        app = Application()
    else:
        print('Завершаю работу...')
        for i in range(5, 0, -1):
            time.sleep(1)
