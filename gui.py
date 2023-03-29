import tkinter as tk
from tkinter import ttk
from tkcalendar import Calendar
from waybill import check_files, compare_dates, generate_waybills, filter_df, change_path, get_df_from_excel_file, old_data
import time
from settings import gui_data

change_path()


class Application:

    def __init__(self, __df, color='green', message='Проверка файлов завершена.'):
        # Настройки основного окна
        self.window = tk.Tk()
        # https://www.iconarchive.com/show/material-icons-by-pictogrammers/excavator-icon.html
        self.window.iconbitmap('icon.ico')
        self.window.title(gui_data['title'])
        self.window.geometry(f"{gui_data['width']}x{gui_data['height']}")
        self.window.resizable(False, False)

        # Закрыть программу по нажатиюю Escape
        self.window.bind('<Escape>', lambda event: self.window.quit())

        # Запустить генерацию по нажатию Enter
        self.window.bind('<Return>', lambda event: self.create_data(__df))

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

        self.start_date_show_text = ttk.Entry(
            self.dates_frame, font='Calibri 14 bold', width=10, textvariable=self.start_date_user_input)
        self.start_date_show_text.grid(row=0, column=1, sticky=tk.E)

        self.end_date_label = ttk.Label(
            self.dates_frame, text='Конец периода:', font='Calibri 14')
        self.end_date_label.grid(row=1, column=0, sticky=tk.W, pady=5, padx=3)

        self.end_date_show_text = ttk.Entry(
            self.dates_frame, font='Calibri 14 bold', width=10, textvariable=self.end_date_user_input)
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

        # Separator
        self.separator1 = ttk.Separator(self.window)
        self.separator1.pack(pady=6, fill='x')

        # Передает event "<<CalendarSelected>>" в функцию self.show_date. Пользователь выбрал дату и выбранная дата передается в show_date.
        self.calendar.bind('<<CalendarSelected>>', self.show_date)

        # Кнопка запуска
        self.run_button_style = ttk.Style()
        self.run_button_style.configure('run.TButton', font=('Calibri 16 bold'))
        self.run_button = ttk.Button(
            text='Запустить', style='run.TButton', width=20, command=lambda: self.create_data(__df))
        self.run_button.pack(pady=2)

        # Separator
        self.separator2 = ttk.Separator(self.window)
        self.separator2.pack(pady=6, fill='x')

        # Вывод служебной информации
        self.output_data_variable = tk.StringVar(value=message)
        self.output_data = ttk.Label(
            self.window, foreground=color, font='Calibri 16 bold', textvariable=self.output_data_variable)
        self.output_data.pack()

        self.window.mainloop()

    def show_date(self, event):
        """Записывает выбранные даты в переменные 'начало периода' и 'конец периода'"""
        if self.flag == 0:
            self.start_date_user_input.set(self.calendar.get_date())
            self.flag = 1
        elif self.flag == 1:
            self.end_date_user_input.set(self.calendar.get_date())
            self.flag = 0

    def create_data(self, __df):
        """Генерирует путевые листы."""
        
        vin = self.vin_user_input.get().strip()
        start_date = self.start_date_user_input.get()
        end_date = self.end_date_user_input.get()
        shift = self.shift_user_input.get()

        input_data = {vin: [start_date, end_date, shift]}

        if bool(vin) == False:
            self.output_data.configure(foreground='red')
            self.output_data_variable.set('Введите VIN!')
        elif not compare_dates(start_date, end_date):
            self.output_data.configure(foreground='red')
            self.output_data_variable.set('Проверьте даты!')
        else:
            self.output_data.configure(foreground='green')
            self.output_data_variable.set('Работаю...')


            df_filtered = filter_df(input_data, __df)
            is_generated = generate_waybills(df_filtered)
            if not is_generated:
                self.output_data.configure(foreground='red')
                self.output_data_variable.set(f'Нет данных.')
            elif is_generated:
                self.output_data.configure(foreground='green')
                self.output_data_variable.set(
                    f'Сгенерировано {len(df_filtered)} путевых листов.')
            else:
                self.output_data.configure(foreground='red')
                self.output_data_variable.set(f'Неизвестная ошибка.')


if __name__ == '__main__':

    if check_files():
        if not old_data():
           df = get_df_from_excel_file('waybill_data.xlsx')
           app = Application(df)
        else:
           df = get_df_from_excel_file('waybill_data.xlsx')
           app = Application(df, color='red', message='Данные устарели!')    
    else:
        print('Завершаю работу...')
        for i in range(5, 0, -1):
            time.sleep(1)
