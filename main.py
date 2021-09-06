import tkinter as tk
import tkinter.ttk as ttk
from tkinter import *
from tkinter.ttk import *


from tkinter import messagebox
from tkinter import Menu

from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg, NavigationToolbar2Tk)
# Implement the default Matplotlib key bindings.

import matplotlib.animation as animation
import matplotlib.pyplot as plt  # $ pip install matplotlib
from matplotlib import dates

import random
from collections import deque

import datetime


from pymodbus.constants import Endian
from pymodbus.payload import BinaryPayloadDecoder
from pymodbus.client.sync import ModbusSerialClient

import os
import xlsxwriter
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

class App(tk.Tk):

    def __init__(self):
        super(App, self).__init__()

        ###########################################################################
        self.temperatures = []

        self.npoints = 85000
        self.fig_1, self.ax = plt.subplots(figsize=(10, 6))
        # self.ax.set_xlabel('Время замера')
        self.ax.set_ylabel('Температура (С)')
        self.fig_1.suptitle(t='График температур')
        self.ax.grid()
        self.fig_1.subplots_adjust(left=0.07, right=0.974, top=0.945, bottom=0.1)
        self.canvas = FigureCanvasTkAgg(self.fig_1, master=self)  # A tk.DrawingArea.
        self.ani = None

        # pack_toolbar=False will make it easier to use a layout manager later on.
        # self.toolbar = NavigationToolbar2Tk(self.canvas, self, pack_toolbar=False)
        # self.toolbar.update()

        self.canvas.get_tk_widget().grid(row=0, column=0, columnspan=3)
        # self.toolbar.grid(row=1, column=0, columnspan=3)

        ###########################################################################
        # настройка внешнего вида окна (размещение полей с показаниями датчиков и параметрами замеров

        self.Parametr_frame = ttk.LabelFrame(self, text='Параметры замера')
        self.Parametr_frame.grid(row=2, column=0, columnspan=1)

        self.cur_temp = ttk.LabelFrame(self, text='Текущая температура')
        self.cur_temp.grid(row=2, column=1)

        self.time_frame = ttk.LabelFrame(self, text='Время замера')
        self.time_frame.grid(row=2, column=2)

        ttk.Label(self.Parametr_frame, text='Длина ребра кубического контейнера, мм').grid(row=0, column=0, sticky=tk.W)
        ttk.Label(self.Parametr_frame, text='Температура воздуха в печи, ℃').grid(row=1, column=0, sticky=tk.W)
        ttk.Label(self.Parametr_frame, text='Время замера, час').grid(row=2, column=0, sticky=tk.W)
        ttk.Label(self.Parametr_frame, text='Влажность образца, %').grid(row=3, column=0, sticky=tk.W)

        self.len_entry = Entry(self.Parametr_frame, width=5)
        self.len_entry.grid(row=0, column=1)
        self.temp_entry = Entry(self.Parametr_frame, width=5)
        self.temp_entry.grid(row=1, column=1)
        self.time_entry = Entry(self.Parametr_frame, width=5)
        self.time_entry.grid(row=2, column=1)
        self.den_entry = Entry(self.Parametr_frame, width=5)
        self.den_entry.grid(row=3, column=1)

        ttk.Label(self.cur_temp, text='Температура в печи, ℃').grid(row=0, column=0)
        ttk.Label(self.cur_temp, text='Датчик 1, ℃').grid(row=1, column=0)
        ttk.Label(self.cur_temp, text='Датчик 2, ℃').grid(row=2, column=0)
        ttk.Label(self.cur_temp, text='Датчик 3, ℃').grid(row=3, column=0)
        self.temp_1 = ttk.Label(self.cur_temp, justify='left')
        self.temp_1.grid(row=0, column=1)
        self.temp_2 = ttk.Label(self.cur_temp, justify='left')
        self.temp_2.grid(row=1, column=1)
        self.temp_3 = ttk.Label(self.cur_temp, justify='left')
        self.temp_3.grid(row=2, column=1)
        self.temp_4 = ttk.Label(self.cur_temp, justify='left')
        self.temp_4.grid(row=3, column=1)

        ttk.Label(self.time_frame, text='Текущее время', justify='left').grid(row=0, column=0, sticky=tk.W)
        ttk.Label(self.time_frame, text='Время начала\nэксперимента', justify='left').grid(row=1, column=0, sticky=tk.W)
        ttk.Label(self.time_frame, text='Время окончания эксперимента', justify='left').grid(row=2, column=0,
                                                                                             sticky=tk.W)
        ttk.Label(self.time_frame, text='Фактическое время\nокончания эксперимента', justify='left').grid(row=3,
                                                                                                          column=0,
                                                                                                          sticky=tk.W)
        self.cur_time_label = ttk.Label(self.time_frame, justify='left',
                                        text=datetime.datetime.now().time().strftime('%H:%M:%S'))
        self.cur_time_label.grid(row=0, column=1)
        self.start_time = ttk.Label(self.time_frame, justify='left')
        self.start_time.grid(row=1, column=1)
        self.fihish_time_label = ttk.Label(self.time_frame, justify='left')
        self.fihish_time_label.grid(row=2, column=1)
        self.real_finish_label = ttk.Label(self.time_frame, justify='left')
        self.real_finish_label.grid(row=3, column=1)

        self.protocol('WM_DELETE_WINDOW', self.on_closing)
        ###########################################################################
        # Указание начальных параметров для замера

        self.temperatures = [round(100 * random.random(), 1) for i in range(4)]
        self.Measuring = BooleanVar()
        self.Measuring.set(FALSE)
        self.Error = BooleanVar()
        self.Error.set(FALSE)
        self.start_button = ttk.Button(self, text='Начать эксперимент', command=self.start_measuring)
        self.start_button.grid(row=3, column=0)

        self.timer = 3000  # время опроса датчиков и обновления графиков в миллисекундах

        self.fmt = dates.DateFormatter('%d.%m.%Y %H:%M:%S')  # формат отображения даты измерения на графиках

        self.start = datetime.datetime.now()
        self.finish = datetime.datetime.now() + datetime.timedelta(hours=24)
        self.real_finish = datetime.datetime.now()

    ###########################################################################
    # Отключение от COM порта и закрытие окна

    def on_closing(self):
        if messagebox.askokcancel("Выход", "Вы действительно хотите выйти?"):
            try:
                client.write_registers(3, 0, unit=1)
            except:
                pass
            if self.Measuring.get():
                self.ani.event_source.stop()
                self.Measuring.set(FALSE)
                if messagebox.askyesno('Остановка эксперимента', 'Эксперимент не завершён. Создать отчёт?'):
                    self.stop_measuring()
            try:
                print('Отключение устройств...')
                client.close()
                print('Выполнено.')
            except:
                pass
            tk.Tk.quit(self)

    # Обновление показаний температур на экране

    def update_temperatures_on_plot(self, dy):
        self.dates.append(datetime.datetime.now())
        self.y_temp_1.append(self.temperatures[0])
        self.y_temp_2.append(self.temperatures[1])
        self.y_temp_3.append(self.temperatures[2])
        self.y_temp_4.append(self.temperatures[3])

        self.line_1.set_data(self.dates, self.y_temp_1)
        self.line_2.set_data(self.dates, self.y_temp_2)
        self.line_3.set_data(self.dates, self.y_temp_3)
        self.line_4.set_data(self.dates, self.y_temp_4)

        self.ax.relim()  # update axes limits
        # self.ax.figure.autofmt_xdate()
        self.ax.xaxis.set_major_formatter(self.fmt)
        self.fig_1.autofmt_xdate()
        self.ax.autoscale_view(True, True, True)

        self.real_finish = datetime.datetime.now()

        if max(self.temperatures) > self.max_temp:
            self.max_temp = max(self.temperatures)

        if self.real_finish >= self.finish:
            print('Эксперимент окончен успешно!')
            print('Время начала: ', self.start)
            print('Время окончания: ', self.real_finish)
            return self.stop_measuring()

        if any(self.temperatures[i] - self.temperatures[0] >= 60 for i in range(1, 4, 1)):
            self.Error.set(TRUE)
            return self.stop_measuring()

        return self.line_1, self.line_2, self.line_3, self.line_4, self.ax, self.y_temp_1, self.y_temp_2, \
               self.y_temp_3, self.y_temp_4

    # Начальные параметры для генерации графиков температур

    def init_temperatures(self):
        try:
            self.line_1.remove()
            self.line_2.remove()
            self.line_3.remove()
            self.line_4.remove()
        except:
            pass
        self.max_temp = 0
        self.start = datetime.datetime.now()
        self.finish = self.start + datetime.timedelta(hours=24)
        self.dates = deque([datetime.datetime.now()], maxlen=self.npoints)
        self.start_time.configure(text=self.start.strftime('%d/%m/%y %H:%M:%S'))
        self.fihish_time_label.configure(text=self.finish.strftime('%d/%m/%y %H:%M:%S'))

        self.ax.xaxis_date()
        self.fig_1.autofmt_xdate()
        self.ax.autoscale_view(True, True, True)
        self.y_temp_1 = deque([self.temperatures[0]], maxlen=self.npoints)
        self.y_temp_2 = deque([self.temperatures[1]], maxlen=self.npoints)
        self.y_temp_3 = deque([self.temperatures[2]], maxlen=self.npoints)
        self.y_temp_4 = deque([self.temperatures[3]], maxlen=self.npoints)
        [self.line_1] = self.ax.plot(self.dates, self.y_temp_1, label='Температура в печи')
        [self.line_2] = self.ax.plot(self.dates, self.y_temp_2, label='Датчик 1')
        [self.line_3] = self.ax.plot(self.dates, self.y_temp_3, label='Датчик 2')
        [self.line_4] = self.ax.plot(self.dates, self.y_temp_4, label='Датчик 3')
        self.ax.legend()
        return self.dates, self.line_1, self.line_2, self.line_3, self.line_4

    # Начало замера

    def start_measuring(self):
        if self.Measuring.get() == FALSE:
            if self.len_entry.get() == '' or self.time_entry.get() == '' \
                    or self.temp_entry.get() == '' or self.den_entry.get() == '':
                messagebox.showerror('Ошибка', 'Все параметры замера должны быть заполнены!')
                return
            if self.len_entry.get().isdigit() and self.temp_entry.get().isdigit() and \
                    self.time_entry.get().isdigit() and self.den_entry.get().isdigit():
                pass
            else:
                messagebox.showerror('Ошибка', 'Провертьте корректность введёных парамтеров')
                return
            params = f'Длина ребра: {self.len_entry.get()} мм\nТемпература воздуха в печи: {self.temp_entry.get()}' \
                     f' С\nВлажность воздуха: {self.den_entry.get()}%.'

            if not messagebox.askyesno("Подтвердите введёные параметры", params):
                return
            self.Measuring.set(TRUE)
            self.Error.set(FALSE)
            self.ani = animation.FuncAnimation(self.fig_1, self.update_temperatures_on_plot,
                                               init_func=self.init_temperatures,
                                               interval=self.timer, repeat=False)
            client.write_registers(3, 16256, unit=1)
            self.ani._start()
            self.real_finish_label.configure(text='')
            self.start_button.configure(text='Остановить эксперимент')
            self.len_entry.configure(state='readonly')
            self.temp_entry.configure(state='readonly')
            self.time_entry.configure(state='readonly')
            self.den_entry.configure(state='readonly')
            print('Эксперимент начат')
            self.start_button.configure(command=self.ask_to_stop)

    def ask_to_stop(self):
        if messagebox.askyesno('Прервать эксперимент', 'Вы действительно хотите прервать эксперимент?'):
            print('Эксперимент окончен вручную')
            return self.stop_measuring()

    # Остановка замера по превышению температуры или по нажатию кнопки

    def stop_measuring(self):
        client.write_registers(3, 0, unit=1)
        self.real_finish_label.configure(text=self.real_finish.strftime('%d/%m/%y %H:%M:%S'))
        self.ani.event_source.stop()
        self.Measuring.set(FALSE)
        self.create_report()
        self.len_entry.configure(state='NORMAL')
        self.temp_entry.configure(state='NORMAL')
        self.time_entry.configure(state='NORMAL')
        self.den_entry.configure(state='NORMAL')
        self.start_button.configure(text='Начать эксперимент')
        if self.Error.get():
            messagebox.showwarning('Остановка эксперимента', 'Произошло возгорание!')
        else:
            messagebox.showinfo('Остановка эксперимента', 'Эксперимент прошёл успешно!')
        self.start_button.configure(command=self.start_measuring)

    def update_temperatures(self):
        try:
            self.temperatures = [round(read_value(24 + i * 2, 2, 1), 1) for i in range(4)]
        except:
            return self.temperatures, self.after(int(self.timer / 10), self.update_temperatures)
        self.temp_1.configure(text=self.temperatures[0])
        self.temp_2.configure(text=self.temperatures[1])
        self.temp_3.configure(text=self.temperatures[2])
        self.temp_4.configure(text=self.temperatures[3])
        return self.temperatures, self.after(int(self.timer / 10), self.update_temperatures)

    def update_current_time(self):
        self.cur_time_label.configure(text=datetime.datetime.now().strftime('%d/%m/%y %H:%M:%S'))
        self.after(1000, self.update_current_time)

    def create_report(self):
        folder_name = self.start.strftime("%d_%m_%Y %H-%M-%S")
        os.mkdir(folder_name)
        self.fig_1.savefig(f'{folder_name}\\График температур.png')
        workbook = xlsxwriter.Workbook(f'{folder_name}\\Результат эксперимента.xlsx')
        worksheet = workbook.add_worksheet('Параметры замера')
        date_format = workbook.add_format({'num_format': 'hh:mm:ss'})
        expenses = (
            ['Длина ребра кубического контейнера, мм', self.len_entry.get()],
            ['Температура воздуха в печи, ℃', self.temp_entry.get()],
            ['Время замера, час', self.time_entry.get()],
            ['Влажность образца, %', self.den_entry.get()],
            ['Максимаяльная температура, ℃', self.max_temp]
        )
        row = 0
        col = 0

        # Iterate over the data and write it out row by row.
        for item, param in expenses:
            worksheet.write(row, col, item)
            worksheet.write_number(row, col + 1, float(param))
            row += 1
        if self.Error.get():
            worksheet.write(row, col, 'Результат эксперимента')
            worksheet.write(row, col+1, 'Произошло возгорание!')
        else:
            worksheet.write(row, col, 'Результат эксперимента')
            worksheet.write(row, col+1, 'Успешный')

        row = 0
        worksheet = workbook.add_worksheet('Результаты замеров')
        for i in range(len(self.y_temp_1)):
            worksheet.write(row, col, self.y_temp_1[i])
            worksheet.write(row, col + 1, self.y_temp_2[i])
            worksheet.write(row, col + 2, self.y_temp_3[i])
            worksheet.write(row, col + 3, self.y_temp_4[i])
            worksheet.write(row, col + 4, self.dates[i], date_format)
            row += 1
        workbook.close()

        doc = DocxTemplate('Шаблон.docx')
        myimage = InlineImage(doc, image_descriptor=f'{folder_name}\\График температур.png', width=Mm(175), height=Mm(105))
        context = {'test_date': self.start.strftime("%d/%m/%Y"),
                   'test_number': 1,
                    'len': str(self.len_entry.get()),
                   'temperature': str(self.temp_entry.get()),
                   'time': str(self.time_entry.get()),
                   'den': str(self.den_entry.get()),
                   'graph_image': myimage,
                   'max_temp': self.max_temp}
        if not self.Error.get():
            context.update({
                'success': 'отрицательный',
                'not_raised': ' не',
                'diff': 60
            })
        else:
            context.update({
                'success': 'положительный',
                'not_raised': '',
                'diff': self.max_temp-float(self.temp_entry.get())
            })
        doc.render(context)
        doc.save(f'{folder_name}\\Отчёт.docx')

    def create_child1(self):
        self.Child1Window = tk.Toplevel(master=self)
        self.heat_on_btn = ttk.Button(master=self.Child1Window, text='Включить печь', command=heat_on, width=20)
        self.heat_off_btn = ttk.Button(master=self.Child1Window, text='Отключить печь', command=heat_off, width=20)

        self.heat_on_btn.pack()
        self.heat_off_btn.pack()

        self.Child1Window.resizable(height=False, width=False)


###########################################################################

def heat_off():
    client.write_registers(3, 0, unit=1)
    print('Печь выключена')

def heat_on():
    client.write_registers(3, 16256, unit=1)
    print('Печь включена')

def read_value(adr, cnt=2, unt=1):
    result = client.read_holding_registers(address=adr, count=cnt, unit=unt)
    decoder = BinaryPayloadDecoder.fromRegisters(result.registers, Endian.Big, wordorder=Endian.Little)
    value = decoder.decode_32bit_float()
    return value


if __name__ == "__main__":
    client = ModbusSerialClient(
        method='rtu',
        port='COM3',
        baudrate=9600,
        timeout=3,
        parity='N',
        stopbits=1,
        bytesize=8
    )

    client.connect()

    root = App()
    root.title("Испытание по методике ООН")
    mainmenu = Menu(root)
    root.config(menu=mainmenu)
    filemenu = Menu(mainmenu, tearoff=0)
    filemenu.add_command(label='Управление печкой', command=root.create_child1)
    filemenu.add_command(label="Выход", command=root.on_closing)
    helpmenu = Menu(mainmenu, tearoff=0)
    helpmenu.add_command(label="Помощь")
    helpmenu.add_command(label="О программе")
    mainmenu.add_cascade(label="Меню", menu=filemenu)
    mainmenu.add_cascade(label="Справка", menu=helpmenu)

    ###########################################################################
    root.resizable(height=False, width=False)
    root.update_current_time()
    root.update_temperatures()
    root.mainloop()
