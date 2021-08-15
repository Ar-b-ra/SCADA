# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import *
from tkinter.ttk import *
# from mysql.connecor import errorcode
import mysql.connector

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

import pandas as pd

import datetime

import OpenOPC
from getpass import getpass

from pymodbus.constants import Endian
from pymodbus.payload import BinaryPayloadDecoder
from pymodbus.client.sync import ModbusSerialClient

import os
import xlsxwriter


def connect_opc():
    opc = OpenOPC.client()
    servers = opc.servers()
    count = 0
    for server in servers:
        print(count, ':', server)
        count += 1
    opc_connect = False
    while not opc_connect:
        try:
            choice = int(input('Выберите номер сервера: '))
            print('Выполняется подключение к ', servers[choice])
            opc.connect(servers[choice])
            opc_connect = True
        except:
            print("Ошибка подключения. Попробуйте снова")
    print(f'Подключение к серверу {servers[choice]} произошло успешно!')
    return opc


def connect_db():
    try:
        cnx = mysql.connector.connect(user=input("user: "),
                                      password=getpass(prompt="Password: "),
                                      host=input("host: "),
                                      auth_plugin='mysql_native_password')
        print("Connected")
        return cnx
    except mysql.connector.Error as err:
        if err.errno == mysql.connector.errorcode.ER_ACCESS_DENIED_ERROR:
            print("Something is wrong with your user name or password")
            return connect_db()
        else:
            print(err)
            return connect_db()


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
        self.cur_time_laber = ttk.Label(self.time_frame, justify='left',
                                        text=datetime.datetime.now().time().strftime('%H:%M:%S'))
        self.cur_time_laber.grid(row=0, column=1)
        self.start_time = ttk.Label(self.time_frame, justify='left')
        self.start_time.grid(row=1, column=1)
        self.fihish_time_label = ttk.Label(self.time_frame, justify='left')
        self.fihish_time_label.grid(row=2, column=1)
        self.real_finish_label = ttk.Label(self.time_frame, justify='left')
        self.real_finish_label.grid(row=3, column=1)

        self.protocol('WM_DELETE_WINDOW', self.on_closing)
        ###########################################################################

        self.temperatures = [round(100 * random.random(), 1) for i in range(4)]
        self.Measuring = BooleanVar()
        self.Measuring.set(FALSE)
        self.Error = BooleanVar()
        self.Error.set(FALSE)
        self.start_button = ttk.Button(self, text='Начать эксперимент', command=self.start_measuring)
        self.start_button.grid(row=3, column=0)

        self.timer = 300  # время опроса датчиков и обновления графиков в миллисекундах

        self.fmt = dates.DateFormatter('%d.%m.%Y %H:%M:%S')  # формат отображения даты измерения на графиках

        self.start = datetime.datetime.now()
        self.finish = datetime.datetime.now() + datetime.timedelta(hours=24)
        self.real_finish = datetime.datetime.now()

    ###########################################################################

    def on_closing(self):
        if messagebox.askokcancel("Выход", "Вы действительно хотите выйти?"):
            try:
                print('Отключение устройств...')
                client.close()
                print('Выполнено.')
            except:
                pass
            # try:
            #     print('MySQL server disconnecting...')
            #     # cnx.disconnect()
            #     print('Done.')
            # except:
            #     pass
            tk.Tk.quit(self)

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

        if self.real_finish >= self.finish:

            print('Эксперимент окончен успешно!')
            print('Время начала: ', self.start)
            print('Время окончания: ', self.real_finish)
            return self.stop_measuring()

        return self.line_1, self.line_2, self.line_3, self.line_4, self.ax, self.y_temp_1, self.y_temp_2, self.y_temp_3, self.y_temp_4

    def init_temperatures(self):
        try:
            self.line_1.remove()
            self.line_2.remove()
            self.line_3.remove()
            self.line_4.remove()
        except:
            pass
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
            params = f'Длина ребра: {self.len_entry.get()} мм\nТемпература воздуха в печи: {self.temp_entry.get()} С\nВлажность воздуха: {self.den_entry.get()}%.'

            if not messagebox.askyesno("Подтвердите введёные параметры", params):
                return
            self.Measuring.set(TRUE)
            self.Error.set(FALSE)
            self.ani = animation.FuncAnimation(self.fig_1, self.update_temperatures_on_plot,
                                               init_func=self.init_temperatures,
                                               interval=self.timer, repeat=False)
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
            return self.stop_measuring()

    def stop_measuring(self):
        self.real_finish_label.configure(text=self.real_finish.strftime('%d/%m/%y %H:%M:%S'))
        self.ani.event_source.stop()
        self.Measuring.set(FALSE)
        self.create_report()
        self.len_entry.configure(state='NORMAL')
        self.temp_entry.configure(state='NORMAL')
        self.time_entry.configure(state='NORMAL')
        self.den_entry.configure(state='NORMAL')
        self.start_button.configure(text='Начать эксперимент')
        print('Эксперимент закончен')
        self.start_button.configure(command=self.start_measuring)

    def update_temperatures(self):
        self.temperatures = [round(100 * random.random(), 1) for i in range(4)]
        self.temp_1.configure(text=self.temperatures[0])
        self.temp_2.configure(text=self.temperatures[1])
        self.temp_3.configure(text=self.temperatures[2])
        self.temp_4.configure(text=self.temperatures[3])
        return self.temperatures, self.after(self.timer, self.update_temperatures)

    def update_current_time(self):
        self.cur_time_laber.configure(text=datetime.datetime.now().strftime('%d/%m/%y %H:%M:%S'))
        self.after(1000, self.update_current_time)

    def create_report(self):
        folder_name = datetime.datetime.now().strftime("%d_%m_%Y %H-%M-%S")
        os.mkdir(folder_name)
        workbook = xlsxwriter.Workbook(f'{folder_name}\\Результат эксперимента.xlsx')
        worksheet = workbook.add_worksheet('Параметры замера')
        date_format = workbook.add_format({'num_format': 'hh:mm:ss'})
        expenses = (
            ['Длина ребра кубического контейнера, мм', self.len_entry.get()],
            ['Температура воздуха в печи, ℃', self.temp_entry.get()],
            ['Время замера, час', self.time_entry.get()],
            ['Влажность образца, %', self.den_entry.get()]
        )
        row = 0
        col = 0

        # Iterate over the data and write it out row by row.
        for item, parametr in expenses:
            worksheet.write(row, col, item)
            worksheet.write(row, col + 1, int(parametr))
            row += 1

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


###########################################################################

def read_value(adr, cnt=2, unt=1):
    result = client.read_holding_registers(address=adr, count=cnt, unit=unt)
    decoder = BinaryPayloadDecoder.fromRegisters(result.registers, Endian.Big, wordorder=Endian.Little)
    value = decoder.decode_32bit_float()
    return value


if __name__ == "__main__":
    # client = connect_opc()
    # cnx = connect_db()
    # cursor = cnx.cursor(buffered=True)
    # databases = "show databases"
    # cursor.execute(databases)
    # count = 0
    # databases = list(cursor)
    # for i in databases:
    #     print(count, i[0])
    #     count += 1
    # database = databases[int(input("Select database: "))][0]
    # cursor.execute(f"USE {database}")
    # cnx.commit()
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
    filemenu.add_command(label='Создать отчёт')
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
