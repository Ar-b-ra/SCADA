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

import time

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
        self.npoints = 85000

        self.fig_1, self.ax = plt.subplots(figsize=(10, 6))
        self.ax.set_xlabel('Время замера')
        self.ax.set_ylabel('Температура (С)')
        self.fig_1.suptitle(t='График температур')
        #self.fig_1.subplots_adjust(left=0.1, right=0.974, top=0.9, bottom=0.1)
        self.ax.xaxis_date()
        self.canvas = FigureCanvasTkAgg(self.fig_1, master=self)  # A tk.DrawingArea.
        self.ani = None

        # pack_toolbar=False will make it easier to use a layout manager later on.
        self.toolbar = NavigationToolbar2Tk(self.canvas, self, pack_toolbar=False)
        self.toolbar.update()

        self.canvas.get_tk_widget().grid(row=0, column=7, rowspan=16, columnspan=2)
        self.toolbar.grid(row=16, column=7)

        ###########################################################################

        self.Parametr_frame = ttk.LabelFrame(self, text='Параметры замера')
        self.Parametr_frame.grid(row=0, column=0, columnspan=6)

        self.cur_temp = ttk.LabelFrame(self, text='Текущая температура')
        self.cur_temp.grid(row=1, column=0)

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

        ttk.Label(self.cur_temp, text='Датчик 1, ℃').grid(row=0, column=0)
        ttk.Label(self.cur_temp, text='Датчик 2, ℃').grid(row=1, column=0)
        ttk.Label(self.cur_temp, text='Датчик 3, ℃').grid(row=2, column=0)
        ttk.Label(self.cur_temp, text='Датчик 4, ℃').grid(row=3, column=0)
        self.temp_1 = ttk.Label(self.cur_temp, justify='left')
        self.temp_1.grid(row=0, column=1)
        self.temp_2 = ttk.Label(self.cur_temp, justify='left')
        self.temp_2.grid(row=1, column=1)
        self.temp_3 = ttk.Label(self.cur_temp, justify='left')
        self.temp_3.grid(row=2, column=1)
        self.temp_4 = ttk.Label(self.cur_temp, justify='left')
        self.temp_4.grid(row=3, column=1)

        self.protocol('WM_DELETE_WINDOW', self.on_closing)
        ###########################################################################

        global temperatures
        temperatures = [round(100 * random.random(), 1) for i in range(4)]
        self.Measuring = BooleanVar()
        self.Measuring.set(FALSE)
        self.Error = BooleanVar()
        self.Error.set(FALSE)
        self.start_button = ttk.Button(self, text='Начать эксперимент', command=self.start_measuring)
        self.start_button.grid(row=4, column=0)

        self.timer = 1000

        self.fmt = dates.DateFormatter('%Y-%m-%d %H:%M:%S')

        self.update_temperatures()

    ###########################################################################

    def on_closing(self):
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            try:
                print('Отключение от устройства опроса температуры...')
                client.close()
                print('Done.')
            except:
                pass
            try:
                print('MySQL server disconnecting...')
                # cnx.disconnect()
                print('Done.')
            except:
                pass
            self.destroy()

    def update_temperatures_on_plot(self, dy):
        self.x.append(self.x[-1] + 1)  # update data
        #self.dates.append(datetime.datetime.now().time())
        self.dates.append(datetime.datetime.now())
        self.y_temp_1.append(float(self.temp_1.cget('text')))
        self.y_temp_2.append(float(self.temp_2.cget('text')))
        self.y_temp_3.append(float(self.temp_3.cget('text')))
        self.y_temp_4.append(float(self.temp_4.cget('text')))

        self.line_1.set_data(self.dates, self.y_temp_1)
        self.line_2.set_data(self.dates, self.y_temp_2)
        self.line_3.set_data(self.dates, self.y_temp_3)
        self.line_4.set_data(self.dates, self.y_temp_4)

        self.ax.relim()  # update axes limits
        # self.ax.figure.autofmt_xdate()
        self.ax.xaxis.set_major_formatter(self.fmt)
        self.fig_1.autofmt_xdate()
        self.ax.autoscale_view(True, True, True)
        return self.line_1, self.line_2, self.line_3, self.line_4, self.ax, self.y_temp_1, self.y_temp_2, self.y_temp_3, self.y_temp_4

    def init_temperatures(self):
          # clear your figure
        #self.fig_1.clear()  # redraw your canvas so it becomes empty
        self.dates = deque([datetime.datetime.now()], maxlen=self.npoints)
        print(self.dates)
        self.x = deque([0], maxlen=self.npoints)
        self.y_temp_1 = deque([self.temp_1.cget('text')], maxlen=self.npoints)
        self.y_temp_2 = deque([self.temp_2.cget('text')], maxlen=self.npoints)
        self.y_temp_3 = deque([self.temp_3.cget('text')], maxlen=self.npoints)
        self.y_temp_4 = deque([self.temp_4.cget('text')], maxlen=self.npoints)
        [self.line_1] = self.ax.plot(self.dates, self.y_temp_1, label='Температура в печи')
        [self.line_2] = self.ax.plot(self.dates, self.y_temp_2, label='Датчик 1')
        [self.line_3] = self.ax.plot(self.dates, self.y_temp_3, label='Датчик 2')
        [self.line_4] = self.ax.plot(self.dates, self.y_temp_4, label='Датчик 3')
        self.ax.legend()
        return self.x, self.line_1, self.line_2, self.line_3, self.line_4

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
            self.Measuring.set(TRUE)
            self.Error.set(FALSE)
            self.ani = animation.FuncAnimation(self.fig_1, self.update_temperatures_on_plot, init_func=self.init_temperatures,
                                               interval=self.timer, repeat=False)
            self.ani._start()
            self.start_button.configure(text='Остановить эксперимент')
            self.len_entry.configure(state='readonly')
            self.temp_entry.configure(state='readonly')
            self.time_entry.configure(state='readonly')
            self.den_entry.configure(state='readonly')
            print('Starting measuring')
        else:
            self.ani.event_source.stop()
            self.Measuring.set(FALSE)
            self.create_report()
            self.len_entry.configure(state='NORMAL')
            self.temp_entry.configure(state='NORMAL')
            self.time_entry.configure(state='NORMAL')
            self.den_entry.configure(state='NORMAL')
            self.start_button.configure(text='Начать эксперимент')
            print('Stop measuring')

    def update_temperatures(self):
        temperatures = [round(100 * random.random(), 1) for i in range(4)]
        self.temp_1.configure(text=temperatures[0])
        self.temp_2.configure(text=temperatures[1])
        self.temp_3.configure(text=temperatures[2])
        self.temp_4.configure(text=temperatures[3])
        self.after(self.timer, self.update_temperatures)

    def create_report(self):
        folder_name = datetime.datetime.now().strftime("%d_%m_%Y %H-%M-%S")
        os.mkdir(folder_name)
        workbook = xlsxwriter.Workbook(f'{folder_name}\Результат эксперимента.xlsx')
        worksheet = workbook.add_worksheet()
        date_format = workbook.add_format({'num_format': 'd mmmm yyyy h:mm:ss'})
        expenses = (
            ['Длина ребра кубического контейнера, мм', self.len_entry.get()],
            ['Температура воздуха в печи, ℃', self.temp_entry.get()],
            ['Время замера, час', self.time_entry.get()],
            ['Влажность образца, %', self.den_entry.get()]
        )
        row = 0
        col = 0

        # Iterate over the data and write it out row by row.
        for item, parametr in (expenses):
            worksheet.write(row, col, item)
            worksheet.write(row, col + 1, int(parametr))
            row += 1

        for i in range(len(self.y_temp_1)):
            worksheet.write(row, col, self.y_temp_1[i])
            worksheet.write(row, col + 1, self.y_temp_2[i])
            worksheet.write(row, col + 2, self.y_temp_3[i])
            worksheet.write(row, col + 3, self.y_temp_4[i])
            worksheet.write_datetime(row, col + 4, self.dates[i], date_format)
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
    root.mainloop()
