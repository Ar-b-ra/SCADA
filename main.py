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

import random
from collections import deque

import pandas as pd
import numpy as np

import time

import OpenOPC
from getpass import getpass

from pymodbus.constants import Endian
from pymodbus.payload import BinaryPayloadDecoder
from pymodbus.payload import BinaryPayloadBuilder
from pymodbus.client.sync import ModbusSerialClient

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


def create_report():
    pass


class App(tk.Tk):

    def __init__(self):
        super(App, self).__init__()

        ###########################################################################
        self.npoints = 85000
        self.x = deque([0], maxlen=self.npoints)
        self.y = deque([0], maxlen=self.npoints)

        self.fig_1, self.ax = plt.subplots()
        [self.line] = self.ax.step(self.x, self.y)

        self.fig_1.suptitle(t='График температур')
        self.canvas = FigureCanvasTkAgg(self.fig_1, master=self)  # A tk.DrawingArea.
        self.ani = None

        # pack_toolbar=False will make it easier to use a layout manager later on.
        self.toolbar = NavigationToolbar2Tk(self.canvas, self, pack_toolbar=False)
        self.toolbar.update()

        self.canvas.get_tk_widget().grid(row=0, column=7, rowspan=15)
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

        self.Measuring = BooleanVar()
        self.Measuring.set(FALSE)
        self.Error = BooleanVar()
        self.Error.set(FALSE)
        self.start_button = ttk.Button(self, text='Начать эксперимент', command=self.start_measuring)
        self.start_button.grid(row=4, column=0)

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
            App.destroy()

    def update(self, dy):
        self.x.append(self.x[-1] + 1)  # update data
        self.y.append(self.y[-1] + dy)

        self.line.set_xdata(self.x)  # update plot data
        self.line.set_ydata(self.y)

        self.ax.relim()  # update axes limits
        self.ax.autoscale_view(True, True, True)
        return self.line, self.ax

    @property
    def data_gen(self):
        while True:
            self.temp_1.configure(text=round(read_value(adr=3), 1))
            self.temp_2.configure(text=round(read_value(adr=9), 1))
            self.temp_3.configure(text=round(read_value(adr=15), 1))
            self.temp_4.configure(text=round(read_value(adr=21), 1))
            yield 1 if random.random() < 0.5 else -1

    def start_measuring(self):
        if self.Measuring.get() == FALSE:
            self.Measuring.set(TRUE)
            self.Error.set(FALSE)
            self.ani = animation.FuncAnimation(self.fig_1, self.update, self.data_gen, interval=1000, repeat=False)
            self.ani._start()
            self.start_button.configure(text='Остановить замер')
            self.len_entry.configure(state='readonly')
            self.temp_entry.configure(state='readonly')
            self.time_entry.configure(state='readonly')
            self.den_entry.configure(state='readonly')
            print('Starting measuring')
        else:
            self.ani.event_source.stop()
            self.Measuring.set(FALSE)
            self.len_entry.configure(state='NORMAL')
            self.temp_entry.configure(state='NORMAL')
            self.time_entry.configure(state='NORMAL')
            self.den_entry.configure(state='NORMAL')
            self.start_button.configure(text='Начать замер')
            print('Stop measuring')


###########################################################################

def read_value(adr, cnt=2, unt=1):
    result = client.read_holding_registers(address=adr, count=cnt, unit=unt)
    decoder = BinaryPayloadDecoder.fromRegisters(result.registers, Endian.Big, wordorder=Endian.Little)
    value = decoder.decode_32bit_float()
    return value


if __name__ == "__main__":
    #client = connect_opc()
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

    root.mainloop()
