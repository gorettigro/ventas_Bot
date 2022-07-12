import ctypes
import sys
import pywinauto
import win32clipboard
import json
import codecs
import psutil

from concurrent.futures import process
from importlib.resources import path
from socket import timeout
from turtle import tilt, title
from unicodedata import name
from time import sleep, time
from pywinauto import keyboard
from pywinauto.application import Application


PROCNAME = "SAEWIN80.exe"

pid = None

for proc in psutil.process_iter():
    if proc.name() == PROCNAME:
        pid=proc.pid

if not pid:
    print("No se ha encontrado el proceso")
    ctypes.windll.user32.MessageBoxW(0,"No se ha encontrado el proceso","Error",1)
    sys.exit(0)

#app = Application().start('SAEWIN80.exe')
app = pywinauto.Application(backend="win32").connect(process=pid)
main=app.window(title_re='.*Aspel.*')
lg=app.window(title='Abrir empresa')

with codecs.open("ventasData.json", "r", encoding="utf-8-sig") as file: 
    data = json.load(file)
    print(data)
    
if lg.exists(timeout=5):
    app['Abrir empresa']['Edit5'].type_keys(data['user_name'])
    app['Abrir empresa']['Edit7'].type_keys(data['password'])
    app['Abrir empresa']['Button3'].click()

if main.exists(timeout=5 and not lg.exists(timeout=5)):
    main.set_focus()
else:
    for i in range(25):
        try:
            app.set_focus()
            keyboard.send_keys('%x')
            keyboard.send_keys('%a')

            if lg.exists(timeout=5):
                app['Abrir empresa']['Edit5'].type_keys(data['user_name'])
                app['Abrir empresa']['Edit7'].type_keys(data['password'])
                app['Abrir empresa']['Button3'].click()
            
            main.set_focus()
            break

        except Exception as e:
            pass

sleep(6)
keyboard.send_keys('%v')
sleep(3)
keyboard.send_keys('%p1')
sleep(5)

ven=main.window(title="Pedidos")
keyboard.send_keys('{F5}')
sleep(3)

filtrop=app.window(title="Filtro de pedidos")
if filtrop.exists(timeout=2):
        app['Filtro de pedidos']['ComboBox3'].type_keys("%{DOWN}")
        app['Filtro de pedidos']['ComboBox4'].type_keys(data['filtro_pedidos'])
        app['Filtro de pedidos']['ComboBox4'].click_input()
        app['Filtro de pedidos']['Button21'].click()
        print("Pedido Found")
else:
    print("Pedido not Found")

sleep(3)

keyboard.send_keys('^e')
exp = app.window(title="Exportar información")
if exp.exists(timeout=2):
        app['Exportar información']['ComboBox2'].type_keys("%{DOWN}")
        app['Exportar información']['ComboBox2'].type_keys(data['formato'])
        app['Exportar información']['ComboBox2'].click()
        app['Exportar Información']['Edit7'].type_keys(data['repositorio'], with_spaces=True)
        sleep(3)
        app['Exportar información']['Button3'].click()
        
sleep(6)

error = app.window(title="Error")
    
if error.exists(timeout=5):
    app['Error']['Button'].click()
    app['Exportar Información']['Edit7'].type_keys(data['repositorio_pre'], with_spaces=True)
    app['Exportar información']['Button3'].click()

confi=app.window(title="Confirmación")
if confi.exists(timeout=2):
    app['Confirmación']['Button1'].click()

info=app.window(title="Información")
if info.exists(timeout=2):
    app['Información']['Button'].click()

sleep(5)

main.set_focus()

keyboard.send_keys("+^F")
sleep(3)

ven2=main.window(title="Facturas")
keyboard.send_keys('{F5}')
sleep(3)

filtrov=app.window(title="Filtro de facturas")
if filtrov.exists(timeout=2):
        app['Filtro de facturas']['ComboBox3'].type_keys("%{DOWN}")
        app['Filtro de facturas']['ComboBox4'].type_keys(data['filtro_facturas'])
        app['Filtro de facturas']['ComboBox4'].click_input()
        app['Filtro de facturas']['Button22'].click()
        print("Factura Found")
else:
    print("Factura Not Found")

sleep(3)

keyboard.send_keys('^e')

expo2 = app.window(title="Exportar información")
if expo2.exists(timeout=2):
        app['Exportar información']['ComboBox2'].type_keys("%{DOWN}")
        app['Exportar información']['ComboBox2'].type_keys(data['formato'])
        app['Exportar información']['ComboBox2'].click()
        app['Exportar Información']['Edit7'].type_keys(data['repositorio'], with_spaces=True)
        sleep(3)
        app['Exportar información']['Button3'].click()
sleep(6)

error2 = app.window(title="Error")
    
if error2.exists(timeout=5):
    app['Error']['Button'].click()
    app['Exportar Información']['Edit7'].type_keys(data['repositorio_pre'], with_spaces=True)
    app['Exportar información']['Button3'].click()

confi2=app.window(title="Confirmación")
if confi2.exists(timeout=2):
    app['Confirmación']['Button1'].click()

sleep(15)

info2=app.window(title="Información")
if info2.exists(timeout=2):
    app['Información']['Button1'].click()

if main.exists(timeout=2):
    main.set_focus()
else:
    for i in range(25):
        try:
            app.set_focus()
            keyboard.send_keys('^%e')
            main.set_focus()
            break

        except Exception as e:
            pass

#main.set_focus()

keyboard.send_keys('%{F4}')

confi3=app.window(title="Confirmación")
if confi3.exists(timeout=2):
    app['Confirmación']['Button1'].click()