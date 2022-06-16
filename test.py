from importlib.resources import path
from socket import timeout
from turtle import tilt, title
from unicodedata import name
import pywinauto
import win32clipboard
import json
import codecs
import os
from time import sleep, time
from pywinauto import keyboard
from pywinauto.application import Application

#app = Application().start('SAEWIN80.exe')
app = Application(backend="win32").connect(path=r"C:\Program Files (x86)\Aspel\Aspel-SAE 8.0\SAEWIN80.exe")
main=app.window(title_re='.*Aspel-SAE.*')
lg=app.window(title='Abrir empresa')

with codecs.open("ventasData.json", "r", encoding="utf-8-sig") as file: 
    data = json.load(file)
    
if lg.exists(timeout=5):
    app['Abrir empresa']['Edit5'].type_keys(data['user_name'])
    app['Abrir empresa']['Edit7'].type_keys(data['password'])
    app['Abrir empresa']['Button3'].click()

main.set_focus()
keyboard.send_keys('%v')
sleep(3)
keyboard.send_keys('%p1')
sleep(3)

ven=main.window(title="Pedidos")
keyboard.send_keys('{F5}')
sleep(3)

filtrop=app.window(title="Filtro de pedidos")
if filtrop.exists(timeout=2):
        app['Filtro de pedidos']['ComboBox3'].type_keys("%{DOWN}")
        app['Filtro de pedidos']['ComboBox4'].type_keys(data['filtro_pedidos'])
        app['Filtro de pedidos']['ComboBox4'].click_input()
        app['Filtro de pedidos']['Button21'].click()

sleep(3)
keyboard.send_keys('^e')
exp = app.window(title="Exportar información")
if exp.exists(timeout=2):
        app['Exportar información']['ComboBox2'].type_keys("%{DOWN}")
        app['Exportar información']['ComboBox2'].type_keys("te")
        app['Exportar información']['ComboBox2'].click()
        app['Exportar Información']['Edit7'].type_keys(data['repositorio'])
        app['Exportar información']['Button3'].click()
        
sleep(3)

error = app.window(title="Error")
    
if error.exists(timeout=5):
    app['Error']['Button'].click()
    app['Exportar Información']['Edit7'].type_keys("C:/Users/auditor/Desktop")
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

sleep(3)
keyboard.send_keys('^e')

expo = app.window(title="Exportar información")
if expo.exists(timeout=2):
        app['Exportar información']['ComboBox2'].type_keys("%{DOWN}")
        app['Exportar información']['ComboBox2'].type_keys("te")
        app['Exportar información']['ComboBox2'].click()
        app['Exportar Información']['Edit7'].type_keys(data['repositorio'])
        app['Exportar información']['Button3'].click()
sleep(3)

error2 = app.window(title="Error")
    
if error2.exists(timeout=5):
    app['Error']['Button'].click()
    app['Exportar Información']['Edit7'].type_keys("C:/Users/auditor/Desktop")
    app['Exportar información']['Button3'].click()

confi=app.window(title="Confirmación")
if confi.exists(timeout=2):
    app['Confirmación']['Button1'].click()

info=app.window(title="Información")
if info.exists(timeout=2):
    app['Información']['Button'].click()

main.set_focus()
keyboard.send_keys('%{F4}')

confi2=app.window(title="Confirmación")
if confi2.exists(timeout=2):
        app['Confirmación']['Button1'].click()