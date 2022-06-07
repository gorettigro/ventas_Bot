from importlib.resources import path
from socket import timeout
from turtle import tilt, title
from unicodedata import name
import pywinauto
import win32clipboard
import json
import os
from time import sleep, time
from pywinauto import application, keyboard

app = application.Application()
app.connect(path=r"C:\Program Files (x86)\Aspel\Aspel-SAE 8.0\SAEWIN80.exe")
main=app.window(title_re='.*Aspel-SAE.*')
lg=app.window(title='Abrir empresa')

with open("ventasData.json", "r", encoding="utf-8") as file: 
    data = json.load(file)
    
if lg.exists(timeout=5):
    app['Abrir empresa']['Edit5'].type_keys(data['user_name'])
    app['Abrir empresa']['Edit7'].type_keys(data['password'])
    app['Abrir empresa']['Button3'].click()

main.set_focus()
keyboard.send_keys('%v')
sleep(3)
keyboard.send_keys('%p1')

ven=main.window(title="Pedidos")
keyboard.send_keys('^e')
exp = app.window(title="Exportar información")
#exp.print_control_identifiers()
if exp.exists(timeout=2):
        app['Exportar información']['ComboBox2'].type_keys("%{DOWN}")
        app['Exportar información']['ComboBox2'].type_keys("te")
        app['Exportar información']['ComboBox2'].click()
        app['Exportar Información']['Edit7'].type_keys(data['repositorio'])
        app['Exportar información']['Button3'].click()

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

main.set_focus()

keyboard.send_keys("+^F")

ven2=main.window(title="Facturas")

keyboard.send_keys('^e')

expo = app.window(title="Exportar información")
if expo.exists(timeout=2):
        app['Exportar información']['ComboBox2'].type_keys("%{DOWN}")
        app['Exportar información']['ComboBox2'].type_keys("te")
        app['Exportar información']['ComboBox2'].click()
        app['Exportar Información']['Edit7'].type_keys(data['repositorio'])
        app['Exportar información']['Button3'].click()

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
