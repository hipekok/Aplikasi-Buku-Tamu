from re import T
import PySimpleGUI as sg
import pandas as pd
import tkinter as tk
from ttkthemes import ThemedTk

sg.theme('Kayak')

EXCEL_FILE = 'Pemesanan.xlsx'

df = pd.read_excel(EXCEL_FILE)

layout=[
[sg.Text('Masukan Data Pelanggan: ')],
[sg.Text('Nama',size=(15,1)), sg.InputText(key='Nama')],
[sg.Text('NIK',size=(15,1)), sg.InputText(key='NIK')],
[sg.Text('No. Telp',size=(15,1)), sg.InputText(key='Tlp')],
[sg.Text('Tgl Resevasi',size=(15,1)), sg.InputText(key='Tgl Resevasi'), sg.CalendarButton(' ', target='Tgl Resevasi', format=('%d-%m-%y'))],
[sg.Text('Tgl Checkout',size=(15,1)), sg.InputText(key='Tgl Checkout'), sg.CalendarButton(' ', target='Tgl Checkout', format=('%d-%m-%y'))],
[sg.Text('Jenis Kelamin',size=(15,1)), sg.Combo(['pria','wanita'],key='Jenis Kelamin')],
                            
[sg.Submit(), sg.Button('Clear', key='Clear'), sg.Exit('EXIT')]

]

window=sg.Window('Form Pemesanan',layout)

def create_calendar(root):
    pass   

def clear_input():
    for key in values.keys():
        window[key].update('')

    

while True :
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'EXIT':
        break
    elif event == 'Clear':
        clear_input()
    elif event == 'Submit':
        df =df._append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data Berhasil Di Simpan')
        clear_input()
window.close()