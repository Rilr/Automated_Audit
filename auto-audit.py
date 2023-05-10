import PySimpleGUI as sg
import pandas as pd
import datetime as dt

sg.theme('DarkAmber') 
layout = [[sg.Text('BMK Audit Excel')], #Automate CSV Upload 
    [sg.InputText(key="-audit-"), 
    sg.FileBrowse(file_types=[("Microsoft Excel Workbook","*.xlsx")])],
    [sg.Text('Automate CSV')], #Automate CSV Upload
    [sg.InputText(key="-automate-"), 
    sg.FileBrowse(file_types=[("CSV Files","*.csv")])], 
    [sg.Text('Webroot CSV(Not_Yet_Suported)')], #Webroot CSV Upload 
    [sg.InputText(key="-webroot-"), 
    sg.FileBrowse(file_types=[("CSV Files","*.csv")])], 
    [sg.Button("Submit"), sg.Cancel()]] #Submit and Cancel Button
window = sg.Window('AutoAutomate', layout)

while True:
    event, values = window.read()
    autoData, webrootData, auditData = values['-automate-'], values['-webroot-'], values['-audit-']
    if event in (sg.WIN_CLOSED, 'Cancel'): 
        break 
    elif event == "Submit":
        print(autoData)