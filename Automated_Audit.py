import PySimpleGUI as sg
import pandas as pd
from datetime import datetime, timedelta
from dateutil import parser    

#-----define UI-----#
sg.theme('DarkAmber')
layout = [[sg.Text('Automate CSV')],                #Automate CSV Upload
          [sg.InputText(key="-automate-"),
           sg.FileBrowse(file_types=[("CSV Files","*.csv")])],
          [sg.Text('Webroot CSV(Not_Yet_Suported)')],   #Webroot CSV Upload
          [sg.InputText(key="-webroot-"),
           sg.FileBrowse(file_types=[("CSV Files","*.csv")])],
          [sg.Button("Submit"), sg.Cancel()]]       #Submit and Cancel Button
window = sg.Window('AutoAutomate', layout)

#-----define time variables-----#
current_date = datetime.now()
one_month_ago = current_date - timedelta(days=30)

#-----what to do with the imported data-----#
while True:
    event, values = window.read() 
    automateData, webrootData = values['-automate-'], values['-webroot-'] #Passing the CSV dataframes into variables, f2 is not used
    if event in (sg.WIN_CLOSED, 'Cancel'):
        break
    elif event == "Submit":
        df = pd.read_csv(automateData, header=0)          #Read the csv
        #date_format = '%Y-%m-%d %H:%M:%S.%f%f%f' ###Old variable, might want to reuse for datetime.strptime()###
        for index, row in df.iterrows():        #Iterate through the rows
            date_str = row[3]
            deviceName = row[1]
            date_obj = parser.parse(date_str)   #Covert date into readable format
            if date_obj < one_month_ago:
                print(deviceName)               #Eventually plan to pass through to it's own CSV, highlighting problem configs and their issues
window.close()