import PySimpleGUI as sg
import pandas as pd
from datetime import datetime, timedelta

sg.theme('DarkAmber')
layout = [[sg.Text('Audit XLSX')],
          [sg.InputText(key='fpaudit'),
           sg.FileBrowse(file_types=[("xlsx Files","*.xlsx")])],
          [sg.Text('Automate XLSX')],
          [sg.InputText(key="fpauto"),
           sg.FileBrowse(file_types=[("xlsx Files","*.xlsx")])],
          [sg.Text('Intune CSV')],
          [sg.InputText(key="fpintune"),
           sg.FileBrowse(file_types=[("CSV Files","*.csv")])],
          [sg.Text('Webroot CSV')],
          [sg.InputText(key="fpwebroot"),
           sg.FileBrowse(file_types=[("CSV Files","*.csv")])],
          [sg.Button("Submit"),sg.Cancel()]]
window = sg.Window('AutoAutomate', layout)

while True:
    event, values = window.read()
    audit_file_path = values['fpaudit']
    auto_file_path = values['fpauto']
    intune_file_path = values['fpintune']
    webroot_file_path = values['fpwebroot']
    audit_df = pd.read_excel(audit_file_path, sheet_name=0, header=2, usecols=[4], engine='openpyxl')
    auto_df, intune_df, webroot_df = [], [], []
    auto_audit_out = r'C:\scripting\Output\audit-discrepancies.xlsx'
    
    if event in (sg.WIN_CLOSED,'Cancel'):
        break
    
    elif event == "Submit":        
        if bool(auto_file_path) == True:
        # Automate difference checker
            def diffAuto(audit_df, auto_df):
                audit_df = audit_df.dropna(how='all')
                auto_df = pd.read_excel(auto_file_path, header=0, usecols=[0], engine='openpyxl')
                diffAuto_df = pd.merge(audit_df, auto_df, left_on='Configuration Name', right_on='Name', how='outer', suffixes=['_audit', '_auto'], indicator=True)
                diffAuto_df = diffAuto_df.rename(columns={'_merge': 'Found In'})
                diffAuto_df['Found In'] = diffAuto_df['Found In'].replace({'left_only': 'Only in Audit', 'right_only': 'Only in Automate'})
                return diffAuto_df
        elif bool(audit_df) == False:
            sg.popup_error("No Audit XLSX File Uploaded!")
            break
        
        elif bool(intune_file_path) == True:
            # Intune difference checker
            def diffIntune(audit_df, intune_df):
                audit_df = audit_df.dropna(how='all')
                intune_df = pd.read_csv(intune_file_path, header=0, usecols=[1])
                diffIntune_df = pd.merge(audit_df, intune_df, left_on='Configuration Name', right_on='Device name', how='outer', suffixes=['_audit', '_intune'], indicator=True)
                diffIntune_df = diffIntune_df.rename(columns={'_merge': 'Found In'})
                diffIntune_df['Found In'] = diffIntune_df['Found In'].replace({'left_only': 'Only in Audit', 'right_only': 'Only in Intune'})
                return diffIntune_df
        elif bool(False):
            continue
        
        elif bool(webroot_file_path) == True:
        # Webroot difference checker
            def diffWebroot(audit_df, webroot_df):
                audit_df = audit_df.dropna(how='all')
                webroot_df = pd.read_csv(webroot_file_path, header=0, usecols=[0])
                diffWebroot_df = pd.merge(audit_df, webroot_df, left_on='Configuration Name', right_on='Hostname', how='outer', suffixes=['_audit', '_intune'], indicator=True)
                diffWebroot_df = diffWebroot_df.rename(columns={'_merge': 'Found In'})
                diffWebroot_df['Found In'] = diffWebroot_df['Found In'].replace({'left_only': 'Only in Audit', 'right_only': 'Only in Webroot'})
                return diffWebroot_df
        elif bool(False):
            continue

        elif bool(auto_file_path) == True:
        # Automate date checker
            def dateAuto(auto_df):
                auto_df = pd.read_excel(auto_file_path, header=0, engine='openpyxl')
                two_weeks_ago = datetime.now() - timedelta(days=14)
                datemask = auto_df['Last Contact'] < two_weeks_ago
                dateAuto_df = auto_df.loc[datemask]
                return dateAuto_df
        elif bool(auto_file_path) == False:
            print('No Auto DF for Dates!')
            continue
        
        elif bool(intune_file_path) == True:
        # Intune date checker
            def dateIntune(intune_df):
                intune_df = pd.read_csv(intune_file_path, header=0)
                two_weeks_ago = datetime.now() - timedelta(days=14)
                datemask = pd.to_datetime(intune_df['Last check-in']) < two_weeks_ago
                dateIntune_df = intune_df.loc[datemask]
                return dateIntune_df
        elif bool(intune_file_path) == False:
            print('No Intune DF for Dates!')
            continue
        
        elif bool(webroot_file_path) == True:            
            # Webroot date checker
            def dateWebroot(webroot_df):
                webroot_df = pd.read_csv(webroot_file_path, header=0)
                two_weeks_ago = datetime.now() - timedelta(days=14)
                datemask = pd.to_datetime(webroot_df['Last Seen']) < two_weeks_ago
                dateWebroot_df = webroot_df.loc[datemask]  
                return dateWebroot_df
        elif bool(webroot_file_path) == False:
            print('No Webroot DF for Dates!')
            continue
                    
        diffAuto_df = diffAuto(audit_df, auto_df)
        diffIntune_df = diffIntune(audit_df, intune_df)
        diffWebroot_df = diffWebroot(audit_df, webroot_df)
        dateAuto_df = dateAuto(auto_df)
        dateIntune_df = dateIntune(intune_df)
        dateWebroot_df = dateWebroot(webroot_df)

        with pd.ExcelWriter(auto_audit_out, mode='w', engine='xlsxwriter') as writer:
            diffAuto_df.to_excel(writer, sheet_name='diffAuto', index=False)
            diffIntune_df.to_excel(writer, sheet_name='diffIntune', index=False)
            diffWebroot_df.to_excel(writer, sheet_name="diffWebroot", index=False)
            dateAuto_df.to_excel(writer, sheet_name='dateAuto', index=False)
            dateAuto_df.to_excel(writer, sheet_name='dateAuto', index=False)
            dateAuto_df.to_excel(writer, sheet_name='dateAuto', index=False)
            dateIntune_df.to_excel(writer, sheet_name='dateIntune', index=False)
            dateWebroot_df.to_excel(writer, sheet_name='dateWebroot', index=False)
    sg.popup('Your report was generated at ' + auto_audit_out)
window.close()