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
          [sg.Text('File Output Location')],
          [sg.InputText(key="outpath"),
           sg.FolderBrowse()],
          [sg.Button("Submit"),sg.Cancel()]],
window = sg.Window('AutoAutomate', layout)

<<<<<<< Updated upstream
while True:
    event, values = window.read()
    audit_file_path = values['fpaudit']
    auto_file_path = values['fpauto']
    intune_file_path = values['fpintune']
    webroot_file_path = values['fpwebroot']
    audit_df = pd.read_excel(audit_file_path, sheet_name=0, header=2, usecols=[4], engine='openpyxl')
    auto_df, intune_df, webroot_df = [], [], []
    auto_audit_out = values["outpath"] + '/audit-discrepancies.xlsx'

    if event in (sg.WIN_CLOSED,'Cancel'):
        break
    
    elif event == "Submit":        
        if not audit_file_path:
            sg.popup_error("No Audit XLSX File Uploaded!")
            continue
        # Automate difference checker
        if audit_file_path:
            def diffAuto(audit_df, auto_df):
                audit_df = audit_df.dropna(how='all')
                auto_df = pd.read_excel(auto_file_path, header=0, usecols=[0], engine='openpyxl')
                audit_df['Configuration Name'] = audit_df['Configuration Name'].str.upper()
                audit_df['Configuration Name'] = audit_df['Configuration Name'].replace(r"^ +| +$", r"", regex=True)
                auto_df['Name'] = auto_df['Name'].str.upper()
                auto_df['Name'] = auto_df['Name'].replace(r"^ +| +$", r"", regex=True)
                diffAuto_df = pd.merge(audit_df, auto_df, left_on='Configuration Name', right_on='Name', how='outer', suffixes=['_audit', '_auto'], indicator=True)
                diffAuto_df = diffAuto_df.rename(columns={'_merge': 'Found In'})
                diffAuto_df['Found In'] = diffAuto_df['Found In'].replace({'left_only': 'Only in Audit', 'right_only': 'Only in Automate'})
                return diffAuto_df
            diffAuto_df = diffAuto(audit_df, auto_df)
        else:
            print('No Audit File Uploaded!')

        if intune_file_path:
            # Intune difference checker
            def diffIntune(audit_df, intune_df):
                audit_df = audit_df.dropna(how='all')
                intune_df = pd.read_csv(intune_file_path, header=0, usecols=[1])
                audit_df['Configuration Name'] = audit_df['Configuration Name'].str.upper()
                audit_df['Configuration Name'] = audit_df['Configuration Name'].replace(r"^ +| +$", r"", regex=True)
                intune_df['Device name'] = intune_df['Device name'].str.upper()
                intune_df['Device name'] = intune_df['Device name'].replace(r"^ +| +$", r"", regex=True)
                diffIntune_df = pd.merge(audit_df, intune_df, left_on='Configuration Name', right_on='Device name', how='outer', suffixes=['_audit', '_intune'], indicator=True)
                diffIntune_df = diffIntune_df.rename(columns={'_merge': 'Found In'})
                diffIntune_df['Found In'] = diffIntune_df['Found In'].replace({'left_only': 'Only in Audit', 'right_only': 'Only in Intune'})
                return diffIntune_df
            diffIntune_df = diffIntune(audit_df, intune_df)
        else:
            print('No Intune File Uploaded!')
            continue

        if webroot_file_path:
        # Webroot difference checker
            def diffWebroot(audit_df, webroot_df):
                audit_df = audit_df.dropna(how='all')
                webroot_df = pd.read_csv(webroot_file_path, header=0, usecols=[0])
                audit_df['Configuration Name'] = audit_df['Configuration Name'].str.upper()
                audit_df['Configuration Name'] = audit_df['Configuration Name'].replace(r"^ +| +$", r"", regex=True)
                webroot_df['Hostname'] = webroot_df['Hostname'].str.upper()
                webroot_df['Hostname'] = webroot_df['Hostname'].replace(r"^ +| +$", r"", regex=True)
                diffWebroot_df = pd.merge(audit_df, webroot_df, left_on='Configuration Name', right_on='Hostname', how='outer', suffixes=['_audit', '_intune'], indicator=True)
                diffWebroot_df = diffWebroot_df.rename(columns={'_merge': 'Found In'})
                diffWebroot_df['Found In'] = diffWebroot_df['Found In'].replace({'left_only': 'Only in Audit', 'right_only': 'Only in Webroot'})
                return diffWebroot_df
            diffWebroot_df = diffWebroot(audit_df, webroot_df)
        else:
            print('No Webroot File Uploaded!')
            continue

        if auto_file_path:
        # Automate date checker
            def dateAuto(auto_df):
                auto_df = pd.read_excel(auto_file_path, header=0, engine='openpyxl')
                two_weeks_ago = datetime.now() - timedelta(days=14)
                datemask = auto_df['Last Contact'] < two_weeks_ago
                dateAuto_df = auto_df.loc[datemask]
                return dateAuto_df
            dateAuto_df = dateAuto(auto_df)
        else:
            print('No Automate file for Dates!')
            continue
        
        if intune_file_path:
        # Intune date checker
            def dateIntune(intune_df):
                intune_df = pd.read_csv(intune_file_path, header=0)
                two_weeks_ago = datetime.now() - timedelta(days=14)
                datemask = pd.to_datetime(intune_df['Last check-in']) < two_weeks_ago
                dateIntune_df = intune_df.loc[datemask]
                return dateIntune_df
            dateIntune_df = dateIntune(intune_df)
        else:
            print('No Intune DF for Dates!')
            continue
        
        if webroot_file_path:            
            # Webroot date checker
            def dateWebroot(webroot_df):
                webroot_df = pd.read_csv(webroot_file_path, header=0)
                two_weeks_ago = datetime.now() - timedelta(days=14)
                datemask = pd.to_datetime(webroot_df['Last Seen']) < two_weeks_ago
                dateWebroot_df = webroot_df.loc[datemask]  
                return dateWebroot_df
            dateWebroot_df = dateWebroot(webroot_df)
        else:
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
    break
window.close()
=======
    Returns:
        New DataFrame with calculated comparision
    '''

    # Port file_path to a DataFrame; based on file_type
    if inventory_lib['file_type'] == "excel":
        input_df = pd.read_excel(inventory_lib['file_path'], header=inventory_lib['header'], usecols=[inventory_lib['column']], engine='openpyxl')
    elif inventory_lib['file_type'] == "csv":
        input_df = pd.read_csv(inventory_lib['file_path'], header=inventory_lib['header'], usecols=[inventory_lib['column']])
    
    # Clean the DataFrames with uppercase conversion and removing outter spaces
    audit_df = audit_df.dropna(how='all')
    audit_df[inventory_lib['config_name']] = audit_df[inventory_lib['config_name']].str.upper()
    audit_df[inventory_lib['config_name']] = audit_df[inventory_lib['config_name']].replace(r"^ +| +$", r"", regex=True)
    input_df[inventory_lib['device_name']] = input_df[inventory_lib['device_name']].str.upper()
    input_df[inventory_lib['device_name']] = input_df[inventory_lib['device_name']].replace(r"^ +| +$", r"", regex=True)
    # Merge, compare and label the data
    diffinput_df = pd.merge(audit_df, input_df, left_on=inventory_lib['config_name'], right_on=inventory_lib['device_name'], how='outer', suffixes=['_audit', f'_{inventory_lib["inv_system"]}'], indicator=True)
    diffinput_df = diffinput_df.rename(columns={'_merge': 'Found In'})
    diffinput_df['Found In'] = diffinput_df['Found In'].replace({'left_only': 'Only in AUDIT', 'right_only': f'Only in {inventory_lib["inv_system"].upper()}'})
    
    return diffinput_df

def dateChecker(inventory_lib):
    '''Loads pandas DataFrame and checks if date is older than 2 weeks
    
    Args:
        inventory_lib: {
            "inv_system": str,
            "config_name": str,
            "device_name": str,
            "file_type": str,
            "header": int,
            "column": int,
            "file_path": str,
            "check_in": str
        }

    Return:
        New DataFrame with only items older than two_weeks_ago
    '''
    
    two_weeks_ago = datetime.now() - timedelta(days=14)

    # Port the file_path to a DataFrame and define datemask handling based on file_type
    if inventory_lib['file_type'] == "excel":
        diff_df = pd.read_excel(inventory_lib['file_path'], header=inventory_lib['header'], engine='openpyxl')
        datemask = diff_df[inventory_lib['check_in']] < two_weeks_ago

    elif inventory_lib['file_type'] == "csv":
        diff_df = pd.read_csv(inventory_lib['file_path'], header=inventory_lib['header'])
        datemask = pd.to_datetime(diff_df[inventory_lib['check_in']]) < two_weeks_ago

    # Apply datemask to DataFrame
    dateDiff_df = diff_df.loc[datemask]
    
    return dateDiff_df
  
def main():

    # Define UI parameters
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
            [sg.Text('File Output Location')],
            [sg.InputText(key="outpath"),
            sg.FolderBrowse()],
            [sg.Button("Submit"),sg.Cancel()]],
    window = sg.Window('Auto-Automate by JLS', layout)

    while True:
        event, values = window.read()
        
        # Graceful program shutdown
        if event in (sg.WIN_CLOSED,'Cancel'):
            break  
        
        # Declare static variables used in functions
        audit_file_path = values['fpaudit']
        audit_df = pd.read_excel(audit_file_path, sheet_name=0, header=2, usecols=[4], engine='openpyxl')
        auto_audit_out = values['outpath'] + '/audit-discrepancies.xlsx'

        # Define libraries
        autoLib = {
            "inv_system": "auto",
            "config_name": "Configuration Name",
            "device_name": "Name",
            "file_type": "excel",
            "header": 0,
            "column": 0,
            "file_path": values['fpauto'],
            "check_in": "Last Contact"
        }

        intuneLib = {
            "inv_system": "intune",
            "config_name": "Configuration Name",
            "device_name": "Device name",
            "file_type": "csv",
            "header": 0,
            "column": 1,
            "file_path": values['fpintune'],
            "check_in": "Last check-in"
        }
        
        webrootLib = {
            "inv_system": "webroot",
            "config_name": "Configuration Name",
            "device_name": "Hostname",
            "file_type": "csv",
            "header": 0,
            "column": 0,
            "file_path": values['fpwebroot'],
            "check_in": "Last Seen"
        }

        if event == "Submit":
            # Calls on the above functions and libraries to write a processed DataFrame into an .xlsx sheet; dependant on the file's presence
            with pd.ExcelWriter(auto_audit_out, mode='w', engine='xlsxwriter') as writer:
                if audit_file_path and autoLib['file_path']:
                    diffChecker(audit_df, autoLib,).to_excel(writer, sheet_name='diffAuto', index=False)
                if audit_file_path and intuneLib['file_path']:
                    diffChecker(audit_df, intuneLib).to_excel(writer, sheet_name='diffIntune', index=False)
                if audit_file_path and webrootLib['file_path']:
                    diffChecker(audit_df, webrootLib).to_excel(writer, sheet_name='diffWebroot', index=False)
                if autoLib['file_path']:
                    dateChecker(autoLib).to_excel(writer, sheet_name='dateAuto', index=False)
                if intuneLib['file_path']:
                    dateChecker(intuneLib).to_excel(writer, sheet_name='dateIntune', index=False)
                if webrootLib['file_path']:
                    dateChecker(webrootLib).to_excel(writer, sheet_name='dateWebroot', index=False)
            sg.popup('Your report was generated at ' + auto_audit_out)
            break
    window.close()

if __name__ == "__main__":
    main()
# TODO Grab file type from path
# TODO Hook data from APIs
>>>>>>> Stashed changes
