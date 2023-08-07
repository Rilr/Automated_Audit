import PySimpleGUI as sg
import pandas as pd
from datetime import datetime, timedelta

def diffChecker(source_lib, inventory_lib):
    '''Cleans and compares data from two frames
    
    Args:
        audit_df: Pandas data frame of audit file to compare
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

    Returns:
        New data frame with calculated comparision
    '''
    # Port source_lib['file_path'] to a DataFrame; based on file_type
    if source_lib['file_type'] == "excel":
        source_df = pd.read_excel(source_lib['file_path'], sheet_name=source_lib['sheet_name'], header=source_lib['header'], usecols=[source_lib['config_name']], engine='openpyxl')
    elif source_lib['file_type'] == "csv":
        source_df = pd.read_csv(source_lib['file_path'], header=source_lib['header'], usecols=[source_lib['config_name']])
    # Port inventory_lib['file_path'] to a DataFrame; based on file_type
    if inventory_lib['file_type'] == "excel":
        input_df = pd.read_excel(inventory_lib['file_path'], header=inventory_lib['header'], usecols=[inventory_lib['column']], engine='openpyxl')
    elif inventory_lib['file_type'] == "csv":
        input_df = pd.read_csv(inventory_lib['file_path'], header=inventory_lib['header'], usecols=[inventory_lib['column']])
    
    # Clean the DataFrames by removing extraneous data and uppercase conversion
    source_df = source_df.dropna(how='all')
    source_df[inventory_lib['config_name']] = source_df[inventory_lib['config_name']].str.upper()
    source_df[inventory_lib['config_name']] = source_df[inventory_lib['config_name']].replace(r"^ +| +$", r"", regex=True)
    input_df[inventory_lib['device_name']] = input_df[inventory_lib['device_name']].str.upper()
    input_df[inventory_lib['device_name']] = input_df[inventory_lib['device_name']].replace(r"^ +| +$", r"", regex=True)
    
    # Merge, compare and label the data
    diffinput_df = pd.merge(source_df, input_df, left_on=source_lib['config_name'], right_on=inventory_lib['device_name'], how='outer', suffixes=[f'_{source_lib["inv_system"]}', f'_{inventory_lib["inv_system"]}'], indicator=True)
    diffinput_df = diffinput_df.rename(columns={'_merge': 'Found In'})
    diffinput_df['Found In'] = diffinput_df['Found In'].replace({'left_only': f'Only in {source_lib["inv_system"].upper()}', 'right_only': f'Only in {inventory_lib["inv_system"].upper()}'})
    
    return diffinput_df

def dateChecker(inventory_lib):
    '''Loads pandas dataframe and checks if date is older than two_weeks_ago
    
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

# Supresses SettingWithCopyWarning log messages; pandas gets confused with the library implementation
pd.options.mode.chained_assignment = None

def main():
    
    # Define UI parameters
    sg.theme('DarkAmber')
    layout = [[sg.Text('Audit XLSX')],
            [sg.InputText(key='fpaudit'),
            sg.FileBrowse(file_types=[("xlsx Files","*.xlsx")])],
            
            [sg.Text('Manage CSV')],
            [sg.InputText(key="fpmanage"),
            sg.FileBrowse(file_types=[("CSV Files","*.csv")])],
            
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
        #audit_file_path = values['fpaudit']
        auto_audit_out = values['outpath'] + '/audit-discrepancies.xlsx'

        # Define libraries
        auditLib = {
            "inv_system": "audit",
            "config_name": "Configuration Name",
            "sheet_name": 0,
            "file_type": "excel",
            "header": 2,
            "column": "Configuration Name",
            "file_path": values['fpaudit'],
            "check_in": "Last Contact"
        }

        manageLib = {
            "inv_system": "manage",
            "config_name": "Configuration Name",
            "device_name": "Device name",
            "file_type": "csv",
            "header": 0,
            "column": 1,
            "file_path": values['fpmanage'],
            "check_in": "Last check-in"
        }

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

        # Calls on the above functions and libraries to write a processed DataFrame into an .xlsx sheet; dependant on the file's presence
        if event == "Submit":
            with pd.ExcelWriter(auto_audit_out, mode='w', engine='xlsxwriter') as writer:
                if auditLib['file_path'] and autoLib['file_path']:
                    diffChecker(auditLib, autoLib,).to_excel(writer, sheet_name='Audit v Auto', index=False)
                if auditLib['file_path'] and intuneLib['file_path']:
                    diffChecker(auditLib, intuneLib).to_excel(writer, sheet_name='Audit v Intune', index=False)
                if auditLib['file_path'] and webrootLib['file_path']:
                    diffChecker(auditLib, webrootLib).to_excel(writer, sheet_name='Audit v Webroot', index=False)
                if manageLib['file_path'] and autoLib['file_path']:
                    diffChecker(manageLib, autoLib).to_excel(writer, sheet_name='Manage v Auto', index=False)
                if autoLib['file_path']:
                    dateChecker(autoLib).to_excel(writer, sheet_name='dateAuto', index=False)
                if intuneLib['file_path']:
                    dateChecker(intuneLib).to_excel(writer, sheet_name='dateIntune', index=False)
                if webrootLib['file_path']:
                    dateChecker(webrootLib).to_excel(writer, sheet_name='dateWebroot', index=False)
        # TODO Add button to open export location via values['outpath'], also button alignment is weird here
        sg.popup('Your report was generated at ' + auto_audit_out)
        
        # Change me to "break" if program should close after clicking "OK" on pop-up
        continue
    window.close()

if __name__ == "__main__":
    main()
# TODO Grab file type from path
# TODO Hook data from APIs