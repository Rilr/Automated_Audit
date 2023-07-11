import PySimpleGUI as sg
import pandas as pd
from datetime import datetime, timedelta

def diffChecker(audit_df, inventory_lib):
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

    if inventory_lib['file_type'] == "excel":
        input_df = pd.read_excel(inventory_lib['file_path'], header=inventory_lib['header'], usecols=[inventory_lib['column']], engine='openpyxl')
    elif inventory_lib['file_type'] == "csv":
        input_df = pd.read_csv(inventory_lib['file_path'], header=inventory_lib['header'], usecols=[inventory_lib['column']])
    
    audit_df = audit_df.dropna(how='all')
    audit_df[inventory_lib['config_name']] = audit_df[inventory_lib['config_name']].str.upper()
    audit_df[inventory_lib['config_name']] = audit_df[inventory_lib['config_name']].replace(r"^ +| +$", r"", regex=True)
    input_df[inventory_lib['device_name']] = input_df[inventory_lib['device_name']].str.upper()
    input_df[inventory_lib['device_name']] = input_df[inventory_lib['device_name']].replace(r"^ +| +$", r"", regex=True)
    diffinput_df = pd.merge(audit_df, input_df, left_on=inventory_lib['config_name'], right_on=inventory_lib['device_name'], how='outer', suffixes=['_audit', f'_{inventory_lib["inv_system"]}'], indicator=True)
    diffinput_df = diffinput_df.rename(columns={'_merge': 'Found In'})
    diffinput_df['Found In'] = diffinput_df['Found In'].replace({'left_only': 'Only in AUDIT', 'right_only': f'Only in {inventory_lib["inv_system"].upper()}'})
    
    return diffinput_df

def dateChecker(inventory_lib):
    '''Loads pandas dataframe and checks if date is older than 2 weeks
    
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
        New data frame with calculated comparision
    '''
    two_weeks_ago = datetime.now() - timedelta(days=14)

    # TODO Grab file type from path
    if inventory_lib['file_type'] == "excel":
        diff_df = pd.read_excel(inventory_lib['file_path'], header=inventory_lib['header'], engine='openpyxl')
        datemask = diff_df[inventory_lib['check_in']] < two_weeks_ago

    elif inventory_lib['file_type'] == "csv":
        diff_df = pd.read_csv(inventory_lib['file_path'], header=inventory_lib['header'])
        datemask = pd.to_datetime(diff_df[inventory_lib['check_in']]) < two_weeks_ago

    dateDiff_df = diff_df.loc[datemask]
    
    return dateDiff_df
  
def main():
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
            [sg.Text('File Output Location')],
            [sg.InputText(key="outpath"),
            sg.FolderBrowse()],
            [sg.Button("Submit"),sg.Cancel()]],
    window = sg.Window('Auto-Automate by JLS', layout)

    while True:
        event, values = window.read()

        # Declare Variables for the DataFrames
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

        if event in (sg.WIN_CLOSED,'Cancel'):
            break

        if event == "Submit":
            with pd.ExcelWriter(auto_audit_out, mode='w', engine='xlsxwriter') as writer:
                if audit_file_path and autoLib['file_path']:
                    diffChecker(audit_df, autoLib,).to_excel(writer, sheet_name='diffAuto', index=False)
                if audit_file_path and intuneLib['file_path']:
                    diffChecker(audit_df, intuneLib).to_excel(writer, sheet_name='diffIntune', index=False)
                if autoLib['file_path']:
                    dateChecker(autoLib).to_excel(writer, sheet_name='dateAuto', index=False)
                if intuneLib['file_path']:
                    dateChecker(intuneLib).to_excel(writer, sheet_name='dateIntune', index=False)
        sg.popup('Your report was generated at ' + auto_audit_out)
        break
    window.close()

if __name__ == "__main__":
    main()

