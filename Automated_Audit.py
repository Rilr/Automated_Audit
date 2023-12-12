import PySimpleGUI as sg
import pandas as pd
from datetime import datetime, timedelta

#TODO Look into class implementation to handle the data more cleanly. 

def read_file_to_df(lib):
    if lib['file_type'] == "xlsx":
        return pd.read_excel(lib['file_path'], header=lib['header'], usecols=[lib['column']], engine='openpyxl')
    elif lib['file_type'] == "csv":
        return pd.read_csv(lib['file_path'], header=lib['header'], usecols=[lib['column']])

def execSummary():
    '''Aggregates all submitted data into one sheet for simple reading'''
    # Combine all data
    # Index off of the device name
    # Iterate through each name and get data from each dataframe
    # Return one large dataframe to be exported to its own excel sheet
    pass

def probSummary():
    '''Highlights problem workstations for review'''
    # Get data from diffChecker() for each library (exclude devices labelled "both")
    # Get data from dateChecker() for each relevant library
    # Concat along device name and include columns for each issue. (inner merge maybe?)
    pass

def diffChecker(source_lib, input_lib):
    '''Cleans and compares data from two frames
    
    Args:
        audit_df: Pandas data frame of audit file to compare
        input_lib: {
            "inv_system": str,
            "config_name": str,
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
    source_df = read_file_to_df(source_lib)
    input_df = read_file_to_df(input_lib)

    # Clean the DataFrames by removing extraneous data and uppercase conversion
    source_df = source_df.dropna(how='all')
    source_df[source_lib['config_name']] = source_df[source_lib['config_name']].str.upper()
    source_df[source_lib['config_name']] = source_df[source_lib['config_name']].replace(r"^ +| +$", r"", regex=True)
    input_df[input_lib['config_name']] = input_df[input_lib['config_name']].str.upper()
    input_df[input_lib['config_name']] = input_df[input_lib['config_name']].replace(r"^ +| +$", r"", regex=True)
    
    # Merge, compare and label the data
    
    diffinput_df = pd.merge(source_df, input_df, left_on=source_lib['config_name'], right_on=input_lib['config_name'], how='outer', suffixes=[f'_{source_lib["inv_system"]}', f'_{input_lib["inv_system"]}'], indicator=True)
    diffinput_df = diffinput_df.rename(columns={'_merge': 'Found In'})
    diffinput_df['Found In'] = diffinput_df['Found In'].replace({'left_only': f'Only in {source_lib["inv_system"].upper()}', 'right_only': f'Only in {input_lib["inv_system"].upper()}'})
    
    return diffinput_df

def dateChecker(input_lib):
    '''Loads pandas dataframe and checks if date is older than two_weeks_ago
    
    Args:
        input_lib: {
            "inv_system": str,
            "config_name": str,
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
    if input_lib['file_type'] == "xlsx":
        diff_df = pd.read_excel(input_lib['file_path'], header=input_lib['header'], engine='openpyxl')
        datemask = diff_df[input_lib['check_in']] < two_weeks_ago

    elif input_lib['file_type'] == "csv":
        diff_df = pd.read_csv(input_lib['file_path'], header=input_lib['header'])
        datemask = pd.to_datetime(diff_df[input_lib['check_in']]) < two_weeks_ago

    # Apply datemask to DataFrame
    dateDiff_df = diff_df.loc[datemask]
    
    return dateDiff_df

def main():
    
    # Define UI parameters
    sg.theme('DarkAmber')
    layout = [[sg.Text('Audit XLSX')],
            [sg.InputText(key='fpaudit'),
            sg.FileBrowse(file_types=[("XLSX Files","*.xlsx")])],
            
            [sg.Text('Manage CSV')],
            [sg.InputText(key="fpmanage"),
            sg.FileBrowse(file_types=[("CSV Files","*.csv")])],
            
            [sg.Text('Automate XLSX or CSV')],
            [sg.InputText(key="fpauto"),
            sg.FileBrowse(file_types=[("XLSX Files", "*.xlsx"), ("CSV Files", "*.csv")])],
            
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
        
        autoExt = values["fpauto"].split('.')[-1]
        if autoExt == "xlsx":
            autoConfName = "Name"
        elif autoExt == "csv":
            autoConfName = "Computer Name"

        # Define dictionaries
        auditLib = {
            "inv_system": "audit",
            "config_name": "Configuration Name",
            "sheet_name": 0,
            "file_type": "xlsx",
            "header": 2,
            "column": "Configuration Name",
            "file_path": values['fpaudit'],
        }

        manageLib = {
            "inv_system": "manage",
            "config_name": "Configuration Name",
            "file_type": "csv",
            "header": 0,
            "column": "Configuration Name",
            "file_path": values['fpmanage'],
        }

        autoLib = {
            "inv_system": "auto",
            "config_name": autoConfName,
            "file_type": autoExt,
            "header": 0,
            "column": autoConfName,
            "file_path": values['fpauto'],
            "check_in": "Last Contact"
        }

        intuneLib = {
            "inv_system": "intune",
            "config_name": "Device name",
            "file_type": "csv",
            "header": 0,
            "column": 1,
            "file_path": values['fpintune'],
            "check_in": "Last check-in"
        }
        
        webrootLib = {
            "inv_system": "webroot",
            "config_name": "Name",
            "file_type": "csv",
            "header": 0,
            "column": "Name",
            "file_path": values['fpwebroot'],
            "check_in": "Last Seen"
        }
        
        # Outpath defined in the UI
        auto_audit_out = values['outpath'] + '/audit-discrepancies.xlsx'
        
        # Supresses SettingWithCopyWarning log messages; pandas gets confused with the library implementation
        pd.options.mode.chained_assignment = None
        
        # Calls on the above functions and dictionaries to write a processed DataFrame into an .xlsx sheet; dependent on the file's presence.
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
                    dateChecker(autoLib).to_excel(writer, sheet_name='Old Auto', index=False)
                    
                if intuneLib['file_path']:
                    dateChecker(intuneLib).to_excel(writer, sheet_name='Old Intune', index=False)
                    
                if webrootLib['file_path']:
                    dateChecker(webrootLib).to_excel(writer, sheet_name='Old Webroot', index=False)
        # TODO Add button to open export location via values['outpath'], also button alignment is weird here
        sg.popup('Location: ' + auto_audit_out)
        
        # Change me to "break" if program should close after clicking "OK" on pop-up
        continue
    window.close()

if __name__ == "__main__":
    main()
# TODO Grab file type from path
# TODO Hook data from APIs