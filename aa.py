import PySimpleGUI as sg
import pandas as pd

def file_to_df(lib):
    if lib['file_type'] == "xlsx":
        return pd.read_excel(lib['file_path'], header=lib['header_row'], usecols=[lib['config_col'], lib['user_col']], engine='openpyxl')
    elif lib['file_type'] == "csv":
        return pd.read_csv(lib['file_path'], header=lib['header_row'], usecols=[lib['config_col'], lib['user_col']])
    return

def main():
    # Define UI parameters
    sg.theme('DarkAmber')
    layout = [[sg.Text('Audit XLSX')],
            [sg.InputText(key='fpaudit'),
            sg.FileBrowse(file_types=[("XLSX Files","*.xlsx")])],
            
            [sg.Text('Manage CSV')],
            [sg.InputText(key="fpmanage"),
            sg.FileBrowse(file_types=[("CSV Files","*.csv")])],
            
            [sg.Text('ITS247 XLSX')],
            [sg.InputText(key="fpits"),
            sg.FileBrowse(file_types=[("XLSX Files","*.xlsx")])],
            
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
        
        # Define dictionaries
        auditLib = {
            "inv_system": "audit",
            "file_type": "xlsx",
            "sheet_name": 0,
            "header_row": 2,
            "config_col": "Configuration Name",
            "user_col": "Current User",
            "file_path": values['fpaudit'],
        }
        
        manageLib = {
            "inv_system": "manage",
            "file_type": "csv",
            "sheet_name": 0,
            "header_row": 0,
            "config_col": "Configuration Name",
            "user_col": "Contact",
            "file_path": values['fpmanage'],
        }
        
        itsLib = {
            "inv_system": "its",
            "file_type": "xlsx",
            "sheet_name": 0,
            "header_row": 0,
            "config_col": "Name",
            "user_col": "Last User",
            "date_col": "Last Online",
            "file_path": values['fpits'],
        }
        
        # Outpath defined in the UI
        auto_audit_out = values['outpath'] + '/audit-discrepancies.xlsx'
        
        # Supresses SettingWithCopyWarning log messages; pandas gets confused with the library implementation
        pd.options.mode.chained_assignment = None
        
        # Calls on the above functions and dictionaries to write a processed DataFrame into an .xlsx sheet; dependent on the file's presence.
        if event == "Submit":
            #TODO grab all data and port it to it's own sheet
            devices_df = pd.DataFrame()
            if auditLib['file_path']:
                audit_df = file_to_df(auditLib)
                devices_df= audit_df[auditLib['config_col']].to_frame()
        
            if manageLib['file_path']:
                manage_df = file_to_df(manageLib)
                manage_df = manage_df[manageLib['config_col']].to_frame()
                devices_df = pd.merge(devices_df, manage_df, on=[auditLib['config_col']], how='outer')
            
            if itsLib['file_path']:
                its_df = file_to_df(itsLib)
                its_df = its_df[itsLib['config_col']].to_frame()
                its_df = its_df.rename(columns={"Name": "Configuration Name"})
                devices_df = pd.merge(devices_df, its_df, on=[auditLib['config_col']], how='outer')
                
            devices_df.to_excel(auto_audit_out, index=False)
                
        sg.popup('Location: ' + auto_audit_out)
        # Change me to "break" if program should close after clicking "OK" on pop-up
        continue
    window.close()

if __name__ == "__main__":
    main()