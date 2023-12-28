import PySimpleGUI as sg
import pandas as pd

def file_to_df(lib):
    if lib['file_type'] == "xlsx":
        if lib['inv_system'] == "audit":
            return pd.read_excel(lib['file_path'], header=lib['header_row'], usecols=[lib['config_col'], lib['user_col']], engine='openpyxl')
        elif lib['inv_system'] == "auto":
            df = pd.read_excel(lib['file_path'], header=lib['header_row'], usecols=[lib['config_col'], lib['user_col'], lib['ctype_col'], lib['date_col']], engine='openpyxl')
            df.rename(columns={lib['config_col']: 'Configuration Name'}, inplace=True)
            return
        return 
    elif lib['file_type'] == "csv":
        if lib['inv_system'] == "manage":
            return pd.read_csv(lib['file_path'], header=lib['header_row'], usecols=[lib['config_col'], lib['user_col'], lib['ctype_col']])
        elif lib['inv_system'] == "intune":
            return pd.read_csv(lib['file_path'], header=lib['header_row'], usecols=[lib['config_col'], lib['user_col'], lib['date_col']])
        elif lib['inv_system'] == "webroot":
            return pd.read_csv(lib['file_path'], header=lib['header_row'], usecols=[lib['config_col'], lib['user_col'], lib['date_col'], lib['status_col']])
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
            "header_row": 0,
            "config_col": "Configuration Name",
            "user_col": "Contact",
            "ctype_col": "Configuration Type",
            "file_path": values['fpmanage'],
        }
            
        autoLib = {
            "inv_system": "auto",
            "file_type": autoExt,
            "header_row": 0,
            "config_col": autoConfName,
            "user_col": "Last Logged in User",
            "ctype_col": "Type",
            "date_col": "Last Contact",
            "file_path": values['fpauto'],
        }

        intuneLib = {
            "inv_system": "intune",
            "file_type": "csv",
            "header_row": 0,
            "config_col": "Device name",
            "user_col": "Primary user display name",
            "date_col": "Last check-in",
            "file_path": values['fpintune'],
        }

        webrootLib = {
            "inv_system": "webroot",
            "file_type": "csv",
            "header_row": 0,
            "config_col": "Name",
            "user_col": "Current User",
            "date_col": "Last Seen",
            "status_col": "Status",
            "file_path": values['fpwebroot'],
        }
        # Outpath defined in the UI
        auto_audit_out = values['outpath'] + '/audit-discrepancies.xlsx'
        
        # Supresses SettingWithCopyWarning log messages; pandas gets confused with the library implementation
        pd.options.mode.chained_assignment = None
        
        # Calls on the above functions and dictionaries to write a processed DataFrame into an .xlsx sheet; dependent on the file's presence.
        if event == "Submit":
            if auditLib['file_path']:
                audit_df = file_to_df(auditLib)
                print(audit_df)
            else:
                sg.popup('Audit file not found.')
                break
            if manageLib['file_path']:
                manage_df = file_to_df(manageLib)
                print(manage_df)
            if autoLib['file_path']:
                auto_df = file_to_df(autoLib)
                print(auto_df)
            if intuneLib['file_path']:
                intune_df = file_to_df(intuneLib)
                print(intune_df)
            if webrootLib['file_path']:
                webroot_df = file_to_df(webrootLib)
                print(webroot_df)
            
        sg.popup('Location: ' + auto_audit_out)
        # Change me to "break" if program should close after clicking "OK" on pop-up
        continue
    window.close()

if __name__ == "__main__":
    main()