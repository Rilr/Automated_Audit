import PySimpleGUI as sg
import pandas as pd

# Supresses SettingWithCopyWarning log messages; pandas gets confused with the library implementation
pd.options.mode.chained_assignment = None

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
        
        [sg.Button("Submit"),sg.Cancel("Close")]],
window = sg.Window('Auto-Automate by JLS', layout)
event, values = window.read()
auto_audit_out = values['outpath'] + '/audit-discrepancies.xlsx'

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

def file_to_df(lib):
    if lib['file_type'] == 'xlsx':
        return pd.read_excel(lib['file_path'], header=lib['header_row'], usecols=[lib['config_col'], lib['user_col']], engine='openpyxl')
    elif lib['file_type'] == 'csv':
        return pd.read_csv(lib['file_path'], header=lib['header_row'], usecols=[lib['config_col'], lib['user_col']])
    return

def df_to_file(auditLib, manageLib, itsLib, devices_df):
    with pd.ExcelWriter(auto_audit_out, mode='w', engine='xlsxwriter') as writer:
        if auditLib['file_path']:
            audit_df = file_to_df(auditLib)
            audit_df.to_excel(writer, sheet_name='Audit', index=False)
        if manageLib['file_path']:
            manage_df = file_to_df(manageLib)
            manage_df.to_excel(writer, sheet_name='Manage', index=False)
        if itsLib['file_path']:
            its_df = file_to_df(itsLib)
            its_df.to_excel(writer, sheet_name='ITS', index=False)
        devices_df.to_excel(writer, sheet_name='Devices', index=False)

def agg_dfs(auditLib, manageLib, itsLib):
    if auditLib['file_path']:
        audit_df = file_to_df(auditLib)
        devices_df = audit_df[auditLib['config_col']].to_frame()
    if manageLib['file_path']:
        manage_df = file_to_df(manageLib)
        manage_df = manage_df[manageLib['config_col']].to_frame()
        devices_df = pd.merge(devices_df, manage_df, on=auditLib['config_col'], how='outer')
    if itsLib['file_path']:
        its_df = file_to_df(itsLib)
        its_df = its_df[itsLib['config_col']].to_frame()
        its_df = its_df.rename(columns={"Name": "Configuration Name"})
        devices_df = pd.merge(devices_df, its_df, on=auditLib['config_col'], how='outer')
    return devices_df

def main():
    while True:
        # Graceful program shutdown
        if event in (sg.WIN_CLOSED,'Cancel'):
            break
        # Calls on the above functions and dictionaries to write a processed DataFrame into an .xlsx sheet; dependent on the file's presence.
        if event == "Submit":
            agg_dfs(auditLib, manageLib, itsLib) 
            df_to_file(auditLib, manageLib, itsLib, devices_df=agg_dfs(auditLib, manageLib, itsLib))
            sg.popup("Audit Discrepancies have been written to " + auto_audit_out)
            break
        # Change me to "break" if program should close after clicking "OK" on pop-up
        continue
    window.close()

if __name__ == "__main__":
    main()