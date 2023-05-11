import PySimpleGUI as sg
import pandas as pd
from datetime import datetime, timedelta

audit_file_path = r'C\scripting\Input\audit.xlsx'
auto_file_path = r'C:\scripting\Input\auto.xlsx'
intune_file_path = r'C:\scripting\Input\intune.csv'
webroot_file_path = r'C:\scripting\Input\webroot.csv'

auto_diff_out = r'C:\scripting\Output\auditdiff.xlsx'
auto_date_out = r'C:\scripting\Output\auditdate.xlsx'
intune_diff_out = r'C:\scripting\Output\intunediff.xlsx'
intune_date_out = r'C:\scripting\Output\intunedate'
webroot_diff_out = r'C:\scripting\Output\webrootdiff.xlsx'
webroot_date_out = r'C:\scripting\Output\webrootdate.xlsx'

audit_df = audit_df = pd.read_excel(audit_file_path, sheet_name=0, header=2, usecols=[4], engine='openpyxl')
auto_df = pd.read_excel(audit_file_path, engine='openpyxl')
intune_df = pd.read_csv(intune_file_path, header=0, usecols=[1])
webroot_df = pd.read_csv(webroot_file_path, header=0, usecols=[0])

# Automate difference checker
def diffAuto(audit_df, auto_df):
    audit_df = audit_df.dropna(how='all')
    diffAuto_df = pd.merge(audit_df, auto_df, left_on='Configuration Name', right_on='Name', how='outer', suffixes=['_audit', '_auto'], indicator=True)
    diffAuto_df = diffAuto_df.rename(columns={'_merge': 'Found In'})
    diffAuto_df['Found In'] = diffAuto_df['Found In'].replace({'left_only': 'Only in Audit', 'right_only': 'Only in Automate'})
    return(diffAuto_df)

# Intune difference checker
def diffIntune(audit_df, intune_df):
    audit_df = audit_df.dropna(how='all')
    diffIntune_df = pd.merge(audit_df, intune_df, left_on='Configuration Name', right_on='Device name', how='outer', suffixes=['_audit', '_intune'], indicator=True)
    diffIntune_df = diffIntune_df.rename(columns={'_merge': 'Found In'})
    diffIntune_df['Found In'] = diffIntune_df['Found In'].replace({'left_only': 'Only in Audit', 'right_only': 'Only in Intune'})
    return(diffIntune_df)

# Webroot difference checker
def diffWebroot(audit_df, webroot_df):
    audit_df = audit_df.dropna(how='all')
    diffWebroot_df = pd.merge(audit_df, webroot_df, left_on='Configuration Name', right_on='Hostname', how='outer', suffixes=['_audit', '_intune'], indicator=True)
    diffWebroot_df = diffWebroot_df.rename(columns={'_merge': 'Found In'})
    diffWebroot_df['Found In'] = diffWebroot_df['Found In'].replace({'left_only': 'Only in Audit', 'right_only': 'Only in Webroot'})
    return(diffWebroot_df)

# Automate date checker
def dateAuto(auto_df):
    two_weeks_ago = datetime.now() - timedelta(days=14)
    datemask = auto_df['Last Contact'] < two_weeks_ago
    dateAuto_df = auto_df.loc[datemask]
    return(dateAuto_df)

# Intune date checker
def dateIntune(intune_df):
    two_weeks_ago = datetime.now() - timedelta(days=14)
    datemask = pd.to_datetime(intune_df['Last check-in']) < two_weeks_ago
    dateIntune_df = intune_df.loc[datemask]
    return(dateIntune_df)

# Webroot date checker
def dateWebroot(webroot_df):
    two_weeks_ago = datetime.now() - timedelta(days=14)
    datemask = pd.to_datetime(webroot_df['Last Seen']) < two_weeks_ago
    dateWebroot_df = webroot_df.loc[datemask]  
    return(dateWebroot_df)