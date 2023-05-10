import pandas as pd

#-----get the files-----#
auditPath = r'C:\scripting\Input\Centerlink Monthly Audit.xlsx'
webrootPath = r'C:\scripting\Input\centerlinkwr0523.csv'
outPath = r'C:\scripting\Output\missing_audit-webroot.xlsx'

#-----read the files-----#
audit_df = pd.read_excel(auditPath, header=2, usecols=[4], engine='openpyxl')
webroot_df = pd.read_csv(webrootPath, header=0, usecols=[0])
audit_df = audit_df.dropna(how='all')

#-----merge the files-----#
merged_df = pd.merge(audit_df, webroot_df, left_on='Configuration Name', right_on='Hostname', how='outer', suffixes=['_audit', '_intune'], indicator=True)

#-----make column values readable-----#
merged_df = merged_df.rename(columns={'_merge': 'Found In'})
merged_df['Found In'] = merged_df['Found In'].replace({'left_only': 'Only in Audit', 'right_only': 'Only in Webroot'})

#-----export the file-----#
merged_df.to_excel(outPath, index=False)