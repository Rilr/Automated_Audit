import pandas as pd

#-----get the files-----#
auditPath = r'C:\scripting\Input\Centerlink Monthly Audit.xlsx'
autoPath = r'C:\scripting\Input\CLAutomate05-03-23.xlsx'
outPath = r'C:\scripting\Output\missing_audit-auto.xlsx'

#-----read the files-----#
audit_df = pd.read_excel(auditPath, sheet_name=0, header=2, usecols=[4], engine='openpyxl')
auto_df = pd.read_excel(autoPath, usecols=[0], engine="openpyxl")

#-----clean the data-----#
audit_df = audit_df.dropna(how='all')

#-----merge the files-----#
merged_df = pd.merge(audit_df, auto_df, left_on='Configuration Name', right_on='Name', how='outer', suffixes=['_audit', '_auto'], indicator=True)

#-----make column values readable-----#
merged_df = merged_df.rename(columns={'_merge': 'Found In'})
merged_df['Found In'] = merged_df['Found In'].replace({'left_only': 'Only in Audit', 'right_only': 'Only in Automate'})

#-----export the file-----#
merged_df.to_excel(outPath, index=False)