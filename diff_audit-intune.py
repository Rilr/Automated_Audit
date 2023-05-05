import pandas as pd

#-----get the files-----#
auditPath = r'C:\scripting\Input\The Childrens Trust Monthly Audit.xlsx'
intunePath = r'C:\scripting\Input\tctintune0523.csv'
outPath = r'C:\scripting\Output\missing_audit-intune.xlsx'

#-----read the files-----#
audit_df = pd.read_excel(auditPath, sheet_name=0, header=2, usecols=[4], engine='openpyxl')
intune_df = pd.read_csv(intunePath, header=0, usecols=[1])

#-----merge the files-----#
merged_df = pd.merge(audit_df, intune_df, left_on='Configuration Name', right_on='Device name', how='outer', suffixes=['_audit', '_intune'], indicator=True)

#-----make column values readable-----#
merged_df = merged_df.rename(columns={'_merge': 'Found In'})
merged_df['Found In'] = merged_df['Found In'].replace({'left_only': 'Only in Audit', 'right_only': 'Only in Intune'})

#-----export the file-----#
merged_df.to_excel(outPath, index=False)