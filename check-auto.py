import pandas as pd
from  datetime import datetime, timedelta

# Define variables

autoIn = r'C:\scripting\Input\CLAutomate05-03-23.xlsx'
autoOut = r'C:\scripting\Output\date_auto_results.xlsx'

# Read data from input.xlsx
auto_df = pd.read_excel(autoIn, engine='openpyxl')

# Check if date is older than two weeks
two_weeks_ago = datetime.now() - timedelta(days=14)
datemask = auto_df['Last Contact'] < two_weeks_ago

# Filter dataframe
filtered_auto_df = auto_df.loc[datemask]

# Output filtered data to output.xlsx
filtered_auto_df.to_excel(autoOut, index=False)