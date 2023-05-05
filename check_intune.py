import pandas as pd
import datetime as dt

# Define variables

intuneIn = r'C:\scripting\Input\DevicesWithInventory_7fdc4978-2d90-45d5-8a76-cbd53daad3f7.csv'
intuneOut = r'C:\scripting\Output\intune_date_results.xlsx'

# Read data from input.xlsx
intune_df = pd.read_csv(intuneIn)

# Check if date is older than two weeks
two_weeks_ago = dt.datetime.now() - dt.timedelta(days=14)
datemask = pd.to_datetime(intune_df['Last check-in']) < two_weeks_ago

# Filter dataframe
filtered_intune_df = intune_df.loc[datemask]

# Output filtered data to output.xlsx
filtered_intune_df.to_excel(intuneOut, index=False)