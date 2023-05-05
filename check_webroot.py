import pandas as pd
import datetime as dt

# Define variables

webrootIn = r'C:\scripting\Input\centerlinkwr0523.csv'
webrootOut = r'C:\scripting\Output\date_webroot_results.xlsx'

# Read data from input.xlsx
webroot_df = pd.read_csv(webrootIn)

# Check if date is older than two weeks
two_weeks_ago = dt.datetime.now() - dt.timedelta(days=14)
datemask = pd.to_datetime(webroot_df['Last Seen']) < two_weeks_ago

# Filter dataframe
filtered_webroot_df = webroot_df.loc[datemask]

# Output filtered data to output.xlsx
filtered_webroot_df.to_excel(webrootOut, index=False)