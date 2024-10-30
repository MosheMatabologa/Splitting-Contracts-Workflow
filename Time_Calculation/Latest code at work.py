import pandas as pd

# File path
file1 = r'C:\Users\Q624157\Desktop\Time_Calculation\Time_Calculation_Web_App\1wtestrearranged_data.xlsx'
# Read data from Excel file
df = pd.read_excel(file1, sheet_name="Sheet1")

# Sort the DataFrame
df.sort_values(by=['Personnel number', 'Time Event Date', 'Time Event In/Out Time'], inplace=True)

# Convert time strings to timedelta with the correct format including seconds
df['Time Event In/Out Time'] = df['Time Event In/Out Time'].astype(str).str.replace('24:00', '00:00:00')
df['Time Event In/Out Time'] = pd.to_timedelta(df['Time Event In/Out Time'], errors='coerce')

# Determine Clock-in and Clock-out
df['Event Type'] = df.groupby(['Personnel number', 'Time Event Date'])['Time Event In/Out Time'].transform(lambda x: x.diff().fillna(pd.Timedelta(seconds=0)) >= pd.Timedelta(seconds=0)).map({True: 'IN', False: 'OUT'})

# Calculate Worked Hours
def calculate_worked_hours(group):
    hours_worked = pd.Timedelta(seconds=0)
    in_time = None
    for _, row in group.iterrows():
        if row['Event Type'] == 'IN':
            in_time = row['Time Event In/Out Time']
        elif row['Event Type'] == 'OUT' and in_time is not None:
            hours_worked += row['Time Event In/Out Time'] - in_time
            in_time = None
    return hours_worked

df['Worked Hours'] = df.groupby(['Personnel number', 'Time Event Date']).apply(calculate_worked_hours).reset_index(drop=True)

df.to_excel('TimeDataOutput.xlsx', index=False, engine='xlsxwriter')
# Calculate hours worked between 14:00 and 18:00