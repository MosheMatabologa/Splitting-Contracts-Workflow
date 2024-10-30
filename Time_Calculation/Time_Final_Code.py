import pandas as pd
from datetime import timedelta

file1 = r'C:\Users\Q624157\Desktop\Time_Calculation\Shift Allowance Calculations.xlsx'

# Read data from Excel file
df = pd.read_excel(file1, sheet_name="Sheet1")

hours_in_day = 48

# Convert time strings to timedelta with the correct format including seconds
df['Clock Out Time'] = pd.to_timedelta(df['Clock Out Time'].str.replace('24:00', '00:00:00'), errors='coerce')
df['Clock In Time'] = pd.to_timedelta(df['Clock In Time'].str.replace('24:00', '00:00:00'), errors='coerce')

# Calculate total time clocked for the day
df['Clocking Totals (Day)'] = df['Clock Out Time'] - df['Clock In Time']

# Calculate hours worked between 14:00 and 18:00
def calculate_hours_worked_14_to_18(row):
    start_time = row['Clock In Time']
    end_time = row['Clock Out Time']
    
    # Define the time boundaries for a 48-hour day
    start_boundary = pd.Timedelta(hours=14)  # 14:00
    end_boundary = pd.Timedelta(hours=18)  # 18:00
    
    # Ensure start and end times are within the boundaries
    start_time_within_range = max(start_time, start_boundary)
    end_time_within_range = min(end_time, end_boundary)
    
    # Calculate hours worked within the boundaries, ensuring non-negative
    hours_worked = max(0, (end_time_within_range - start_time_within_range).total_seconds() / 3600)
    
    return pd.Timedelta(hours=hours_worked)

# Calculate hours worked between 18:00 and 06:00

def calculate_hours_worked_18_to_06(row):
    start_time = row['Clock In Time']
    end_time = row['Clock Out Time']
    
    # Define the time boundaries for a 48-hour day
    start_boundary = pd.Timedelta(hours=18)  # 18:00
    end_boundary = pd.Timedelta(days=1, hours=6)  # 06:00 the next day
    
    # Ensure start and end times are within the boundaries
    start_time_within_range = max(start_time, start_boundary)
    end_time_within_range = min(end_time, end_boundary)
    
    # Calculate hours worked within the boundaries, ensuring non-negative
    hours_worked = max(0, (end_time_within_range - start_time_within_range).total_seconds() / 3600)
    
    return pd.Timedelta(hours=hours_worked)

# Apply the functions to calculate hours worked between 14:00 and 18:00 and between 18:00 and 06:00
df['17,5%\n14:00 - 18:00'] = df.apply(calculate_hours_worked_14_to_18, axis=1)
df['23%\n18:00 -06:00'] = df.apply(calculate_hours_worked_18_to_06, axis=1)

# Save the DataFrame to Excel
df.to_excel("MsRaymondeMosheSheet1Output.xlsx", index=False, engine='xlsxwriter')
