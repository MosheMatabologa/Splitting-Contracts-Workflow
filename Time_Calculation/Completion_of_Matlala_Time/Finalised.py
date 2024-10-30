import pandas as pd

# Define the file path
file_path = r'C:\Users\Q624157\Desktop\Time_Calculation\Clocking Data.XLSX'

# Read the data into a DataFrame and specify that the first row contains the header
df = pd.read_excel(file_path, sheet_name="Sheet1", header=0)

# Convert the 'Time Event Date' column to datetime
df['Time Event Date'] = pd.to_datetime(df['Time Event Date'])

# Replace '24:00' with '00:00:00' and convert to timedelta
df['Time Event In/Out Time'] = pd.to_timedelta(
    df['Time Event In/Out Time'].astype(str).str.replace('24:00', '00:00:00'), 
    errors='coerce'
)

# Sort by 'Time Event Date' and 'Time Event In/Out Time'
df.sort_values(by=['Time Event Date', 'Time Event In/Out Time'], inplace=True)

# Function to calculate worked hours within a specific time range
def calculate_hours(group):
    # Initialize variables to store clock-in and clock-out times
    clock_in_time = None
    hours_14_18 = pd.Timedelta(0)
    hours_18_06 = pd.Timedelta(0)

    for _, row in group.iterrows():
        if row['Time Event Type'] == 'P10':  # Clock-in event
            clock_in_time = row['Time Event In/Out Time']
        elif row['Time Event Type'] == 'P20' and clock_in_time is not None:  # Clock-out event
            clock_out_time = row['Time Event In/Out Time']
            # Adjust for shifts crossing over midnight
            if clock_out_time < clock_in_time:
                clock_out_time += pd.Timedelta(days=1)
            # Calculate hours worked between 14:00-18:00
            shift_start = max(clock_in_time, pd.Timedelta(hours=14))
            shift_end = min(clock_out_time, pd.Timedelta(hours=18))
            if shift_start < shift_end:
                hours_14_18 += shift_end - shift_start
            # Calculate hours worked between 18:00-06:00
            shift_start = max(clock_in_time, pd.Timedelta(hours=18))
            shift_end = min(clock_out_time, pd.Timedelta(hours=30))  # 30 represents 06:00 next day
            if shift_start < shift_end:
                hours_18_06 += shift_end - shift_start
            # Reset clock-in time for the next pair
            clock_in_time = None

    return hours_14_18, hours_18_06

# Group by 'Time Event Date' and apply the function
hours_14_18_list = []
hours_18_06_list = []
for date, group in df.groupby('Time Event Date'):
    hours_14_18, hours_18_06 = calculate_hours(group)
    hours_14_18_list.extend([hours_14_18.total_seconds() / 3600.0] * len(group))
    hours_18_06_list.extend([hours_18_06.total_seconds() / 3600.0] * len(group))

# Assign calculated hours to DataFrame
df['Hours @ 17.5%'] = hours_14_18_list
df['Hours @ 23%'] = hours_18_06_list

# Save the results to an Excel file
output_file_path = r'C:\Users\Q624157\Desktop\Time_Calculation\UpdatedOutput.xlsx'
df.to_excel(output_file_path, index=False, engine='xlsxwriter')
