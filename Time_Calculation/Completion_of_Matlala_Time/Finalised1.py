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
def calculate_hours(row):
    clock_in_time = row['Time Event In/Out Time']
    clock_out_time = row['Time Event In/Out Time']
 
    # Adjust for shifts crossing over midnight
    if clock_out_time < clock_in_time:
        clock_out_time += pd.Timedelta(days=1)
 
    # Convert clock-in and clock-out times to Timestamp objects
    clock_in_timestamp = pd.Timestamp(row['Time Event Date']) + clock_in_time
    clock_out_timestamp = pd.Timestamp(row['Time Event Date']) + clock_out_time
 
    # Initialize hours worked for each time range
    hours_14_18 = 0
    hours_18_06 = 0
 
    # Check if the shift overlaps with the time ranges and calculate the hours accordingly
    if pd.Timestamp(row['Time Event Date']).replace(hour=14) <= clock_out_timestamp <= pd.Timestamp(row['Time Event Date']).replace(hour=18):
        hours_14_18 = (min(clock_out_timestamp, pd.Timestamp(row['Time Event Date']).replace(hour=18)) - max(clock_in_timestamp, pd.Timestamp(row['Time Event Date']).replace(hour=14))).total_seconds() / 3600
 
    if pd.Timestamp(row['Time Event Date']).replace(hour=18) <= clock_out_timestamp <= pd.Timestamp(row['Time Event Date']).replace(hour=23, minute=59):
        hours_14_18 = (pd.Timestamp(row['Time Event Date']).replace(hour=18) - max(clock_in_timestamp, pd.Timestamp(row['Time Event Date']).replace(hour=14))).total_seconds() / 3600
        hours_18_06 = (min(clock_out_timestamp, pd.Timestamp(row['Time Event Date']).replace(hour=23, minute=59)) - pd.Timestamp(row['Time Event Date']).replace(hour=18)).total_seconds() / 3600
 
    if pd.Timestamp(row['Time Event Date']).replace(hour=0) <= clock_out_timestamp <= pd.Timestamp(row['Time Event Date']).replace(hour=6):
        hours_18_06 = (min(clock_out_timestamp, pd.Timestamp(row['Time Event Date']).replace(hour=6)) - max(clock_in_timestamp, pd.Timestamp(row['Time Event Date']).replace(hour=0))).total_seconds() / 3600
 
    return hours_14_18, hours_18_06
 
# Calculate worked hours for each row
hours_14_18_list = []
hours_18_06_list = []
for index, row in df.iterrows():
    hours_14_18, hours_18_06 = calculate_hours(row)
    hours_14_18_list.append(hours_14_18)
    hours_18_06_list.append(hours_18_06)
 
# Assign calculated hours to DataFrame
df['Hours @ 14:00-18:00'] = hours_14_18_list
df['Hours @ 18:00-06:00'] = hours_18_06_list
 
# Save the results to an Excel file
output_file_path = r'C:\Users\Q624157\Desktop\Time_Calculation\UpdatedOutput.xlsx'
df.to_excel(output_file_path, index=False, engine='xlsxwriter')