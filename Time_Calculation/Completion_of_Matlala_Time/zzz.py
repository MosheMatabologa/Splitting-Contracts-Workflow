import pandas as pd

# Define the file path
file_path = r'C:\Users\Q624157\Desktop\Time_Calculation\Clocking Data.XLSX'

# Read the data into a DataFrame
df = pd.read_excel(file_path, sheet_name="Sheet1")

# Convert the 'Time Event Date' column to datetime
df['Time Event Date'] = pd.to_datetime(df['Time Event Date'])

# Replace '24:00' with '00:00:00' and convert to timedelta
df['Time Event In/Out Time'] = pd.to_timedelta(
    df['Time Event In/Out Time'].astype(str).str.replace('24:00', '00:00:00'), 
    errors='coerce'
)

# Sort by 'Time Event Date' and 'Time Event In/Out Time'
df.sort_values(by=['Time Event Date', 'Time Event In/Out Time'], inplace=True)

# Function to calculate worked hours in a shift
def calculate_shift_hours(group, start_hour, end_hour):
    shift_hours = pd.Timedelta(0)
    clock_in_time = None

    for _, row in group.iterrows():
        if row['Time Event Type'] == 'In':
            clock_in_time = row['Time Event In/Out Time']
        elif row['Time Event Type'] == 'Out' and clock_in_time is not None:
            clock_out_time = row['Time Event In/Out Time']
            # Calculate overlap with the shift
            shift_start = max(clock_in_time, pd.Timedelta(hours=start_hour))
            shift_end = min(clock_out_time, pd.Timedelta(hours=end_hour))
            if shift_start < shift_end:  # There was an overlap with the shift
                shift_hours += shift_end - shift_start
            clock_in_time = None  # Reset for the next pair

    return shift_hours

# Group by 'Time Event Date' and apply the function for each shift
shift_14_18 = df.groupby(['Time Event Date']).apply(lambda g: calculate_shift_hours(g, 14, 18))
shift_18_06 = df.groupby(['Time Event Date']).apply(lambda g: calculate_shift_hours(g, 18, 6))  # Use 6 to represent 06:00

# Add the results to the DataFrame
df['Hours Worked 14:00-18:00'] = shift_14_18.reset_index(level=0, drop=True)
df['Hours Worked 18:00-06:00'] = shift_18_06.reset_index(level=0, drop=True)

# Save the results to an Excel file
output_file_path = r'C:\Users\Q624157\Desktop\Time_Calculation\UpdatedOutput.xlsx'
df.to_excel(output_file_path, index=False, engine='xlsxwriter')
