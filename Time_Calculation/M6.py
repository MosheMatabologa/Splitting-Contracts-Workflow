import pandas as pd
from datetime import timedelta, datetime, time

# Read data from Excel file
file_path = r'C:\Users\Q624157\Desktop\Clocking Data.XLSX'
df = pd.read_excel(file_path)

# Define function to calculate hours worked between specific time ranges
def calculate_hours_worked(start_time, end_time):
    # Define time boundaries for the specified time ranges
    start_boundary_14_to_18 = time(14, 0)  # 14:00
    end_boundary_14_to_18 = time(18, 0)    # 18:00
    end_boundary_18_to_06 = time(6, 0)     # 06:00 the next day

    # Adjust date to current date to avoid comparisons between timedelta and time objects
    current_date = datetime.now().date()

    # Convert time objects to datetime objects for comparison
    start_time = datetime.combine(current_date, start_time)
    end_time = datetime.combine(current_date, end_time)

    # Calculate adjusted end time for shifts ending after midnight
    if end_time < start_time:
        end_time += timedelta(days=1)

    # Ensure start and end times are within the boundaries
    start_time_within_range = max(start_time, datetime.combine(current_date, start_boundary_14_to_18))
    end_time_within_range = min(end_time, datetime.combine(current_date, end_boundary_18_to_06))

    # Calculate hours worked within the boundaries
    hours_worked_14_to_18 = max(0, (min(end_time_within_range, datetime.combine(current_date, end_boundary_14_to_18)) - start_time_within_range).total_seconds() / 3600)
    hours_worked_18_to_06 = max(0, (end_time_within_range - max(start_time_within_range, datetime.combine(current_date, end_boundary_14_to_18))).total_seconds() / 3600)

    return hours_worked_14_to_18, hours_worked_18_to_06

# Iterate over rows to calculate hours worked and update the DataFrame
for index, row in df.iterrows():
    if index % 2 == 0:  # Even numbered rows, representing clock in times
        shift_start_time = row['Shift Start Time']
        shift_end_time = row['Shift End Time']
        clock_out_time = df.at[index + 1, 'Time Event In/Out Time']
        hours_14_to_18, hours_18_to_06 = calculate_hours_worked(shift_start_time, clock_out_time)
        df.at[index, 'Hours @ 17,5%'] = hours_14_to_18
        df.at[index, 'Hours @ 23%'] = hours_18_to_06
    else:  # Odd numbered rows, representing clock out times
        pass  # Skip as we already processed these in the even numbered rows

# Convert the columns to float with 2 decimal places
df['Hours @ 17,5%'] = df['Hours @ 17,5%'].astype(float).round(2)
df['Hours @ 23%'] = df['Hours @ 23%'].astype(float).round(2)

# Save the updated DataFrame to Excel
output_file_path = "M8MsRaymondeMosheSheet1Output.xlsx"
df.to_excel(output_file_path, index=False, engine='xlsxwriter')

print(f"Output saved to {output_file_path}")
