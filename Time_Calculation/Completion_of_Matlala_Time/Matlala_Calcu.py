"""
This is the final piece of the code
step 1 exctract file
step 2 read file
step 3 identify times
step 4 calculate times in various zones
"""

import pandas as pd

file1 = r'C:\Users\Q624157\Desktop\Clocking Data.XLSX'


# Define the custom function to calculate hours worked between 14:00-18:00
def calculate_hours_worked_14_to_18(row):
   clock_in = pd.to_datetime(row['Time Event In/Out Time']).time()
   clock_out = pd.to_datetime(row['Time Event In/Out Time']).time()
   
   # Check if the clock in/out times fall within the specified time range
   if clock_in >= pd.to_datetime('14:00').time() and clock_out <= pd.to_datetime('18:00').time():
       return (clock_out.hour - clock_in.hour) + (clock_out.minute - clock_in.minute) / 60
   else:
       return 0

# Define the custom function to calculate hours worked between 18:00-06:00
def calculate_hours_worked_18_to_06(row):
   clock_in = pd.to_datetime(row['Time Event In/Out Time']).time()
   clock_out = pd.to_datetime(row['Time Event In/Out Time']).time()
   
   # Check if the clock in/out times fall within the specified time range
   if clock_in >= pd.to_datetime('18:00').time() or clock_out <= pd.to_datetime('06:00').time():
       return (clock_out.hour - clock_in.hour) + (clock_out.minute - clock_in.minute) / 60
   else:
       return 0

# Load the DataFrame from your data source
df = pd.read_csv("your_data.csv")

# Apply the custom functions to calculate hours worked between 14:00-18:00 and 18:00-06:00
df['Hours @ 17,5%'] = df.apply(calculate_hours_worked_14_to_18, axis=1)
df['Hours @ 23%'] = df.apply(calculate_hours 
