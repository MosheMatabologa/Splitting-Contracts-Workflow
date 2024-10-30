import pandas as pd

# Read the data from an Excel file
df = pd.read_excel(r'C:\Users\Q624157\Desktop\correct_Copy of Clocking Data.xlsx')

# Create an empty list to hold the modified rows
new_rows = []

# Iterate over the rows of the original DataFrame
for index, row in df.iterrows():
    # Copy the entire row to a new dictionary
    new_row = row.to_dict()
    
    # Check if the row index is odd
    if index % 2 != 0:
        # Extract the "Time Event In/Out Time" data from odd-numbered rows
        clock_out = row['Time Event In/Out Time']
        # Add the "Clock Out" data to the new row
        new_row['Clock Out'] = clock_out
    
    # Append the new row to the list
    new_rows.append(new_row)

# Create a new DataFrame from the list of modified rows
new_df = pd.DataFrame(new_rows)

# Save the rearranged data to a new Excel file
new_df.to_excel('workingwithArnavtestrearranged_data.xlsx', index=False)
