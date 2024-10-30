import pandas as pd

# Assuming your data is stored in a CSV file named 'data.csv'
# You can load the data into a DataFrame
df = pd.read_excel(r'C:\Users\Q624157\Desktop\Clocking Data.XLSX')

# Sort the DataFrame based on 'Time Event Date'
df.sort_values(by='Time Event Date', inplace=True)

# Group the data by 'Personnel number', 'Time Event Date', and 'Time Event Type'
grouped = df.groupby(['Personnel number', 'Time Event Date', 'Time Event Type'])

# Create empty lists to store rearranged data
new_rows = []
temp_row = {}

# Iterate over each group
for name, group in grouped:
    for _, row in group.iterrows():
        # Store relevant data in temporary dictionary
        if row['Time Event Type'] == 'In':
            temp_row['Personnel number'] = row['Personnel number']
            temp_row['Surname'] = row['Surname']
            temp_row['Name'] = row['Name']
            temp_row['Cost Center'] = row['Cost Center']
            temp_row['Shift Start Time'] = row['Shift Start Time']
            temp_row['Shift End Time'] = row['Shift End Time']
            temp_row['Time Event Date'] = row['Time Event Date']
            temp_row['In Time'] = row['Time Event In/Out Time']
            temp_row['Out Time'] = ''
            temp_row['Time Event Type'] = 'In'
            temp_row['Hours @ 17,5%'] = row['Hours @ 17,5%']
            temp_row['Hours @ 23%'] = row['Hours @ 23%']
        elif row['Time Event Type'] == 'Out':
            temp_row['Out Time'] = row['Time Event In/Out Time']
            temp_row['Time Event Type'] = 'Out'
    
    # Append the completed row to the new_rows list
    new_rows.append(temp_row.copy())

# Convert the new_rows list to a DataFrame
new_df = pd.DataFrame(new_rows)

# Save the rearranged data to a new CSV file
new_df.to_excel('rearranged_data.xlsx', index=False)
