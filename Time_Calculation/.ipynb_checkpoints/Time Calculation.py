#Time Calculation 
import pandas as pd

file1 = r'C:\temp\Time_Problem.xlsx'

#read data from excel file

df = pd.read_excel(file1, sheet_name = "Problem")

#change format from string to Date

df['Clock Out Time'] = df['Clock Out Time'].astype('datetime64[ns]')
df['Clock In Time'] = df['Clock In Time'].astype('datetime64[ns]')

#difference between clock out and clockin . How long he was there
df['Clocking Totals (Day)']=df['Clock Out Time']-df['Clock In Time']

#preparing timestamp for 18:00 and 22:00
from datetime import datetime
# datetime(year, month, day, hour, minute, second, microsecond)
var18= datetime(2024, 4, 16, 18, 0, 0)

print(var18)

df['17,5%\n14:00 - 18:00']=var18-df['Clock In Time']
df['23%\n18:00 -06:00']=df['Clock Out Time']-var18


# datetime(year, month, day, hour, minute, second, microsecond)
var18= datetime(2024, 4, 16, 18, 0, 0)

print(var18)

df.to_excel("output.xlsx")