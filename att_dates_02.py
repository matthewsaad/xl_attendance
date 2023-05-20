import openpyxl
import pandas as pd
from sys import argv # Not needed yet
from datetime import datetime
from openpyxl import Workbook
import tkinter as tk


# Open the excel spread sheet nd
wb = openpyxl.load_workbook('Test_Attendance.xlsx')
sheet = wb.get_sheet_by_name('Sheet1')

cells = ['C6', 'E6', 'G6', 'I6', 'K6', 
		'C16', 'E16', 'G16', 'I16', 'K16',
		'C26', 'E26', 'G26', 'I26', 'K26']

start_date = '2023-05-16'
end_date = '2023-05-31'
print(str(start_date))
# Takes start and end date arguments and gives list of only Mon-Fri dates.
dates = pd.bdate_range(start='2023-07-01', end='2023-07-15').tolist()

work_dates = []
for date in dates:
	work_dates.append(date.strftime('%m/%d/%Y'))

# Change the first_date to first_day: 'Monday, Tuesday...'
first_date = dates[0]
first_day = []
first_day.append(first_date.strftime('%A'))

# Match the day with the correct cell and list remaining cells.
if first_day[0] == 'Monday':
	sheet['C6'] = first_date
	remain_cells = cells[1:15]
elif first_day[0] == 'Tuesday':
	sheet['E6'] = first_date
	remain_cells = cells[2:15]
elif first_day[0] == 'Wednesday':
	sheet['G6'] = first_date
	remain_cells = cells[3:15]
elif first_day[0] == 'Thursday':
	sheet['I6'] = first_date
	remain_cells = cells[4:15]
elif first_day[0] == 'Friday':
	sheet['K6'] = first_date
	remain_cells = cells[5:15]
	
# Subtract the difference of remaining cells from the number of remaining dates and add 1.
# Use that to slice only the needed remaining cells.
no_remain_cells = len(remain_cells)
no_remain_wd = len(work_dates)
y = int(no_remain_cells) - (int(no_remain_cells) - int(no_remain_wd) + 1)

remain_cells = remain_cells[0:y]
remain_dates = work_dates[1:15]

# Use enumerate to access both index and value
for i, cell in enumerate(remain_cells): 
# Assign value to cell using '.value' 
    sheet[cell].value = remain_dates[i]  

wb.save('Test_Attendance.xlsx')



