#!/usr/bin/env python

## Initial Variables
#    largest_positive_change = 0
#    largest_positive_change_week = 0
#    previous_week_value = 0
#    current_week_value = -99999

## Pseudo Code
#    Open data file for reading.
#    While has next line
#        Set prvious_week_value to current week value
#        Read Line
#        Parse Line
#        Read current_week_value
#        Read current_week_date
#        if previous_week_value = -99990
#            Break
#        else
#            current change  = current week value - previous week value
#            if current_change > largest_positive_change:
#                largest_positive_change = current_change
#                largest_positive_change_week = current_week_date

# Import Library to Support Excel 2010 Workbook
from openpyxl import load_workbook

# Read in workbook and use the default sheet
wb = load_workbook('2016-11_SP_10years1.xlsx', read_only=True)
ws = wb.active

# Initial Values
largest_positive_change = 0
largest_positive_change_week = "1970-01-01 00:00:00"
previous_week_value = 0
current_week_value = -99999

# Processing logic.
# Let's hardcode the range and make assumptions on the data layout
# as the first pass through it would be difficult to analize the file
# for data.
for i in range(391, 1, -1):
    # print 'previous_week_value: %s, current_week_value: %s' % (previous_week_value, current_week_value)
    previous_week_value = current_week_value
    # print 'previous_week_value: %s, current_week_value: %s' % (previous_week_value, current_week_value)
    start_column = 'A%s' % i
    end_column = 'E%s' % i
    # print 'start_column: %s, end_column: %s' % (start_column, end_column)
    columns_range = ws[start_column:end_column]
    ((close_date, close, open, high, low),) = columns_range
    current_week_value = close.value
    # print close_date.value
    # print close.value
    if previous_week_value != -99999:
        current_week_difference = current_week_value - previous_week_value
        if current_week_difference > largest_positive_change:
            largest_positive_change = current_week_difference
            largest_positive_change_week = '%s' % close_date.value

# Print out what we found to be the largest change and the week of the change.
print 'largest_positive_change: %s, largest_positive_change_week: %s' % (largest_positive_change, largest_positive_change_week)
