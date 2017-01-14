#!/usr/bin/env python

## Initial Variables
#    largest_positive_change = 0
#    largest_positive_change_week = 0
#    previous_week_value = 0
#    current_week_value = -99999

## Pseudo Code
#    Open data file for reading.
#    While has next line
#        Set previous_week_value to current week value
#        Read Line
#        Parse Line
#        Read current_week_value
#        Read current_week_date
#        if previous_week_value = -99990
#            continue
#        else
#            current change  = current week value - previous week value
#            if current_change > largest_positive_change:
#                largest_positive_change = current_change
#                largest_positive_change_week = current_week_date
