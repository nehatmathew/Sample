# -*- coding: utf-8 -*-
"""
Created on Tue Nov  7 15:57:48 2023

@author: AshisMohanty
#HEY!
#step 1
#step 2
This is Step 2. Run the Supramax.py file before running this file
"""
import pandas as pd
import numpy as np 
from datetime import datetime

baltic_data = pd.read_csv(str('C:/Users/NehaMathew/Documents/Baltic/Baltic Exchange - Historic_Step1_'+str(datetime.now().date())+'.csv'))

'''
try:
    baltic_data['Date'] = pd.to_datetime(baltic_data['Date'], format = '%d-%m-%Y').dt.strftime('%d/%m/%Y')
    baltic_data['Date'] = pd.to_datetime(baltic_data['Date'], format='%d/%m/%Y')
    
except ValueError:
    baltic_data['Date'] = pd.to_datetime(baltic_data['Date'], format='%d/%m/%Y')
 
    
baltic_data.set_index('Date', inplace=True)


try:
    baltic_data['Date'] = pd.to_datetime(baltic_data['Date'], format='%d-%b-%Y').dt.strftime('%d/%m/%Y')
    baltic_data['Date'] = pd.to_datetime(baltic_data['Date'], format='%d/%m/%Y')
except ValueError:
    baltic_data['Date'] = pd.to_datetime(baltic_data['Date'], format='%d/%b/%Y')

baltic_data.set_index('Date', inplace=True)



try:
    # First, try the format with a two-digit year
    baltic_data['Date'] = pd.to_datetime(baltic_data['Date'], format='%d-%b-%y').dt.strftime('%d/%m/%Y')
    # Convert the string back to datetime format
    baltic_data['Date'] = pd.to_datetime(baltic_data['Date'], format='%d/%m/%Y')
except ValueError:
    # Fallback to handle other possible formats
    baltic_data['Date'] = pd.to_datetime(baltic_data['Date'], format='%d/%b/%Y')

# Set the 'Date' column as the index
baltic_data.set_index('Date', inplace=True) 


'''
# Convert the 'Date' column to datetime, handling mixed formats
baltic_data['Date'] = pd.to_datetime(baltic_data['Date'], format='mixed', dayfirst=True, errors='coerce')

# Check for and handle any invalid (NaT) dates
if baltic_data['Date'].isna().any():
    print("Some dates couldn't be parsed. Investigate invalid rows:")
    print(baltic_data[baltic_data['Date'].isna()])

# Set the 'Date' column as the index
baltic_data.set_index('Date', inplace=True)





#create a custom frequency for week ending on Saturday
custom_weekly_frequency = 'W-SAT'
#calculate the mean
weekly_average = baltic_data.resample(custom_weekly_frequency).mean()

#rename the index
weekly_average = weekly_average.rename_axis('Week_Starting_Date')

#add columns for Year and Week
weekly_average['Year'] = weekly_average.index.year
weekly_average['Week'] = weekly_average.index.isocalendar().week
#Display index as Start of the week (as displayed on Connect)
weekly_average.index = weekly_average.index - pd.Timedelta(days=6)
# Specify the column you want to move to position 0
column_to_move = ['Year', 'Week']  # Example: Moving column 'B' to position 0
# Create a new DataFrame with the specified column at position 0
weekly_average = pd.concat([weekly_average[column_to_move], weekly_average.drop(column_to_move, axis=1)], axis=1)


#Aggregate the daily data (pivot table) on monthly level
monthly_average = baltic_data.resample('MS').mean()
#add columns for Year and Week
monthly_average['Year'] = monthly_average.index.year
monthly_average['Month'] = monthly_average.index.month
# Specify the column you want to move to position 0
column_to_move = ['Year', 'Month']  # Example: Moving column 'B' to position 0
# Create a new DataFrame with the specified column at position 0
monthly_average = pd.concat([monthly_average[column_to_move], monthly_average.drop(column_to_move, axis=1)], axis=1)


annual_average = baltic_data.resample('YS').mean()
#add columns for Year and Week
annual_average['Year'] = annual_average.index.year
# Specify the column you want to move to position 0
column_to_move = ['Year']  # Example: Moving column 'B' to position 0
# Create a new DataFrame with the specified column at position 0
annual_average = pd.concat([annual_average[column_to_move], annual_average.drop(column_to_move, axis=1)], axis=1)



baltic_data['Year'] = baltic_data.index.year
baltic_data['Week'] = baltic_data.index.isocalendar().week

# Specify the column you want to move to position 0
column_to_move = ['Year', 'Week']  # Example: Moving column 'B' to position 0
# Create a new DataFrame with the specified column at position 0
baltic_data = pd.concat([baltic_data[column_to_move], baltic_data.drop(column_to_move, axis=1)], axis=1)

# Provide names for the Excel tabs
sheet_names = ['Daily', 'Weekly', 'Monthly', 'Yearly']

# Specify the Excel file name with the current date
current_date = datetime.now().strftime('%Y-%m-%d')  # Format the current date as YYYY-MM-DD
excel_file = f'C:/Users/NehaMathew/Documents/Baltic/Baltic_Data_average_{current_date}.xlsx'  # Add the formatted date to the filename



with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
    for sheet_name, data in zip(sheet_names, [baltic_data, weekly_average, monthly_average, annual_average]):
        data.to_excel(writer, sheet_name=sheet_name, index=True)